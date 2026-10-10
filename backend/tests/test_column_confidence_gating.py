"""列级置信度阈值：高置信自动通过，低置信才进入人工确认。"""
from backend.utils import compute_precheck as pre
from backend.utils.compute_ingest import IngestMeta, resolve_with_confirmations
from backend.utils.fast_header_matcher import FastHeaderMatcher


def _meta():
    train = [{'file_name': 'train.xlsx', 'sheet_name': '工资',
              'headers': {'工号': 'A', '金额': 'B'}}]
    actual = [{'file_name': 'upload.xlsx', 'file_path': '/tmp/upload.xlsx',
               'sheet_name': 'Payroll', 'headers': {'ID': 'A', 'Pay': 'B'}}]
    return IngestMeta(
        train_sheets=train,
        input_sheets=actual,
        source_structure={'files': {'train.xlsx': {
            'sheets': {'工资': {'headers': train[0]['headers']}}}}},
        ai_provider_name='claude',
    )


def _ai_mapping(pay_confidence):
    return {
        'success': True,
        'mapping': {'file_mapping': {
            'upload.xlsx': {
                'expected_file': 'train.xlsx',
                'sheet_mapping': {'Payroll': '工资'},
                'header_mapping': {'ID': '工号', 'Pay': '金额'},
                'header_mapping_by_sheet': {'Payroll': {'ID': '工号', 'Pay': '金额'}},
                'header_confidence_by_sheet': {'Payroll': {'ID': 0.99, 'Pay': pay_confidence}},
                'sheet_confidence': {'Payroll': 0.95},
                'file_confidence': 0.95,
                'needs_rewrite': True,
            }
        }},
        'needs_confirmation': True,
        'ai_suggestions': [
            {'expected_path': 'train.xlsx > 工资 > 金额',
             'suggested_path': 'upload.xlsx > Payroll > Pay',
             'confidence': pay_confidence, 'reason': '测试建议'},
            {'expected_path': 'train.xlsx > 工资 > 工号',
             'suggested_path': 'upload.xlsx > Payroll > ID',
             'confidence': 0.99, 'reason': '测试建议'},
        ],
        'source_sheet_reviews': [{
            'expected_file': 'train.xlsx', 'expected_sheet': '工资',
            'suggested_file': 'upload.xlsx', 'suggested_sheet': 'Payroll',
            'confidence': 0.95, 'reason': '测试 Sheet 建议',
            'recommendation_source': 'ai',
        }],
    }


def test_low_confidence_columns_do_not_enter_confirmation(monkeypatch):
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *args: None)
    monkeypatch.setattr(FastHeaderMatcher, 'match_headers_only',
                        lambda self, *args, **kwargs: _ai_mapping(0.70))
    result = resolve_with_confirmations(_meta(), skip_history_check=True)
    assert not result.ok
    assert result.mapping_requires_confirmation
    assert result.missing_columns == []
    assert result.ai_suggestions == []
    assert result.source_sheet_reviews[0]['suggested_sheet'] == 'Payroll'


def test_high_confidence_columns_do_not_require_column_confirmation(monkeypatch):
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *args: None)
    monkeypatch.setattr(FastHeaderMatcher, 'match_headers_only',
                        lambda self, *args, **kwargs: _ai_mapping(0.96))
    result = resolve_with_confirmations(_meta(), skip_history_check=True)
    # AI 文件/Sheet 层仍需要确认；完整 AI 结果都进入一次性审核弹窗。
    assert not result.ok
    assert result.mapping_requires_confirmation
    assert result.missing_columns == []
    assert result.ai_suggestions == []
    assert result.source_sheet_reviews[0]['suggested_sheet'] == 'Payroll'

def test_ai_column_payload_shape_does_not_affect_file_sheet_entries(monkeypatch):
    import json
    from backend.ai_engine import ai_provider
    from backend.utils.ai_source_mapping import match_sources_with_ai

    training = [
        {'file_name': 'a.xlsx', 'sheet_name': '工资',
         'headers': {'ID': 'A', '金额': 'B'}},
        {'file_name': 'b.xlsx', 'sheet_name': '社保',
         'headers': {'ID': 'A', '社保': 'B'}},
    ]
    actual = [
        {'file_name': 'up_a.xlsx', 'file_path': '/tmp/up_a.xlsx',
         'sheet_name': 'Pay', 'headers': {'ID': 'A', 'Pay': 'B'}},
        {'file_name': 'up_b.xlsx', 'file_path': '/tmp/up_b.xlsx',
         'sheet_name': 'SI', 'headers': {'ID': 'A', 'SI': 'B'}},
    ]
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', lambda *args, **kwargs: json.dumps({
        'mappings': [
            {'training_id': 0, 'actual_id': 0,
             'columns': {'ID': 'ID', 'Pay': '金额'},
             'file_confidence': 0.9, 'sheet_confidence': 0.9,
             'column_confidence': {'ID': 1.0, 'Pay': 0.8}},
            {'training_id': 1, 'actual_id': 1, 'columns': ['ID', 'SI']},
        ]
    }, ensure_ascii=False))
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, 'claude')
    mapping = result['mapping']['file_mapping']
    assert mapping['up_a.xlsx']['sheet_mapping'] == {'Pay': '工资'}
    assert mapping['up_b.xlsx']['sheet_mapping'] == {'SI': '社保'}
    assert mapping['up_a.xlsx']['header_mapping'] == {'ID': 'ID'}
    assert mapping['up_b.xlsx']['header_mapping'] == {'ID': 'ID'}


def test_multiple_uploaded_files_cannot_map_to_same_training_file(monkeypatch):
    from backend.utils import compute_precheck as pre
    from backend.utils.compute_ingest import IngestMeta, resolve_with_confirmations

    monkeypatch.setattr(pre, '_check_target_sheets', lambda *args: None)
    train = [{'file_name': 'train.xlsx', 'sheet_name': '工资',
              'headers': {'工号': 'A', '金额': 'B'}}]
    inputs = [
        {'file_name': 'upload1.xlsx', 'file_path': '/tmp/upload1.xlsx',
         'sheet_name': 'Sheet1', 'headers': {'工号': 'A', '金额': 'B'}},
        {'file_name': 'upload2.xlsx', 'file_path': '/tmp/upload2.xlsx',
         'sheet_name': 'Sheet2', 'headers': {'工号': 'A', '金额': 'B'}},
    ]
    meta = IngestMeta(
        train_sheets=train,
        input_sheets=inputs,
        source_structure={'files': {'train.xlsx': {'sheets': {'工资': {'headers': train[0]['headers']}}}}},
        rename_candidates=[
            {'uploaded': 'upload1.xlsx', 'candidates': [{'expected': 'train.xlsx', 'score': 1.0}]},
            {'uploaded': 'upload2.xlsx', 'candidates': [{'expected': 'train.xlsx', 'score': 1.0}]},
        ],
    )
    result = resolve_with_confirmations(
        meta,
        confirmed_renames={'upload1.xlsx': 'train.xlsx', 'upload2.xlsx': 'train.xlsx'},
        skip_history_check=True,
    )
    assert not result.ok
    assert '多个上传文件不能映射到同一个训练文件' in result.mapping_notice
    assert {item['uploaded'] for item in result.rename_candidates} == {'upload1.xlsx', 'upload2.xlsx'}
