"""One bounded semantic matching attempt over already parsed source schemas."""
import json


def match_sources_with_ai(matcher, training, actual, provider_name, determined=None):
    from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout

    locked = []
    for match in determined or []:
        ti = next((i for i, s in enumerate(training) if (s['file_name'], s['sheet_name']) ==
                   (match['train_file'], match['train_sheet'])), None)
        ai = next((i for i, s in enumerate(actual) if (s['file_name'], s['sheet_name']) ==
                   (match['input_file'], match['input_sheet'])), None)
        if ti is not None and ai is not None and set(match['col_mapping'].values()) == set(training[ti]['headers']):
            locked.append({'training_id': ti, 'actual_id': ai, 'columns': match['col_mapping']})
    used_t = {m['training_id'] for m in locked}
    used_a = {m['actual_id'] for m in locked}
    candidates, fixed = [], {}
    for ti, t in enumerate(training):
        if ti in used_t:
            continue
        choices = []
        for ai, a in enumerate(actual):
            if ai in used_a:
                continue
            same = {c: c for c in a['headers'] if c in t['headers']}
            fixed[ti, ai] = same
            choices.append({'actual_id': ai, 'file': a['file_name'], 'sheet': a['sheet_name'],
                'matched_column_count': len(same),
                'missing_columns': [c for c in t['headers'] if c not in same],
                'candidate_columns': [c for c in a['headers'] if c not in same]})
        candidates.append({'training_id': ti, 'file': t['file_name'], 'sheet': t['sheet_name'], 'candidates': choices})

    prompt = (
        '将训练数据源映射到本次上传的数据源。根据完整文件名、sheet名和字段语义选择，'
        '不能只凭一个关键词。每个训练表必须唯一对应一个实际表，不得复用实际表。'
        '字段只能映射语义相同的字段，不能猜测缺失数据，不确定请返回空 mappings。'
        '输出 JSON: {"mappings":[{"training_id":0,"actual_id":0,'
        '"columns":{"实际列名":"训练列名"}}]}。只返回缺失列映射；同名列已由程序锁定，无需输出。'
        '选择正确候选，覆盖其 missing_columns；candidate_columns 中无对应则不要猜测。\n'
        + json.dumps({'unresolved': candidates}, ensure_ascii=False)
    )
    raw = chat_with_timeout(AIProviderFactory.create_provider(provider_name),
                            [{'role': 'user', 'content': prompt}], max_tokens=6000)
    if not raw:
        raise ValueError('AI 源数据匹配超时或未返回结果')
    text = raw.strip()
    if text.startswith('```'):
        text = text.split('\n', 1)[1].rsplit('```', 1)[0].strip()
    response = json.loads(text)
    entries = response.get('mappings') if isinstance(response, dict) else None
    if not isinstance(entries, list):
        raise ValueError('AI 映射格式无效')
    for entry in entries:
        ti, ai = entry.get('training_id'), entry.get('actual_id')
        if type(ti) is not int or type(ai) is not int or (ti, ai) not in fixed:
            raise ValueError('AI 修改了已确定映射或选择了无效候选')
        columns = entry.get('columns')
        if not isinstance(columns, dict):
            raise ValueError('AI 列映射格式无效')
        if any(k in fixed[ti, ai] and v != fixed[ti, ai][k] for k, v in columns.items()):
            raise ValueError('AI 覆盖了已确定列')
        entry['columns'] = {**fixed[ti, ai], **columns}
    return validate_mapping(matcher, training, actual, {'mappings': locked + entries})


def validate_mapping(matcher, training, actual, response):
    entries = response.get('mappings') if isinstance(response, dict) else None
    if not isinstance(entries, list) or not training or len(entries) != len(training):
        raise ValueError('AI 映射未覆盖全部训练表')
    used_training, used_actual, matches = set(), set(), []
    file_targets, file_columns, target_files = {}, {}, {}
    for item in entries:
        ti, ai = item.get('training_id'), item.get('actual_id')
        if (type(ti) is not int or type(ai) is not int or not 0 <= ti < len(training)
                or not 0 <= ai < len(actual) or ti in used_training or ai in used_actual):
            raise ValueError('AI 表映射重复或引用不存在的表')
        t, a = training[ti], actual[ai]
        columns = item.get('columns')
        if (not isinstance(columns, dict) or not columns
                or any(not isinstance(v, str) for v in columns.values())
                or not set(columns).issubset(a['headers'])
                or set(columns.values()) != set(t['headers'])
                or len(set(columns.values())) != len(columns)):
            raise ValueError('AI 列映射缺失、重复或引用不存在的字段')
        # 当前加载器按文件共享列映射，不允许跨 sheet 的同名列含义冲突。
        name = a['file_name']
        if file_targets.setdefault(name, t['file_name']) != t['file_name']:
            raise ValueError('同一上传文件不能映射到多个训练文件')
        if target_files.setdefault(t['file_name'], name) != name:
            raise ValueError('多个上传文件不能覆盖同一训练文件')
        merged = file_columns.setdefault(name, {})
        for source, target in columns.items():
            if merged.setdefault(source, target) != target:
                raise ValueError('同一文件不同 sheet 的列映射冲突')
        renamed = [columns.get(h, h) for h in a['headers']]
        if len(set(renamed)) != len(renamed):
            raise ValueError('AI 列映射与原有列重名')
        used_training.add(ti)
        used_actual.add(ai)
        matches.append({'train_file': t['file_name'], 'train_sheet': t['sheet_name'],
                        'input_file': name, 'input_file_path': a['file_path'],
                        'input_sheet': a['sheet_name'], 'col_mapping': columns,
                        'needs_rewrite': True})
    for match in matches:
        a = next(s for s in actual if s['file_name'] == match['input_file']
                 and s['sheet_name'] == match['input_sheet'])
        merged = file_columns[match['input_file']]
        if any(merged[k] != v for k, v in match['col_mapping'].items()):
            raise ValueError('合并后的字段映射冲突')
        renamed = [merged.get(h, h) for h in a['headers']]
        if len(set(renamed)) != len(renamed):
            raise ValueError('合并后的字段映射产生重名列')
    return {'success': True, 'mapping': {'file_mapping': matcher._build_file_mapping(matches)}}
