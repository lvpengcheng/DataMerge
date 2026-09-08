"""Current input contract shared by every conversational editing mode."""
import json


def build_training_edit_context(config, context):
    sources = context.get('source_structure') or {}
    payload = {
        'input_revision': config.get('input_revision'),
        'mode': config.get('mode'),
        'source_structure': sources,
        'expected_structure': config.get('expected_structure') or context.get('expected_structure') or {},
        'target_sheets': config.get('target_sheets'),
        'manual_headers': config.get('manual_headers'),
        'multi_sheet_source': config.get('multi_sheet_source', False),
        'salary_year': config.get('salary_year'),
        'salary_month': config.get('salary_month'),
    }
    if not sources:
        payload['source_structure_description'] = config.get('source_structure_desc', '')
    return (
        '\n## 当前已保存的训练输入（优先于旧脚本中的文件结构和历史对话）\n'
        + json.dumps(payload, ensure_ascii=False, default=str)
        + '\n文件或结构是事实背景，不是自动修改全部逻辑的授权。只落实本轮指示及其必要依赖。'
        '\n可精确修改读取、清洗、去重、关联、汇总、组装、输出结构和公式。'
        '新增步骤必须同时接入执行调用链，不能只定义函数；同步必要的 import、字段及调用点。'
        '明确关联主键、关联类型、重复键处理、缺失值处理和汇总粒度；不清楚时不要臆造业务规则。'
        '使用执行环境提供的当前源数据和模板，不硬编码旧上传路径；公共工具无法直接编辑时，'
        '在当前脚本中增加必要的局部处理。保留无关规则及模板样式。'
        + ('\n输入已更新：历史评分和差异已失效，不得据此额外修改旧差异列。' if config.get('validation_stale') else '')
        + '\n'
    )
