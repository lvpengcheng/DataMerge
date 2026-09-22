"""One bounded semantic matching attempt over already parsed source schemas."""
import json
import ast
import copy
import logging
import os

from .period_matching import period_candidate_allowed, period_context

logger = logging.getLogger(__name__)


def _script_matching_context(context, training, actual):
    """Expose schema names and code operations, without baked row samples or credentials."""
    context = context or {}
    allowed = {str(value) for sheet in training + actual
               for value in [sheet['file_name'], sheet['sheet_name'], *sheet['headers']]}
    allowed.update(str(s) for s in context.get('target_sheets', []))
    allowed.update(str(c) for cols in context.get('target_columns', {}).values() for c in cols)
    allowed.update({'left', 'right', 'inner', 'outer', 'sum', 'mean', 'first', 'last', 'coerce', 'ignore'})
    result = {key: context.get(key) for key in ('template_name', 'target_sheets', 'target_columns', 'original_file_names')
              if context.get(key)}
    try:
        tree = ast.parse(context.get('script_content') or '')
    except (SyntaxError, ValueError):
        result['logic_unavailable'] = True
        return result
    # Preserve the script's template reference as a basename, never its deployment path.
    template_paths = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in
                {'TEMPLATE_NAME', 'TEMPLATE_PATH', '_TEMPLATE_PATH'} for t in node.targets):
            if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                result['script_template_name'] = node.value.value.replace('\\', '/').rsplit('/', 1)[-1]
                allowed.add(node.value.value)
                template_paths.add(node.value.value)

    class SchemaOnly(ast.NodeTransformer):
        def visit_Constant(self, node):
            if isinstance(node.value, str):
                value = node.value
                if value in allowed:
                    if value in template_paths:
                        value = value.replace('\\', '/').rsplit('/', 1)[-1]
                    return ast.copy_location(ast.Constant(value), node)
                return ast.copy_location(ast.Constant('<literal>'), node)
            if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
                return ast.copy_location(ast.Constant(0), node)
            return node

    snippets = []
    remaining = 16000
    # Include aliases plus function bodies containing lookups, joins, filters,
    # arithmetic and output assignments. Never execute uploaded script code.
    def relevance(node):
        references = sum(isinstance(n, ast.Constant) and isinstance(n.value, str) and n.value in allowed
                         for n in ast.walk(node))
        return references + (10 if getattr(node, 'name', '') in {'main', 'fill_template', 'fill_result_sheets'} else 0)

    for node in sorted(tree.body, key=relevance, reverse=True):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Assign, ast.AnnAssign)):
            continue
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            if not any(isinstance(t, ast.Name) and t.id in
                       {'_SOURCE_MAP', '_SK_TO_SHEET', '_COL_MAP', 'TEMPLATE_NAME'} for t in targets):
                continue
        text = ast.unparse(SchemaOnly().visit(copy.deepcopy(node)))
        if remaining <= 0:
            result['logic_truncated'] = True
            break
        snippet = text[:min(4000, remaining)]
        snippets.append(snippet)
        remaining -= len(snippet)
        if len(snippet) < len(text):
            result['logic_truncated'] = True
    result['script_logic'] = snippets
    return result


def _ai_call_options(provider_name: str) -> dict:
    options = {'max_tokens': 120000}
    if provider_name == 'deepseek':
        options.update(max_tokens=120000, require_complete=True,
                       extra_body={'thinking': {'type': 'disabled'}},
                       response_format={'type': 'json_object'})
    return options


def _call_ai_json(provider_name: str, prompt: str) -> dict:
    from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout
    raw = chat_with_timeout(AIProviderFactory.create_provider(provider_name),
                            [{'role': 'user', 'content': prompt}],
                            raise_on_error=True, **_ai_call_options(provider_name))
    if not raw:
        raise ValueError('AI_RESPONSE_EMPTY: AI 返回了空的匹配结果')
    text = raw.strip()
    if text.startswith('```'):
        text = text.split('\n', 1)[1].rsplit('```', 1)[0].strip()
    return json.loads(text)


def _build_stage1_prompt(candidates, training, actual, matching_context):
    unresolved_training_ids = {
        item.get('training_id') for item in candidates
        if isinstance(item, dict) and type(item.get('training_id')) is int
    }
    unresolved_actual_ids = {
        choice.get('actual_id')
        for item in candidates if isinstance(item, dict)
        for choice in (item.get('candidates') or []) if isinstance(choice, dict)
        if type(choice.get('actual_id')) is int
    }
    return (
        '你是 Excel 源数据映射专家。你的唯一任务：在训练数据源和本次可用数据源（本次上传 + 自动补入的基础资料）之间，建立文件到文件、Sheet 到 Sheet 的一级和二级匹配关系。'
        '本阶段不要处理列映射。'
        '\n一、层级原则：文件是第一层，Sheet 是第二层，列是第三层；必须先确定文件，再在对应文件内确定 Sheet；不允许跨文件拼凑 Sheet。'
        '一个上传文件只能对应一个训练文件，一个训练文件只能对应一个上传文件；一个上传文件内的所有 Sheet 只能映射到同一个训练文件下的 Sheet。'
        '文件关系无法确定时，不要猜测其 Sheet 关系。'
        '\n二、匹配优先级：先硬约束，再结构证据，最后名称相似性。'
        '硬约束包括：工资年月对应的本月/上月角色、用户已确认关系、代码中明确引用的源文件/源 Sheet、名称中的 YYYYMM/地区/机构/人员类别。'
        '结构证据包括：表头结构、列数量、关键列、join key、主键、筛选字段、代码从该 Sheet 读取的列。'
        '名称相似只能作为辅助，不能替代结构和代码证据。'
        'global_actual 和 candidates 中的 origin 表示来源：upload=本次上传、tenant_base=租户基础资料、global_base=全局基础资料。'
        '程序只会为本次上传尚未被确定性结构规则覆盖的训练角色补入基础资料；当候选在结构和语义上等价时，必须按 upload、tenant_base、global_base 的顺序优先。'
        '但来源优先级不能替代结构与代码证据，不能把无关的本次上传强行匹配到训练角色。'
        '\n三、月份角色硬约束：本次工资年月由 salary_year/salary_month 给出；训练名称含本月/当月/本期/当期时，必须对应 salary_year+salary_month 的 YYYYMM；'
        '训练名称含上月/前月/上期/前期时，必须对应前一个月的 YYYYMM；不能把 202607 和 202608 反向匹配。上传名称没有 YYYYMM 时不要仅凭月份猜测。'
        '\n四、输入结构：salary_year、salary_month、current_period_index、global_training、global_actual、training、script_context。'
        'global_training 和 global_actual 只包含程序尚未明确的训练表与可用源表；已经确定的文件和 Sheet 不会发送给你，也不得重新判断；'
        'training 的每一项是一个待匹配训练表，candidates 是允许选择的 actual 候选，actual_id 是唯一候选标识。'
        '\n五、输出要求：严格只输出 JSON，不要输出解释文字、Markdown、代码块标记。格式：'
        '{"mappings":[{"training_id":0,"actual_id":0,"actual_sheet":"上传Sheet名","file_confidence":0.0-1.0,"sheet_confidence":0.0-1.0,'
        '"file_reason":"文件对应依据","sheet_reason":"Sheet 对应依据","code_evidence":"代码引用与用途"}]}。'
        'confidence 是 0~1。没有把握的 mapping 不要输出；不能确定就返回空 mappings，不要为了凑满所有训练表而猜测。'
        '\n输入数据：\n'
        + json.dumps({
            'salary_year': (matching_context or {}).get('salary_year'),
            'salary_month': (matching_context or {}).get('salary_month'),
            'current_period_index': period_context(matching_context),
            'global_training': [
                {'training_id': ti, 'file': t['file_name'], 'sheet': t['sheet_name'],
                 'columns': list(t['headers'].keys())}
                for ti, t in enumerate(training) if ti in unresolved_training_ids
            ],
            'global_actual': [
                {'actual_id': ai, 'file': a['file_name'],
                 'original_file': a.get('original_file_name', a['file_name']),
                 'origin': a.get('source_origin', 'upload'),
                 'asset_name': a.get('source_asset_name'),
                 'sheet': a['sheet_name'],
                 'original_sheet': a.get('original_sheet_name', a['sheet_name']),
                 'columns': list(a['headers'].keys())}
                for ai, a in enumerate(actual) if ai in unresolved_actual_ids
            ],
            'training': candidates,
            'script_context': _script_matching_context(matching_context, training, actual),
        }, ensure_ascii=False)
    )


def _stage1_entry_valid(item, training, actual, fixed, matching_context, used_t, used_a, file_targets, target_files):
    ti = item.get('training_id')
    ai = item.get('actual_id')
    if type(ti) is not int or type(ai) is not int or not 0 <= ti < len(training) or not 0 <= ai < len(actual):
        return None, 'AI 修改了已确定映射或选择了无效候选'
    if (ti, ai) not in fixed:
        return None, 'AI 选择的候选不在允许范围内'
    if ti in used_t or ai in used_a:
        return None, 'AI 表映射重复或引用不存在的表'
    if not period_candidate_allowed(training[ti], actual[ai], matching_context):
        return None, 'AI 月份角色与工资年月冲突'
    actual_sheet = str(item.get('actual_sheet') or actual[ai].get('sheet_name') or '')
    if actual_sheet and actual_sheet != actual[ai].get('sheet_name'):
        return None, 'AI 返回的 actual_sheet 与 actual_id 不一致'
    actual_file = actual[ai]['file_name']
    train_file = training[ti]['file_name']
    if file_targets.get(actual_file, train_file) != train_file:
        return None, '同一上传文件不能映射到多个训练文件'
    if target_files.get(train_file, actual_file) != actual_file:
        return None, '多个上传文件不能覆盖同一训练文件'
    return {
        'training_id': ti,
        'actual_id': ai,
        'actual_sheet': actual_sheet or actual[ai].get('sheet_name'),
        'file_confidence': item.get('file_confidence'),
        'sheet_confidence': item.get('sheet_confidence'),
        'file_reason': str(item.get('file_reason') or ''),
        'sheet_reason': str(item.get('sheet_reason') or ''),
        'code_evidence': str(item.get('code_evidence') or ''),
    }, None


def _parse_stage1_to_pairs(entries, training, actual, fixed, matching_context):
    if not isinstance(entries, list):
        return []
    used_t, used_a = set(), set()
    file_targets, target_files = {}, {}
    pairs = []
    for item in entries:
        if not isinstance(item, dict):
            continue
        pair, reason = _stage1_entry_valid(item, training, actual, fixed, matching_context,
                                           used_t, used_a, file_targets, target_files)
        if pair is None:
            logger.warning('[源数据映射] 丢弃无效一阶段 AI 映射: %s', reason)
            continue
        used_t.add(pair['training_id'])
        used_a.add(pair['actual_id'])
        file_targets[actual[pair['actual_id']]['file_name']] = training[pair['training_id']]['file_name']
        target_files[training[pair['training_id']]['file_name']] = actual[pair['actual_id']]['file_name']
        pairs.append(pair)
    return pairs


def _build_stage2_prompt(pairs, training, actual, fixed, matching_context):
    matches = []
    for pair in pairs:
        ti, ai = pair['training_id'], pair['actual_id']
        matches.append({
            'training_id': ti,
            'actual_id': ai,
            'training_file': training[ti]['file_name'],
            'training_sheet': training[ti]['sheet_name'],
            'actual_file': actual[ai]['file_name'],
            'actual_sheet': actual[ai]['sheet_name'],
            'origin': actual[ai].get('source_origin', 'upload'),
            'asset_name': actual[ai].get('source_asset_name'),
            'training_columns': list(training[ti]['headers'].keys()),
            'actual_columns': list(actual[ai]['headers'].keys()),
            'fixed_mapping': dict(fixed.get((ti, ai)) or {}),
        })
    unresolved_training_ids = {pair['training_id'] for pair in pairs}
    unresolved_actual_ids = {pair['actual_id'] for pair in pairs}
    global_training = [
        {'training_id': ti, 'file': t['file_name'], 'sheet': t['sheet_name'],
         'columns': list(t['headers'].keys())}
        for ti, t in enumerate(training) if ti in unresolved_training_ids
    ]
    global_actual = [
        {'actual_id': ai, 'file': a['file_name'], 'original_file': a.get('original_file_name', a['file_name']),
         'origin': a.get('source_origin', 'upload'), 'asset_name': a.get('source_asset_name'),
         'sheet': a['sheet_name'], 'original_sheet': a.get('original_sheet_name', a['sheet_name']),
         'columns': list(a['headers'].keys())}
        for ai, a in enumerate(actual) if ai in unresolved_actual_ids
    ]
    return (
        '你是 Excel 字段匹配专家。文件关系和 Sheet 关系已经确定，不要修改。你的唯一任务：在每个已确认的 Sheet 对内部，把训练需要的列映射到上传文件的列。'
        '\n一、范围原则：只能在该 Sheet 对内部匹配列；不能跨 Sheet 拼凑列；上传列只能来自当前 actual_columns，训练列只能来自当前 training_columns；'
        '一个训练列最多对应一个上传列，一个上传列最多对应一个训练列；fixed_mapping 不可修改；找不到就放到 unresolved_columns，不要编造。'
        '\n二、匹配优先级：完全同名 confidence=1.0；去空格/大小写/全半角后同名 confidence>=0.98；同义名、业务别名结合代码用途、单位、join key、筛选条件判断；结构位置只能辅助；无法确定不要输出。'
        '注意区分应发/实发、个人/单位、上月/本月、金额/比例、税前/税后、收入/扣除。'
        '\n三、全局上下文：global_training 和 global_actual 只包含程序尚未明确的训练表与源表。'
        '已确定关系不参与本阶段，也不得重新判断；当前内容只用于排除剩余表之间的歧义，不能作为跨 Sheet 拼凑字段的理由。实际列匹配仍然只能使用当前 match 的 training_columns 和 actual_columns。'
        '\n四、输出要求：严格只输出 JSON，不要输出解释文字、Markdown、代码块标记。格式：'
        '{"mappings":[{"training_id":0,"actual_id":0,"columns":{"上传列名":"训练列名"},'
        '"column_confidence":{"上传列名":0.0-1.0},"column_reasons":{"上传列名":"匹配依据"},'
        '"unresolved_columns":[{"training_column":"训练列名","reason":"未找到可靠对应"}]}]}。'
        'columns 必须是对象，不能是数组；confidence 是 0~1，证据不足时给低分。'
        '\n输入数据：\n'
        + json.dumps({'matches': matches,
                      'global_training': global_training,
                      'global_actual': global_actual,
                      'script_context': _script_matching_context(matching_context, training, actual)},
                     ensure_ascii=False)
    )


def _parse_stage2_entries(entries, pairs, training, actual, fixed):
    by_pair = {}
    if isinstance(entries, list):
        for item in entries:
            if isinstance(item, dict):
                key = (item.get('training_id'), item.get('actual_id'))
                by_pair[key] = item
    out = []
    for pair in pairs:
        ti, ai = pair['training_id'], pair['actual_id']
        raw = by_pair.get((ti, ai)) or {}
        fixed_cols = dict(fixed.get((ti, ai)) or {})
        columns = raw.get('columns') if isinstance(raw.get('columns'), dict) else {}
        merged = {**fixed_cols, **columns}
        if len(set(merged.values())) != len(merged):
            logger.warning('[源数据映射] 丢弃重复二阶段列映射: training_id=%s actual_id=%s', ti, ai)
            merged = dict(fixed_cols)
        valid = {}
        for source, target in merged.items():
            if (isinstance(source, str) and isinstance(target, str)
                    and source in actual[ai]['headers'] and target in training[ti]['headers']):
                valid[source] = target
        conf = raw.get('column_confidence') if isinstance(raw.get('column_confidence'), dict) else {}
        reasons = raw.get('column_reasons') if isinstance(raw.get('column_reasons'), dict) else {}
        out.append({
            'training_id': ti,
            'actual_id': ai,
            'columns': valid,
            'column_confidence': conf,
            'column_reasons': reasons,
            'file_confidence': pair.get('file_confidence'),
            'sheet_confidence': pair.get('sheet_confidence'),
            'file_reason': pair.get('file_reason'),
            'sheet_reason': pair.get('sheet_reason'),
            'code_evidence': pair.get('code_evidence'),
        })
    return out


def _is_stage1_entries(entries):
    if not isinstance(entries, list) or not entries:
        return False
    for item in entries:
        if not isinstance(item, dict):
            return False
        if 'columns' in item:
            return False
        if 'file_confidence' in item or 'sheet_confidence' in item or 'actual_sheet' in item:
            return True
    return False


def _complete_two_stage_entries(entries, training, actual, fixed, matching_context, provider_name):
    pairs = _parse_stage1_to_pairs(entries, training, actual, fixed, matching_context)
    # 智算的 AI 仅负责文件/Sheet 身份判断。列关系不再发送给 AI，也不要求
    # 操作员确认；这里只保留代码能够百分百确定的同名列，供旧执行结构兼容。
    completed = []
    for pair in pairs:
        ti, ai = pair['training_id'], pair['actual_id']
        columns = dict(fixed.get((ti, ai)) or {})
        completed.append({
            **pair,
            'columns': columns,
            'column_confidence': {name: 1.0 for name in columns},
            'column_reasons': {},
        })
    logger.info('[源数据映射] AI 文件/Sheet 匹配完成 %s 对；已跳过 AI 列匹配', len(completed))
    return completed


def match_sources_with_ai(matcher, training, actual, provider_name, determined=None, matching_context=None):
    from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout

    locked = []
    for match in determined or []:
        ti = next((i for i, s in enumerate(training) if (s['file_name'], s['sheet_name']) ==
                   (match['train_file'], match['train_sheet'])), None)
        ai = next((i for i, s in enumerate(actual) if (s['file_name'], s['sheet_name']) ==
                   (match['input_file'], match['input_sheet'])), None)
        if ti is not None and ai is not None and set(match['col_mapping'].values()) == set(training[ti]['headers']):
            locked_columns = dict(match['col_mapping'])
            locked.append({'training_id': ti, 'actual_id': ai, 'columns': locked_columns,
                           'file_confidence': 1.0, 'sheet_confidence': 1.0,
                           'column_confidence': {source: 1.0 for source in locked_columns}})
    used_t = {m['training_id'] for m in locked}
    used_a = {m['actual_id'] for m in locked}
    file_targets = {actual[m['actual_id']]['file_name']: training[m['training_id']]['file_name'] for m in locked}
    target_files = {v: k for k, v in file_targets.items()}
    candidates, fixed = [], {}
    for ti, t in enumerate(training):
        if ti in used_t:
            continue
        choices = []
        for ai, a in enumerate(actual):
            if ai in used_a:
                continue
            if not matcher._is_candidate_allowed(t, a, actual):
                continue
            if not period_candidate_allowed(t, a, matching_context):
                continue
            if file_targets.get(a['file_name'], t['file_name']) != t['file_name']:
                continue
            if target_files.get(t['file_name'], a['file_name']) != a['file_name']:
                continue
            same = {c: c for c in a['headers'] if c in t['headers']}
            fixed[ti, ai] = same
            choices.append({'actual_id': ai, 'file': a['file_name'], 'sheet': a['sheet_name'],
                'original_file': a.get('original_file_name', a['file_name']),
                'original_sheet': a.get('original_sheet_name', a['sheet_name']),
                'origin': a.get('source_origin', 'upload'),
                'asset_name': a.get('source_asset_name'),
                'matched_column_count': len(same),
                'matched_columns': list(same),
                'missing_columns': [c for c in t['headers'] if c not in same],
                'candidate_columns': [c for c in a['headers'] if c not in same]})
        candidates.append({'training_id': ti, 'file': t['file_name'], 'sheet': t['sheet_name'],
                           'expected_columns': list(t['headers']), 'candidates': choices})

    stage1_prompt = _build_stage1_prompt(candidates, training, actual, matching_context)
    if os.getenv('LOG_AI_SOURCE_MATCH_PROMPT', 'false').strip().lower() in ('true', '1', 'yes', 'on'):
        logger.info('[源数据映射] AI 一阶段提示词完整内容 START\n%s\n[源数据映射] AI 一阶段提示词完整内容 END', stage1_prompt)
        logger.info('[源数据映射] AI 一阶段提示词长度: %s 字符', len(stage1_prompt))
    response = _call_ai_json(provider_name, stage1_prompt)
    entries = response.get('mappings') if isinstance(response, dict) else None
    if not isinstance(entries, list):
        raise ValueError('AI 映射格式无效')
    if _is_stage1_entries(entries):
        entries = _complete_two_stage_entries(entries, training, actual, fixed, matching_context, provider_name)
    # 逐条校验 AI 映射：无效条目只丢弃自己，不再让一条坏数据拖垮整轮 AI 匹配。
    valid_entries = []
    dropped_entries = []
    _used_t = set(used_t)
    _used_a = set(used_a)
    _file_targets = dict(file_targets)
    _target_files = dict(target_files)
    for entry in entries:
        try:
            if not isinstance(entry, dict):
                raise ValueError('AI 映射条目格式无效')
            ti, ai = entry.get('training_id'), entry.get('actual_id')
            if (type(ti) is not int or type(ai) is not int
                    or not 0 <= ti < len(training) or not 0 <= ai < len(actual)):
                raise ValueError('AI 修改了已确定映射或选择了无效候选')
            if not period_candidate_allowed(training[ti], actual[ai], matching_context):
                raise ValueError('AI 月份角色与工资年月冲突')
            if (ti, ai) not in fixed:
                raise ValueError('AI 修改了已确定映射或选择了无效候选')
            if ti in _used_t or ai in _used_a:
                raise ValueError('AI 表映射重复或引用不存在的表')
            # 无论供应商是否额外返回 columns，都不采纳 AI 列语义判断。
            columns = dict(fixed[ti, ai])
            if (any(not isinstance(k, str) for k in columns)
                    or any(not isinstance(v, str) for v in columns.values())
                    or not set(columns).issubset(actual[ai]['headers'])
                    or not set(columns.values()).issubset(training[ti]['headers'])):
                raise ValueError('AI 列映射引用不存在的字段')
            merged_columns = columns
            if len(set(merged_columns.values())) != len(merged_columns):
                raise ValueError('AI 列映射与原有列重名')
            if not matcher._is_candidate_allowed(training[ti], actual[ai], actual):
                raise ValueError('AI 映射不能覆盖人工确认的文件及月份分表关系')
            if not period_candidate_allowed(training[ti], actual[ai], matching_context):
                raise ValueError('AI 月份角色与工资年月冲突')
            actual_file = actual[ai]['file_name']
            train_file = training[ti]['file_name']
            if _file_targets.get(actual_file, train_file) != train_file:
                raise ValueError('同一上传文件不能映射到多个训练文件')
            if _target_files.get(train_file, actual_file) != actual_file:
                raise ValueError('多个上传文件不能覆盖同一训练文件')
            entry['columns'] = merged_columns
            valid_entries.append(entry)
            _used_t.add(ti)
            _used_a.add(ai)
            _file_targets[actual_file] = train_file
            _target_files[train_file] = actual_file
        except ValueError as exc:
            dropped_entries.append({
                'training_id': entry.get('training_id'),
                'actual_id': entry.get('actual_id'),
                'reason': str(exc),
            })
            continue
    if dropped_entries:
        logger.warning('[源数据映射] 丢弃 %s 条无效 AI 映射，保留 %s 条有效映射: %s',
                       len(dropped_entries), len(valid_entries),
                       json.dumps(dropped_entries, ensure_ascii=False, default=str)[:2000])
    entries = valid_entries
    result = validate_mapping(matcher, training, actual, {'mappings': locked + entries}, allow_partial=True)
    def sheet_reason(entry):
        parts = [f'{label}：{entry[key]}' for key, label in
                 [('file_reason', '文件'), ('sheet_reason', 'Sheet'), ('code_evidence', '代码')]
                 if isinstance(entry.get(key), str) and entry[key].strip()]
        return ('；'.join(parts) or str(entry.get('reason') or
                'AI 结合文件、Sheet、字段语义与脚本逻辑推荐，请确认'))[:700]

    def column_reason(entry, source):
        reasons = entry.get('column_reasons')
        detail = reasons.get(source) if isinstance(reasons, dict) else None
        return detail[:500] if isinstance(detail, str) and detail.strip() else sheet_reason(entry)

    result['ai_suggestions'] = []
    result['source_sheet_reviews'] = []
    for e in entries:
        confidence = e.get('sheet_confidence', e.get('file_confidence'))
        try:
            confidence = float(confidence) if confidence is not None else None
        except (TypeError, ValueError):
            confidence = None
        result['source_sheet_reviews'].append({
            'expected_file': training[e['training_id']]['file_name'],
            'expected_sheet': training[e['training_id']]['sheet_name'],
            'suggested_file': actual[e['actual_id']]['file_name'],
            'suggested_sheet': actual[e['actual_id']]['sheet_name'],
            'confidence': confidence,
            'reason': sheet_reason(e),
            'recommendation_source': 'ai',
        })
    return result


def validate_mapping(matcher, training, actual, response, allow_partial=False):
    entries = response.get('mappings') if isinstance(response, dict) else None
    if not isinstance(entries, list) or not training or (not allow_partial and len(entries) != len(training)):
        raise ValueError('AI 映射未覆盖全部训练表')
    used_training, used_actual, matches = set(), set(), []
    file_targets, target_files = {}, {}
    for item in entries:
        ti, ai = item.get('training_id'), item.get('actual_id')
        if (type(ti) is not int or type(ai) is not int or not 0 <= ti < len(training)
                or not 0 <= ai < len(actual) or ti in used_training or ai in used_actual):
            raise ValueError('AI 表映射重复或引用不存在的表')
        t, a = training[ti], actual[ai]
        if not matcher._is_candidate_allowed(t, a, actual):
            raise ValueError('AI 映射不能覆盖人工确认的文件及月份分表关系')
        columns = item.get('columns')
        if (not isinstance(columns, dict) or (not columns and not allow_partial)
                or any(not isinstance(v, str) for v in columns.values())
                or not set(columns).issubset(a['headers'])
                or not set(columns.values()).issubset(t['headers'])
                or (not allow_partial and set(columns.values()) != set(t['headers']))
                or len(set(columns.values())) != len(columns)):
            raise ValueError('AI 列映射缺失、重复或引用不存在的字段')
        def _conf(value):
            if value is None:
                return None
            try:
                number = float(value)
            except (TypeError, ValueError):
                return None
            return number if 0.0 <= number <= 1.0 else None
        file_confidence = _conf(item.get('file_confidence'))
        sheet_confidence = _conf(item.get('sheet_confidence'))
        raw_column_confidence = item.get('column_confidence')
        column_confidence = {}
        for source_col, target_col in columns.items():
            if raw_column_confidence is not None and not isinstance(raw_column_confidence, dict):
                raise ValueError('AI 列置信度格式无效')
            column_confidence[source_col] = _conf((raw_column_confidence or {}).get(source_col))
        # 文件关系保持唯一；字段关系按 Sheet 隔离，不跨业务表混用。
        name = a['file_name']
        if file_targets.setdefault(name, t['file_name']) != t['file_name']:
            raise ValueError('同一上传文件不能映射到多个训练文件')
        if target_files.setdefault(t['file_name'], name) != name:
            raise ValueError('多个上传文件不能覆盖同一训练文件')
        renamed = [columns.get(h, h) for h in a['headers']]
        if len(set(renamed)) != len(renamed):
            raise ValueError('AI 列映射与原有列重名')
        used_training.add(ti)
        used_actual.add(ai)
        matches.append({'train_file': t['file_name'], 'train_sheet': t['sheet_name'],
                        'input_file': name, 'input_file_path': a['file_path'],
                        'input_sheet': a['sheet_name'], 'col_mapping': columns,
                        'column_confidence': column_confidence,
                        'sheet_confidence': sheet_confidence,
                        'file_confidence': file_confidence,
                        'needs_rewrite': True})
    # 新链路的完整性只检查训练文件/Sheet 是否全部得到唯一来源；列覆盖率不再
    # 决定智算能否启动。
    complete = len(entries) == len(training)
    result = {'success': complete, 'mapping': {'file_mapping': matcher._build_file_mapping(matches)}}
    if not complete:
        result['error'] = 'AI 仅确定部分文件或 Sheet，已保留有效推荐，其余需要确认'
    return result
