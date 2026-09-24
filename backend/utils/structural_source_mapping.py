"""Conservative schema matching; filenames are tie breakers, never substitutes for columns."""
import re
import unicodedata
from difflib import SequenceMatcher
from .period_matching import period_candidate_allowed


_EXCEL_SUFFIX = re.compile(r'\.(?:xlsx|xlsm|xls)$', re.IGNORECASE)
_PERIOD_TOKEN = re.compile(
    r'(?<!\d)20\d{2}(?:\s*[年./_-]?\s*)(?:0?[1-9]|1[0-2])\s*月?(?!\d)')
_IDENTITY_NOISE = re.compile(r'[\s\-_.·—（）()【】\[\]]+')
_IDENTITY_LEAD = 0.08


def normalized_source_identity(value, *, filename=False):
    """用于文件/Sheet 业务名称关联；移除扩展名、工资月份和分隔符。"""
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().strip()
    if filename:
        text = _EXCEL_SUFFIX.sub('', text)
    text = _PERIOD_TOKEN.sub('', text)
    return _IDENTITY_NOISE.sub('', text)


def raw_source_identity(value, *, filename=False):
    """只做 Unicode/大小写/扩展名规范化；用于识别真正的完全同名。"""
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().strip()
    if filename:
        text = _EXCEL_SUFFIX.sub('', text).strip()
    return text


def source_identity_similarity(left, right, *, filename=False):
    # 完全同名必须具有最高优先级。不能先删除月份再比较：纯月份文件名
    # （202607.xlsx / 202608.xlsx）删除月份后都会变成空串，导致本来完全
    # 相同的智训/智算文件被误判为“同结构歧义”。
    raw_left = raw_source_identity(left, filename=filename)
    raw_right = raw_source_identity(right, filename=filename)
    if raw_left and raw_left == raw_right:
        return 1.0
    left = normalized_source_identity(left, filename=filename)
    right = normalized_source_identity(right, filename=filename)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    score = SequenceMatcher(None, left, right).ratio()
    # 很弱的偶然字符重合不能用来拆解同结构歧义。
    return round(score, 6) if score >= 0.45 else 0.0


def source_identity_rank(training, actual):
    """结构同分时：先比较业务文件名，再比较 Sheet 名。"""
    file_score = max(
        source_identity_similarity(training.get('file_name'), actual.get(name), filename=True)
        for name in ('file_name', 'original_file_name') if actual.get(name)
    )
    sheet_score = max(
        source_identity_similarity(training.get('sheet_name'), actual.get(name))
        for name in ('sheet_name', 'original_sheet_name') if actual.get(name)
    )
    return file_score, sheet_score


def ranked_choice_is_unique(ranks):
    """结构分相同时，名称关联必须有明确领先，不能因微小相似度误配。"""
    ranks = sorted(ranks, reverse=True)
    if len(ranks) <= 1:
        return True
    top, second = ranks[0], ranks[1]
    identity_offset = 0
    if len(top) >= 3:
        if top[0] != second[0]:
            return True
        identity_offset = 1
    if top[identity_offset] != second[identity_offset]:
        return top[identity_offset] - second[identity_offset] >= _IDENTITY_LEAD
    if len(top) > identity_offset + 1 and top[identity_offset + 1] != second[identity_offset + 1]:
        return top[identity_offset + 1] - second[identity_offset + 1] >= _IDENTITY_LEAD
    return False


def suggest_file_relations(training, actual, resolved, year=None, month=None):
    """Filename candidates are metadata only; never rename uploads or assume 1:1 means correct."""
    from .source_auto_filler import _filename_similarity, _period_index, _period_role
    expected, uploaded = {}, {}
    for sheet in training:
        expected.setdefault(sheet['file_name'], []).append(sheet)
    for sheet in actual:
        uploaded.setdefault(sheet['file_name'], []).append(sheet)
    covered = {info['expected_file'] for info in resolved.values()}
    missing = set(expected) - covered - set(uploaded)
    extras = set(uploaded) - set(resolved) - set(expected)
    plans = []
    if year and month and 1 <= int(month) <= 12:
        current = int(year) * 12 + int(month) - 1
        for role, period in [('current', current), ('previous', current - 1)]:
            targets = [f for f in missing if _period_role(f) == role]
            sources = [f for f in extras if _period_index(f) == period]
            if len(targets) == len(sources) == 1:
                plans.append({'from': sources[0], 'to': targets[0], 'decision': 'period_role',
                              'reason': '按指定工资年月识别文件角色，保留原始上传名称', 'score': 1.0})
                missing.remove(targets[0])
                extras.remove(sources[0])
    candidates = []
    for filename in sorted(extras):
        source_cols = {normalized_header(c) for s in uploaded[filename] for c in s['headers']}
        options = []
        for target in sorted(missing):
            target_cols = {normalized_header(c) for s in expected[target] for c in s['headers']}
            union = source_cols | target_cols
            overlap = len(source_cols & target_cols) / len(union) if union else 0
            similarity = _filename_similarity(filename, target)
            options.append({'expected': target, 'score': round(.7 * overlap + .3 * similarity, 3),
                            'header_jaccard': round(overlap, 3), 'name_similarity': round(similarity, 3)})
        if options:
            candidates.append({'uploaded': filename, 'candidates': sorted(options, key=lambda c: -c['score']),
                               'uploaded_sheets': [{'name': s['sheet_name'], 'headers': list(s['headers'])}
                                                   for s in uploaded[filename]]})
    return plans, candidates


def normalized_header(value):
    # Formatting-only differences. Do not remove units, months or business words.
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value))).casefold()


def structural_columns(matcher, training, actual):
    expected = [h for h in training if matcher._is_valid_header(h)]
    index = {}
    for column in actual:
        index.setdefault(normalized_header(column), []).append(column)
    mapping = {}
    for column in expected:
        candidates = index.get(normalized_header(column), [])
        if len(candidates) != 1:
            return None
        source = candidates[0]
        if source in mapping:
            return None
        mapping[source] = column
    # Extra actual columns are allowed, but cannot collide after normalization.
    renamed = [mapping.get(h, h) for h in actual]
    if not mapping or len(set(renamed)) != len(renamed):
        return None
    return mapping


def match_structural_sources(matcher, training, actual, matching_context=None):
    """Only lock unique, complete schema matches, retaining unresolved tables for AI."""
    pending = set(range(len(training)))
    used = set()
    file_targets, target_files, matches = {}, {}, []

    # 文件名 + Sheet 名完全相同是用户定义的硬关系。即使本次表头被解析到
    # 不同区域或列有增减，也不应再进入 AI/人工“来源”审核；列差异由训练脚本
    # 在确定的来源内处理。这里只锁唯一的完全同名项，绝不靠列表顺序配对。
    exact_actual = {}
    for ai, item in enumerate(actual):
        key = (raw_source_identity(item.get('file_name'), filename=True),
               raw_source_identity(item.get('sheet_name')))
        exact_actual.setdefault(key, []).append(ai)
    for ti in sorted(tuple(pending)):
        item = training[ti]
        key = (raw_source_identity(item.get('file_name'), filename=True),
               raw_source_identity(item.get('sheet_name')))
        candidates = [ai for ai in exact_actual.get(key, []) if ai not in used]
        if not all(key) or len(candidates) != 1:
            continue
        ai = candidates[0]
        source = actual[ai]
        if file_targets.get(source['file_name'], item['file_name']) != item['file_name']:
            continue
        if target_files.get(item['file_name'], source['file_name']) != source['file_name']:
            continue
        columns = structural_columns(matcher, item.get('headers') or {}, source.get('headers') or {})
        if columns is None:
            # 来源身份已经确定时只保留同名列，不做语义列匹配。
            source_by_normalized = {}
            for name in (source.get('headers') or {}):
                source_by_normalized.setdefault(normalized_header(name), []).append(name)
            columns = {}
            for target_name in (item.get('headers') or {}):
                options = source_by_normalized.get(normalized_header(target_name), [])
                if len(options) == 1:
                    columns[options[0]] = target_name
        file_targets[source['file_name']] = item['file_name']
        target_files[item['file_name']] = source['file_name']
        used.add(ai)
        pending.remove(ti)
        matches.append({
            'train_file': item['file_name'], 'train_sheet': item['sheet_name'],
            'input_file': source['file_name'], 'input_file_path': source['file_path'],
            'input_sheet': source['sheet_name'], 'col_mapping': columns,
            'column_confidence': {name: 1.0 for name in columns},
            'sheet_confidence': 1.0, 'file_confidence': 1.0,
            'needs_rewrite': not matcher._is_fully_identical(item, source, columns),
            'identity_exact': True,
        })

    while pending:
        proposals = {}
        for ti in sorted(pending):
            t = training[ti]
            choices = []
            for ai, a in enumerate(actual):
                if ai in used or not matcher._is_candidate_allowed(t, a, actual):
                    continue
                if not period_candidate_allowed(t, a, matching_context):
                    continue
                if file_targets.get(a['file_name'], t['file_name']) != t['file_name']:
                    continue
                if target_files.get(t['file_name'], a['file_name']) != a['file_name']:
                    continue
                columns = structural_columns(matcher, t['headers'], a['headers'])
                if columns is None:
                    continue
                rank = source_identity_rank(t, a)
                choices.append((rank, ai, columns))
            if not choices:
                continue
            best = max(item[0] for item in choices)
            winners = [item for item in choices if item[0] == best]
            if len(winners) == 1 and ranked_choice_is_unique([item[0] for item in choices]):
                proposals[ti] = winners[0]
        accepted = []
        for ti, (rank, ai, columns) in proposals.items():
            # Two expected sheets competing for the same actual sheet are ambiguous.
            rivals = [item for item in proposals.values() if item[1] == ai]
            if len(rivals) != 1:
                continue
            t, a = training[ti], actual[ai]
            if any((actual[other_ai]['file_name'] == a['file_name'] and
                    training[other_ti]['file_name'] != t['file_name']) or
                   (training[other_ti]['file_name'] == t['file_name'] and
                    actual[other_ai]['file_name'] != a['file_name'])
                   for other_ti, (_, other_ai, _) in proposals.items()):
                continue
            if file_targets.get(a['file_name'], t['file_name']) != t['file_name']:
                continue
            if target_files.get(t['file_name'], a['file_name']) != a['file_name']:
                continue
            file_targets[a['file_name']] = t['file_name']
            target_files[t['file_name']] = a['file_name']
            used.add(ai)
            accepted.append(ti)
            matches.append({'train_file': t['file_name'], 'train_sheet': t['sheet_name'],
                            'input_file': a['file_name'], 'input_file_path': a['file_path'],
                            'input_sheet': a['sheet_name'], 'col_mapping': columns,
                            'column_confidence': {source: 1.0 for source in columns},
                            'sheet_confidence': 1.0, 'file_confidence': 1.0,
                            'needs_rewrite': not matcher._is_fully_identical(t, a, columns)})
        if not accepted:
            break
        pending.difference_update(accepted)
    # Preserve additional instances in an already resolved workbook (e.g. new
    # employee/month sheets), as the existing preload contract requires.
    if not pending:
        for ai, a in enumerate(actual):
            if ai in used or a['file_name'] not in file_targets:
                continue
            target_file = file_targets[a['file_name']]
            candidates = [t for t in training if t['file_name'] == target_file
                          and matcher._is_candidate_allowed(t, a, actual)
                          and period_candidate_allowed(t, a, matching_context)]
            columns = [structural_columns(matcher, t['headers'], a['headers']) for t in candidates
                       if len(t['headers']) >= 3]
            columns = [c for c in columns if c]
            if not columns or any(c != columns[0] for c in columns):
                continue
            if any(m['train_file'] == target_file and m['train_sheet'] == a['sheet_name'] for m in matches):
                continue
            matches.append({'train_file': target_file, 'train_sheet': a['sheet_name'],
                            'input_file': a['file_name'], 'input_file_path': a['file_path'],
                            'input_sheet': a['sheet_name'], 'col_mapping': columns[0],
                            'column_confidence': {source: 1.0 for source in columns[0]},
                            'sheet_confidence': 1.0, 'file_confidence': 1.0,
                            'needs_rewrite': any(k != v for k, v in columns[0].items())})
    result = {'success': not pending, 'determined': matches, 'match_method': 'structure',
              'mapping': {'file_mapping': matcher._build_file_mapping(matches)}}
    if pending:
        result['error'] = '以下来源结构不完整或存在多个候选，需要语义推荐或人工确认: ' + '；'.join(
            f"{training[i]['file_name']}/{training[i]['sheet_name']}" for i in sorted(pending))
    return result
