"""Conservative schema matching; filenames are tie breakers, never substitutes for columns."""
import re
import unicodedata
from .period_matching import period_candidate_allowed


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
                rank = (a['file_name'] == t['file_name'], a['sheet_name'] == t['sheet_name'])
                choices.append((rank, ai, columns))
            if not choices:
                continue
            best = max(item[0] for item in choices)
            winners = [item for item in choices if item[0] == best]
            if len(winners) == 1:
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
