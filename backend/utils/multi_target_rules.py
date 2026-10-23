"""Portable, readable multi-target rules with explicit workbook boundaries."""
import copy
import json
import re
from pathlib import Path

RULES_CONTENT_LIMIT = 200000
START = '<!-- DM_RULES BUNDLE START -->'
END = '<!-- DM_RULES BUNDLE END -->'
MANIFEST = '<!-- DM_RULES MANIFEST '


def limit_rules_content(content):
    text = str(content or '')
    if len(text) > RULES_CONTENT_LIMIT and START in text:
        raise ValueError('多目标规则超过20万字符，请精简规则或减少本次目标文件，不能截断部分目标规则')
    return text[:RULES_CONTENT_LIMIT]


def json_response(raw):
    text = re.sub(r'^```(?:json)?\s*|\s*```$', '', str(raw or '').strip())
    result = json.loads(text)
    if not isinstance(result, dict):
        raise ValueError('规则分析结果必须是JSON对象')
    return result


def validate_target(target):
    if not isinstance(target.get('rules'), str) or not target['rules'].strip():
        raise ValueError(f'{target.get("file_name")}: 目标规则为空')
    if '<!-- DM_RULES' in target['rules']:
        raise ValueError('目标规则中不能包含内部边界标记')
    required = {field['id']: field for sheet in target['sheets'] for field in sheet['fields']}
    coverage = target.get('covered_field_ids')
    if (not isinstance(coverage, list) or any(not isinstance(item, str) for item in coverage)
            or len(coverage) != len(required) or set(coverage) != set(required)):
        raise ValueError(f'{target["file_name"]}: 字段覆盖不完整或出现未知字段')
    missing = [field['name'] for field in required.values() if field['name'] not in target['rules']]
    missing += [sheet['name'] for sheet in target['sheets'] if sheet['name'] not in target['rules']]
    if missing:
        raise ValueError(f'{target["file_name"]}: 规则正文缺少字段或Sheet：' + '、'.join(missing[:20]))
    if not isinstance(target.get('external_output_dependencies', []), list):
        raise ValueError('跨结果依赖必须是目标ID列表')


def validate_bundle(bundle):
    targets = bundle.get('targets')
    if bundle.get('version') != 1 or not isinstance(targets, list) or len(targets) < 2:
        raise ValueError('多目标规则格式无效')
    ids, names = set(), set()
    for target in targets:
        if target['id'] in ids or target['file_name'].casefold() in names:
            raise ValueError('多目标规则存在重复目标')
        ids.add(target['id'])
        names.add(target['file_name'].casefold())
        validate_target(target)
    for target in targets:
        if any(not isinstance(dep, str) or dep not in ids or dep == target['id']
               for dep in target.get('external_output_dependencies', [])):
            raise ValueError('多目标规则引用了未知目标或自身作为生成依赖')
    graph = {target['id']: target.get('external_output_dependencies', []) for target in targets}
    done, visiting = set(), set()
    def visit(node):
        if node in visiting:
            raise ValueError('目标文件之间存在循环依赖，请调整规则')
        if node in done:
            return
        visiting.add(node)
        for child in graph[node]:
            visit(child)
        visiting.remove(node)
        done.add(node)
    for node in graph:
        visit(node)
    for key in ('common_rules', 'cross_target_checks', 'source_format'):
        if not isinstance(bundle.get(key, ''), str) or '<!-- DM_RULES' in bundle.get(key, ''):
            raise ValueError('公共规则或核对规则格式无效')
    return bundle


def render_bundle(bundle):
    validate_bundle(bundle)
    manifest = {key: value for key, value in bundle.items() if key not in ('common_rules', 'cross_target_checks', 'source_format', 'targets', 'supplemental')}
    manifest['targets'] = [{key: value for key, value in target.items() if key != 'rules'} for target in bundle['targets']]
    meta = json.dumps(manifest, ensure_ascii=False).replace('<', '\\u003c').replace('>', '\\u003e')
    def block(name, text):
        return f'<!-- DM_RULES {name} START -->\n{text.strip()}\n<!-- DM_RULES {name} END -->'
    parts = [START, '# 数据处理规则', '## 基本信息\n- 输出目标：' + '、'.join(t['file_name'] for t in bundle['targets']),
             MANIFEST + meta + ' -->', '## 公共规则', block('COMMON', bundle['common_rules'])]
    for target in bundle['targets']:
        parts.extend(['## 目标文件：' + target['file_name'], block(target['id'], target['rules'])])
    parts.extend(['## 跨目标核对规则', block('CHECKS', bundle.get('cross_target_checks', '')),
                  '## 源数据格式', block('SOURCE', bundle.get('source_format', '')), END])
    return limit_rules_content((bundle.get('supplemental', '') + '\n\n' + '\n\n'.join(parts)).strip())


def parse_bundle(content):
    text = str(content or '')
    if START not in text:
        return None
    if text.count(START) != 1 or text.count(END) != 1 or text.index(END) < text.index(START):
        raise ValueError('多目标规则边界不完整，请重新下载完整rules.md')
    document = text[text.index(START):text.index(END) + len(END)]
    metadata = re.findall(re.escape(MANIFEST) + r'(.*?) -->', document, re.S)
    if len(metadata) != 1:
        raise ValueError('多目标规则清单缺失或重复')
    bundle = json.loads(metadata[0])
    if not isinstance(bundle, dict):
        raise ValueError('多目标规则清单无效')
    def block(name):
        matches = re.findall(re.escape(f'<!-- DM_RULES {name} START -->') + r'\s*(.*?)\s*' +
                             re.escape(f'<!-- DM_RULES {name} END -->'), document, re.S)
        if len(matches) != 1:
            raise ValueError(f'多目标规则的{name}段落缺失或重复')
        return matches[0]
    bundle.update(common_rules=block('COMMON'), cross_target_checks=block('CHECKS'), source_format=block('SOURCE'))
    for target in bundle.get('targets', []):
        if not re.fullmatch(r'T\d+', str(target.get('id', ''))):
            raise ValueError('目标规则ID无效')
        target['rules'] = block(target['id'])
    bundle['supplemental'] = (text[:text.index(START)] + '\n' + text[text.index(END) + len(END):]).strip()
    return validate_bundle(bundle)


def rules_for_workbook(content, name):
    bundle = parse_bundle(content)
    if not bundle:
        return content or ''
    normalized = Path(name).with_suffix('.xlsx').name.casefold()
    matches = [t for t in bundle['targets'] if Path(t['file_name']).with_suffix('.xlsx').name.casefold() == normalized]
    if len(matches) != 1:
        raise ValueError(f'多目标规则中没有唯一对应的目标文件：{name}，请保持目标名称一致或重新整理对应关系')
    target = matches[0]
    if target.get('external_output_dependencies'):
        raise ValueError(f'{name}需要其他结果文件作为输入，当前仅支持独立目标，请先改为从源数据计算或提供上游结果作为源文件')
    return '\n\n'.join(part for part in [bundle.get('supplemental', ''), '## 公共规则\n' + bundle['common_rules'],
        f'## 本次目标文件：{name}', target['rules'], '## 源数据格式\n' + bundle.get('source_format', '')] if part)


def rules_for_structure(content, structure, fallback_name=None):
    if START not in str(content or ''):
        return content
    info = json.loads(structure) if isinstance(structure, str) else (structure or {})
    name = info.get('file_name') or fallback_name
    if not name:
        raise ValueError('多目标规则需要明确当前目标文件名，不能按第一个目标默认套用')
    return rules_for_workbook(content, name)


def patch_bundle(bundle, patch):
    result = copy.deepcopy(bundle)
    targets = {target['id']: target for target in result['targets']}
    changed = set()
    updates = patch.get('targets', [])
    if not isinstance(updates, list) or any(not isinstance(update, dict) for update in updates):
        raise ValueError('目标规则修改必须是对象列表')
    for update in updates:
        tid = update.get('id')
        if tid not in targets or tid in changed:
            raise ValueError('修改引用了未知或重复的目标文件')
        changed.add(tid)
        for key in ('rules', 'covered_field_ids', 'external_output_dependencies'):
            if key in update:
                targets[tid][key] = update[key]
    for key in ('common_rules', 'cross_target_checks'):
        if key in patch:
            result[key] = patch[key]
    validate_bundle(result)
    return result
