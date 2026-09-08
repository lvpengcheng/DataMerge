"""生成脚本共用的多表合并：保留 schema，显式验证关联基数。"""
import json

import pandas as pd

from backend.utils.data_helpers import _schema_text_value


def merge_source_tables(source_data, merge_config):
    if isinstance(merge_config, str):
        merge_config = json.loads(merge_config)
    merged = dict(source_data)
    output_key = merge_config.get('output_key', 'derived_main_table')
    groups = merge_config.get('vertical_groups', [])
    if len(groups) > 1:
        raise ValueError('多个纵向合并组需要分别配置输出表，不能覆盖同一个 output_key')
    for group in groups:
        if not group or any(k not in merged for k in group):
            raise ValueError(f'纵向合并源表缺失: {group}')
        entries = [merged[k] for k in group]
        schemas, formats = {}, {}
        for entry in entries:
            for name, schema in (entry.get('column_schemas') or {}).items():
                old_kind = (schemas.get(name) or {}).get('field_type')
                new_kind = schema.get('field_type')
                if old_kind and new_kind and old_kind != new_kind:
                    raise ValueError(f'纵向合并字段类型不一致: {name} ({old_kind}/{new_kind})')
                schemas[name] = schema
            formats.update(entry.get('column_formats') or {})
        combined = pd.concat([e['df'] for e in entries], ignore_index=True, sort=False)
        merged[output_key] = {**entries[0], 'df': combined, 'columns': list(combined.columns),
                              'column_schemas': schemas, 'column_formats': formats}
    for spec in merge_config.get('horizontal_joins', []):
        left_key, right_key = spec['left'], spec['right']
        if left_key not in merged or right_key not in merged:
            raise ValueError(f'关联源表不存在: {left_key}, {right_key}')
        keys = spec['on']
        keys = [keys] if isinstance(keys, str) else list(keys)
        if not keys:
            raise ValueError('关联键不能为空')
        left, right = merged[left_key], merged[right_key]
        for entry, name in ((left, left_key), (right, right_key)):
            if not entry['df'].columns.is_unique:
                raise ValueError(f'关联表存在重复列名: {name}')
            missing = [key for key in keys if key not in entry['df'].columns]
            if missing:
                raise ValueError(f'关联表 {name} 缺少键: {missing}')
        # 不覆盖左表原始主键值；只在临时 key 数组中统一数字/文本类型。
        left_keys = [left['df'][k].map(_schema_text_value).replace('', None) for k in keys]
        right_keys = pd.DataFrame({k: right['df'][k].map(_schema_text_value).replace('', None)
                                   for k in keys})
        valid = right_keys.notna().all(axis=1)
        duplicate = right_keys.loc[valid].duplicated(keep=False)
        if duplicate.any():
            raise ValueError(f'关联表 {right_key} 的键 {keys} 存在重复；请配置复合键或先按业务规则汇总')
        added = [c for c in right['df'].columns if c not in left['df'].columns]
        lookup = right['df'].loc[valid, added].copy()
        lookup.index = pd.MultiIndex.from_frame(right_keys.loc[valid, keys])
        requested = pd.MultiIndex.from_arrays(left_keys, names=keys)
        matched = lookup.reindex(requested)
        result = left['df'].copy()
        for name in added:
            result[name] = matched[name].to_numpy()
        # reindex preserves left order/count, and missing keys never match blank rows.
        metadata = {}
        for field in ('column_schemas', 'column_formats'):
            metadata[field] = dict(left.get(field) or {})
            metadata[field].update({c: right[field][c] for c in added if c in (right.get(field) or {})})
        merged[left_key] = {**left, **metadata, 'df': result, 'columns': list(result.columns)}
    return merged
