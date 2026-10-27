"""Lightweight solution metadata for tenant selection; execution still verifies packages."""
import json
from .store import artifact


def tenant_versions(rows):
    versions, pending, unavailable = [], [], []
    for row in rows:
        identity = {'id': row.id, 'name': row.name, 'version': row.digest[:12]}
        if row.status == 'candidate':
            pending.append({**identity, 'validation_passed': bool((row.validation or {}).get('passed'))})
            continue
        if row.status != 'published' or not row.is_active:
            continue
        try:
            manifest = json.loads((artifact(row.digest) / 'manifest.json').read_text(encoding='utf-8'))
            if not isinstance(manifest, dict) or manifest.get('digest') != row.digest:
                raise ValueError('方案清单与版本不一致')
        except (OSError, ValueError, UnicodeError):
            unavailable.append({**identity, 'reason': '方案文件缺失或清单无效，请恢复完整方案包或重新导入原方案'})
            continue
        versions.append({'script_id': 'solution_' + row.id,
                         'name': row.name + ' · 方案 ' + row.digest[:12],
                         'is_active': not versions, 'solution_id': row.id,
                         'solution_digest': row.digest,
                         'score': (row.validation or {}).get('accuracy', 1.0),
                         'mode': manifest.get('mode')})
    return versions, pending, unavailable
