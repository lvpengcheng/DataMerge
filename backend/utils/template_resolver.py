"""模板路径解析器：让脚本不依赖烘焙的绝对路径，按当前环境重新定位模板。

跨机器 / 跨 session 复用脚本时，脚本里烘焙的训练机绝对路径 TEMPLATE_PATH 往往不存在。
本解析器按优先级链算出当前环境下的**有效模板路径**，供各执行入口注入 `_template_override_path`：

  1. 用户上传的覆盖模板（uploaded_override）
  2. 随脚本绑定的模板副本（bundled_path，阶段2产出）
  3. 当前租户目录里按 TEMPLATE_NAME 命中（有 TEMPLATE_HASH 则优先校验一致）
  4. 全局资源目录 global_assets 里按名命中
  5. 烘焙的 TEMPLATE_PATH（仅当本机存在）
  6. 都没有 → 返回 None，由调用方报清晰错误

脚本里由生成器烘焙 TEMPLATE_NAME / TEMPLATE_HASH / TEMPLATE_PATH 三个逻辑引用（老脚本可能只有 TEMPLATE_PATH，此时用其文件名兜底）。
"""

import os
import re
import hashlib
import logging
import ntpath
import posixpath
import ast

logger = logging.getLogger(__name__)

_STR_PREFIX = r"(?:[rRuUbB]{0,2})?"
_NAME_RE = re.compile(r"TEMPLATE_NAME\s*=\s*" + _STR_PREFIX + r"(['\"])(.+?)\1")
_HASH_RE = re.compile(r"TEMPLATE_HASH\s*=\s*" + _STR_PREFIX + r"(['\"])([0-9a-fA-F]*)\1")
_PATH_RE = re.compile(r"TEMPLATE_PATH\s*=\s*" + _STR_PREFIX + r"(['\"])(.+?)\1")


def _md5(path):
    try:
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def extract_template_ref(script_code):
    """从脚本源码抠出 (TEMPLATE_NAME, TEMPLATE_HASH, TEMPLATE_PATH)；缺名则用路径文件名兜底。"""
    name = hsh = baked = None
    if script_code:
        m = _NAME_RE.search(script_code); name = m.group(2) if m else None
        m = _HASH_RE.search(script_code); hsh = m.group(2) if m else None
        m = _PATH_RE.search(script_code); baked = m.group(2) if m else None
    if not name and baked:
        # os.path.basename 只理解当前操作系统的分隔符：Linux 处理
        # ``E:\\...\\模板.xlsx`` 时会把整条路径当文件名。迁移场景必须同时兼容
        # Windows 与 POSIX 路径。
        name = portable_basename(baked)
    return name, hsh, baked


def portable_basename(path):
    """跨 Windows/POSIX 提取文件名，不依赖当前运行平台。"""
    value = str(path or "").strip().rstrip("/\\")
    if not value:
        return ""
    return ntpath.basename(posixpath.basename(value))


def find_nonportable_absolute_paths(script_code):
    """返回除 TEMPLATE_PATH 外的 Windows/Docker/Ubuntu 环境绝对路径。"""
    try:
        tree = ast.parse(script_code or "")
    except Exception:
        return []
    template_literals = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            value = node.value
            if (any(isinstance(t, ast.Name) and t.id == "TEMPLATE_PATH" for t in targets)
                    and isinstance(value, ast.Constant) and isinstance(value.value, str)):
                template_literals.add(value.value)
    found = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Constant) and isinstance(node.value, str)):
            continue
        value = node.value.strip()
        if not value or value in template_literals:
            continue
        if re.match(r"^[A-Za-z]:[\\/]", value) or value.startswith(("/app/", "/www/")):
            found.append(value)
    return list(dict.fromkeys(found))


def _find_by_name(root, name, want_hash=None):
    """在 root 下递归找文件；指定哈希时只允许返回哈希一致的文件。"""
    if not root or not name or not os.path.isdir(root):
        return None
    for dirpath, _dirs, files in os.walk(root):
        if name in files:
            full = os.path.join(dirpath, name)
            if want_hash:
                if _md5(full) == want_hash:
                    return full
            else:
                return full
    return None


def resolve_template_path(*, tenant_id=None, script_code=None, uploaded_override=None,
                          project_root=None, bundled_path=None):
    """按优先级链解析当前环境下的有效模板路径；找不到返回 None（并记录尝试过的位置）。"""
    tried = []

    # 1) 用户上传的覆盖模板
    if uploaded_override and os.path.exists(uploaded_override):
        return uploaded_override
    if uploaded_override:
        tried.append(f"上传模板:{uploaded_override}")

    # 2) 随脚本绑定的模板副本（阶段2）
    if bundled_path and os.path.exists(bundled_path):
        return bundled_path
    if bundled_path:
        tried.append(f"随脚本副本:{bundled_path}")

    name, hsh, baked = extract_template_ref(script_code)
    root = str(project_root) if project_root else os.getcwd()

    # 3) 当前租户目录按文件名（+哈希校验）
    if tenant_id and name:
        tdir = os.path.join(root, "tenants", str(tenant_id))
        hit = _find_by_name(tdir, name, hsh)
        if hit:
            return hit
        tried.append(f"租户目录:{tdir}/**/{name}")

    # 4) 全局资源目录
    if name:
        gdir = os.path.join(root, "global_assets")
        hit = _find_by_name(gdir, name, hsh)
        if hit:
            return hit
        tried.append(f"全局资源:{gdir}/**/{name}")

    # 5) 烘焙的绝对路径（仅当本机存在）
    if baked and os.path.exists(baked):
        return baked
    if baked:
        tried.append(f"烘焙路径:{baked}")

    logger.warning("[模板解析] 未定位到模板 name=%s hash=%s，尝试过: %s",
                   name, hsh, " | ".join(tried) if tried else "(无线索)")
    return None
