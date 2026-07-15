"""目标模板表语义解析：把训练固化的 _COL_MAP 键映射到当月模板里实际的 sheet 名。

供两处**共用同一套逻辑**，保证「事前校验」与「运行时」判定完全一致（否则 precheck
判 OK、运行时却因写法差异跳过，又会退化成静默漏填）：
- backend/utils/compute_precheck.py —— 提交时校验目标表能否唯一解析，歧义/落空产出
  候选交前端人工选择
- backend/ai_engine/template_code_generator.py 生成的骨架 —— 运行时按解析结果把实际
  sheet 临时改名对齐 _COL_MAP 键，填完改回

匹配分层（越靠前越优先，命中即锁定）：
  1) 精确名     —— key 本身就是 wb 里的 sheet 名（训练=智算模板同名时零变化）
  2) 人工映射   —— manual_map 里用户已指定的
  3) 列签名语义 —— 顶部多行合并词表，每个列名按「整名或去 前缀- 取尾段」命中；
                   覆盖度≥阈值且唯一显著领先才自动认领
出现歧义（多候选并列、无法唯一确定）→ 计入 ambiguous，**绝不按出现位置猜**；
彻底无候选的键 → 计入 unresolved（当月缺该表的合法场景）。
"""

from typing import Any, Dict, List, Tuple


def _norm_name(s) -> str:
    """列名归一：去空白/换行/全角空格、转小写，便于跨模板比对。"""
    return "".join(str(s or "").split()).replace("　", "").lower()


def _name_variants(name: str) -> set:
    """一个列名的可匹配变体：整名 + 去掉「太保填写-」「合杰更新-」等前缀后的尾段。

    训练时列名常是「横幅-子表头」合并名（太保填写-姓名 / 合杰更新-上个月个税），
    当月模板单行表头只有子表头（姓名 / 上个月个税），故整名与尾段都要能命中。
    """
    n = _norm_name(name)
    out = {n}
    for sep in ("-", "_", "|", "：", ":"):
        if sep in name:
            out.add(_norm_name(name.split(sep)[-1]))
    return {v for v in out if v}


def _sheet_header_vocab(ws, max_rows: int = 6) -> set:
    """把 sheet 顶部若干行的所有非空单元格文本归一后汇成词表（覆盖多行表头）。

    不信任固化 header_row：当月模板可能有标题横幅致表头行偏移，合并顶部多行天然
    覆盖「横幅行 + 子表头行」。跳过公式串（以 = 开头）。
    """
    vocab = set()
    try:
        mr = min(max_rows, ws.max_row or 1)
        for r in range(1, mr + 1):
            for cell in ws[r]:
                v = cell.value
                if v is None:
                    continue
                s = str(v).strip()
                if not s or s.startswith("Unnamed") or s.startswith("="):
                    continue
                vocab.add(_norm_name(s))
    except Exception:
        pass
    return vocab


def _key_signature(col_map: Dict[str, Any], key: str) -> List[str]:
    """_COL_MAP[key] 各区域列名列表（训练时固化的目标表列签名，未归一）。"""
    sig = []
    for region in (col_map.get(key) or {}).get("regions", []):
        for col in region.get("columns", []):
            nm = str(col.get("name") or "").strip()
            if nm and not nm.startswith("Unnamed"):
                sig.append(nm)
    return sig


def resolve_target_sheets(
    wb,
    col_map: Dict[str, Any],
    source_prefix: str = "源_",
    manual_map: Dict[str, str] = None,
    cover_threshold: float = 0.6,
    lead_margin: float = 0.15,
) -> Tuple[Dict[str, str], Dict[str, list], List[str]]:
    """把 col_map 的键语义映射到 wb 里实际的 sheet 名。

    Returns:
        resolved:   {col_map_key: 实际sheet名}
        ambiguous:  {col_map_key: [{"sheet":.., "score":..}, ...]}  多候选并列、需人工确认
        unresolved: [col_map_key, ...]  无任何候选（当月缺该表，合法，仅告警）
    纯函数、不抛异常；歧义如何处理由调用方决定（运行时抛错 / 校验时产出候选）。
    """
    manual = manual_map or {}
    resolved: Dict[str, str] = {}
    claimed = set()

    def _candidates():
        out = []
        for sn in wb.sheetnames:
            if sn.startswith(source_prefix):
                continue
            try:
                if wb[sn].sheet_state == "hidden":
                    continue
            except Exception:
                pass
            out.append(sn)
        return out

    # 1) 精确名
    for key in col_map.keys():
        if key in wb.sheetnames:
            resolved[key] = key
            claimed.add(key)

    # 2) 人工映射（用户在前端已指定的，优先于语义猜测）
    for key in col_map.keys():
        if key in resolved:
            continue
        tgt = manual.get(key)
        if tgt and tgt in wb.sheetnames and tgt not in claimed:
            resolved[key] = tgt
            claimed.add(tgt)

    # 3) 列签名语义匹配（仅对还没解决的键）
    ambiguous: Dict[str, list] = {}
    unresolved: List[str] = []
    for key in [k for k in col_map.keys() if k not in resolved]:
        sig = _key_signature(col_map, key)
        if not sig:
            unresolved.append(key)
            continue
        variants = [_name_variants(n) for n in sig]
        scored = []
        for sn in _candidates():
            if sn in claimed:
                continue
            vocab = _sheet_header_vocab(wb[sn])
            if not vocab:
                continue
            hit = sum(1 for vs in variants if vs & vocab)
            cover = hit / len(sig)
            if cover > 0:
                scored.append((cover, sn))
        scored.sort(key=lambda x: (-x[0], x[1]))
        strong = [s for s in scored if s[0] >= cover_threshold]
        if len(strong) == 1:
            resolved[key] = strong[0][1]
            claimed.add(strong[0][1])
        elif len(strong) >= 2 and (strong[0][0] - strong[1][0]) >= lead_margin:
            resolved[key] = strong[0][1]
            claimed.add(strong[0][1])
        elif strong:
            ambiguous[key] = [{"sheet": sn, "score": round(cov, 3)} for cov, sn in scored[:5]]
        else:
            unresolved.append(key)

    return resolved, ambiguous, unresolved
