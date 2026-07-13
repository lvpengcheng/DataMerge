# -*- coding: utf-8 -*-
"""模版模式 - 规则驱动的结构化行清洗（阶段0 预处理）。

在"填列"之前，按主键对模版数据区做**增/删行**，并保住模版已设公式：
  - 每行公式（如 =ROUND((P4+...)*6.78%,2)）随 CopyRow 复制到新行、相对引用自动指向本行；
  - 本表汇总公式（=SUM(P4:P303)）随插删自动伸缩；
  - **跨 sheet 引用**（如 请款!=SUM(本月工资明细!P11:P20)）也随之调整。

────────────────────────────────────────────────────────────────────
【硬约束 —— 娇韵诗真实文件实测得出，勿改】
1. 增删行必须开启 updateReference：
     Cells.DeleteRows(idx, cnt, True)                      # 布尔三参重载=含跨表引用更新
     Cells.InsertRows(idx, cnt, InsertOptions(UpdateReference=True))
   默认的两参重载 **只更新本表引用、不更新其他 sheet 指向本表的引用**，
   会导致 请款/开票 等下游汇总静默错位（数字自洽但分组错）。
2. 新增行必须插在数据区 **末行之前**（落在 SUM 范围内部），否则汇总/跨表 SUM
   不会把新行纳入（Excel 同款行为：在范围末尾之后插入不扩展范围）。
────────────────────────────────────────────────────────────────────

本模块只做"行结构"，不填业务列（填列仍由阶段1 的 fill_template 负责）。
"""

import logging

logger = logging.getLogger(__name__)

# 数据区末尾的汇总/合计行：其主键列可能有值（如"总计"落在证件号码列），
# 若只按"主键列连续非空"会把汇总行误当数据行 → 用关键字在此截断。
_SUMMARY_MARKERS = ("合计", "总计", "小计", "汇总", "共计", "累计", "total", "subtotal")


def _is_summary_key(nk):
    low = nk.lower()
    return any(m in nk or m in low for m in _SUMMARY_MARKERS)


def _norm_key(v):
    """主键归一：None→''，去首尾空白；数字统一成不带小数点的字符串以便与文本主键对齐。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def detect_data_region(cells, key_col_idx, data_start=None, max_scan=None):
    """圈定数据区 [data_start, data_end]（0-based，闭区间）。

    以**主键列连续有值**为准，即便解析器给不出 data_end 也稳。
      - data_start 未给时：从第 0 行往下找第一行主键列非空处作为起点。
      - data_end：从 data_start 起，主键列连续非空的最后一行。

    返回 (data_start, data_end)；数据区为空返回 (None, None)。
    """
    last = cells.MaxDataRow
    if last is None or last < 0:
        return (None, None)
    scan_to = last if max_scan is None else min(last, (data_start or 0) + max_scan)

    if data_start is None:
        data_start = None
        for r in range(0, last + 1):
            if _norm_key(cells[r, key_col_idx].Value):
                data_start = r
                break
        if data_start is None:
            return (None, None)

    data_end = data_start - 1
    for r in range(data_start, scan_to + 1):
        k = _norm_key(cells[r, key_col_idx].Value)
        if not k or _is_summary_key(k):   # 空行或汇总行 → 数据区到此为止
            break
        data_end = r
    if data_end < data_start:
        return (None, None)
    return (data_start, data_end)


def build_row_plan(existing_keys, add_keys=None, remove_keys=None):
    """产出确定性行计划。

    existing_keys: 现有数据区主键，按行顺序。
    add_keys:      要新增的项；每项可以是：
                     - 字符串 key（无锚点 → 追加到数据区末尾）
                     - dict {"key": 新主键, "after": 锚点主键}（插在锚点行之后、加入其分组）
    remove_keys:   要删除的主键集合。

    返回 dict：
      remove: set(实际删的主键)
      add:    list[{"key":.., "after":锚点或None}]（保序去重、剔除已存在的）
      keep:   list(保留主键)
    """
    existing = [_norm_key(k) for k in existing_keys]
    existing_set = set(existing)
    remove_set = {_norm_key(k) for k in (remove_keys or [])} & existing_set

    seen = set()
    add_list = []
    for item in (add_keys or []):
        if isinstance(item, dict):
            nk = _norm_key(item.get("key"))
            after = _norm_key(item.get("after")) or None
            before = _norm_key(item.get("before")) or None
        else:
            nk, after, before = _norm_key(item), None, None
        if nk and nk not in existing_set and nk not in seen:
            add_list.append({"key": nk, "after": after, "before": before})
            seen.add(nk)

    keep = [k for k in existing if k not in remove_set]
    return {"remove": remove_set, "add": add_list, "keep": keep}


def _merge_desc_runs(rows):
    """把行号列表合并成连续区间，按起点降序返回 [(start, count), ...]，供倒序删除。"""
    if not rows:
        return []
    rows = sorted(set(rows))
    runs = []
    s = p = rows[0]
    for r in rows[1:]:
        if r == p + 1:
            p = r
        else:
            runs.append((s, p - s + 1))
            s = p = r
    runs.append((s, p - s + 1))
    runs.sort(key=lambda x: x[0], reverse=True)
    return runs


def _mutate(cells, key_col_idx, ds, existing_keys, plan):
    """在已打开的 cells 上执行删/增（含跨表引用更新），返回 (removed, added)。

    ds: 数据区起始行(0based)。existing_keys: ds 起逐行主键。plan: build_row_plan 结果。
    不做 CalculateFormula/Save（由调用方负责）。
    """
    from Aspose.Cells import InsertOptions

    # 1) 删行（倒序、连续合并、开启跨表更新）
    removed = 0
    if plan["remove"]:
        del_rows = [ds + i for i, k in enumerate(existing_keys) if k in plan["remove"]]
        for start, cnt in _merge_desc_runs(del_rows):
            cells.DeleteRows(start, cnt, True)   # True=更新引用(含跨表)
            removed += cnt

    # 2) 增行：解析每项的插入位置 P（0based，在 P 处插入=插在原 P 行之前）
    #    - after=锚点  → P = 锚点行 + 1（插在锚点之后）
    #    - before=锚点 → P = 锚点行（插在锚点之前）；用于"入组"：锚点取该组最后一行，
    #      插在其之前 → 落在该组 SUM 范围**内部**，范围随之扩展（若插在组末行之后则不扩展）。
    #    - 无锚点      → 追加到数据区末行之前。
    #    按 P 降序处理，避免下方插入移动上方待插位置。CopyRow 样板取 P-1(同区相邻行)。
    added = 0
    if plan["add"]:
        io = InsertOptions()
        io.UpdateReference = True

        def _kmap():
            ds3, de3 = detect_data_region(cells, key_col_idx, data_start=ds)
            m = {}
            if ds3 is not None:
                for r in range(ds3, de3 + 1):
                    m[_norm_key(cells[r, key_col_idx].Value)] = r
            return m, ds3, de3

        kmap, _, de_now = _kmap()
        by_pos = {}
        floating = []
        for a in plan["add"]:
            pos = None
            if a.get("before"):
                row = kmap.get(a["before"])
                pos = row if row is not None else None
            elif a.get("after"):
                row = kmap.get(a["after"])
                pos = (row + 1) if row is not None else None
            if pos is None:
                floating.append(a["key"])
            else:
                by_pos.setdefault(pos, []).append(a["key"])

        # 无锚点：统一追加到当前数据区末行之前
        if floating:
            by_pos.setdefault(de_now, []).extend(floating)

        for pos in sorted(by_pos, reverse=True):
            keys = by_pos[pos]
            n = len(keys)
            cells.InsertRows(pos, n, io)
            sample = pos - 1 if pos - 1 >= ds else pos + n   # 相邻同区行作样板
            for i, k in enumerate(keys):
                dest = pos + i
                cells.CopyRow(cells, sample, dest)           # 复制格式+每行公式(相对引用自动校正)
                cells[dest, key_col_idx].PutValue(k)
            added += n

    return removed, added


def apply_row_plan_aspose(xlsx_path, sheet_name, key_col_idx,
                          add_keys=None, remove_keys=None,
                          data_start=None, dispose=True):
    """在 xlsx 上对指定 sheet 按主键增删行（Aspose，含跨表引用更新）。

    流程：打开 → 圈数据区 → 删(倒序、含跨表更新) → 重新圈区 → 在末行前插入并 CopyRow
          复制样板行(格式+每行公式) → 给新行写主键 → Save+Dispose。

    返回布局 dict：
      {"key_to_row": {主键: 0based行}, "new_rows": [0based...],
       "data_start": int, "data_end": int, "removed": int, "added": int}
    失败抛异常（调用方决定是否阻断）。
    """
    import aspose_init
    aspose_init.ensure_license()
    from Aspose.Cells import Workbook, InsertOptions

    wb = Workbook(str(xlsx_path))
    try:
        ws = None
        for i in range(wb.Worksheets.Count):
            if str(wb.Worksheets[i].Name).strip() == str(sheet_name).strip():
                ws = wb.Worksheets[i]
                break
        if ws is None:
            raise ValueError(f"[行清洗] 模版缺少目标 sheet: {sheet_name!r}")
        cells = ws.Cells

        ds, de = detect_data_region(cells, key_col_idx, data_start=data_start)
        if ds is None:
            raise ValueError(f"[行清洗] 无法圈定数据区（主键列={key_col_idx}）: {sheet_name!r}")

        existing_keys = [_norm_key(cells[r, key_col_idx].Value) for r in range(ds, de + 1)]
        plan = build_row_plan(existing_keys, add_keys=add_keys, remove_keys=remove_keys)
        removed, added = _mutate(cells, key_col_idx, ds, existing_keys, plan)
        if removed or added:
            logger.info(f"[行清洗] {sheet_name}: 删除 {removed} / 新增 {added}")

        wb.CalculateFormula()
        wb.Save(str(xlsx_path))

        # 3) 回读最终布局
        fds, fde = detect_data_region(cells, key_col_idx, data_start=ds)
        key_to_row, new_rows = {}, []
        add_set = {a["key"] for a in plan["add"]}
        if fds is not None:
            for r in range(fds, fde + 1):
                k = _norm_key(cells[r, key_col_idx].Value)
                if k:
                    key_to_row[k] = r
                    if k in add_set:
                        new_rows.append(r)
        return {"key_to_row": key_to_row, "new_rows": new_rows,
                "data_start": fds, "data_end": fde,
                "removed": removed, "added": added}
    finally:
        if dispose:
            try:
                wb.Dispose()
            except Exception:
                pass


def clean_template_rows(xlsx_path, sheet_name, key_col_idx, *,
                        group_col_idx=None, add_rows=None, remove_keys=None,
                        where=None, keep_only_keys=None, data_start=None, dispose=True):
    """高层入口：一次打开，按"分组列自动入组 + 三种删除 + 可选以源表为准对齐"解析并执行行清洗。

    参数（列均为 0-based 列号）：
      key_col_idx:   主键列。
      group_col_idx: 分组列（如成本中心/公司）。给定时，新增行按其分组值插到
                     "同组现有行的末尾"（该组的跨表 SUM 块内）；为空则追加到数据区末尾。
      add_rows:      list[{"key": 新主键, "group": 分组值(可选)}] —— 来自源表(入职名单/本月名单)。
      remove_keys:   显式/源表命中的待删主键集合。
      where:         (col_idx, equals_value) 按模版某列值删行（如 状态列==离职）。
      keep_only_keys: 若给定（以源表为准对齐/sync_to_source）：模版中**不在该集合**的主键
                     一律删除（离职/差异行）。与 add_rows 配合即"删差集+增差集+留交集"。
      data_start:    数据区起始行(0based) 提示；不传则自动圈定(汇总关键字截断)。

    返回同 apply_row_plan_aspose 的布局 dict。
    """
    import aspose_init
    aspose_init.ensure_license()
    from Aspose.Cells import Workbook

    wb = Workbook(str(xlsx_path))
    try:
        ws = None
        for i in range(wb.Worksheets.Count):
            if str(wb.Worksheets[i].Name).strip() == str(sheet_name).strip():
                ws = wb.Worksheets[i]
                break
        if ws is None:
            raise ValueError(f"[行清洗] 模版缺少目标 sheet: {sheet_name!r}")
        cells = ws.Cells

        ds, de = detect_data_region(cells, key_col_idx, data_start=data_start)
        if ds is None:
            raise ValueError(f"[行清洗] 无法圈定数据区（主键列={key_col_idx}）: {sheet_name!r}")

        # 读现有行：主键 + 分组值 + where 列值
        existing_keys = []
        group_of = {}          # 分组值 -> 该组"最后一个现有主键"(用作锚点)
        remove_set = {_norm_key(k) for k in (remove_keys or [])}
        where_col, where_val = (where or (None, None))
        for r in range(ds, de + 1):
            k = _norm_key(cells[r, key_col_idx].Value)
            existing_keys.append(k)
            if group_col_idx is not None:
                gv = _norm_key(cells[r, group_col_idx].Value)
                group_of[gv] = k          # 顺序覆盖 → 最终为该组最后一行的主键
            if where_col is not None and _norm_key(cells[r, where_col].Value) == _norm_key(where_val):
                remove_set.add(k)

        # 以源表为准对齐：模版里不在 keep_only_keys 的主键全删（离职/差异行）
        if keep_only_keys is not None:
            _keep_norm = {_norm_key(k) for k in keep_only_keys}
            for k in existing_keys:
                if k and k not in _keep_norm:
                    remove_set.add(k)

        remove_set &= set(existing_keys)   # 只删确实存在的

        # 解析新增：按分组值找锚点（同组最后一行，且未被删）；找不到组 → 追加末尾
        kept = [k for k in existing_keys if k not in remove_set]
        kept_set = set(kept)
        add_items = []
        new_key_group = {}     # 新主键 -> 分组值（用于显式回填分组列，见下）
        for row in (add_rows or []):
            nk = _norm_key(row.get("key"))
            if not nk:
                continue
            gval = row.get("group")
            before = None
            if group_col_idx is not None and gval is not None:
                cand = group_of.get(_norm_key(gval))   # 该组最后一行主键
                if cand in kept_set:
                    before = cand    # 插在该组末行之前 → 落在该组 SUM 范围内
            add_items.append({"key": nk, "before": before})
            if group_col_idx is not None and gval is not None:
                new_key_group[nk] = _norm_key(gval)

        plan = build_row_plan(existing_keys, add_keys=add_items, remove_keys=remove_set)
        removed, added = _mutate(cells, key_col_idx, ds, existing_keys, plan)
        logger.info(f"[行清洗] {sheet_name}: 删除 {removed} / 新增 {added}（分组入组）")

        # 显式回填新增行的分组列值：CopyRow 复制的是物理邻行的分组值，
        # 当模版分组列非连续/交错排列时，邻行可能属于别的组 → 分组值会错。
        # 故按 add_rows 指定的分组值精确回填（业务列仍由 fill_template 按主键填）。
        if new_key_group:
            gds, gde = detect_data_region(cells, key_col_idx, data_start=ds)
            if gds is not None:
                pos = {}
                for r in range(gds, gde + 1):
                    pos[_norm_key(cells[r, key_col_idx].Value)] = r
                for nk, gval in new_key_group.items():
                    r = pos.get(nk)
                    if r is not None:
                        cells[r, group_col_idx].PutValue(gval)

        wb.CalculateFormula()
        wb.Save(str(xlsx_path))

        fds, fde = detect_data_region(cells, key_col_idx, data_start=ds)
        key_to_row, new_rows = {}, []
        add_set = {a["key"] for a in plan["add"]}
        if fds is not None:
            for r in range(fds, fde + 1):
                k = _norm_key(cells[r, key_col_idx].Value)
                if k:
                    key_to_row[k] = r
                    if k in add_set:
                        new_rows.append(r)
        return {"key_to_row": key_to_row, "new_rows": new_rows,
                "data_start": fds, "data_end": fde,
                "removed": removed, "added": added}
    finally:
        if dispose:
            try:
                wb.Dispose()
            except Exception:
                pass
