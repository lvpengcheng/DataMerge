"""计算输出后处理：统一文件名、模板格式兜底、纯值副本生成。

集中放置三类计算结果后处理逻辑，供 run_compute_task / compute2 等所有计算路径复用：
1. build_result_filename：固化结果文件名（脚本名_薪资年月_时间戳 / 脚本名_时间戳）。
2. restore_formats_from_template：把输出单元格格式刷回模板原格式（修复 openpyxl 写
   datetime 自动改日期格式的问题；旧脚本兜底，幂等，不改值不重算）。
3. make_values_only_copy：把带公式结果转成纯值副本，并只保留目标 sheet（去掉 源_ 源数据 sheet）。
"""

import os
import re
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

_ILLEGAL = re.compile(r'[\\/:*?"<>|]+')


def _safe_name(s, fallback="result", max_bytes=120):
    s = (str(s) if s else "").strip() or fallback
    s = _ILLEGAL.sub("_", s)
    # 按 UTF-8 字节截断脚本名部分：防止文件名过长触发文件系统(单段≤255字节)/路径长度上限，
    # 否则原版能存、更长的"_纯值"版存不出来 → 纯值版按钮缺失。中文按 3 字节计。
    b = s.encode("utf-8")
    if len(b) > max_bytes:
        s = b[:max_bytes].decode("utf-8", "ignore").rstrip("_ ") or fallback
    return s


def build_result_filename(script_name, salary_year=None, salary_month=None, ext=".xlsx") -> str:
    """固化结果文件名。

    - 填了薪资年月 → 脚本名_YYYYMM_YYYYMMDDHHMMSS
    - 没填        → 脚本名_YYYYMMDDHHMMSS
    """
    safe = _safe_name(script_name)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    if not ext.startswith("."):
        ext = "." + ext
    if salary_year and salary_month:
        try:
            ym = f"{int(salary_year):04d}{int(salary_month):02d}"
            return f"{safe}_{ym}_{ts}{ext}"
        except (ValueError, TypeError):
            pass
    return f"{safe}_{ts}{ext}"


def values_only_name(formula_filename: str, suffix="_纯值") -> str:
    """由带公式版文件名派生纯值版文件名（在扩展名前插入后缀）。"""
    root, ext = os.path.splitext(formula_filename)
    return f"{root}{suffix}{ext or '.xlsx'}"


def dual_output_enabled() -> bool:
    """读取 .env 开关：是否在带公式版之外再出一个纯值版（默认 true=出两个）。
    显式设为 false/0/no/off 可关闭，只出带公式版。"""
    val = os.getenv("COMPUTE_OUTPUT_VALUES_COPY", "true").strip().lower()
    return val not in ("false", "0", "no", "off", "")


def restore_formats_from_template(output_path, template_path) -> int:
    """把输出文件的单元格 number_format 刷回模板原格式。

    解决 openpyxl 写入 datetime 值时自动把 General 改成日期格式、导致模板常规列
    在输出里变日期的问题。仅按"同名 sheet + 同坐标"恢复格式，不改值、不重算，幂等。
    返回恢复的单元格数。

    **必须用 Aspose，不能用 openpyxl**：openpyxl `load→save` 会把整个文件重写，
    而它 load 时把"数字+日期格式"的单元格（如 源_ 里被套日期格式的金额 60.74）读成
    datetime，save 时按 1900 幽灵闰日塌缩成 59.74——即使本函数只想改模板 sheet 的格式，
    也会顺带把 源_ 的值毁掉。Aspose 只改样式、底层 DoubleValue 精确保留。
    """
    if not template_path or not os.path.exists(template_path) or not os.path.exists(output_path):
        return 0
    if not str(output_path).lower().endswith((".xlsx", ".xlsm")):
        return 0
    try:
        from Aspose.Cells import Workbook
        import aspose_init
        aspose_init.ensure_license()
    except Exception as e:
        logger.warning(f"[fmt兜底] Aspose 不可用，跳过: {e}")
        return 0
    try:
        tpl = Workbook(str(template_path))
    except Exception as e:
        logger.warning(f"[fmt兜底] 打开模板失败，跳过: {template_path} - {e}")
        return 0

    restored = 0
    out = None
    try:
        out = Workbook(str(output_path))
        # 模板 sheet 名 -> worksheet
        tpl_by_name = {str(tpl.Worksheets[ti].Name): tpl.Worksheets[ti]
                       for ti in range(tpl.Worksheets.Count)}
        touched = False
        for oi in range(out.Worksheets.Count):
            ows = out.Worksheets[oi]
            tws = tpl_by_name.get(str(ows.Name))
            if tws is None:
                continue   # 源_ 等模板没有的 sheet 不动（也不会被 save 毁值，Aspose 保底层数值）
            ocells = ows.Cells
            tcells = tws.Cells
            # 只遍历输出中"实际存在"的单元格（GetEnumerator 不含空格 → 不会给输出灌空单元格）
            it = ocells.GetEnumerator()
            while it.MoveNext():
                ocell = it.Current
                try:
                    tcell = tcells[ocell.Row, ocell.Column]   # 访问模板空格仅在模板侧实例化，不落盘
                    tstyle = tcell.GetStyle()
                    ostyle = ocell.GetStyle()
                    if ostyle.Number != tstyle.Number or (ostyle.Custom or "") != (tstyle.Custom or ""):
                        ostyle.Number = tstyle.Number
                        ostyle.Custom = tstyle.Custom
                        ocell.SetStyle(ostyle)
                        restored += 1
                        touched = True
                except Exception:
                    continue
        if touched:
            out.Save(str(output_path))
            logger.info(f"[fmt兜底] 已按模板恢复 {restored} 个单元格格式（Aspose 保值）: {output_path}")
    except Exception as e:
        logger.warning(f"[fmt兜底] 处理异常，按原文件: {output_path} - {e}")
        return 0
    finally:
        for _wb in (out, tpl):
            try:
                if _wb is not None:
                    _wb.Dispose()
            except Exception:
                pass
    return restored


def _template_default_date_format(template_path):
    """解析模板 styles.xml，取"默认(Normal)单元格样式"的 numFmt formatCode。

    仅当它是日期格式时返回该 formatCode，否则返回 None。
    取 cellStyleXfs[0]（Normal 样式基准），回退 cellXfs[0]（默认单元格格式）。
    直接解析 XML（不依赖 openpyxl 的 named_styles，后者跨版本 key 不稳定）。
    """
    import zipfile
    try:
        from openpyxl.styles.numbers import is_date_format, builtin_format_code, BUILTIN_FORMATS
    except Exception:
        return None
    try:
        with zipfile.ZipFile(str(template_path)) as z:
            s = z.read("xl/styles.xml").decode("utf-8", "ignore")
    except Exception:
        return None
    fmts = dict(re.findall(r'<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', s))

    def _first_xf_numfmtid(block_name):
        mb = re.search(r"<%s[^>]*>(.*?)</%s>" % (block_name, block_name), s, re.S)
        if not mb:
            return None
        m0 = re.search(r'<xf\b[^>]*?numFmtId="(\d+)"', mb.group(1))
        return m0.group(1) if m0 else None

    numfmtid = _first_xf_numfmtid("cellStyleXfs") or _first_xf_numfmtid("cellXfs")
    if numfmtid is None:
        return None
    code = fmts.get(numfmtid)
    if code is None:
        # 内建格式（0-163）：尝试取内建 formatCode
        try:
            code = BUILTIN_FORMATS.get(int(numfmtid))
        except Exception:
            code = None
    if not code:
        return None
    try:
        return code if is_date_format(code) else None
    except Exception:
        return None


def normalize_source_sheet_formats(output_path, template_path, source_sheet_prefix="源_") -> int:
    """修复 源_ sheet 继承模板"日期默认样式"导致数字被显示成日期的问题。

    根因：部分模板的 Normal/默认单元格样式 numFmt 是日期格式（如 `[$-409]dd/mmm/yy;@`），
    脚本用 openpyxl 追加 `源_` sheet 时，未显式设格式的单元格会继承这个默认样式，
    于是序号/工资等**数字**被显示成日期（底层值不变，仅格式错）。

    零误伤策略：仅当模板的默认(Normal)样式**本身是日期格式**时才介入，且**只把格式
    恰好等于该默认格式**的 `源_` 单元格拉回 `General`。真正的日期列（写入时显式设为
    `yyyy-mm-dd`）与其它任何显式格式都不动。幂等，不改值、不重算。返回修复的单元格数。

    该兜底覆盖所有脚本（含未更新骨架的旧脚本），无需重新训练即可生效。

    **必须用 Aspose 改格式，不能用 openpyxl**：openpyxl `load` 会把"数字+日期格式"读成
    datetime，`save` 时 `to_excel` 又按 1900 幽灵闰日塌缩，导致 60.74→59.74 值损坏。
    Aspose 只改样式、底层 `DoubleValue` 精确保留，不会动值。
    """
    if not template_path or not os.path.exists(template_path) or not os.path.exists(output_path):
        return 0
    if not str(output_path).lower().endswith((".xlsx", ".xlsm")):
        return 0

    default_fmt = _template_default_date_format(template_path)
    if not default_fmt:
        return 0  # 模板默认样式不是日期格式 → 无此问题，跳过

    # 列名关键词判定日期列：真日期列（入职时间/日期等）保留成日期，仅非日期列清成常规
    try:
        from backend.utils.source_sheet_writer import is_date_keyword_column
    except Exception:
        try:
            from .source_sheet_writer import is_date_keyword_column
        except Exception:
            def is_date_keyword_column(_n):  # 兜底：拿不到就一律按非日期处理
                return False

    try:
        from Aspose.Cells import Workbook
        import aspose_init
        aspose_init.ensure_license()
    except Exception as e:
        logger.warning(f"[源_格式兜底] Aspose 不可用，跳过: {e}")
        return 0

    def _norm_fmt(f):
        return re.sub(r"\s+", "", str(f or "")).lower()
    _target = _norm_fmt(default_fmt)

    fixed = 0
    try:
        wb = Workbook(str(output_path))
    except Exception as e:
        logger.warning(f"[源_格式兜底] 打开失败，跳过: {output_path} - {e}")
        return 0
    try:
        touched = False
        for si in range(wb.Worksheets.Count):
            ws = wb.Worksheets[si]
            if not (source_sheet_prefix and str(ws.Name).startswith(source_sheet_prefix)):
                continue
            cells = ws.Cells
            try:
                mc = cells.MaxDataColumn   # 0-based，-1 表示空
            except Exception:
                mc = -1
            if mc < 0:
                continue
            # 表头（第 0 行）→ 逐列日期列判定
            _date_col = {}
            for ci in range(0, mc + 1):
                try:
                    hv = cells[0, ci].Value
                    _date_col[ci] = bool(hv is not None and is_date_keyword_column(hv))
                except Exception:
                    _date_col[ci] = False
            it = cells.GetEnumerator()
            while it.MoveNext():
                cell = it.Current
                try:
                    style = cell.GetStyle()
                    # 只动"格式恰好等于模板默认日期格式"的单元格（继承来的伪日期）
                    if _norm_fmt(style.Custom) != _target:
                        continue
                    if cell.Row > 0 and _date_col.get(cell.Column):
                        style.Custom = "yyyy-mm-dd"   # 真日期列 → 规范日期显示
                    else:
                        style.Custom = "General"      # 表头/非日期列 → 常规
                    cell.SetStyle(style)
                    fixed += 1
                    touched = True
                except Exception:
                    continue
        if touched:
            wb.Save(str(output_path))
            logger.info(f"[源_格式兜底] 已规范 {fixed} 个继承模板日期默认格式的单元格"
                        f"（日期列→yyyy-mm-dd，其余→General；Aspose 保值）: {output_path}")
    except Exception as e:
        logger.warning(f"[源_格式兜底] 处理异常，按原文件: {output_path} - {e}")
        return 0
    finally:
        try:
            wb.Dispose()
        except Exception:
            pass
    return fixed


def make_values_only_copy(src_xlsx, dst_xlsx, source_sheet_prefix="源_",
                          keep_sheets=None, template_path=None, sheet_name_map=None):
    """生成纯值副本：新填列公式→值，模版原有公式保留。

    公式取舍（选择性拍平）：
      - template_path 给定时：先从**实际生成结果所用的模版**收集"原本就是公式的单元格坐标"
        作为保护名单，拍平时只把**不在名单内**的公式（= AI 新填列）转成值，模版自带公式保留。
      - template_path 为 None 时：退化为旧行为，全部公式拍平（向后兼容）。
      - sheet_name_map：{模版sheet名: 输出sheet名} 别名映射。当输出文件的 sheet 名与模版
        不同（上传模版 + sheet 改名场景）时，用它把保护名单的 key 同时登记到输出侧名字，
        避免因名字对不上导致模版公式被误拍平。

    sheet 取舍：只删除 源_ 前缀的源数据 sheet，**保留模版所有其他 sheet**，
      避免保留下来的模版公式因引用的辅助 sheet 被删而变成 #REF!。
      （keep_sheets 参数已废弃，保留仅为签名兼容，不再据此删 sheet。）

    返回 dst_xlsx（成功）或 None（失败/被跳过）。
    """
    try:
        from Aspose.Cells import Workbook
        import aspose_init
        aspose_init.ensure_license()
    except Exception as e:
        logger.warning(f"[纯值版] Aspose 不可用，跳过: {e}")
        return None
    try:
        wb = Workbook(str(src_xlsx))
    except Exception as e:
        logger.warning(f"[纯值版] 打开失败，跳过: {src_xlsx} - {e}")
        return None
    try:
        try:
            wb.CalculateFormula()
        except Exception as _ce:
            logger.warning(f"[纯值版] CalculateFormula 跳过: {_ce}")

        # 0) 从"实际生成结果所用的模版"收集"原有公式单元格"坐标 → 保护名单
        #    模版自带公式列保留为公式，只把新填列（模版里原本无公式）拍成值。
        #    key: sheet名(去空格)；value: set((row, col))，Aspose 0 基坐标
        protected = {}
        if template_path:
            try:
                _twb = Workbook(str(template_path))
                try:
                    for _ti in range(_twb.Worksheets.Count):
                        _tws = _twb.Worksheets[_ti]
                        _tname = str(_tws.Name).strip()
                        _pset = set()
                        _tit = _tws.Cells.GetEnumerator()
                        while _tit.MoveNext():
                            _tc = _tit.Current
                            try:
                                if _tc.IsFormula:
                                    _pset.add((_tc.Row, _tc.Column))
                            except Exception:
                                continue
                        # 登记到模版名；若有别名映射(输出侧改了名)，同名字也登记一份
                        protected[_tname] = _pset
                        if sheet_name_map:
                            _alias = sheet_name_map.get(str(_tws.Name)) or sheet_name_map.get(_tname)
                            if _alias:
                                protected[str(_alias).strip()] = _pset
                finally:
                    try:
                        _twb.Dispose()
                    except Exception:
                        pass
                logger.info("[纯值版] 模版公式保护名单: " +
                            (", ".join(f"{k}={len(v)}格" for k, v in protected.items() if v) or "无"))
            except Exception as _pe:
                logger.warning(f"[纯值版] 读取模版公式名单失败，将全部拍平: {template_path} - {_pe}")
                protected = {}

        # 1) 公式格转纯值：仅拍平"不在模版保护名单里"的公式（= 新填列）
        for i in range(wb.Worksheets.Count):
            try:
                ws = wb.Worksheets[i]
                _pset = protected.get(str(ws.Name).strip(), set())
                cells = ws.Cells
                it = cells.GetEnumerator()
                while it.MoveNext():
                    cell = it.Current
                    try:
                        if cell.IsFormula and (cell.Row, cell.Column) not in _pset:
                            cell.PutValue(cell.Value)
                    except Exception:
                        continue
            except Exception as _we:
                logger.warning(f"[纯值版] sheet 转值跳过(忽略): idx={i} - {_we}")
                continue

        # 2) 只删除 源_ 前缀的源数据 sheet，保留模版所有其他 sheet（避免模版公式断链）
        try:
            names = [str(wb.Worksheets[i].Name) for i in range(wb.Worksheets.Count)]
            # 删除 源_ 前缀的源数据 sheet（倒序删，避免索引错位）
            # 保护：至少保留一个 sheet —— 若删完会为空（极端：全是 源_ sheet），则不删，
            # 避免 Aspose Save 因"工作簿无 sheet"报错导致整个纯值版生成失败。
            src_idx = [i for i, n in enumerate(names)
                       if source_sheet_prefix and n.startswith(source_sheet_prefix)]
            non_src_count = wb.Worksheets.Count - len(src_idx)
            if non_src_count >= 1:
                for i in sorted(src_idx, reverse=True):
                    try:
                        wb.Worksheets.RemoveAt(i)
                    except Exception as _re:
                        logger.warning(f"[纯值版] 删除源sheet失败(忽略): idx={i} - {_re}")
            elif src_idx:
                logger.warning(f"[纯值版] 全部为源数据sheet，保留以避免空工作簿: {names}")
        except Exception as _se:
            logger.warning(f"[纯值版] 处理sheet时异常(忽略，继续保存): {_se}")

        wb.Save(str(dst_xlsx))
        logger.info(f"[纯值版] 已生成: {dst_xlsx}")
        return str(dst_xlsx)
    except Exception as e:
        logger.exception(f"[纯值版] 生成失败（返回 None）: {src_xlsx} -> {dst_xlsx}: {e}")
        return None
    finally:
        try:
            wb.Dispose()
        except Exception:
            pass
