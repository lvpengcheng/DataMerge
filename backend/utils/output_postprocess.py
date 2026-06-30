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
    """
    try:
        import openpyxl
    except Exception as e:
        logger.warning(f"[fmt兜底] openpyxl 不可用，跳过: {e}")
        return 0
    if not template_path or not os.path.exists(template_path) or not os.path.exists(output_path):
        return 0
    if not str(output_path).lower().endswith((".xlsx", ".xlsm")):
        return 0  # .xls 老格式 openpyxl 不支持
    try:
        tpl = openpyxl.load_workbook(template_path, data_only=False)
    except Exception as e:
        logger.warning(f"[fmt兜底] 打开模板失败，跳过: {template_path} - {e}")
        return 0

    restored = 0
    try:
        out = openpyxl.load_workbook(output_path, data_only=False,
                                     keep_vba=str(output_path).lower().endswith(".xlsm"))
        # 遍历"输出"已用范围（脚本填过值的单元格在这里），按坐标取模板单元格的格式刷回。
        # 用输出范围而非模板范围：模板数据行常为空(max_row 偏小)，但 openpyxl 访问空单元格
        # 仍返回其 number_format（默认 General），正好是我们要恢复的目标。
        for ws_name in out.sheetnames:
            if ws_name not in tpl.sheetnames:
                continue
            ows = out[ws_name]
            tws = tpl[ws_name]
            try:
                max_row = ows.max_row or 0
                max_col = ows.max_column or 0
            except Exception:
                continue
            if max_row <= 0 or max_col <= 0:
                continue
            for row in ows.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                for ocell in row:
                    try:
                        tfmt = tws.cell(row=ocell.row, column=ocell.column).number_format
                        if ocell.number_format != tfmt:
                            ocell.number_format = tfmt
                            restored += 1
                    except Exception:
                        continue
        if restored > 0:
            out.save(output_path)
            logger.info(f"[fmt兜底] 已按模板恢复 {restored} 个单元格格式: {output_path}")
    except Exception as e:
        logger.warning(f"[fmt兜底] 处理异常，按原文件: {output_path} - {e}")
        return 0
    finally:
        try:
            tpl.close()
        except Exception:
            pass
    return restored


def make_values_only_copy(src_xlsx, dst_xlsx, source_sheet_prefix="源_"):
    """生成纯值副本：公式→值，并删除 源_ 前缀的源数据 sheet（只保留目标 sheet）。

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

        # 1) 公式格转纯值
        for i in range(wb.Worksheets.Count):
            try:
                ws = wb.Worksheets[i]
                cells = ws.Cells
                it = cells.GetEnumerator()
                while it.MoveNext():
                    cell = it.Current
                    try:
                        if cell.IsFormula:
                            cell.PutValue(cell.Value)
                    except Exception:
                        continue
            except Exception as _we:
                logger.warning(f"[纯值版] sheet 转值跳过(忽略): idx={i} - {_we}")
                continue

        # 2) 删除 源_ 前缀的源数据 sheet（倒序删，避免索引错位）
        #    保护：至少保留一个 sheet —— 若删完会为空（极端：全是 源_ sheet），则不删，
        #    避免 Aspose Save 因"工作簿无 sheet"报错导致整个纯值版生成失败。
        try:
            names = [str(wb.Worksheets[i].Name) for i in range(wb.Worksheets.Count)]
            src_idx = [i for i, n in enumerate(names)
                       if source_sheet_prefix and n.startswith(source_sheet_prefix)]
            non_src_count = wb.Worksheets.Count - len(src_idx)
            if non_src_count >= 1:
                # 有目标 sheet：安全删除全部 源_ sheet（倒序）
                for i in sorted(src_idx, reverse=True):
                    try:
                        wb.Worksheets.RemoveAt(i)
                    except Exception as _re:
                        logger.warning(f"[纯值版] 删除源sheet失败(忽略): idx={i} - {_re}")
            elif src_idx:
                logger.warning(f"[纯值版] 全部为源数据sheet，保留以避免空工作簿: {names}")
        except Exception as _se:
            logger.warning(f"[纯值版] 处理源sheet时异常(忽略，继续保存): {_se}")

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
