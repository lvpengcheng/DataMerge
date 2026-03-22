"""
Excel差异对比组件 - 用于对比生成结果和预期结果
"""

import logging
import os
import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional, List
import platform
import shutil

logger = logging.getLogger(__name__)


def normalize_emp_code(emp_code) -> str:
    """标准化工号: 转换为8位字符串, 不足前面补0

    示例:
    - "123" -> "00000123"
    - "12345678" -> "12345678"
    - 12345 -> "00012345"
    """
    if pd.isna(emp_code) or emp_code == "":
        return ""
    code_str = str(emp_code).strip()
    if code_str.isdigit():
        return code_str.zfill(8)
    return code_str


def calculate_excel_formulas(file_path: str) -> str:
    """计算Excel文件中的所有公式并保存

    优先使用 Aspose.Cells 内存计算（无需 Excel 软件），失败则回退 win32com。

    Args:
        file_path: Excel文件路径

    Returns:
        计算后的文件路径（原文件会被覆盖）
    """
    # ---- 方案1: Aspose.Cells（推荐，无进程开销） ----
    try:
        import aspose_init  # noqa: F401
        from Aspose.Cells import Workbook as AsposeWorkbook, LoadOptions as AsposeLoadOptions

        logger.info(f"[Aspose] 开始计算公式: {file_path}")
        wb = AsposeWorkbook(str(file_path))
        wb.CalculateFormula()
        wb.Save(str(file_path))
        logger.info(f"[Aspose] 公式计算完成: {file_path}")
        return file_path
    except ImportError:
        logger.info("Aspose.Cells 不可用，尝试 win32com")
    except Exception as e:
        logger.warning(f"[Aspose] 公式计算失败: {e}，尝试 win32com 回退")

    # ---- 方案2: win32com 回退 ----
    if platform.system() != 'Windows':
        logger.warning("非Windows系统且Aspose不可用，跳过公式计算")
        return file_path

    try:
        import pythoncom
        import win32com.client as win32
        import tempfile

        logger.info(f"[win32com] 开始计算公式: {file_path}")

        pythoncom.CoInitialize()

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            abs_path = str(Path(file_path).resolve())

            # 非ASCII路径处理
            use_temp = False
            temp_path = None
            try:
                abs_path.encode('ascii')
            except UnicodeEncodeError:
                use_temp = True
                temp_dir = tempfile.mkdtemp()
                temp_path = os.path.join(temp_dir, "temp_excel.xlsx")
                shutil.copy(abs_path, temp_path)
                work_path = temp_path
            else:
                work_path = abs_path

            wb = excel.Workbooks.Open(work_path, UpdateLinks=0)
            for sheet in wb.Sheets:
                sheet.Calculate()
            wb.Application.Calculate()
            wb.Application.CalculateFull()
            wb.SaveAs(work_path, FileFormat=51)
            wb.Close(SaveChanges=False)

            if use_temp and temp_path:
                shutil.copy(temp_path, abs_path)
                try:
                    os.remove(temp_path)
                    os.rmdir(os.path.dirname(temp_path))
                except Exception:
                    pass

            logger.info(f"[win32com] 公式计算完成: {file_path}")
            return file_path

        finally:
            try:
                excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    except ImportError:
        logger.warning("未安装pywin32，跳过Excel公式计算")
        return file_path
    except Exception as e:
        logger.warning(f"[win32com] 公式计算失败: {e}")
        return file_path


def read_excel_with_formulas_calculated(file_path: str) -> pd.DataFrame:
    """读取Excel文件，获取公式计算后的值

    对于包含公式的Excel文件，尝试获取公式计算后的值：
    1. 首先用data_only=True读取（获取Excel已缓存的计算值）
    2. 如果某些单元格值为None（公式未被计算），则尝试自己计算公式
    3. 使用多遍迭代计算，处理公式间的依赖关系（如实发工资依赖应发工资）

    Args:
        file_path: Excel文件路径

    Returns:
        包含计算值的DataFrame
    """
    import openpyxl

    try:
        # 同时打开data_only和公式版本
        wb_data = openpyxl.load_workbook(file_path, data_only=True)
        ws_data = wb_data.active

        wb_formula = openpyxl.load_workbook(file_path, data_only=False)
        ws_formula = wb_formula.active

        # 获取表头（第一行）
        headers = []
        for cell in ws_data[1]:
            headers.append(cell.value)

        # 过滤掉None的表头
        valid_headers = [h for h in headers if h is not None]
        logger.debug(f"读取到表头: {valid_headers}")

        max_row = ws_data.max_row
        max_col = len(headers)

        # 第一遍：收集所有值和公式信息
        cell_values = {}  # (row, col) -> value
        formulas = {}     # (row, col) -> formula_string
        formula_count = 0

        for row_idx in range(2, max_row + 1):
            for col_idx in range(1, max_col + 1):
                if col_idx > len(headers) or headers[col_idx - 1] is None:
                    continue

                data_cell = ws_data.cell(row=row_idx, column=col_idx)
                formula_cell = ws_formula.cell(row=row_idx, column=col_idx)

                cell_formula = formula_cell.value
                is_formula = cell_formula is not None and str(cell_formula).startswith('=')

                if is_formula:
                    formula_count += 1
                    # 先尝试用data_only的缓存值
                    if data_cell.value is not None:
                        cell_values[(row_idx, col_idx)] = data_cell.value
                    else:
                        # 标记需要计算的公式
                        formulas[(row_idx, col_idx)] = str(cell_formula)
                        cell_values[(row_idx, col_idx)] = None
                else:
                    cell_values[(row_idx, col_idx)] = data_cell.value

        # 第二遍：多次迭代计算公式（处理公式间依赖）
        max_iterations = 10
        calculated_count = 0

        for iteration in range(max_iterations):
            # 统计还有多少公式未计算
            remaining = sum(1 for k, v in cell_values.items() if k in formulas and v is None)
            if remaining == 0:
                break

            calculated_this_round = 0
            for (row_idx, col_idx), formula in formulas.items():
                if cell_values[(row_idx, col_idx)] is not None:
                    continue

                # 尝试计算公式，使用缓存的值
                calculated_value = _try_calculate_formula_with_cache(
                    formula, cell_values, row_idx, max_col
                )
                if calculated_value is not None:
                    cell_values[(row_idx, col_idx)] = calculated_value
                    calculated_count += 1
                    calculated_this_round += 1
                    logger.debug(f"迭代{iteration+1} 计算公式 ({row_idx},{col_idx}): {formula} = {calculated_value}")

            if calculated_this_round == 0:
                # 本轮没有计算任何公式，说明无法继续计算了
                break

        wb_data.close()
        wb_formula.close()

        if formula_count > 0:
            logger.info(f"文件包含 {formula_count} 个公式单元格，成功计算 {calculated_count} 个")

        # 构建DataFrame
        data = []
        for row_idx in range(2, max_row + 1):
            row_data = {}
            for col_idx, col_name in enumerate(headers, start=1):
                if col_name is None:
                    continue
                row_data[col_name] = cell_values.get((row_idx, col_idx))
            data.append(row_data)

        return pd.DataFrame(data, columns=valid_headers)

    except Exception as e:
        logger.warning(f"使用openpyxl读取失败: {e}，回退到pandas读取")
        import traceback
        traceback.print_exc()
        # 回退到普通pandas读取
        return pd.read_excel(file_path)


def _try_calculate_formula_with_cache(formula: str, cell_values: dict, row_idx: int, max_col: int) -> Any:
    """使用缓存的单元格值计算公式（支持公式间依赖）

    Args:
        formula: Excel公式字符串（如 =A2+B2）
        cell_values: 已计算的单元格值缓存 {(row, col): value}
        row_idx: 当前行号
        max_col: 最大列数

    Returns:
        计算结果或None（如果无法计算，可能是依赖的值还未计算）
    """
    import re
    import openpyxl.utils

    if not formula or not formula.startswith('='):
        return None

    try:
        expr = formula[1:]

        def get_cell_value(col_letter: str, ref_row: int) -> float:
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            cached_value = cell_values.get((ref_row, col_idx))
            if cached_value is None:
                # 依赖的值还未计算，抛出异常中断本次计算
                raise ValueError(f"依赖的单元格 {col_letter}{ref_row} 尚未计算")
            try:
                return float(cached_value)
            except (ValueError, TypeError):
                return 0.0

        def replace_cell_ref(match):
            col_letter = match.group(1)
            ref_row = int(match.group(2))
            return str(get_cell_value(col_letter, ref_row))

        # 处理SUM范围
        def replace_sum_range(match):
            start_col = match.group(1)
            start_row = int(match.group(2))
            end_col = match.group(3)
            end_row = int(match.group(4))

            total = 0.0
            start_col_idx = openpyxl.utils.column_index_from_string(start_col)
            end_col_idx = openpyxl.utils.column_index_from_string(end_col)

            for r in range(start_row, end_row + 1):
                for c in range(start_col_idx, end_col_idx + 1):
                    cached_value = cell_values.get((r, c))
                    if cached_value is None:
                        raise ValueError(f"依赖的单元格 ({r},{c}) 尚未计算")
                    if cached_value is not None:
                        try:
                            total += float(cached_value)
                        except (ValueError, TypeError):
                            pass
            return str(total)

        # 先处理SUM范围
        expr = re.sub(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', replace_sum_range, expr, flags=re.IGNORECASE)

        # 匹配单元格引用
        expr = re.sub(r'([A-Z]+)(\d+)', replace_cell_ref, expr)

        # 替换Excel函数为Python函数
        expr = re.sub(r'ROUND\s*\(', 'round(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'MAX\s*\(', 'max(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'MIN\s*\(', 'min(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'ABS\s*\(', 'abs(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'SUM\s*\(', 'sum([', expr, flags=re.IGNORECASE)
        if 'sum([' in expr:
            expr = re.sub(r'sum\(\[([^\)]+)\)', r'sum([\1])', expr)

        # 处理IF函数
        if_match = re.search(r'IF\s*\(([^,]+),([^,]+),([^\)]+)\)', expr, flags=re.IGNORECASE)
        if if_match:
            condition = if_match.group(1).strip()
            true_val = if_match.group(2).strip()
            false_val = if_match.group(3).strip()
            condition = condition.replace('<>', '!=')
            expr = re.sub(r'IF\s*\([^\)]+\)', f'({true_val} if {condition} else {false_val})', expr, flags=re.IGNORECASE)

        # 安全计算
        safe_pattern = r'^[\d\.\+\-\*\/\(\)\s\,\<\>\=\!\[\]a-z]+$'
        if re.match(safe_pattern, expr, re.IGNORECASE):
            allowed_names = {
                'round': round, 'max': max, 'min': min, 'abs': abs, 'sum': sum,
                'True': True, 'False': False
            }
            result = eval(expr, {"__builtins__": {}}, allowed_names)
            return result
        else:
            logger.debug(f"公式表达式不安全，跳过计算: {expr}")
            return None

    except ValueError as e:
        # 依赖的值还未计算，等待下一轮迭代
        logger.debug(f"公式暂无法计算 {formula}: {e}")
        return None
    except Exception as e:
        logger.debug(f"公式计算失败 {formula}: {e}")
        return None


def _try_calculate_simple_formula(formula: str, ws, row_idx: int) -> Any:
    """尝试计算Excel公式

    支持的公式格式：
    - =A2+B2+C2 (简单加法)
    - =A2-B2 (简单减法)
    - =A2*B2 (简单乘法)
    - =ROUND(A2+B2, 2) (ROUND函数)
    - =MAX(A2, 0) (MAX函数)
    - =MIN(A2, B2) (MIN函数)
    - =SUM(A2:C2) (SUM函数)
    - =IF(A2>0, A2, 0) (IF函数)

    Args:
        formula: Excel公式字符串（如 =A2+B2）
        ws: openpyxl worksheet（data_only模式）
        row_idx: 当前行号

    Returns:
        计算结果或None（如果无法计算）
    """
    import re
    import openpyxl.utils

    if not formula or not formula.startswith('='):
        return None

    try:
        # 移除等号
        expr = formula[1:]

        # 替换单元格引用为实际值
        def get_cell_value(col_letter: str, ref_row: int) -> float:
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            cell_value = ws.cell(row=ref_row, column=col_idx).value
            if cell_value is None:
                return 0.0
            try:
                return float(cell_value)
            except (ValueError, TypeError):
                return 0.0

        def replace_cell_ref(match):
            col_letter = match.group(1)
            ref_row = int(match.group(2))
            return str(get_cell_value(col_letter, ref_row))

        # 处理SUM范围（如SUM(A2:C2)）
        def replace_sum_range(match):
            start_col = match.group(1)
            start_row = int(match.group(2))
            end_col = match.group(3)
            end_row = int(match.group(4))

            total = 0.0
            start_col_idx = openpyxl.utils.column_index_from_string(start_col)
            end_col_idx = openpyxl.utils.column_index_from_string(end_col)

            for r in range(start_row, end_row + 1):
                for c in range(start_col_idx, end_col_idx + 1):
                    cell_value = ws.cell(row=r, column=c).value
                    if cell_value is not None:
                        try:
                            total += float(cell_value)
                        except (ValueError, TypeError):
                            pass
            return str(total)

        # 先处理SUM范围
        expr = re.sub(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', replace_sum_range, expr, flags=re.IGNORECASE)

        # 匹配单元格引用（如A2, BC15）
        expr = re.sub(r'([A-Z]+)(\d+)', replace_cell_ref, expr)

        # 替换Excel函数为Python函数
        expr = re.sub(r'ROUND\s*\(', 'round(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'MAX\s*\(', 'max(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'MIN\s*\(', 'min(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'ABS\s*\(', 'abs(', expr, flags=re.IGNORECASE)
        expr = re.sub(r'SUM\s*\(', 'sum([', expr, flags=re.IGNORECASE)
        # 处理SUM的结尾括号（简单情况）
        if 'sum([' in expr:
            expr = re.sub(r'sum\(\[([^\)]+)\)', r'sum([\1])', expr)

        # 处理IF函数: IF(condition, true_val, false_val) -> (true_val if condition else false_val)
        if_match = re.search(r'IF\s*\(([^,]+),([^,]+),([^\)]+)\)', expr, flags=re.IGNORECASE)
        if if_match:
            condition = if_match.group(1).strip()
            true_val = if_match.group(2).strip()
            false_val = if_match.group(3).strip()
            # 替换Excel比较运算符
            condition = condition.replace('<>', '!=')
            expr = re.sub(r'IF\s*\([^\)]+\)', f'({true_val} if {condition} else {false_val})', expr, flags=re.IGNORECASE)

        # 安全计算表达式（允许数字、运算符、函数）
        safe_pattern = r'^[\d\.\+\-\*\/\(\)\s\,\<\>\=\!\[\]a-z]+$'
        if re.match(safe_pattern, expr, re.IGNORECASE):
            # 使用安全的eval
            allowed_names = {
                'round': round, 'max': max, 'min': min, 'abs': abs, 'sum': sum,
                'True': True, 'False': False
            }
            result = eval(expr, {"__builtins__": {}}, allowed_names)
            return result
        else:
            logger.debug(f"公式表达式不安全，跳过计算: {expr}")
            return None

    except Exception as e:
        logger.debug(f"公式计算失败 {formula}: {e}")
        return None


def _standardize_column_name(col_name) -> str:
    """标准化列名：去除换行符、统一空格、去除首尾空格、统一全角半角

    Args:
        col_name: 原始列名

    Returns:
        标准化后的列名
    """
    if pd.isna(col_name):
        return ""

    col_str = str(col_name)
    # 替换换行符为空格
    col_str = col_str.replace('\n', ' ').replace('\r', ' ')
    # 统一全角括号为半角
    col_str = col_str.replace('（', '(').replace('）', ')')
    # 合并多个连续空格为一个
    import re
    col_str = re.sub(r'\s+', ' ', col_str)
    # 去除首尾空格
    col_str = col_str.strip()

    return col_str


def _standardize_key_value(value) -> str:
    """标准化主键值（去空格、统一格式、补零等）

    Args:
        value: 原始主键值

    Returns:
        标准化后的字符串
    """
    if pd.isna(value):
        return ""

    value_str = str(value).strip()

    # 如果是纯数字字符串，补零到 8 位（兼容工号场景）
    if value_str.isdigit():
        return value_str.zfill(8)

    # 其他情况：去空格、转小写（统一格式）
    return value_str.lower().replace(" ", "")


def compare_excel_files(
    result_file: str,
    expected_file: str,
    output_file: Optional[str] = None,
    primary_keys: Optional[List[str]] = None,
    auto_detect_keys: bool = False
) -> Dict[str, Any]:
    """对比两个Excel文件的差异

    优化说明:
    1. 主键灵活配置: 支持手动指定或自动检测主键列
    2. 主键标准化: 统一转换主键值后再进行匹配
    3. 逐行数据对比: 按主键逐行对比, 而不是按行号直接对比

    Args:
        result_file: 生成的结果文件路径
        expected_file: 预期的结果文件路径
        output_file: 差异报告输出文件路径(可选)
        primary_keys: 主键列名列表，如 ["工号", "姓名"] 或 ["订单号"]。默认 ["工号", "中文姓名"]
        auto_detect_keys: 如果 primary_keys 为 None，是否自动检测主键（暂未实现，预留）

    Returns:
        对比统计结果字典
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    logger.info("开始差异对比...")

    # 用Excel计算公式（两个文件都可能包含公式）
    # 生成的结果文件包含AI生成的公式
    # 预期文件也可能包含VLOOKUP等公式（不一定是纯值）
    calculate_excel_formulas(result_file)
    calculate_excel_formulas(expected_file)

    # 读取结果文件的公式（用于差异分析时展示）
    result_formulas = {}  # {列名: 公式字符串}
    try:
        import openpyxl
        wb_formulas = openpyxl.load_workbook(result_file, data_only=False)
        ws_formulas = wb_formulas.active
        # 获取表头映射
        header_row = {ws_formulas.cell(row=1, column=col).value: col for col in range(1, ws_formulas.max_column + 1)}
        # 获取第2行的公式（假设所有数据行使用相同公式模式）
        if ws_formulas.max_row >= 2:
            for col_name, col_idx in header_row.items():
                if col_name:
                    cell = ws_formulas.cell(row=2, column=col_idx)
                    cell_value = cell.value
                    if cell_value and isinstance(cell_value, str) and cell_value.startswith('='):
                        result_formulas[col_name] = cell_value
        wb_formulas.close()
    except Exception as e:
        logger.warning(f"读取公式失败: {e}")

    # 读取两个Excel文件（支持公式计算值）
    result_df = read_excel_with_formulas_calculated(result_file)
    expected_df = read_excel_with_formulas_calculated(expected_file)

    # 标准化列名（去除换行符、统一空格）
    result_df.columns = [_standardize_column_name(col) for col in result_df.columns]
    expected_df.columns = [_standardize_column_name(col) for col in expected_df.columns]

    logger.info(f"生成结果: {len(result_df)} 行")
    logger.info(f"预期结果: {len(expected_df)} 行")

    # 确定主键列
    if primary_keys is None:
        # 默认使用工号+姓名（向后兼容）
        primary_keys = ["工号", "中文姓名"]
        logger.info(f"使用默认主键: {primary_keys}")
    else:
        logger.info(f"使用指定主键: {primary_keys}")

    # 检查主键列是否存在，如果不存在则回退
    available_keys = []
    for key in primary_keys:
        if key in expected_df.columns and key in result_df.columns:
            available_keys.append(key)
        else:
            logger.warning(f"主键列 '{key}' 在某个文件中不存在，跳过")

    if not available_keys:
        # 兜底：使用第一列作为主键
        first_col = expected_df.columns[0] if len(expected_df.columns) > 0 else None
        if first_col and first_col in result_df.columns:
            available_keys = [first_col]
            logger.warning(f"所有指定主键都不存在，回退到第一列: {first_col}")
        else:
            logger.error("无法确定有效的主键列，对比可能不准确")
            available_keys = []

    primary_keys = available_keys

    # 标准化主键列
    for key_col in primary_keys:
        expected_df[f"标准化_{key_col}"] = expected_df[key_col].apply(_standardize_key_value)
        result_df[f"标准化_{key_col}"] = result_df[key_col].apply(_standardize_key_value)

    # 如果没有指定输出文件, 使用默认路径
    if output_file is None:
        output_dir = Path(result_file).parent
        output_file = str(output_dir / "差异对比.xlsx")

    # 创建差异报告
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "差异对比"

    # 写入表头（动态主键列名）
    key_header_names = []
    for i, key in enumerate(primary_keys[:3]):
        key_header_names.append(f"主键{i+1}({key})")
    while len(key_header_names) < 3:
        key_header_names.append(f"主键{len(key_header_names)+1}")
    headers = key_header_names + ["字段名", "预期值", "生成值", "差异", "差异率", "匹配状态"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")

    row_idx = 2
    total_differences = 0
    unmatched_expected = 0
    unmatched_result = 0
    total_cells = 0  # 总单元格数
    matched_cells = 0  # 匹配的单元格数

    # 按字段分类的差异样本（用于AI修正）
    # 结构: {字段名: {"formula": 公式, "count": 差异数量}}
    field_diff_samples = {}

    # 只对比预期结果中存在的列（忽略生成结果中的中间列/过程列）
    # 排除主键列、标准化列和匹配键
    exclude_cols = {"匹配键"}
    for key in primary_keys:
        exclude_cols.add(key)
        exclude_cols.add(f"标准化_{key}")
    compare_columns = set(expected_df.columns) - exclude_cols
    # 检查生成结果中缺少的列
    missing_in_result = compare_columns - set(result_df.columns)
    if missing_in_result:
        logger.warning(f"生成结果中缺少以下列: {missing_in_result}")
        # 缺少的列应该计入差异！每缺少一列，对每行都是差异
        total_differences += len(missing_in_result) * len(expected_df)
        logger.info(f"缺少 {len(missing_in_result)} 列 × {len(expected_df)} 行 = {len(missing_in_result) * len(expected_df)} 处差异")
    # 实际对比的列（预期结果中的列，且在生成结果中也存在）
    common_columns = compare_columns & set(result_df.columns)

    # 计算总单元格数（预期结果的行数 × 需要对比的列数，排除主键列）
    compare_data_columns = common_columns - set(primary_keys) - {"匹配键"}
    for key in primary_keys:
        compare_data_columns.discard(f"标准化_{key}")
    total_cells = len(expected_df) * len(compare_data_columns)
    # 加上缺失列的单元格数
    total_cells += len(missing_in_result) * len(expected_df)

    # 创建复合匹配键
    if primary_keys:
        standardized_key_cols = [f"标准化_{k}" for k in primary_keys]

        # 检查标准化列是否都存在
        if all(col in expected_df.columns for col in standardized_key_cols):
            expected_df["匹配键"] = expected_df[standardized_key_cols].astype(str).agg("_".join, axis=1)
        else:
            logger.warning("预期文件缺少某些标准化主键列，使用原始列")
            expected_df["匹配键"] = expected_df[primary_keys].astype(str).agg("_".join, axis=1)

        if all(col in result_df.columns for col in standardized_key_cols):
            result_df["匹配键"] = result_df[standardized_key_cols].astype(str).agg("_".join, axis=1)
        else:
            logger.warning("生成文件缺少某些标准化主键列，使用原始列")
            result_df["匹配键"] = result_df[primary_keys].astype(str).agg("_".join, axis=1)

        # 使用匹配键进行合并
        merged_df = pd.merge(
            expected_df, result_df,
            on="匹配键",
            how="outer",
            suffixes=("_expected", "_result"),
            indicator=True
        )

        logger.info(f"匹配结果统计:")
        logger.info(f"  - 两边都有: {len(merged_df[merged_df['_merge'] == 'both'])} 条")
        logger.info(f"  - 仅预期有: {len(merged_df[merged_df['_merge'] == 'left_only'])} 条")
        logger.info(f"  - 仅生成有: {len(merged_df[merged_df['_merge'] == 'right_only'])} 条")

        # 遍历每一行
        for idx, row in merged_df.iterrows():
            merge_status = row.get("_merge", "")

            # 动态获取主键值（用于差异报告的前3列）
            key_values = []
            for key in primary_keys[:3]:  # 最多显示前3个主键
                std_key = f"标准化_{key}"
                key_val = row.get(f"{std_key}_expected", "") or row.get(f"{std_key}_result", "")
                if not key_val:  # 如果标准化列不存在，使用原始列
                    key_val = row.get(f"{key}_expected", "") or row.get(f"{key}_result", "")
                key_values.append(key_val)

            # 补齐到3列（兼容现有报告格式）
            while len(key_values) < 3:
                key_values.append("")

            if merge_status == "left_only":
                # 生成文件缺少这一行，所有单元格都算差异
                unmatched_expected += 1
                ws.cell(row=row_idx, column=1, value=key_values[0])
                ws.cell(row=row_idx, column=2, value=key_values[1])
                ws.cell(row=row_idx, column=3, value=key_values[2])
                ws.cell(row=row_idx, column=4, value="整行")
                ws.cell(row=row_idx, column=5, value="存在")
                ws.cell(row=row_idx, column=6, value="不存在")
                ws.cell(row=row_idx, column=9, value="仅预期有")
                for c_idx in range(1, 10):
                    ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                        start_color="FFFF99", end_color="FFFF99", fill_type="solid"
                    )
                row_idx += 1
                # 修复：缺少一整行，应该增加该行所有数据列的单元格数
                total_differences += len(compare_data_columns)
                continue
            elif merge_status == "right_only":
                # 生成文件多出这一行（不影响分数，因为total_cells基于预期文件）
                unmatched_result += 1
                ws.cell(row=row_idx, column=1, value=key_values[0])
                ws.cell(row=row_idx, column=2, value=key_values[1])
                ws.cell(row=row_idx, column=3, value=key_values[2])
                ws.cell(row=row_idx, column=4, value="整行")
                ws.cell(row=row_idx, column=5, value="不存在")
                ws.cell(row=row_idx, column=6, value="存在")
                ws.cell(row=row_idx, column=9, value="仅生成有")
                for c_idx in range(1, 10):
                    ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                        start_color="FFCC99", end_color="FFCC99", fill_type="solid"
                    )
                row_idx += 1
                total_differences += 1
                continue

            # 遍历每个字段进行对比
            for col in common_columns:
                # 跳过主键列和匹配键
                if col in primary_keys or col == "匹配键":
                    continue
                # 跳过标准化主键列
                if any(col == f"标准化_{k}" for k in primary_keys):
                    continue

                expected_col = f"{col}_expected" if f"{col}_expected" in merged_df.columns else col
                result_col = f"{col}_result" if f"{col}_result" in merged_df.columns else col

                expected_value = row.get(expected_col, 0)
                result_value = row.get(result_col, 0)

                if pd.isna(expected_value):
                    expected_value = 0
                if pd.isna(result_value):
                    result_value = 0

                try:
                    expected_num = float(expected_value) if expected_value != "" else 0
                    result_num = float(result_value) if result_value != "" else 0
                    difference = result_num - expected_num

                    if abs(difference) > 0.01:
                        total_differences += 1
                        diff_rate_str = f"{(difference / expected_num * 100):.2f}%" if expected_num != 0 else "N/A"

                        # 收集差异样本（每个字段只保留公式和差异数量）
                        if col not in field_diff_samples:
                            field_diff_samples[col] = {
                                "formula": result_formulas.get(col, ""),  # 该字段使用的公式
                                "count": 1
                            }
                        else:
                            field_diff_samples[col]["count"] += 1

                        ws.cell(row=row_idx, column=1, value=key_values[0])
                        ws.cell(row=row_idx, column=2, value=key_values[1])
                        ws.cell(row=row_idx, column=3, value=key_values[2])
                        ws.cell(row=row_idx, column=4, value=col)
                        ws.cell(row=row_idx, column=5, value=expected_num)
                        ws.cell(row=row_idx, column=6, value=result_num)
                        ws.cell(row=row_idx, column=7, value=difference)
                        ws.cell(row=row_idx, column=8, value=diff_rate_str)
                        ws.cell(row=row_idx, column=9, value="匹配成功")

                        if abs(difference) > 100 or (expected_num != 0 and abs(difference / expected_num) > 0.1):
                            for c_idx in range(1, 10):
                                ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                                    start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"
                                )
                        row_idx += 1
                    else:
                        # 单元格匹配成功
                        matched_cells += 1
                except (ValueError, TypeError):
                    if str(expected_value) != str(result_value):
                        total_differences += 1

                        # 收集文本差异样本（只保留公式和差异数量）
                        if col not in field_diff_samples:
                            field_diff_samples[col] = {
                                "formula": result_formulas.get(col, ""),
                                "count": 1
                            }
                        else:
                            field_diff_samples[col]["count"] += 1

                        ws.cell(row=row_idx, column=1, value=key_values[0])
                        ws.cell(row=row_idx, column=2, value=key_values[1])
                        ws.cell(row=row_idx, column=3, value=key_values[2])
                        ws.cell(row=row_idx, column=4, value=col)
                        ws.cell(row=row_idx, column=5, value=str(expected_value))
                        ws.cell(row=row_idx, column=6, value=str(result_value))
                        ws.cell(row=row_idx, column=7, value="文本不同")
                        ws.cell(row=row_idx, column=9, value="匹配成功")
                        row_idx += 1
                    else:
                        # 文本值匹配成功
                        matched_cells += 1

    # 调整列宽
    column_widths = [15, 15, 12, 20, 15, 15, 12, 10, 15]
    for col_idx, width in enumerate(column_widths, start=1):
        ws.column_dimensions[chr(64 + col_idx)].width = width

    wb.save(output_file)
    logger.info(f"差异对比文件已保存: {output_file}")
    logger.info(f"共发现 {total_differences} 处差异")
    logger.info(f"其中: 仅预期有 {unmatched_expected} 条, 仅生成有 {unmatched_result} 条")

    # 计算匹配率
    match_rate = matched_cells / total_cells if total_cells > 0 else 0.0
    logger.info(f"单元格匹配: {matched_cells}/{total_cells} = {match_rate:.2%}")

    return {
        "total_differences": total_differences,
        "unmatched_expected": unmatched_expected,
        "unmatched_result": unmatched_result,
        "output_file": output_file,
        "success": total_differences == 0,
        "total_cells": total_cells,
        "matched_cells": matched_cells,
        "match_rate": match_rate,
        "field_diff_samples": field_diff_samples  # 按字段分类的差异样本
    }


def compare_dataframes(
    result_df: pd.DataFrame,
    expected_df: pd.DataFrame,
    output_file: str = "差异对比.xlsx",
    primary_keys: Optional[List[str]] = None
) -> Dict[str, Any]:
    """对比两个DataFrame的差异(用于沙箱内调用)

    Args:
        result_df: 生成的结果DataFrame
        expected_df: 预期的结果DataFrame
        output_file: 差异报告输出文件路径
        primary_keys: 主键列名列表，默认 ["工号", "中文姓名"]

    Returns:
        对比统计结果字典
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    logger.info("开始差异对比...")

    # 标准化列名（去除换行符、统一空格）
    result_df.columns = [_standardize_column_name(col) for col in result_df.columns]
    expected_df.columns = [_standardize_column_name(col) for col in expected_df.columns]

    logger.info(f"生成结果: {len(result_df)} 行")
    logger.info(f"预期结果: {len(expected_df)} 行")

    # 确定主键列
    if primary_keys is None:
        primary_keys = ["工号", "中文姓名"]

    available_keys = []
    for key in primary_keys:
        if key in expected_df.columns and key in result_df.columns:
            available_keys.append(key)
    if not available_keys:
        first_col = expected_df.columns[0] if len(expected_df.columns) > 0 else None
        if first_col and first_col in result_df.columns:
            available_keys = [first_col]
            logger.warning(f"所有指定主键都不存在，回退到第一列: {first_col}")
    primary_keys = available_keys

    # 标准化主键列
    expected_df = expected_df.copy()
    result_df = result_df.copy()
    for key_col in primary_keys:
        expected_df[f"标准化_{key_col}"] = expected_df[key_col].apply(_standardize_key_value)
        result_df[f"标准化_{key_col}"] = result_df[key_col].apply(_standardize_key_value)

    # 创建差异报告
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "差异对比"

    # 写入表头（动态主键列名）
    key_header_names = []
    for i, key in enumerate(primary_keys[:3]):
        key_header_names.append(f"主键{i+1}({key})")
    while len(key_header_names) < 3:
        key_header_names.append(f"主键{len(key_header_names)+1}")
    headers = key_header_names + ["字段名", "预期值", "生成值", "差异", "差异率", "匹配状态"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")

    row_idx = 2
    total_differences = 0
    unmatched_expected = 0
    unmatched_result = 0
    total_cells = 0  # 总单元格数
    matched_cells = 0  # 匹配的单元格数

    # 只对比预期结果中存在的列（忽略生成结果中的中间列/过程列）
    # 排除主键列、标准化列和匹配键
    exclude_cols = {"匹配键"}
    for key in primary_keys:
        exclude_cols.add(key)
        exclude_cols.add(f"标准化_{key}")
    compare_columns = set(expected_df.columns) - exclude_cols
    # 检查生成结果中缺少的列
    missing_in_result = compare_columns - set(result_df.columns)
    if missing_in_result:
        logger.warning(f"生成结果中缺少以下列: {missing_in_result}")
        # 缺少的列应该计入差异！每缺少一列，对每行都是差异
        total_differences += len(missing_in_result) * len(expected_df)
        logger.info(f"缺少 {len(missing_in_result)} 列 × {len(expected_df)} 行 = {len(missing_in_result) * len(expected_df)} 处差异")
    # 实际对比的列（预期结果中的列，且在生成结果中也存在）
    common_columns = compare_columns & set(result_df.columns)

    # 计算总单元格数（预期结果的行数 × 需要对比的列数，排除主键列）
    compare_data_columns = common_columns - set(primary_keys) - {"匹配键"}
    for key in primary_keys:
        compare_data_columns.discard(f"标准化_{key}")
    total_cells = len(expected_df) * len(compare_data_columns)
    # 加上缺失列的单元格数
    total_cells += len(missing_in_result) * len(expected_df)

    # 创建复合匹配键
    if primary_keys:
        standardized_key_cols = [f"标准化_{k}" for k in primary_keys]

        # 检查标准化列是否都存在
        if all(col in expected_df.columns for col in standardized_key_cols):
            expected_df["匹配键"] = expected_df[standardized_key_cols].astype(str).agg("_".join, axis=1)
        else:
            logger.warning("预期文件缺少某些标准化主键列，使用原始列")
            expected_df["匹配键"] = expected_df[primary_keys].astype(str).agg("_".join, axis=1)

        if all(col in result_df.columns for col in standardized_key_cols):
            result_df["匹配键"] = result_df[standardized_key_cols].astype(str).agg("_".join, axis=1)
        else:
            logger.warning("生成文件缺少某些标准化主键列，使用原始列")
            result_df["匹配键"] = result_df[primary_keys].astype(str).agg("_".join, axis=1)

        # 使用匹配键进行合并
        merged_df = pd.merge(
            expected_df, result_df,
            on="匹配键",
            how="outer",
            suffixes=("_expected", "_result"),
            indicator=True
        )

        logger.info(f"匹配结果统计:")
        logger.info(f"  - 两边都有: {len(merged_df[merged_df['_merge'] == 'both'])} 条")
        logger.info(f"  - 仅预期有: {len(merged_df[merged_df['_merge'] == 'left_only'])} 条")
        logger.info(f"  - 仅生成有: {len(merged_df[merged_df['_merge'] == 'right_only'])} 条")

        # 遍历每一行进行对比(与compare_excel_files相同的逻辑)
        for idx, row in merged_df.iterrows():
            merge_status = row.get("_merge", "")

            # 动态获取主键值
            key_values = []
            for key in primary_keys[:3]:
                std_key = f"标准化_{key}"
                key_val = row.get(f"{std_key}_expected", "") or row.get(f"{std_key}_result", "")
                if not key_val:
                    key_val = row.get(f"{key}_expected", "") or row.get(f"{key}_result", "")
                key_values.append(key_val)
            while len(key_values) < 3:
                key_values.append("")

            if merge_status == "left_only":
                # 生成文件缺少这一行，所有单元格都算差异
                unmatched_expected += 1
                ws.cell(row=row_idx, column=1, value=key_values[0])
                ws.cell(row=row_idx, column=2, value=key_values[1])
                ws.cell(row=row_idx, column=3, value=key_values[2])
                ws.cell(row=row_idx, column=4, value="整行")
                ws.cell(row=row_idx, column=5, value="存在")
                ws.cell(row=row_idx, column=6, value="不存在")
                ws.cell(row=row_idx, column=9, value="仅预期有")
                for c_idx in range(1, 10):
                    ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                        start_color="FFFF99", end_color="FFFF99", fill_type="solid"
                    )
                row_idx += 1
                # 修复：缺少一整行，应该增加该行所有数据列的单元格数
                total_differences += len(compare_data_columns)
                continue
            elif merge_status == "right_only":
                # 生成文件多出这一行（不影响分数，因为total_cells基于预期文件）
                unmatched_result += 1
                ws.cell(row=row_idx, column=1, value=key_values[0])
                ws.cell(row=row_idx, column=2, value=key_values[1])
                ws.cell(row=row_idx, column=3, value=key_values[2])
                ws.cell(row=row_idx, column=4, value="整行")
                ws.cell(row=row_idx, column=5, value="不存在")
                ws.cell(row=row_idx, column=6, value="存在")
                ws.cell(row=row_idx, column=9, value="仅生成有")
                for c_idx in range(1, 10):
                    ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                        start_color="FFCC99", end_color="FFCC99", fill_type="solid"
                    )
                row_idx += 1
                total_differences += 1
                continue

            for col in common_columns:
                # 跳过主键列和匹配键
                if col in primary_keys or col == "匹配键":
                    continue
                # 跳过标准化主键列
                if any(col == f"标准化_{k}" for k in primary_keys):
                    continue

                expected_col = f"{col}_expected" if f"{col}_expected" in merged_df.columns else col
                result_col = f"{col}_result" if f"{col}_result" in merged_df.columns else col

                expected_value = row.get(expected_col, 0)
                result_value = row.get(result_col, 0)

                if pd.isna(expected_value):
                    expected_value = 0
                if pd.isna(result_value):
                    result_value = 0

                try:
                    expected_num = float(expected_value) if expected_value != "" else 0
                    result_num = float(result_value) if result_value != "" else 0
                    difference = result_num - expected_num

                    if abs(difference) > 0.01:
                        total_differences += 1
                        diff_rate_str = f"{(difference / expected_num * 100):.2f}%" if expected_num != 0 else "N/A"

                        ws.cell(row=row_idx, column=1, value=key_values[0])
                        ws.cell(row=row_idx, column=2, value=key_values[1])
                        ws.cell(row=row_idx, column=3, value=key_values[2])
                        ws.cell(row=row_idx, column=4, value=col)
                        ws.cell(row=row_idx, column=5, value=expected_num)
                        ws.cell(row=row_idx, column=6, value=result_num)
                        ws.cell(row=row_idx, column=7, value=difference)
                        ws.cell(row=row_idx, column=8, value=diff_rate_str)
                        ws.cell(row=row_idx, column=9, value="匹配成功")

                        if abs(difference) > 100 or (expected_num != 0 and abs(difference / expected_num) > 0.1):
                            for c_idx in range(1, 10):
                                ws.cell(row=row_idx, column=c_idx).fill = PatternFill(
                                    start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"
                                )
                        row_idx += 1
                    else:
                        # 单元格匹配成功
                        matched_cells += 1
                except (ValueError, TypeError):
                    if str(expected_value) != str(result_value):
                        total_differences += 1
                        ws.cell(row=row_idx, column=1, value=key_values[0])
                        ws.cell(row=row_idx, column=2, value=key_values[1])
                        ws.cell(row=row_idx, column=3, value=key_values[2])
                        ws.cell(row=row_idx, column=4, value=col)
                        ws.cell(row=row_idx, column=5, value=str(expected_value))
                        ws.cell(row=row_idx, column=6, value=str(result_value))
                        ws.cell(row=row_idx, column=7, value="文本不同")
                        ws.cell(row=row_idx, column=9, value="匹配成功")
                        row_idx += 1
                    else:
                        # 文本值匹配成功
                        matched_cells += 1

    # 调整列宽
    column_widths = [15, 15, 12, 20, 15, 15, 12, 10, 15]
    for col_idx, width in enumerate(column_widths, start=1):
        ws.column_dimensions[chr(64 + col_idx)].width = width

    wb.save(output_file)
    logger.info(f"差异对比文件已保存: {output_file}")
    logger.info(f"共发现 {total_differences} 处差异")

    # 计算匹配率
    match_rate = matched_cells / total_cells if total_cells > 0 else 0.0
    logger.info(f"单元格匹配: {matched_cells}/{total_cells} = {match_rate:.2%}")

    return {
        "total_differences": total_differences,
        "unmatched_expected": unmatched_expected,
        "unmatched_result": unmatched_result,
        "output_file": output_file,
        "success": total_differences == 0,
        "total_cells": total_cells,
        "matched_cells": matched_cells,
        "match_rate": match_rate
    }
