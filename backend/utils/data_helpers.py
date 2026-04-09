"""
数据处理公共辅助函数
这些函数可以在沙箱内外统一使用
"""

import pandas as pd
from typing import Dict, Any, List, Optional


# ============================================================
# 同义词映射表 - 人力资源行业常见列名变体
# ============================================================

SYNONYM_GROUPS = {
    "员工标识": ["工号", "员工编号", "员工编码", "员工号", "雇员工号", "雇员编号", "人员编号", "职工号", "编号", "ID"],
    "姓名": ["姓名", "员工姓名", "雇员姓名", "中文名", "中文姓名", "名字", "员工名称"],
    "部门": ["部门", "部门名称", "所属部门", "归属部门", "组织", "组织名称"],
    "基本工资": ["基本工资", "基本薪资", "底薪", "月薪", "基础工资", "固定工资", "标准工资"],
    "绩效工资": ["绩效工资", "绩效", "绩效奖金", "考核工资", "业绩工资"],
    "加班费": ["加班费", "加班工资", "OT费用", "加班补贴"],
    "出勤天数": ["出勤天数", "实际出勤", "出勤日数", "应出勤天数", "工作天数"],
    "加班小时": ["加班小时", "加班时数", "加班工时", "OT小时"],
    "社保": ["社保", "社保个人", "社保扣款", "社会保险", "五险一金个人"],
    "公积金": ["公积金", "住房公积金", "公积金个人"],
    "个税": ["个税", "个人所得税", "所得税", "税款"],
    "应发工资": ["应发工资", "应发合计", "应发金额", "税前工资"],
    "实发工资": ["实发工资", "实发合计", "实发金额", "到手工资", "净工资"],
}


def find_column(df: pd.DataFrame, target_name: str, synonyms: List[str] = None) -> Optional[str]:
    """在DataFrame中查找语义匹配的列名

    Args:
        df: DataFrame
        target_name: 目标列名（可能是规则中的名称）
        synonyms: 同义词列表，如["工号", "员工编号"]

    Returns:
        实际存在的列名，如果都不存在返回None
    """
    # 1. 精确匹配
    if target_name in df.columns:
        return target_name

    # 2. 同义词匹配
    if synonyms:
        for syn in synonyms:
            if syn in df.columns:
                return syn

    # 3. 从全局同义词表查找
    for group_name, group_synonyms in SYNONYM_GROUPS.items():
        if target_name in group_synonyms:
            for syn in group_synonyms:
                if syn in df.columns:
                    return syn

    # 4. 包含匹配
    for col in df.columns:
        if target_name in col or col in target_name:
            return col

    return None


def safe_get_column(df: pd.DataFrame, col_name: str, default=0, synonyms: List[str] = None):
    """安全获取DataFrame列，支持同义词查找

    Args:
        df: DataFrame
        col_name: 目标列名
        default: 默认值（列不存在时使用）
        synonyms: 同义词列表

    Returns:
        列数据(Series)，如果列不存在则返回填充默认值的Series
    """
    actual_col = find_column(df, col_name, synonyms)
    if actual_col:
        return df[actual_col].fillna(default)
    else:
        print(f"警告: 列 '{col_name}' 不存在，可用列: {list(df.columns)}")
        # 返回与DataFrame长度相同的Series，而不是单个值
        # 这样可以安全地参与后续的计算
        return pd.Series([default] * len(df), index=df.index)


def convert_region_to_dataframe(region) -> pd.DataFrame:
    """将ExcelRegion转换为pandas DataFrame

    Args:
        region: ExcelRegion对象，包含head_data和data

    Returns:
        转换后的DataFrame，列名为中文表头名称
    """
    if not region.data:
        return pd.DataFrame()

    # 创建列字母到列名的反向映射
    col_letter_to_name = {v: k for k, v in region.head_data.items()}

    # 转换数据
    converted_data = []
    for row in region.data:
        new_row = {}
        for col_letter, value in row.items():
            col_name = col_letter_to_name.get(col_letter, col_letter)
            new_row[col_name] = value
        converted_data.append(new_row)

    # 创建DataFrame
    columns = list(region.head_data.keys())
    return pd.DataFrame(converted_data, columns=columns)


def normalize_emp_code(emp_code) -> str:
    """标准化工号：转换为8位字符串，不足前面补0

    示例：
    - "123" -> "00000123"
    - 12345 -> "00012345"
    """
    if pd.isna(emp_code) or emp_code == "":
        return ""
    code_str = str(emp_code).strip()
    if code_str.isdigit():
        return code_str.zfill(8)
    return code_str


def print_available_columns(data_store: Dict[str, Any]):
    """打印所有源文件的列名，用于调试"""
    print("=" * 50)
    print("【源文件列名】")
    for filename, file_data in data_store.get("files", {}).items():
        if isinstance(file_data, pd.DataFrame):
            print(f"  {filename}: {list(file_data.columns)}")
    print("=" * 50)


def load_files_to_dataframes(data_store: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
    """从data_store提取所有DataFrame，返回{文件名: DataFrame}"""
    result = {}
    for filename, file_data in data_store.get("files", {}).items():
        if isinstance(file_data, pd.DataFrame):
            result[filename] = file_data
    return result


def make_unique_sheet_key(name: str, existing_keys: set, max_len: int = 31) -> str:
    """将 sheet key 截断到 max_len 字符，并在碰撞时追加 _2, _3... 后缀保证唯一。

    截断是因为 Excel sheet 名最长 31 字符，source_data 的 key 与 sheet 名保持一致。
    会自动将生成的 key 加入 existing_keys 集合。
    """
    if len(name) > max_len:
        name = name[:max_len]
    base = name
    counter = 2
    while name in existing_keys:
        suffix = f"_{counter}"
        name = base[:max_len - len(suffix)] + suffix
        counter += 1
    existing_keys.add(name)
    return name
