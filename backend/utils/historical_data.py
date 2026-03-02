"""
历史计算数据查询工具

提供对租户历史计算结果的查询能力，支持按薪资年月加载数据，
并提供汇总、平均、计数、分组等聚合查询方法。
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Union

import pandas as pd

logger = logging.getLogger(__name__)


class HistoricalDataProvider:
    """历史计算数据查询工具

    在沙箱中通过 history_provider 变量访问，提供对历史薪资计算结果的查询。

    用法示例:
        # 获取2025年1月的基本工资汇总
        total = history_provider.get_sum('基本工资', 2025, months=[1])

        # 获取2025年1-6月某员工的历史数据
        df = history_provider.get_employee_history('00001234', 2025, months=[1,2,3,4,5,6])

        # 按部门分组获取平均工资
        df = history_provider.get_data(['部门', '基本工资'], 2025, months=[1], group_by='部门')
    """

    def __init__(self, tenant_id: str, base_dir: str = "tenants"):
        self.tenant_id = tenant_id
        self.tenant_dir = Path(base_dir) / tenant_id
        self.history_dir = self.tenant_dir / "history"
        self._cache: Dict[str, pd.DataFrame] = {}
        self._history_meta = self._load_meta()

    def _load_meta(self) -> Dict[str, Any]:
        history_file = self.tenant_dir / "calculation_history.json"
        if history_file.exists():
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {"records": []}

    def get_history_file(self, year: int, month: int) -> Optional[str]:
        """获取指定年月的历史输出文件路径"""
        path = self.history_dir / f"{year}_{month:02d}" / "output.xlsx"
        return str(path) if path.exists() else None

    def get_available_months(self, year: int) -> List[int]:
        """获取指定年份有历史数据的月份列表"""
        months = []
        for r in self._history_meta.get("records", []):
            if r["salary_year"] == year:
                months.append(r["salary_month"])
        return sorted(set(months))

    def load_history(self, year: int, month: int, sheet_name: str = None) -> Optional[pd.DataFrame]:
        """加载指定年月的历史数据为DataFrame

        Args:
            year: 薪资年份
            month: 薪资月份
            sheet_name: Sheet名称，默认读取第一个sheet

        Returns:
            DataFrame，如果文件不存在返回None
        """
        cache_key = f"{year}_{month:02d}_{sheet_name or 'default'}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        file_path = self.get_history_file(year, month)
        if not file_path:
            return None

        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            self._cache[cache_key] = df
            return df
        except Exception as e:
            logger.warning(f"加载历史数据失败 {year}年{month}月: {e}")
            return None

    def _load_multi_months(self, year: int, months: List[int] = None, sheet_name: str = None) -> Optional[pd.DataFrame]:
        """加载多个月份的数据并合并，添加_salary_month列"""
        if months is None:
            months = self.get_available_months(year)
        if not months:
            return None

        frames = []
        for m in months:
            df = self.load_history(year, m, sheet_name)
            if df is not None:
                df = df.copy()
                df["_salary_month"] = m
                frames.append(df)

        if not frames:
            return None
        return pd.concat(frames, ignore_index=True)

    def _apply_condition(self, df: pd.DataFrame, condition) -> pd.DataFrame:
        """应用筛选条件

        condition 格式:
            单条件: {"field": "部门", "op": "==", "value": "销售部"}
            多条件: [{"field": ..., "op": ..., "value": ...}, ...]
            op 支持: ==, !=, >, <, >=, <=, in, not_in, contains
        """
        if condition is None:
            return df

        if isinstance(condition, dict):
            condition = [condition]

        for cond in condition:
            field = cond["field"]
            op = cond.get("op", "==")
            value = cond["value"]

            if field not in df.columns:
                continue

            if op == "==":
                df = df[df[field] == value]
            elif op == "!=":
                df = df[df[field] != value]
            elif op == ">":
                df = df[df[field] > value]
            elif op == "<":
                df = df[df[field] < value]
            elif op == ">=":
                df = df[df[field] >= value]
            elif op == "<=":
                df = df[df[field] <= value]
            elif op == "in":
                df = df[df[field].isin(value)]
            elif op == "not_in":
                df = df[~df[field].isin(value)]
            elif op == "contains":
                df = df[df[field].astype(str).str.contains(str(value), na=False)]

        return df

    def get_sum(self, field: str, year: int, months: List[int] = None,
                condition=None, sheet_name: str = None) -> float:
        """获取指定字段的汇总值

        Args:
            field: 字段名（列名）
            year: 薪资年份
            months: 月份列表，None表示该年所有可用月份
            condition: 筛选条件
            sheet_name: Sheet名称
        """
        df = self._load_multi_months(year, months, sheet_name)
        if df is None or field not in df.columns:
            return 0.0
        df = self._apply_condition(df, condition)
        return float(pd.to_numeric(df[field], errors='coerce').sum())

    def get_avg(self, field: str, year: int, months: List[int] = None,
                condition=None, sheet_name: str = None) -> float:
        """获取指定字段的平均值"""
        df = self._load_multi_months(year, months, sheet_name)
        if df is None or field not in df.columns:
            return 0.0
        df = self._apply_condition(df, condition)
        result = pd.to_numeric(df[field], errors='coerce').mean()
        return float(result) if pd.notna(result) else 0.0

    def get_count(self, field: str, year: int, months: List[int] = None,
                  condition=None, sheet_name: str = None) -> int:
        """获取指定字段的非空计数"""
        df = self._load_multi_months(year, months, sheet_name)
        if df is None or field not in df.columns:
            return 0
        df = self._apply_condition(df, condition)
        return int(df[field].notna().sum())

    def get_data(self, fields: List[str] = None, year: int = None,
                 months: List[int] = None, condition=None,
                 group_by: Union[str, List[str]] = None,
                 sort_by: Union[str, List[str]] = None,
                 agg: str = "sum", sheet_name: str = None) -> Optional[pd.DataFrame]:
        """灵活查询历史数据

        Args:
            fields: 要返回的字段列表，None返回全部
            year: 薪资年份
            months: 月份列表
            condition: 筛选条件
            group_by: 分组字段
            sort_by: 排序字段
            agg: 聚合方式（sum/mean/count/min/max），仅group_by时生效
            sheet_name: Sheet名称
        """
        df = self._load_multi_months(year, months, sheet_name)
        if df is None:
            return None

        df = self._apply_condition(df, condition)

        if fields:
            valid_fields = [f for f in fields if f in df.columns]
            if group_by:
                gb = [group_by] if isinstance(group_by, str) else group_by
                valid_fields = list(set(gb + valid_fields))
            df = df[valid_fields]

        if group_by:
            numeric_cols = df.select_dtypes(include='number').columns.tolist()
            gb = [group_by] if isinstance(group_by, str) else group_by
            numeric_cols = [c for c in numeric_cols if c not in gb]
            if numeric_cols:
                df = df.groupby(gb, as_index=False)[numeric_cols].agg(agg)

        if sort_by:
            sb = [sort_by] if isinstance(sort_by, str) else sort_by
            valid_sb = [s for s in sb if s in df.columns]
            if valid_sb:
                df = df.sort_values(valid_sb)

        return df

    def get_employee_history(self, emp_code: str, year: int,
                             months: List[int] = None, fields: List[str] = None,
                             emp_code_col: str = "工号",
                             sheet_name: str = None) -> Optional[pd.DataFrame]:
        """获取指定员工的历史数据

        Args:
            emp_code: 员工工号
            year: 薪资年份
            months: 月份列表
            fields: 要返回的字段列表
            emp_code_col: 工号列名
            sheet_name: Sheet名称
        """
        condition = {"field": emp_code_col, "op": "==", "value": emp_code}
        return self.get_data(fields=fields, year=year, months=months,
                             condition=condition, sheet_name=sheet_name)
