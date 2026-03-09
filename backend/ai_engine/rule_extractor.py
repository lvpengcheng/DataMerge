"""
规则提取器 - 从规则文档中提取数据清洗规则和警告规则
"""

import re
from typing import Dict, List, Any, Tuple


class RuleExtractor:
    """从规则文档中提取结构化的数据清洗规则和警告规则"""

    def __init__(self):
        pass

    def extract_rules(self, rules_content: str) -> Dict[str, Any]:
        """
        从规则文档中提取数据清洗规则和警告规则

        Args:
            rules_content: 规则文档内容

        Returns:
            包含清洗规则和警告规则的字典
        """
        result = {
            "data_cleaning_rules": [],
            "warning_rules": [],
            "import_validation_rules": {}
        }

        # 提取数据清洗规则
        cleaning_rules = self._extract_data_cleaning_rules(rules_content)
        result["data_cleaning_rules"] = cleaning_rules

        # 提取警告规则
        warning_rules = self._extract_warning_rules(rules_content)
        result["warning_rules"] = warning_rules

        # 提取导入校验规则
        import_rules = self._extract_import_validation_rules(rules_content)
        result["import_validation_rules"] = import_rules

        return result

    def _extract_data_cleaning_rules(self, content: str) -> List[Dict[str, Any]]:
        """提取数据清洗规则"""
        rules = []

        # 查找"数据清洗规则"章节
        # 支持多种格式：## 六、数据清洗规则  ## 三、数据清洗规则  ## 数据清洗规则
        cleaning_section_match = re.search(
            r'#+\s*(?:[一二三四五六七八九十\d]+[、.．]\s*)?数据清洗规则\s*\n(.*?)(?=\n#+\s*|\Z)',
            content,
            re.DOTALL | re.IGNORECASE
        )

        if not cleaning_section_match:
            return rules

        cleaning_content = cleaning_section_match.group(1)

        # 提取每条规则（以 - 或 数字. 开头）
        rule_lines = re.findall(r'^[\s]*[-•]\s*(.+?)$', cleaning_content, re.MULTILINE)

        for rule_text in rule_lines:
            rule_text = rule_text.strip()
            if not rule_text or rule_text.startswith('---'):
                continue

            parsed_rule = self._parse_cleaning_rule(rule_text)
            if parsed_rule:
                rules.append(parsed_rule)

        return rules

    def _parse_cleaning_rule(self, rule_text: str) -> Dict[str, Any]:
        """解析单条数据清洗规则"""
        rule = {
            "original_text": rule_text,
            "tables": [],
            "conditions": [],
            "action": "exclude"  # 默认动作是排除
        }

        # 提取表名（【表名】格式）
        table_matches = re.findall(r'【([^】]+)】', rule_text)
        rule["tables"] = table_matches

        # 解析常见的规则模式
        if "重复" in rule_text:
            rule["conditions"].append({
                "type": "duplicate",
                "field": self._extract_field_name(rule_text)
            })

        if "为空" in rule_text or "空" in rule_text:
            fields = self._extract_field_names_for_empty_check(rule_text)
            for field in fields:
                rule["conditions"].append({
                    "type": "empty",
                    "field": field
                })

        if "必须包含" in rule_text:
            rule["conditions"].append({
                "type": "must_contain",
                "field": self._extract_field_name(rule_text),
                "values": self._extract_quoted_values(rule_text)
            })

        if "必须与" in rule_text and "一致" in rule_text:
            rule["conditions"].append({
                "type": "must_match",
                "description": rule_text
            })

        if "不能作为计薪雇员" in rule_text or "不计入计薪雇员" in rule_text:
            rule["action"] = "exclude_from_payroll"

        if "不参与报表统计" in rule_text:
            rule["action"] = "exclude_from_report"

        return rule

    def _extract_warning_rules(self, content: str) -> List[Dict[str, Any]]:
        """提取警告规则"""
        rules = []

        # 查找"警告信息"章节
        # 支持多种格式：## 六、警告信息  ## 四、警告信息  ## 警告信息
        warning_section_match = re.search(
            r'#+\s*(?:[一二三四五六七八九十\d]+[、.．]\s*)?警告信息\s*\n(.*?)(?=\n#+\s*|\Z)',
            content,
            re.DOTALL | re.IGNORECASE
        )

        if not warning_section_match:
            return rules

        warning_content = warning_section_match.group(1)

        # 提取每条警告规则
        # 匹配段落（非空行）
        paragraphs = [p.strip() for p in warning_content.split('\n') if p.strip() and not p.strip().startswith('---')]

        for para in paragraphs:
            parsed_rule = self._parse_warning_rule(para)
            if parsed_rule:
                rules.append(parsed_rule)

        return rules

    def _parse_warning_rule(self, rule_text: str) -> Dict[str, Any]:
        """解析单条警告规则"""
        rule = {
            "original_text": rule_text,
            "source_table": None,
            "target_table": None,
            "field": None,
            "message": None
        }

        # 提取表名
        table_matches = re.findall(r'【?([^】\s]+表)】?', rule_text)
        if len(table_matches) >= 2:
            rule["source_table"] = table_matches[0]
            rule["target_table"] = table_matches[1]
        elif len(table_matches) == 1:
            rule["source_table"] = table_matches[0]

        # 提取字段名
        if "中的" in rule_text:
            field_match = re.search(r'中的([^\s必须需要]+)', rule_text)
            if field_match:
                rule["field"] = field_match.group(1)

        # 提取提示信息
        if "提示" in rule_text:
            message_match = re.search(r'提示["""]([^"""]+)["""]', rule_text)
            if message_match:
                rule["message"] = message_match.group(1)

        # 判断警告类型
        if "必须在" in rule_text and "存在" in rule_text:
            rule["type"] = "must_exist_in"
        elif "未匹配" in rule_text:
            rule["type"] = "not_matched"

        return rule

    def _extract_import_validation_rules(self, content: str) -> Dict[str, List[str]]:
        """提取导入校验规则"""
        rules = {}

        # 查找所有"导入校验规则"部分
        validation_sections = re.finditer(
            r'### 源文件\d+[：:]\s*([^\n]+).*?\*\*导入校验规则[：:]\*\*\s*\n(.*?)(?=\n---|###|\Z)',
            content,
            re.DOTALL
        )

        for match in validation_sections:
            table_name = match.group(1).strip()
            rules_content = match.group(2).strip()

            # 提取规则列表
            rule_lines = re.findall(r'^[\s]*[-•]\s*(.+?)$', rules_content, re.MULTILINE)
            rules[table_name] = [r.strip() for r in rule_lines if r.strip()]

        return rules

    def _extract_field_name(self, text: str) -> str:
        """从文本中提取字段名"""
        # 尝试提取引号中的内容
        quoted = re.search(r'["""]([^"""]+)["""]', text)
        if quoted:
            return quoted.group(1)

        # 尝试提取常见字段名
        for field in ["工号", "姓名", "中文姓名", "身份证", "身份证号码", "账单供应商", "客户代码"]:
            if field in text:
                return field

        return ""

    def _extract_field_names_for_empty_check(self, text: str) -> List[str]:
        """提取需要检查为空的字段名"""
        fields = []

        # 查找"或"连接的字段
        if "或" in text:
            parts = text.split("为空")[0].split("或")
            for part in parts:
                field = self._extract_field_name(part)
                if field:
                    fields.append(field)
        else:
            field = self._extract_field_name(text)
            if field:
                fields.append(field)

        return fields

    def _extract_quoted_values(self, text: str) -> List[str]:
        """提取引号中的值"""
        return re.findall(r'["""]([^"""]+)["""]', text)

    def format_rules_for_prompt(self, rules: Dict[str, Any]) -> str:
        """
        将提取的规则格式化为适合AI理解的提示文本

        Args:
            rules: 提取的规则字典

        Returns:
            格式化后的规则文本
        """
        sections = []

        # 格式化数据清洗规则
        if rules.get("data_cleaning_rules"):
            sections.append("## 数据清洗规则\n")
            sections.append("在复制基础数据时，必须应用以下清洗规则过滤数据：\n")
            for i, rule in enumerate(rules["data_cleaning_rules"], 1):
                sections.append(f"{i}. {rule['original_text']}\n")
            sections.append("\n")

        # 格式化警告规则
        if rules.get("warning_rules"):
            sections.append("## 警告信息规则\n")
            sections.append("在处理数据时，需要检查以下情况并生成警告信息：\n")
            for i, rule in enumerate(rules["warning_rules"], 1):
                sections.append(f"{i}. {rule['original_text']}\n")
            sections.append("\n**重要**: 警告信息需要收集到一个列表中，在计算完成后返回。\n\n")

        # 格式化导入校验规则
        if rules.get("import_validation_rules"):
            sections.append("## 导入校验规则\n")
            sections.append("在导入数据时，需要验证以下规则：\n")
            for table_name, table_rules in rules["import_validation_rules"].items():
                sections.append(f"\n### {table_name}\n")
                for rule in table_rules:
                    sections.append(f"- {rule}\n")
            sections.append("\n")

        return "".join(sections)
