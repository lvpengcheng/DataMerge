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
            "import_validation_rules": {},
            "conditional_format_rules": [],
            "precision_rules": []
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

        # 提取条件格式规则（标红、高亮、颜色标记等）
        cf_rules = self._extract_conditional_format_rules(rules_content)
        result["conditional_format_rules"] = cf_rules

        # 提取数值精度规则（小数位数、四舍五入等）
        precision_rules = self._extract_precision_rules(rules_content)
        result["precision_rules"] = precision_rules

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

        # 合并/关联操作识别
        if any(kw in rule_text for kw in ["纵向合并", "上下合并", "拼接", "concat", "追加合并"]):
            rule["action"] = "vertical_concat"
            rule["conditions"].append({
                "type": "merge",
                "merge_type": "vertical_concat",
                "description": rule_text
            })

        if any(kw in rule_text for kw in ["横向合并", "关联合并", "匹配合并", "join", "左连接", "左关联"]):
            rule["action"] = "horizontal_join"
            rule["conditions"].append({
                "type": "merge",
                "merge_type": "horizontal_join",
                "description": rule_text
            })

        if "合并" in rule_text and rule["action"] == "exclude":
            rule["conditions"].append({
                "type": "merge",
                "merge_type": "unspecified",
                "description": rule_text
            })

        return rule

    def _extract_warning_rules(self, content: str) -> List[Dict[str, Any]]:
        """提取警告规则"""
        rules = []

        # 查找"警告信息/警告规则"章节
        # 支持多种格式：## 六、警告信息  ## 四、警告规则  ## 警告信息
        warning_section_match = re.search(
            r'#+\s*(?:[一二三四五六七八九十\d]+[、.．]\s*)?警告(?:信息|规则)\s*\n(.*?)(?=\n#+\s*|\Z)',
            content,
            re.DOTALL | re.IGNORECASE
        )

        if not warning_section_match:
            return rules

        warning_content = warning_section_match.group(1)

        # 提取每条警告规则
        rule_lines = re.findall(r'^[\s]*[-•]\s*(.+?)$', warning_content, re.MULTILINE)

        for line in rule_lines:
            line = line.strip()
            if not line or line.startswith('---'):
                continue
            parsed_rule = self._parse_warning_rule(line)
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

    def _extract_precision_rules(self, content: str) -> List[Dict[str, Any]]:
        """提取数值精度规则（小数位数、四舍五入等）"""
        rules = []

        # 查找"数值精度"或"精度规则"章节
        for section_keyword in [r'数值精度', r'精度规则', r'小数精度', r'数据精度']:
            section_match = re.search(
                r'#+\s*(?:[一二三四五六七八九十\d]+[、.．]\s*)?' + section_keyword + r'\s*\n(.*?)(?=\n#+\s*|\Z)',
                content,
                re.DOTALL | re.IGNORECASE
            )
            if section_match:
                section_content = section_match.group(1)
                rule_lines = re.findall(r'^[\s]*[-•\d.]\s*(.+?)$', section_content, re.MULTILINE)
                for line in rule_lines:
                    line = line.strip()
                    if line and not line.startswith('---'):
                        rules.append({"original_text": line})

        # 全文扫描精度相关描述
        precision_patterns = re.findall(
            r'(.{0,30}(?:保留\d+位小数|四舍五入|取整|精度).{0,50})',
            content
        )
        for match in precision_patterns:
            match = re.sub(r'^[\s\-•]+', '', match).strip()
            if not match:
                continue
            if any(match in r['original_text'] or r['original_text'] in match for r in rules):
                continue
            if match.startswith('#') or match.startswith('//'):
                continue
            rules.append({"original_text": match})

        return rules

    def _extract_conditional_format_rules(self, content: str) -> List[Dict[str, Any]]:
        """提取条件格式规则（标红、高亮、颜色标记、加粗等样式规则）"""
        rules = []

        # 方式1：查找专门的"条件格式"或"特殊规则"章节
        for section_keyword in [r'条件格式[规则]*', r'特殊规则', r'显示规则', r'样式规则']:
            cf_section_match = re.search(
                r'#+\s*(?:[一二三四五六七八九十\d]+[、.．]\s*)?' + section_keyword + r'\s*\n(.*?)(?=\n#+\s*|\Z)',
                content,
                re.DOTALL | re.IGNORECASE
            )
            if cf_section_match:
                cf_content = cf_section_match.group(1)
                rule_lines = re.findall(r'^[\s]*[-•\d.]\s*(.+?)$', cf_content, re.MULTILINE)
                for line in rule_lines:
                    line = line.strip()
                    if line and not line.startswith('---'):
                        # 只提取包含样式相关关键词的行
                        if self._is_conditional_format_line(line):
                            parsed = self._parse_conditional_format_rule(line)
                            if parsed:
                                rules.append(parsed)

        # 方式2：全文扫描含有条件格式关键词的规则行
        cf_patterns = [
            r'(.{0,50}(?:标红|标黄|标绿|标蓝|变红|变黄|高亮|底色|背景色|字体颜色|加粗显示|颜色标记|条件格式|标注[红黄绿蓝橙]色|[红黄绿蓝橙]色(?:标[注记]|显示|提[醒示])|填充[红黄绿蓝橙]色).{0,80})',
        ]
        for pattern in cf_patterns:
            matches = re.findall(pattern, content)
            for match in matches:
                match = re.sub(r'^[\s\-•]+', '', match).strip()
                if not match:
                    continue
                # 避免重复（已在章节中提取的）
                if any(match in r['original_text'] or r['original_text'] in match for r in rules):
                    continue
                # 排除非规则性内容（如代码注释、日志等）
                if match.startswith('#') or match.startswith('//') or 'import' in match:
                    continue
                parsed = self._parse_conditional_format_rule(match)
                if parsed:
                    rules.append(parsed)

        return rules

    def _is_conditional_format_line(self, text: str) -> bool:
        """判断一行文本是否包含条件格式相关关键词"""
        keywords = [
            '标红', '标黄', '标绿', '标蓝', '变红', '变黄',
            '高亮', '底色', '背景色', '字体颜色', '加粗显示', '加粗',
            '颜色标记', '条件格式', '红色', '黄色', '绿色', '蓝色', '橙色',
            '标注', '着色', '变色',
        ]
        return any(kw in text for kw in keywords)

    def _parse_conditional_format_rule(self, rule_text: str) -> Dict[str, Any]:
        """解析单条条件格式规则"""
        rule = {
            "original_text": rule_text,
            "field": None,
            "condition": None,
            "style": None
        }

        # 提取样式类型
        style_map = {
            "标红": {"type": "fill", "color": "FF0000"},
            "变红": {"type": "fill", "color": "FF0000"},
            "红色": {"type": "fill", "color": "FF0000"},
            "标黄": {"type": "fill", "color": "FFFF00"},
            "变黄": {"type": "fill", "color": "FFFF00"},
            "黄色": {"type": "fill", "color": "FFFF00"},
            "标绿": {"type": "fill", "color": "00FF00"},
            "绿色": {"type": "fill", "color": "00FF00"},
            "标蓝": {"type": "fill", "color": "0000FF"},
            "蓝色": {"type": "fill", "color": "0000FF"},
            "橙色": {"type": "fill", "color": "FFA500"},
            "高亮": {"type": "fill", "color": "FFFF00"},
            "加粗显示": {"type": "font", "bold": True},
            "加粗": {"type": "font", "bold": True},
        }
        for keyword, style in style_map.items():
            if keyword in rule_text:
                rule["style"] = style
                break

        # 提取条件（大于、小于、等于、不等于等）
        condition_match = re.search(
            r'(大于|小于|等于|不等于|超过|低于|>=|<=|>|<|=)\s*(\d+\.?\d*)',
            rule_text
        )
        if condition_match:
            op_text = condition_match.group(1)
            value = condition_match.group(2)
            op_map = {
                "大于": "greaterThan", "超过": "greaterThan", ">": "greaterThan",
                "小于": "lessThan", "低于": "lessThan", "<": "lessThan",
                "等于": "equal", "=": "equal",
                "不等于": "notEqual",
                ">=": "greaterThanOrEqual",
                "<=": "lessThanOrEqual",
            }
            rule["condition"] = {
                "operator": op_map.get(op_text, "greaterThan"),
                "value": value
            }

        # 提取字段名
        field_match = re.search(r'[【""]?([^】"""\s]{2,8})[】""]?\s*(?:大于|小于|等于|不等于|超过|低于|>=|<=|>|<)', rule_text)
        if field_match:
            rule["field"] = field_match.group(1)

        return rule

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

        # 格式化条件格式规则
        if rules.get("conditional_format_rules"):
            sections.append("## 条件格式规则\n")
            sections.append("在填充公式后，需要对以下情况应用条件格式（使用openpyxl的conditional_formatting）：\n")
            for i, rule in enumerate(rules["conditional_format_rules"], 1):
                sections.append(f"{i}. {rule['original_text']}\n")
            sections.append("\n**重要**: 使用`from openpyxl.formatting.rule import CellIsRule`实现条件格式，放在公式填充之后。\n\n")

        # 格式化数值精度规则
        if rules.get("precision_rules"):
            sections.append("## 数值精度规则\n")
            sections.append("在生成公式时，必须按以下精度要求使用ROUND函数：\n")
            for i, rule in enumerate(rules["precision_rules"], 1):
                sections.append(f"{i}. {rule['original_text']}\n")
            sections.append("\n")

        return "".join(sections)
