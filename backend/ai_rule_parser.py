"""
AI增强规则解析器 - 使用第三方AI API进行智能规则解析
"""

import os
import json
import re
from typing import Dict, List, Any, Optional
from pathlib import Path

# 导入现有的AI引擎和规则解析器
from backend.ai_engine.ai_provider import AIProviderFactory
from backend.rule_parser import RuleParser, RuleSet, FileRule, SheetRule, ColumnRule


class AIRuleParser:
    """AI增强规则解析器"""

    def __init__(self, ai_provider_type: str = None, ai_config: Optional[Dict[str, Any]] = None):
        """
        初始化AI增强规则解析器

        Args:
            ai_provider_type: AI提供者类型，如果为None则从配置读取
            ai_config: AI配置，如果为None则使用默认配置
        """
        if ai_provider_type is None and ai_config is None:
            # 从配置自动创建AI提供者
            self.ai_provider = AIProviderFactory.create_with_fallback()
        else:
            # 使用指定的提供者和配置
            self.ai_config = ai_config or {}
            self.ai_provider = AIProviderFactory.create_provider(ai_provider_type, self.ai_config)

        self.rule_parser = RuleParser()

    def parse_with_ai(self, file_path: str, use_ai_for_unclear: bool = True) -> RuleSet:
        """
        使用AI增强解析规则文件

        Args:
            file_path: 规则文件路径
            use_ai_for_unclear: 是否对不明确的规则使用AI解析

        Returns:
            解析后的规则集
        """
        # 首先使用基础解析器
        try:
            rule_set = self.rule_parser.parse(file_path)
        except Exception as e:
            print(f"基础解析失败: {e}")
            # 如果基础解析失败，完全使用AI解析
            return self._parse_fully_with_ai(file_path)

        # 检查规则完整性
        if use_ai_for_unclear and self._needs_ai_enhancement(rule_set):
            # 使用AI增强解析
            enhanced_rule_set = self._enhance_rules_with_ai(file_path, rule_set)
            return enhanced_rule_set

        return rule_set

    def _needs_ai_enhancement(self, rule_set: RuleSet) -> bool:
        """检查是否需要AI增强"""
        # 检查预期文件是否有足够的列规则
        if not rule_set.expected_file.sheets:
            return True

        for sheet in rule_set.expected_file.sheets:
            if not sheet.columns:
                return True
            for column in sheet.columns:
                # 检查列规则是否完整
                if not column.data_source or column.data_source.strip() == '':
                    return True
                # 检查是否有计算规则但数据来源不明确
                if column.calculation_rule and not self._is_valid_data_source(column.data_source):
                    return True

        # 检查映射规则是否完整
        if rule_set.mapping_rules and len(rule_set.mapping_rules) == 0:
            return True

        return False

    def _is_valid_data_source(self, data_source: str) -> bool:
        """检查数据来源是否有效"""
        # 有效的格式: file.xlsx!Sheet1!A列 或 file.xlsx!A1:B10
        patterns = [
            r'.+\.(xlsx?|xls|csv)!.+',
            r'.+\.(xlsx?|xls|csv)!.+!.+',
            r'[A-Z]+[0-9]+:[A-Z]+[0-9]+',
            r'[A-Z]+列'
        ]

        for pattern in patterns:
            if re.match(pattern, data_source):
                return True

        return False

    def _parse_fully_with_ai(self, file_path: str) -> RuleSet:
        """完全使用AI解析规则文件"""
        # 提取文件内容
        file_content = self._extract_file_content(file_path)

        # 生成AI提示词
        prompt = self._generate_ai_prompt(file_content, file_path)

        # 调用AI
        ai_response = self.ai_provider.generate_code(prompt)

        # 解析AI响应
        return self._parse_ai_response(ai_response, file_path)

    def _enhance_rules_with_ai(self, file_path: str, base_rule_set: RuleSet) -> RuleSet:
        """使用AI增强现有规则"""
        # 提取文件内容
        file_content = self._extract_file_content(file_path)

        # 生成增强提示词
        prompt = self._generate_enhancement_prompt(file_content, base_rule_set, file_path)

        # 调用AI
        ai_response = self.ai_provider.generate_code(prompt)

        # 解析AI增强响应
        enhanced_rule_set = self._parse_enhancement_response(ai_response, base_rule_set)

        return enhanced_rule_set

    def _extract_file_content(self, file_path: str) -> str:
        """提取文件内容"""
        file_ext = Path(file_path).suffix.lower()

        if file_ext == '.pdf':
            # 使用PDF解析器提取文本
            from backend.rule_parser import PDFRuleParser
            parser = PDFRuleParser()
            # 这里简化处理，实际应该调用PDF解析器的方法
            return "PDF内容提取"
        elif file_ext in ['.docx', '.doc']:
            # 使用Word解析器提取文本
            from backend.rule_parser import WordRuleParser
            parser = WordRuleParser()
            # 这里简化处理，实际应该调用Word解析器的方法
            return "Word内容提取"
        elif file_ext in ['.xlsx', '.xls']:
            # 使用Excel解析器提取数据
            from backend.rule_parser import ExcelRuleParser
            parser = ExcelRuleParser()
            # 这里简化处理，实际应该调用Excel解析器的方法
            return "Excel内容提取"
        else:
            # 文本文件
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            except:
                return ""

    def _generate_ai_prompt(self, file_content: str, file_path: str) -> str:
        """生成AI解析提示词"""
        file_name = Path(file_path).name

        prompt = f"""你是一个专业的规则文件解析专家。请分析以下规则文件内容，并提取出完整的数据处理规则。

## 规则文件: {file_name}
{file_content}

## 需要提取的信息:

1. **预期输出文件**
   - 文件名
   - 工作表结构
   - 每列的规则

2. **源文件规则**
   - 每个源文件的文件名
   - 数据来源位置
   - 数据格式

3. **数据映射规则**
   - 源数据到目标数据的映射关系
   - 格式: 源位置 -> 目标位置

4. **计算规则**
   - 需要计算的列
   - 计算公式
   - 格式: 列名 = 公式

## 输出格式要求:
请以JSON格式输出，结构如下:
{{
  "expected_file": {{
    "file_name": "输出文件名.xlsx",
    "sheets": [
      {{
        "sheet_name": "工作表名",
        "columns": [
          {{
            "column_name": "列名",
            "data_source": "数据来源",
            "calculation_rule": "计算规则(可选)",
            "validation_rule": "验证规则(可选)",
            "description": "描述(可选)"
          }}
        ]
      }}
    ]
  }},
  "source_files": [
    {{
      "file_name": "源文件名.xlsx",
      "sheets": [
        {{
          "sheet_name": "工作表名",
          "columns": [
            {{
              "column_name": "列名",
              "data_source": "数据位置",
              "description": "描述(可选)"
            }}
          ]
        }}
      ]
    }}
  ],
  "mapping_rules": {{
    "源位置1": "目标位置1",
    "源位置2": "目标位置2"
  }},
  "calculation_rules": {{
    "目标列1": "计算公式1",
    "目标列2": "计算公式2"
  }}
}}

## 注意事项:
1. 数据来源格式: 文件名!工作表名!列名 或 文件名!单元格范围
2. 计算规则使用Excel公式语法
3. 确保所有规则都是明确和可执行的

请只输出JSON格式的结果，不要有其他内容。"""

        return prompt

    def _generate_enhancement_prompt(self, file_content: str, base_rule_set: RuleSet, file_path: str) -> str:
        """生成AI增强提示词"""
        file_name = Path(file_path).name

        # 将基础规则集转换为JSON
        base_rules_json = self._rule_set_to_dict(base_rule_set)

        prompt = f"""你是一个专业的规则文件解析专家。请基于已有的规则分析和规则文件内容，补充和完善规则。

## 规则文件: {file_name}
{file_content}

## 已有规则分析:
{json.dumps(base_rules_json, ensure_ascii=False, indent=2)}

## 需要补充和完善的内容:

1. **缺失的数据来源**
   - 检查每个列是否有明确的数据来源
   - 补充缺失的数据来源信息

2. **不明确的映射关系**
   - 完善源数据到目标数据的映射
   - 确保映射关系是明确和可执行的

3. **计算规则细化**
   - 完善计算公式
   - 确保公式语法正确

4. **验证规则**
   - 添加必要的数据验证规则

## 输出格式要求:
请以JSON格式输出完整的、增强后的规则，结构如下:
{{
  "expected_file": {{
    "file_name": "输出文件名.xlsx",
    "sheets": [
      {{
        "sheet_name": "工作表名",
        "columns": [
          {{
            "column_name": "列名",
            "data_source": "数据来源",
            "calculation_rule": "计算规则(可选)",
            "validation_rule": "验证规则(可选)",
            "description": "描述(可选)"
          }}
        ]
      }}
    ]
  }},
  "source_files": [
    {{
      "file_name": "源文件名.xlsx",
      "sheets": [
        {{
          "sheet_name": "工作表名",
          "columns": [
            {{
              "column_name": "列名",
              "data_source": "数据位置",
              "description": "描述(可选)"
            }}
          ]
        }}
      ]
    }}
  ],
  "mapping_rules": {{
    "源位置1": "目标位置1",
    "源位置2": "目标位置2"
  }},
  "calculation_rules": {{
    "目标列1": "计算公式1",
    "目标列2": "计算公式2"
  }}
}}

## 注意事项:
1. 保持已有规则的结构
2. 只补充和完善缺失或不明确的部分
3. 确保所有规则都是明确和可执行的
4. 数据来源格式: 文件名!工作表名!列名 或 文件名!单元格范围

请只输出JSON格式的结果，不要有其他内容。"""

        return prompt

    def _parse_ai_response(self, ai_response: str, file_path: str) -> RuleSet:
        """解析AI响应"""
        try:
            # 提取JSON部分
            json_match = re.search(r'\{.*\}', ai_response, re.DOTALL)
            if json_match:
                json_str = json_match.group(0)
                rules_dict = json.loads(json_str)
            else:
                # 尝试直接解析
                rules_dict = json.loads(ai_response)

            # 转换为RuleSet对象
            return self._dict_to_rule_set(rules_dict, file_path)

        except json.JSONDecodeError as e:
            print(f"AI响应JSON解析失败: {e}")
            print(f"AI响应内容: {ai_response[:500]}...")

            # 尝试从文本中提取规则
            return self._extract_rules_from_text(ai_response, file_path)

    def _parse_enhancement_response(self, ai_response: str, base_rule_set: RuleSet) -> RuleSet:
        """解析AI增强响应"""
        try:
            # 提取JSON部分
            json_match = re.search(r'\{.*\}', ai_response, re.DOTALL)
            if json_match:
                json_str = json_match.group(0)
                enhanced_dict = json.loads(json_str)
            else:
                enhanced_dict = json.loads(ai_response)

            # 合并基础规则和增强规则
            merged_dict = self._merge_rules(base_rule_set, enhanced_dict)

            # 转换为RuleSet对象
            return self._dict_to_rule_set(merged_dict, "")

        except json.JSONDecodeError as e:
            print(f"AI增强响应JSON解析失败: {e}")
            return base_rule_set

    def _rule_set_to_dict(self, rule_set: RuleSet) -> Dict[str, Any]:
        """将RuleSet转换为字典"""
        return {
            "expected_file": {
                "file_name": rule_set.expected_file.file_name,
                "sheets": [
                    {
                        "sheet_name": sheet.sheet_name,
                        "columns": [
                            {
                                "column_name": column.column_name,
                                "data_source": column.data_source,
                                "calculation_rule": column.calculation_rule,
                                "validation_rule": column.validation_rule,
                                "description": column.description
                            }
                            for column in sheet.columns
                        ]
                    }
                    for sheet in rule_set.expected_file.sheets
                ]
            },
            "source_files": [
                {
                    "file_name": file.file_name,
                    "sheets": [
                        {
                            "sheet_name": sheet.sheet_name,
                            "columns": [
                                {
                                    "column_name": column.column_name,
                                    "data_source": column.data_source,
                                    "description": column.description
                                }
                                for column in sheet.columns
                            ]
                        }
                        for sheet in file.sheets
                    ]
                }
                for file in rule_set.source_files
            ],
            "mapping_rules": rule_set.mapping_rules,
            "calculation_rules": rule_set.calculation_rules
        }

    def _dict_to_rule_set(self, rules_dict: Dict[str, Any], file_path: str) -> RuleSet:
        """将字典转换为RuleSet"""
        # 解析预期文件
        expected_file_dict = rules_dict.get("expected_file", {})
        expected_file = FileRule(
            file_name=expected_file_dict.get("file_name", f"{Path(file_path).stem}_output.xlsx"),
            sheets=[
                SheetRule(
                    sheet_name=sheet_dict["sheet_name"],
                    columns=[
                        ColumnRule(
                            column_name=col_dict["column_name"],
                            data_source=col_dict.get("data_source", ""),
                            calculation_rule=col_dict.get("calculation_rule"),
                            validation_rule=col_dict.get("validation_rule"),
                            description=col_dict.get("description")
                        )
                        for col_dict in sheet_dict.get("columns", [])
                    ]
                )
                for sheet_dict in expected_file_dict.get("sheets", [])
            ]
        )

        # 解析源文件
        source_files = []
        for source_file_dict in rules_dict.get("source_files", []):
            source_file = FileRule(
                file_name=source_file_dict["file_name"],
                sheets=[
                    SheetRule(
                        sheet_name=sheet_dict["sheet_name"],
                        columns=[
                            ColumnRule(
                                column_name=col_dict["column_name"],
                                data_source=col_dict.get("data_source", ""),
                                description=col_dict.get("description")
                            )
                            for col_dict in sheet_dict.get("columns", [])
                        ]
                    )
                    for sheet_dict in source_file_dict.get("sheets", [])
                ]
            )
            source_files.append(source_file)

        # 获取映射规则和计算规则
        mapping_rules = rules_dict.get("mapping_rules", {})
        calculation_rules = rules_dict.get("calculation_rules", {})

        return RuleSet(
            expected_file=expected_file,
            source_files=source_files,
            mapping_rules=mapping_rules,
            calculation_rules=calculation_rules
        )

    def _merge_rules(self, base_rule_set: RuleSet, enhanced_dict: Dict[str, Any]) -> Dict[str, Any]:
        """合并基础规则和增强规则"""
        base_dict = self._rule_set_to_dict(base_rule_set)

        # 深度合并字典
        def deep_merge(base: Any, enhancement: Any) -> Any:
            if isinstance(base, dict) and isinstance(enhancement, dict):
                result = base.copy()
                for key, value in enhancement.items():
                    if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                        result[key] = deep_merge(result[key], value)
                    elif key in result and isinstance(result[key], list) and isinstance(value, list):
                        # 合并列表，优先使用增强版本
                        result[key] = value
                    else:
                        result[key] = value
                return result
            elif isinstance(base, list) and isinstance(enhancement, list):
                # 对于列表，使用增强版本
                return enhancement
            else:
                # 其他情况使用增强版本
                return enhancement

        return deep_merge(base_dict, enhanced_dict)

    def _extract_rules_from_text(self, text: str, file_path: str) -> RuleSet:
        """从文本中提取规则（备用方法）"""
        # 这里可以使用基础解析器的文本解析逻辑
        # 简化实现，实际应该调用基础解析器的方法
        return RuleSet(
            expected_file=FileRule(
                file_name=f"{Path(file_path).stem}_output.xlsx",
                sheets=[]
            ),
            source_files=[],
            mapping_rules={},
            calculation_rules={}
        )


# 使用示例
if __name__ == "__main__":
    # 配置AI（需要设置API密钥）
    ai_config = {
        "api_key": os.getenv("OPENAI_API_KEY"),
        "model": "gpt-4"
    }

    # 创建AI增强解析器
    parser = AIRuleParser(ai_provider_type="openai", ai_config=ai_config)

    # 解析规则文件
    try:
        rule_set = parser.parse_with_ai("rules.pdf")
        print(f"AI增强解析成功:")
        print(f"预期输出文件: {rule_set.expected_file.file_name}")
        print(f"源文件数量: {len(rule_set.source_files)}")
        print(f"映射规则数量: {len(rule_set.mapping_rules)}")
        print(f"计算规则数量: {len(rule_set.calculation_rules)}")
    except Exception as e:
        print(f"AI增强解析失败: {e}")