"""
AI驱动的校验规则生成器

在训练时调用AI分析规则文件，自动提取数据校验规则。
生成的校验规则将保存到脚本信息中，供计算时使用。
"""

import json
import logging
from typing import Dict, Any, List, Optional

logger = logging.getLogger(__name__)


class ValidationRuleGenerator:
    """使用AI生成数据校验规则"""

    def __init__(self, ai_provider):
        """初始化

        Args:
            ai_provider: AI提供者实例
        """
        self.ai_provider = ai_provider

    def generate_validation_rules(
        self,
        rules_content: str,
        source_structure: Dict[str, Any]
    ) -> Dict[str, Any]:
        """使用AI分析规则内容，生成校验规则

        Args:
            rules_content: 规则文件内容
            source_structure: 源数据结构

        Returns:
            校验规则字典，格式：
            {
                "value_constraints": [
                    {"column": "加班小时", "operator": ">", "value": 20, "message": "加班小时数不能大于20"}
                ],
                "required_columns": {"员工信息": ["工号", "姓名"]},
                "custom_rules": ["其他自定义规则描述"]
            }
        """
        if not rules_content:
            return {"value_constraints": [], "required_columns": {}, "custom_rules": []}

        # 构建源数据字段信息
        source_columns = self._extract_source_columns(source_structure)

        prompt = f"""分析以下规则文件内容，提取其中的数据校验规则。

## 源数据可用字段
{json.dumps(source_columns, ensure_ascii=False, indent=2)}

## 规则文件内容
{rules_content}

## 任务
请从规则文件中提取所有数据校验规则，包括：
1. 数值范围约束（如：加班小时不能大于20）
2. 必需字段约束
3. 数据格式约束
4. 其他校验规则

## 输出格式
请输出JSON格式，结构如下：
```json
{{
    "value_constraints": [
        {{
            "column": "列名（必须与源数据字段名完全匹配）",
            "operator": "操作符（>、>=、<、<=、==、!=）",
            "value": 数值,
            "message": "错误提示信息"
        }}
    ],
    "required_columns": {{
        "文件名关键词": ["必需列1", "必需列2"]
    }},
    "custom_rules": [
        "其他无法用简单规则表达的校验描述"
    ]
}}
```

注意：
1. column字段名必须与源数据中的字段名完全一致（如"加班小时"而非"加班小时数"）
2. 只提取规则文件中明确要求的校验规则
3. 如果没有找到校验规则，返回空数组
4. 只输出JSON，不要有其他内容"""

        try:
            # 检查 provider 是否支持流式调用
            messages = [
                {"role": "user", "content": prompt}
            ]

            # 尝试使用流式调用（DeepSeek 要求）
            response = ""
            try:
                # 检查是否有流式方法
                if hasattr(self.ai_provider, '_openai_chat_stream'):
                    for chunk, finish_reason in self.ai_provider._openai_chat_stream(messages):
                        if chunk:
                            response += chunk
                elif hasattr(self.ai_provider, '_claude_chat_stream'):
                    for chunk, finish_reason in self.ai_provider._claude_chat_stream("", messages):
                        if chunk:
                            response += chunk
                else:
                    # 回退到非流式
                    response = self.ai_provider.chat(messages)
            except Exception as stream_error:
                logger.warning(f"流式调用失败，回退到非流式: {stream_error}")
                response = self.ai_provider.chat(messages)

            rules = self._parse_ai_response(response)
            logger.info(f"AI生成校验规则: {len(rules.get('value_constraints', []))} 个数值约束")
            return rules
        except Exception as e:
            logger.error(f"AI生成校验规则失败: {e}")
            return {"value_constraints": [], "required_columns": {}, "custom_rules": []}

    def _extract_source_columns(self, source_structure: Dict[str, Any]) -> Dict[str, List[str]]:
        """从源数据结构中提取所有字段名

        Args:
            source_structure: 源数据结构

        Returns:
            {文件名: [字段列表]}
        """
        columns = {}
        files = source_structure.get("files", {})
        for file_name, file_info in files.items():
            sheets = file_info.get("sheets", {})
            for sheet_name, sheet_info in sheets.items():
                headers = sheet_info.get("headers", {})
                if isinstance(headers, dict):
                    columns[f"{file_name}/{sheet_name}"] = list(headers.keys())
                elif isinstance(headers, list):
                    columns[f"{file_name}/{sheet_name}"] = headers
        return columns

    def _parse_ai_response(self, response: str) -> Dict[str, Any]:
        """解析AI响应，提取JSON

        Args:
            response: AI响应文本

        Returns:
            解析后的规则字典
        """
        import re

        # 尝试提取JSON代码块
        json_pattern = r'```json\s*(.*?)```'
        matches = re.findall(json_pattern, response, re.DOTALL)

        if matches:
            try:
                return json.loads(matches[0])
            except json.JSONDecodeError:
                pass

        # 尝试直接解析
        try:
            # 找到第一个{和最后一个}
            start = response.find('{')
            end = response.rfind('}')
            if start >= 0 and end > start:
                return json.loads(response[start:end + 1])
        except json.JSONDecodeError:
            pass

        logger.warning("无法解析AI响应为JSON")
        return {"value_constraints": [], "required_columns": {}, "custom_rules": []}


def generate_validation_rules_with_ai(
    ai_provider,
    rules_content: str,
    source_structure: Dict[str, Any]
) -> Dict[str, Any]:
    """便捷函数：使用AI生成校验规则

    Args:
        ai_provider: AI提供者实例
        rules_content: 规则文件内容
        source_structure: 源数据结构

    Returns:
        校验规则字典
    """
    generator = ValidationRuleGenerator(ai_provider)
    return generator.generate_validation_rules(rules_content, source_structure)
