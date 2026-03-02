"""
AI规则生成器 - 使用AI分析需求文档并生成结构化规则
"""

import json
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path

from .ai_provider import BaseAIProvider
from .ai_provider import AIProviderFactory


class AIRuleGenerator:
    """AI规则生成器"""

    def __init__(self, ai_provider: BaseAIProvider = None):
        """初始化AI规则生成器"""
        if ai_provider is None:
            self.ai_provider = AIProviderFactory.create_with_fallback()
        else:
            self.ai_provider = ai_provider

        self.logger = logging.getLogger(__name__)

    def generate_rules_from_document(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        从需求文档生成结构化规则

        Args:
            document_content: 需求文档内容（PDF/Word提取的文本）
            source_structure: 源文件结构
            expected_structure: 预期输出文件结构
            manual_headers: 手动表头规则

        Returns:
            结构化规则字典
        """
        try:
            self.logger.info("开始使用AI生成规则...")

            # 检查文档长度，决定使用哪种策略
            doc_length = len(document_content)
            self.logger.info(f"文档长度: {doc_length} 字符")

            if doc_length > 50000:  # 超长文档
                self.logger.info("检测到超长文档，使用分步处理策略...")
                return self._generate_rules_step_by_step(
                    document_content, source_structure, expected_structure, manual_headers
                )
            elif doc_length > 20000:  # 长文档
                self.logger.info("检测到长文档，使用压缩摘要策略...")
                return self._generate_rules_with_compression(
                    document_content, source_structure, expected_structure, manual_headers
                )
            else:  # 短文档
                self.logger.info("文档长度适中，使用标准策略...")
                return self._generate_rules_standard(
                    document_content, source_structure, expected_structure, manual_headers
                )

        except Exception as e:
            self.logger.error(f"AI规则生成失败: {e}")
            # 返回默认规则
            return self._get_default_rules()

    def _generate_rules_standard(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """标准规则生成策略"""
        # 准备提示词
        prompt = self._create_rule_generation_prompt(
            document_content, source_structure, expected_structure, manual_headers
        )

        # 调用AI生成规则
        self.logger.info("调用AI分析文档并生成规则...")
        response = self.ai_provider.generate_completion(prompt)

        # 解析AI响应
        rules = self._parse_ai_response(response)

        self.logger.info(f"AI规则生成完成，生成 {len(rules.get('mapping_rules', {}))} 个映射规则")
        return rules

    def _generate_rules_with_compression(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """使用压缩摘要的规则生成策略"""
        # 创建高度压缩的提示词
        prompt = self._create_compressed_prompt(
            document_content, source_structure, expected_structure, manual_headers
        )

        # 调用AI生成规则
        self.logger.info("调用AI分析压缩摘要并生成规则...")
        response = self.ai_provider.generate_completion(prompt)

        # 解析AI响应
        rules = self._parse_ai_response(response)

        self.logger.info(f"AI规则生成完成（压缩策略），生成 {len(rules.get('mapping_rules', {}))} 个映射规则")
        return rules

    def _generate_rules_step_by_step(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """分步处理策略：先分析结构，再生成规则"""
        self.logger.info("步骤1: 分析文档结构...")

        # 第一步：分析文档结构
        structure_prompt = self._create_structure_analysis_prompt(document_content)
        structure_response = self.ai_provider.generate_completion(structure_prompt)

        # 解析结构分析结果
        doc_structure = self._parse_structure_analysis(structure_response)
        self.logger.info(f"文档结构分析完成: {doc_structure.get('summary', '无总结')}")

        self.logger.info("步骤2: 基于分析结果生成规则...")

        # 第二步：基于分析结果生成规则
        rules_prompt = self._create_rules_from_structure_prompt(
            doc_structure, source_structure, expected_structure, manual_headers
        )
        rules_response = self.ai_provider.generate_completion(rules_prompt)

        # 解析规则
        rules = self._parse_ai_response(rules_response)

        self.logger.info(f"分步规则生成完成，生成 {len(rules.get('mapping_rules', {}))} 个映射规则")
        return rules

    def _create_compressed_prompt(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """创建高度压缩的提示词"""
        # 提取极简摘要
        doc_summary = self._extract_minimal_summary(document_content, max_length=800)

        # 极简文件结构
        simple_source = self._simplify_structure(source_structure, "源文件")
        simple_expected = self._simplify_structure(expected_structure, "预期文件")

        prompt = f"""你是一个专业的数据处理规则分析师。请基于以下高度压缩的信息生成数据处理规则。

## 文档核心内容
{doc_summary}

## 源文件关键信息
{simple_source}

## 预期输出关键信息
{simple_expected}

## 任务要求
基于以上信息，推断并生成：
1. 最可能的数据映射关系
2. 关键的计算规则
3. 必要的验证规则

请以JSON格式输出规则，结构参考之前的格式，但可以更简洁。
直接输出JSON，不要有其他文本。"""

        return prompt

    def _extract_minimal_summary(self, document_content: str, max_length: int = 800) -> str:
        """提取极简摘要"""
        if not document_content:
            return "[无内容]"

        lines = document_content.split('\n')
        if len(lines) == 0:
            return "[无内容]"

        # 只提取最关键的信息
        key_info = []

        # 1. 文档标题（第一行）
        if lines[0].strip():
            title = lines[0].strip()[:100]
            key_info.append(f"标题: {title}")

        # 2. 查找包含关键业务词汇的行
        business_keywords = ['计算', '映射', '规则', '公式', '报表', '薪资', '工资', '考勤', '绩效']
        business_lines = []
        for line in lines:
            line_lower = line.lower()
            for keyword in business_keywords:
                if keyword in line_lower:
                    # 提取简洁版本
                    clean_line = ' '.join(line.split()[:8])  # 只取前8个词
                    if len(clean_line) > 80:
                        clean_line = clean_line[:80] + "..."
                    business_lines.append(clean_line)
                    if len(business_lines) >= 5:  # 最多5行
                        break
            if len(business_lines) >= 5:
                break

        if business_lines:
            key_info.append("关键业务描述:\n" + '\n'.join(business_lines))

        # 3. 文档统计
        total_chars = len(document_content)
        key_info.append(f"文档规模: {total_chars}字符")

        summary = '\n\n'.join(key_info)

        if len(summary) > max_length:
            summary = summary[:max_length] + f"... [已极简压缩]"

        return summary

    def _simplify_structure(self, structure: Dict[str, Any], structure_type: str) -> str:
        """极简文件结构"""
        if not structure:
            return f"{structure_type}: 无结构信息"

        try:
            if structure_type == "源文件":
                files = structure.get('files', {})
                if not files:
                    return "源文件: 无文件信息"

                # 只列出文件名和主要sheet
                file_list = []
                for file_name in list(files.keys())[:2]:  # 最多2个文件
                    file_info = files[file_name]
                    sheets = file_info.get('sheets', {})
                    sheet_names = list(sheets.keys())[:1]  # 每个文件最多1个sheet
                    if sheet_names:
                        file_list.append(f"{file_name} ({sheet_names[0]})")
                    else:
                        file_list.append(file_name)

                files_text = ', '.join(file_list)
                if len(files) > 2:
                    files_text += f" 等{len(files)}个文件"

                return f"源文件: {files_text}"

            elif structure_type == "预期文件":
                sheets = structure.get('sheets', {})
                if not sheets:
                    return "预期文件: 无sheet信息"

                # 只提取主要列
                sheet_info = []
                for sheet_name, sheet_data in sheets.items():
                    headers = list(sheet_data.get('headers', {}).keys())[:4]  # 最多4个列
                    if headers:
                        sheet_info.append(f"{sheet_name}: {', '.join(headers)}")
                    else:
                        sheet_info.append(sheet_name)

                sheets_text = '; '.join(sheet_info)
                return f"预期文件: {sheets_text}"

            else:
                return f"{structure_type}: 简化结构信息"

        except Exception:
            return f"{structure_type}: [结构信息]"

    def _create_structure_analysis_prompt(self, document_content: str) -> str:
        """创建文档结构分析提示词"""
        # 使用中等长度的摘要
        doc_summary = self._extract_document_summary(document_content, max_length=1200)

        prompt = f"""你是一个文档分析师。请分析以下需求文档的结构和主要内容。

## 文档内容摘要
{doc_summary}

## 分析要求
请分析文档的以下方面：
1. **文档类型**：是什么类型的文档？（需求文档、技术规范、操作手册等）
2. **核心业务**：文档涉及的核心业务是什么？
3. **数据处理需求**：文档中描述了哪些数据处理需求？
4. **关键表格**：文档中提到了哪些关键表格或数据结构？
5. **计算规则**：是否有明确的计算公式或规则？
6. **输出要求**：对输出结果有什么要求？

请以JSON格式输出分析结果，结构如下：
```json
{{
  "document_type": "文档类型",
  "core_business": "核心业务描述",
  "data_processing_needs": ["需求1", "需求2", ...],
  "key_tables": ["表1", "表2", ...],
  "calculation_rules": ["规则1", "规则2", ...],
  "output_requirements": ["要求1", "要求2", ...],
  "summary": "分析总结"
}}
```

直接输出JSON，不要有其他文本。"""

        return prompt

    def _parse_structure_analysis(self, response: str) -> Dict[str, Any]:
        """解析结构分析结果"""
        try:
            # 尝试提取JSON
            import json
            lines = response.strip().split('\n')
            json_start = -1
            json_end = -1

            # 查找JSON开始位置
            for i, line in enumerate(lines):
                line = line.strip()
                if line.startswith('{'):
                    json_start = i
                    break

            if json_start == -1:
                # 如果没有找到JSON开始，尝试整个响应
                json_text = response.strip()
            else:
                # 查找JSON结束位置
                brace_count = 0
                for j in range(json_start, len(lines)):
                    line = lines[j]
                    brace_count += line.count('{')
                    brace_count -= line.count('}')

                    if brace_count == 0:
                        json_end = j
                        break

                if json_end == -1:
                    json_end = len(lines) - 1

                json_text = '\n'.join(lines[json_start:json_end + 1])

            structure = json.loads(json_text)
            return structure

        except Exception as e:
            self.logger.warning(f"解析结构分析结果失败: {e}")
            return {
                "document_type": "需求文档",
                "core_business": "未知",
                "data_processing_needs": [],
                "key_tables": [],
                "calculation_rules": [],
                "output_requirements": [],
                "summary": "结构分析失败"
            }

    def _create_rules_from_structure_prompt(
        self,
        doc_structure: Dict[str, Any],
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """基于文档结构创建规则生成提示词"""
        # 压缩文件结构
        compressed_source = self._compress_structure(source_structure, "源文件")
        compressed_expected = self._compress_structure(expected_structure, "预期文件")

        prompt = f"""你是一个专业的数据处理规则生成器。基于以下文档分析结果和文件结构，生成数据处理规则。

## 文档分析结果
- 文档类型: {doc_structure.get('document_type', '未知')}
- 核心业务: {doc_structure.get('core_business', '未知')}
- 数据处理需求: {', '.join(doc_structure.get('data_processing_needs', []))}
- 关键表格: {', '.join(doc_structure.get('key_tables', []))}
- 计算规则: {', '.join(doc_structure.get('calculation_rules', []))}
- 输出要求: {', '.join(doc_structure.get('output_requirements', []))}

## 源文件结构
{compressed_source}

## 预期输出结构
{compressed_expected}

## 任务要求
基于文档分析结果和文件结构，生成合理的数据处理规则，包括：
1. 文件映射关系
2. 列映射规则
3. 计算规则
4. 验证规则

请以JSON格式输出规则，结构参考标准格式。
直接输出JSON，不要有其他文本。"""

        return prompt

    def _get_default_rules(self) -> Dict[str, Any]:
        """获取默认规则（当AI生成失败时使用）"""
        return {
            "file_mapping": {
                "source_files": [],
                "expected_file": {
                    "file_name": "expected_output.xlsx",
                    "description": "默认预期输出文件",
                    "sheets": []
                }
            },
            "column_mapping": {},
            "calculation_rules": {},
            "validation_rules": {},
            "processing_steps": ["读取源文件", "处理数据", "生成输出"],
            "summary": "AI规则生成失败，使用默认规则"
        }

    def _create_rule_generation_prompt(
        self,
        document_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """创建规则生成提示词"""
        # 1. 提取文档关键信息（避免token超限）
        doc_summary = self._extract_document_summary(document_content)

        # 2. 压缩文件结构信息
        compressed_source = self._compress_structure(source_structure, "源文件")
        compressed_expected = self._compress_structure(expected_structure, "预期输出文件")

        prompt = f"""你是一个专业的数据处理规则分析师。请分析以下需求文档摘要、源文件结构和预期输出结构，生成完整的数据处理规则。

## 需求文档关键信息
{doc_summary}

## 源文件结构（压缩版）
{compressed_source}

## 预期输出结构（压缩版）
{compressed_expected}"""

        if manual_headers:
            prompt += f"""

## 手动表头规则
```json
{json.dumps(manual_headers, ensure_ascii=False, indent=2)}
```
"""

        prompt += """

## 任务要求
请分析以上信息，生成完整的数据处理规则，包括：

1. **文件映射规则**：源文件与预期输出文件的对应关系
2. **列映射规则**：源文件列与预期输出列的映射关系
3. **计算规则**：需要进行的计算和转换
4. **验证规则**：数据验证要求

## 输出格式要求
请以JSON格式输出，包含以下结构：
```json
{{
  "file_mapping": {{
    "source_files": [
      {{
        "file_name": "源文件名.xlsx",
        "description": "文件描述",
        "sheets": [
          {{
            "sheet_name": "Sheet1",
            "columns": ["列1", "列2", ...]
          }}
        ]
      }}
    ],
    "expected_file": {{
      "file_name": "预期输出文件名.xlsx",
      "description": "文件描述",
      "sheets": [
        {{
          "sheet_name": "Sheet1",
          "columns": ["输出列1", "输出列2", ...]
        }}
      ]
    }}
  }},
  "column_mapping": {{
    "源列名1": "目标列名1",
    "源列名2": "目标列名2",
    ...
  }},
  "calculation_rules": {{
    "目标列名": "计算表达式或说明",
    ...
  }},
  "validation_rules": {{
    "列名": "验证条件",
    ...
  }},
  "processing_steps": [
    "步骤1描述",
    "步骤2描述",
    ...
  ],
  "summary": "规则总结说明"
}}
```

## 注意事项
1. 仔细分析需求文档中的业务逻辑
2. 根据源文件和预期文件的结构推断映射关系
3. 如果需求文档中提到计算公式，请提取出来
4. 如果需求文档中提到数据验证要求，请提取出来
5. 确保生成的规则完整且可执行

请直接输出JSON，不要有其他文本。"""

        return prompt

    def _parse_ai_response(self, response: str) -> Dict[str, Any]:
        """解析AI响应，提取JSON格式的规则"""
        try:
            # 尝试从响应中提取JSON
            lines = response.strip().split('\n')
            json_start = -1
            json_end = -1

            for i, line in enumerate(lines):
                if line.strip().startswith('{'):
                    json_start = i
                    break

            if json_start == -1:
                # 如果没有找到JSON开始，尝试整个响应
                json_text = response.strip()
            else:
                # 查找JSON结束
                brace_count = 0
                for j in range(json_start, len(lines)):
                    line = lines[j]
                    brace_count += line.count('{')
                    brace_count -= line.count('}')

                    if brace_count == 0:
                        json_end = j
                        break

                if json_end == -1:
                    json_end = len(lines) - 1

                json_text = '\n'.join(lines[json_start:json_end + 1])

            # 解析JSON
            rules = json.loads(json_text)
            return rules

        except json.JSONDecodeError as e:
            self.logger.error(f"解析AI响应JSON失败: {e}")
            self.logger.error(f"响应内容: {response[:500]}...")

            # 返回默认规则结构
            return {
                "file_mapping": {
                    "source_files": [],
                    "expected_file": {
                        "file_name": "expected_output.xlsx",
                        "description": "AI生成的预期输出文件",
                        "sheets": []
                    }
                },
                "column_mapping": {},
                "calculation_rules": {},
                "validation_rules": {},
                "processing_steps": [],
                "summary": "AI规则解析失败，使用默认规则"
            }

    def convert_to_rule_set(self, ai_rules: Dict[str, Any]) -> Dict[str, Any]:
        """
        将AI生成的规则转换为系统可用的规则集格式

        Args:
            ai_rules: AI生成的规则字典

        Returns:
            系统规则集格式
        """
        # 这里需要根据系统的RuleSet结构进行转换
        # 由于RuleSet是dataclass，我们返回字典格式，由调用者处理

        rule_set = {
            "expected_file": {
                "file_name": ai_rules.get("file_mapping", {}).get("expected_file", {}).get("file_name", "expected_output.xlsx"),
                "sheets": ai_rules.get("file_mapping", {}).get("expected_file", {}).get("sheets", [])
            },
            "source_files": ai_rules.get("file_mapping", {}).get("source_files", []),
            "mapping_rules": ai_rules.get("column_mapping", {}),
            "calculation_rules": ai_rules.get("calculation_rules", {}),
            "validation_rules": ai_rules.get("validation_rules", {}),
            "processing_steps": ai_rules.get("processing_steps", []),
            "summary": ai_rules.get("summary", "AI生成的规则")
        }

        return rule_set

    def _extract_document_summary(self, document_content: str, max_length: int = 1000) -> str:
        """提取文档关键信息摘要（更激进的压缩）"""
        if not document_content:
            return "[文档内容为空]"

        # 如果文档很短，直接返回
        if len(document_content) <= 800:
            return document_content

        lines = document_content.split('\n')
        if len(lines) == 0:
            return "[文档无内容]"

        # 更激进的摘要策略
        summary_parts = []

        # 1. 提取文档标题（前3行）
        title_lines = []
        for i in range(min(3, len(lines))):
            line = lines[i].strip()
            if line and len(line) < 100:  # 避免过长的行作为标题
                title_lines.append(line)

        if title_lines:
            title = ' | '.join(title_lines[:2])  # 最多2行标题
            if len(title) > 150:
                title = title[:150] + "..."
            summary_parts.append(f"文档标题: {title}")

        # 2. 提取最关键的业务逻辑部分
        critical_keywords = [
            '计算规则', '计算公式', '数据映射', '映射关系',
            '输出要求', '预期结果', '报表项', '列对应',
            '薪资计算', '工资计算', '考勤计算', '绩效计算'
        ]

        critical_sections = []
        for i, line in enumerate(lines):
            line_lower = line.lower()
            for keyword in critical_keywords:
                if keyword.lower() in line_lower:
                    # 提取关键段落（前后各2行）
                    start = max(0, i - 2)
                    end = min(len(lines), i + 3)
                    section = '\n'.join(lines[start:end])
                    critical_sections.append(section)
                    break

        if critical_sections:
            # 只取最重要的3个部分
            critical_text = '\n\n'.join(critical_sections[:3])
            if len(critical_text) > 600:
                critical_text = critical_text[:600] + "..."
            summary_parts.append(f"关键业务逻辑:\n{critical_text}")

        # 3. 提取表格结构信息（不提取具体数据）
        table_headers = []
        for line in lines:
            if '列名' in line or ('需要' in line and '存储' in line) or '报表项' in line:
                # 清理表格行，只保留关键信息
                clean_line = ' '.join(line.split()[:10])  # 只取前10个词
                if len(clean_line) > 100:
                    clean_line = clean_line[:100] + "..."
                table_headers.append(clean_line)

        if table_headers:
            tables_text = '\n'.join(table_headers[:5])  # 最多5个表头
            summary_parts.append(f"表格结构:\n{tables_text}")

        # 4. 提取数据流程描述
        flow_keywords = ['数据流程', '处理流程', '业务流程', '工作流程', '步骤']
        flow_sections = []
        for i, line in enumerate(lines):
            line_lower = line.lower()
            for keyword in flow_keywords:
                if keyword in line_lower:
                    # 提取流程描述（当前行和后4行）
                    start = max(0, i)
                    end = min(len(lines), i + 5)
                    flow = '\n'.join(lines[start:end])
                    flow_sections.append(flow)
                    break

        if flow_sections:
            flow_text = '\n\n'.join(flow_sections[:2])  # 最多2个流程
            if len(flow_text) > 400:
                flow_text = flow_text[:400] + "..."
            summary_parts.append(f"处理流程:\n{flow_text}")

        # 5. 文档元信息
        total_chars = len(document_content)
        total_lines = len(lines)
        summary_parts.append(f"文档规模: {total_lines}行, {total_chars}字符")

        summary = '\n\n'.join(summary_parts)

        # 强制限制长度
        if len(summary) > max_length:
            # 进一步压缩：移除较长的部分
            parts = summary.split('\n\n')
            compressed_parts = []
            current_length = 0

            for part in parts:
                if current_length + len(part) + 2 <= max_length * 0.8:  # 留20%空间
                    compressed_parts.append(part)
                    current_length += len(part) + 2
                else:
                    # 压缩这个部分
                    if len(part) > 200:
                        compressed_part = part[:200] + "..."
                        if current_length + len(compressed_part) + 2 <= max_length:
                            compressed_parts.append(compressed_part)
                            current_length += len(compressed_part) + 2
                    break

            summary = '\n\n'.join(compressed_parts)
            summary += f"\n\n[文档已大幅压缩，原文档{total_chars}字符]"

        return summary

    def _compress_structure(self, structure: Dict[str, Any], structure_type: str) -> str:
        """压缩文件结构信息"""
        if not structure:
            return f"{structure_type}: [空结构]"

        try:
            # 提取关键信息
            if structure_type == "源文件":
                files = structure.get('files', {})
                compressed_info = []

                for file_name, file_info in files.items():
                    # 只提取基本信息
                    sheets = file_info.get('sheets', {})
                    sheet_list = []

                    for sheet_name, sheet_info in sheets.items():
                        headers = list(sheet_info.get('headers', {}).keys())
                        # 只取前5个表头
                        if len(headers) > 5:
                            headers = headers[:5] + [f"...等{len(headers)-5}列"]

                        sheet_list.append(f"  - {sheet_name}: {', '.join(headers)}")

                    sheets_text = '\n'.join(sheet_list[:3])  # 最多3个sheet
                    if len(sheets) > 3:
                        sheets_text += f"\n  ...等{len(sheets)-3}个sheet"

                    compressed_info.append(f"{file_name}:\n{sheets_text}")

                files_text = '\n\n'.join(compressed_info[:3])  # 最多3个文件
                if len(files) > 3:
                    files_text += f"\n\n...等{len(files)-3}个文件"

                return f"源文件结构:\n{files_text}"

            elif structure_type == "预期输出文件":
                sheets = structure.get('sheets', {})
                compressed_info = []

                for sheet_name, sheet_info in sheets.items():
                    headers = list(sheet_info.get('headers', {}).keys())
                    # 只取前8个表头
                    if len(headers) > 8:
                        headers = headers[:8] + [f"...等{len(headers)-8}列"]

                    data_sample = sheet_info.get('data_sample', [])
                    sample_count = len(data_sample)

                    compressed_info.append(f"  - {sheet_name}: {', '.join(headers)}")
                    if sample_count > 0:
                        compressed_info[-1] += f" (有{sample_count}条数据示例)"

                sheets_text = '\n'.join(compressed_info)
                return f"预期输出结构:\n{sheets_text}"

            else:
                # 通用压缩
                return f"{structure_type}:\n{json.dumps(structure, ensure_ascii=False)[:500]}..."

        except Exception as e:
            self.logger.warning(f"压缩结构失败: {e}")
            # 返回简化版本
            return f"{structure_type}: [结构信息，详细内容已压缩]"