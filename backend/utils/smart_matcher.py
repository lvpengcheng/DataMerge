"""
智能文件和表头匹配器 - 使用AI进行灵活匹配
"""

import os
import json
import logging
from typing import Dict, List, Any, Tuple, Optional
from pathlib import Path
import pandas as pd

logger = logging.getLogger(__name__)


class SmartMatcher:
    """智能匹配器 - 使用AI匹配文件名、sheet名和表头"""

    def __init__(self, ai_provider=None):
        """初始化

        Args:
            ai_provider: AI提供者实例
        """
        self.ai_provider = ai_provider

    def match_files_and_headers(
        self,
        training_structure: Dict[str, Any],
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """匹配文件和表头

        Args:
            training_structure: 训练时的文件结构
            input_files: 输入文件列表
            manual_headers: 手动表头配置

        Returns:
            (是否成功, 错误信息, 映射关系)
            映射关系格式: {
                "file_mapping": {
                    "实际文件路径": {
                        "expected_file": "训练时的文件名",
                        "sheet_mapping": {
                            "实际sheet名": "训练时的sheet名"
                        },
                        "header_mapping": {
                            "实际列名": "训练时的列名"
                        }
                    }
                }
            }
        """
        # 1. 先尝试完全匹配
        exact_match, mapping = self._try_exact_match(training_structure, input_files, manual_headers)
        if exact_match:
            logger.info("文件和表头完全匹配，直接使用")
            return True, None, mapping

        # 2. 完全匹配失败，使用AI进行智能匹配
        logger.info("文件或表头不完全匹配，使用AI进行智能匹配...")

        if not self.ai_provider:
            return False, "文件或表头不匹配，且未配置AI进行智能匹配", None

        try:
            ai_match, error_msg, ai_mapping = self._ai_match(training_structure, input_files, manual_headers)
            if ai_match:
                logger.info("AI匹配成功")
                return True, None, ai_mapping
            else:
                logger.error(f"AI匹配失败: {error_msg}")
                return False, error_msg, None
        except Exception as e:
            logger.error(f"AI匹配过程出错: {e}")
            return False, f"AI匹配失败: {str(e)}", None

    def _try_exact_match(
        self,
        training_structure: Dict[str, Any],
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, Optional[Dict[str, Any]]]:
        """尝试完全匹配

        Args:
            training_structure: 训练时的文件结构
            input_files: 输入文件列表
            manual_headers: 手动表头配置

        Returns:
            (是否完全匹配, 映射关系)
        """
        training_files = training_structure.get("files", {})

        # 提取输入文件的基本信息
        input_file_info = {}
        for file_path in input_files:
            file_name = os.path.basename(file_path)
            try:
                # 使用 IntelligentExcelParser 读取Excel文件
                from excel_parser import IntelligentExcelParser

                # 如果有手动表头配置，传递给 parser
                file_manual_headers = None
                if manual_headers:
                    file_manual_headers = manual_headers.get(file_name)

                parser = IntelligentExcelParser(file_path, manual_headers=file_manual_headers)

                sheets = {}
                for sheet_name in parser.sheet_names:
                    df = parser.parse_sheet(sheet_name)
                    if df is not None:
                        sheets[sheet_name] = list(df.columns)

                input_file_info[file_path] = {
                    "file_name": file_name,
                    "sheets": sheets
                }
            except Exception as e:
                logger.warning(f"读取文件 {file_path} 失败: {e}")
                continue

        # 检查是否完全匹配
        mapping = {"file_mapping": {}}

        for file_path, file_info in input_file_info.items():
            file_name = file_info["file_name"]

            # 查找训练时的对应文件
            matched_training_file = None
            for training_file_name in training_files.keys():
                if training_file_name == file_name:
                    matched_training_file = training_file_name
                    break

            if not matched_training_file:
                # 文件名不匹配
                return False, None

            # 检查sheet和表头是否匹配
            training_sheets = training_files[matched_training_file].get("sheets", {})
            file_mapping = {
                "expected_file": matched_training_file,
                "sheet_mapping": {},
                "header_mapping": {}
            }

            for sheet_name, headers in file_info["sheets"].items():
                # 查找对应的训练sheet
                if sheet_name not in training_sheets:
                    # sheet名不匹配
                    return False, None

                training_headers = training_sheets[sheet_name].get("headers", {})
                if isinstance(training_headers, dict):
                    training_header_list = list(training_headers.keys())
                else:
                    training_header_list = training_headers

                # 检查表头是否完全匹配
                if set(headers) != set(training_header_list):
                    # 表头不匹配
                    return False, None

                # 记录映射（完全匹配时是恒等映射）
                file_mapping["sheet_mapping"][sheet_name] = sheet_name
                for header in headers:
                    file_mapping["header_mapping"][header] = header

            mapping["file_mapping"][file_path] = file_mapping

        return True, mapping

    def _ai_match(
        self,
        training_structure: Dict[str, Any],
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """使用AI进行智能匹配

        Args:
            training_structure: 训练时的文件结构
            input_files: 输入文件列表
            manual_headers: 手动表头配置

        Returns:
            (是否匹配成功, 错误信息, 映射关系)
        """
        # 提取输入文件信息
        input_file_info = {}
        for file_path in input_files:
            file_name = os.path.basename(file_path)
            try:
                # 使用 IntelligentExcelParser 读取Excel文件
                from excel_parser import IntelligentExcelParser

                # 如果有手动表头配置，传递给 parser
                file_manual_headers = None
                if manual_headers:
                    file_manual_headers = manual_headers.get(file_name)

                parser = IntelligentExcelParser(file_path, manual_headers=file_manual_headers)

                sheets = {}
                for sheet_name in parser.sheet_names:
                    df = parser.parse_sheet(sheet_name)
                    if df is not None:
                        sheets[sheet_name] = list(df.columns)

                input_file_info[file_name] = {
                    "file_path": file_path,
                    "sheets": sheets
                }
            except Exception as e:
                logger.warning(f"读取文件 {file_path} 失败: {e}")
                continue

        # 提取训练时的文件结构
        training_files = training_structure.get("files", {})
        training_file_info = {}
        for file_name, file_data in training_files.items():
            sheets = {}
            for sheet_name, sheet_data in file_data.get("sheets", {}).items():
                headers = sheet_data.get("headers", {})
                if isinstance(headers, dict):
                    sheets[sheet_name] = list(headers.keys())
                else:
                    sheets[sheet_name] = headers
            training_file_info[file_name] = sheets

        # 构建AI提示词
        prompt = f"""你是一个Excel文件结构匹配专家。请分析以下两组文件结构，判断它们是否可以匹配，并生成映射关系。

## 训练时的文件结构（期望的结构）
{json.dumps(training_file_info, ensure_ascii=False, indent=2)}

## 当前上传的文件结构
{json.dumps({k: v["sheets"] for k, v in input_file_info.items()}, ensure_ascii=False, indent=2)}

## 任务
1. 判断当前上传的文件是否可以匹配到训练时的文件
2. 如果可以匹配，生成详细的映射关系（文件名映射、sheet名映射、列名映射）
3. 如果无法匹配，说明原因

## 匹配规则
- 文件名可以不完全相同，但业务含义应该一致（如"员工信息.xlsx"可以匹配"员工基本信息表.xlsx"）
- Sheet名可以不完全相同，但业务含义应该一致
- 列名可以不完全相同，但业务含义应该一致（如"姓名"可以匹配"员工姓名"）
- 如果训练时有多个文件，当前上传的文件数量应该匹配
- 如果某些必需的列在当前文件中找不到对应列，则匹配失败

## 输出格式
请输出JSON格式，结构如下：
```json
{{
    "success": true/false,
    "error_message": "如果失败，说明原因",
    "file_mapping": {{
        "当前文件名": {{
            "expected_file": "训练时的文件名",
            "sheet_mapping": {{
                "当前sheet名": "训练时的sheet名"
            }},
            "header_mapping": {{
                "当前列名": "训练时的列名"
            }}
        }}
    }}
}}
```

注意：
1. 只输出JSON，不要有其他内容
2. 如果无法匹配，success设为false，并在error_message中详细说明原因
3. 映射关系要完整，包含所有文件、sheet和列的映射
"""

        try:
            # 调用AI（使用流式调用）
            messages = [{"role": "user", "content": prompt}]

            # 尝试使用流式调用
            response = ""
            logger.info("开始AI智能匹配...")
            try:
                # 检查是否有流式方法
                if hasattr(self.ai_provider, '_openai_chat_stream'):
                    logger.info("使用 OpenAI 流式调用")
                    for chunk, finish_reason in self.ai_provider._openai_chat_stream(messages):
                        if chunk:
                            response += chunk
                            # 实时输出到控制台
                            import sys
                            sys.stdout.write(chunk)
                            sys.stdout.flush()
                elif hasattr(self.ai_provider, '_claude_chat_stream'):
                    logger.info("使用 Claude 流式调用")
                    for chunk, finish_reason in self.ai_provider._claude_chat_stream("", messages):
                        if chunk:
                            response += chunk
                            # 实时输出到控制台
                            import sys
                            sys.stdout.write(chunk)
                            sys.stdout.flush()
                else:
                    # 回退到非流式
                    logger.warning("AI provider 不支持流式调用，使用非流式")
                    response = self.ai_provider.chat(messages)
            except Exception as stream_error:
                logger.warning(f"流式调用失败，回退到非流式: {stream_error}")
                response = self.ai_provider.chat(messages)

            logger.info(f"\nAI响应长度: {len(response)} 字符")

            # 解析AI响应
            result = self._parse_ai_response(response)

            if not result:
                return False, "AI响应格式错误", None

            if not result.get("success"):
                error_msg = result.get("error_message", "AI判断无法匹配")
                return False, error_msg, None

            # 转换映射关系（将文件名映射转换为文件路径映射）
            file_mapping = result.get("file_mapping", {})
            path_mapping = {}

            for current_file_name, mapping_info in file_mapping.items():
                # 找到对应的文件路径
                file_path = input_file_info.get(current_file_name, {}).get("file_path")
                if file_path:
                    path_mapping[file_path] = mapping_info

            final_mapping = {"file_mapping": path_mapping}
            return True, None, final_mapping

        except Exception as e:
            logger.error(f"AI匹配失败: {e}")
            return False, f"AI匹配过程出错: {str(e)}", None

    def _parse_ai_response(self, response: str) -> Optional[Dict[str, Any]]:
        """解析AI响应

        Args:
            response: AI响应文本

        Returns:
            解析后的JSON对象
        """
        try:
            # 提取JSON内容
            import re
            json_match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # 尝试直接解析整个响应
                json_str = response.strip()

            result = json.loads(json_str)
            return result
        except Exception as e:
            logger.error(f"解析AI响应失败: {e}, 响应内容: {response[:500]}")
            return None
