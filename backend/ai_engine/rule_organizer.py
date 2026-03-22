"""
规则整理模块
将源文件结构、目标文件结构（含公式）、设计文档发送给 AI，
生成结构化的 rules.md 供下载使用。
"""

import logging
import os
from pathlib import Path
from typing import List, Dict, Optional, Any

logger = logging.getLogger(__name__)

# 文档最大字符数（防止超出 AI 上下文窗口），可通过 .env 配置
MAX_DOC_CHARS = int(os.getenv("RULE_ORGANIZE_MAX_DOC_CHARS", "30000"))


SYSTEM_PROMPT = """\
你是一位专业的数据处理规则分析师，擅长分析Excel文件结构和业务规则，\
尤其熟悉人力资源领域的薪资计算、考勤管理、社保税务等业务场景。

## 任务
根据提供的源文件结构、目标文件结构（含公式）、以及设计文档，\
整理出一份结构化的数据处理规则文档(Markdown格式)。

## 分析要求
1. 分析目标文件中每一列的数据来源和计算逻辑
2. 结合公式信息推断计算规则
3. 结合设计文档中的业务描述补充规则细节
4. 识别源文件与目标文件之间的列映射关系
5. 识别需要跨表查找的数据关联关系
6. 识别数据清洗、精度、条件格式等规则
7. **识别中间计算项**：某些目标列的计算依赖于中间结果（如累计税前收入、累计已缴税额等），\
这些中间结果在源文件和目标文件中都不存在，但计算逻辑必须依赖它们。\
请识别出所有需要的中间计算项。

## 中间计算项处理规则
当目标列的计算需要依赖不在源文件/目标文件中的中间值时：
- 这些中间计算项应作为**新增列**，按列字母顺序紧接在目标文件最后一列之后编排
- 中间项列使用 **淡蓝色背景(#DCE6F1)** 标识，与原始目标列区分
- **中间项必须写在"列级规则"章节内**，与原始目标列统一编排，不要单独分章节
- 中间项的类型标注为"中间计算项"，并注明被哪些目标列依赖
- 常见中间项举例：累计税前收入、累计专项扣除、累计已预缴税额、本月应纳税所得额等

## 输出格式要求
请严格按照以下 Markdown 格式输出：

# 数据处理规则

## 基本信息
- 数据来源: [列出所有源文件名]
- 输出目标: [目标文件名]

## 列级规则

### {列字母}列: {列名}
- 类型: [计算/直接复制/条件判断/查找匹配/固定值]
- 公式: [目标文件中的原始公式，如有]
- 规则描述: [用自然语言描述该列的数据来源和处理逻辑]
- 数据来源: [源文件名.Sheet名.列名 → 目标列名]
- 关联键: [如果涉及跨表查找，说明用什么字段关联]
- 精度: [小数位数要求，如有]
- 特殊处理: [边界条件、空值处理等]

（先按目标文件的列字母顺序逐列输出原始列，然后紧接着输出中间计算项列）

### {列字母}列: {中间项名称}（中间项，淡蓝色背景 #DCE6F1）
- 类型: 中间计算项
- 规则描述: [该中间项的计算逻辑]
- 依赖: [计算该中间项需要的输入列]
- 被依赖: [哪些目标列的计算需要用到此中间项]

（如不需要中间项则不输出此部分）

## 数据清洗规则
- [列出需要过滤、去重、格式转换等清洗规则]

## 全局规则
- [列出适用于所有列的通用规则，如主键选择、排序方式等]

## 注意事项
- [列出实现时需要注意的特殊情况]
"""


class RuleOrganizer:
    """规则整理器：利用 AI 从源文件/目标文件/设计文档生成结构化规则"""

    def __init__(self, ai_provider):
        self.ai_provider = ai_provider

        # 延迟导入，避免模块级循环依赖
        import sys, os
        project_root = str(Path(__file__).resolve().parent.parent.parent)
        if project_root not in sys.path:
            sys.path.insert(0, project_root)

        from excel_parser import IntelligentExcelParser
        self.excel_parser = IntelligentExcelParser()

        from backend.ai_engine.document_parser import get_document_parser
        self.doc_parser = get_document_parser()

    # ------------------------------------------------------------------
    # 公开入口
    # ------------------------------------------------------------------
    def organize_rules(
        self,
        source_files: List[str],
        target_file: str,
        design_doc_files: List[str],
        file_passwords: Optional[Dict[str, str]] = None,
    ) -> str:
        """主入口：返回生成的 Markdown 字符串"""
        source_info = self._extract_source_structures(source_files, file_passwords)
        target_info = self._extract_target_structure(target_file, file_passwords)
        design_info = self._extract_design_docs(design_doc_files)

        messages = self._build_messages(source_info, target_info, design_info)
        total_len = sum(len(m["content"]) for m in messages)
        logger.info(f"[规则整理] 开始调用 AI, 消息总长度: {total_len}")

        result = self.ai_provider.chat(messages)

        logger.info(f"[规则整理] AI 返回, 结果长度: {len(result)}")
        return result

    def organize_rules_stream(
        self,
        source_files: List[str],
        target_file: str,
        design_doc_files: List[str],
        file_passwords: Optional[Dict[str, str]] = None,
        chunk_callback: Any = None,
    ) -> str:
        """流式整理规则：逐块回调 + 返回完整结果"""
        source_info = self._extract_source_structures(source_files, file_passwords)
        target_info = self._extract_target_structure(target_file, file_passwords)
        design_info = self._extract_design_docs(design_doc_files)

        messages = self._build_messages(source_info, target_info, design_info)
        total_len = sum(len(m["content"]) for m in messages)
        logger.info(f"[规则整理-流式] 开始调用 AI, 消息总长度: {total_len}")

        result = self.ai_provider.chat_stream(messages, chunk_callback=chunk_callback)

        logger.info(f"[规则整理-流式] AI 返回, 结果长度: {len(result)}")
        return result

    def chat_followup(
        self,
        messages: List[Dict[str, str]],
        chunk_callback: Any = None,
    ) -> str:
        """多轮追问：接受完整对话历史，返回 AI 新回复"""
        logger.info(f"[规则整理-追问] 消息轮数: {len(messages)}")
        result = self.ai_provider.chat_stream(messages, chunk_callback=chunk_callback)
        logger.info(f"[规则整理-追问] AI 返回, 结果长度: {len(result)}")
        return result

    # ------------------------------------------------------------------
    # 内部方法
    # ------------------------------------------------------------------
    def _extract_source_structures(
        self,
        source_files: List[str],
        file_passwords: Optional[Dict[str, str]] = None,
    ) -> str:
        """解析源文件，提取表头 + 1 行数据样本"""
        passwords = file_passwords or {}
        parts: List[str] = []

        for file_path in source_files:
            file_name = Path(file_path).name
            try:
                parsed_data = self.excel_parser.parse_excel_file(
                    file_path,
                    max_data_rows=1,
                    active_sheet_only=True,
                    best_region_only=True,
                    password=passwords.get(file_name),
                )
                part = f"=== 源文件: {file_name} ===\n"
                for sheet_data in parsed_data:
                    part += f"Sheet: {sheet_data.sheet_name}\n"
                    for region in sheet_data.regions:
                        headers_str = ", ".join(
                            f"{name}={col}" for name, col in region.head_data.items()
                        )
                        part += f"  表头: {headers_str}\n"
                        if region.data:
                            sample_str = ", ".join(
                                f"{col}={val}" for col, val in region.data[0].items()
                            )
                            part += f"  数据样本(第1行): {sample_str}\n"
                parts.append(part)
            except Exception as e:
                logger.warning(f"[规则整理] 解析源文件失败 {file_name}: {e}")
                parts.append(f"=== 源文件: {file_name} (解析失败: {e}) ===\n")

        return "\n".join(parts)

    def _extract_target_structure(
        self,
        target_file: str,
        file_passwords: Optional[Dict[str, str]] = None,
    ) -> str:
        """解析目标文件，提取表头 + 1 行数据样本 + 公式"""
        passwords = file_passwords or {}
        file_name = Path(target_file).name

        try:
            parsed_data = self.excel_parser.parse_excel_file(
                target_file,
                max_data_rows=1,
                active_sheet_only=True,
                best_region_only=True,
                password=passwords.get(file_name),
            )
            result = f"=== 目标文件: {file_name} ===\n"
            for sheet_data in parsed_data:
                result += f"Sheet: {sheet_data.sheet_name}\n"
                for region in sheet_data.regions:
                    # 表头
                    headers_str = ", ".join(
                        f"{name}={col}" for name, col in region.head_data.items()
                    )
                    result += f"  表头: {headers_str}\n"
                    # 数据样本
                    if region.data:
                        sample_str = ", ".join(
                            f"{col}={val}" for col, val in region.data[0].items()
                        )
                        result += f"  数据样本(第1行): {sample_str}\n"
                    # 公式（关键：目标文件的公式能帮助 AI 推断计算逻辑）
                    if region.formula:
                        result += "  公式:\n"
                        for cell_addr, formula in sorted(region.formula.items()):
                            result += f"    {cell_addr}: {formula}\n"
            return result
        except Exception as e:
            logger.warning(f"[规则整理] 解析目标文件失败 {file_name}: {e}")
            return f"=== 目标文件: {file_name} (解析失败: {e}) ===\n"

    def _extract_design_docs(self, design_doc_files: List[str]) -> str:
        """用 DocumentParser 提取设计文档文本"""
        if not design_doc_files:
            return "(无设计文档)"

        parts: List[str] = []
        for file_path in design_doc_files:
            file_name = Path(file_path).name
            try:
                content = self.doc_parser.parse_document(file_path)
                if len(content) > MAX_DOC_CHARS:
                    logger.warning(
                        f"[规则整理] 设计文档 {file_name} 过长 ({len(content)} 字符)，截断至 {MAX_DOC_CHARS}"
                    )
                    content = content[:MAX_DOC_CHARS] + "\n... (文档过长，已截断)"
                parts.append(f"=== 设计文档: {file_name} ===\n{content}\n")
            except Exception as e:
                logger.warning(f"[规则整理] 读取设计文档失败 {file_name}: {e}")
                parts.append(f"=== 设计文档: {file_name} (读取失败: {e}) ===\n")

        return "\n".join(parts)

    def _build_messages(
        self, source_info: str, target_info: str, design_info: str
    ) -> List[Dict[str, str]]:
        """构建发送给 AI 的消息列表"""
        user_content = f"""\
以下是需要分析的数据文件结构和设计文档。

## 源文件结构
{source_info}

## 目标文件结构（含公式）
{target_info}

## 设计文档
{design_info}

请根据以上信息，整理出完整的数据处理规则文档。"""

        return [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ]
