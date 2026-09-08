"""
规则整理模块
将源文件结构、目标文件结构（含公式）、设计文档发送给 AI，
生成结构化的 rules.md 供下载使用。
"""

import logging
import os
import re
from pathlib import Path
from typing import List, Dict, Optional, Any

logger = logging.getLogger(__name__)

# 文档最大字符数（防止超出 AI 上下文窗口），可通过 .env 配置
MAX_DOC_CHARS = max(1000, int(os.getenv("RULE_ORGANIZE_MAX_DOC_CHARS", "30000")))

# 源数据格式段落标题（用于检测和替换）
SOURCE_DATA_FORMAT_HEADER = "## 源数据格式"


def format_source_data_section(source_info: str) -> str:
    """将 _extract_source_structures() 的输出包装为标准 Markdown 段落。

    此段落会追加到 AI 生成的规则文档末尾，使 rules.md 同时包含业务规则和源数据格式。
    智训时 FormulaCodeGenerator 会检测并替换此段落为实际训练数据的源结构。
    """
    return f"\n\n{SOURCE_DATA_FORMAT_HEADER}\n\n{source_info.strip()}\n"


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
8. **列处理分层分析**：根据每列的数据来源和依赖关系进行分层分类：\
   - 首先确定**主表**：即生成代码时遍历行的基准表（通常是主键所在的表，但需根据业务逻辑判断，不一定是行数最多的表）\
   - L0（主键层）：用于行匹配的主键列（如工号/员工编号/身份证号）\
   - L1（同源层）：与主键在同一张源表（主表）中的列，运行时可直接按行复制，不需要任何查找公式\
   - L2（跨表层）：需要通过主键去其他源表查找的列，运行时必须使用VLOOKUP/INDEX+MATCH等公式\
   - L3（计算层）：依赖L1/L2列进行公式计算的列（如 应发合计=基本工资+绩效+加班费）\
   - L4（复合层）：依赖L3列结果的列（如 实发工资=应发合计-社保-公积金-个税）\
   确定每列的层级，在列级规则的"类型"中标注层级，并在最后的"列处理分层"章节汇总。
9. **字段类型必须落入规则**：excel_parser 已根据列头含义、样例值和单元格格式给出每列的
   `字段类型/格式类型/number_format`。源列参与筛选、计算、排序和主键关联时必须先按字段类型读取；
   非模板输出写入时必须应用对应 number_format。主键两侧类型必须一致，文本主键不得转数字。
10. **证据与冲突检查**：原始公式、SOP 条款、样本推断必须明确区分；每条规则注明依据的
    文件/Sheet/单元格或章节。除用户明确要求覆盖外，样例 Excel 中已经明确的公式逻辑
    优先于 SOP、其他规则文档和样本数值推断；保留条件分支、舍入顺序、绝对引用和例外。
    与 SOP 或其他文档冲突时按样例明确公式整理，并单独列出冲突及采用依据，
    不得擅自补造常量或条件；没有明确公式的部分再依据 SOP 和规则文档补充。
11. **主键与依赖检查**：只凭少量样本不能断定主键全量唯一；说明记录粒度、空键/重复键策略，
    必要时采用工号+月份等复合键。跨表查找遇到一对多必须明确汇总或展开规则。
    输出前检查列引用是否存在、公式是否循环依赖、日期空值/边界、同源与跨表引用是否一致。
    不确定或图片中尚未提取的内容列入待确认事项，不要猜测。

## 中间计算项处理规则
当目标列的计算需要依赖不在源文件/目标文件中的中间值时：
- 这些中间计算项应作为**新增列**，按列字母顺序紧接在目标文件最后一列之后编排
- 中间项列使用 **淡蓝色背景(#DCE6F1)** 标识，与原始目标列区分
- **中间项必须写在"列级规则"章节内**，与原始目标列统一编排，不要单独分章节
- 中间项的类型标注为"中间计算项"，并注明被哪些目标列依赖
- 常见中间项举例：累计税前收入、累计专项扣除、累计已预缴税额、本月应纳税所得额等

### 重要：识别设计文档中的公共中间字段
设计文档中经常用以下方式定义**公共中间字段**，这些字段被多个目标列共用，\
必须识别出来作为中间计算项，否则每个依赖它的列都要内联完整公式，导致公式极长：

1. **@前缀变量**：如 `@工资基数 = 全量表.入职杰浦基本工资`，`@` 明确标识为可复用的中间值
2. **"其中："段落里的命名子表达式**：如
   ```
   工资基数 = IF(特殊人员表.工资基数 不为空, ...)
   其中：
     试用期内金额 = IF(...)
     最低工资标准 = IF(地区=上海, 2740, 2320)
   ```
   这里 `试用期内金额` 和 `最低工资标准` 就是中间字段
3. **多个目标列共用的子表达式**：如果同一个计算逻辑（如 `历史月已离职判断`、\
   `试用期判断`）在2个以上目标列中出现，应抽取为中间项
4. **临时变量/辅助字段**：如 `PD2 = 薪资结束日（实习生为当月25日，其他为当月最后一天）`

**关键原则**：宁可多建中间列，也不要让公式嵌套超过3层。\
如果一个列的计算公式 IF 嵌套超过3层或长度超过200字符，\
请考虑将其中的子表达式拆分为独立中间列。

## 输出格式要求
请严格按照以下 Markdown 格式输出：

# 数据处理规则

## 基本信息
- 数据来源: [列出所有源文件名]
- 输出目标: [目标文件名]

## 列级规则

### {列字母}列: {列名}
- 类型: [计算/直接复制/条件判断/查找匹配/固定值]（层级标注：L0-主键/L1-同源/L2-跨表/L3-计算/L4-复合）
- 公式: [目标文件中的原始公式，如有]
- 规则描述: [用自然语言描述该列的数据来源和处理逻辑]
- 数据来源: [源文件名.Sheet名.列名 → 目标列名]
- 关联键: [如果涉及跨表查找，说明用什么字段关联]
- 精度: [小数位数要求，如有]
- 特殊处理: [边界条件、空值处理等]
- 字段类型: [text/integer/decimal/date/datetime/time/boolean]
- 格式类型: [general/text/integer/decimal/percentage/currency/date/datetime/time/custom]
- number_format: [Excel格式码，如 @、0、0.00、yyyy-mm-dd]

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

## 列处理分层

### 主键: {主键列名}
### 主键来源表: {源文件名.Sheet名}
### 主表: {源文件名.Sheet名}（即生成代码时遍历行的基准表，其所有列属于L1同源层）
说明：主表通常是主键所在的表，但也可能由业务逻辑决定。代码中用 find_source_sheet() 定位，禁止用 max(len(df))。

### L1-同源直接复制（与主键同表，直接按行复制，不需要VLOOKUP）
- {列名}（{源表.列名}）
...

### L2-跨表查找（需通过主键查找其他源表）
- {列名}（{源表名} → VLOOKUP/INDEX+MATCH）
...

### L3-计算列（依赖L1/L2的公式计算）
- {列名} = {计算表达式}（依赖: {依赖列}）
...

### L4-复合列（依赖L3的复合计算）
- {列名} = {计算表达式}（依赖: {依赖列}）
...

（如某层级没有列则省略该层级）

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
        target_file: Optional[str] = None,
        design_doc_files: Optional[List[str]] = None,
        file_passwords: Optional[Dict[str, str]] = None,
        user_message: Optional[str] = None,
    ) -> str:
        """主入口：返回生成的 Markdown 字符串"""
        source_info = self._extract_source_structures(source_files, file_passwords)
        target_info = self._extract_target_structure(target_file, file_passwords)
        design_info = self._extract_design_docs(design_doc_files or [])
        design_info = self._prepare_design_context(design_info, user_message)

        messages = self._build_messages(source_info, target_info, design_info, user_message)
        total_len = sum(len(m["content"]) for m in messages)
        logger.info(f"[规则整理] 开始调用 AI, 消息总长度: {total_len}")

        result = self.ai_provider.chat(messages)

        # 追加源数据格式段落，使 rules.md 更完整
        result = result.rstrip() + format_source_data_section(source_info)

        logger.info(f"[规则整理] AI 返回, 结果长度: {len(result)}")
        return result

    def organize_rules_stream(
        self,
        source_files: List[str],
        target_file: Optional[str] = None,
        design_doc_files: Optional[List[str]] = None,
        file_passwords: Optional[Dict[str, str]] = None,
        user_message: Optional[str] = None,
        chunk_callback: Any = None,
        thinking_callback: Any = None,
    ) -> str:
        """流式整理规则：逐块回调 + 返回完整结果。

        thinking_callback: 推理模型思考过程逐块回调（与正式内容分开，不污染结果）。
        """
        source_info = self._extract_source_structures(source_files, file_passwords)
        target_info = self._extract_target_structure(target_file, file_passwords)
        design_info = self._extract_design_docs(design_doc_files or [])
        design_info = self._prepare_design_context(design_info, user_message, thinking_callback)

        messages = self._build_messages(source_info, target_info, design_info, user_message)
        total_len = sum(len(m["content"]) for m in messages)
        logger.info(f"[规则整理-流式] 开始调用 AI, 消息总长度: {total_len}")

        result = self.ai_provider.chat_stream(
            messages, chunk_callback=chunk_callback, thinking_callback=thinking_callback)

        # 追加源数据格式段落，使 rules.md 更完整
        source_section = format_source_data_section(source_info)
        result = result.rstrip() + source_section
        if chunk_callback:
            chunk_callback(source_section)

        logger.info(f"[规则整理-流式] AI 返回, 结果长度: {len(result)}")
        return result

    def chat_followup(
        self,
        messages: List[Dict[str, str]],
        chunk_callback: Any = None,
        thinking_callback: Any = None,
    ) -> str:
        """多轮追问：接受完整对话历史，返回 AI 新回复"""
        logger.info(f"[规则整理-追问] 消息轮数: {len(messages)}")
        result = self.ai_provider.chat_stream(
            messages, chunk_callback=chunk_callback, thinking_callback=thinking_callback)
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
        """解析源文件结构与多行样例，并独立扫描全表公式。"""
        passwords = file_passwords or {}
        parts: List[str] = []

        for file_path in source_files:
            file_name = Path(file_path).name
            try:
                parsed_data = self.excel_parser.parse_excel_file(
                    file_path,
                    max_data_rows=5,
                    active_sheet_only=False,
                    best_region_only=False,
                    password=passwords.get(file_name),
                    read_formulas=True,
                )
                part = f"=== 源文件: {file_name} ===\n"
                for sheet_data in parsed_data:
                    part += f"Sheet: {sheet_data.sheet_name}\n"
                    for region in sheet_data.regions:
                        headers_str = ", ".join(
                            f"{name}={col}" for name, col in region.head_data.items()
                        )
                        part += f"  表头: {headers_str}\n"
                        part += self._format_region_schemas(region)
                        for sample_index, sample in enumerate(region.data, 1):
                            sample_str = ", ".join(
                                f"{col}={val}" for col, val in sample.items()
                            )
                            part += f"  数据样本({sample_index}): {sample_str}\n"
                        for address, formula in region.formula.items():
                            part += f"  样本公式 {address}: {formula}\n"
                part += self._formula_context(file_path)
                parts.append(part)
            except Exception as e:
                logger.warning(f"[规则整理] 解析源文件失败 {file_name}: {e}")
                parts.append(f"=== 源文件: {file_name} (解析失败: {e}) ===\n")

        return "\n".join(parts)

    def _extract_target_structure(
        self,
        target_file: Optional[str],
        file_passwords: Optional[Dict[str, str]] = None,
    ) -> str:
        """解析目标文件，提取表头 + 1 行数据样本 + 公式。target_file 为 None 时返回占位说明。"""
        if not target_file:
            return (
                "（用户未提供目标文件，请根据源文件结构和设计文档自由设计输出格式：\n"
                "  - 自行规划目标列名、类型、来源、计算逻辑\n"
                "  - 在规则文档中明确说明你设计的输出表结构）\n"
            )
        passwords = file_passwords or {}
        file_name = Path(target_file).name

        try:
            parsed_data = self.excel_parser.parse_excel_file(
                target_file,
                max_data_rows=5,
                active_sheet_only=False,
                best_region_only=False,
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
                    result += self._format_region_schemas(region)
                    # 数据样本
                    for sample_index, sample in enumerate(region.data, 1):
                        sample_str = ", ".join(
                            f"{col}={val}" for col, val in sample.items()
                        )
                        result += f"  数据样本({sample_index}): {sample_str}\n"
                    # 公式（关键：目标文件的公式能帮助 AI 推断计算逻辑）
                    if region.formula:
                        result += "  公式:\n"
                        for cell_addr, formula in sorted(region.formula.items()):
                            result += f"    {cell_addr}: {formula}\n"
            return result + self._formula_context(target_file)
        except Exception as e:
            logger.warning(f"[规则整理] 解析目标文件失败 {file_name}: {e}")
            return f"=== 目标文件: {file_name} (解析失败: {e}) ===\n"

    @staticmethod
    def _format_region_schemas(region) -> str:
        """把解析器字段定义写进 AI 可读结构；列名而非样例值决定引用身份。"""
        schemas = getattr(region, "column_schemas", None) or {}
        if not schemas:
            return ""
        rows = []
        for name, letter in region.head_data.items():
            schema = schemas.get(letter) or {}
            rows.append(
                f"{name}({letter}):字段类型={schema.get('field_type', 'text')},"
                f"格式类型={schema.get('format_type', 'general')},"
                f"number_format={schema.get('number_format', 'General')}"
            )
        return "  字段定义: " + "; ".join(rows) + "\n"

    def _extract_design_docs(self, design_doc_files: List[str]) -> str:
        """用 DocumentParser 提取设计文档文本"""
        if not design_doc_files:
            return "(无设计文档)"

        parts: List[str] = []
        for file_path in design_doc_files:
            file_name = Path(file_path).name
            try:
                content = self.doc_parser.parse_document(file_path)
                if not content.strip() or re.match(r'^\[.*(?:失败|未提取到文本)', content):
                    raise ValueError(f'设计文档未能完整读取: {file_name}: {content}')
                parts.append(f"=== 设计文档: {file_name} ===\n{content}\n")
            except Exception as e:
                logger.warning(f"[规则整理] 读取设计文档失败 {file_name}: {e}")
                raise ValueError(f'读取设计文档失败 {file_name}: {e}') from e

        return "\n".join(parts)

    @staticmethod
    def _formula_context(file_path):
        from backend.utils.formula_evidence import collect_formula_evidence, format_formula_evidence
        try:
            return format_formula_evidence(collect_formula_evidence(file_path))
        except Exception as exc:
            # 加密 xlsx / 旧 xls 的公式仍来自带密码的解析器；必须标明扫描范围限制。
            logger.warning('完整公式扫描失败 %s: %s', Path(file_path).name, exc)
            return '注意：全表公式扫描未完成，只提供了解析样本中的公式，不得推断其他区域没有公式。\n'

    def _prepare_design_context(self, content, user_message=None, progress=None):
        """长文档逐段阅读后整合；每段有全局目录及相邻上下文，绝不截掉原文尾部。"""
        if len(content) <= MAX_DOC_CHARS:
            return content
        outline = '\n'.join(line for line in content.splitlines()
                            if re.match(r'^(#{1,6}\s|===|\d+[.、]\s*)', line))
        chunks, start = [], 0
        while start < len(content):
            end = min(start + MAX_DOC_CHARS, len(content))
            if end < len(content):
                boundary = content.rfind('\n', start + MAX_DOC_CHARS // 2, end)
                if boundary > start:
                    end = boundary + 1
            chunks.append((start, end))
            start = end
        notes = []
        for index, (start, end) in enumerate(chunks, 1):
            message = f'正在分析 SOP 第 {index}/{len(chunks)} 段（全文 {len(content)} 字符）'
            logger.info(message)
            if progress:
                progress(message + '\n')
            prompt = (f'用户需求：{user_message or "整理业务规则"}\n全文目录：\n{outline}\n'
                      f'本段 {index}/{len(chunks)}，原文字符范围 [{start}, {end})：\n{content[start:end]}\n'
                      f'前文衔接：{content[max(0, start-500):start]}\n后文衔接：{content[end:end+500]}\n'
                      '逐条提取本段规则，保留原章节、表格对应关系、公式、参数、条件、例外、跨章节引用；'
                      '不要猜测未定义项，列出歧义和图片内容缺失。用于后续全局整合，不能只写概括性摘要。')
            result = self.ai_provider.chat([
                {'role': 'system', 'content': '你负责阅读 SOP 分段并提取可追溯的完整规则条目。'},
                {'role': 'user', 'content': prompt},
            ])
            if not result or not result.strip():
                raise ValueError(f'SOP 第 {index} 段分析结果为空，已停止以避免遗漏规则')
            notes.append(f'### 分段 {index} [{start}, {end})\n{result}')
        return ('全文已逐段送交 AI 阅读；以下为分段提取结果，整合时必须检查跨段依赖和冲突。\n'
                f'## 原文目录\n{outline}\n' + '\n\n'.join(notes))

    def _build_messages(
        self, source_info: str, target_info: str, design_info: str,
        user_message: Optional[str] = None,
    ) -> List[Dict[str, str]]:
        """构建发送给 AI 的消息列表"""
        user_req_section = ""
        if user_message and user_message.strip():
            user_req_section = (
                "\n## 用户额外需求（最高优先级）\n"
                "以下是用户在对话中明确提出的需求，请在生成规则文档时严格遵守，"
                "不得遗漏；如与设计文档冲突，以用户需求为准。\n\n"
                f"{user_message.strip()}\n"
            )

        user_content = f"""\
以下是需要分析的数据文件结构和设计文档。
{user_req_section}
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
