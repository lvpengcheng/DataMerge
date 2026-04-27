"""
模块化代码生成器 - 将复杂规则拆分成5个固定步骤生成代码

步骤1: 数据加载模块 - 加载所有数据，获取结构、data和公式
步骤2: 映射关系模块 - 根据规则生成字段映射关系
步骤3: 预期文件模板模块 - 生成空的预期文件结构
步骤4: 公式生成模块 - 分析字段前后关系，生成所有数据公式
步骤5: 数据填充模块 - 填充基础数据，然后根据公式填充其他列
"""

import os
import json
import logging
from typing import Dict, List, Any, Optional, Tuple
from .ai_provider import BaseAIProvider, AIProviderFactory
from .prompt_generator import PromptGenerator

logger = logging.getLogger(__name__)


class ModularCodeGenerator:
    """模块化代码生成器

    使用固定的5步策略生成代码：
    1. 数据加载
    2. 映射关系
    3. 预期文件模板
    4. 公式生成
    5. 数据填充
    """

    # 固定的6个模块定义
    MODULES = [
        {
            "module_id": "module_1_data_loader",
            "module_name": "数据加载模块",
            "file_name": "data_loader.py",
            "description": "加载所有源数据文件，获取数据结构、数据内容和Excel公式",
            "function_name": "load_all_data"
        },
        {
            "module_id": "module_2_mapping",
            "module_name": "映射关系模块",
            "file_name": "field_mapping.py",
            "description": "根据规则生成源字段到目标字段的映射关系",
            "function_name": "generate_field_mapping"
        },
        {
            "module_id": "module_3_template",
            "module_name": "预期文件模板模块",
            "file_name": "output_template.py",
            "description": "生成空的预期输出文件结构（表头、Sheet等）",
            "function_name": "create_output_template"
        },
        {
            "module_id": "module_4_formulas",
            "module_name": "公式生成模块",
            "file_name": "formula_generator.py",
            "description": "分析字段的前后依赖关系，生成所有数据计算公式（包含Excel公式）",
            "function_name": "generate_formulas"
        },
        {
            "module_id": "module_5_filler",
            "module_name": "数据填充模块",
            "file_name": "data_filler.py",
            "description": "填充数据并记录数据来源、计算公式、中间计算列",
            "function_name": "fill_data"
        },
        {
            "module_id": "module_6_excel_output",
            "module_name": "Excel增强输出模块",
            "file_name": "excel_output.py",
            "description": "保存Excel并添加表头批注、过程列淡蓝色背景、Excel公式",
            "function_name": "save_excel_with_details"
        }
    ]

    def __init__(self, ai_provider: BaseAIProvider = None, training_logger = None):
        """初始化模块化代码生成器

        Args:
            ai_provider: AI提供者
            training_logger: 训练日志记录器（可选）
        """
        if ai_provider is None:
            self.ai_provider = AIProviderFactory.create_with_fallback()
        else:
            self.ai_provider = ai_provider

        self.training_logger = training_logger
        self.max_tokens = int(os.getenv("OPENAI_MAX_TOKENS", "8000"))
        self.prompt_generator = PromptGenerator()
        # 批量生成模式开关，可通过环境变量配置
        self.use_batch_mode = os.getenv("MODULAR_BATCH_MODE", "true").lower() == "true"

    def generate_modular_code(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None,
        stream_callback: callable = None,
        salary_year: Optional[int] = None,
        salary_month: Optional[int] = None,
        monthly_standard_hours: Optional[float] = None,
        analysis=None,
    ) -> Tuple[str, List[Dict[str, Any]]]:
        """完整的模块化代码生成流程

        Args:
            rules_content: 规则内容
            source_structure: 源文件结构
            expected_structure: 预期文件结构
            manual_headers: 手动表头配置
            stream_callback: 流式回调函数
            salary_year: 薪资年份（可选）
            salary_month: 薪资月份（可选）
            monthly_standard_hours: 当月标准工时（可选）
            analysis: TableAnalysisResult 预分析结果（可选）

        Returns:
            (完整代码, 模块信息列表)
        """
        # 保存薪资参数供后续使用
        self.salary_year = salary_year
        self.salary_month = salary_month
        self.monthly_standard_hours = monthly_standard_hours
        self._analysis = analysis

        def log(msg):
            logger.info(msg)
            if stream_callback:
                stream_callback(msg)

        # 如果有薪资参数，记录日志
        if monthly_standard_hours is not None:
            log(f"薪资参数: {salary_year}年{salary_month}月, 标准工时: {monthly_standard_hours}")

        # 检查是否使用批量生成模式
        if self.use_batch_mode:
            log("=== 使用批量模式生成所有模块（单次AI调用）===")
            return self._generate_batch_mode(
                rules_content, source_structure, expected_structure,
                manual_headers, stream_callback, log,
                salary_year=salary_year,
                salary_month=salary_month,
                monthly_standard_hours=monthly_standard_hours
            )

        # 原有的逐步生成模式
        log("=== 开始5步模块化代码生成 ===")

        generated_modules = []

        # 依次生成5个模块
        for i, module_def in enumerate(self.MODULES):
            module_num = i + 1
            log(f"\n--- 步骤 {module_num}/5: {module_def['module_name']} ---")

            try:
                code = self._generate_module_with_stream(
                    module_def,
                    rules_content,
                    source_structure,
                    expected_structure,
                    manual_headers,
                    generated_modules,  # 传入已生成的模块供参考
                    stream_callback
                )

                generated_modules.append({
                    "module_id": module_def["module_id"],
                    "module_name": module_def["module_name"],
                    "file_name": module_def["file_name"],
                    "function_name": module_def["function_name"],
                    "code": code
                })

                log(f"模块 {module_def['module_name']} 生成完成，代码长度: {len(code)}")

            except Exception as e:
                log(f"模块 {module_def['module_name']} 生成失败: {e}")
                # 生成占位代码
                placeholder = self._generate_placeholder(module_def)
                generated_modules.append({
                    "module_id": module_def["module_id"],
                    "module_name": module_def["module_name"],
                    "file_name": module_def["file_name"],
                    "function_name": module_def["function_name"],
                    "code": placeholder,
                    "error": str(e)
                })

        # 合并所有模块生成完整代码
        log("\n--- 合并所有模块 ---")
        full_code = self._merge_all_modules(
            generated_modules,
            source_structure,
            expected_structure
        )

        log(f"代码生成完成，总长度: {len(full_code)} 字符")

        # 记录最终生成的代码到日志文件（逐步生成模式）
        if self.training_logger:
            self.training_logger.log_generated_code(full_code)

        return full_code, generated_modules

    def _generate_batch_mode(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]],
        stream_callback: callable,
        log: callable,
        salary_year: Optional[int] = None,
        salary_month: Optional[int] = None,
        monthly_standard_hours: Optional[float] = None
    ) -> Tuple[str, List[Dict[str, Any]]]:
        """批量模式：单次AI调用生成所有5个模块

        通过设计一个包含所有模块要求的提示词，让AI一次性生成所有代码，
        然后使用分隔符解析拆分成各个模块。

        如果批量生成失败（未能解析出足够的模块），自动回退到分步生成模式。
        """
        # 生成批量提示词
        # 如果有analysis，在规则内容前注入分层摘要
        _rules_for_prompt = rules_content
        if self._analysis:
            from backend.ai_engine.table_analyzer import TableAnalyzer
            _ta = TableAnalyzer()
            _layer_summary = _ta.generate_layer_summary(self._analysis)
            if _layer_summary:
                _rules_for_prompt = _layer_summary + "\n\n" + rules_content

        prompt = self.prompt_generator.generate_batch_modular_prompt(
            rules_content=_rules_for_prompt,
            source_structure=source_structure,
            expected_structure=expected_structure,
            manual_headers=manual_headers,
            modules=self.MODULES,
            salary_year=salary_year,
            salary_month=salary_month,
            monthly_standard_hours=monthly_standard_hours
        )

        log(f"批量提示词生成完成，长度: {len(prompt)}")

        # 记录完整提示词到日志文件
        if self.training_logger:
            self.training_logger.log_full_prompt(prompt, "modular_batch")

        # 使用流式调用生成代码
        full_response = ""
        has_stream = hasattr(self.ai_provider, 'generate_code_with_stream') and stream_callback

        if has_stream:
            log("使用流式生成...")
            raw_response = ""

            def chunk_handler(chunk):
                nonlocal raw_response
                raw_response += chunk
                # 直接输出到控制台
                import sys
                sys.stdout.write(chunk)
                sys.stdout.flush()

            # generate_code_with_stream会返回经过extract_python_code处理后的代码
            # 但我们需要原始响应来解析模块，所以使用raw_response
            extracted_code = self.ai_provider.generate_code_with_stream(prompt, chunk_callback=chunk_handler)

            # 使用原始响应来解析模块，因为它包含完整的函数定义
            # 如果原始响应为空，使用提取后的代码
            full_response = raw_response if raw_response else extracted_code

            # 如果还是空，记录警告
            if not full_response:
                log("警告: AI响应为空，可能是API调用问题")
        else:
            log("使用普通生成...")
            full_response = self.ai_provider.generate_code(prompt)

        log(f"\nAI响应完成，总长度: {len(full_response)}")

        # 记录完整AI响应到日志文件
        if self.training_logger:
            self.training_logger.log_full_ai_response(full_response, "modular_batch")

        # 解析响应，拆分成各个模块
        generated_modules, success_count = self._parse_batch_response(full_response, log)

        # 如果成功解析的模块数少于3个，回退到分步生成模式
        if success_count < 3:
            log(f"\n批量生成只成功解析了 {success_count} 个模块，回退到分步生成模式...")
            return self._fallback_to_step_mode(
                rules_content, source_structure, expected_structure,
                manual_headers, stream_callback, log
            )

        # 合并所有模块生成完整代码
        log("\n--- 合并所有模块 ---")
        full_code = self._merge_all_modules(
            generated_modules,
            source_structure,
            expected_structure
        )

        log(f"代码生成完成，总长度: {len(full_code)} 字符")

        # 记录最终生成的代码到日志文件（批量模式）
        if self.training_logger:
            self.training_logger.log_generated_code(full_code)

        return full_code, generated_modules

    def _fallback_to_step_mode(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]],
        stream_callback: callable,
        log: callable
    ) -> Tuple[str, List[Dict[str, Any]]]:
        """回退到分步生成模式"""
        log("=== 使用分步生成模式 ===")

        generated_modules = []

        # 依次生成5个模块
        for i, module_def in enumerate(self.MODULES):
            module_num = i + 1
            log(f"\n--- 步骤 {module_num}/5: {module_def['module_name']} ---")

            try:
                code = self._generate_module_with_stream(
                    module_def,
                    rules_content,
                    source_structure,
                    expected_structure,
                    manual_headers,
                    generated_modules,
                    stream_callback
                )

                generated_modules.append({
                    "module_id": module_def["module_id"],
                    "module_name": module_def["module_name"],
                    "file_name": module_def["file_name"],
                    "function_name": module_def["function_name"],
                    "code": code
                })

                log(f"模块 {module_def['module_name']} 生成完成，代码长度: {len(code)}")

            except Exception as e:
                log(f"模块 {module_def['module_name']} 生成失败: {e}")
                placeholder = self._generate_placeholder(module_def)
                generated_modules.append({
                    "module_id": module_def["module_id"],
                    "module_name": module_def["module_name"],
                    "file_name": module_def["file_name"],
                    "function_name": module_def["function_name"],
                    "code": placeholder,
                    "error": str(e)
                })

        # 合并所有模块生成完整代码
        log("\n--- 合并所有模块 ---")
        full_code = self._merge_all_modules(
            generated_modules,
            source_structure,
            expected_structure
        )

        log(f"代码生成完成，总长度: {len(full_code)} 字符")

        # 记录最终生成的代码到日志文件（回退模式）
        if self.training_logger:
            self.training_logger.log_generated_code(full_code)

        return full_code, generated_modules

    def _parse_batch_response(
        self,
        response: str,
        log: callable
    ) -> Tuple[List[Dict[str, Any]], int]:
        """解析批量生成的响应，拆分成各个模块

        Returns:
            (模块列表, 成功解析的模块数量)
        """
        import re

        generated_modules = []
        success_count = 0

        # 首先尝试提取markdown代码块中的代码
        code_block_pattern = r'```python\s*(.*?)```'
        code_blocks = re.findall(code_block_pattern, response, re.DOTALL)
        if code_blocks:
            # 合并所有代码块
            clean_response = '\n\n'.join(code_blocks)
            log(f"从markdown代码块中提取了 {len(code_blocks)} 个代码段")
        else:
            clean_response = response

        # 预期的分隔符模式: ### MODULE_1: xxx 或 # ===== MODULE_1 =====
        module_patterns = [
            r'###\s*MODULE_(\d+)[:：]?\s*(.*?)(?=###\s*MODULE_\d+|$)',
            r'#\s*=+\s*MODULE_(\d+)\s*=+\s*(.*?)(?=#\s*=+\s*MODULE_\d+|$)',
            r'"""MODULE_(\d+)"""(.*?)(?="""MODULE_\d+"""|$)',
            r'# --- 模块\s*(\d+)[:：]?\s*(.*?)(?=# --- 模块\s*\d+|$)',
        ]

        parsed = False
        for pattern in module_patterns:
            matches = re.findall(pattern, clean_response, re.DOTALL | re.IGNORECASE)
            if matches and len(matches) >= 3:  # 至少找到3个模块才算成功
                parsed = True
                log(f"使用模式匹配成功，找到 {len(matches)} 个模块")

                for module_num_str, code in matches:
                    module_num = int(module_num_str)
                    if 1 <= module_num <= 6:  # 支持6个模块
                        idx = module_num - 1
                        module_def = self.MODULES[idx]

                        # 清理代码
                        cleaned_code = self._clean_module_code(code)

                        # 检查代码是否有效（包含函数定义）
                        if f"def {module_def['function_name']}" in cleaned_code:
                            success_count += 1

                        generated_modules.append({
                            "module_id": module_def["module_id"],
                            "module_name": module_def["module_name"],
                            "file_name": module_def["file_name"],
                            "function_name": module_def["function_name"],
                            "code": cleaned_code
                        })
                break

        if not parsed:
            # 尝试按函数定义拆分
            log("尝试按函数定义拆分...")

            for i, module_def in enumerate(self.MODULES):
                func_name = module_def["function_name"]
                # 查找函数定义（使用clean_response而不是原始response）
                pattern = rf'(def\s+{func_name}\s*\(.*?)(?=def\s+\w+\s*\(|$)'
                match = re.search(pattern, clean_response, re.DOTALL)

                if match:
                    code = match.group(1).strip()
                    generated_modules.append({
                        "module_id": module_def["module_id"],
                        "module_name": module_def["module_name"],
                        "file_name": module_def["file_name"],
                        "function_name": module_def["function_name"],
                        "code": code
                    })
                    success_count += 1
                else:
                    # 生成占位代码
                    log(f"未找到函数 {func_name}，使用占位代码")
                    placeholder = self._generate_placeholder(module_def)
                    generated_modules.append({
                        "module_id": module_def["module_id"],
                        "module_name": module_def["module_name"],
                        "file_name": module_def["file_name"],
                        "function_name": module_def["function_name"],
                        "code": placeholder,
                        "error": "未在AI响应中找到此函数"
                    })

        # 确保有6个模块
        if len(generated_modules) < 6:
            log(f"警告: 只解析到 {len(generated_modules)} 个模块，补充占位代码")
            existing_ids = {m["module_id"] for m in generated_modules}
            for module_def in self.MODULES:
                if module_def["module_id"] not in existing_ids:
                    placeholder = self._generate_placeholder(module_def)
                    generated_modules.append({
                        "module_id": module_def["module_id"],
                        "module_name": module_def["module_name"],
                        "file_name": module_def["file_name"],
                        "function_name": module_def["function_name"],
                        "code": placeholder,
                        "error": "批量生成中未包含此模块"
                    })

        # 按模块ID排序
        generated_modules.sort(key=lambda x: x["module_id"])

        log(f"解析完成: 共 {len(generated_modules)} 个模块，成功解析 {success_count} 个")

        return generated_modules, success_count

    def _clean_module_code(self, code: str) -> str:
        """清理模块代码，移除markdown标记和非Python代码行"""
        import re

        # 移除markdown代码块标记
        code = re.sub(r'^```python\s*', '', code, flags=re.MULTILINE)
        code = re.sub(r'^```\s*$', '', code, flags=re.MULTILINE)

        # 移除中文模块名称行（如：数据加载模块、映射关系模块等）
        code = re.sub(r'^[数据加载映射关系预期文件模板公式生成填充]+模块\s*$', '', code, flags=re.MULTILINE)

        # 移除只包含中文的行（非注释）
        lines = code.split('\n')
        cleaned_lines = []
        in_code = False

        for line in lines:
            stripped = line.strip()

            # 跳过空行（在代码开始前）
            if not stripped and not in_code:
                continue

            # 检查是否是有效的Python代码开始
            if stripped.startswith('def ') or stripped.startswith('import ') or stripped.startswith('from ') or stripped.startswith('class '):
                in_code = True

            # 如果已经在代码中，保留所有行
            if in_code:
                cleaned_lines.append(line)
            # 如果还没开始代码，检查是否是注释或有效代码
            elif stripped.startswith('#'):
                # 保留注释
                cleaned_lines.append(line)
            elif stripped and not re.match(r'^[\u4e00-\u9fa5\s]+$', stripped):
                # 不是纯中文行，可能是有效代码
                if any(c in stripped for c in ['=', '(', ')', '[', ']', '{', '}', ':', '+', '-', '*', '/']):
                    in_code = True
                    cleaned_lines.append(line)

        return '\n'.join(cleaned_lines).strip()

    def _generate_module_with_stream(
        self,
        module_def: Dict[str, Any],
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]],
        previous_modules: List[Dict[str, Any]],
        stream_callback: callable = None
    ) -> str:
        """生成单个模块的代码（支持流式输出）"""

        module_id = module_def["module_id"]

        # 获取步骤编号
        step_number = int(module_id.split("_")[1])

        # 使用 prompt_generator 生成提示词
        prompt = self.prompt_generator.generate_modular_step_prompt(
            step_number=step_number,
            step_name=module_def["module_name"],
            rules_content=rules_content,
            source_structure=source_structure,
            expected_structure=expected_structure,
            previous_modules=previous_modules,
            manual_headers=manual_headers
        )

        # 检查是否支持流式生成
        has_stream = hasattr(self.ai_provider, 'generate_code_with_stream') and stream_callback

        if has_stream:
            # 使用流式生成
            full_code = ""

            def chunk_handler(chunk):
                nonlocal full_code
                full_code += chunk
                # 直接输出到控制台
                import sys
                sys.stdout.write(chunk)
                sys.stdout.flush()

            self.ai_provider.generate_code_with_stream(prompt, chunk_callback=chunk_handler)
            return full_code
        else:
            # 使用普通生成
            return self.ai_provider.generate_code(prompt)

    def _generate_placeholder(self, module_def: Dict[str, Any]) -> str:
        """生成占位代码"""
        return f'''
def {module_def["function_name"]}(*args, **kwargs):
    """
    {module_def["description"]}

    TODO: 此函数生成失败，需要手动实现
    """
    logger.warning("{module_def["function_name"]} 尚未实现")
    return None
'''

    def _merge_all_modules(
        self,
        modules: List[Dict[str, Any]],
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any]
    ) -> str:
        """合并所有模块生成完整代码"""

        # 文件头部
        header = '''"""
自动生成的数据处理脚本 - 6步模块化版本

模块结构:
1. 数据加载模块 (load_all_data) - 使用excel_parser.py加载数据
2. 映射关系模块 (generate_field_mapping)
3. 预期文件模板模块 (create_output_template)
4. 公式生成模块 (generate_formulas)
5. 数据填充模块 (fill_data) - 返回数据和透明化信息
6. Excel增强输出模块 (save_excel_with_details) - 添加批注、公式、过程列样式
"""

import pandas as pd
import numpy as np
import os
import json
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable, Tuple
import logging

# 导入Excel智能解析器
from excel_parser import IntelligentExcelParser

# 导入openpyxl用于Excel增强输出
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ============================================================
# 辅助函数：将ExcelRegion转换为DataFrame
# ============================================================

def convert_region_to_dataframe(region) -> pd.DataFrame:
    """将ExcelRegion转换为pandas DataFrame

    Args:
        region: ExcelRegion对象，包含head_data和data

    Returns:
        转换后的DataFrame，列名为中文表头名称
        即使没有数据行，也会返回带列名的空DataFrame
    """
    # 创建列字母到列名的反向映射
    col_letter_to_name = {v: k for k, v in region.head_data.items()}
    columns = list(region.head_data.keys())

    # 如果没有数据，返回带列名的空DataFrame
    if not region.data:
        return pd.DataFrame(columns=columns)

    # 转换数据
    converted_data = []
    for row in region.data:
        new_row = {}
        for col_letter, value in row.items():
            col_name = col_letter_to_name.get(col_letter, col_letter)
            new_row[col_name] = value
        converted_data.append(new_row)

    # 创建DataFrame，确保列顺序与表头一致
    return pd.DataFrame(converted_data, columns=columns)


# ============================================================
# 工号标准化函数
# ============================================================

def normalize_emp_code(emp_code) -> str:
    """标准化工号：转换为8位字符串，不足前面补0

    示例：
    - "123" -> "00000123"
    - "12345678" -> "12345678"
    - 12345 -> "00012345"
    """
    if pd.isna(emp_code) or emp_code == "":
        return ""
    # 转换为字符串并去除空格
    code_str = str(emp_code).strip()
    # 如果是纯数字，补齐到8位
    if code_str.isdigit():
        return code_str.zfill(8)
    return code_str

'''

        # 添加各模块代码
        modules_code = "\n# " + "=" * 60 + "\n"
        modules_code += "# 模块函数定义\n"
        modules_code += "# " + "=" * 60 + "\n\n"

        for mod in modules:
            modules_code += f"\n# --- 模块: {mod['module_name']} ---\n"
            modules_code += f"# 文件: {mod['file_name']}\n\n"
            modules_code += mod['code']
            modules_code += "\n\n"

        # 主函数 - 预期结构通过读取Excel获取
        output_filename = expected_structure.get('file_name', 'output.xlsx')

        # 提取预期文件的Sheet名称列表和列名信息（用于生成代码中的结构读取）
        expected_sheets_info = self._extract_expected_sheets_info(expected_structure)

        main_function = f'''
# {"=" * 60}
# 主函数
# {"=" * 60}

def load_expected_structure(expected_file_path: str) -> Dict[str, Any]:
    """从预期文件读取结构信息

    Args:
        expected_file_path: 预期文件路径

    Returns:
        预期文件的结构信息
    """
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(
        expected_file_path,
        manual_headers=manual_headers,
        active_sheet_only=True  # 只加载激活的sheet
    )

    structure = {{
        "file_name": os.path.basename(expected_file_path),
        "sheets": {{}}
    }}

    for sheet_data in results:
        sheet_name = sheet_data.sheet_name
        if sheet_data.regions:
            region = sheet_data.regions[0]
            structure["sheets"][sheet_name] = {{
                "headers": list(region.head_data.keys()),
                "head_data": region.head_data,
                "row_count": len(region.data) if region.data else 0
            }}

    return structure

def process_excel_files(input_folder: str, output_folder: str, expected_file: str = None,
                        salary_year: int = None, salary_month: int = None,
                        monthly_standard_hours: float = None) -> bool:
    """处理Excel文件的主函数

    执行6步处理流程:
    1. 加载所有源数据
    2. 生成字段映射关系
    3. 创建输出文件模板
    4. 生成计算公式
    5. 填充数据（返回数据和透明化信息）
    6. 保存Excel（添加批注、公式、过程列样式）

    Args:
        input_folder: 输入文件夹路径
        output_folder: 输出文件夹路径
        expected_file: 预期文件路径（可选，用于读取预期结构）
        salary_year: 薪资年份（可选）
        salary_month: 薪资月份（可选）
        monthly_standard_hours: 当月标准工时（可选，由调用方计算后传入）

    Returns:
        处理是否成功
    """
    try:
        logger.info("=" * 60)
        logger.info("开始执行6步数据处理流程")
        logger.info("=" * 60)

        # 薪资参数日志（如果有传入）
        if monthly_standard_hours is not None:
            logger.info(f"薪资年月: {{salary_year}}年{{salary_month}}月, 当月标准工时: {{monthly_standard_hours}}")

        # 确保输出目录存在
        os.makedirs(output_folder, exist_ok=True)

        # 规则内容（如果需要可以从外部传入）
        rules_content = """规则内容会在运行时传入"""

        # 预期结构 - 如果提供了预期文件则读取，否则使用默认结构
        if expected_file and os.path.exists(expected_file):
            logger.info(f"从预期文件读取结构: {{expected_file}}")
            expected_structure = load_expected_structure(expected_file)
        else:
            # 使用训练时提取的基本结构信息
            expected_structure = {{
                "file_name": "{output_filename}",
                "sheets": {expected_sheets_info}
            }}

        # 步骤1: 加载所有数据
        logger.info("")
        logger.info(">>> 步骤1/6: 加载源数据")
        logger.info("-" * 40)
        data_store = load_all_data(input_folder)
        if not data_store:
            logger.error("数据加载失败")
            return False
        logger.info(f"数据加载完成，共加载 {{len(data_store.get('files', {{}}))}} 个文件")

        # 步骤2: 生成映射关系
        logger.info("")
        logger.info(">>> 步骤2/6: 生成字段映射")
        logger.info("-" * 40)
        mapping = generate_field_mapping(data_store, rules_content)
        if not mapping:
            logger.warning("映射关系生成失败，使用默认映射")
            mapping = {{"direct_mapping": {{}}, "calculated_fields": []}}
        logger.info("映射关系生成完成")

        # 步骤3: 创建输出模板
        logger.info("")
        logger.info(">>> 步骤3/6: 创建输出模板")
        logger.info("-" * 40)
        template = create_output_template(mapping, expected_structure)
        if not template:
            logger.error("输出模板创建失败")
            return False
        logger.info(f"输出模板创建完成，共 {{len(template)}} 个Sheet")

        # 步骤4: 生成计算公式
        logger.info("")
        logger.info(">>> 步骤4/6: 生成计算公式")
        logger.info("-" * 40)
        formulas = generate_formulas(mapping, rules_content)
        if formulas is None:
            formulas = []
        logger.info(f"公式生成完成，共 {{len(formulas)}} 个计算任务")

        # 步骤5: 填充数据
        logger.info("")
        logger.info(">>> 步骤5/6: 填充数据")
        logger.info("-" * 40)
        fill_result = fill_data(data_store, template, mapping, formulas)

        # 处理fill_data的返回值（可能是元组或字典）
        if isinstance(fill_result, tuple) and len(fill_result) == 4:
            # 新格式：返回(result_data, column_sources, column_formulas, intermediate_columns)
            result, column_sources, column_formulas, intermediate_columns = fill_result
        elif isinstance(fill_result, dict):
            # 旧格式：只返回数据字典
            result = fill_result
            column_sources = {{}}
            column_formulas = {{}}
            intermediate_columns = []
        else:
            logger.error("数据填充返回格式错误")
            return False

        if not result:
            logger.error("数据填充失败")
            return False
        logger.info("数据填充完成")

        # 步骤6: 保存Excel（添加批注、公式、过程列样式）
        logger.info("")
        logger.info(">>> 步骤6/6: 保存Excel（带透明化信息）")
        logger.info("-" * 40)
        output_path = os.path.join(output_folder, "{output_filename}")

        # 直接调用save_excel_with_details（该函数在模块级别已定义）
        save_excel_with_details(result, output_path, column_sources, column_formulas, intermediate_columns)
        logger.info("保存完成（带批注和公式）")

        logger.info(f"结果已保存到: {{output_path}}")
        logger.info("")
        logger.info("=" * 60)
        logger.info("处理完成!")
        logger.info("=" * 60)

        return True

    except Exception as e:
        logger.error(f"处理失败: {{e}}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数入口 - 用于沙箱调用"""
    # 使用全局变量 input_folder, output_folder
    return process_excel_files(input_folder, output_folder)


# 注意：此脚本设计为通过沙箱调用 main() 或 process_excel_files(input_folder, output_folder)
# 不再使用 if __name__ == "__main__" 块，以避免与沙箱环境冲突
'''

        # 合并完整代码
        full_code = header + modules_code + main_function

        # 验证生成的代码语法
        full_code = self._validate_and_fix_merged_code(full_code)

        return full_code

    def _validate_and_fix_merged_code(self, code: str) -> str:
        """验证并修复合并后的代码语法

        Args:
            code: 合并后的代码

        Returns:
            验证/修复后的代码
        """
        import ast

        try:
            ast.parse(code)
            logger.info("合并代码语法验证通过")
            return code
        except SyntaxError as e:
            logger.warning(f"合并代码存在语法错误: {e}")

            # 尝试使用AI提供者的修复方法
            if hasattr(self.ai_provider, 'validate_and_fix_code_format'):
                logger.info("尝试自动修复代码语法...")
                fixed_code = self.ai_provider.validate_and_fix_code_format(code)

                try:
                    ast.parse(fixed_code)
                    logger.info("代码语法修复成功")
                    return fixed_code
                except SyntaxError as e2:
                    logger.error(f"代码修复后仍有语法错误: {e2}")

            return code

    def _extract_expected_sheets_info(self, expected_structure: Dict[str, Any]) -> str:
        """从预期结构中提取Sheet信息（只包含列名，不含数据）

        Args:
            expected_structure: 预期文件结构

        Returns:
            格式化的Sheet信息字符串，用于嵌入生成的代码
        """
        sheets_info = {}

        # 处理sheets结构
        if "sheets" in expected_structure:
            for sheet_name, sheet_data in expected_structure.get("sheets", {}).items():
                if isinstance(sheet_data, dict):
                    # 只提取列名/表头信息
                    headers = []
                    if "headers" in sheet_data:
                        headers = sheet_data["headers"]
                    elif "head_data" in sheet_data:
                        headers = list(sheet_data["head_data"].keys())

                    sheets_info[sheet_name] = {
                        "headers": headers
                    }

        # 格式化为Python字典字面量
        result = "{\n"
        for sheet_name, info in sheets_info.items():
            headers_str = json.dumps(info["headers"], ensure_ascii=False)
            result += f'                "{sheet_name}": {{"headers": {headers_str}}},\n'
        result += "            }"

        return result

    # ============ 兼容旧接口 ============

    def analyze_and_split_rules(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """分析规则并返回固定的5个任务（兼容旧接口）"""
        tasks = []
        for i, mod in enumerate(self.MODULES):
            tasks.append({
                "task_id": mod["module_id"],
                "task_name": mod["function_name"],
                "task_type": mod["module_name"],
                "description": mod["description"],
                "dependencies": [self.MODULES[j]["module_id"] for j in range(i)]
            })
        return tasks
