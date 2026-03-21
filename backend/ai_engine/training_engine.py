"""
训练引擎 - 管理AI训练过程
"""

import os
import json
import time
import logging
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

from .ai_provider import BaseAIProvider
from .prompt_generator import PromptGenerator
from .training_logger import TrainingLogger
from excel_parser import IntelligentExcelParser
from ..sandbox.code_sandbox import CodeSandbox
from ..utils.excel_comparator import compare_excel_files


class TrainingEngine:
    """训练引擎"""

    def __init__(self, ai_provider: BaseAIProvider = None, max_iterations: int = None,
                 stream_callback: Optional[callable] = None, use_modular: bool = None,
                 use_formula_mode: bool = None, db_persistence=None, user_id: int = None):
        """初始化训练引擎

        Args:
            ai_provider: AI提供者，如果为None则从配置自动创建
            max_iterations: 最大训练迭代次数，如果为None则从环境变量读取
            stream_callback: 流式回调函数，用于实时显示日志
            use_modular: 是否使用模块化代码生成，如果为None则根据规则复杂度自动判断
            use_formula_mode: 是否使用公式模式（生成Excel公式而非Python代码）
            db_persistence: TrainingPersistence 实例（可选，传入则写DB）
            user_id: 当前用户ID（可选，用于DB记录）
        """
        self.db_persistence = db_persistence
        self.user_id = user_id
        if ai_provider is None:
            # 从配置自动创建AI提供者
            from .ai_provider import AIProviderFactory
            self.ai_provider = AIProviderFactory.create_with_fallback()
        else:
            self.ai_provider = ai_provider

        # 设置最大迭代次数
        if max_iterations is None:
            # 从环境变量读取，默认为5（对于复杂规则需要更多迭代）
            self.max_iterations = int(os.getenv("MAX_TRAINING_ITERATIONS", "5"))
        else:
            self.max_iterations = max_iterations

        # 是否使用模块化生成
        if use_modular is None:
            # 从环境变量读取，默认为auto（自动判断）
            modular_setting = os.getenv("USE_MODULAR_GENERATION", "auto").lower()
            if modular_setting == "true":
                self.use_modular = True
            elif modular_setting == "false":
                self.use_modular = False
            else:
                self.use_modular = None  # auto模式
        else:
            self.use_modular = use_modular

        # 是否使用公式模式（默认启用）
        if use_formula_mode is None:
            # 从环境变量读取，默认为true
            formula_setting = os.getenv("USE_FORMULA_MODE", "true").lower()
            self.use_formula_mode = formula_setting == "true"
        else:
            self.use_formula_mode = use_formula_mode

        self.prompt_generator = PromptGenerator()
        self.excel_parser = IntelligentExcelParser()
        self.sandbox = CodeSandbox()
        self.logger = logging.getLogger(__name__)
        self.stream_callback = stream_callback
        self.training_logger = None

        # 训练成功阈值配置
        self.training_success_threshold = float(os.getenv("TRAINING_SUCCESS_THRESHOLD", "0.95"))
        self.training_perfect_threshold = float(os.getenv("TRAINING_PERFECT_THRESHOLD", "1.0"))

    # ==================== DB 持久化辅助方法 ====================

    def _db_record_iteration(self, iteration_num, code=None, accuracy=None,
                             prompt_text=None, ai_response=None,
                             execution_result=None, error_details=None,
                             duration_seconds=None, status="completed"):
        """将迭代结果写入DB（如果启用了持久化）"""
        if not self.db_persistence or not self._db_session_id:
            return
        try:
            self.db_persistence.record_iteration(
                session_id=self._db_session_id,
                iteration_num=iteration_num,
                prompt_text=prompt_text[:50000] if prompt_text else None,
                ai_response=ai_response[:50000] if ai_response else None,
                generated_code=code,
                accuracy=accuracy,
                execution_result=execution_result,
                error_details=error_details,
                duration_seconds=duration_seconds,
                status=status,
            )
        except Exception as e:
            if hasattr(self, 'training_logger') and self.training_logger:
                self.training_logger.log_warning(f"DB记录迭代失败: {e}")

    def _db_complete_session(self, best_score, total_iterations, best_code=None,
                             status="completed", error_message=None, tenant_id=None, mode="formula"):
        """完成DB训练会话并保存脚本"""
        if not self.db_persistence or not self._db_session_id:
            return
        try:
            script_id = None
            if best_code and best_score and best_score > 0 and tenant_id:
                script = self.db_persistence.save_script(
                    tenant_id=tenant_id,
                    name=f"training_{tenant_id}",
                    code=best_code,
                    mode=mode,
                    source_session_id=self._db_session_id,
                    accuracy=best_score,
                    created_by=self.user_id,
                )
                script_id = script.id
            self.db_persistence.complete_session(
                session_id=self._db_session_id,
                status=status,
                best_accuracy=best_score,
                total_iterations=total_iterations,
                final_script_id=script_id,
                error_message=error_message,
            )
        except Exception as e:
            if hasattr(self, 'training_logger') and self.training_logger:
                self.training_logger.log_warning(f"DB完成会话失败: {e}")

    def _is_natural_language_document(self, text: str) -> bool:
        """检查文本是否是自然语言文档（而非结构化规则）"""
        if not text:
            return False

        # 检查是否包含结构化规则关键词
        structured_keywords = [
            '预期输出:', 'Expected Output:', 'Output File:',
            '映射规则:', 'Mapping Rules:', '数据映射:',
            '计算规则:', 'Calculation Rules:', '计算公式:',
            '列名', '数据来源', '计算规则'
        ]

        for keyword in structured_keywords:
            if keyword in text:
                return False  # 包含结构化关键词，不是纯自然语言

        # 检查是否包含自然语言特征
        natural_language_indicators = [
            '需求文档', '需求说明', '需求背景', '版本信息', '变更日志',
            '文档说明', '名词解释', '前期准备', '导入表说明', '报表项',
            '业务逻辑', '处理流程', '数据流程', '系统需求'
        ]

        text_lower = text.lower()
        for indicator in natural_language_indicators:
            if indicator.lower() in text_lower:
                return True

        # 检查文本长度和结构
        lines = text.split('\n')
        if len(lines) > 10:
            # 如果文本较长且没有明显的结构化格式，可能是自然语言
            return True

        return False

    def _should_use_modular_generation(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any]
    ) -> bool:
        """判断是否应该使用模块化代码生成

        根据规则复杂度自动判断：
        1. 规则内容超过5000字符
        2. 源文件超过3个
        3. 预期输出有多个Sheet
        4. 规则中包含多个独立的计算逻辑
        """
        # 如果明确指定，直接返回
        if self.use_modular is not None:
            return self.use_modular

        # 规则1: 规则内容长度（改为3000字符触发）
        if len(rules_content) > 3000:
            self.logger.info("规则内容超过3000字符，使用模块化生成")
            return True

        # 规则2: 源文件数量
        source_file_count = len(source_structure.get('files', {}))
        if source_file_count > 3:
            self.logger.info(f"源文件超过3个({source_file_count})，使用模块化生成")
            return True

        # 规则3: 预期输出Sheet数量
        expected_sheet_count = len(expected_structure.get('sheets', {}))
        if expected_sheet_count > 2:
            self.logger.info(f"预期输出超过2个Sheet({expected_sheet_count})，使用模块化生成")
            return True

        # 规则4: 检查规则中的复杂度指标
        complexity_indicators = [
            '计算公式', '汇总', '合并', '关联', '匹配',
            'VLOOKUP', 'SUMIF', 'COUNTIF', '透视',
            '分组', '排序', '筛选', '去重'
        ]
        indicator_count = sum(1 for ind in complexity_indicators if ind in rules_content)
        if indicator_count >= 4:
            self.logger.info(f"规则包含{indicator_count}个复杂度指标，使用模块化生成")
            return True

        return False

    def _get_best_history_path(self, tenant_id: str) -> Path:
        """获取历史最佳分数文件路径"""
        return Path("tenants") / tenant_id / "best_history.json"

    def _get_historical_best_file(self, tenant_id: str) -> Path:
        """获取历史最佳分数文件路径（别名方法）"""
        return self._get_best_history_path(tenant_id)

    def _load_historical_best(self, tenant_id: str) -> Dict[str, Any]:
        """加载租户的历史最佳分数和代码

        Args:
            tenant_id: 租户ID

        Returns:
            包含历史最佳分数和代码的字典，如果不存在则返回空字典
        """
        history_path = self._get_best_history_path(tenant_id)
        if history_path.exists():
            try:
                with open(history_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                self.logger.warning(f"加载历史最佳分数失败: {e}")
        return {"best_score": 0.0, "best_code": None, "updated_at": None}

    def _save_historical_best(self, tenant_id: str, score: float, code: str) -> None:
        """保存租户的历史最佳分数和代码

        Args:
            tenant_id: 租户ID
            score: 最佳分数
            code: 最佳代码
        """
        history_path = self._get_best_history_path(tenant_id)
        history_path.parent.mkdir(parents=True, exist_ok=True)

        from datetime import datetime
        history_data = {
            "best_score": score,
            "best_code": code,
            "updated_at": datetime.now().isoformat()
        }

        try:
            with open(history_path, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, ensure_ascii=False, indent=2)
            self.logger.info(f"历史最佳分数已保存: {score:.2%}")
        except Exception as e:
            self.logger.warning(f"保存历史最佳分数失败: {e}")

    def _format_detailed_diff(self, field_diff_samples: Dict[str, Any], total_diff: int, matched_cells: int, total_cells: int) -> str:
        """格式化详细的差异信息用于AI修正

        按字段分类汇总差异，显示字段名、使用的公式和差异数量。

        Args:
            field_diff_samples: 按字段分类的差异样本 {字段名: {formula, count}}
            total_diff: 总差异数
            matched_cells: 匹配的单元格数
            total_cells: 总单元格数

        Returns:
            格式化的差异描述文本
        """
        if not field_diff_samples:
            return f"总体匹配率: {matched_cells}/{total_cells}, 差异 {total_diff} 处"

        lines = [
            "## 差异汇总",
            f"- 总体匹配率: {matched_cells}/{total_cells} ({matched_cells/total_cells*100:.1f}%)" if total_cells > 0 else "- 总体匹配率: N/A",
            f"- 总差异数: {total_diff} 处",
            f"- 涉及字段数: {len(field_diff_samples)} 个",
            "",
            "## 有差异的字段及其公式",
            ""
        ]

        # 按差异数量排序，优先显示差异最多的字段
        sorted_fields = sorted(field_diff_samples.items(), key=lambda x: x[1].get("count", 1), reverse=True)

        for field_name, sample in sorted_fields:
            formula = sample.get("formula", "")
            count = sample.get("count", 1)

            if formula:
                lines.append(f"- **{field_name}** ({count}处差异): `{formula}`")
            else:
                lines.append(f"- **{field_name}** ({count}处差异): [非公式列/直接赋值]")

        lines.append("")
        lines.append("## 修正要求")
        lines.append("请根据计算规则检查上述字段的Excel公式，确保：")
        lines.append("1. VLOOKUP列号计算正确（列号=目标列-起始列+1）")
        lines.append("2. 数据源sheet名称正确引用")
        lines.append("3. 条件判断逻辑符合规则要求")
        lines.append("4. 主键类型和数据类型是否一致，不一致vlookup需要转换")
        return "\n".join(lines)

    def train(
        self,
        source_files: List[str],
        expected_file: str,
        rule_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None,
        tenant_id: str = "default",
        salary_year: Optional[int] = None,
        salary_month: Optional[int] = None,
        monthly_standard_hours: Optional[float] = None,
        force_retrain: bool = False
    ) -> Dict[str, Any]:
        """训练AI生成数据处理脚本

        Args:
            source_files: 源数据文件列表
            expected_file: 预期结果文件
            rule_files: 规则文件列表
            manual_headers: 手动表头配置
            tenant_id: 租户ID
            salary_year: 薪资年份（可选）
            salary_month: 薪资月份（可选）
            monthly_standard_hours: 当月标准工时（可选，由调用方计算并传入）
            force_retrain: 是否强制重新训练（默认False）
                - False: 如果历史最佳分数=100%，直接使用历史最佳代码；如果<100%，重新训练
                - True: 清除所有历史训练数据和最佳代码，从头开始全新训练
        """
        # 保存薪资参数供后续使用
        self.salary_year = salary_year
        self.salary_month = salary_month
        self.monthly_standard_hours = monthly_standard_hours

        # 从规则文件名中提取关键词
        keyword = "training"
        if rule_files:
            # 使用第一个规则文件的文件名（去掉扩展名）作为关键词
            rule_filename = Path(rule_files[0]).stem
            # 清理关键词，移除特殊字符
            keyword = rule_filename.replace(" ", "_").replace("-", "_")

        # 初始化训练日志记录器
        self.training_logger = TrainingLogger(tenant_id, keyword=keyword)
        if self.stream_callback:
            self.training_logger.set_stream_callback(self.stream_callback)

        # 处理强制重新训练
        if force_retrain:
            self.training_logger.log_info("=" * 60)
            self.training_logger.log_info("强制重新训练模式：清除所有历史数据")
            self.training_logger.log_info("=" * 60)

            # 清除历史最佳分数和代码
            historical_best_file = self._get_historical_best_file(tenant_id)
            if historical_best_file.exists():
                historical_best_file.unlink()
                self.training_logger.log_info(f"已删除历史最佳分数文件: {historical_best_file}")

            # 清除历史训练日志（可选，保留日志便于追溯）
            # 注意：这里不删除training_logs目录，只清除历史最佳数据
            # 如果需要清除所有训练日志，可以取消下面的注释
            """
            training_logs_dir = Path(f"tenants/{tenant_id}/training_logs")
            if training_logs_dir.exists():
                import shutil
                shutil.rmtree(training_logs_dir)
                training_logs_dir.mkdir(parents=True, exist_ok=True)
                self.training_logger.log_info(f"已清除历史训练日志目录: {training_logs_dir}")
            """

            self.training_logger.log_info("历史数据清除完成，开始全新训练")

        # 开始训练记录
        self.training_logger.start_training(
            self.max_iterations, source_files, expected_file, rule_files
        )

        # DB 持久化：创建训练会话
        self._db_session_id = None
        if self.db_persistence:
            try:
                mode = "formula" if self.use_formula_mode else "auto"
                db_session = self.db_persistence.create_session(
                    tenant_id=tenant_id,
                    session_key=self.training_logger.session_dir.name if hasattr(self.training_logger, 'session_dir') else f"{tenant_id}_{keyword}",
                    mode=mode,
                    user_id=self.user_id,
                    config={
                        "max_iterations": self.max_iterations,
                        "salary_year": salary_year,
                        "salary_month": salary_month,
                        "monthly_standard_hours": monthly_standard_hours,
                        "force_retrain": force_retrain,
                    },
                )
                self._db_session_id = db_session.id
            except Exception as e:
                self.training_logger.log_warning(f"DB持久化创建会话失败: {e}")

        # 0. 验证输入文件
        self._validate_input_files(source_files, expected_file, rule_files)

        # 1. 解析源文件结构
        source_structure = self._analyze_source_structure(source_files, manual_headers)
        self.training_logger.log_info(f"解析源文件结构完成，共 {len(source_structure.get('files', {}))} 个文件")

        # 2. 解析预期文件结构
        expected_structure = self._analyze_expected_structure(expected_file, manual_headers)
        self.training_logger.log_info("解析预期文件结构完成")

        # 3. 提取规则内容
        rules_content = self.prompt_generator.extract_rules_from_files(rule_files)
        self.training_logger.log_info(f"提取规则内容完成，共 {len(rules_content)} 字符")

        # 3.05 使用AI生成数据校验规则（暂时禁用，因为生成的规则未用于提示词）
        validation_rules = {}
        # TODO: 未来如果需要将校验规则加入提示词，可以重新启用此功能
        """
        try:
            from .validation_rule_generator import generate_validation_rules_with_ai
            from .ai_provider import AIProviderFactory

            # 检查是否配置了专用的校验规则AI提供者
            validation_provider_type = os.getenv("VALIDATION_AI_PROVIDER")
            if validation_provider_type and validation_provider_type != os.getenv("AI_PROVIDER"):
                # 创建专用的校验规则AI提供者
                self.training_logger.log_info(f"使用 {validation_provider_type} AI分析规则文件，生成数据校验规则...")
                validation_ai_provider = AIProviderFactory.create_provider(validation_provider_type)
            else:
                # 使用默认AI提供者
                self.training_logger.log_info("使用AI分析规则文件，生成数据校验规则...")
                validation_ai_provider = self.ai_provider

            validation_rules = generate_validation_rules_with_ai(
                validation_ai_provider, rules_content, source_structure
            )
            if validation_rules.get("value_constraints"):
                self.training_logger.log_info(
                    f"生成 {len(validation_rules['value_constraints'])} 条数值校验规则"
                )
        except Exception as e:
            self.training_logger.log_warning(f"AI生成校验规则失败: {e}")
        """

        # 3.1 检查是否使用公式模式
        if self.use_formula_mode:
            self.training_logger.log_info("启用公式模式：生成Excel公式而非Python代码...")
            result = self._train_formula_mode(
                source_files, expected_file, rules_content,
                source_structure, expected_structure,
                manual_headers, tenant_id, force_retrain
            )
            # 添加校验规则到结果
            result["validation_rules"] = validation_rules
            return result

        # 3.5 判断是否使用模块化生成
        use_modular = self._should_use_modular_generation(
            rules_content, source_structure, expected_structure
        )

        if use_modular:
            self.training_logger.log_info("检测到复杂规则，启用模块化代码生成...")
            result = self._train_modular(
                source_files, expected_file, rules_content,
                source_structure, expected_structure,
                manual_headers, tenant_id
            )
            # 添加校验规则到结果
            result["validation_rules"] = validation_rules
            return result

        # 4. 使用AI分析文档并生成结构化规则（如果规则内容是自然语言描述）
        ai_rules = None
        if self._is_natural_language_document(rules_content):
            self.training_logger.log_info("检测到自然语言规则文档，使用AI生成结构化规则...")
            try:
                from .rule_generator import AIRuleGenerator
                rule_generator = AIRuleGenerator(self.ai_provider)
                ai_rules = rule_generator.generate_rules_from_document(
                    rules_content, source_structure, expected_structure, manual_headers
                )
                summary = ai_rules.get('summary', '无总结')
                self.training_logger.log_info(f"AI规则生成完成: {summary}")
            except Exception as e:
                self.training_logger.log_warning(f"AI规则生成失败，使用原始规则内容: {e}")
                ai_rules = None

        # 5. 生成初始代码
        best_code = None
        best_score = 0.0
        iteration_results = []

        for iteration in range(self.max_iterations):
            iteration_num = iteration + 1
            iteration_type = "training" if iteration == 0 else "correction"
            self.training_logger.start_iteration(iteration_num, iteration_type)

            try:
                # 生成或修正代码
                if iteration == 0:
                    if ai_rules:
                        # 使用AI生成的规则
                        prompt = self.prompt_generator.generate_training_prompt_with_ai_rules(
                            source_structure, expected_structure, ai_rules, manual_headers
                        )
                        self.training_logger.log_info("使用AI生成的规则创建提示词")
                    else:
                        # 使用原始规则内容
                        prompt = self.prompt_generator.generate_training_prompt(
                            source_structure, expected_structure, rules_content, manual_headers
                        )
                        self.training_logger.log_info("使用原始规则内容创建提示词")

                    # 记录提示词生成
                    prompt_preview = f"前500字符: {prompt[:500]}... 后500字符: {prompt[-500:] if len(prompt) > 500 else prompt}"
                    self.training_logger.log_prompt_generation(
                        "training", len(prompt), prompt_preview
                    )

                    # 记录完整提示词到文件
                    self.training_logger.log_full_prompt(prompt, "generate")

                    # 生成代码（支持流式输出）
                    code = self._generate_code_with_logging(prompt, iteration_num, tenant_id)
                else:
                    # 使用上一次的对比结果生成修正提示
                    last_result = iteration_results[-1]
                    if ai_rules:
                        prompt = self.prompt_generator.generate_correction_prompt_with_ai_rules(
                            original_code=last_result["code"],
                            error_description=last_result["error_description"],
                            comparison_result=last_result["comparison_result"],
                            source_structure=source_structure,
                            expected_structure=expected_structure,
                            ai_rules=ai_rules,
                            manual_headers=manual_headers
                        )
                    else:
                        prompt = self.prompt_generator.generate_correction_prompt(
                            original_code=last_result["code"],
                            error_description=last_result["error_description"],
                            comparison_result=last_result["comparison_result"],
                            source_structure=source_structure,
                            expected_structure=expected_structure,
                            rules_content=rules_content,
                            manual_headers=manual_headers
                        )

                    # 记录修正提示词生成
                    self.training_logger.log_prompt_generation("correction", len(prompt))

                    # 记录完整提示词到文件
                    self.training_logger.log_full_prompt(prompt, "correct")

                    # 生成代码（支持流式输出）
                    code = self._generate_code_with_logging(prompt, iteration_num, tenant_id)

                # 保存生成的脚本（无论成功与否）
                script_path = self._save_generated_script(code, iteration_num, tenant_id)
                if script_path:
                    self.training_logger.log_info(f"脚本已保存到: {script_path}")

                # 执行代码并验证结果
                self.training_logger.log_execution_start()
                execution_result = self._execute_and_validate(
                    code, source_files, expected_file, manual_headers,
                    tenant_id=tenant_id, iteration_num=iteration_num
                )

                # 记录执行结果
                self.training_logger.log_execution_result(
                    execution_result["success"],
                    execution_result.get("execution_time", 0),
                    execution_result.get("error"),
                    execution_result.get("output_file")
                )

                # 计算匹配分数
                score = self._calculate_match_score(execution_result)

                # 记录比较结果
                comparison = execution_result.get("comparison", "")
                self.training_logger.log_comparison_result(comparison, score)

                # 保存迭代结果
                iteration_result = {
                    "iteration": iteration_num,
                    "code": code,
                    "script_path": script_path,
                    "execution_result": execution_result,
                    "score": score,
                    "error_description": execution_result.get("error", ""),
                    "comparison_result": execution_result.get("comparison", ""),
                    "raw_response": getattr(self.ai_provider, 'last_raw_response', "")
                }
                iteration_results.append(iteration_result)
                self._db_record_iteration(
                    iteration_num=iteration + 1, code=code, accuracy=score,
                    ai_response=iteration_result.get("raw_response"),
                    execution_result={"score": score, "error": iteration_result.get("error_description", "")},
                )

                # 更新最佳代码
                is_best = score > best_score
                if is_best:
                    best_score = score
                    best_code = code

                # 记录迭代完成
                self.training_logger.log_iteration_complete(iteration_num, score, is_best)

                # 如果达到完美匹配阈值，提前结束
                if score >= self.training_perfect_threshold:
                    self.training_logger.log_info(f"达到{self.training_perfect_threshold*100:.0f}%匹配，提前结束训练")
                    break

            except Exception as e:
                self.training_logger.log_warning(
                    f"第 {iteration_num} 次迭代失败: {e}，保留已有最佳结果(score={best_score:.2%})"
                )
                self.logger.error(f"迭代 {iteration_num} 异常: {e}", exc_info=True)
                continue

        # 5. 返回训练结果
        result = {
            "tenant_id": tenant_id,
            "best_code": best_code,
            "best_score": best_score,
            "source_structure": source_structure,
            "expected_structure": expected_structure,
            "manual_headers": manual_headers,
            "rules_content": rules_content,
            "validation_rules": validation_rules,
            "iteration_results": iteration_results,
            "total_iterations": len(iteration_results),
            "success": best_score >= self.training_success_threshold,
            "log_file": self.training_logger.get_log_file_path(),
            "training_summary": self.training_logger.get_training_summary_path()
        }

        # 记录训练完成
        self.training_logger.log_training_complete(
            best_score, len(iteration_results), result["success"], len(best_code) if best_code else 0
        )
        self._db_complete_session(best_score, len(iteration_results), best_code, tenant_id=tenant_id, mode="simple")

        return result

    def _train_formula_mode(
        self,
        source_files: List[str],
        expected_file: str,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None,
        tenant_id: str = "default",
        force_retrain: bool = False
    ) -> Dict[str, Any]:
        """使用公式模式训练 - 生成使用Excel公式的Python代码

        思路：
        - 把源数据放入不同sheet
        - 基础列直接填充数据
        - 计算列使用Excel公式（VLOOKUP/IF等）

        Args:
            source_files: 源文件列表
            expected_file: 预期结果文件
            rules_content: 规则内容
            source_structure: 源文件结构
            expected_structure: 预期文件结构
            manual_headers: 手动表头配置
            tenant_id: 租户ID
            force_retrain: 是否强制重新训练

        Returns:
            训练结果字典
        """
        from .formula_code_generator import FormulaCodeGenerator

        self.training_logger.log_info("=" * 60)
        self.training_logger.log_info("公式模式训练开始")
        self.training_logger.log_info("说明: 基础列填充数据，计算列使用Excel公式")
        self.training_logger.log_info("=" * 60)

        # 加载历史最佳分数和代码
        historical_best = self._load_historical_best(tenant_id)
        historical_best_score = historical_best.get("best_score", 0.0)
        historical_best_code = historical_best.get("best_code")
        if historical_best_score > 0:
            self.training_logger.log_info(f"历史最佳分数: {historical_best_score:.2%}")

        # 如果不是强制重新训练，且历史最佳分数已经达到完美匹配阈值，直接返回，不需要再训练
        if not force_retrain and historical_best_score >= self.training_perfect_threshold and historical_best_code:
            self.training_logger.log_info(f"历史最佳分数已达到{self.training_perfect_threshold*100:.0f}%，跳过训练，直接使用历史最佳代码")
            self.training_logger.log_training_complete(
                best_score=historical_best_score,
                total_iterations=0,
                success=True,
                best_code_length=len(historical_best_code) if historical_best_code else 0
            )
            return {
                "success": True,
                "best_score": historical_best_score,
                "current_score": historical_best_score,
                "historical_best_score": historical_best_score,
                "total_iterations": 0,
                "output_path": "",
                "result_excel": "",
                "comparison_excel": "",
                "best_code": historical_best_code,
                "mode": "formula",
                "iteration_results": [],
                "source_structure": source_structure,
                "expected_structure": expected_structure,
                "validation_rules": {}
            }

        # 创建公式代码生成器
        formula_generator = FormulaCodeGenerator(
            ai_provider=self.ai_provider,
            training_logger=self.training_logger
        )

        # 获取输入输出文件夹
        input_folder = str(Path(source_files[0]).parent) if source_files else ""
        output_folder = str(Path(expected_file).parent) if expected_file else input_folder

        best_code = None
        best_score = 0.0
        best_output_path = ""
        best_saved_output_path = ""
        best_saved_comparison_path = ""
        iteration_results = []
        source_structure_desc = ""  # 用于修正时传递

        for iteration in range(self.max_iterations):
            iteration_num = iteration + 1
            self.training_logger.start_iteration(iteration_num, "formula")
            self.training_logger.log_info(f"--- 公式模式迭代 {iteration_num}/{self.max_iterations} ---")

            try:
                if iteration == 0:
                    # 首次生成代码
                    code, ai_response = formula_generator.generate_code(
                        input_folder=input_folder,
                        rules_content=rules_content,
                        expected_structure=expected_structure,
                        manual_headers=manual_headers,
                        stream_callback=self.stream_callback
                    )
                    # 保存源数据结构描述用于后续修正
                    source_structure_desc = formula_generator.formula_builder.get_source_structure_for_prompt()
                else:
                    # 修正代码
                    # 优先使用详细差异信息，如果没有则使用简单文本
                    comparison_result = ""
                    if iteration_results:
                        last_result = iteration_results[-1]
                        comparison_result = last_result.get("detailed_diff", "") or last_result.get("comparison", "")

                    # 确定用于修正的原始代码：优先使用best_code，否则使用上一次迭代的代码
                    original_code_for_correction = best_code
                    if original_code_for_correction is None and iteration_results:
                        # 从最近的迭代结果中获取代码
                        original_code_for_correction = iteration_results[-1].get("code")

                    if original_code_for_correction is None:
                        # 如果仍然没有可用代码，重新生成
                        self.training_logger.log_warning("没有可用的原始代码，重新生成...")
                        code, ai_response = formula_generator.generate_code(
                            input_folder=input_folder,
                            rules_content=rules_content,
                            expected_structure=expected_structure,
                            manual_headers=manual_headers,
                            stream_callback=self.stream_callback
                        )
                    else:
                        code = formula_generator.generate_correction_code(
                            original_code=original_code_for_correction,
                            comparison_result=comparison_result,
                            rules_content=rules_content,
                            source_structure=source_structure_desc,
                            stream_callback=self.stream_callback
                        )

                if not code:
                    self.training_logger.log_warning("代码生成失败，跳过此迭代")
                    continue

                # 注意：log_generated_code已在formula_code_generator内部调用，这里不再重复

                # 在执行代码之前，清理output_folder中的旧输出文件，避免干扰
                output_dir_for_clean = Path(output_folder)
                old_output_file = output_dir_for_clean / "薪资汇总表.xlsx"
                if old_output_file.exists():
                    old_output_file.unlink()
                    self.training_logger.log_info(f"清理旧输出文件: {old_output_file.name}")

                # 在沙箱中执行代码
                self.training_logger.log_info("执行生成的代码...")
                start_time = time.time()

                execution_env = {
                    "input_folder": input_folder,
                    "output_folder": output_folder,
                    "source_files": [Path(f).name for f in source_files],
                    "manual_headers": manual_headers or {}
                }

                # 添加薪资参数（如果有）
                if hasattr(self, 'salary_year') and self.salary_year is not None:
                    execution_env["salary_year"] = self.salary_year
                if hasattr(self, 'salary_month') and self.salary_month is not None:
                    execution_env["salary_month"] = self.salary_month
                if hasattr(self, 'monthly_standard_hours') and self.monthly_standard_hours is not None:
                    execution_env["monthly_standard_hours"] = self.monthly_standard_hours
                    self.training_logger.log_info(f"薪资参数 - 年: {self.salary_year}, 月: {self.salary_month}, 标准工时: {self.monthly_standard_hours}")

                execution_result = self.sandbox.execute_script(code, execution_env)
                execution_time = time.time() - start_time

                # 记录沙箱执行结果
                self.training_logger.log_info(f"沙箱执行结果: success={execution_result['success']}")
                if execution_result.get('output'):
                    for line in execution_result['output'].split('\n'):
                        if line.strip():
                            self.training_logger.log_info(f"[沙箱] {line}")
                if execution_result.get('error'):
                    self.training_logger.log_error(f"沙箱错误: {execution_result['error'][:500]}...")

                if not execution_result["success"]:
                    self.training_logger.log_execution_result(
                        False, execution_time, error=execution_result.get("error", "执行失败")
                    )
                    # 记录失败的迭代
                    iteration_results.append({
                        "iteration": iteration_num,
                        "score": 0.0,
                        "error": execution_result.get("error", "执行失败"),
                        "code": code
                    })
                    continue

                # 查找生成的输出文件
                output_dir = Path(output_folder)

                # 优先查找固定名称的输出文件（薪资汇总表.xlsx）
                expected_output_file = output_dir / "薪资汇总表.xlsx"
                if expected_output_file.exists():
                    output_path = str(expected_output_file)
                    self.training_logger.log_info(f"找到预期输出文件: {output_path}")
                else:
                    # 如果没有固定名称的文件，查找所有xlsx文件，但排除差异对比文件
                    output_files = list(output_dir.glob("*.xlsx"))
                    # 过滤掉差异对比文件和comparison文件
                    output_files = [
                        f for f in output_files
                        if "差异对比" not in f.name
                        and "comparison" not in f.name.lower()
                        and not f.name.startswith("rex_")  # 排除日志保存的文件
                    ]
                    self.training_logger.log_info(f"在输出目录 {output_dir} 中找到 {len(output_files)} 个有效Excel文件")

                    if not output_files:
                        self.training_logger.log_execution_result(
                            False, execution_time, error="未找到输出Excel文件"
                        )
                        iteration_results.append({
                            "iteration": iteration_num,
                            "score": 0.0,
                            "error": "未找到输出Excel文件",
                            "code": code
                        })
                        continue

                    output_path = str(output_files[0])
                self.training_logger.log_execution_result(True, execution_time, output_file=output_path)

                # 使用training_logger保存输出Excel到训练日志目录
                saved_output_path = self.training_logger.save_output_excel(output_path, mode_type="formula")

                # 比较结果 - 直接保存到training_logs目录
                comparison_output_file = str(self.training_logger.log_dir / f"差异对比_{iteration_num}.xlsx")
                comparison_result = compare_excel_files(
                    result_file=output_path,
                    expected_file=expected_file,
                    output_file=comparison_output_file
                )

                # 计算匹配分数（基于单元格匹配数量）
                total_cells = comparison_result.get("total_cells", 0)
                matched_cells = comparison_result.get("matched_cells", 0)
                total_diff = comparison_result.get("total_differences", 0)

                # 使用单元格匹配率作为分数
                if total_cells > 0:
                    score = matched_cells / total_cells
                else:
                    # 如果无法获取单元格数，使用旧的计算方式
                    score = 1.0 if total_diff == 0 else max(0, 1.0 - total_diff / 100.0)

                comparison_text = f"差异对比完成: 匹配 {matched_cells}/{total_cells} 个单元格, 匹配率 {score:.2%}, 差异 {total_diff} 处"
                self.training_logger.log_info(comparison_text)

                # 生成详细的差异描述用于AI修正
                field_diff_samples = comparison_result.get("field_diff_samples", {})
                detailed_diff_text = self._format_detailed_diff(field_diff_samples, total_diff, matched_cells, total_cells)

                # 重命名差异对比文件为统一格式
                saved_comparison_path = ""
                if Path(comparison_output_file).exists():
                    saved_comparison_path = self.training_logger.save_comparison_excel(comparison_output_file, mode_type="formula")
                    # 删除临时的差异对比文件
                    Path(comparison_output_file).unlink(missing_ok=True)

                # 记录迭代结果
                iteration_result = {
                    "iteration": iteration_num,
                    "score": score,
                    "output_path": output_path,
                    "saved_output_path": saved_output_path,
                    "saved_comparison_path": saved_comparison_path,
                    "comparison": comparison_text,
                    "detailed_diff": detailed_diff_text,  # 详细差异信息
                    "field_diff_samples": field_diff_samples,  # 按字段分类的差异
                    "code": code
                }
                iteration_results.append(iteration_result)
                self._db_record_iteration(
                    iteration_num=iteration_num, code=code, accuracy=score,
                    execution_result={"score": score, "output_path": output_path},
                )

                # 更新最佳结果
                if score > best_score:
                    best_score = score
                    best_code = code
                    best_output_path = output_path
                    best_saved_output_path = saved_output_path
                    best_saved_comparison_path = saved_comparison_path
                    self.training_logger.log_iteration_complete(iteration_num, score, is_best=True)
                else:
                    self.training_logger.log_iteration_complete(iteration_num, score, is_best=False)

                # 如果分数达到完美匹配阈值，提前结束
                if score >= self.training_perfect_threshold:
                    self.training_logger.log_info(f"达到{self.training_perfect_threshold*100:.0f}%匹配，提前结束训练")
                    break

            except Exception as e:
                self.training_logger.log_error(f"迭代 {iteration_num} 发生错误: {e}", e)
                import traceback
                traceback.print_exc()

        # 更新历史最佳分数（如果本次分数更高）
        if best_score > historical_best_score and best_code:
            self._save_historical_best(tenant_id, best_score, best_code)
            historical_best_score = best_score
            historical_best_code = best_code
            self.training_logger.log_info(f"新的历史最佳分数: {best_score:.2%}")

        # 使用历史最佳代码（如果历史最佳分数更高）
        final_code = best_code
        final_score = best_score
        if historical_best_score > best_score and historical_best_code:
            final_code = historical_best_code
            final_score = historical_best_score
            self.training_logger.log_info(f"使用历史最佳代码（分数: {historical_best_score:.2%}）")

        # 构建返回结果
        result = {
            "success": final_score >= self.training_success_threshold,
            "best_score": final_score,
            "current_score": best_score,
            "historical_best_score": historical_best_score,
            "total_iterations": len(iteration_results),
            "output_path": best_output_path,
            "result_excel": best_saved_output_path,
            "comparison_excel": best_saved_comparison_path,
            "best_code": final_code,
            "mode": "formula",
            "iteration_results": iteration_results,
            "source_structure": source_structure,
            "expected_structure": expected_structure,
            "manual_headers": manual_headers,
            "rules_content": rules_content,
            "tenant_id": tenant_id,
            "log_file": self.training_logger.get_log_file_path(),
            "training_summary": self.training_logger.get_training_summary_path()
        }

        self.training_logger.log_info("=" * 60)
        self.training_logger.log_info(f"公式模式训练完成")
        self.training_logger.log_info(f"本次最佳分数: {best_score:.2%}")
        self.training_logger.log_info(f"历史最佳分数: {historical_best_score:.2%}")
        self.training_logger.log_info(f"采用分数: {final_score:.2%}")
        self.training_logger.log_info("=" * 60)

        # 记录训练完成
        self.training_logger.log_training_complete(
            final_score, len(iteration_results), result["success"],
            len(final_code) if final_code else 0
        )
        self._db_complete_session(final_score, len(iteration_results), final_code, tenant_id=tenant_id, mode="formula")

        return result

    def _train_modular(
        self,
        source_files: List[str],
        expected_file: str,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None,
        tenant_id: str = "default"
    ) -> Dict[str, Any]:
        """使用模块化方式训练生成代码

        将复杂规则拆分成多个子任务，分别生成代码后合并。

        Args:
            source_files: 源文件列表
            expected_file: 预期结果文件
            rules_content: 规则内容
            source_structure: 源文件结构
            expected_structure: 预期文件结构
            manual_headers: 手动表头配置
            tenant_id: 租户ID

        Returns:
            训练结果字典
        """
        from .modular_code_generator import ModularCodeGenerator

        self.training_logger.log_info("=== 启动模块化代码生成 ===")

        # 创建模块化代码生成器，传入training_logger以记录提示词和AI响应
        modular_generator = ModularCodeGenerator(self.ai_provider, self.training_logger)

        best_code = None
        best_score = 0.0
        iteration_results = []

        for iteration in range(self.max_iterations):
            iteration_num = iteration + 1
            iteration_type = "modular_training" if iteration == 0 else "modular_correction"
            self.training_logger.start_iteration(iteration_num, iteration_type)

            try:
                if iteration == 0:
                    # 第一次迭代：完整的模块化生成
                    self.training_logger.log_info("开始模块化代码生成...")

                    code, tasks = modular_generator.generate_modular_code(
                        rules_content,
                        source_structure,
                        expected_structure,
                        manual_headers,
                        stream_callback=self.stream_callback,
                        salary_year=self.salary_year,
                        salary_month=self.salary_month,
                        monthly_standard_hours=self.monthly_standard_hours
                    )

                    self.training_logger.log_info(f"模块化生成完成，共 {len(tasks)} 个模块，代码长度: {len(code)}")
                else:
                    # 后续迭代：基于错误进行修正
                    last_result = iteration_results[-1]

                    # 生成修正提示
                    correction_prompt = self._build_modular_correction_prompt(
                        last_result["code"],
                        last_result["error_description"],
                        last_result["comparison_result"],
                        source_structure,
                        expected_structure,
                        rules_content
                    )

                    # 记录完整提示词
                    self.training_logger.log_full_prompt(correction_prompt, "correct")

                    self.training_logger.log_info("生成修正代码...")
                    code = self._generate_code_with_logging(correction_prompt, iteration_num, tenant_id)

                # 保存生成的脚本
                script_path = self._save_generated_script(code, iteration_num, tenant_id)
                if script_path:
                    self.training_logger.log_info(f"脚本已保存到: {script_path}")

                # 执行代码并验证结果
                self.training_logger.log_execution_start()
                execution_result = self._execute_and_validate(
                    code, source_files, expected_file, manual_headers,
                    tenant_id=tenant_id, iteration_num=iteration_num
                )

                # 记录执行结果
                self.training_logger.log_execution_result(
                    execution_result["success"],
                    execution_result.get("execution_time", 0),
                    execution_result.get("error"),
                    execution_result.get("output_file")
                )

                # 计算匹配分数
                score = self._calculate_match_score(execution_result)

                # 记录比较结果
                comparison = execution_result.get("comparison", "")
                self.training_logger.log_comparison_result(comparison, score)

                # 保存迭代结果
                iteration_result = {
                    "iteration": iteration_num,
                    "code": code,
                    "script_path": script_path,
                    "execution_result": execution_result,
                    "score": score,
                    "error_description": execution_result.get("error", ""),
                    "comparison_result": execution_result.get("comparison", ""),
                    "generation_mode": "modular"
                }
                iteration_results.append(iteration_result)
                self._db_record_iteration(
                    iteration_num=iteration_num, code=code, accuracy=score,
                    execution_result={"score": score, "error": iteration_result.get("error_description", "")},
                )

                # 更新最佳代码
                is_best = score > best_score
                if is_best:
                    best_score = score
                    best_code = code

                # 记录迭代完成
                self.training_logger.log_iteration_complete(iteration_num, score, is_best)

                # 如果达到完美匹配阈值，提前结束
                if score >= self.training_perfect_threshold:
                    self.training_logger.log_info(f"达到{self.training_perfect_threshold*100:.0f}%匹配，提前结束训练")
                    break

            except Exception as e:
                self.training_logger.log_error(f"模块化生成迭代 {iteration_num} 失败: {e}")
                # 获取已生成的代码（如果有）
                failed_code = code if 'code' in dir() and code else ""
                iteration_results.append({
                    "iteration": iteration_num,
                    "code": failed_code,
                    "error_description": str(e),
                    "comparison_result": f"迭代失败: {str(e)}",  # 必须包含此字段，否则修正提示词无法生成
                    "score": 0.0,
                    "generation_mode": "modular"
                })

        # 从最佳迭代结果中获取输出文件路径
        best_result_excel = ""
        best_comparison_excel = ""
        for iter_result in iteration_results:
            if iter_result.get("score") == best_score:
                execution_result = iter_result.get("execution_result", {})
                best_result_excel = execution_result.get("training_output_file", "")
                best_comparison_excel = execution_result.get("comparison_excel_file", "")
                break

        # 返回训练结果
        result = {
            "tenant_id": tenant_id,
            "best_code": best_code,
            "best_score": best_score,
            "result_excel": best_result_excel,
            "comparison_excel": best_comparison_excel,
            "source_structure": source_structure,
            "expected_structure": expected_structure,
            "manual_headers": manual_headers,
            "rules_content": rules_content,
            "iteration_results": iteration_results,
            "total_iterations": len(iteration_results),
            "success": best_score >= self.training_success_threshold,
            "generation_mode": "modular",
            "log_file": self.training_logger.get_log_file_path(),
            "training_summary": self.training_logger.get_training_summary_path()
        }

        # 记录训练完成
        self.training_logger.log_training_complete(
            best_score, len(iteration_results), result["success"], len(best_code) if best_code else 0
        )
        self._db_complete_session(best_score, len(iteration_results), best_code, tenant_id=tenant_id, mode="modular")

        return result

    def _build_modular_correction_prompt(
        self,
        original_code: str,
        error_description: str,
        comparison_result: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        rules_content: str
    ) -> str:
        """构建模块化修正提示词

        使用与第一次训练相同的骨架结构，只是在前面添加错误信息和上次代码。
        这样AI可以参考完整的骨架来修正代码，而不是生成完全不同的结构。
        """
        # 首先生成与第一次训练相同的基础提示词（包含完整骨架）
        base_prompt = self.prompt_generator.generate_batch_modular_prompt(
            rules_content=rules_content,
            source_structure=source_structure,
            expected_structure=expected_structure,
            manual_headers=None,
            modules=None,
            salary_year=getattr(self, 'salary_year', None),
            salary_month=getattr(self, 'salary_month', None),
            monthly_standard_hours=getattr(self, 'monthly_standard_hours', None)
        )

        # 在基础提示词前添加上次的代码和错误信息
        correction_header = f"""## 【修正模式】这是第二次或后续迭代，请修正上次代码中的错误

## 上次生成的代码（有错误，需要修正）
```python
{original_code[:15000]}
```

## 上次执行的错误
{error_description[:3000]}

## 上次对比结果（前10行差异）
{comparison_result[:5000]}

## 【修正要求】
1. 参考下方的"代码骨架"结构，修正上述代码中的错误
2. 保持6个模块的结构不变
3. 特别注意：上面的错误信息告诉你具体哪里出错了
4. 必须包含完整的main()函数

---
以下是正确的代码骨架和规则要求，请参考修正：

"""

        return correction_header + base_prompt

    def _generate_code_with_logging(self, prompt: str, iteration: int = 0, tenant_id: str = "") -> str:
        """生成代码并记录完整日志（支持流式输出）

        Args:
            prompt: 提示词
            iteration: 迭代次数（用于日志标记）
            tenant_id: 租户ID（用于日志标记）

        Returns:
            生成的代码
        """
        self.training_logger.log_info(f"开始调用AI生成代码 (迭代: {iteration}, 租户: {tenant_id})...")

        # 调试信息
        has_stream_method = hasattr(self.ai_provider, 'generate_code_with_stream')
        has_callback = self.stream_callback is not None
        self.training_logger.log_info(f"流式生成检查: 支持流式={has_stream_method}, 有回调={has_callback}")

        # 检查AI提供者是否支持流式生成
        if has_stream_method and has_callback:
            # 使用流式生成
            self.training_logger.log_info("使用流式API生成代码...")

            def chunk_handler(chunk):
                # 实时输出到日志
                self.training_logger.log_streaming_chunk(chunk)

            code = self.ai_provider.generate_code_with_stream(prompt, chunk_callback=chunk_handler)
        else:
            # 使用普通生成
            self.training_logger.log_info("使用普通API生成代码（非流式）...")
            code = self.ai_provider.generate_code(prompt)

        # 记录完整的AI响应（原始响应）
        raw_response = getattr(self.ai_provider, 'last_raw_response', '')
        if raw_response:
            self.training_logger.log_full_ai_response(raw_response, "generate")

        # 记录生成的代码
        self.training_logger.log_generated_code(code, "modular")
        self.training_logger.log_code_generated(len(code), code[:200] if code else "")

        return code

    def _save_generated_script(self, code: str, iteration: int, tenant_id: str, score: float = 0.0) -> str:
        """保存生成的脚本到文件

        Args:
            code: 生成的Python代码
            iteration: 迭代次数
            tenant_id: 租户ID
            score: 匹配分数

        Returns:
            保存的文件路径
        """
        from pathlib import Path
        import hashlib

        # 创建脚本保存目录
        scripts_dir = Path("tenants") / tenant_id / "scripts"
        scripts_dir.mkdir(parents=True, exist_ok=True)

        # 生成文件名：迭代次数_分数_哈希值前8位.py
        code_hash = hashlib.md5(code.encode('utf-8')).hexdigest()[:8]
        score_str = f"{score:.2f}" if score > 0 else "0.00"
        filename = f"script_{iteration:02d}_{score_str}_{code_hash}.py"
        file_path = scripts_dir / filename

        # 保存脚本
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(code)

            self.training_logger.log_info(f"保存生成的脚本: {filename} (长度: {len(code)} 字符)")
            return str(file_path)
        except Exception as e:
            self.training_logger.log_error(f"保存脚本失败: {e}")
            return ""

    def _save_api_response(self, raw_response: str, iteration: int, tenant_id: str) -> str:
        """保存完整的API响应到文件

        Args:
            raw_response: 原始的API响应
            iteration: 迭代次数
            tenant_id: 租户ID

        Returns:
            保存的文件路径
        """
        from pathlib import Path
        import hashlib

        # 创建响应保存目录
        responses_dir = Path("tenants") / tenant_id / "api_responses"
        responses_dir.mkdir(parents=True, exist_ok=True)

        # 生成文件名：响应_迭代次数_哈希值前8位.txt
        response_hash = hashlib.md5(raw_response.encode('utf-8')).hexdigest()[:8]
        filename = f"response_{iteration:02d}_{response_hash}.txt"
        file_path = responses_dir / filename

        # 保存响应
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(raw_response)

            self.training_logger.log_info(f"保存API响应: {filename} (长度: {len(raw_response)} 字符)")
            return str(file_path)
        except Exception as e:
            self.training_logger.log_error(f"保存API响应失败: {e}")
            return ""

    def _validate_input_files(
        self, source_files: List[str], expected_file: str, rule_files: List[str]
    ) -> None:
        """验证输入文件"""
        from pathlib import Path

        # 检查源文件
        for i, file_path in enumerate(source_files, 1):
            if not Path(file_path).exists():
                raise FileNotFoundError(
                    f"源文件 {i} 不存在: {file_path}\n"
                    f"请确保文件存在且路径正确。"
                )

        # 检查预期文件
        if not Path(expected_file).exists():
            raise FileNotFoundError(
                f"预期结果文件不存在: {expected_file}\n"
                f"请确保文件存在且路径正确。"
            )

        # 检查规则文件
        for i, file_path in enumerate(rule_files, 1):
            if not Path(file_path).exists():
                raise FileNotFoundError(
                    f"规则文件 {i} 不存在: {file_path}\n"
                    f"请确保文件存在且路径正确。"
                )

        self.training_logger.log_info(f"文件验证通过: {len(source_files)}个源文件, 1个预期文件, {len(rule_files)}个规则文件")

    def _analyze_source_structure(
        self, source_files: List[str], manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """分析源文件结构"""
        structure = {
            "files": {},
            "total_sheets": 0,
            "total_regions": 0
        }

        for file_path in source_files:
            try:
                # 检查文件是否存在
                if not Path(file_path).exists():
                    raise FileNotFoundError(f"源文件不存在: {file_path}")

                file_name = Path(file_path).name
                parsed_data = self.excel_parser.parse_excel_file(
                      file_path,
                    max_data_rows=10,  # 训练时只读取10行数据用于分析结构
          manual_headers=manual_headers,
                    active_sheet_only=True  # 只加载激活的sheet
                )

                file_structure = {
                    "file_name": file_name,
                    "sheets": {},
                    "total_regions": 0
                }

                for sheet_data in parsed_data:
                    sheet_structure = {
                        "sheet_name": sheet_data.sheet_name,
                        "regions": len(sheet_data.regions),
                        "headers": {},
                        "data_sample": []
                    }

                    for region in sheet_data.regions:
                        # 记录表头映射
                        sheet_structure["headers"].update(region.head_data)

                        # 记录数据样本（最多3行）
                        if region.data and len(sheet_structure["data_sample"]) < 3:
                            sheet_structure["data_sample"].append(region.data[0])

                    file_structure["sheets"][sheet_data.sheet_name] = sheet_structure
                    file_structure["total_regions"] += len(sheet_data.regions)

                structure["files"][file_name] = file_structure
                structure["total_sheets"] += len(parsed_data)
                structure["total_regions"] += file_structure["total_regions"]

            except Exception as e:
                self.training_logger.log_error(f"解析源文件 {file_path} 失败: {e}")
                structure["files"][Path(file_path).name] = {
                    "error": str(e),
                    "file_name": Path(file_path).name
                }

        return structure

    def _analyze_expected_structure(
        self, expected_file: str, manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """分析预期文件结构"""
        try:
            # 检查文件是否存在
            if not Path(expected_file).exists():
                raise FileNotFoundError(f"预期结果文件不存在: {expected_file}")

            parsed_data = self.excel_parser.parse_excel_file(
                expected_file,
             max_data_rows=10,  # 训练时只读取10行数据用于分析结构
                manual_headers=manual_headers,
                active_sheet_only=True  # 只加载激活的sheet
            )

            structure = {
                "file_name": Path(expected_file).name,
                "sheets": {},
                "total_regions": 0
            }

            for sheet_data in parsed_data:
                sheet_structure = {
                    "sheet_name": sheet_data.sheet_name,
                    "regions": len(sheet_data.regions),
                    "headers": {},
                    "data_sample": []
                }

                for region in sheet_data.regions:
                    # 记录表头映射
                    sheet_structure["headers"].update(region.head_data)

                    # 记录数据样本（最多3行）
                    if region.data and len(sheet_structure["data_sample"]) < 3:
                        sheet_structure["data_sample"].append(region.data[0])

                structure["sheets"][sheet_data.sheet_name] = sheet_structure
                structure["total_regions"] += len(sheet_data.regions)

            return structure

        except Exception as e:
            self.training_logger.log_error(f"解析预期文件 {expected_file} 失败: {e}")
            return {
                "file_name": Path(expected_file).name,
                "error": str(e),
                "sheets": {}
            }

    def _execute_and_validate(
        self,
        code: str,
        source_files: List[str],
        expected_file: str,
        manual_headers: Optional[Dict[str, Any]] = None,
        tenant_id: str = "default",
        iteration_num: int = 0
    ) -> Dict[str, Any]:
        """执行代码并验证结果"""
        result = {
            "success": False,
            "error": "",
            "output_file": "",
            "comparison": "",
            "execution_time": 0
        }

        try:
            # 创建临时输入文件夹
            import tempfile
            import shutil
            import os

            temp_dir = tempfile.mkdtemp()
            # 将短路径转换为长路径，避免Windows 8.3短路径格式导致的问题
            temp_dir = str(Path(temp_dir).resolve())
            input_dir = Path(temp_dir) / "input"
            output_dir = Path(temp_dir) / "output"
            input_dir.mkdir(exist_ok=True)
            output_dir.mkdir(exist_ok=True)

            # 复制源文件到输入文件夹
            for source_file in source_files:
                shutil.copy(source_file, input_dir / Path(source_file).name)

            # 准备执行环境
            execution_env = {
                "input_folder": str(input_dir),
                "output_folder": str(output_dir),
                "manual_headers": manual_headers or {},
                "source_files": [Path(f).name for f in source_files]
            }

            # 添加薪资参数（如果有）- 直接使用传入的值，不自动计算
            if hasattr(self, 'salary_year') and self.salary_year is not None:
                execution_env["salary_year"] = self.salary_year
            if hasattr(self, 'salary_month') and self.salary_month is not None:
                execution_env["salary_month"] = self.salary_month
            if hasattr(self, 'monthly_standard_hours') and self.monthly_standard_hours is not None:
                execution_env["monthly_standard_hours"] = self.monthly_standard_hours

            # 在沙箱中执行代码
            start_time = time.time()
            execution_result = self.sandbox.execute_script(code, execution_env)
            execution_time = time.time() - start_time

            result["execution_time"] = execution_time

            # 记录沙箱执行结果
            self.training_logger.log_info(f"沙箱执行结果: success={execution_result['success']}")
            if execution_result.get('output'):
                # 显示沙箱的完整输出（包括print语句）
                sandbox_output = execution_result['output']
                # 分行显示，让前端能看到每一行
                for line in sandbox_output.split('\n'):
                    if line.strip():
                        self.training_logger.log_info(f"[沙箱] {line}")
            if execution_result.get('error'):
                self.training_logger.log_error(f"沙箱错误: {execution_result['error'][:500]}...")

            if execution_result["success"]:
                # 查找生成的输出文件
                output_files = list(output_dir.glob("*.xlsx"))
                self.training_logger.log_info(f"在输出目录 {output_dir} 中找到 {len(output_files)} 个Excel文件")

                # 列出所有文件用于调试
                all_files = list(output_dir.glob("*"))
                self.training_logger.log_debug(f"输出目录中的所有文件: {[f.name for f in all_files]}")

                if output_files:
                    output_file = output_files[0]
                    result["output_file"] = str(output_file)
                    self.training_logger.log_info(f"找到输出文件: {output_file}")

                    # 将输出文件复制到training文件夹（加上时间戳）
                    from datetime import datetime
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    training_output_dir = Path(f"tenants/{tenant_id}/training_logs")
                    training_output_dir.mkdir(parents=True, exist_ok=True)
                    training_output_file = training_output_dir / f"output_iter{iteration_num:02d}_{timestamp}_{output_file.name}"
                    shutil.copy(output_file, training_output_file)
                    self.training_logger.log_info(f"输出文件已复制到: {training_output_file}")
                    result["training_output_file"] = str(training_output_file)

                    # 使用独立的差异对比组件进行对比
                    comparison_output_file = str(output_dir / "差异对比.xlsx")
                    comparison_result = compare_excel_files(
                        result_file=str(output_file),
                        expected_file=expected_file,
                        output_file=comparison_output_file
                    )

                    # 保存差异对比Excel到training_logs，然后删除临时文件
                    if Path(comparison_output_file).exists():
                        saved_comparison_file = self.training_logger.save_comparison_excel(comparison_output_file)
                        if saved_comparison_file:
                            result["comparison_excel_file"] = saved_comparison_file
                        # 删除临时的差异对比文件，只保留training_logs中的版本
                        Path(comparison_output_file).unlink(missing_ok=True)

                    # 根据对比结果判断是否成功
                    total_diff = comparison_result.get("total_differences", 0)
                    result["success"] = total_diff == 0
                    result["comparison"] = f"差异对比完成: 共发现 {total_diff} 处差异"
                    if total_diff == 0:
                        result["comparison"] = "所有检查项都通过！"
                    self.training_logger.log_info(f"对比结果: {result['comparison']}")
                else:
                    # 没有生成输出文件，记录沙箱的错误信息
                    sandbox_error = execution_result.get('error', '')
                    if sandbox_error:
                        result["error"] = f"未生成输出文件。沙箱错误: {sandbox_error}"
                    else:
                        result["error"] = "未生成输出文件"
                    self.training_logger.log_error(result["error"][:500])
            else:
                result["error"] = execution_result.get("error", "执行失败")
                self.training_logger.log_error(f"沙箱执行失败: {result['error'][:500]}...")

            # 清理临时目录
            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            result["error"] = f"执行验证过程中出错: {str(e)}"
            self.training_logger.log_error(f"执行验证失败: {e}")

        return result

    def _compare_files(
        self, actual_file: str, expected_file: str, manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """比较两个Excel文件"""
        try:
            # 解析实际输出文件
            actual_data = self.excel_parser.parse_excel_file(
                actual_file,
                manual_headers=manual_headers,
                active_sheet_only=True  # 只加载激活的sheet
            )
            expected_data = self.excel_parser.parse_excel_file(
                expected_file,
                manual_headers=manual_headers,
                active_sheet_only=True  # 只加载激活的sheet
            )

            # 转换为结构化的字典
            actual_structure = self._convert_to_structure(actual_data)
            expected_structure = self._convert_to_structure(expected_data)

            # 使用提示词生成器格式化对比结果
            return self.prompt_generator.format_comparison_result(actual_structure, expected_structure)

        except Exception as e:
            return f"比较文件时出错: {str(e)}"

    def _convert_to_structure(self, parsed_data: List[Any]) -> Dict[str, Any]:
        """将解析数据转换为结构化的字典"""
        structure = {
            "sheets": {}
        }

        for sheet_data in parsed_data:
            sheet_structure = {
                "sheet_name": sheet_data.sheet_name,
                "headers": {},
                "data": []
            }

            for region in sheet_data.regions:
                # 合并表头
                sheet_structure["headers"].update(region.head_data)

                # 添加数据
                for row in region.data:
                    sheet_structure["data"].append(row)

            structure["sheets"][sheet_data.sheet_name] = sheet_structure

        return structure

    def _calculate_match_score(self, execution_result: Dict[str, Any]) -> float:
        """计算匹配分数"""
        if not execution_result["success"]:
            return 0.0

        comparison = execution_result.get("comparison", "")
        if "所有检查项都通过！" in comparison:
            return 1.0

        # 简单的启发式评分
        # 可以根据实际需求实现更复杂的评分逻辑
        lines = comparison.split('\n')
        error_count = sum(1 for line in lines if '不一致' in line or '失败' in line)
        total_checks = len(lines)

        if total_checks == 0:
            return 0.0

        return max(0.0, 1.0 - (error_count / total_checks))