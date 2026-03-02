"""
训练日志记录器 - 提供详细的训练过程日志和流式显示支持
"""

import json
import logging
import sys
import time
from typing import Dict, List, Any, Optional, Callable
from pathlib import Path
from datetime import datetime


class TrainingLogger:
    """训练日志记录器"""

    def __init__(self, tenant_id: str, log_dir: Optional[str] = None, keyword: str = "training"):
        """初始化训练日志记录器

        Args:
            tenant_id: 租户ID
            log_dir: 日志目录，如果为None则使用默认目录
            keyword: 关键词，用于日志文件命名
        """
        self.tenant_id = tenant_id
        self.keyword = keyword
        self.logger = logging.getLogger(f"training.{tenant_id}")
        self.logger.setLevel(logging.INFO)  # 设置日志级别

        # 清除已有的处理器，避免重复
        self.logger.handlers.clear()
        self.logger.propagate = False  # 防止日志重复输出

        # 设置日志目录（与StorageManager保持一致，使用相对路径tenants）
        if log_dir:
            self.log_dir = Path(log_dir)
        else:
            # 使用与StorageManager一致的路径：tenants/{tenant_id}/training_logs
            self.log_dir = Path("tenants") / tenant_id / "training_logs"

        self.log_dir.mkdir(parents=True, exist_ok=True)
        print(f"[TrainingLogger] 日志目录: {self.log_dir.absolute()}")

        # 创建日志文件（使用新命名格式：租户_关键词_日期_时间）
        self.session_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = self.log_dir / f"{tenant_id}_{keyword}_{self.session_timestamp}.log"

        # 添加文件处理器
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        ))
        self.logger.addHandler(file_handler)

        # 添加控制台处理器，确保日志输出到终端
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        ))
        self.logger.addHandler(console_handler)

        # 训练状态
        self.current_iteration = 0
        self.total_iterations = 0
        self.start_time = None
        self.stream_callback = None

    def set_stream_callback(self, callback: Callable[[str], None]):
        """设置流式回调函数，用于实时显示日志

        Args:
            callback: 回调函数，接收日志消息作为参数
        """
        self.stream_callback = callback

    def _stream_log(self, message: str, level: str = "INFO"):
        """流式输出日志

        Args:
            message: 日志消息
            level: 日志级别
        """
        if self.stream_callback:
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] [{level}] {message}"
            self.stream_callback(formatted_message)

    def start_training(self, total_iterations: int, source_files: List[str],
                      expected_file: str, rule_files: List[str]):
        """开始训练记录

        Args:
            total_iterations: 总迭代次数
            source_files: 源文件列表
            expected_file: 预期结果文件
            rule_files: 规则文件列表
        """
        self.start_time = time.time()
        self.total_iterations = total_iterations
        self.current_iteration = 0

        log_message = f"开始训练 - 租户: {self.tenant_id}"
        self.logger.info(log_message)
        self._stream_log(log_message)

        # 记录训练配置
        config_info = {
            "total_iterations": total_iterations,
            "source_files": [Path(f).name for f in source_files],
            "expected_file": Path(expected_file).name,
            "rule_files": [Path(f).name for f in rule_files],
            "start_time": datetime.now().isoformat()
        }

        self.logger.info(f"训练配置: {json.dumps(config_info, ensure_ascii=False)}")
        self._stream_log(f"训练配置: 总迭代次数={total_iterations}, 源文件={len(source_files)}, 规则文件={len(rule_files)}")

    def start_iteration(self, iteration: int, iteration_type: str = "training"):
        """开始迭代记录

        Args:
            iteration: 迭代次数
            iteration_type: 迭代类型 (training/correction)
        """
        self.current_iteration = iteration
        iteration_type_str = "训练" if iteration_type == "training" else "修正"

        log_message = f"开始第 {iteration}/{self.total_iterations} 次迭代 ({iteration_type_str})"
        self.logger.info(log_message)
        self._stream_log(log_message)

    def log_prompt_generation(self, prompt_type: str, prompt_length: int,
                             prompt_preview: str = None):
        """记录提示词生成

        Args:
            prompt_type: 提示词类型 (training/correction)
            prompt_length: 提示词长度
            prompt_preview: 提示词预览
        """
        prompt_type_str = "训练" if prompt_type == "training" else "修正"

        log_message = f"生成{prompt_type_str}提示词 - 长度: {prompt_length} 字符"
        self.logger.info(log_message)
        self._stream_log(log_message)

        if prompt_preview:
            preview_log = f"提示词预览: {prompt_preview}"
            self.logger.debug(preview_log)
            self._stream_log(preview_log, level="DEBUG")

    def log_full_prompt(self, prompt: str, prompt_type: str = "generate"):
        """记录完整的提示词到单独文件

        Args:
            prompt: 完整的提示词
            prompt_type: 提示词类型 (generate/correct)
        """
        # 统一命名格式：租户_关键词_prompt_日期时间_第几次_类型.txt
        iteration_str = f"{self.current_iteration:02d}"
        prompt_file = self.log_dir / f"{self.tenant_id}_{self.keyword}_prompt_{self.session_timestamp}_{iteration_str}_{prompt_type}.txt"

        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(f"# 提示词类型: {prompt_type}\n")
            f.write(f"# 迭代次数: {self.current_iteration}\n")
            f.write(f"# 时间: {datetime.now().isoformat()}\n")
            f.write(f"# 长度: {len(prompt)} 字符\n")
            f.write("=" * 60 + "\n\n")
            f.write(prompt)

        self.logger.info(f"完整提示词已保存到: {prompt_file}")
        self._stream_log(f"完整提示词已保存到: {prompt_file.name}")

    def log_full_ai_response(self, response: str, response_type: str = "generate"):
        """记录完整的AI响应到单独文件

        Args:
            response: 完整的AI响应
            response_type: 响应类型 (generate/correct)
        """
        # 统一命名格式：租户_关键词_response_日期时间_第几次_类型.txt
        iteration_str = f"{self.current_iteration:02d}"
        response_file = self.log_dir / f"{self.tenant_id}_{self.keyword}_response_{self.session_timestamp}_{iteration_str}_{response_type}.txt"

        with open(response_file, 'w', encoding='utf-8') as f:
            f.write(f"# AI响应类型: {response_type}\n")
            f.write(f"# 迭代次数: {self.current_iteration}\n")
            f.write(f"# 时间: {datetime.now().isoformat()}\n")
            f.write(f"# 长度: {len(response)} 字符\n")
            f.write("=" * 60 + "\n\n")
            f.write(response)

        self.logger.info(f"完整AI响应已保存到: {response_file}")
        self._stream_log(f"完整AI响应已保存到: {response_file.name}")

    def log_generated_code(self, code: str, mode_type: str = "formula"):
        """记录生成的代码到单独文件

        Args:
            code: 生成的Python代码
            mode_type: 模式类型 (formula/modular/standard)
        """
        # 统一命名格式：租户_关键词_code_日期时间_第几次_模式.py
        iteration_str = f"{self.current_iteration:02d}"
        code_file = self.log_dir / f"{self.tenant_id}_{self.keyword}_code_{self.session_timestamp}_{iteration_str}_{mode_type}.py"

        with open(code_file, 'w', encoding='utf-8') as f:
            f.write(f"# 迭代次数: {self.current_iteration}\n")
            f.write(f"# 生成时间: {datetime.now().isoformat()}\n")
            f.write(f"# 模式: {mode_type}\n")
            f.write(f"# 代码长度: {len(code)} 字符\n")
            f.write("# " + "=" * 58 + "\n\n")
            f.write(code)

        self.logger.info(f"生成的代码已保存到: {code_file}")
        self._stream_log(f"生成的代码已保存到: {code_file.name}")

    def log_streaming_chunk(self, chunk: str):
        """记录流式响应块（实时显示AI生成的代码）

        Args:
            chunk: 响应块
        """
        if chunk:
            # 在控制台直接打印（不换行，实现流式效果）
            import sys
            sys.stdout.write(chunk)
            sys.stdout.flush()

            # 同时发送到SSE队列
            if self.stream_callback:
                # 包装成特定格式，让前端能识别这是AI代码流式输出
                # 格式: [HH:MM:SS] [CODE] chunk内容
                from datetime import datetime
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_msg = f"[{timestamp}] [CODE] {chunk}"
                self.stream_callback(formatted_msg)

    def log_ai_api_call(self, api_type: str, request_data: Dict[str, Any] = None):
        """记录AI API调用

        Args:
            api_type: API类型 (generate_code/stream_generate_code)
            request_data: 请求数据
        """
        log_message = f"调用AI API - 类型: {api_type}"
        self.logger.info(log_message)
        self._stream_log(log_message)

        if request_data:
            self.logger.debug(f"API请求数据: {json.dumps(request_data, ensure_ascii=False)[:500]}...")

    def log_streaming_response(self, chunk: str, is_complete: bool = False):
        """记录流式响应

        Args:
            chunk: 响应块
            is_complete: 是否完成
        """
        if chunk.strip():
            self._stream_log(f"AI响应: {chunk}", level="DEBUG")

            if is_complete:
                self.logger.info(f"AI响应完成，总长度: {len(chunk)} 字符")
                self._stream_log(f"AI响应完成，总长度: {len(chunk)} 字符")

    def log_code_generated(self, code_length: int, code_preview: str = None):
        """记录代码生成

        Args:
            code_length: 代码长度
            code_preview: 代码预览
        """
        log_message = f"生成代码 - 长度: {code_length} 字符"
        self.logger.info(log_message)
        self._stream_log(log_message)

        if code_preview:
            preview_log = f"代码预览: {code_preview}"
            self.logger.debug(preview_log)
            self._stream_log(preview_log, level="DEBUG")

    def log_execution_start(self):
        """记录代码执行开始"""
        log_message = "开始执行生成的代码"
        self.logger.info(log_message)
        self._stream_log(log_message)

    def log_execution_result(self, success: bool, execution_time: float,
                            error: str = None, output_file: str = None):
        """记录代码执行结果

        Args:
            success: 是否成功（注意：这是验证结果，不是执行结果）
            execution_time: 执行时间
            error: 错误信息
            output_file: 输出文件
        """
        if output_file:
            # 代码执行成功并生成了输出文件，但验证可能未通过
            if success:
                log_message = f"代码执行成功，验证通过 - 耗时: {execution_time:.2f}秒, 输出文件: {Path(output_file).name}"
            else:
                log_message = f"代码执行成功，但验证未通过 - 耗时: {execution_time:.2f}秒, 输出文件: {Path(output_file).name}"
        else:
            # 没有输出文件，代码执行失败
            log_message = f"代码执行失败 - 耗时: {execution_time:.2f}秒"

        self.logger.info(log_message)
        self._stream_log(log_message)

        if error:
            self.logger.error(f"执行错误: {error}")
            self._stream_log(f"执行错误: {error}", level="ERROR")

    def log_comparison_result(self, comparison: str, score: float):
        """记录比较结果

        Args:
            comparison: 比较结果文本
            score: 匹配分数
        """
        log_message = f"文件比较完成 - 匹配分数: {score:.2%}"
        self.logger.info(log_message)
        self._stream_log(log_message)

        # 记录详细的比较结果
        if comparison:
            lines = comparison.split('\n')
            for line in lines[:10]:  # 只显示前10行
                if line.strip():
                    self.logger.debug(f"比较结果: {line}")
                    self._stream_log(f"比较结果: {line}", level="DEBUG")

            if len(lines) > 10:
                self.logger.debug(f"... 还有 {len(lines) - 10} 行比较结果")
                self._stream_log(f"... 还有 {len(lines) - 10} 行比较结果", level="DEBUG")

    def save_comparison_excel(self, comparison_excel_path: str, mode_type: str = "formula") -> str:
        """将差异对比Excel保存到training_logs目录

        Args:
            comparison_excel_path: 原始对比Excel文件路径
            mode_type: 模式类型 (formula/modular/standard)

        Returns:
            保存后的文件路径
        """
        import shutil

        source_path = Path(comparison_excel_path)
        if not source_path.exists():
            self.logger.warning(f"差异对比文件不存在: {comparison_excel_path}")
            return ""

        # 统一命名格式：租户_关键词_comparison_日期时间_第几次_模式.xlsx
        iteration_str = f"{self.current_iteration:02d}"
        dest_filename = f"{self.tenant_id}_{self.keyword}_comparison_{self.session_timestamp}_{iteration_str}_{mode_type}.xlsx"
        dest_path = self.log_dir / dest_filename

        shutil.copy2(source_path, dest_path)

        self.logger.info(f"差异对比Excel已保存到: {dest_path}")
        self._stream_log(f"差异对比Excel已保存到: {dest_path.name}")

        return str(dest_path)

    def save_output_excel(self, output_excel_path: str, mode_type: str = "formula") -> str:
        """将生成的输出Excel保存到training_logs目录

        Args:
            output_excel_path: 原始输出Excel文件路径
            mode_type: 模式类型 (formula/modular/standard)

        Returns:
            保存后的文件路径
        """
        import shutil

        source_path = Path(output_excel_path)
        if not source_path.exists():
            self.logger.warning(f"输出Excel文件不存在: {output_excel_path}")
            return ""

        # 统一命名格式：租户_关键词_output_日期时间_第几次_模式.xlsx
        iteration_str = f"{self.current_iteration:02d}"
        dest_filename = f"{self.tenant_id}_{self.keyword}_output_{self.session_timestamp}_{iteration_str}_{mode_type}.xlsx"
        dest_path = self.log_dir / dest_filename

        shutil.copy2(source_path, dest_path)

        self.logger.info(f"输出Excel已保存到: {dest_path}")
        self._stream_log(f"输出Excel已保存到: {dest_path.name}")

        return str(dest_path)

    def log_iteration_complete(self, iteration: int, score: float, is_best: bool = False):
        """记录迭代完成

        Args:
            iteration: 迭代次数
            score: 匹配分数
            is_best: 是否是最佳结果
        """
        best_marker = " (最佳)" if is_best else ""
        log_message = f"第 {iteration} 次迭代完成 - 分数: {score:.2%}{best_marker}"
        self.logger.info(log_message)
        self._stream_log(log_message)

    def log_training_complete(self, best_score: float, total_iterations: int,
                             success: bool, best_code_length: int):
        """记录训练完成

        Args:
            best_score: 最佳分数
            total_iterations: 总迭代次数
            success: 是否成功
            best_code_length: 最佳代码长度
        """
        elapsed_time = time.time() - self.start_time if self.start_time else 0
        status = "成功" if success else "失败"

        log_message = (
            f"训练完成 - 状态: {status}, "
            f"最佳分数: {best_score:.2%}, "
            f"总迭代次数: {total_iterations}, "
            f"总耗时: {elapsed_time:.2f}秒, "
            f"最佳代码长度: {best_code_length} 字符"
        )

        self.logger.info(log_message)
        self._stream_log(log_message)

        # 保存训练摘要
        self._save_training_summary(best_score, total_iterations, success, elapsed_time)

    def _save_training_summary(self, best_score: float, total_iterations: int,
                              success: bool, elapsed_time: float):
        """保存训练摘要

        Args:
            best_score: 最佳分数
            total_iterations: 总迭代次数
            success: 是否成功
            elapsed_time: 耗时
        """
        summary = {
            "tenant_id": self.tenant_id,
            "training_completed": datetime.now().isoformat(),
            "success": success,
            "best_score": best_score,
            "total_iterations": total_iterations,
            "elapsed_time_seconds": elapsed_time,
            "log_file": str(self.log_file),
            "summary": f"训练{'成功' if success else '失败'}，最佳分数: {best_score:.2%}"
        }

        summary_file = self.log_dir / "training_summary.json"
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

        self.logger.info(f"训练摘要已保存到: {summary_file}")

    def log_error(self, error_message: str, exception: Exception = None):
        """记录错误

        Args:
            error_message: 错误消息
            exception: 异常对象
        """
        self.logger.error(error_message, exc_info=exception)
        self._stream_log(f"错误: {error_message}", level="ERROR")
        # 强制输出到控制台
        sys.stdout.write(f"[训练错误] {error_message}\n")
        sys.stdout.flush()

        if exception:
            self._stream_log(f"异常详情: {str(exception)}", level="ERROR")
            sys.stdout.write(f"[训练错误] 异常详情: {str(exception)}\n")
            sys.stdout.flush()

    def log_warning(self, warning_message: str):
        """记录警告

        Args:
            warning_message: 警告消息
        """
        self.logger.warning(warning_message)
        self._stream_log(f"警告: {warning_message}", level="WARNING")
        # 强制输出到控制台
        sys.stdout.write(f"[训练警告] {warning_message}\n")
        sys.stdout.flush()

    def log_info(self, info_message: str):
        """记录信息

        Args:
            info_message: 信息消息
        """
        self.logger.info(info_message)
        self._stream_log(info_message)
        # 强制输出到控制台，确保后端能看到
        sys.stdout.write(f"[训练日志] {info_message}\n")
        sys.stdout.flush()

    def log_debug(self, debug_message: str):
        """记录调试信息

        Args:
            debug_message: 调试消息
        """
        self.logger.debug(debug_message)
        self._stream_log(debug_message, level="DEBUG")

    def get_log_file_path(self) -> str:
        """获取日志文件路径

        Returns:
            日志文件路径
        """
        return str(self.log_file)

    def get_training_summary_path(self) -> str:
        """获取训练摘要文件路径

        Returns:
            训练摘要文件路径
        """
        return str(self.log_dir / "training_summary.json")


class StreamAwareAIProvider:
    """支持流式日志的AI提供者包装器"""

    def __init__(self, ai_provider, training_logger: TrainingLogger):
        """初始化

        Args:
            ai_provider: 原始AI提供者
            training_logger: 训练日志记录器
        """
        self.ai_provider = ai_provider
        self.training_logger = training_logger

    def generate_code(self, prompt: str) -> str:
        """生成代码（支持流式日志）

        Args:
            prompt: 提示词

        Returns:
            生成的代码
        """
        # 记录API调用
        self.training_logger.log_ai_api_call("generate_code", {"prompt_length": len(prompt)})

        try:
            # 尝试使用流式生成（如果支持）
            if hasattr(self.ai_provider, 'stream_generate_code'):
                return self._stream_generate_code_with_logging(prompt)
            else:
                # 使用普通生成
                code = self.ai_provider.generate_code(prompt)
                self.training_logger.log_code_generated(len(code), code[:200])
                return code
        except Exception as e:
            self.training_logger.log_error(f"AI代码生成失败: {str(e)}", e)
            raise

    def _stream_generate_code_with_logging(self, prompt: str) -> str:
        """使用流式生成代码并记录日志

        Args:
            prompt: 提示词

        Returns:
            生成的代码
        """
        full_response = ""

        # 记录流式API调用
        self.training_logger.log_ai_api_call("stream_generate_code", {"prompt_length": len(prompt)})

        try:
            # 获取流式响应
            stream = self.ai_provider.stream_generate_code(prompt)

            for chunk in stream:
                if chunk:
                    full_response += chunk
                    # 记录流式响应
                    self.training_logger.log_streaming_response(chunk)

            # 记录完成
            self.training_logger.log_streaming_response("", is_complete=True)
            self.training_logger.log_code_generated(len(full_response), full_response[:200])

            return full_response

        except Exception as e:
            self.training_logger.log_error(f"AI流式代码生成失败: {str(e)}", e)
            raise