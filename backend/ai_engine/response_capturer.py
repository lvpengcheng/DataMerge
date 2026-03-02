"""
API响应捕获器 - 捕获AI提供者的原始响应
"""

import logging
from typing import Dict, List, Any, Optional
from .ai_provider import BaseAIProvider


class ResponseCapturer(BaseAIProvider):
    """API响应捕获器，包装AI提供者以捕获原始响应"""

    def __init__(self, wrapped_provider: BaseAIProvider, logger: logging.Logger = None):
        """初始化响应捕获器

        Args:
            wrapped_provider: 被包装的AI提供者
            logger: 日志记录器
        """
        self.wrapped_provider = wrapped_provider
        self.logger = logger or logging.getLogger(__name__)
        self.last_raw_response = None
        self.last_extracted_code = None

    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码并捕获原始响应"""
        try:
            # 调用被包装的提供者
            code = self.wrapped_provider.generate_code(prompt, **kwargs)

            # 尝试获取原始响应（如果被包装的提供者支持）
            if hasattr(self.wrapped_provider, 'last_raw_response'):
                self.last_raw_response = self.wrapped_provider.last_raw_response
                self.logger.info(f"ResponseCapturer: 捕获到原始响应，长度: {len(self.last_raw_response) if self.last_raw_response else 0}")

            self.last_extracted_code = code
            return code

        except Exception as e:
            self.logger.error(f"ResponseCapturer: 生成代码失败: {e}")
            raise

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口"""
        return self.wrapped_provider.chat(messages, **kwargs)

    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        return self.wrapped_provider.generate_completion(prompt, **kwargs)

    def extract_python_code(self, ai_response: str) -> str:
        """从AI响应中提取Python代码"""
        return self.wrapped_provider.extract_python_code(ai_response)

    def validate_and_fix_code_format(self, code: str) -> str:
        """验证和修复Python代码格式"""
        return self.wrapped_provider.validate_and_fix_code_format(code)

    def get_last_raw_response(self) -> Optional[str]:
        """获取最后一次的原始响应"""
        return self.last_raw_response

    def get_last_extracted_code(self) -> Optional[str]:
        """获取最后一次提取的代码"""
        return self.last_extracted_code