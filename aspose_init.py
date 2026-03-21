"""
Aspose.Cells for .NET 全局初始化模块

程序启动时 import 一次即可，后续所有模块直接 from Aspose.Cells import ... 使用。
避免重复初始化 pythonnet 运行时和重复注册许可证。

用法:
    import aspose_init                          # 触发初始化
    from Aspose.Cells import Workbook           # 直接使用
"""

import os
import logging

logger = logging.getLogger(__name__)

# ==================== 初始化状态 ====================
_initialized = False
_license_applied = False

# 路径配置
_project_root = os.path.dirname(os.path.abspath(__file__))
_libs_dir = os.path.join(_project_root, "libs")
_lic_path = os.path.join(_libs_dir, "Aspose.Total.NET.lic")


def init_aspose():
    """一次性初始化 pythonnet + Aspose.Cells .NET 运行时 + 许可证

    幂等调用：多次调用安全，只执行一次。
    """
    global _initialized
    if _initialized:
        return True

    try:
        # 1. 注册 DLL 目录（Windows Python 3.8+ 必需）
        if hasattr(os, 'add_dll_directory'):
            os.add_dll_directory(_libs_dir)

        # 2. 加载 .NET Core 运行时
        import pythonnet
        pythonnet.load("coreclr", runtime_config=os.path.join(_libs_dir, "runtimeconfig.json"))

        # 3. 加载程序集（SkiaSharp 必须先于 Aspose.Cells）
        import clr
        clr.AddReference(os.path.join(_libs_dir, "SkiaSharp.dll"))
        clr.AddReference(os.path.join(_libs_dir, "Aspose.Cells.dll"))
        clr.AddReference("System.Text.Encoding.CodePages")

        # 4. 注册中文编码支持
        import System.Text
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

        _initialized = True
        logger.info("[Aspose] .NET 运行时初始化完成")

        # 5. 自动注册许可证
        _apply_license()

        return True
    except Exception as e:
        logger.error(f"[Aspose] 初始化失败: {e}")
        return False


def _apply_license(lic_path: str = None):
    """注册 Aspose 许可证"""
    global _license_applied
    if _license_applied:
        return True

    path = lic_path or _lic_path
    if not os.path.exists(path):
        logger.warning(f"[Aspose] 许可证文件不存在: {path}")
        return False

    try:
        from Aspose.Cells import License
        lic = License()
        lic.SetLicense(path)
        _license_applied = True
        logger.info(f"[Aspose] 许可证已生效: {os.path.basename(path)}")
        return True
    except Exception as e:
        logger.error(f"[Aspose] 许可证设置失败: {e}")
        return False


def is_initialized() -> bool:
    """检查是否已初始化"""
    return _initialized


def is_licensed() -> bool:
    """检查许可证是否已生效"""
    return _license_applied


# ==================== 模块加载时自动初始化 ====================
init_aspose()
