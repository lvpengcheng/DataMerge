"""
Aspose.Cells for .NET 全局初始化模块

程序启动时 import 一次即可，后续所有模块直接 from Aspose.Cells import ... 使用。
避免重复初始化 pythonnet 运行时和重复注册许可证。

用法:
    import aspose_init                          # 触发初始化
    from Aspose.Cells import Workbook           # 直接使用
"""

import os
import sys
import logging
import threading

logger = logging.getLogger(__name__)

# ==================== 初始化状态 ====================
_initialized = False
_license_applied = False
_license_obj = None          # 持有 License 对象引用，防止 GC 回收导致许可证失效
_lock = threading.Lock()     # 线程安全锁

# 路径配置
_project_root = os.path.dirname(os.path.abspath(__file__))
_libs_dir = os.path.join(_project_root, "libs")
_lic_path = os.path.join(_libs_dir, "Aspose.Total.NET.lic")


def init_aspose():
    """一次性初始化 pythonnet + Aspose.Cells .NET 运行时 + 许可证

    幂等调用：多次调用安全，只执行一次。线程安全。
    """
    global _initialized
    if _initialized:
        return True

    with _lock:
        if _initialized:
            return True

        try:
            # 1. 将 libs 目录加入程序集搜索路径
            if _libs_dir not in sys.path:
                sys.path.append(_libs_dir)

            # 2. Windows: 注册 DLL 目录（Python 3.8+ 必需）
            if hasattr(os, 'add_dll_directory'):
                os.add_dll_directory(_libs_dir)

            # 3. Linux: 用 ctypes 预加载 libSkiaSharp.so 原生库
            if sys.platform == 'linux':
                import ctypes
                so_path = os.path.join(_libs_dir, "libSkiaSharp.so")
                if os.path.exists(so_path):
                    try:
                        ctypes.cdll.LoadLibrary(so_path)
                        logger.info(f"[Aspose] Linux: 预加载 libSkiaSharp.so 成功")
                    except OSError as e:
                        logger.warning(f"[Aspose] Linux: 预加载 libSkiaSharp.so 失败: {e}")
                        try:
                            ctypes.cdll.LoadLibrary("libSkiaSharp.so")
                            logger.info(f"[Aspose] Linux: 从系统路径加载 libSkiaSharp.so 成功")
                        except OSError as e2:
                            logger.error(f"[Aspose] Linux: libSkiaSharp.so 加载全部失败: {e2}")

            # 4. 加载 .NET Core 运行时
            import pythonnet
            pythonnet.load("coreclr", runtime_config=os.path.join(_libs_dir, "runtimeconfig.json"))

            # 5. 加载程序集（用短名，从 sys.path 搜索；SkiaSharp 必须先于 Aspose.Cells）
            import clr
            clr.AddReference("SkiaSharp")
            clr.AddReference("Aspose.Cells")

            # 6. 注册中文编码支持（可选，失败不影响核心功能）
            try:
                import System.Text
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
                logger.info("[Aspose] CodePages 编码支持已注册")
            except Exception as enc_err:
                logger.warning(f"[Aspose] CodePages 编码注册跳过（不影响核心功能）: {enc_err}")

            _initialized = True
            logger.info("[Aspose] .NET 运行时初始化完成")

            # 7. 自动注册许可证
            _apply_license()

            return True
        except Exception as e:
            logger.error(f"[Aspose] 初始化失败: {e}")
            return False


def _apply_license(lic_path: str = None):
    """注册 Aspose 许可证

    将 License 对象保存到模块级变量 _license_obj，
    防止 Python GC 回收 .NET 对象导致许可证失效。
    """
    global _license_applied, _license_obj

    path = lic_path or _lic_path
    if not os.path.exists(path):
        logger.warning(f"[Aspose] 许可证文件不存在: {path}")
        return False

    try:
        from Aspose.Cells import License
        lic = License()

        # 优先使用 Stream 方式（pythonnet 下更可靠）
        try:
            from System.IO import FileStream, FileMode, FileAccess, FileShare
            stream = FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
            lic.SetLicense(stream)
            stream.Close()
            logger.info(f"[Aspose] 许可证已生效（Stream 方式）: {os.path.basename(path)}")
        except Exception as stream_err:
            logger.warning(f"[Aspose] Stream 方式失败: {stream_err}, 尝试路径方式")
            lic = License()
            lic.SetLicense(path)
            logger.info(f"[Aspose] 许可证已生效（路径方式）: {os.path.basename(path)}")

        # 关键：保持 License 对象的模块级引用，防止 GC 回收
        _license_obj = lic
        _license_applied = True
        return True
    except Exception as e:
        logger.error(f"[Aspose] 许可证设置失败: {e}")
        return False


def is_initialized() -> bool:
    """检查是否已初始化"""
    return _initialized


def is_licensed() -> bool:
    """检查许可证是否已生效（Python 端标志）"""
    return _license_applied



def ensure_license():
    """确保 Aspose 许可证已注册

    轻量级检查：由于 _license_obj 持有 License 对象引用防止 GC，
    正常情况下许可证不会丢失，这里只做 Python 标志检查。
    仅在未初始化或标志异常时才重新注册。
    """
    global _license_applied

    if not _initialized:
        init_aspose()
        return

    # 快速路径：License 对象存活 + 标志已设置 → 直接返回
    if _license_applied and _license_obj is not None:
        return

    # 异常路径（加锁）：标志丢失，需要重新注册
    with _lock:
        if _license_applied and _license_obj is not None:
            return

        logger.warning("[Aspose] 许可证标志异常，重新注册")
        _license_applied = False
        _apply_license()
        if _license_applied:
            logger.info("[Aspose] 许可证重新注册成功")
        else:
            logger.error("[Aspose] 许可证重新注册失败！将以评估模式运行")


# ==================== 模块加载时自动初始化 ====================
init_aspose()
