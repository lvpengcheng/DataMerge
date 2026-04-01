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
_license_bytes = None        # 许可证文件内容缓存（避免并发读文件冲突）
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

    首次用 .NET 原生 File.ReadAllBytes 读取字节并缓存，避免 Python→.NET
    字节转换导致 "license file is corrupted"。后续使用缓存的 .NET 字节数组
    通过 MemoryStream 设置，线程安全且无文件 IO。
    """
    global _license_applied, _license_obj, _license_bytes

    path = lic_path or _lic_path

    try:
        from Aspose.Cells import License
        from System.IO import MemoryStream

        # 首次：用 .NET 原生读取许可证文件到 .NET byte[]（避免 Python bytes 转换损坏）
        if _license_bytes is None:
            if not os.path.exists(path):
                logger.warning(f"[Aspose] 许可证文件不存在: {path}")
                return False
            from System.IO import File as NetFile
            _license_bytes = NetFile.ReadAllBytes(path)
            logger.info(f"[Aspose] 许可证文件已缓存到 .NET 内存 ({_license_bytes.Length} bytes)")

        # 用 MemoryStream 设置许可证（线程安全，无文件 IO）
        lic = License()
        ms = MemoryStream(_license_bytes)
        ms.Position = 0
        lic.SetLicense(ms)
        ms.Close()

        _license_obj = lic
        _license_applied = True
        return True
    except Exception as e:
        logger.error(f"[Aspose] 许可证设置失败: {e}")
        # 回退：直接用文件路径（仅初始化时，无并发风险）
        try:
            logger.info("[Aspose] 尝试回退方案：直接用文件路径设置许可证...")
            from Aspose.Cells import License
            lic = License()
            lic.SetLicense(lic_path or _lic_path)
            _license_obj = lic
            _license_applied = True
            logger.info("[Aspose] 回退方案成功")
            return True
        except Exception as e2:
            logger.error(f"[Aspose] 回退方案也失败: {e2}")
            return False


def is_initialized() -> bool:
    """检查是否已初始化"""
    return _initialized


def is_licensed() -> bool:
    """检查许可证是否已生效（Python 端标志）"""
    return _license_applied



def ensure_license():
    """确保 Aspose 许可证已注册

    每次调用都通过 MemoryStream 重新 SetLicense。
    MemoryStream 无文件 IO，不会有并发 "corrupted" 问题，无需冷却。
    加锁防止多线程同时调用 SetLicense。
    """
    if not _initialized:
        init_aspose()
        return

    with _lock:
        _apply_license()


# ==================== 模块加载时自动初始化 ====================
init_aspose()
