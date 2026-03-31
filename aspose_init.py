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
import time
import logging
import threading

logger = logging.getLogger(__name__)

# ==================== 初始化状态 ====================
_initialized = False
_license_applied = False
_license_obj = None          # 持有 License 对象引用，防止 GC 回收导致许可证失效
_license_bytes = None        # 许可证文件内容缓存（避免并发读文件冲突）
_lock = threading.Lock()     # 线程安全锁
_last_license_time = 0       # 上次成功注册许可证的时间戳
_LICENSE_COOLDOWN = 30       # 冷却时间（秒），此间隔内不重复 SetLicense

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

    首次从文件读取字节缓存到内存，后续使用 MemoryStream 设置，
    避免并发 FileStream 读取导致 "license file is corrupted" 错误。
    """
    global _license_applied, _license_obj, _license_bytes, _last_license_time

    path = lic_path or _lic_path

    try:
        # 首次：读取许可证文件到内存缓存
        if _license_bytes is None:
            if not os.path.exists(path):
                logger.warning(f"[Aspose] 许可证文件不存在: {path}")
                return False
            with open(path, 'rb') as f:
                _license_bytes = f.read()
            logger.info(f"[Aspose] 许可证文件已缓存到内存 ({len(_license_bytes)} bytes)")

        # 用 MemoryStream 设置许可证（线程安全，无文件 IO）
        from Aspose.Cells import License
        from System.IO import MemoryStream
        import System

        lic = License()
        byte_array = System.Array[System.Byte](list(_license_bytes))
        ms = MemoryStream(byte_array)
        lic.SetLicense(ms)
        ms.Close()

        _license_obj = lic
        _license_applied = True
        _last_license_time = time.time()
        logger.info(f"[Aspose] 许可证已生效（MemoryStream 方式）")
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

    冷却机制：30秒内不重复 SetLicense，避免高频调用。
    加锁：防止多线程同时 SetLicense 冲突。
    """
    if not _initialized:
        init_aspose()
        return

    # 冷却期内跳过（无锁快速路径）
    if _license_applied and (time.time() - _last_license_time) < _LICENSE_COOLDOWN:
        return

    with _lock:
        # 双重检查：拿到锁后再确认
        if _license_applied and (time.time() - _last_license_time) < _LICENSE_COOLDOWN:
            return
        _apply_license()


# ==================== 模块加载时自动初始化 ====================
init_aspose()
