"""
Aspose.Cells 工具类
集中管理：格式转换（PDF）、加密/解密Excel、密码保护等
依赖 aspose_init 模块完成全局初始化（pythonnet + 许可证）
"""

import logging
import tempfile

import aspose_init  # noqa: F401 — 确保 Aspose 已初始化

from Aspose.Cells import (  # type: ignore
    Workbook, SaveFormat, PdfSaveOptions, LoadOptions, EncryptionType,
    FileFormatUtil,
)
from Aspose.Cells.Rendering.PdfSecurity import PdfSecurityOptions  # type: ignore

logger = logging.getLogger(__name__)


def is_encrypted(file_path: str) -> bool:
    """检测 Excel 文件是否有打开密码"""
    try:
        info = FileFormatUtil.DetectFileFormat(file_path)
        return bool(info.IsEncrypted)
    except Exception:
        return False


def convert_to_pdf(input_path: str, output_path: str = None, active_sheet_only: bool = True) -> str:
    """Excel 转 PDF

    Args:
        active_sheet_only: True=只导出激活sheet（默认），False=导出所有sheet
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".pdf")

    wb = Workbook(input_path)

    if active_sheet_only:
        # 隐藏非激活 sheet，只导出目标计算 sheet
        active_index = wb.Worksheets.ActiveSheetIndex
        for i in range(wb.Worksheets.Count):
            if i != active_index:
                wb.Worksheets[i].IsVisible = False

    opts = PdfSaveOptions()
    opts.CalculateFormula = True
    wb.Save(output_path, opts)

    logger.info(f"PDF转换完成: {input_path} -> {output_path} (active_only={active_sheet_only})")
    return output_path


def convert_to_encrypted_pdf(
    input_path: str,
    output_path: str = None,
    user_password: str = "123456",
    owner_password: str = "admin",
    active_sheet_only: bool = True,
) -> str:
    """Excel 转加密 PDF"""
    if not output_path:
        output_path = tempfile.mktemp(suffix=".pdf")

    wb = Workbook(input_path)

    if active_sheet_only:
        active_index = wb.Worksheets.ActiveSheetIndex
        for i in range(wb.Worksheets.Count):
            if i != active_index:
                wb.Worksheets[i].IsVisible = False

    opts = PdfSaveOptions()
    opts.CalculateFormula = True

    security = PdfSecurityOptions()
    security.UserPassword = user_password
    security.OwnerPassword = owner_password
    security.PrintPermission = True
    security.FullQualityPrintPermission = True
    opts.SecurityOptions = security

    wb.Save(output_path, opts)

    logger.info(f"加密PDF转换完成: {input_path} -> {output_path}")
    return output_path


def encrypt_excel(
    input_path: str,
    output_path: str = None,
    password: str = "123456",
) -> str:
    """Excel 文件加密（设置打开密码）
    使用 Aspose.Cells 原生加密（StrongCryptographicProvider, 128位）
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    wb = Workbook(input_path)
    wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
    wb.Settings.Password = password
    wb.Save(output_path)

    logger.info(f"Excel加密完成: {input_path} -> {output_path}")
    return output_path


def decrypt_excel(
    input_path: str,
    output_path: str = None,
    password: str = "",
) -> str:
    """Excel 文件解密（移除打开密码）
    使用 Aspose.Cells 原生解密
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    logger.info(f"[decrypt_excel] path={input_path}, pwd_len={len(password)}, pwd_repr={repr(password)}, pwd_type={type(password).__name__}")

    try:
        load_opts = LoadOptions()
        load_opts.Password = password
        wb = Workbook(input_path, load_opts)
    except Exception as e:
        err_str = str(e)
        if 'Invalid password' in err_str:
            # 密码不对。尝试无密码打开（文件可能实际未加密）
            logger.warning(f"[decrypt_excel] 密码无效, 尝试无密码打开: {input_path}")
            try:
                wb = Workbook(input_path)
            except Exception as e2:
                # 两种方式都失败，文件确实加密但密码不对
                logger.error(f"[decrypt_excel] 无密码打开也失败: {e2}")
                raise ValueError(
                    f"文件 '{input_path}' 加密且密码不正确 (pwd_len={len(password)}, pwd_repr={repr(password)})"
                ) from e
        else:
            raise
    wb.Settings.Password = None
    wb.Save(output_path)

    logger.info(f"Excel解密完成: {input_path} -> {output_path}")
    return output_path


def write_protect_excel(
    input_path: str,
    output_path: str = None,
    password: str = "123456",
) -> str:
    """Excel 写保护（可打开查看，编辑需密码）"""
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    wb = Workbook(input_path)
    wb.Settings.WriteProtection.Password = password
    wb.Settings.WriteProtection.RecommendReadOnly = True
    wb.Save(output_path)

    logger.info(f"Excel写保护完成: {input_path} -> {output_path}")
    return output_path


def convert_format(
    input_path: str,
    output_format: str,
    output_path: str = None,
) -> str:
    """通用格式转换（pdf, csv, html, xlsx, xls）"""
    format_map = {
        "pdf": (SaveFormat.Pdf, ".pdf"),
        "csv": (SaveFormat.Csv, ".csv"),
        "html": (SaveFormat.Html, ".html"),
        "xlsx": (SaveFormat.Xlsx, ".xlsx"),
        "xls": (SaveFormat.Excel97To2003, ".xls"),
    }

    fmt = output_format.lower()
    if fmt not in format_map:
        raise ValueError(f"不支持的格式: {output_format}，可选: {list(format_map.keys())}")

    save_fmt, ext = format_map[fmt]
    if not output_path:
        output_path = tempfile.mktemp(suffix=ext)

    wb = Workbook(input_path)
    wb.Save(output_path, save_fmt)

    logger.info(f"格式转换完成: {input_path} -> {output_path} ({fmt})")
    return output_path
