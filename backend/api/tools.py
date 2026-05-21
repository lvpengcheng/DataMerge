"""智能小工具 API"""

import io
import os
import sys
import shutil
import tempfile
import zipfile
import logging
from pathlib import Path
from typing import List

from fastapi import APIRouter, Depends, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse

from ..auth.dependencies import get_current_user

router = APIRouter(prefix="/api/tools", tags=["智能小工具"])

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}


def _import_split_one_file():
    """延迟导入根目录的 split_by_banner.split_one_file"""
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    from split_by_banner import split_one_file
    return split_one_file


@router.post("/split-by-banner")
async def split_by_banner(
    files: List[UploadFile] = File(...),
    current_user=Depends(get_current_user),
):
    """按 banner 拆分 sheet:接收多个 Excel,逐个拆分,打包为 zip 返回"""
    if not files:
        raise HTTPException(status_code=400, detail="未上传文件")

    split_one_file = _import_split_one_file()
    work_dir = Path(tempfile.mkdtemp(prefix="split_banner_"))
    src_dir = work_dir / "src"
    out_dir = work_dir / "out"
    src_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    errors: List[str] = []
    success_files: List[Path] = []

    try:
        for uf in files:
            name = uf.filename or "unnamed.xlsx"
            ext = Path(name).suffix.lower()
            if ext not in EXCEL_EXTS:
                errors.append(f"{name}: 不支持的扩展名({ext})")
                continue
            src_path = src_dir / name
            try:
                content = await uf.read()
                src_path.write_bytes(content)
            except Exception as e:
                errors.append(f"{name}: 写入失败 {e}")
                continue

            out_name = f"{src_path.stem}_split.xlsx"
            out_path = out_dir / out_name
            try:
                split_one_file(src_path, out_path)
                if out_path.exists():
                    success_files.append(out_path)
                else:
                    errors.append(f"{name}: 拆分未生成输出")
            except Exception as e:
                logger.exception(f"拆分失败: {name}")
                errors.append(f"{name}: {e}")

        if not success_files:
            detail = "全部失败:\n" + "\n".join(errors) if errors else "未生成任何输出"
            raise HTTPException(status_code=400, detail=detail)

        # 打包 zip
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in success_files:
                zf.write(p, arcname=p.name)
            if errors:
                zf.writestr("_errors.txt", "\n".join(errors))
        buf.seek(0)

        # 同步清理(StreamingResponse 已读完 buffer)
        shutil.rmtree(work_dir, ignore_errors=True)

        headers = {"Content-Disposition": 'attachment; filename="split_results.zip"'}
        if errors:
            headers["X-Split-Errors"] = str(len(errors))
        return StreamingResponse(buf, media_type="application/zip", headers=headers)

    except HTTPException:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        logger.exception("拆分接口异常")
        raise HTTPException(status_code=500, detail=str(e))
