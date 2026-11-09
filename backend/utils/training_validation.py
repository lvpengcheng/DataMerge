"""Shared post-execution validation for every training entry point."""
import logging

from backend.utils.subprocess_runner import run_in_subprocess, default_timeout, default_max_memory_mb

logger = logging.getLogger(__name__)


def training_comparison_sheets(code, expected_structure, selected_sheets=None):
    """Prefer the training selection, then the template targets recorded in the script."""
    from backend.utils.compute_precheck import _extract_colmap

    if isinstance(selected_sheets, dict):
        selected_sheets = selected_sheets.get((expected_structure or {}).get("file_name"))
    if selected_sheets:
        return list(selected_sheets)
    targets = _extract_colmap(code) if code and "_COL_MAP" in code else None
    return list(targets or (expected_structure or {}).get("sheets") or {}) or None


def compare_for_training(compare):
    """Comparison runs after output generation and may leave accuracy unassessed."""
    from backend.utils.excel_comparator import require_complete_comparison

    comparison = {}
    try:
        comparison = compare()
        return require_complete_comparison(comparison)
    except Exception as exc:
        logger.warning("智训结果对比未完成，已保留生成结果: %s", exc)
        reason = "结果对比超时" if comparison.get("timed_out") else f"结果对比未完成：{exc}"
        return {
            "comparison_complete": False,
            "warning": f"{reason}。结果已生成，可下载，成功率尚未评估。",
        }


class TrainingResourceFailure(RuntimeError):
    """Stop AI correction retries when execution resources failed."""


def cleanup_training_directory(path):
    from backend.utils.resource_guard import ensure_healthy
    try:
        ensure_healthy()
    except RuntimeError:
        return  # An unreaped child may still own these files.
    import shutil
    shutil.rmtree(path, ignore_errors=True)


def finalize_training_output(output_path, code, expected_structure=None, template_path=None):
    result = run_in_subprocess(
        'backend.utils.training_validation:_finalize_training_output_impl',
        (str(output_path), code, expected_structure, template_path),
        timeout=default_timeout('write'), max_memory_mb=default_max_memory_mb())
    if not result.success:
        if result.killed or result.termination_failed:
            raise TrainingResourceFailure(result.error)
        raise ValueError(result.error or '训练输出后处理失败')
    return result.result


def _finalize_training_output_impl(output_path, code, expected_structure, template_path):
    if not template_path:
        import ast
        for node in ast.parse(code).body:
            if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'TEMPLATE_PATH' for t in node.targets):
                try:
                    candidate = ast.literal_eval(node.value)
                    if isinstance(candidate, str):
                        template_path = candidate
                except (ValueError, TypeError):
                    pass
    from backend.utils.output_postprocess import finalize_output_workbook
    from backend.utils.excel_comparator import inspect_formula_cache
    result = finalize_output_workbook(output_path, template_path, expected_structure, code)
    report = inspect_formula_cache(output_path)
    # Excel 已经完成填充、重算和保存时，公式错误值是可以对比和下载的
    # 业务结果。例如 VLOOKUP 未命中会产生 #N/A，这应该进入准确率对比并
    # 作为下一轮修正依据，不能把已生成的工作簿判为“执行失败”。
    # 文件无法打开/计算/保存仍会由 finalize_output_workbook 直接抛错。
    result['formula_report'] = report
    result['formula_warning'] = bool(
        report.get('empty_cache_count')
        or report.get('invalid_ref_formula_count')
        or report.get('error_cache_count')
        or report.get('external_formula_count')
    )
    return result
