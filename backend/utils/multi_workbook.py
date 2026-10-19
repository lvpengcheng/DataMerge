"""Multi-workbook contracts shared by training, generated scripts and delivery."""
import ast
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path


def workbook_definitions(structure):
    if isinstance(structure, str):
        structure = json.loads(structure)
    return (structure or {}).get("workbooks") or {}


def combine_structures(structures):
    """Keep workbook schemas separate; the union is only for reserved sheet names."""
    sheets = {}
    for structure in structures.values():
        sheets.update(structure.get("sheets") or {})
    return {"workbooks": structures, "sheets": sheets, "delivery": "zip"}


def script_targets(code):
    for node in ast.parse(code or "").body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == "MULTI_WORKBOOK_SCRIPTS" for t in node.targets):
            return ast.literal_eval(node.value)
    return {}


def build_script(targets):
    for name in targets:
        validate_name(name)
        compile(targets[name]["code"], name + ".py", "exec")
    return ('"""多个目标工作簿：按文件独立执行规则，全部成功后交付。"""\n'
            'from backend.utils.multi_workbook import execute_workbooks\n'
            'MULTI_WORKBOOK_SCRIPTS = ' + repr(targets) + '\n\n'
            'def main():\n'
            '    return execute_workbooks(MULTI_WORKBOOK_SCRIPTS, globals())\n')


def validate_name(name):
    if (not name or "/" in name or "\\" in name or name.startswith(("~", "_"))
            or Path(name).suffix.lower() != ".xlsx"):
        raise ValueError(f"交付文件名必须是无路径的 xlsx 文件名: {name}")


def generate_workbooks(generator, kwargs):
    structures = workbook_definitions(kwargs.get("expected_structure"))
    targets, responses, prompts = {}, [], []
    for name, structure in structures.items():
        validate_name(name)
        structure = dict(structure)
        structure["reserved_source_sheet_names"] = list((kwargs["expected_structure"].get("sheets") or {}).keys())
        options = dict(kwargs)
        options["expected_structure"] = structure
        options["rules_content"] = (kwargs.get("rules_content", "") +
            f"\n\n本次只生成交付文件【{name}】。请遵循规则中该文件的生成要求及公共规则。"
            "其他交付文件由平台分别生成，不要将其 sheet 混入本文件。")
        if "template_path" in options:
            options["template_path"] = structure.get("template_path") or structure.get("expected_file")
            selections = options.get("target_sheets")
            if isinstance(selections, dict):
                options["target_sheets"] = selections.get(name)
        code, response = generator.generate_code(**options)
        if not code:
            raise ValueError(f"{name}: 未生成有效脚本")
        targets[name] = {"code": code, "structure": structure}
        responses.append(f"### {name}\n{response or ''}")
        prompts.append(f"### {name}\n{getattr(generator, 'last_prompt', '') or ''}")
    generator.last_prompt = "\n\n".join(prompts)
    return build_script(targets), "\n\n".join(responses)


def correct_workbooks(generator, kwargs):
    targets = script_targets(kwargs["original_code"])
    for name, target in targets.items():
        options = dict(kwargs)
        options["original_code"] = target["code"]
        # Formula corrections rebuild the skeleton from generator state.
        generator._expected_structure = target["structure"]
        if "rules_content" in options:
            options["rules_content"] = (options.get("rules_content") or "") + f"\n只修正交付文件【{name}】。"
        code = generator.generate_correction_code(**options)
        if not code:
            raise ValueError(f"{name}: 修正未生成有效脚本")
        target["code"] = code
    return build_script(targets)


def resolve_child_template(code, tenant_id=None):
    from backend.utils.template_resolver import extract_template_ref, resolve_template_path
    name, _, baked = extract_template_ref(code)
    if not name and not baked:
        return None
    path = resolve_template_path(tenant_id=tenant_id, script_code=code)
    if not path:
        raise ValueError(f"找不到目标模板 {name or baked}，请同步该租户的模板文件")
    return path


def execute_workbooks(targets, environment):
    from backend.utils.script_entry import invoke_script_main
    from backend.utils.result_selection import pick_result_output
    output = Path(environment["output_folder"])
    output.mkdir(parents=True, exist_ok=True)
    # Isolated globals and folders prevent same-named functions/sheets/files colliding.
    with tempfile.TemporaryDirectory(prefix="multi_", dir=output) as staging:
        completed = []
        for index, (name, target) in enumerate(targets.items()):
            validate_name(name)
            child_dir = Path(staging) / str(index)
            child_dir.mkdir()
            env = {key: value for key, value in environment.items()
                   if key in ("input_folder", "salary_year", "salary_month", "monthly_standard_hours",
                              "tenant_id", "manual_headers", "file_passwords", "history_provider",
                              "report_progress", "SYNONYM_GROUPS", "find_column", "safe_get_column",
                              "convert_region_to_dataframe", "normalize_emp_code", "print_available_columns",
                              "load_files_to_dataframes", "EMPTY", "ZERO", "excel_text", "write_cell")}
            env.update(output_folder=str(child_dir), __name__="multi_workbook_child")
            overrides = environment.get('_template_override_paths') or {}
            template = overrides.get(name) or resolve_child_template(target["code"], environment.get("tenant_id"))
            if template:
                env["_template_override_path"] = template
            from .compute_templates import workbook_target_map
            manual_map = workbook_target_map(environment.get('_target_sheet_manual_map'), name)
            if manual_map:
                env['_target_sheet_manual_map'] = manual_map
            exec(compile(target["code"], name + ".py", "exec"), env)
            for key in ("salary_year", "salary_month", "monthly_standard_hours", "tenant_id"):
                if environment.get(key) is not None:
                    env[key] = environment[key]
            if template:
                env["TEMPLATE_PATH"] = template
            result_info = invoke_script_main(env["main"], env)
            if isinstance(result_info, dict) and result_info.get("success") is False:
                raise ValueError(f"{name}: {result_info.get('error') or '脚本报告执行失败'}")
            result = pick_result_output(child_dir.glob("*.xlsx"), template)
            if result is None:
                raise ValueError(f"{name}: 未生成结果")
            ready = Path(staging) / name
            shutil.copy2(result, ready)
            completed.append((ready, output / name))
            # Keep the fill report associated with its own workbook.
            report = child_dir / "fill_report.json"
            if report.is_file():
                staged_report = Path(staging) / f"{index}.json"
                shutil.copy2(report, staged_report)
                completed.append((staged_report, output / (name + ".fill_report.json")))
        for source, destination in completed:
            os.replace(source, destination)
    return {"output_files": list(targets)}


def precise_edit_workbooks(provider, code, feedback, **kwargs):
    from backend.ai_engine.precise_edit import run_precise_edit
    targets = script_targets(code)
    named = [name for name in targets if name in feedback]
    for name, target in targets.items():
        if named and name not in named:
            continue
        edited = run_precise_edit(provider, target["code"], feedback,
            **{**kwargs, "extra_context": kwargs.get("extra_context", "") + f"\n当前交付文件: {name}"})
        if not edited:
            return None
        target["code"] = edited
    return build_script(targets)


def select_outputs(files, structure, template_path=None):
    from backend.utils.result_selection import pick_result_output
    definitions = workbook_definitions(structure)
    if len(definitions) <= 1:
        result = pick_result_output(files, template_path)
        return [result] if result else []
    available = {Path(path).name: Path(path) for path in files if Path(path).is_file()}
    missing = [name for name in definitions if name not in available]
    if missing:
        raise ValueError("缺少交付文件: " + "、".join(missing))
    return [available[name] for name in definitions]


def create_zip(destination, files):
    """Publish an atomic ZIP containing only explicit deliverables, never inputs."""
    files = [Path(path) for path in files]
    names = [path.name.casefold() for path in files]
    if not files or len(set(names)) != len(names):
        raise ValueError("交付文件清单为空或存在重名")
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, staged = tempfile.mkstemp(suffix=".zip", dir=destination.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(staged, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for path in files:
                if not path.is_file():
                    raise ValueError(f"交付文件不存在: {path.name}")
                archive.write(path, arcname=path.name)
        os.replace(staged, destination)
    finally:
        if os.path.exists(staged):
            os.remove(staged)
    return destination


def validate_workbooks(output_dir, code, structure, rules_content=None, tenant_id=None):
    from backend.utils.excel_comparator import (
        compare_excel_files_multi_sheet, extract_primary_keys_from_rules, require_complete_comparison)
    from backend.utils.training_validation import finalize_training_output
    definitions = workbook_definitions(structure)
    targets = script_targets(code)
    results = select_outputs(Path(output_dir).glob("*.xlsx"), structure)
    details, reports, diffs, text = {}, {}, [], []
    total = matched = 0
    warning = False
    for index, result in enumerate(results):
        expected = definitions[result.name]
        child_code = targets.get(result.name, {}).get("code", code)
        template = resolve_child_template(child_code, tenant_id)
        finalized = finalize_training_output(str(result), child_code, expected, template)
        warning = warning or bool(finalized.get("formula_warning"))
        reports[result.name] = finalized.get("formula_report") or {}
        diff = Path(output_dir) / f"_diff_{index + 1}.xlsx"
        comparison = compare_excel_files_multi_sheet(
            str(result), expected["expected_file"], str(diff),
            primary_keys=extract_primary_keys_from_rules(rules_content) if rules_content else None,
            result_calculated=True, calculate_expected_formulas=True)
        require_complete_comparison(comparison)
        total += comparison.get("total_cells", 0)
        matched += comparison.get("matched_cells", 0)
        details[result.name] = comparison
        text.append(f"### 交付文件: {result.name}\n{comparison.get('detailed_text', '')}")
        if diff.exists():
            diffs.append(diff)
    return {"success": True, "accuracy": matched / total if total else 0,
            "total_cells": total, "matched_cells": matched,
            "total_differences": total - matched, "diff_details": details,
            "detailed_diff": "\n\n".join(text), "output_dir": str(output_dir),
            "output_files": [str(path) for path in results],
            "diff_files": [str(path) for path in diffs],
            "formula_warning": warning, "formula_report": {"workbooks": reports}}


def deliver_workbooks(output_dir, destination_dir, code, structure, tenant_id, display_name,
                      salary_year=None, salary_month=None, template_override_paths=None):
    """Finalize every declared workbook before publishing one ZIP."""
    from backend.utils.output_postprocess import (
        finalize_output_workbook, build_result_filename, dual_output_enabled,
        make_values_only_copy, values_only_name)
    from backend.utils.excel_comparator import inspect_formula_cache
    definitions = workbook_definitions(structure)
    targets = script_targets(code)
    outputs = select_outputs(Path(output_dir).glob("*.xlsx"), structure)
    destination_dir = Path(destination_dir)
    destination_dir.mkdir(parents=True, exist_ok=True)
    rows, reports, summaries, paths = 0, {}, {}, []
    values_failed = False
    package_name = Path(build_result_filename(display_name, salary_year, salary_month, ".zip"))
    import uuid
    package_name = package_name.with_stem(package_name.stem + "_" + uuid.uuid4().hex[:8])
    # A per-delivery staging folder avoids exposing a half-built result package.
    with tempfile.TemporaryDirectory(prefix="delivery_", dir=destination_dir) as staged:
        for output in outputs:
            child_code = targets.get(output.name, {}).get("code", code)
            template = (template_override_paths or {}).get(output.name) or resolve_child_template(child_code, tenant_id)
            result = finalize_output_workbook(str(output), template, definitions[output.name], child_code)
            report = inspect_formula_cache(str(output))
            reports[output.name] = report
            if (report.get("error_cache_count") or report.get("invalid_ref_formula_count")) and (
                    os.getenv("COMPUTE_STRICT_FORMULA_ERRORS", "false").lower() in ("true", "1", "yes", "on")):
                raise ValueError(f"{output.name}: 公式错误，未发布交付包")
            rows += result.get("rows_processed", 0)
            summaries[output.name] = result.get("sheet_summary")
            path = Path(staged) / output.name
            shutil.copy2(output, path)
            paths.append(path)
            if dual_output_enabled():
                value_path = Path(staged) / values_only_name(output.name)
                reserved = {name.casefold() for name in definitions} | {path.name.casefold() for path in paths}
                suffix = 2
                original_stem = value_path.stem
                while value_path.name.casefold() in reserved:
                    value_path = value_path.with_name(f"{original_stem}_{suffix}.xlsx")
                    suffix += 1
                if make_values_only_copy(str(path), str(value_path), "源_", None, template,
                                         fill_report_path=str(Path(output_dir) / (output.name + ".fill_report.json")),
                                         recalculate=False):
                    paths.append(value_path)
                else:
                    values_failed = True
        archive = create_zip(destination_dir / package_name, paths)
    return {"output_file": archive.name, "output_files": [archive.name],
            "business_files": [path.name for path in outputs], "delivery_type": "zip",
            "rows_processed": rows, "formula_report": {"workbooks": reports},
            "formula_warning": any(report.get(key) for report in reports.values() for key in (
                "empty_cache_count", "error_cache_count", "invalid_ref_formula_count", "external_formula_count")),
            "values_copy_failed": values_failed, "sheet_summaries": summaries}
