"""Multi-workbook contracts shared by training, generated scripts and delivery."""
import ast
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from .multi_target_rules import rules_for_workbook


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


def _target_assignment(tree):
    return next((node for node in tree.body if isinstance(node, ast.Assign) and any(
        isinstance(t, ast.Name) and t.id == "MULTI_WORKBOOK_SCRIPTS" for t in node.targets)), None)


def _entry_block(node, entry):
    return (isinstance(node, ast.If) and isinstance(node.test, ast.Compare)
            and isinstance(node.test.left, ast.Name) and node.test.left.id == '__name__'
            and len(node.test.ops) == 1 and isinstance(node.test.ops[0], ast.Eq)
            and len(node.test.comparators) == 1
            and isinstance(node.test.comparators[0], ast.Constant)
            and node.test.comparators[0].value == entry)


def _section_indent(source, remove=False):
    """Move code into/out of a section without changing multiline string values."""
    import io
    import tokenize
    protected, fstrings = set(), []
    for token in tokenize.generate_tokens(io.StringIO(source).readline):
        kind = tokenize.tok_name[token.type]
        if token.type == tokenize.STRING:
            protected.update(range(token.start[0] + 1, token.end[0] + 1))
        elif kind == 'FSTRING_START':
            fstrings.append(token.start[0])
        elif kind == 'FSTRING_END':
            protected.update(range(fstrings.pop() + 1, token.end[0] + 1))
    lines = []
    for row, line in enumerate(source.splitlines(keepends=True), 1):
        if row not in protected:
            line = line[4:] if remove and line.startswith('    ') else line if remove else '    ' + line
        lines.append(line)
    return ''.join(lines)


def script_targets(code):
    tree = ast.parse(code or "")
    assignment = _target_assignment(tree)
    if assignment is None:
        return {}
    targets = ast.literal_eval(assignment.value)
    for name, target in targets.items():
        if 'entry' not in target:
            continue  # Legacy dictionary containing source strings.
        block = next((n for n in tree.body if _entry_block(n, target['entry'])), None)
        if block is None:
            raise ValueError(f'{name}: 缺少目标文件的代码段')
        source = ast.get_source_segment(code, block).split('\n', 1)[1]
        source = _section_indent(source + '\n', remove=True)
        future = target.get('future_imports') or []
        target['code'] = ('\n'.join(future) + '\n' if future else '') + source
        for key in ('entry', 'future_imports'):
            target.pop(key, None)
    return targets


def script_wrapper(code):
    """Inspect outer code without counting child template paths a second time."""
    tree = ast.parse(code or '')
    assignment = _target_assignment(tree)
    if assignment is None:
        return code
    entries = [t.get('entry') for t in ast.literal_eval(assignment.value).values() if t.get('entry')]
    tree.body = [n for n in tree.body if n is not assignment
                 and not any(_entry_block(n, entry) for entry in entries)]
    return ast.unparse(tree)


def build_script(targets, wrapper=None):
    """One Python module with visible executable code sections for all targets."""
    import textwrap
    import pprint
    metadata, sections, futures = {}, [], []
    for index, (name, target) in enumerate(targets.items(), 1):
        validate_name(name)
        source = target['code'].replace('\r\n', '\n').replace('\r', '\n')
        compile(source, name + '.py', 'exec')
        lines = source.splitlines(keepends=True)
        child_futures = []
        for node in ast.parse(source).body:
            if isinstance(node, ast.ImportFrom) and node.module == '__future__':
                declaration = ast.get_source_segment(source, node)
                child_futures.append(declaration)
                if declaration not in futures:
                    futures.append(declaration)
                for row in range(node.lineno - 1, node.end_lineno):
                    lines[row] = '\n'
        source = ''.join(lines).rstrip('\n') + '\n'
        entry = f'_dm_workbook_{index}'
        info = {key: value for key, value in target.items() if key != 'code'}
        info.update(entry=entry)
        if child_futures:
            info['future_imports'] = child_futures
        metadata[name] = info
        sections.append('# ' + '=' * 18 + ' ' + name + ' ' + '=' * 18 + '\n'
                        + 'if __name__ == ' + repr(entry) + ':\n'
                        + _section_indent(source))
    if wrapper is None:
        wrapper = ('def main():\n'
                   '    return execute_workbooks(MULTI_WORKBOOK_SCRIPTS, globals())\n')
    # Future imports are module-wide in a single Python file.
    outer = ast.parse(wrapper)
    for node in outer.body:
        if isinstance(node, ast.ImportFrom) and node.module == '__future__':
            declaration = ast.unparse(node)
            if declaration not in futures:
                futures.append(declaration)
    outer.body = [n for n in outer.body if not (isinstance(n, ast.ImportFrom) and n.module == '__future__')]
    wrapper = ast.unparse(outer) + '\n'
    data = 'MULTI_WORKBOOK_SCRIPTS = ' + pprint.pformat(metadata, width=100, sort_dicts=False) + '\n'
    return (('\n'.join(futures) + '\n' if futures else '')
            + '# 多目标文件脚本：各目标代码直接在下方编辑，统一由 main 生成全部文件。\n'
            + 'from backend.utils.multi_workbook import execute_workbooks\n'
            + 'import sys as _dm_sys\n'
            + '_DM_WORKBOOK_PROGRAM = _dm_sys._getframe().f_code\n\n'
            + data + '\n' + '\n'.join(sections) + '\n'
            + 'if __name__ not in ' + repr(tuple(info['entry'] for info in metadata.values())) + ':\n'
            + textwrap.indent(wrapper, '    '))


def readable_script(code):
    """Upgrade old dictionary bundles on download; new visible sections stay intact."""
    tree = ast.parse(code or '')
    assignment = _target_assignment(tree)
    if assignment is None:
        return code
    targets = ast.literal_eval(assignment.value)
    if all('entry' in target for target in targets.values()):
        return code
    return build_script(script_targets(code), wrapper=script_wrapper(code))


def script_download_response(code, filename, headers=None):
    from fastapi.responses import Response
    from urllib.parse import quote
    return Response(readable_script(code), media_type='text/x-python; charset=utf-8',
                    headers={**(headers or {}),
                             'Content-Disposition': "attachment; filename*=utf-8''" + quote(filename)})


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
        options["rules_content"] = (rules_for_workbook(kwargs.get("rules_content", ""), name) +
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
            options["rules_content"] = rules_for_workbook(options.get("rules_content"), name) + f"\n只修正交付文件【{name}】。"
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


def template_reference_code(code):
    """Stable template-only identity shared by visible sections and frozen packages."""
    from .template_resolver import extract_template_ref
    name, digest, path = extract_template_ref(code)
    return '\n'.join(key + ' = ' + repr(value) for key, value in (
        ('TEMPLATE_NAME', name), ('TEMPLATE_HASH', digest), ('TEMPLATE_PATH', path)) if value)


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
                              "load_files_to_dataframes", "EMPTY", "ZERO", "excel_text", "write_cell",
                              "_pre_loaded_source_data")}
            env.update(output_folder=str(child_dir), __name__=target.get('entry', 'multi_workbook_child'))
            overrides = environment.get('_template_override_paths') or {}
            template = overrides.get(name)
            if not target.get('entry'):
                template = template or resolve_child_template(target['code'], environment.get('tenant_id'))
            if template:
                env["_template_override_path"] = template
            from .compute_templates import workbook_target_map
            manual_map = workbook_target_map(environment.get('_target_sheet_manual_map'), name)
            if manual_map:
                env['_target_sheet_manual_map'] = manual_map
            if target.get('entry'):
                program = environment.get('_DM_WORKBOOK_PROGRAM')
                if program is None:
                    raise ValueError(f'{name}: 缺少多文件脚本的统一入口')
                exec(program, env)
            else:
                exec(compile(target["code"], name + ".py", "exec"), env)
            if target.get('entry'):
                # Read the actual edited declarations, never a stale metadata copy.
                ref_code = '\n'.join(key + ' = ' + repr(env[key]) for key in (
                    'TEMPLATE_NAME', 'TEMPLATE_HASH', 'TEMPLATE_PATH') if env.get(key))
                template = template or resolve_child_template(template_reference_code(ref_code), environment.get('tenant_id'))
                if template:
                    env['_template_override_path'] = template
            # Optional target-level configuration overrides the child declaration.
            if 'source_files' in target:
                env['SOURCE_FILES'] = target['source_files']
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
            **{**kwargs, "extra_context": rules_for_workbook(kwargs.get("extra_context", ""), name) + f"\n当前交付文件: {name}"})
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


def create_zip(destination, files, *, values_files=None):
    """Publish an atomic ZIP containing only explicit deliverables, never inputs."""
    files = [Path(path) for path in files]
    value_paths = {Path(path) for path in values_files} if values_files is not None else None
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
                folder = ('纯值版' if path in value_paths else '原版') if value_paths is not None else None
                archive.write(path, arcname=f'{folder}/{path.name}' if folder else path.name)
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
            primary_keys=extract_primary_keys_from_rules(rules_for_workbook(rules_content, result.name)) if rules_content else None,
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
    original_paths, value_paths = [], []
    values_failed = False
    package_name = Path(build_result_filename(display_name, salary_year, salary_month, ".zip"))
    from datetime import datetime
    package_name = package_name.with_stem(package_name.stem + datetime.now().strftime('%f'))
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
            original_paths.append(path)
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
                    value_paths.append(value_path)
                else:
                    values_failed = True
        # Never offer an incomplete values-only package as a successful delivery.
        if values_failed:
            value_paths = []
        originals = create_zip(destination_dir / package_name.with_stem(package_name.stem + '_原版'), original_paths)
        downloads = [originals.name]
        if value_paths:
            values = create_zip(destination_dir / package_name.with_stem(package_name.stem + '_纯值'), value_paths)
            downloads.append(values.name)
        archive = create_zip(destination_dir / package_name, original_paths + value_paths,
                             values_files=value_paths)
    return {"output_file": archive.name, "output_files": downloads,
            "business_files": [path.name for path in outputs], "delivery_type": "zip",
            "rows_processed": rows, "formula_report": {"workbooks": reports},
            "formula_warning": any(report.get(key) for report in reports.values() for key in (
                "empty_cache_count", "error_cache_count", "invalid_ref_formula_count", "external_formula_count")),
            "values_copy_failed": values_failed, "sheet_summaries": summaries}
