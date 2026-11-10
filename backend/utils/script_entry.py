"""Shared calling convention for training and compute script main functions."""
import inspect
import os
from pathlib import Path

_TEMPLATE_ARGUMENTS = ("_template_override_path", "template_path", "template_file")
_MISSING = object()


def invoke_script_main(main, environment):
    from backend.utils.script_formats import track_script_formats, protect_explicit_values, preserve_template_formats
    from backend.utils.script_source_scope import selected_input_folder, filter_source_data, validate_source_files
    namespace = main.__globals__
    version2 = namespace.get('TEMPLATE_PIPELINE_VERSION') == 2
    selection = validate_source_files(namespace.get('SOURCE_FILES'))
    environment = dict(environment)
    original_input = namespace.get('input_folder')
    template_path = environment.get('_template_override_path') or namespace.get('_template_override_path')
    template_bindings = {}
    if template_path:
        template = Path(template_path).resolve()
        if not template.is_file():
            raise ValueError(f'指定的模板文件不可用：{template.name}。请重新上传模板或检查智训模板绑定。')
        input_folder = environment.get('input_folder') or original_input
        if input_folder and template.is_relative_to(Path(input_folder).resolve()):
            raise ValueError(f'模板「{template.name}」位于源文件目录，请通过模板入口上传。')
        template_bindings = dict.fromkeys(_TEMPLATE_ARGUMENTS, str(template))
        if 'TEMPLATE_PATH' in namespace:
            template_bindings['TEMPLATE_PATH'] = str(template)
        environment.update(template_bindings)
    loader = namespace.get('load_source_data')
    schema_helper = namespace.get('apply_target_column_schemas')
    original_cache = namespace.get('_pre_loaded_source_data', _MISSING)
    cached = environment.get('_pre_loaded_source_data', namespace.get('_pre_loaded_source_data'))
    with selected_input_folder(environment.get('input_folder', original_input), selection, cached) as folder:
        from backend.utils.column_contract import (
            normalize_column_specs, normalize_source_keys)
        # 格式兼容模式不应让源键重新出现两套类型规则；保留类型须逐列声明 raw。
        specs = normalize_column_specs(namespace.get('_COLUMN_SPECS'))
        reserved = ()
        if selection is not None:
            reserved = namespace.get('_RESULT_SHEET_NAMES') or (namespace.get('_COL_MAP') or {}).keys()
        try:
            if selection is not None:
                environment['input_folder'] = folder
                namespace['input_folder'] = folder
            if cached is not None:
                selected_cache = filter_source_data(cached, selection, reserved, folder) if cached else cached
                normalized_cache = normalize_source_keys(selected_cache, specs)
                namespace['_pre_loaded_source_data'] = normalized_cache
            if callable(loader):
                def spec_loader(*args, **kwargs):
                    if cached:
                        print(f"[性能优化] 使用预加载源数据（{len(normalized_cache)}个sheet，跳过Excel解析）")
                        return normalized_cache
                    data = filter_source_data(loader(*args, **kwargs), selection, reserved, folder)
                    return normalize_source_keys(data, specs)
                namespace['load_source_data'] = spec_loader
            with (preserve_template_formats() if version2 else track_script_formats()):
                if callable(schema_helper) and not version2:
                    namespace['apply_target_column_schemas'] = protect_explicit_values(schema_helper)
                saved_template_bindings = {key: namespace.get(key, _MISSING) for key in template_bindings}
                namespace.update(template_bindings)
                try:
                    return _invoke_script_main(main, environment)
                finally:
                    for key, value in saved_template_bindings.items():
                        if value is _MISSING:
                            namespace.pop(key, None)
                        else:
                            namespace[key] = value
                    if callable(schema_helper):
                        namespace['apply_target_column_schemas'] = schema_helper
        finally:
            if selection is not None:
                if original_input is None:
                    namespace.pop('input_folder', None)
                else:
                    namespace['input_folder'] = original_input
            if callable(loader):
                namespace['load_source_data'] = loader
            if cached is not None:
                if original_cache is _MISSING:
                    namespace.pop('_pre_loaded_source_data', None)
                else:
                    namespace['_pre_loaded_source_data'] = original_cache


def _invoke_script_main(main, environment):
    from backend.utils.openpyxl_compat import ensure_custom_filter_compat
    ensure_custom_filter_compat()
    aliases = {
        'input_folder': 'input_folder', 'input_dir': 'input_folder', 'input_path': 'input_folder',
        'source_dir': 'input_folder', 'source_folder': 'input_folder',
        'output_folder': 'output_folder', 'output_dir': 'output_folder',
        'output_path': 'output_folder',
        'salary_year': 'salary_year', 'salary_month': 'salary_month',
        'monthly_standard_hours': 'monthly_standard_hours', 'standard_hours': 'monthly_standard_hours',
        **dict.fromkeys(_TEMPLATE_ARGUMENTS, '_template_override_path'),
    }
    signature = inspect.signature(main)
    positional, keywords = [], {}
    for name, parameter in signature.parameters.items():
        if parameter.kind in (parameter.VAR_POSITIONAL, parameter.VAR_KEYWORD):
            continue
        if name == 'output_file':
            value = os.path.join(environment['output_folder'], 'output.xlsx')
        elif name in aliases and aliases[name] in environment:
            value = environment[aliases[name]]
        elif parameter.default is not parameter.empty:
            value = parameter.default
        else:
            raise ValueError(f'脚本 main 的必填参数 {name!r} 不在支持的执行参数中')
        if parameter.kind == parameter.POSITIONAL_ONLY:
            positional.append(value)
        else:
            keywords[name] = value
    return main(*positional, **keywords)
