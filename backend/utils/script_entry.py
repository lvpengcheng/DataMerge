"""Shared calling convention for training and compute script main functions."""
import inspect
import os


def invoke_script_main(main, environment):
    from backend.utils.script_formats import track_script_formats, protect_explicit_values, preserve_template_formats
    from backend.utils.script_source_scope import selected_input_folder, filter_source_data, validate_source_files
    namespace = main.__globals__
    version2 = namespace.get('TEMPLATE_PIPELINE_VERSION') == 2
    selection = validate_source_files(namespace.get('SOURCE_FILES'))
    environment = dict(environment)
    original_input = namespace.get('input_folder')
    loader = namespace.get('load_source_data')
    schema_helper = namespace.get('apply_target_column_schemas')
    cached = namespace.get('_pre_loaded_source_data')
    with selected_input_folder(environment.get('input_folder', original_input), selection, cached) as folder:
        if selection is not None:
            reserved = namespace.get('_RESULT_SHEET_NAMES') or (namespace.get('_COL_MAP') or {}).keys()
            filtered_cache = (filter_source_data(cached, selection, reserved, folder)
                              if cached else cached)
            environment['input_folder'] = folder
            namespace['input_folder'] = folder
            if cached is not None:
                namespace['_pre_loaded_source_data'] = filtered_cache
            if callable(loader):
                def selected_loader(*args, **kwargs):
                    return filter_source_data(loader(*args, **kwargs), selection, reserved,
                                              environment.get('input_folder'))
                namespace['load_source_data'] = selected_loader
        try:
            with (preserve_template_formats() if version2 else track_script_formats()):
                if callable(schema_helper) and not version2:
                    namespace['apply_target_column_schemas'] = protect_explicit_values(schema_helper)
                try:
                    return _invoke_script_main(main, environment)
                finally:
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
                    namespace['_pre_loaded_source_data'] = cached


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
