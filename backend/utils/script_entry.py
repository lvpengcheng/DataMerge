"""Shared calling convention for training and compute script main functions."""
import inspect
import os


def invoke_script_main(main, environment):
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
