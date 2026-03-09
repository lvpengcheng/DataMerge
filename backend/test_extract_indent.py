"""测试代码提取时的缩进问题"""
import re

# 读取AI响应
with open('../tenants/rex104/training_logs/rex104_薪资计算规则_response_20260309_123811_01_generate.txt', 'r', encoding='utf-8') as f:
    ai_response = f.read()

print("=" * 80)
print("1. 提取代码块")
print("=" * 80)

# 提取代码块
code_block_pattern = r'```python\s*(.*?)```'
matches = re.findall(code_block_pattern, ai_response, re.DOTALL)

if matches:
    code = matches[0]
    print(f"找到代码块，长度: {len(code)} 字符\n")

    # 显示前20行的缩进
    lines = code.split('\n')
    print("前20行的缩进:")
    for i, line in enumerate(lines[:20]):
        spaces = len(line) - len(line.lstrip())
        print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")

    print("\n" + "=" * 80)
    print("2. 提取clean_source_data函数")
    print("=" * 80)

    # 提取clean_source_data函数
    clean_start_idx = code.find("def clean_source_data")
    if clean_start_idx != -1:
        # 查找下一个顶级函数定义
        next_def_idx = code.find("\ndef ", clean_start_idx + 1)
        if next_def_idx > clean_start_idx:
            clean_end_idx = next_def_idx
        else:
            clean_end_idx = len(code)

        clean_function = code[clean_start_idx:clean_end_idx].strip()

        print(f"提取的clean_source_data函数长度: {len(clean_function)} 字符\n")

        # 显示提取的函数的缩进
        func_lines = clean_function.split('\n')
        print("提取的函数的缩进:")
        for i, line in enumerate(func_lines):
            spaces = len(line) - len(line.lstrip())
            print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")
else:
    print("未找到代码块")
