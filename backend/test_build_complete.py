"""测试完整的代码构建流程"""
import sys
sys.path.insert(0, '.')

from ai_engine.formula_code_generator import FormulaCodeGenerator

# 读取AI响应
with open('../tenants/rex104/training_logs/rex104_薪资计算规则_response_20260309_123811_01_generate.txt', 'r', encoding='utf-8') as f:
    ai_response = f.read()

# 创建代码生成器实例
generator = FormulaCodeGenerator(ai_provider=None)

# 1. 提取函数
print("=" * 80)
print("步骤1: 提取函数")
print("=" * 80)

extracted_code = generator._extract_fill_result_sheets_function(ai_response)
print(f"提取的代码长度: {len(extracted_code)} 字符")

# 显示提取的clean_source_data函数
lines = extracted_code.split('\n')
print("\n提取的clean_source_data函数前15行:")
for i, line in enumerate(lines[:15]):
    spaces = len(line) - len(line.lstrip())
    print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")

# 2. 构建完整代码
print("\n" + "=" * 80)
print("步骤2: 构建完整代码")
print("=" * 80)

complete_code = generator._build_complete_code(extracted_code)
print(f"完整代码长度: {len(complete_code)} 字符")

# 查找clean_source_data在完整代码中的位置
if "def clean_source_data" in complete_code:
    print("\n[OK] 完整代码中包含clean_source_data函数")

    # 找到函数位置
    start = complete_code.find("def clean_source_data")
    # 找到函数的前后各5行
    lines_before_start = complete_code[:start].split('\n')
    start_line_num = len(lines_before_start)

    all_lines = complete_code.split('\n')
    print(f"\nclean_source_data函数位置: 第{start_line_num}行")
    print("\n函数及其前后5行:")

    for i in range(max(0, start_line_num-6), min(len(all_lines), start_line_num+15)):
        line = all_lines[i]
        spaces = len(line) - len(line.lstrip())
        marker = " <-- 函数定义" if i == start_line_num - 1 else ""
        print(f"行{i+1:4d} (缩进{spaces:2d}): {line[:70]}{marker}")
else:
    print("\n[NO] 完整代码中未找到clean_source_data函数")
