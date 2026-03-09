"""测试完整的代码提取流程"""
import sys
sys.path.insert(0, '.')

from ai_engine.formula_code_generator import FormulaCodeGenerator

# 读取AI响应
with open('../tenants/rex104/training_logs/rex104_薪资计算规则_response_20260309_123811_01_generate.txt', 'r', encoding='utf-8') as f:
    ai_response = f.read()

# 创建代码生成器实例
generator = FormulaCodeGenerator(ai_provider=None)

# 调用提取方法
print("=" * 80)
print("测试 _extract_fill_result_sheets_function 方法")
print("=" * 80)

extracted_code = generator._extract_fill_result_sheets_function(ai_response)

print(f"\n提取的代码长度: {len(extracted_code)} 字符\n")

# 显示提取的代码的前30行
lines = extracted_code.split('\n')
print("提取的代码前30行:")
for i, line in enumerate(lines[:30]):
    spaces = len(line) - len(line.lstrip())
    print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")

# 检查clean_source_data函数
print("\n" + "=" * 80)
print("检查clean_source_data函数")
print("=" * 80)

if "def clean_source_data" in extracted_code:
    print("[OK] 找到clean_source_data函数")

    # 查找函数体
    start = extracted_code.find("def clean_source_data")
    end = extracted_code.find("\n\n\ndef fill_result_sheets", start)
    if end == -1:
        end = extracted_code.find("\ndef fill_result_sheets", start)

    if end > start:
        func_code = extracted_code[start:end]
        func_lines = func_code.split('\n')
        print(f"\nclean_source_data函数共{len(func_lines)}行:")
        for i, line in enumerate(func_lines):
            spaces = len(line) - len(line.lstrip())
            print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")
else:
    print("[NO] 未找到clean_source_data函数")
