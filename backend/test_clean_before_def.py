"""测试_clean_before_function_def方法"""
import sys
sys.path.insert(0, '.')

from ai_engine.formula_code_generator import FormulaCodeGenerator

# 读取AI响应
with open('../tenants/rex104/training_logs/rex104_薪资计算规则_response_20260309_123811_01_generate.txt', 'r', encoding='utf-8') as f:
    ai_response = f.read()

# 创建代码生成器实例
generator = FormulaCodeGenerator(ai_provider=None)

# 1. 提取函数
extracted_code = generator._extract_fill_result_sheets_function(ai_response)

print("=" * 80)
print("清理前的代码")
print("=" * 80)
lines = extracted_code.split('\n')
print(f"总行数: {len(lines)}")
print("\n前20行:")
for i, line in enumerate(lines[:20]):
    spaces = len(line) - len(line.lstrip())
    print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")

# 2. 清理
print("\n" + "=" * 80)
print("调用_clean_before_function_def清理")
print("=" * 80)

cleaned_code = generator._clean_before_function_def(extracted_code)

lines2 = cleaned_code.split('\n')
print(f"清理后总行数: {len(lines2)}")
print("\n清理后前20行:")
for i, line in enumerate(lines2[:20]):
    spaces = len(line) - len(line.lstrip())
    print(f"行{i+1:3d} (缩进{spaces:2d}): {line[:70]}")

# 3. 对比
print("\n" + "=" * 80)
print("对比分析")
print("=" * 80)
print(f"清理前行数: {len(lines)}")
print(f"清理后行数: {len(lines2)}")
print(f"删除了 {len(lines) - len(lines2)} 行")
