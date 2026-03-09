"""
测试公式模式提示词是否包含数据清洗规则
"""

import sys
from pathlib import Path

# 添加backend到路径
sys.path.insert(0, str(Path(__file__).parent))

from ai_engine.prompt_generator import PromptGenerator

def test_formula_prompt_with_cleaning_rules():
    """测试公式模式提示词生成"""

    print("=" * 80)
    print("测试公式模式提示词生成（包含数据清洗规则）")
    print("=" * 80)

    # 读取测试规则文件
    rules_file = Path("test_data/data_cleaning_test/rules/员工薪资计算规则.md")

    if not rules_file.exists():
        print(f"\n错误: 规则文件不存在: {rules_file}")
        return

    with open(rules_file, 'r', encoding='utf-8') as f:
        rules_content = f.read()

    print(f"\n1. 读取规则文件: {rules_file}")
    print(f"   规则文件大小: {len(rules_content)} 字符")

    # 创建提示词生成器
    prompt_gen = PromptGenerator()

    # 模拟源数据结构
    source_structure = """
## 源数据结构
- **员工基础表_员工信息** (8列): 工号, 姓名, 身份证号, 部门, 职位, 基本工资, 入职日期, 员工状态
- **考勤表_考勤记录** (5列): 工号, 姓名, 出勤天数, 迟到次数, 请假天数
- **社保表_社保信息** (5列): 身份证号, 姓名, 社保基数, 个人缴纳, 公司缴纳
- **部门信息表_部门信息** (3列): 部门名称, 部门经理, 部门预算
"""

    # 模拟预期输出结构
    expected_structure = {
        "sheets": {
            "薪资报表": {
                "headers": {
                    "工号": {},
                    "姓名": {},
                    "身份证号": {},
                    "部门": {},
                    "职位": {},
                    "基本工资": {},
                    "出勤天数": {},
                    "迟到次数": {},
                    "请假天数": {},
                    "迟到扣款": {},
                    "请假扣款": {},
                    "应发工资": {},
                    "社保个人缴纳": {},
                    "实发工资": {}
                }
            }
        }
    }

    # 生成提示词
    print("\n2. 生成公式模式提示词...")
    prompt = prompt_gen.generate_formula_mode_prompt(
        source_structure=source_structure,
        expected_structure=expected_structure,
        rules_content=rules_content,
        manual_headers=None
    )

    print(f"   提示词长度: {len(prompt)} 字符")

    # 检查提示词内容
    print("\n3. 检查提示词内容:")

    has_cleaning_rules = "数据清洗规则" in prompt
    has_clean_function = "def clean_source_data" in prompt
    has_fill_function = "def fill_result_sheets" in prompt
    has_cleaning_instruction = "应用数据清洗规则" in prompt or "应用清洗规则" in prompt

    print(f"   - 包含数据清洗规则章节: {'[OK]' if has_cleaning_rules else '[NO]'}")
    print(f"   - 包含clean_source_data函数签名: {'[OK]' if has_clean_function else '[NO]'}")
    print(f"   - 包含fill_result_sheets函数签名: {'[OK]' if has_fill_function else '[NO]'}")
    print(f"   - 包含清洗指令: {'[OK]' if has_cleaning_instruction else '[NO]'}")

    # 显示数据清洗规则部分
    if has_cleaning_rules:
        print("\n4. 数据清洗规则部分:")
        start = prompt.find("## 数据清洗规则")
        if start != -1:
            end = prompt.find("\n## ", start + 1)
            if end == -1:
                end = start + 1000
            snippet = prompt[start:end]
            print(f"   找到数据清洗规则章节，长度: {len(snippet)} 字符")

    # 显示函数签名部分
    if has_clean_function:
        print("\n5. clean_source_data函数签名部分:")
        print("   找到clean_source_data函数定义")

    # 总结
    print("\n" + "=" * 80)
    print("测试总结:")
    print("=" * 80)

    all_passed = has_cleaning_rules and has_clean_function and has_fill_function and has_cleaning_instruction

    if all_passed:
        print("\n[OK] 所有检查通过！")
        print("\n提示词已正确包含:")
        print("  1. 数据清洗规则章节")
        print("  2. clean_source_data函数签名和说明")
        print("  3. fill_result_sheets函数签名")
        print("  4. 清洗指令和流程说明")
        print("\nAI应该能够生成包含数据清洗逻辑的代码。")
    else:
        print("\n[NO] 部分检查失败")
        if not has_cleaning_rules:
            print("  - 缺少数据清洗规则章节")
        if not has_clean_function:
            print("  - 缺少clean_source_data函数签名")
        if not has_fill_function:
            print("  - 缺少fill_result_sheets函数签名")
        if not has_cleaning_instruction:
            print("  - 缺少清洗指令")

    print("\n" + "=" * 80)

    # 保存提示词到文件供检查
    output_file = Path("test_data/data_cleaning_test/generated_prompt.txt")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(prompt)
    print(f"\n完整提示词已保存到: {output_file}")


if __name__ == "__main__":
    test_formula_prompt_with_cleaning_rules()
