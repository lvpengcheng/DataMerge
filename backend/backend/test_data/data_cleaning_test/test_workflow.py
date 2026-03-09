"""
快速测试数据清洗和警告功能
"""

import sys
from pathlib import Path

# 添加backend到路径
current_dir = Path(__file__).parent
backend_path = current_dir.parent.parent
sys.path.insert(0, str(backend_path))

from ai_engine.rule_extractor import RuleExtractor
from ai_engine.prompt_generator import PromptGenerator

def test_data_cleaning_workflow():
    """测试数据清洗工作流程"""

    print("=" * 80)
    print("数据清洗和警告功能测试")
    print("=" * 80)

    # 1. 读取规则文件
    rules_file = Path("test_data/data_cleaning_test/rules/员工薪资计算规则.md")

    if not rules_file.exists():
        print(f"\n错误: 规则文件不存在: {rules_file}")
        print("请先运行 generate_test_data.py 生成测试数据")
        return

    print(f"\n1. 读取规则文件: {rules_file}")
    with open(rules_file, 'r', encoding='utf-8') as f:
        rules_content = f.read()

    print(f"   规则文件大小: {len(rules_content)} 字符")

    # 2. 提取规则
    print("\n2. 提取数据清洗和警告规则...")
    extractor = RuleExtractor()
    extracted_rules = extractor.extract_rules(rules_content)

    print(f"   - 数据清洗规则: {len(extracted_rules['data_cleaning_rules'])} 条")
    print(f"   - 警告规则: {len(extracted_rules['warning_rules'])} 条")
    print(f"   - 导入校验规则: {len(extracted_rules['import_validation_rules'])} 个表")

    # 3. 显示提取的规则
    print("\n3. 数据清洗规则详情:")
    for i, rule in enumerate(extracted_rules['data_cleaning_rules'], 1):
        if rule['original_text'] and not rule['original_text'].startswith('--'):
            print(f"   {i}. {rule['original_text'][:60]}...")

    print("\n4. 警告规则详情:")
    for i, rule in enumerate(extracted_rules['warning_rules'], 1):
        if rule['original_text'] and not rule['original_text'].startswith('*'):
            print(f"   {i}. {rule['original_text'][:60]}...")

    # 4. 格式化为提示词
    print("\n5. 格式化规则为AI提示词...")
    formatted_rules = extractor.format_rules_for_prompt(extracted_rules)
    print(f"   格式化后长度: {len(formatted_rules)} 字符")

    # 5. 测试集成到PromptGenerator
    print("\n6. 测试集成到PromptGenerator...")
    prompt_gen = PromptGenerator()

    # 模拟源文件结构
    source_structure = {
        "file_name": "测试输入",
        "sheets": [
            {
                "sheet_name": "员工信息",
                "regions": [
                    {
                        "headers": ["工号", "姓名", "身份证号", "部门", "职位", "基本工资"],
                        "sample_data": [
                            {"工号": "E001", "姓名": "张三", "基本工资": 8000}
                        ]
                    }
                ]
            }
        ]
    }

    # 模拟预期输出结构
    expected_structure = {
        "file_name": "员工薪资报表",
        "sheets": [
            {
                "sheet_name": "薪资报表",
                "regions": [
                    {
                        "headers": ["工号", "姓名", "部门", "基本工资", "实发工资"],
                        "sample_data": []
                    }
                ]
            }
        ]
    }

    # 生成训练提示词
    prompt = prompt_gen.generate_training_prompt(
        source_structure=source_structure,
        expected_structure=expected_structure,
        rules_content=rules_content,
        manual_headers=None
    )

    print(f"   生成的提示词长度: {len(prompt)} 字符")

    # 检查提示词中是否包含清洗和警告规则
    has_cleaning_rules = "数据清洗规则" in prompt
    has_warning_rules = "警告信息规则" in prompt
    has_return_warnings = "warnings" in prompt and "return" in prompt

    print(f"\n7. 验证提示词内容:")
    print(f"   - 包含数据清洗规则: {'✓' if has_cleaning_rules else '✗'}")
    print(f"   - 包含警告信息规则: {'✓' if has_warning_rules else '✗'}")
    print(f"   - 包含返回警告要求: {'✓' if has_return_warnings else '✗'}")

    # 8. 显示提示词片段
    print("\n8. 提示词中的数据清洗规则片段:")
    if "## 数据清洗规则" in prompt:
        start = prompt.find("## 数据清洗规则")
        end = prompt.find("\n\n##", start + 1)
        if end == -1:
            end = start + 500
        snippet = prompt[start:end]
        print("   " + snippet[:300].replace("\n", "\n   "))
        print("   ...")

    print("\n9. 提示词中的警告信息规则片段:")
    if "## 警告信息规则" in prompt:
        start = prompt.find("## 警告信息规则")
        end = prompt.find("\n\n##", start + 1)
        if end == -1:
            end = start + 500
        snippet = prompt[start:end]
        print("   " + snippet[:300].replace("\n", "\n   "))
        print("   ...")

    # 9. 总结
    print("\n" + "=" * 80)
    print("测试总结:")
    print("=" * 80)

    all_passed = has_cleaning_rules and has_warning_rules and has_return_warnings

    if all_passed:
        print("\n✓ 所有测试通过！")
        print("\n功能验证:")
        print("  1. 规则提取器成功提取数据清洗规则")
        print("  2. 规则提取器成功提取警告规则")
        print("  3. PromptGenerator成功集成规则")
        print("  4. 生成的提示词包含所有必要的规则和要求")
        print("\n下一步:")
        print("  - 使用训练引擎训练脚本")
        print("  - 验证AI生成的代码包含清洗和警告逻辑")
        print("  - 执行生成的代码并检查警告信息")
    else:
        print("\n✗ 部分测试失败")
        if not has_cleaning_rules:
            print("  - 提示词中缺少数据清洗规则")
        if not has_warning_rules:
            print("  - 提示词中缺少警告信息规则")
        if not has_return_warnings:
            print("  - 提示词中缺少返回警告的要求")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    test_data_cleaning_workflow()
