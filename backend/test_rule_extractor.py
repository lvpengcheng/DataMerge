"""
测试规则提取器
"""

from ai_engine.rule_extractor import RuleExtractor
from pathlib import Path


def test_rule_extractor():
    """测试规则提取器"""

    # 读取达美乐规则文件
    rules_file = Path(r"c:\Users\Administrator\Desktop\training\rules\达美乐薪资计算规则.md")

    if not rules_file.exists():
        print(f"规则文件不存在: {rules_file}")
        return

    with open(rules_file, 'r', encoding='utf-8') as f:
        rules_content = f.read()

    # 创建规则提取器
    extractor = RuleExtractor()

    # 提取规则
    print("=" * 80)
    print("开始提取规则...")
    print("=" * 80)

    extracted_rules = extractor.extract_rules(rules_content)

    # 显示数据清洗规则
    print("\n" + "=" * 80)
    print("数据清洗规则:")
    print("=" * 80)
    for i, rule in enumerate(extracted_rules["data_cleaning_rules"], 1):
        print(f"\n{i}. {rule['original_text']}")
        print(f"   涉及表: {rule['tables']}")
        print(f"   条件: {rule['conditions']}")
        print(f"   动作: {rule['action']}")

    # 显示警告规则
    print("\n" + "=" * 80)
    print("警告规则:")
    print("=" * 80)
    for i, rule in enumerate(extracted_rules["warning_rules"], 1):
        print(f"\n{i}. {rule['original_text']}")
        print(f"   源表: {rule.get('source_table')}")
        print(f"   目标表: {rule.get('target_table')}")
        print(f"   字段: {rule.get('field')}")
        print(f"   类型: {rule.get('type')}")
        print(f"   提示信息: {rule.get('message')}")

    # 显示导入校验规则
    print("\n" + "=" * 80)
    print("导入校验规则:")
    print("=" * 80)
    for table_name, rules in extracted_rules["import_validation_rules"].items():
        print(f"\n{table_name}:")
        for rule in rules:
            print(f"  - {rule}")

    # 格式化为提示词
    print("\n" + "=" * 80)
    print("格式化后的提示词:")
    print("=" * 80)
    formatted = extractor.format_rules_for_prompt(extracted_rules)
    print(formatted)

    print("\n" + "=" * 80)
    print("测试完成!")
    print("=" * 80)


if __name__ == "__main__":
    test_rule_extractor()
