"""
规则解析器使用示例
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from backend.rule_parser import RuleParserFactory, RuleSet
from backend.ai_rule_parser import AIRuleParser


class RuleParserDemo:
    """规则解析器演示"""

    def __init__(self):
        self.examples_dir = project_root / "examples"
        self.examples_dir.mkdir(exist_ok=True)

    def create_example_rule_file(self):
        """创建示例规则文件"""
        rule_content = """# 员工工资数据处理规则

## 预期输出文件
输出文件: 员工工资汇总_202401.xlsx

## 文件结构说明

### 1. 预期输出文件
文件: 员工工资汇总_202401.xlsx
工作表: 工资汇总
  列定义:
    - 员工编号: 员工信息.xlsx!基本信息!A列
    - 姓名: 员工信息.xlsx!基本信息!B列
    - 部门: 员工信息.xlsx!基本信息!C列
    - 基本工资: 工资明细.xlsx!1月工资!C列
    - 绩效奖金: =基本工资 * 绩效系数
    - 加班费: 考勤记录.xlsx!1月考勤!E列
    - 扣款: 考勤记录.xlsx!1月考勤!F列
    - 实发工资: =基本工资 + 绩效奖金 + 加班费 - 扣款

### 2. 源文件说明

文件: 员工信息.xlsx
工作表: 基本信息
  数据列:
    - A列: 员工编号
    - B列: 姓名
    - C列: 部门
    - D列: 入职日期

文件: 工资明细.xlsx
工作表: 1月工资
  数据列:
    - A列: 员工编号
    - B列: 姓名
    - C列: 基本工资
    - D列: 岗位津贴

文件: 考勤记录.xlsx
工作表: 1月考勤
  数据列:
    - A列: 员工编号
    - B列: 姓名
    - C列: 出勤天数
    - D列: 加班小时
    - E列: 加班费
    - F列: 扣款金额

## 数据映射规则
员工信息.xlsx!基本信息!A列 -> 员工工资汇总_202401.xlsx!工资汇总!员工编号
员工信息.xlsx!基本信息!B列 -> 员工工资汇总_202401.xlsx!工资汇总!姓名
员工信息.xlsx!基本信息!C列 -> 员工工资汇总_202401.xlsx!工资汇总!部门
工资明细.xlsx!1月工资!C列 -> 员工工资汇总_202401.xlsx!工资汇总!基本工资
考勤记录.xlsx!1月考勤!E列 -> 员工工资汇总_202401.xlsx!工资汇总!加班费
考勤记录.xlsx!1月考勤!F列 -> 员工工资汇总_202401.xlsx!工资汇总!扣款

## 计算规则
绩效奖金 = 基本工资 * 0.2  # 绩效系数为20%
实发工资 = 基本工资 + 绩效奖金 + 加班费 - 扣款

## 验证规则
1. 员工编号不能为空
2. 基本工资必须大于0
3. 实发工资不能为负数
"""

        # 保存为文本文件
        text_file = self.examples_dir / "salary_rules.txt"
        with open(text_file, 'w', encoding='utf-8') as f:
            f.write(rule_content)

        print(f"示例规则文件已创建: {text_file}")
        return text_file

    def demo_basic_parsing(self):
        """演示基础解析"""
        print("\n" + "=" * 60)
        print("基础规则解析演示")
        print("=" * 60)

        # 创建示例文件
        rule_file = self.create_example_rule_file()

        try:
            # 使用基础解析器
            print(f"\n1. 解析规则文件: {rule_file.name}")

            # 注意：文本文件需要使用合适的解析器，这里我们模拟PDF解析器
            from backend.rule_parser import PDFRuleParser
            parser = PDFRuleParser()

            # 由于是文本文件，我们直接使用文本解析逻辑
            with open(rule_file, 'r', encoding='utf-8') as f:
                rule_text = f.read()

            # 使用解析器的文本解析方法
            from backend.rule_parser import RuleParser
            base_parser = RuleParser()

            # 提取规则
            file_rules = base_parser._extract_file_rules(rule_text)
            mapping_rules = base_parser._extract_mapping_rules(rule_text)
            calculation_rules = base_parser._extract_calculation_rules(rule_text)

            print(f"2. 解析结果:")
            print(f"   - 找到 {len(file_rules)} 个文件规则")
            print(f"   - 找到 {len(mapping_rules)} 个映射规则")
            print(f"   - 找到 {len(calculation_rules)} 个计算规则")

            # 显示部分规则
            print(f"\n3. 规则示例:")
            if file_rules:
                print(f"   预期输出文件: {file_rules[0].file_name}")
                for sheet in file_rules[0].sheets[:2]:  # 只显示前2个工作表
                    print(f"   工作表: {sheet.sheet_name}")
                    for column in sheet.columns[:3]:  # 只显示前3列
                        print(f"     列: {column.column_name} -> {column.data_source}")

            if mapping_rules:
                print(f"\n   映射规则示例:")
                for i, (source, target) in enumerate(list(mapping_rules.items())[:3]):
                    print(f"     {source} -> {target}")

            if calculation_rules:
                print(f"\n   计算规则示例:")
                for i, (column, formula) in enumerate(list(calculation_rules.items())[:3]):
                    print(f"     {column} = {formula}")

            print("\nV 基础解析演示完成")

        except Exception as e:
            print(f"X 基础解析演示失败: {e}")

    def demo_ai_enhanced_parsing(self):
        """演示AI增强解析"""
        print("\n" + "=" * 60)
        print("AI增强规则解析演示")
        print("=" * 60)

        # 创建示例文件
        rule_file = self.create_example_rule_file()

        try:
            print(f"\n1. 使用AI增强解析规则文件: {rule_file.name}")

            # 使用本地AI提供者（模拟）
            parser = AIRuleParser(ai_provider_type="local")

            # 解析规则
            rule_set = parser.parse_with_ai(str(rule_file), use_ai_for_unclear=True)

            print(f"2. AI增强解析结果:")
            print(f"   预期输出文件: {rule_set.expected_file.file_name}")

            if rule_set.expected_file.sheets:
                print(f"   工作表数量: {len(rule_set.expected_file.sheets)}")
                for sheet in rule_set.expected_file.sheets:
                    print(f"     - {sheet.sheet_name}: {len(sheet.columns)} 列")

            print(f"   源文件数量: {len(rule_set.source_files)}")
            print(f"   映射规则数量: {len(rule_set.mapping_rules)}")
            print(f"   计算规则数量: {len(rule_set.calculation_rules)}")

            print(f"\n3. 生成的代码提示:")
            print("   基于解析的规则，可以生成以下类型的处理代码:")

            code_template = '''
import pandas as pd
import os
from excel_parser import IntelligentExcelParser

def process_salary_data():
    """处理工资数据"""
    # 初始化解析器
    parser = IntelligentExcelParser()

    # 1. 解析源文件
    employee_info = parser.parse_excel_file("员工信息.xlsx")
    salary_detail = parser.parse_excel_file("工资明细.xlsx")
    attendance = parser.parse_excel_file("考勤记录.xlsx")

    # 2. 构建输出数据结构
    output_data = {
        "员工编号": [],
        "姓名": [],
        "部门": [],
        "基本工资": [],
        "绩效奖金": [],
        "加班费": [],
        "扣款": [],
        "实发工资": []
    }

    # 3. 应用映射规则
    # ... 映射逻辑 ...

    # 4. 应用计算规则
    # ... 计算逻辑 ...

    # 5. 保存输出文件
    output_df = pd.DataFrame(output_data)
    output_df.to_excel("员工工资汇总_202401.xlsx", index=False)

    return True

if __name__ == "__main__":
    process_salary_data()
'''

            print(code_template)

            print("\nV AI增强解析演示完成")

        except Exception as e:
            print(f"X AI增强解析演示失败: {e}")

    def demo_integration_with_existing_system(self):
        """演示与现有系统的集成"""
        print("\n" + "=" * 60)
        print("与现有系统集成演示")
        print("=" * 60)

        print("\n1. 系统组件集成:")
        print("   - excel_parser.py: Excel智能解析器")
        print("   - rule_parser.py: 规则文件解析器")
        print("   - ai_rule_parser.py: AI增强解析器")
        print("   - ai_provider.py: AI引擎接口")
        print("   - prompt_generator.py: 提示词生成器")

        print("\n2. 数据处理流程:")
        print("   1. 用户上传规则文件（PDF/Word/Excel）")
        print("   2. 规则解析器提取数据处理规则")
        print("   3. AI引擎增强解析不明确的规则")
        print("   4. 生成数据处理代码")
        print("   5. 使用excel_parser.py处理实际数据")
        print("   6. 输出符合规则的结果文件")

        print("\n3. 代码生成示例:")
        integration_code = '''
# 集成示例代码
from backend.rule_parser import RuleParserFactory
from backend.ai_rule_parser import AIRuleParser
from excel_parser import IntelligentExcelParser

class DataProcessor:
    def __init__(self, rule_file_path: str):
        self.rule_file = rule_file_path
        self.excel_parser = IntelligentExcelParser()

    def process(self):
        # 1. 解析规则
        rule_parser = RuleParserFactory.create_parser(self.rule_file)
        rule_set = rule_parser.parse(self.rule_file)

        # 2. 使用AI增强（如果需要）
        if self._needs_ai_enhancement(rule_set):
            ai_parser = AIRuleParser()
            rule_set = ai_parser.parse_with_ai(self.rule_file)

        # 3. 生成处理代码
        code = self._generate_processing_code(rule_set)

        # 4. 执行处理
        return self._execute_code(code, rule_set)

    def _generate_processing_code(self, rule_set: RuleSet) -> str:
        # 基于规则集生成Python代码
        # 这里可以集成prompt_generator.py
        pass

    def _execute_code(self, code: str, rule_set: RuleSet) -> bool:
        # 执行生成的代码
        # 这里可以集成code_sandbox.py
        pass
'''

        print(integration_code)

        print("\nV 系统集成演示完成")

    def run_all_demos(self):
        """运行所有演示"""
        print("规则解析器系统演示")
        print("=" * 60)

        self.demo_basic_parsing()
        self.demo_ai_enhanced_parsing()
        self.demo_integration_with_existing_system()

        print("\n" + "=" * 60)
        print("演示完成！")
        print("=" * 60)
        print("\n总结:")
        print("1. 规则解析器支持PDF、Word、Excel格式的规则文件")
        print("2. AI增强解析可以处理不明确的规则")
        print("3. 系统可以与现有的excel_parser.py无缝集成")
        print("4. 支持生成完整的数据处理代码")


def main():
    """主函数"""
    demo = RuleParserDemo()
    demo.run_all_demos()


if __name__ == "__main__":
    main()