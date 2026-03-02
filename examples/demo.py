#!/usr/bin/env python3
"""
演示脚本 - 展示系统基本功能
"""

import os
import sys
import tempfile
import pandas as pd
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def create_demo_files():
    """创建演示文件"""
    demo_dir = project_root / "demo_files"
    demo_dir.mkdir(exist_ok=True)

    print(f"创建演示文件到: {demo_dir}")

    # 1. 创建规则文件
    rule_file = demo_dir / "salary_rules.md"
    rule_content = """# 工资计算规则

## 数据源文件
1. employee_info.xlsx - 员工基本信息
2. salary_details.xlsx - 工资明细

## 计算规则
1. 基本工资 = 岗位工资 + 职级工资
2. 绩效工资 = 基本工资 × 绩效系数
3. 应发工资 = 基本工资 + 绩效工资 + 津贴
4. 实发工资 = 应发工资 - 社保 - 公积金 - 个税

## 输出要求
1. 包含员工ID、姓名、部门
2. 包含各项工资明细
3. 包含合计行
4. 格式与示例文件一致
"""
    rule_file.write_text(rule_content, encoding='utf-8')
    print(f"✓ 创建规则文件: {rule_file.name}")

    # 2. 创建员工信息文件
    employee_data = {
        '员工ID': ['E001', 'E002', 'E003', 'E004'],
        '姓名': ['张三', '李四', '王五', '赵六'],
        '部门': ['技术部', '市场部', '财务部', '人事部'],
        '岗位工资': [10000, 8000, 9000, 8500],
        '职级工资': [2000, 1500, 1800, 1600],
        '绩效系数': [1.2, 1.1, 1.3, 1.15]
    }
    employee_df = pd.DataFrame(employee_data)
    employee_file = demo_dir / "employee_info.xlsx"
    employee_df.to_excel(employee_file, index=False)
    print(f"✓ 创建员工信息文件: {employee_file.name}")

    # 3. 创建工资明细文件
    salary_data = {
        '员工ID': ['E001', 'E002', 'E003', 'E004'],
        '津贴': [500, 300, 400, 350],
        '社保': [800, 600, 700, 650],
        '公积金': [400, 300, 350, 320],
        '个税': [1200, 900, 1100, 950]
    }
    salary_df = pd.DataFrame(salary_data)
    salary_file = demo_dir / "salary_details.xlsx"
    salary_df.to_excel(salary_file, index=False)
    print(f"✓ 创建工资明细文件: {salary_file.name}")

    # 4. 创建预期结果文件
    # 计算预期结果
    expected_data = {
        '员工ID': ['E001', 'E002', 'E003', 'E004', '合计'],
        '姓名': ['张三', '李四', '王五', '赵六', ''],
        '部门': ['技术部', '市场部', '财务部', '人事部', ''],
        '基本工资': [12000, 9500, 10800, 10100, 42400],
        '绩效工资': [14400, 10450, 14040, 11615, 50505],
        '应发工资': [14900, 10750, 14440, 11965, 52055],
        '实发工资': [12500, 8950, 12290, 10045, 43785]
    }
    expected_df = pd.DataFrame(expected_data)
    expected_file = demo_dir / "expected_salary.xlsx"
    expected_df.to_excel(expected_file, index=False)
    print(f"✓ 创建预期结果文件: {expected_file.name}")

    return {
        'rule_file': str(rule_file),
        'source_files': [str(employee_file), str(salary_file)],
        'expected_file': str(expected_file)
    }

def test_excel_parser():
    """测试Excel解析器"""
    print("\n" + "="*60)
    print("测试Excel解析器")
    print("="*60)

    from backend.excel_parser import IntelligentExcelParser

    parser = IntelligentExcelParser()

    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        test_data = {
            '测试列1': ['值1', '值2', '值3'],
            '测试列2': [100, 200, 300],
            '测试列3': ['A', 'B', 'C']
        }
        df = pd.DataFrame(test_data)
        df.to_excel(f.name, index=False)

        # 解析文件
        results = parser.parse_excel_file(f.name)

        print(f"解析文件: {Path(f.name).name}")
        print(f"找到 {len(results)} 个Sheet")

        for sheet in results:
            print(f"  Sheet: {sheet.sheet_name}")
            print(f"    区域数量: {len(sheet.regions)}")

            for region in sheet.regions:
                print(f"    表头行: {region.head_row_start}-{region.head_row_end}")
                print(f"    数据行: {region.data_row_start}-{region.data_row_end}")
                print(f"    表头数量: {len(region.head_data)}")
                print(f"    数据行数: {len(region.data)}")

                if region.data:
                    print(f"    第一行数据: {region.data[0]}")

        os.unlink(f.name)

def test_document_validator():
    """测试文档验证器"""
    print("\n" + "="*60)
    print("测试文档验证器")
    print("="*60)

    from backend.document_validator import DocumentValidator
    from backend.excel_parser import IntelligentExcelParser

    validator = DocumentValidator()
    parser = IntelligentExcelParser()

    # 创建两个相同的文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            # 创建相同的数据
            data = {
                '列A': [1, 2, 3],
                '列B': ['X', 'Y', 'Z'],
                '列C': [100.5, 200.5, 300.5]
            }
            df = pd.DataFrame(data)
            df.to_excel(f1.name, index=False)
            df.to_excel(f2.name, index=False)

            # 提取模版
            parsed_data = parser.parse_excel_file(f1.name)
            template_schema = validator.extract_document_schema(parsed_data)

            print(f"提取文档模版:")
            print(f"  Sheet数量: {template_schema['total_sheets']}")
            print(f"  验证规则: {list(template_schema['validation_rules'].keys())}")

            # 验证相同文件
            is_valid, errors = validator.validate_file(f2.name, template_schema)
            print(f"验证相同文件: {'通过' if is_valid else '失败'}")
            if errors:
                for error in errors:
                    print(f"  错误: {error}")

            # 创建不同的文件并验证
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f3:
                different_data = {
                    '列A': [1, 2],  # 少一行
                    '列B': ['X', 'Y'],
                    '新列': [100, 200]  # 不同的列
                }
                df3 = pd.DataFrame(different_data)
                df3.to_excel(f3.name, index=False)

                is_valid, errors = validator.validate_file(f3.name, template_schema)
                print(f"验证不同文件: {'通过' if is_valid else '失败'}")
                if errors:
                    for error in errors:
                        print(f"  错误: {error}")

            os.unlink(f1.name)
            os.unlink(f2.name)
            os.unlink(f3.name)

def test_storage_manager():
    """测试存储管理器"""
    print("\n" + "="*60)
    print("测试存储管理器")
    print("="*60)

    from backend.storage.storage_manager import StorageManager

    with tempfile.TemporaryDirectory() as temp_dir:
        storage = StorageManager(base_dir=temp_dir)

        # 测试租户目录
        tenant_dir = storage.get_tenant_dir("demo_tenant")
        print(f"创建租户目录: {tenant_dir}")
        assert tenant_dir.exists()

        # 测试保存文件
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as test_file:
            test_file.write(b"Test content")
            test_file.flush()

            saved = storage.save_training_files(
                "demo_tenant",
                [test_file.name],
                [test_file.name],
                test_file.name
            )

            print(f"保存训练文件:")
            print(f"  规则文件: {len(saved['rules'])} 个")
            print(f"  源文件: {len(saved['source'])} 个")
            print(f"  预期文件: {saved['expected']}")

        # 测试存储统计
        stats = storage.get_storage_stats("demo_tenant")
        print(f"存储统计:")
        print(f"  总大小: {stats['total_size_human']}")
        print(f"  目录统计: {list(stats['directory_stats'].keys())}")
        print(f"  文件计数: {stats['file_counts']}")

def main():
    """主函数"""
    print("="*60)
    print("AI驱动的Excel数据整合系统 - 演示")
    print("="*60)

    # 检查依赖
    try:
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"缺少依赖: {e}")
        print("请运行: pip install pandas openpyxl")
        return

    # 创建演示文件
    demo_files = create_demo_files()

    # 运行测试
    test_excel_parser()
    test_document_validator()
    test_storage_manager()

    print("\n" + "="*60)
    print("演示完成!")
    print("="*60)
    print("\n下一步:")
    print("1. 配置AI API密钥 (在.env文件中设置OPENAI_API_KEY或ANTHROPIC_API_KEY)")
    print("2. 运行系统: python run.py")
    print("3. 访问API文档: http://localhost:8000/docs")
    print(f"\n演示文件已创建到: {project_root / 'demo_files'}")
    print("可以使用这些文件进行训练和测试。")

if __name__ == "__main__":
    main()