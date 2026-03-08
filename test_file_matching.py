"""
测试文件匹配逻辑
"""
import sys
import json
from pathlib import Path

from excel_parser import IntelligentExcelParser

# 读取活跃脚本配置
with open("tenants/rex/active_script.json", 'r', encoding='utf-8') as f:
    active_script = json.load(f)

script_info = active_script.get("script_info", {})
source_structure = script_info.get("source_structure", {})

print("=== 训练时的文件结构 ===")
expected_files = source_structure.get("files", {})
for filename, file_info in expected_files.items():
    print(f"\n文件: {filename}")
    sheets = file_info.get("sheets", {})
    for sheet_name, sheet_info in sheets.items():
        headers = sheet_info.get("headers", {})
        print(f"  Sheet: {sheet_name}")
        print(f"  表头: {list(headers.keys())}")

print("\n\n=== 附件文件结构 ===")
parser = IntelligentExcelParser()
temp_dir = Path("tenants/rex/calculations/202612/temp")

for attachment in temp_dir.glob("*.xlsx"):
    print(f"\n文件: {attachment.name}")
    try:
        parsed_sheets = parser.parse_excel_file(str(attachment))

        for sheet_data in parsed_sheets:
            sheet_name = sheet_data.sheet_name
            print(f"  Sheet: {sheet_name}")

            for region in sheet_data.regions:
                headers = region.head_data
                print(f"    Region表头: {list(headers.keys())}")

                # 尝试匹配
                print(f"\n    匹配测试:")
                for expected_filename, expected_file_info in expected_files.items():
                    expected_sheets = expected_file_info.get("sheets", {})

                    for expected_sheet_name, expected_sheet_info in expected_sheets.items():
                        expected_headers = set(expected_sheet_info.get("headers", {}).keys())
                        attach_headers = set(headers.keys())

                        if expected_headers and attach_headers:
                            intersection = expected_headers & attach_headers
                            score = len(intersection) / len(expected_headers)

                            if score > 0.5:  # 降低阈值以便看到更多匹配
                                print(f"      vs {expected_filename}/{expected_sheet_name}: {score:.2%}")
                                print(f"        期望表头: {expected_headers}")
                                print(f"        实际表头: {attach_headers}")
                                print(f"        交集: {intersection}")

    except Exception as e:
        print(f"  解析失败: {e}")
        import traceback
        traceback.print_exc()
