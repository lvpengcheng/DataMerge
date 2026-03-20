#!/usr/bin/env python3
"""
验证打包是否成功的测试脚本
"""

import sys
import subprocess
from pathlib import Path

def run_command(cmd, description):
    """运行命令并检查结果"""
    print(f"\n{'='*60}")
    print(f"测试: {description}")
    print(f"命令: {cmd}")
    print('='*60)

    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=30
        )

        if result.returncode == 0:
            print(f"✓ 成功")
            if result.stdout:
                print(result.stdout[:500])
            return True
        else:
            print(f"✗ 失败")
            print(f"错误: {result.stderr}")
            return False

    except subprocess.TimeoutExpired:
        print(f"✗ 超时")
        return False
    except Exception as e:
        print(f"✗ 异常: {e}")
        return False

def main():
    """主测试流程"""
    print("DataMerge 打包验证测试")
    print("="*60)

    tests = []

    # 1. 检查必要文件
    print("\n[1/6] 检查必要文件...")
    required_files = [
        'setup.py',
        'requirements.txt',
        'run.py',
        'MANIFEST.in',
        'Dockerfile',
        'backend/app/main.py',
        'frontend/templates/training.html',
        'frontend/templates/compute.html'
    ]

    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
            print(f"  ✗ 缺少: {file}")
        else:
            print(f"  ✓ 存在: {file}")

    tests.append(len(missing_files) == 0)

    # 2. 检查Python语法
    print("\n[2/6] 检查Python语法...")
    result = run_command(
        "python -m py_compile backend/app/main.py",
        "编译主应用文件"
    )
    tests.append(result)

    # 3. 检查依赖
    print("\n[3/6] 检查依赖安装...")
    result = run_command(
        "pip check",
        "检查依赖冲突"
    )
    tests.append(result)

    # 4. 测试导入
    print("\n[4/6] 测试模块导入...")
    test_imports = [
        "import fastapi",
        "import pandas",
        "import openpyxl",
        "from backend.app import main"
    ]

    import_success = True
    for imp in test_imports:
        try:
            exec(imp)
            print(f"  ✓ {imp}")
        except Exception as e:
            print(f"  ✗ {imp}: {e}")
            import_success = False

    tests.append(import_success)

    # 5. 构建测试
    print("\n[5/6] 测试构建...")
    result = run_command(
        "python -m build --outdir test_dist",
        "构建发布包"
    )
    tests.append(result)

    # 6. 检查构建产物
    print("\n[6/6] 检查构建产物...")
    if Path("test_dist").exists():
        files = list(Path("test_dist").glob("*"))
        if files:
            print(f"  ✓ 生成了 {len(files)} 个文件:")
            for f in files:
                print(f"    - {f.name}")
            tests.append(True)
        else:
            print("  ✗ 没有生成文件")
            tests.append(False)
    else:
        print("  ✗ test_dist 目录不存在")
        tests.append(False)

    # 总结
    print("\n" + "="*60)
    print("测试总结")
    print("="*60)

    passed = sum(tests)
    total = len(tests)

    print(f"通过: {passed}/{total}")

    if passed == total:
        print("\n✓ 所有测试通过！可以打包发布。")
        return 0
    else:
        print(f"\n✗ {total - passed} 个测试失败，请修复后再打包。")
        return 1

if __name__ == "__main__":
    sys.exit(main())
