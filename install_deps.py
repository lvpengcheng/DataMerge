"""
安装依赖脚本 - 解决Windows编码问题
"""

import subprocess
import sys
import os

def install_dependencies():
    """安装依赖"""
    print("正在安装依赖...")

    # 基础依赖
    base_deps = [
        "fastapi==0.104.1",
        "uvicorn[standard]==0.24.0",
        "pydantic==2.5.0",
        "pandas==2.1.4",
        "openpyxl==3.1.2",
        "xlrd==2.0.1",
        "openai==1.3.0",
        "anthropic==0.18.0",
        "python-multipart==0.0.6",
        "python-dotenv==1.0.0",
        "redis==5.0.1",
        "celery==5.3.4",
    ]

    # 规则解析器依赖
    rule_parser_deps = [
        "PyPDF2==3.0.1",
        "python-docx==1.1.0",
        "requests==2.31.0",
    ]

    # 开发工具
    dev_deps = [
        "pytest==7.4.3",
        "pytest-asyncio==0.21.1",
        "pytest-cov==4.1.0",
        "black==23.11.0",
        "flake8==6.1.0",
    ]

    all_deps = base_deps + rule_parser_deps + dev_deps

    print(f"总共需要安装 {len(all_deps)} 个包")

    success_count = 0
    fail_count = 0

    for i, dep in enumerate(all_deps, 1):
        print(f"\n[{i}/{len(all_deps)}] 安装 {dep}...")
        try:
            # 使用subprocess安装
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", dep],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )

            if result.returncode == 0:
                print(f"  [成功] 安装成功")
                success_count += 1
            else:
                print(f"  [失败] 安装失败: {result.stderr[:100]}")
                fail_count += 1

        except Exception as e:
            print(f"  [失败] 安装异常: {str(e)}")
            fail_count += 1

    print(f"\n安装完成:")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")

    if fail_count == 0:
        print("\n所有依赖安装成功！")
        return True
    else:
        print(f"\n有 {fail_count} 个包安装失败，请手动安装。")
        return False

def check_dependencies():
    """检查依赖是否已安装"""
    print("检查依赖...")

    required_modules = [
        "fastapi", "uvicorn", "pydantic", "pandas", "openpyxl",
        "openai", "anthropic", "python_dotenv", "PyPDF2", "docx"
    ]

    missing = []

    for module in required_modules:
        try:
            __import__(module.replace("-", "_"))
            print(f"  [已安装] {module}")
        except ImportError:
            print(f"  [未安装] {module}")
            missing.append(module)

    if missing:
        print(f"\n缺少 {len(missing)} 个依赖: {', '.join(missing)}")
        return False
    else:
        print("\n所有依赖都已安装！")
        return True

def main():
    """主函数"""
    print("=" * 60)
    print("依赖安装工具")
    print("=" * 60)

    # 检查是否已安装
    if check_dependencies():
        print("\n所有依赖都已安装，无需重复安装。")
        return

    print("\n开始安装缺失的依赖...")

    # 安装依赖
    if install_dependencies():
        print("\n依赖安装完成！")

        # 再次检查
        print("\n验证安装结果...")
        check_dependencies()
    else:
        print("\n依赖安装失败，请检查错误信息。")
        sys.exit(1)

if __name__ == "__main__":
    main()