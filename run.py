#!/usr/bin/env python3
"""
运行脚本 - 用于本地开发和测试
"""

import os
import sys
import logging
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def check_dependencies():
    """检查依赖"""
    required_packages = [
        'fastapi',
        'uvicorn',
        'pandas',
        'openpyxl',
        'pydantic'
    ]

    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"缺少依赖包: {', '.join(missing_packages)}")
        print("请运行: pip install -r requirements.txt")
        return False

    return True

def create_example_files():
    """创建示例文件"""
    example_dir = project_root / "examples"
    example_dir.mkdir(exist_ok=True)

    # 创建示例规则文件
    rule_file = example_dir / "rules.md"
    if not rule_file.exists():
        rule_file.write_text("""# 数据处理规则

## 数据源文件
1. source1.xlsx - 员工基本信息
2. source2.xlsx - 工资数据

## 计算规则
1. 基本工资 = 岗位工资 + 绩效工资
2. 应发工资 = 基本工资 + 津贴 - 扣款
3. 实发工资 = 应发工资 - 个税 - 社保

## 输出格式
1. 包含员工ID、姓名、部门
2. 包含各项工资明细
3. 包含合计行
""")

    # 创建示例Excel文件（使用pandas）
    try:
        import pandas as pd

        # 示例源文件1
        source1_data = {
            '员工ID': ['001', '002', '003'],
            '姓名': ['张三', '李四', '王五'],
            '部门': ['技术部', '市场部', '财务部'],
            '岗位工资': [10000, 8000, 9000],
            '绩效工资': [2000, 1500, 1800]
        }
        source1_df = pd.DataFrame(source1_data)
        source1_file = example_dir / "source1.xlsx"
        source1_df.to_excel(source1_file, index=False)

        # 示例源文件2
        source2_data = {
            '员工ID': ['001', '002', '003'],
            '津贴': [500, 300, 400],
            '扣款': [200, 150, 180],
            '个税': [1000, 800, 900],
            '社保': [800, 600, 700]
        }
        source2_df = pd.DataFrame(source2_data)
        source2_file = example_dir / "source2.xlsx"
        source2_df.to_excel(source2_file, index=False)

        # 示例预期结果
        expected_data = {
            '员工ID': ['001', '002', '003', '合计'],
            '姓名': ['张三', '李四', '王五', ''],
            '部门': ['技术部', '市场部', '财务部', ''],
            '基本工资': [12000, 9500, 10800, 32300],
            '应发工资': [12300, 9650, 11020, 32970],
            '实发工资': [10400, 8250, 9420, 28070]
        }
        expected_df = pd.DataFrame(expected_data)
        expected_file = example_dir / "expected_result.xlsx"
        expected_df.to_excel(expected_file, index=False)

        print(f"示例文件已创建到: {example_dir}")

    except ImportError:
        print("pandas未安装，跳过创建示例Excel文件")

def main():
    """主函数"""
    import argparse
    parser = argparse.ArgumentParser(description='AI驱动的Excel数据整合SaaS系统')
    parser.add_argument('--start', action='store_true', help='直接启动服务，不显示菜单')
    args = parser.parse_args()

    print("=" * 60)
    print("AI驱动的Excel数据整合SaaS系统")
    print("=" * 60)

    # 检查依赖
    if not check_dependencies():
        sys.exit(1)

    # 创建示例文件
    create_example_files()

    # 检查环境变量
    from dotenv import load_dotenv
    load_dotenv()

    ai_provider = os.getenv("AI_PROVIDER", "openai").lower()

    # 根据配置的AI提供者检查对应的API密钥
    needs_api_key = False
    provider_name = ""

    if ai_provider == "openai":
        needs_api_key = not os.getenv("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY") == "your_openai_api_key_here"
        provider_name = "OpenAI"
    elif ai_provider == "claude":
        needs_api_key = not os.getenv("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_API_KEY") == "your_anthropic_api_key_here"
        provider_name = "Claude"
    elif ai_provider == "deepseek":
        needs_api_key = not os.getenv("DEEPSEEK_API_KEY") or os.getenv("DEEPSEEK_API_KEY") == "your_deepseek_api_key_here"
        provider_name = "DeepSeek"
    elif ai_provider in ["ollama", "local"]:
        needs_api_key = False  # 本地提供者不需要API密钥
    else:
        needs_api_key = True
        provider_name = ai_provider

    if needs_api_key and ai_provider not in ["ollama", "local"]:
        print(f"\n[警告] 未设置{provider_name} API密钥")
        print(f"当前配置的AI提供者: {ai_provider}")
        print("请设置对应的API密钥:")

        if ai_provider == "openai":
            print("  - OPENAI_API_KEY (OpenAI API密钥)")
        elif ai_provider == "claude":
            print("  - ANTHROPIC_API_KEY (Claude API密钥)")
        elif ai_provider == "deepseek":
            print("  - DEEPSEEK_API_KEY (DeepSeek API密钥)")

        print("\n或者编辑.env文件配置:")
        print("  # 切换到本地提供者进行测试")
        print("  AI_PROVIDER=local")
        print("\n或者复制.env.example文件并配置:")
        print("  cp .env.example .env")
        print("  # 编辑.env文件设置API密钥")

    # 如果指定了--start参数，直接启动服务
    if args.start:
        print("\n启动FastAPI服务...")
        print("API文档地址: http://localhost:8000/docs")
        print("按 Ctrl+C 停止服务\n")

        import uvicorn

        # 配置日志，确保应用日志能输出
        log_config = uvicorn.config.LOGGING_CONFIG
        log_config["formatters"]["default"]["fmt"] = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        log_config["formatters"]["access"]["fmt"] = '%(asctime)s - %(levelname)s - %(client_addr)s - "%(request_line)s" %(status_code)s'

        # 添加根 logger 配置，确保所有应用日志都能输出
        log_config["loggers"][""] = {
            "handlers": ["default"],
            "level": "INFO",
            "propagate": False
        }

        # 添加 training logger 配置
        log_config["loggers"]["training"] = {
            "handlers": ["default"],
            "level": "INFO",
            "propagate": False
        }

        uvicorn.run(
            "backend.app.main:app",
            host="0.0.0.0",
            port=8000,
            reload=True,
            reload_excludes=["*tenants*", "*.xlsx", "*.xls", "*.log"],  # 排除训练产生的文件
            log_level="info",
            log_config=log_config,
            use_colors=True
        )
        return

    # 启动选项
    print("\n启动选项:")
    print("1. 启动FastAPI服务")
    print("2. 运行测试")
    print("3. 退出")

    choice = input("\n请选择 (1-3): ").strip()

    if choice == "1":
        print("\n启动FastAPI服务...")
        print("API文档地址: http://localhost:8000/docs")
        print("按 Ctrl+C 停止服务\n")

        import uvicorn
        uvicorn.run(
            "backend.app.main:app",
            host="0.0.0.0",
            port=8000,
            reload=True,
            reload_excludes=["*tenants*", "*.xlsx", "*.xls", "*.log"],  # 排除训练产生的文件
            log_level="info"
        )

    elif choice == "2":
        print("\n运行测试...")
        os.system("pytest tests/ -v")

    elif choice == "3":
        print("退出")
        sys.exit(0)

    else:
        print("无效选择")

if __name__ == "__main__":
    main()