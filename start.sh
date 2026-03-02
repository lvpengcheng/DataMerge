#!/bin/bash

echo ""
echo "========================================"
echo "  AI驱动的Excel数据整合SaaS系统"
echo "========================================"
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未找到Python3，请先安装Python 3.8+"
    echo "下载地址: https://www.python.org/downloads/"
    exit 1
fi

# 显示Python版本
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "[信息] 检测到Python版本: $PYTHON_VERSION"

# 检查是否在虚拟环境中
if [ -z "$VIRTUAL_ENV" ]; then
    echo "[警告] 未检测到虚拟环境，建议使用虚拟环境"
    echo ""
    echo "创建虚拟环境命令:"
    echo "  python3 -m venv venv"
    echo "  source venv/bin/activate"
    echo ""
    read -p "是否继续？(y/n): " USE_VENV
    if [[ ! "$USE_VENV" =~ ^[Yy]$ ]]; then
        echo "已取消"
        exit 0
    fi
fi

# 检查依赖
echo ""
echo "[步骤1] 检查依赖包..."
python3 -c "import pandas, openpyxl, fastapi, uvicorn, pydantic" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[信息] 缺少依赖包，正在安装..."

    # 检查是否有requirements.txt
    if [ -f "requirements.txt" ]; then
        pip install -r requirements.txt
    else
        echo "[警告] 未找到requirements.txt，安装基础依赖..."
        pip install pandas openpyxl fastapi uvicorn pydantic python-multipart
        pip install PyPDF2 python-docx  # 规则解析器依赖
        pip install openai anthropic    # AI API依赖（可选）
    fi

    if [ $? -ne 0 ]; then
        echo "[错误] 依赖安装失败"
        exit 1
    fi
    echo "[成功] 依赖安装完成"
else
    echo "[成功] 依赖检查通过"
fi

# 检查环境变量
echo ""
echo "[步骤2] 检查环境配置..."
if [ ! -f ".env" ]; then
    echo "[信息] 未找到.env文件，正在从.env.example创建..."
    if [ -f ".env.example" ]; then
        cp .env.example .env
        echo "[成功] 已创建.env文件，请编辑该文件配置API密钥"
        echo ""
        echo "需要配置的项:"
        echo "  1. OPENAI_API_KEY 或 ANTHROPIC_API_KEY"
        echo "  2. 其他可选配置"
        echo ""
        read -p "是否现在编辑.env文件？(y/n): " EDIT_NOW
        if [[ "$EDIT_NOW" =~ ^[Yy]$ ]]; then
            if command -v nano &> /dev/null; then
                nano .env
            elif command -v vim &> /dev/null; then
                vim .env
            else
                echo "请手动编辑 .env 文件"
            fi
        fi
    else
        echo "[错误] 未找到.env.example文件"
        exit 1
    fi
fi

# 检查AI API密钥
echo ""
echo "[步骤3] 检查AI API配置..."
python3 -c "
import os
has_openai = os.getenv('OPENAI_API_KEY', '') not in ['', 'your_openai_api_key_here']
has_anthropic = os.getenv('ANTHROPIC_API_KEY', '') not in ['', 'your_anthropic_api_key_here']
if not has_openai and not has_anthropic:
    print('NO_KEY')
else:
    if has_openai:
        print('OPENAI_OK')
    if has_anthropic:
        print('ANTHROPIC_OK')
"

if [ $? -eq 0 ] && [ -z "$(python3 -c "
import os
has_openai = os.getenv('OPENAI_API_KEY', '') not in ['', 'your_openai_api_key_here']
has_anthropic = os.getenv('ANTHROPIC_API_KEY', '') not in ['', 'your_anthropic_api_key_here']
if not has_openai and not has_anthropic:
    print('NO_KEY')
" 2>/dev/null)" ]; then
    echo "[成功] AI API配置检查通过"
else
    echo "[警告] 未配置AI API密钥，部分功能将受限"
    echo "请编辑.env文件设置:"
    echo "  OPENAI_API_KEY=你的OpenAI密钥"
    echo "  或"
    echo "  ANTHROPIC_API_KEY=你的Claude密钥"
    echo ""
    read -p "是否继续？(y/n): " CONTINUE_WITHOUT_AI
    if [[ ! "$CONTINUE_WITHOUT_AI" =~ ^[Yy]$ ]]; then
        echo "已取消"
        exit 0
    fi
fi

# 创建必要目录
echo ""
echo "[步骤4] 创建必要目录..."
mkdir -p data uploads examples logs
echo "[成功] 目录结构就绪"

# 启动选项
echo ""
echo "========================================"
echo "          启动选项"
echo "========================================"
echo "1. 启动FastAPI服务 (默认)"
echo "2. 运行测试"
echo "3. 运行规则解析器演示"
echo "4. 创建示例文件"
echo "5. 退出"
echo ""

read -p "请选择 (1-5，默认1): " CHOICE
CHOICE=${CHOICE:-1}

case $CHOICE in
    1)
        echo ""
        echo "========================================"
        echo "        启动FastAPI服务"
        echo "========================================"
        echo "服务地址: http://localhost:8000"
        echo "API文档: http://localhost:8000/docs"
        echo "按 Ctrl+C 停止服务"
        echo ""

        python3 run.py
        ;;

    2)
        echo ""
        echo "========================================"
        echo "          运行测试"
        echo "========================================"

        # 检查pytest
        python3 -c "import pytest" 2>/dev/null
        if [ $? -ne 0 ]; then
            echo "[信息] 安装pytest..."
            pip install pytest
        fi

        echo "运行测试套件..."
        python3 -m pytest tests/ -v
        ;;

    3)
        echo ""
        echo "========================================"
        echo "     运行规则解析器演示"
        echo "========================================"

        if [ -f "examples/rule_parser_demo.py" ]; then
            python3 examples/rule_parser_demo.py
        else
            echo "[错误] 未找到演示文件"
        fi
        ;;

    4)
        echo ""
        echo "========================================"
        echo "        创建示例文件"
        echo "========================================"

        python3 -c "
import sys
sys.path.insert(0, '.')
from run import create_example_files
create_example_files()
print('示例文件已创建到 examples/ 目录')
"
        ;;

    5)
        echo "退出"
        ;;

    *)
        echo "无效选择"
        ;;
esac

echo ""