#!/usr/bin/env python3
"""
强制重新加载服务器 - 确保加载最新代码
"""

import os
import sys
import subprocess
import time
import signal

def kill_all_python_processes():
    """杀死所有相关的Python进程"""
    print("停止所有相关进程...")

    # 在Windows上使用taskkill
    if sys.platform == "win32":
        os.system("taskkill /F /IM uvicorn.exe 2>nul")
        os.system("taskkill /F /IM python.exe 2>nul")
    else:
        # 在Unix-like系统上
        os.system("pkill -f uvicorn 2>/dev/null")
        os.system("pkill -f 'python.*main.py' 2>/dev/null")

    print("等待进程停止...")
    time.sleep(3)

def clear_python_cache():
    """清除Python缓存"""
    print("清除Python缓存...")

    cache_dirs = [
        "__pycache__",
        "backend/__pycache__",
        "backend/app/__pycache__",
        "backend/ai_engine/__pycache__",
        ".pytest_cache"
    ]

    import shutil
    for cache_dir in cache_dirs:
        if os.path.exists(cache_dir):
            try:
                shutil.rmtree(cache_dir)
                print(f"  已清除: {cache_dir}")
            except:
                print(f"  清除失败: {cache_dir}")

def start_server():
    """启动服务器"""
    print("\n启动服务器...")

    # 切换到项目目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # 启动命令 - 使用明确的Python路径
    python_exe = sys.executable
    cmd = [
        python_exe,
        "-m", "uvicorn",
        "backend.app.main:app",
        "--reload",
        "--host", "0.0.0.0",
        "--port", "8000",
        "--log-level", "debug"  # 添加详细日志
    ]

    print(f"执行命令: {' '.join(cmd)}")

    # 启动进程
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=True,
        bufsize=1,
        encoding='utf-8'
    )

    print("服务器启动中...")

    # 读取输出并等待启动
    output_lines = []
    for i in range(30):  # 最多等待30秒
        # 检查进程是否退出
        if process.poll() is not None:
            print("服务器进程已退出!")
            output = process.stdout.read()
            print(f"输出:\n{output}")
            return False, None

        # 读取输出
        line = process.stdout.readline()
        if line:
            output_lines.append(line.strip())
            print(f"  {line.strip()}")

            # 检查启动成功消息
            if "Uvicorn running on" in line or "Application startup complete" in line:
                print("\n[OK] 服务器启动成功!")
                return True, process

        # 尝试连接
        try:
            import requests
            response = requests.get("http://localhost:8000/docs", timeout=1)
            if response.status_code == 200:
                print("\n[OK] 服务器启动成功!")
                return True, process
        except:
            pass

        time.sleep(1)

    print("\n[ERROR] 服务器启动超时")
    return False, process

def test_endpoints():
    """测试端点"""
    print("\n测试端点...")

    import requests

    # 测试文档
    try:
        response = requests.get("http://localhost:8000/docs", timeout=5)
        if response.status_code == 200:
            print("[OK] 文档端点正常")
        else:
            print(f"[ERROR] 文档端点异常: {response.status_code}")
            return False
    except Exception as e:
        print(f"[ERROR] 无法访问文档: {e}")
        return False

    # 获取OpenAPI文档
    try:
        response = requests.get("http://localhost:8000/openapi.json", timeout=5)
        if response.status_code == 200:
            import json
            data = response.json()
            paths = data.get('paths', {})

            print(f"[OK] OpenAPI文档获取成功，共有 {len(paths)} 个路由")

            # 检查流式训练路由
            stream_found = False
            for path in paths:
                if '/train/stream' in path:
                    stream_found = True
                    print(f"[OK] 找到流式训练路由: {path}")
                    break

            if not stream_found:
                print("[ERROR] 未找到流式训练路由")
                print("当前路由:")
                for path in paths:
                    print(f"  - {path}")
                return False

            return True

        else:
            print(f"[ERROR] 获取OpenAPI失败: {response.status_code}")
            return False

    except Exception as e:
        print(f"[ERROR] 获取OpenAPI时出错: {e}")
        return False

def main():
    print("=" * 60)
    print("强制重新加载服务器")
    print("=" * 60)

    # 1. 停止所有进程
    kill_all_python_processes()

    # 2. 清除缓存
    clear_python_cache()

    # 3. 启动服务器
    success, process = start_server()

    if not success:
        print("\n[ERROR] 服务器启动失败")
        if process:
            process.terminate()
        return

    # 4. 等待服务器完全启动
    time.sleep(2)

    # 5. 测试端点
    if test_endpoints():
        print("\n" + "=" * 60)
        print("[SUCCESS] 服务器重新加载成功!")
        print("服务地址: http://localhost:8000")
        print("API文档: http://localhost:8000/docs")
        print("流式训练: http://localhost:8000/api/{tenant_id}/train/stream")
        print("=" * 60)

        print("\n服务器正在运行，按 Ctrl+C 停止")

        try:
            # 保持脚本运行并显示服务器输出
            while True:
                if process.poll() is not None:
                    print("\n服务器进程已退出")
                    break

                line = process.stdout.readline()
                if line:
                    print(f"  {line.strip()}")

                time.sleep(0.1)

        except KeyboardInterrupt:
            print("\n\n正在停止服务器...")
            process.terminate()
            process.wait()
            print("服务器已停止")

    else:
        print("\n[ERROR] 端点测试失败")
        if process:
            process.terminate()

if __name__ == "__main__":
    try:
        import requests
    except ImportError:
        print("安装requests...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "requests"])
        import requests

    main()