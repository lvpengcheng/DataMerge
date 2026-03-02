#!/usr/bin/env python3
"""
强制重启FastAPI服务
"""

import os
import sys
import subprocess
import time
import signal
import psutil

def kill_existing_servers():
    """杀死现有的uvicorn进程"""
    print("查找并停止现有的uvicorn进程...")

    killed = False
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            # 检查是否是uvicorn进程
            if proc.info['name'] and 'uvicorn' in proc.info['name'].lower():
                print(f"找到uvicorn进程 PID: {proc.info['pid']}")
                proc.kill()
                killed = True
                print(f"已停止进程 {proc.info['pid']}")

            # 检查是否是python进程运行我们的应用
            elif proc.info['cmdline'] and any('backend.app.main' in cmd for cmd in proc.info['cmdline'] if cmd):
                print(f"找到FastAPI应用进程 PID: {proc.info['pid']}")
                proc.kill()
                killed = True
                print(f"已停止进程 {proc.info['pid']}")

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    if killed:
        print("等待进程停止...")
        time.sleep(2)
    else:
        print("没有找到需要停止的进程")

def start_server():
    """启动新的服务器"""
    print("\n启动新的FastAPI服务器...")

    # 切换到项目目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # 启动命令
    cmd = [
        sys.executable,  # 使用当前Python解释器
        "-m", "uvicorn",
        "backend.app.main:app",
        "--reload",
        "--host", "0.0.0.0",
        "--port", "8000"
    ]

    print(f"执行命令: {' '.join(cmd)}")

    try:
        # 启动进程
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            bufsize=1
        )

        print("服务器启动中...")

        # 等待服务器启动
        for i in range(30):  # 最多等待30秒
            time.sleep(1)

            # 检查进程是否还在运行
            if process.poll() is not None:
                print("服务器进程已退出!")
                output = process.stdout.read()
                print(f"输出:\n{output}")
                return False

            # 尝试连接服务器
            try:
                import requests
                response = requests.get("http://localhost:8000/docs", timeout=1)
                if response.status_code == 200:
                    print(f"\n✅ 服务器启动成功!")
                    print(f"文档地址: http://localhost:8000/docs")
                    print(f"流式训练端点: http://localhost:8000/api/{{tenant_id}}/train/stream")

                    # 打印一些初始输出
                    print("\n服务器输出:")
                    try:
                        # 非阻塞读取
                        import select
                        if select.select([process.stdout], [], [], 0.1)[0]:
                            line = process.stdout.readline()
                            if line:
                                print(f"  {line.strip()}")
                    except:
                        pass

                    return True
            except:
                if i % 5 == 0:
                    print(f"等待服务器启动... ({i+1}/30秒)")

        print("❌ 服务器启动超时")
        return False

    except Exception as e:
        print(f"❌ 启动服务器时出错: {e}")
        return False

def test_endpoints():
    """测试端点"""
    print("\n测试API端点...")

    import requests
    import json

    base_url = "http://localhost:8000"

    # 测试文档
    try:
        response = requests.get(f"{base_url}/docs", timeout=5)
        if response.status_code == 200:
            print("✅ 文档端点正常")
        else:
            print(f"❌ 文档端点异常: {response.status_code}")
    except Exception as e:
        print(f"❌ 无法访问文档: {e}")
        return

    # 获取OpenAPI文档
    try:
        response = requests.get(f"{base_url}/openapi.json", timeout=5)
        if response.status_code == 200:
            data = response.json()
            paths = data.get('paths', {})

            print(f"✅ OpenAPI文档获取成功，共有 {len(paths)} 个路由")

            # 检查流式训练路由
            stream_found = False
            for path in paths:
                if '/train/stream' in path:
                    stream_found = True
                    print(f"✅ 找到流式训练路由: {path}")
                    break

            if not stream_found:
                print("❌ 未找到流式训练路由")
                print("当前路由:")
                for path in paths:
                    print(f"  - {path}")
        else:
            print(f"❌ 获取OpenAPI失败: {response.status_code}")
    except Exception as e:
        print(f"❌ 获取OpenAPI时出错: {e}")

def main():
    print("=" * 60)
    print("FastAPI服务强制重启工具")
    print("=" * 60)

    # 杀死现有进程
    kill_existing_servers()

    # 启动新服务
    if start_server():
        # 等待一下让服务器完全启动
        time.sleep(2)

        # 测试端点
        test_endpoints()

        print("\n" + "=" * 60)
        print("重启完成!")
        print("服务地址: http://localhost:8000")
        print("API文档: http://localhost:8000/docs")
        print("流式训练: http://localhost:8000/api/{tenant_id}/train/stream")
        print("=" * 60)

        print("\n按 Ctrl+C 停止服务器")
        try:
            # 保持脚本运行
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n\n正在停止服务器...")
            kill_existing_servers()
            print("服务器已停止")
    else:
        print("\n❌ 服务器启动失败")
        sys.exit(1)

if __name__ == "__main__":
    try:
        import psutil
        import requests
    except ImportError:
        print("安装依赖...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil", "requests"])
        import psutil
        import requests

    main()