#!/usr/bin/env python3
"""
检查服务器状态和路由
"""

import requests
import json
import sys

def check_server():
    """检查服务器状态"""
    base_url = "http://localhost:8000"

    print("检查服务器状态...")

    # 1. 检查服务器是否运行
    try:
        response = requests.get(f"{base_url}/docs", timeout=5)
        if response.status_code == 200:
            print("[OK] 服务器正在运行")
        else:
            print(f"[ERROR] 服务器返回状态码: {response.status_code}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] 无法连接到服务器: {e}")
        print("请确保服务已启动: uvicorn backend.app.main:app --reload --host 0.0.0.0 --port 8000")
        return False

    # 2. 获取OpenAPI文档
    try:
        response = requests.get(f"{base_url}/openapi.json", timeout=5)
        if response.status_code == 200:
            openapi_data = response.json()
            paths = openapi_data.get('paths', {})

            print(f"[OK] OpenAPI文档获取成功，共有 {len(paths)} 个路由")

            # 检查训练相关路由
            train_routes = []
            for path in paths:
                if '/train' in path:
                    train_routes.append(path)

            print(f"训练相关路由 ({len(train_routes)} 个):")
            for route in train_routes:
                print(f"  - {route}")

                # 检查流式训练路由
                if '/train/stream' in route:
                    print(f"    [FOUND] 找到流式训练路由!")

                    # 检查路由详情
                    route_info = paths[route]
                    if 'post' in route_info:
                        print(f"    Method: POST")
                        print(f"    Summary: {route_info['post'].get('summary', '无')}")
                        print(f"    Description: {route_info['post'].get('description', '无')[:100]}...")

            if not train_routes:
                print("[WARN] 没有找到训练相关路由")

        else:
            print(f"[ERROR] 获取OpenAPI失败: {response.status_code}")

    except Exception as e:
        print(f"[ERROR] 检查OpenAPI时出错: {e}")

    # 3. 直接测试流式训练端点
    print("\n直接测试流式训练端点...")
    try:
        # 发送一个简单的测试请求
        test_data = {
            "rule_files": [],
            "source_files": [],
            "expected_result": None,
            "manual_headers": "{}"
        }

        # 注意：这里只是测试端点是否存在，不期望成功执行
        response = requests.post(
            f"{base_url}/api/test/train/stream",
            files={},  # 空文件
            data=test_data,
            timeout=10
        )

        print(f"响应状态码: {response.status_code}")

        if response.status_code == 422:
            print("[OK] 端点存在！返回422表示参数验证失败（这是预期的，因为我们发送了空数据）")
            print("响应内容:", response.text[:200])
        elif response.status_code == 404:
            print("[ERROR] 端点不存在 (404)")
        else:
            print(f"[INFO] 端点返回状态码: {response.status_code}")
            print("响应内容:", response.text[:500])

    except requests.exceptions.RequestException as e:
        print(f"[ERROR] 测试端点时出错: {e}")

    return True

def check_code_version():
    """检查代码版本"""
    print("\n检查代码版本...")

    try:
        # 导入应用并检查
        import sys
        from pathlib import Path

        project_root = Path(__file__).parent
        sys.path.insert(0, str(project_root))

        from backend.app.main import app

        # 检查路由数量
        route_count = len([r for r in app.routes if hasattr(r, 'path')])
        print(f"[INFO] 应用中有 {route_count} 个路由")

        # 查找流式训练路由
        stream_routes = []
        for route in app.routes:
            if hasattr(route, 'path') and '/train/stream' in route.path:
                stream_routes.append(route.path)

        if stream_routes:
            print(f"[OK] 代码中找到流式训练路由: {stream_routes}")
        else:
            print("[ERROR] 代码中没有找到流式训练路由")

    except ImportError as e:
        print(f"[ERROR] 导入应用失败: {e}")
    except Exception as e:
        print(f"[ERROR] 检查代码时出错: {e}")

if __name__ == "__main__":
    print("=" * 60)
    print("服务器状态检查")
    print("=" * 60)

    # 检查服务器
    server_ok = check_server()

    # 检查代码版本
    check_code_version()

    print("\n" + "=" * 60)
    print("诊断建议:")

    if not server_ok:
        print("1. 服务器未运行，请启动服务:")
        print("   cd e:\\project\\DataMerge")
        print("   uvicorn backend.app.main:app --reload --host 0.0.0.0 --port 8000")
    else:
        print("1. 服务器正在运行")
        print("2. 如果路由不存在，可能是代码未重新加载")
        print("3. 尝试重启服务:")
        print("   a. 按 Ctrl+C 停止当前服务")
        print("   b. 重新运行启动命令")
        print("4. 或者使用我创建的 restart_server.bat")

    print("=" * 60)