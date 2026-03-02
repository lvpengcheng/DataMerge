#!/usr/bin/env python3
"""
直接测试流式训练端点，避免浏览器CORS问题
"""

import requests
import json
import time
import sys
from pathlib import Path

def test_stream_endpoint_directly():
    """直接测试流式端点"""
    print("=" * 60)
    print("直接流式端点测试")
    print("=" * 60)

    base_url = "http://localhost:8000"
    tenant_id = "test"

    # 1. 首先检查服务是否运行
    print("\n1. 检查服务状态...")
    try:
        response = requests.get(f"{base_url}/docs", timeout=5)
        if response.status_code == 200:
            print("[OK] 服务正在运行")
        else:
            print(f"[ERROR] 服务异常: HTTP {response.status_code}")
            return False
    except Exception as e:
        print(f"[ERROR] 无法连接到服务: {e}")
        print("请启动服务: uvicorn backend.app.main:app --reload --host 0.0.0.0 --port 8000")
        return False

    # 2. 检查端点是否存在
    print("\n2. 检查流式训练端点...")
    try:
        # 发送一个简单的测试请求（不带文件）
        response = requests.post(
            f"{base_url}/api/{tenant_id}/train/stream",
            files={},  # 空文件
            data={'manual_headers': '{}'},
            timeout=10
        )

        print(f"响应状态码: {response.status_code}")

        if response.status_code == 422:
            print("[OK] 端点存在！422是预期的（缺少必要文件）")
            print(f"响应: {response.text[:200]}...")
            return True
        elif response.status_code == 404:
            print("[ERROR] 端点不存在 (404)")
            print("可能的原因:")
            print("  - 代码未重新加载")
            print("  - 路由未正确注册")
            print("  - 服务需要重启")
            return False
        else:
            print(f"响应: {response.status_code} - {response.text[:200]}")
            return True

    except requests.exceptions.RequestException as e:
        print(f"[ERROR] 请求失败: {e}")
        return False

def test_with_dummy_files():
    """使用虚拟文件测试"""
    print("\n" + "=" * 60)
    print("使用虚拟文件测试")
    print("=" * 60)

    base_url = "http://localhost:8000"
    tenant_id = "test"

    # 创建虚拟文件
    print("创建虚拟测试文件...")

    # 创建临时目录
    import tempfile
    import os
    temp_dir = tempfile.mkdtemp()

    try:
        # 创建虚拟规则文件
        rule_file = os.path.join(temp_dir, "test_rules.txt")
        with open(rule_file, 'w', encoding='utf-8') as f:
            f.write("测试规则文件\n这是一个测试用的规则文件")

        # 创建虚拟Excel文件（实际上创建文本文件，但命名为.xlsx）
        source_file = os.path.join(temp_dir, "test_source.xlsx")
        with open(source_file, 'w', encoding='utf-8') as f:
            f.write("这不是真正的Excel文件，仅用于测试")

        expected_file = os.path.join(temp_dir, "test_expected.xlsx")
        with open(expected_file, 'w', encoding='utf-8') as f:
            f.write("这不是真正的Excel文件，仅用于测试")

        print(f"虚拟文件创建在: {temp_dir}")

        # 准备请求
        files = [
            ('rule_files', ('test_rules.txt', open(rule_file, 'rb'), 'text/plain')),
            ('source_files', ('test_source.xlsx', open(source_file, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')),
            ('expected_result', ('test_expected.xlsx', open(expected_file, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
        ]

        data = {
            'manual_headers': json.dumps({"Sheet1": {"A1": "测试"}})
        }

        print("\n发送流式请求...")
        print(f"URL: {base_url}/api/{tenant_id}/train/stream")

        # 发送流式请求
        response = requests.post(
            f"{base_url}/api/{tenant_id}/train/stream",
            files=files,
            data=data,
            stream=True,
            headers={'Accept': 'text/event-stream', 'Cache-Control': 'no-cache'}
        )

        print(f"响应状态码: {response.status_code}")

        if response.status_code == 200:
            print("[OK] 流式连接成功建立！")
            print("\n开始接收流式数据...")
            print("-" * 60)

            # 读取流式响应
            line_count = 0
            for line in response.iter_lines():
                if line:
                    line = line.decode('utf-8')
                    line_count += 1

                    if line.startswith('data: '):
                        data_str = line[6:]
                        try:
                            event_data = json.loads(data_str)
                            event_type = event_data.get('type', 'unknown')

                            if event_type == 'start':
                                print(f"[START] {event_data['data']['message']}")
                            elif event_type == 'log':
                                log_data = event_data['data']
                                print(f"[{log_data['timestamp']}] [{log_data['level'].upper()}] {log_data['message']}")
                            elif event_type == 'result':
                                print("\n" + "=" * 60)
                                print("[COMPLETE] 训练完成!")
                                result_data = event_data['data']
                                print(f"状态: {result_data['status']}")
                                break
                            elif event_type == 'error':
                                print(f"[ERROR] 错误: {event_data['data']['error']}")
                                break
                            else:
                                print(f"[{event_type}] {data_str[:100]}...")

                        except json.JSONDecodeError:
                            print(f"[RAW] {data_str[:100]}...")

                    if line_count >= 50:  # 限制输出行数
                        print(f"\n已接收 {line_count} 行，停止接收...")
                        break

            print(f"\n总共接收 {line_count} 行数据")
            return True

        else:
            print(f"[ERROR] 请求失败: {response.status_code}")
            print(f"响应: {response.text[:500]}")
            return False

    except Exception as e:
        print(f"[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # 清理临时文件
        import shutil
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass

        # 关闭文件
        for file_info in files:
            if len(file_info) >= 2 and hasattr(file_info[1], 'close'):
                file_info[1].close()

def main():
    print("流式训练端点直接测试")
    print("=" * 60)

    # 测试1: 基本端点测试
    if not test_stream_endpoint_directly():
        print("\n❌ 基本测试失败，请检查服务状态")
        return

    # 测试2: 使用虚拟文件测试
    print("\n是否进行虚拟文件测试? (y/n)")
    choice = input().strip().lower()

    if choice == 'y':
        test_with_dummy_files()
    else:
        print("跳过虚拟文件测试")

    print("\n" + "=" * 60)
    print("测试完成!")
    print("=" * 60)

if __name__ == "__main__":
    main()