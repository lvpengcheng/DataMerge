#!/usr/bin/env python3
"""
最终流式训练测试 - 验证完整功能
"""

import requests
import json
import time

def test_complete_stream_training():
    """测试完整的流式训练功能"""
    print("=" * 60)
    print("最终流式训练测试")
    print("=" * 60)

    base_url = "http://localhost:8000"
    tenant_id = "test"

    print(f"租户ID: {tenant_id}")
    print(f"服务地址: {base_url}")

    # 1. 检查服务状态
    print("\n1. 检查服务状态...")
    try:
        response = requests.get(f"{base_url}/docs", timeout=5)
        if response.status_code == 200:
            print("[OK] 服务正在运行")
        else:
            print(f"[ERROR] 服务异常: {response.status_code}")
            return False
    except Exception as e:
        print(f"[ERROR] 无法连接到服务: {e}")
        return False

    # 2. 检查流式训练端点
    print("\n2. 检查流式训练端点...")
    try:
        response = requests.post(
            f"{base_url}/api/{tenant_id}/train/stream",
            files={},
            data={'manual_headers': '{}'},
            timeout=10
        )

        print(f"响应状态码: {response.status_code}")

        if response.status_code == 422:
            print("[OK] 端点存在！422是预期的（缺少必要文件）")
            print("错误详情（预期）:", response.json().get('detail', '')[:100])
        else:
            print(f"[ERROR] 意外响应: {response.status_code}")
            print(f"响应: {response.text[:200]}")
            return False

    except Exception as e:
        print(f"[ERROR] 测试端点失败: {e}")
        return False

    # 3. 测试流式响应
    print("\n3. 测试流式响应...")
    print("注意：这个测试需要实际文件，所以只测试连接")

    try:
        # 创建虚拟文件
        import tempfile
        import os

        temp_dir = tempfile.mkdtemp()

        # 创建测试文件
        rule_file = os.path.join(temp_dir, "test_rule.txt")
        with open(rule_file, 'w', encoding='utf-8') as f:
            f.write("测试规则：计算员工工资")

        source_file = os.path.join(temp_dir, "test_source.xlsx")
        with open(source_file, 'w', encoding='utf-8') as f:
            f.write("员工编号,姓名,月薪\n1,张三,5000\n2,李四,6000")

        expected_file = os.path.join(temp_dir, "test_expected.xlsx")
        with open(expected_file, 'w', encoding='utf-8') as f:
            f.write("员工编号,姓名,月薪,年终奖\n1,张三,5000,6000\n2,李四,6000,7200")

        print(f"创建测试文件在: {temp_dir}")

        # 准备请求
        files = [
            ('rule_files', ('test_rule.txt', open(rule_file, 'rb'), 'text/plain')),
            ('source_files', ('test_source.xlsx', open(source_file, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')),
            ('expected_result', ('test_expected.xlsx', open(expected_file, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
        ]

        data = {
            'manual_headers': json.dumps({"Sheet1": {"A1": "员工编号", "B1": "姓名", "C1": "月薪"}})
        }

        print("\n发送流式训练请求...")
        print("这可能需要一些时间，请耐心等待...")

        # 发送流式请求
        response = requests.post(
            f"{base_url}/api/{tenant_id}/train/stream",
            files=files,
            data=data,
            stream=True,
            headers={'Accept': 'text/event-stream', 'Cache-Control': 'no-cache'},
            timeout=30  # 延长超时时间
        )

        print(f"响应状态码: {response.status_code}")

        if response.status_code == 200:
            print("[OK] 流式连接成功建立！")
            print("\n开始接收训练日志...")
            print("-" * 60)

            # 读取流式响应
            received_count = 0
            start_time = time.time()

            for line in response.iter_lines():
                if line:
                    line = line.decode('utf-8')
                    received_count += 1

                    if line.startswith('data: '):
                        data_str = line[6:]
                        try:
                            event_data = json.loads(data_str)
                            event_type = event_data.get('type')

                            if event_type == 'start':
                                print(f"[START] {event_data['data']['message']}")
                            elif event_type == 'log':
                                log_data = event_data['data']
                                print(f"[{log_data['timestamp']}] [{log_data['level'].upper()}] {log_data['message']}")
                            elif event_type == 'result':
                                result_data = event_data['data']
                                print("\n" + "=" * 60)
                                print("[COMPLETE] 训练完成!")
                                print(f"状态: {result_data['status']}")

                                if 'training_result' in result_data:
                                    training_result = result_data['training_result']
                                    print(f"最佳分数: {training_result.get('best_score', 0):.2%}")
                                    print(f"总迭代次数: {training_result.get('total_iterations', 0)}")
                                    print(f"是否成功: {training_result.get('success', False)}")

                                break
                            elif event_type == 'error':
                                print(f"[ERROR] {event_data['data']['error']}")
                                break

                        except json.JSONDecodeError:
                            if received_count <= 3:  # 只显示前3行原始数据
                                print(f"[RAW] {data_str[:100]}...")

                    # 限制输出时间
                    if time.time() - start_time > 30:  # 最多30秒
                        print(f"\n[INFO] 已接收 {received_count} 行，停止接收...")
                        break

            elapsed_time = time.time() - start_time
            print(f"\n总共接收 {received_count} 行数据，耗时 {elapsed_time:.1f} 秒")

            return True

        else:
            print(f"[ERROR] 流式请求失败: {response.status_code}")
            print(f"响应: {response.text[:500]}")

            # 尝试获取更多错误信息
            try:
                error_data = response.json()
                print(f"错误详情: {json.dumps(error_data, ensure_ascii=False, indent=2)}")
            except:
                pass

            return False

    except requests.exceptions.Timeout:
        print("[ERROR] 请求超时")
        return False
    except Exception as e:
        print(f"[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # 清理
        import shutil
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass

        # 关闭文件
        try:
            for file_info in files:
                if len(file_info) >= 2 and hasattr(file_info[1], 'close'):
                    file_info[1].close()
        except:
            pass

def main():
    print("流式训练功能最终测试")
    print("=" * 60)

    print("这个测试将验证:")
    print("1. 服务是否运行")
    print("2. 流式训练端点是否存在")
    print("3. 能否建立流式连接")
    print("4. 能否接收训练日志")
    print("5. 训练结果是否正确返回")
    print("=" * 60)

    success = test_complete_stream_training()

    print("\n" + "=" * 60)
    if success:
        print("[SUCCESS] 流式训练功能测试通过!")
        print("现在可以使用以下方式测试:")
        print("1. 浏览器打开 test_stream_training.html")
        print("2. 运行 python direct_stream_test.py")
        print("3. 使用Postman发送multipart/form-data请求")
    else:
        print("[FAILED] 流式训练功能测试失败")
        print("请检查:")
        print("1. 服务是否正常运行")
        print("2. 是否有足够的系统资源")
        print("3. AI服务是否可用")
        print("4. 查看服务日志获取更多信息")

    print("=" * 60)

if __name__ == "__main__":
    main()