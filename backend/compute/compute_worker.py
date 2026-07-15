"""计算子进程入口。

智算的 run_compute_task 跑大量 Aspose(.NET) 运算，pythonnet 调用 .NET 时持有 GIL，
若在 API 进程内运行会冻结事件循环、导致 SSE/状态接口无响应被前置 WAF 判 502。
本模块让 run_compute_task 在【独立子进程】里运行：进度事件通过 stdout 带前缀回传，
父进程异步读管道再推给真正的 TaskLogBuffer。子进程被 GIL 冻结无所谓（它只算不服务）。

用法（由父进程 spawn）：
    python -u -m backend.compute.compute_worker <params.json>
params.json 含 run_compute_task 所需的可序列化参数。
"""

import os
import sys
import json
import asyncio
from pathlib import Path

# 事件行前缀：父进程据此从子进程 stdout 中区分"进度事件"与其它杂散输出
_EVT_PREFIX = "@@EVT@@"
_DONE = "@@DONE@@"

# 确保项目根在 sys.path（子进程独立启动）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))


class _StdoutBuffer:
    """伪装成 TaskLogBuffer：run_compute_task 调用的 push/finish/create_task。

    push 把事件 JSON 以 _EVT_PREFIX 开头单独成行写到 stdout，父进程读取后转推真缓冲。
    """

    def __init__(self):
        self._n = 0

    def create_task(self, task_id):  # run_compute_task 不调用，留作兼容
        return None

    def push(self, task_id, event_json) -> int:
        self._n += 1
        try:
            # 直接写 UTF-8 字节，绕过 Windows 默认 GBK stdout 编码，避免父进程按 UTF-8 解码时乱码
            sys.stdout.buffer.write((_EVT_PREFIX + str(event_json) + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
        except Exception:
            pass
        return self._n

    def finish(self, task_id):
        try:
            sys.stdout.buffer.write((_DONE + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
        except Exception:
            pass


def main():
    if len(sys.argv) < 2:
        sys.stderr.write("compute_worker: 缺少参数文件路径\n")
        sys.exit(2)

    params_file = sys.argv[1]
    try:
        with open(params_file, "r", encoding="utf-8") as f:
            p = json.load(f)
    except Exception as e:
        sys.stderr.write(f"compute_worker: 读取参数失败: {e}\n")
        sys.exit(2)

    task_id = str(p.get("task_id"))

    # 延迟导入：在子进程内加载应用（含 Aspose 初始化、DB 引擎），与 API 进程隔离
    try:
        import backend.app.main as _m
    except Exception as e:
        # 导入失败也要把错误回传给父进程，便于前端展示
        err = json.dumps({"type": "error", "message": f"计算子进程初始化失败: {e}"},
                         ensure_ascii=False)
        sys.stdout.buffer.write((_EVT_PREFIX + err + "\n").encode("utf-8"))
        sys.stdout.buffer.flush()
        sys.exit(1)

    buf = _StdoutBuffer()
    try:
        asyncio.run(_m.run_compute_task(
            task_id,
            buf,
            p.get("tenant_id"),
            p.get("script_id"),
            p.get("script_content"),
            Path(p.get("source_dir")),
            salary_year=p.get("salary_year"),
            salary_month=p.get("salary_month"),
            standard_hours=p.get("standard_hours"),
            file_passwords=p.get("file_passwords"),
            pre_validated_mapping=p.get("pre_validated_mapping"),
            precheck_auto_filled=p.get("precheck_auto_filled"),
            template_override_path=p.get("template_override_path"),
            target_sheet_manual_map=p.get("target_sheet_manual_map"),
        ))
    except Exception as e:
        err = json.dumps({"type": "error", "message": f"计算执行失败: {e}"},
                         ensure_ascii=False)
        sys.stdout.buffer.write((_EVT_PREFIX + err + "\n").encode("utf-8"))
        sys.stdout.buffer.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
