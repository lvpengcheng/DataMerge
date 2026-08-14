"""智能组表子进程入口。

引擎跑大量 Aspose(.NET) 运算，pythonnet 持 GIL 会冻结 API 事件循环（SSE 假死），
所以整个任务（解析/签名/AI生成/预扩展/执行/后处理）都在独立子进程内运行：
进度事件通过 stdout 带前缀回传，父进程异步读管道再转推 TaskLogBuffer。

用法（由父进程 spawn）：
    python -u -m backend.assemble.assemble_worker <params.json>
"""

import sys
import json
from pathlib import Path

_EVT_PREFIX = "@@EVT@@"
_DONE = "@@DONE@@"

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))


class _StdoutBuffer:
    """伪装成 push 回调：事件 JSON 以 @@EVT@@ 前缀单独成行写 stdout，父进程读取后转推真缓冲。"""

    def __init__(self, task_id: str):
        self.task_id = task_id
        self._n = 0

    def push(self, event_json) -> int:
        self._n += 1
        try:
            # 直接写 UTF-8 字节，绕过 Windows 默认 GBK stdout 编码，避免父进程按 UTF-8 解码时乱码
            sys.stdout.buffer.write(
                (_EVT_PREFIX + json.dumps(event_json, ensure_ascii=False) + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
        except Exception:
            pass
        return self._n

    def finish(self):
        try:
            sys.stdout.buffer.write((_DONE + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
        except Exception:
            pass


def main():
    # 以 .env 文件为准重新加载配置（override=True 覆盖父进程继承的旧环境变量）：
    # 父进程（uvicorn）启动后修改 .env 不会热重载，其 os.environ 里可能是旧值（如
    # CODE_SANDBOX_MAX_MEMORY=1024），worker 若继承旧值会导致内存上限/超时配置失效。
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except Exception:
        pass
    if len(sys.argv) < 2:
        sys.stderr.write("assemble_worker: 缺少参数文件路径\n")
        sys.exit(2)

    params_file = sys.argv[1]
    try:
        with open(params_file, "r", encoding="utf-8") as f:
            p = json.load(f)
    except Exception as e:
        sys.stderr.write(f"assemble_worker: 读取参数失败: {e}\n")
        sys.exit(2)

    task_id = int(p.get("task_id"))
    buf = _StdoutBuffer(str(task_id))

    try:
        from backend.assemble.assemble_engine import run_assemble_task
        run_assemble_task(task_id, buf.push, p)
    except Exception as e:
        err = json.dumps({"type": "error", "message": f"组表子进程失败: {e}"},
                         ensure_ascii=False)
        sys.stdout.buffer.write((_EVT_PREFIX + err + "\n").encode("utf-8"))
        sys.stdout.buffer.flush()
    finally:
        buf.finish()


if __name__ == "__main__":
    main()
