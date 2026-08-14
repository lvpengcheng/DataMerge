"""子进程 worker：subprocess_runner 的对端，在独立子进程里执行目标函数。

用法: python -m backend.utils.subprocess_worker <params_file>

params_file 是一个 pickle 文件，内容为:
    {"entry": "module.path:func_name", "args": tuple, "kwargs": dict}

协议（stdout，逐行）:
    @@PROG@@{json}      进度事件（kwargs 带 __SUBPROC_PROGRESS__ 标记时）
    @@RESULT@@{path}    成功：结果 pickle 文件路径（调用方负责删除）
    @@ERROR@@{b64}      失败：base64 编码的 traceback
    其他行              普通日志（父进程 debug 记录）

为什么结果走文件而不是 stdout: 结果可能含超大 DataFrame（几十万行），
stdout 管道传大 payload 有编码/缓冲风险；pickle 文件更稳。
"""

import base64
import importlib
import os
import pickle
import sys
import tempfile
import traceback

_PROGRESS_MARK = "__SUBPROC_PROGRESS__"   # 与 subprocess_runner 中一致


def _progress_wrapper(msg):
    """把进度消息打到 stdout，父进程按 @@PROG@@ 前缀转发给调用方。

    必须写 sys.__stdout__（真实管道）而非 sys.stdout：目标函数执行时可能被
    contextlib.redirect_stdout 重定向（沙箱把 sys.stdout 换成 StringIO），
    写 sys.stdout 会把 @@PROG@@ 标记吞进输出缓冲，父进程永远收不到进度。
    """
    if not isinstance(msg, str):
        try:
            msg = str(msg)
        except Exception:
            msg = repr(msg)
    msg = msg.replace("\n", " ")[:2000]
    _out = getattr(sys, "__stdout__", None) or sys.stdout
    try:
        _out.write(f"@@PROG@@{msg}\n")
        _out.flush()
    except Exception:
        pass


def _progress_wrapper_for(task_id: str):
    """带 task_id 的进度包装（daemon 模式每任务生成一个）。"""
    def _wrap(msg):
        if not isinstance(msg, str):
            try:
                msg = str(msg)
            except Exception:
                msg = repr(msg)
        msg = msg.replace("\n", " ")[:2000]
        _out = getattr(sys, "__stdout__", None) or sys.stdout
        try:
            _out.write(f"@@PROG@@{task_id}@@{msg}\n")
            _out.flush()
        except Exception:
            pass
    return _wrap


def _run_task(params: dict, task_id: str = "") -> None:
    """执行单个任务（成功 @@RESULT@@{task_id}@@{path}，失败 @@ERROR@@{task_id}@@{b64}）。

    任务级异常就地输出错误行，不向上抛（daemon 模式任务失败不能拖垮 worker）。
    """
    try:
        entry = params["entry"]
        args = params.get("args") or ()
        kwargs = params.get("kwargs") or {}

        # Aspose 许可证/运行时：daemon 模式已在启动时预加载，这里 ensure_license 幂等兜底。
        try:
            import aspose_init
            aspose_init.ensure_license()
        except Exception:
            pass  # 目标函数自身不用 Aspose 时不阻断

        module_path, _, func_name = entry.partition(":")
        if not func_name:
            sys.stdout.write(f"@@ERROR@@{task_id}@@{base64.b64encode(f'entry 格式错误: {entry!r}'.encode()).decode('ascii')}\n")
            return
        module = importlib.import_module(module_path)
        func = getattr(module, func_name)

        # 进度回调：kwargs 里的标记替换为 stdout 包装（pickle 不能传函数）。
        if kwargs.pop(_PROGRESS_MARK, None):
            kwargs["progress_cb"] = _progress_wrapper_for(task_id)

        result = func(*args, **kwargs)

        # 结果写 pickle 临时文件，stdout 只回传路径
        fd, result_path = tempfile.mkstemp(suffix=".pkl", prefix="subproc_result_")
        try:
            with os.fdopen(fd, "wb") as f:
                pickle.dump(result, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception:
            try:
                os.remove(result_path)
            except Exception:
                pass
            raise
        sys.stdout.write(f"@@RESULT@@{task_id}@@{result_path}\n")
    except SystemExit:
        raise
    except Exception:
        tb = traceback.format_exc()
        sys.stdout.write(
            f"@@ERROR@@{task_id}@@{base64.b64encode(tb.encode('utf-8', 'replace')).decode('ascii')}\n")
    finally:
        sys.stdout.flush()


def daemon_main():
    """常驻 worker 循环：预加载 Aspose/excel_parser，然后循环读 stdin 任务行。

    任务行 = 任务 pickle 文件路径（父进程写入）。串行处理（单个 worker），
    任务级异常不退出；父进程负责超时/内存监控与崩溃重启。
    启动完成后输出 @@READY@@ 供父进程等待就绪。
    """
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except Exception:
        pass
    # 预加载 Aspose + excel_parser：这是常驻 worker 省掉 20-30s 初始化开销的关键
    try:
        import aspose_init
        aspose_init.ensure_license()
    except Exception:
        pass
    try:
        import excel_parser  # noqa: F401
    except Exception:
        pass

    print("@@READY@@", flush=True)

    for line in sys.stdin:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        params_file = line
        task_id = ""
        try:
            with open(params_file, "rb") as f:
                params = pickle.load(f)
            task_id = str(params.get("task_id") or "")
            _run_task(params, task_id=task_id)
        except Exception as e:
            sys.stdout.write(
                f"@@ERROR@@{task_id}@@{base64.b64encode(f'daemon 任务异常: {e}'.encode()).decode('ascii')}\n")
            sys.stdout.flush()
        finally:
            try:
                os.remove(params_file)
            except Exception:
                pass


def main():
    # 以 .env 文件为准重新加载配置（override=True 覆盖父进程继承的旧环境变量，
    # 防止父进程启动后改 .env 不生效导致子进程用旧配置运行）。
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except Exception:
        pass
    if len(sys.argv) >= 2 and sys.argv[1] == "--daemon":
        daemon_main()
        return
    if len(sys.argv) < 2:
        print("@@ERROR@@subprocess_worker: 缺少参数文件")
        sys.exit(2)

    params_file = sys.argv[1]
    try:
        with open(params_file, "rb") as f:
            params = pickle.load(f)
        _run_task(params, task_id="")
        sys.exit(0)
    except SystemExit:
        raise
    except Exception:
        tb = traceback.format_exc()
        sys.stdout.write(f"@@ERROR@@{base64.b64encode(tb.encode('utf-8', 'replace')).decode('ascii')}\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
