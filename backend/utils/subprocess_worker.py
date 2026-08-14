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


def main():
    # 以 .env 文件为准重新加载配置（override=True 覆盖父进程继承的旧环境变量，
    # 防止父进程启动后改 .env 不生效导致子进程用旧配置运行）。
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except Exception:
        pass
    if len(sys.argv) < 2:
        print("@@ERROR@@subprocess_worker: 缺少参数文件")
        sys.exit(2)

    params_file = sys.argv[1]
    try:
        with open(params_file, "rb") as f:
            params = pickle.load(f)
        entry = params["entry"]
        args = params.get("args") or ()
        kwargs = params.get("kwargs") or {}

        # Aspose 许可证/运行时：子进程各自初始化（aspose_init.py 是唯一初始化点）。
        # 被调函数可能解析 xlsx（Aspose.Cells），沙箱脚本也可能 monkey-patch excel_parser。
        try:
            import aspose_init
            aspose_init.ensure_license()
        except Exception:
            pass  # 目标函数自身不用 Aspose 时不阻断

        module_path, _, func_name = entry.partition(":")
        if not func_name:
            print(f"@@ERROR@@subprocess_worker: entry 格式错误: {entry!r}（应为 module.path:func_name）")
            sys.exit(2)
        module = importlib.import_module(module_path)
        func = getattr(module, func_name)

        # 进度回调：kwargs 里的标记替换为真正的 stdout 包装（pickle 不能传函数）。
        # 注意用 pop 彻底移除标记键，否则会作为多余 kwarg 传给目标函数报
        # "unexpected keyword argument"。
        if kwargs.pop(_PROGRESS_MARK, None):
            kwargs["progress_cb"] = _progress_wrapper

        result = func(*args, **kwargs)

        # 结果写 pickle 临时文件，stdout 只回传路径
        fd, result_path = tempfile.mkstemp(suffix=".pkl", prefix="subproc_result_")
        try:
            with os.fdopen(fd, "wb") as f:
                pickle.dump(result, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception:
            # 结果不可 pickle（如含不可序列化对象）→ 报错而不是带病回传
            try:
                os.remove(result_path)
            except Exception:
                pass
            raise
        print(f"@@RESULT@@{result_path}")
        sys.stdout.flush()
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
