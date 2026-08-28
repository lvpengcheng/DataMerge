#!/usr/bin/env python3
"""一键发布 DataMerge 到 Ubuntu 原生 systemd 环境（不使用 Docker）。"""

import argparse
import getpass
import json
import os
import shlex
import sys
import time
from pathlib import Path

from deploy import EXCLUDE_DIRS, build_tarball


ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = ROOT / "deploy.native.config.json"
EXAMPLE_CONFIG = ROOT / "deploy.native.config.example.json"


def fail(message: str, code: int = 1):
    print(f"[错误] {message}", file=sys.stderr)
    raise SystemExit(code)


def load_config(path: Path) -> dict:
    if not path.exists():
        if EXAMPLE_CONFIG.exists() and path == DEFAULT_CONFIG:
            path.write_text(EXAMPLE_CONFIG.read_text(encoding="utf-8"), encoding="utf-8")
            print(f"[首次使用] 已生成 {path.name}，请填写服务器信息后重新运行。")
            raise SystemExit(2)
        fail(f"找不到配置文件：{path}")
    try:
        cfg = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        fail(f"配置文件解析失败：{exc}")
    for key in ("host", "user", "remote_path"):
        if not cfg.get(key) or str(cfg[key]).startswith("你的"):
            fail(f"请在 {path.name} 中填写 {key}")
    cfg.setdefault("port", 22)
    cfg.setdefault("app_user", cfg["user"] if cfg["user"] != "root" else "datamerge")
    cfg.setdefault("app_port", 8000)
    cfg.setdefault("service_name", "datamerge")
    cfg.setdefault("memory_max", "6G")
    cfg.setdefault("memory_high", "5G")
    cfg.setdefault("memory_swap_max", "0")
    cfg.setdefault("cpu_quota", "200%")
    cfg.setdefault("tasks_max", 256)
    return cfg


def connect(cfg: dict):
    try:
        import paramiko
    except ImportError:
        fail("缺少 paramiko，请运行：python -m pip install paramiko")
    password = cfg.get("password") or os.getenv("DEPLOY_PASSWORD")
    key_file = cfg.get("key_file") or os.getenv("DEPLOY_KEY_FILE")
    if key_file:
        key_file = str((ROOT / key_file).resolve()) if not os.path.isabs(key_file) else key_file
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        client.connect(
            cfg["host"], port=int(cfg["port"]), username=cfg["user"],
            password=password, key_filename=key_file or None,
            allow_agent=True, look_for_keys=True, timeout=30,
        )
    except Exception as exc:
        if password or key_file:
            fail(f"SSH连接失败：{exc}")
        # 没有显式凭据且本机 SSH Agent/默认密钥不可用时，再交互式询问密码。
        password = getpass.getpass(f"SSH密钥登录失败，请输入 {cfg['user']}@{cfg['host']} 的密码：")
        try:
            client.connect(
                cfg["host"], port=int(cfg["port"]), username=cfg["user"],
                password=password, allow_agent=False, look_for_keys=False, timeout=30,
            )
        except Exception as retry_exc:
            fail(f"SSH连接失败：{retry_exc}")
    return client


def run_remote(ssh, command: str) -> int:
    _stdin, stdout, stderr = ssh.exec_command(command, get_pty=True)
    for line in iter(stdout.readline, ""):
        if line:
            print("  " + line, end="")
    rc = stdout.channel.recv_exit_status()
    err = stderr.read().decode("utf-8", "ignore")
    if rc and err:
        print(err, file=sys.stderr)
    return rc


def q(value) -> str:
    return shlex.quote(str(value))


def upload(sftp, local: Path, remote: str):
    print(f"上传：{local.name} → {remote}")
    sftp.put(str(local), remote)


def main():
    parser = argparse.ArgumentParser(description="一键发布到 Ubuntu 原生 systemd 环境")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG))
    parser.add_argument("--yes", action="store_true", help="跳过发布确认")
    parser.add_argument("--dry-run", action="store_true", help="只打包，不上传")
    parser.add_argument("--update-env", action="store_true", help="本次强制覆盖服务器 .env")
    args = parser.parse_args()

    cfg_path = Path(args.config).resolve()
    cfg = load_config(cfg_path)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    package = ROOT / "temp" / f"datamerge_native_{timestamp}.tar.gz"
    package.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 64)
    print("打包 Ubuntu 原生发布包（自动排除 .env、数据、日志和结果文件）")
    count = build_tarball(package, set(EXCLUDE_DIRS) | {"tmp"})
    print(f"已打包 {count} 个文件，{package.stat().st_size / 1024 / 1024:.1f} MB")
    if args.dry_run:
        print(f"发布包：{package}")
        return

    remote_path = str(cfg["remote_path"]).rstrip("/")
    print(f"目标：{cfg['user']}@{cfg['host']}:{cfg['port']} → {remote_path}")
    print(f"服务：{cfg['service_name']}，端口：{cfg['app_port']}，内存上限：{cfg['memory_max']}")
    if not args.yes and input("确认发布？输入 y 继续：").strip().lower() != "y":
        print("已取消")
        return

    ssh = connect(cfg)
    remote_package = f"/tmp/{package.name}"
    remote_env = f"/tmp/datamerge_env_{timestamp}"
    try:
        sftp = ssh.open_sftp()
        upload(sftp, package, remote_package)

        env_file = cfg.get("env_file")
        local_env = (ROOT / env_file).resolve() if env_file else None
        upload_env = bool(local_env and local_env.is_file())
        if args.update_env and not upload_env:
            fail("指定了 --update-env，但配置中的 env_file 不存在")
        if upload_env:
            upload(sftp, local_env, remote_env)
        sftp.close()

        # root 登录直接执行；普通用户要求已配置 passwordless sudo，避免把密码拼进远程命令。
        sudo = "" if cfg["user"] == "root" else "sudo -n "
        create_user = ""
        if cfg["app_user"] != "root":
            create_user = (
                f"id {q(cfg['app_user'])} >/dev/null 2>&1 || "
                f"{sudo}useradd --system --create-home --shell /usr/sbin/nologin {q(cfg['app_user'])}; "
            )

        env_install = ""
        if upload_env:
            condition = "true" if args.update_env else f"[ ! -f {q(remote_path + '/.env')} ]"
            env_install = (
                f"if {condition}; then {sudo}install -m 600 -o {q(cfg['app_user'])} "
                f"-g {q(cfg['app_user'])} {q(remote_env)} {q(remote_path + '/.env')}; fi; "
                f"rm -f {q(remote_env)}; "
            )

        deploy_env = {
            "APP_DIR": remote_path,
            "APP_USER": cfg["app_user"],
            "APP_PORT": cfg["app_port"],
            "SERVICE_NAME": cfg["service_name"],
            "MEMORY_MAX": cfg["memory_max"],
            "MEMORY_HIGH": cfg["memory_high"],
            "MEMORY_SWAP_MAX": cfg["memory_swap_max"],
            "CPU_QUOTA": cfg["cpu_quota"],
            "TASKS_MAX": cfg["tasks_max"],
        }
        if cfg.get("pip_index_url"):
            deploy_env["PIP_INDEX_URL"] = cfg["pip_index_url"]
        env_args = " ".join(f"{key}={q(value)}" for key, value in deploy_env.items())

        backup_dir = remote_path + "/_deploy_backups"
        remote_script = (
            "set -e; "
            f"{create_user}"
            f"{sudo}mkdir -p {q(remote_path)} {q(backup_dir)}; "
            f"if [ -f {q(remote_path + '/backend/app/main.py')} ]; then "
            f"{sudo}tar czf {q(backup_dir + '/code_' + timestamp + '.tar.gz')} "
            f"-C {q(remote_path)} --exclude=_deploy_backups --exclude=venv --exclude=.env "
            "--exclude=tenants --exclude=data --exclude=logs --exclude=output "
            "--exclude=compare_results --exclude=temp --exclude=global_assets . || true; fi; "
            f"{sudo}find {q(backup_dir)} -name 'code_*.tar.gz' -type f -printf '%T@ %p\n' "
            f"| sort -rn | tail -n +11 | cut -d' ' -f2- | xargs -r {sudo}rm -f; "
            f"{sudo}tar xzf {q(remote_package)} -C {q(remote_path)}; "
            f"rm -f {q(remote_package)}; "
            f"{env_install}"
            f"{sudo}chown -R {q(cfg['app_user'] + ':')} {q(remote_path)}; "
            f"cd {q(remote_path)}; "
            f"if {sudo}systemctl cat {q(cfg['service_name'])} >/dev/null 2>&1; then action=update; else action=install; fi; "
            f"{sudo}env {env_args} bash deploy_native_ubuntu.sh \"$action\""
        )

        print("远程备份、解压、安装依赖并重启服务……")
        rc = run_remote(ssh, remote_script)
        if rc:
            fail(f"远程发布失败，退出码 {rc}", rc)
        print("=" * 64)
        print(f"发布成功：http://{cfg['host']}:{cfg['app_port']}")
        print(f"查看日志：ssh {cfg['user']}@{cfg['host']} 'journalctl -u {cfg['service_name']} -f'")
    finally:
        ssh.close()
        try:
            package.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    main()
