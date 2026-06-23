#!/usr/bin/env python3
"""
一键发布到 Docker 服务器：打包本地代码 → SFTP 上传 → 远程 docker compose up -d --build。

用法：
  1) 安装依赖（仅本机需要，不进镜像）：  pip install paramiko
  2) 配置连接信息（三选一，优先级 命令行 > 环境变量 > deploy.config.json）：
       - 复制 deploy.config.example.json 为 deploy.config.json 填好（已被 .gitignore 忽略）
       - 或设环境变量 DEPLOY_HOST / DEPLOY_PORT / DEPLOY_USER / DEPLOY_PASSWORD / DEPLOY_REMOTE_PATH
       - 或命令行：python deploy.py --host x --user root --remote-path /opt/datamerge
  3) 运行：  python deploy.py            （会先列出要做的事并要求确认）
            python deploy.py --yes      （跳过确认）
            python deploy.py --dry-run  （只打包看看包含哪些文件，不上传）

安全保证：
  - 绝不上传 tenants/ data/ logs/ output/ .env 等（保留服务器上的数据与密钥，不会被覆盖）
  - 远程是"解压覆盖"到 REMOTE_PATH，不删库不动卷挂载目录
  - 密码不写进任何提交的文件；可用环境变量或运行时输入
"""

import os
import sys
import json
import time
import fnmatch
import tarfile
import getpass
import argparse
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

# 绝不上传：运行时数据 / 卷挂载 / 密钥 / 本地杂物（覆盖会毁掉生产数据）
# 注意 global_assets/ 也在 docker-compose 里卷挂载，属运行时数据，不随发布覆盖。
EXCLUDE_DIRS = {
    ".git", ".venv", "venv", "env", "ENV", "__pycache__", "node_modules",
    "tenants", "data", "logs", "output", "compare_results", "temp",
    "global_assets", "releases", ".claude", ".idea", ".vscode",
    ".pytest_cache", "htmlcov",
}
# --code-only 时额外排除的"大而少变"目录（首次全量发布后，服务器已有，无需重传）
CODE_ONLY_EXTRA_DIRS = {"libs", "fonts"}
EXCLUDE_PATTERNS = [
    "*.pyc", "*.pyo", "*.pyd", "*.log", "*.tmp", "*.swp",
    ".env", ".env.*", "data.db", "*.db-journal",
    "deploy.config.json", "deploy.py",  # 脚本与配置不必上服务器
]


def _excluded(rel_path: str, exclude_dirs: set) -> bool:
    parts = Path(rel_path).parts
    if any(p in exclude_dirs for p in parts):
        return True
    name = Path(rel_path).name
    return any(fnmatch.fnmatch(name, pat) for pat in EXCLUDE_PATTERNS)


def build_tarball(out_path: Path, exclude_dirs: set) -> int:
    """打包 PROJECT_ROOT 下未被排除的文件到 out_path，返回文件数。"""
    count = 0
    with tarfile.open(out_path, "w:gz") as tar:
        for root, dirs, files in os.walk(PROJECT_ROOT):
            rel_root = os.path.relpath(root, PROJECT_ROOT)
            dirs[:] = [d for d in dirs if not _excluded(os.path.join(rel_root, d) if rel_root != "." else d, exclude_dirs)]
            for f in files:
                rel = os.path.normpath(os.path.join(rel_root, f)) if rel_root != "." else f
                if _excluded(rel, exclude_dirs):
                    continue
                tar.add(os.path.join(root, f), arcname=rel.replace("\\", "/"))
                count += 1
    return count


def load_config(args) -> dict:
    cfg = {}
    cfg_file = PROJECT_ROOT / "deploy.config.json"
    if cfg_file.exists():
        try:
            cfg.update(json.loads(cfg_file.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"[警告] 读取 deploy.config.json 失败：{e}")
    # 环境变量覆盖
    for k_env, k in [("DEPLOY_HOST", "host"), ("DEPLOY_PORT", "port"), ("DEPLOY_USER", "user"),
                     ("DEPLOY_PASSWORD", "password"), ("DEPLOY_REMOTE_PATH", "remote_path"),
                     ("DEPLOY_HEALTH_URL", "health_url")]:
        if os.getenv(k_env):
            cfg[k] = os.getenv(k_env)
    # 命令行覆盖
    for k in ["host", "port", "user", "password", "remote_path", "health_url"]:
        v = getattr(args, k, None)
        if v:
            cfg[k] = v
    cfg.setdefault("port", 22)
    return cfg


def _ensure_conn(cfg: dict):
    """校验连接信息，必要时提示输入密码。"""
    for k in ["host", "user", "remote_path"]:
        if not cfg.get(k):
            print(f"[错误] 缺少连接配置：{k}（用 deploy.config.json / 环境变量 / 命令行 提供）")
            sys.exit(1)
    if not cfg.get("password"):
        cfg["password"] = getpass.getpass(f"输入 {cfg['user']}@{cfg['host']} 的密码：")


def _connect(cfg: dict):
    try:
        import paramiko
    except ImportError:
        print("[错误] 需要 paramiko：请先运行  pip install paramiko")
        sys.exit(1)
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh.connect(cfg["host"], port=int(cfg["port"]), username=cfg["user"],
                    password=cfg["password"], timeout=30)
    except Exception as e:
        print(f"[错误] SSH 连接失败：{e}"); sys.exit(1)
    return ssh


def _run_remote(ssh, script: str):
    """执行远程脚本，实时回显（隐藏哨兵标记），返回 (rc, 是否健康)。"""
    stdin, stdout, stderr = ssh.exec_command(script, get_pty=True)
    health_ok = False
    for line in iter(stdout.readline, ""):
        if not line:
            continue
        if "__HEALTH_OK__" in line:
            health_ok = True
        if line.strip() in ("__HEALTH_OK__", "__DEPLOY_DONE__", "__ROLLBACK_DONE__"):
            continue
        sys.stdout.write("  " + line)
        sys.stdout.flush()
    rc = stdout.channel.recv_exit_status()
    err = stderr.read().decode("utf-8", "ignore")
    if rc != 0 and err:
        print(err)
    return rc, health_ok


def _health_url(cfg: dict) -> str:
    # 发布后健康检查地址；默认走服务器本机端口，公网反代部署请在配置里填公网地址
    return cfg.get("health_url") or "http://localhost:8000/api/health"


def _health_cmd(url: str) -> str:
    return (
        f"echo '等待服务就绪（{url}）...'; "
        "for i in $(seq 1 20); do "
        f"if curl -fsSk '{url}' >/dev/null 2>&1; then echo '__HEALTH_OK__'; break; fi; "
        "sleep 3; done; "
    )


def do_rollback(cfg: dict, yes: bool):
    """回滚到上一个镜像 datamerge:prev（部署前自动打的标签），不重建、秒级生效。"""
    _ensure_conn(cfg)
    remote_path = cfg["remote_path"].rstrip("/")
    print("=" * 60)
    print(f"回滚目标：{cfg['user']}@{cfg['host']}  →  {remote_path}")
    print("动作：把 datamerge:prev 重新标记为 latest 并重建容器（不 build）")
    if not yes and input("确认回滚到上一版本？输入 y 继续：").strip().lower() != "y":
        print("已取消。"); return
    ssh = _connect(cfg)
    try:
        script = (
            f"cd '{remote_path}' && "
            "if ! docker image inspect datamerge:prev >/dev/null 2>&1; then "
            "echo '没有可回滚的 datamerge:prev 镜像（至少成功发布过一次才会有）'; exit 1; fi && "
            "docker tag datamerge:prev datamerge:latest && "
            "if docker compose version >/dev/null 2>&1; then docker compose up -d --force-recreate; "
            "else docker-compose up -d --force-recreate; fi; "
            f"{_health_cmd(_health_url(cfg))}"
            "echo '__ROLLBACK_DONE__'"
        )
        print("回滚中...")
        rc, health_ok = _run_remote(ssh, script)
        print("=" * 60)
        if rc != 0:
            print(f"[错误] 回滚失败，退出码 {rc}"); sys.exit(rc)
        print("✅ 已回滚到上一版本" + ("，健康检查通过。" if health_ok else "（健康检查未通过，请看 docker compose logs -f app）"))
    finally:
        ssh.close()


def main():
    ap = argparse.ArgumentParser(description="发布 DataMerge 到 Docker 服务器")
    ap.add_argument("--host")
    ap.add_argument("--port", type=int)
    ap.add_argument("--user")
    ap.add_argument("--password")
    ap.add_argument("--remote-path", dest="remote_path")
    ap.add_argument("--health-url", dest="health_url", help="发布后健康检查地址（默认 http://localhost:8000/api/health）")
    ap.add_argument("--yes", action="store_true", help="跳过确认")
    ap.add_argument("--dry-run", action="store_true", help="只打包不上传")
    ap.add_argument("--no-build", action="store_true", help="只上传代码，不执行 docker compose build")
    ap.add_argument("--code-only", action="store_true",
                    help="只传代码，跳过 libs/fonts（需先全量发布过一次，服务器已有这些）")
    ap.add_argument("--rollback", action="store_true",
                    help="回滚到上一个镜像 datamerge:prev（不上传、不 build，秒级生效）")
    args = ap.parse_args()

    cfg = load_config(args)

    # 回滚：不打包不上传，直接远程切回 :prev 镜像
    if args.rollback:
        do_rollback(cfg, args.yes)
        return

    exclude_dirs = set(EXCLUDE_DIRS)
    if args.code_only:
        exclude_dirs |= CODE_ONLY_EXTRA_DIRS

    # dry-run 只需打包
    ts = time.strftime("%Y%m%d_%H%M%S")
    tar_name = f"datamerge_deploy_{ts}.tar.gz"
    local_tar = PROJECT_ROOT / "temp" / tar_name
    local_tar.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("打包本地代码..." + ("（code-only：跳过 libs/fonts）" if args.code_only else ""))
    n = build_tarball(local_tar, exclude_dirs)
    size_mb = local_tar.stat().st_size / 1e6
    print(f"  已打包 {n} 个文件，{size_mb:.1f} MB → {local_tar.name}")
    print(f"  排除目录：{', '.join(sorted(exclude_dirs))}")
    print(f"  排除文件：{', '.join(EXCLUDE_PATTERNS)}")

    if args.dry_run:
        print("[dry-run] 仅打包，未上传。包保留在", local_tar)
        return

    _ensure_conn(cfg)
    remote_path = cfg["remote_path"].rstrip("/")
    remote_tar = f"/tmp/{tar_name}"
    print("=" * 60)
    print(f"目标：{cfg['user']}@{cfg['host']}:{cfg['port']}  →  {remote_path}")
    print(f"动作：备份当前代码({tar_name.replace('deploy','backup')}) + 镜像打 :prev → 上传新代码 → " +
          ("仅上传" if args.no_build else "docker compose up -d --build"))
    if not args.yes:
        if input("确认发布？输入 y 继续：").strip().lower() != "y":
            print("已取消。"); return

    ssh = _connect(cfg)
    try:
        # 1) SFTP 上传压缩包
        print(f"上传 {tar_name} ...")
        sftp = ssh.open_sftp()
        sftp.put(str(local_tar), remote_tar)
        sftp.close()
        print("  上传完成。")

        # 2) 发布前备份：当前代码打 code_时间戳.tar.gz（保留最近10份）+ 当前镜像打 :prev
        backup_cmd = (
            f"BK='{remote_path}/_deploy_backups'; mkdir -p \"$BK\"; "
            f"if [ -e '{remote_path}/docker-compose.yml' ]; then "
            f"tar czf \"$BK/code_{ts}.tar.gz\" -C '{remote_path}' "
            "--exclude=_deploy_backups --exclude=tenants --exclude=data --exclude=logs "
            "--exclude=output --exclude=temp --exclude=global_assets . 2>/dev/null || true; "
            "fi; "
            "ls -1t \"$BK\"/code_*.tar.gz 2>/dev/null | tail -n +11 | xargs -r rm -f || true; "
            "docker image inspect datamerge:latest >/dev/null 2>&1 && docker tag datamerge:latest datamerge:prev || true; "
        )

        build_cmd = "" if args.no_build else (
            "export DOCKER_BUILDKIT=1 COMPOSE_DOCKER_CLI_BUILD=1 && "
            "if docker compose version >/dev/null 2>&1; then docker compose up -d --build; "
            "else docker-compose up -d --build; fi && "
        )
        health_cmd = "" if args.no_build else _health_cmd(_health_url(cfg))
        remote_script = (
            f"set -e && mkdir -p '{remote_path}' && "
            f"{backup_cmd}"
            f"tar xzf '{remote_tar}' -C '{remote_path}' && "
            f"cd '{remote_path}' && "
            f"{build_cmd}"
            f"rm -f '{remote_tar}'; "
            f"{health_cmd}"
            f"echo '__DEPLOY_DONE__'"
        )
        print("远程备份 + 解压 + 发布中（首次 build 可能数分钟）...")
        rc, health_ok = _run_remote(ssh, remote_script)
        if rc != 0:
            print(f"[错误] 远程命令退出码 {rc}（旧容器未受影响，可继续运行）")
            sys.exit(rc)
        print("=" * 60)
        if args.no_build:
            print("✅ 代码已上传（未重建容器）。")
        elif health_ok:
            print("✅ 发布完成，健康检查通过（/api/health 正常）。回滚：python deploy.py --rollback")
        else:
            print("⚠️ 发布完成，但健康检查未通过——容器可能还在启动或启动失败。"
                  "\n   查看日志： docker compose logs -f app ；回滚： python deploy.py --rollback")
    finally:
        ssh.close()
        try:
            local_tar.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    main()
