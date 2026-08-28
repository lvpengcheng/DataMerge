#!/usr/bin/env bash
set -euo pipefail

if [[ ${EUID} -ne 0 ]]; then
  echo "请使用 sudo bash tools/configure_ubuntu_docker_guard.sh"
  exit 1
fi

echo "== 当前资源 =="
free -h
swapon --show || true
df -hT / /var/lib/docker 2>/dev/null || true
docker info --format 'DockerRoot={{.DockerRootDir}} Cgroup={{.CgroupVersion}} Logging={{.LoggingDriver}}'

# 限制内核过度回收和脏页集中回写，降低 xlsx 大文件保存时的长时间 IO 冻结。
install -d -m 0755 /etc/sysctl.d
cat >/etc/sysctl.d/90-datamerge-guard.conf <<'EOF'
vm.swappiness=10
vm.dirty_background_ratio=3
vm.dirty_ratio=10
vm.vfs_cache_pressure=100
EOF
sysctl --system >/dev/null

# Docker 日志默认上限；compose 中已有显式轮转设置，这里保护其他容器。
install -d -m 0755 /etc/docker
if [[ ! -e /etc/docker/daemon.json ]]; then
  cat >/etc/docker/daemon.json <<'EOF'
{
  "log-driver": "json-file",
  "log-opts": {"max-size": "50m", "max-file": "3"},
  "live-restore": true
}
EOF
  systemctl restart docker
else
  echo "注意：/etc/docker/daemon.json 已存在，为避免覆盖，请手工合并 log-opts 和 live-restore。"
fi

echo "== 建议写入项目 .env 的参数 =="
cat <<'EOF'
APP_MEM_LIMIT=10g
APP_MEM_RESERVATION=2g
APP_CPUS=2
SUBPROCESS_MAX_MEMORY_MB=3072
SUBPROCESS_CONCURRENCY=1
SUBPROCESS_QUEUE_TIMEOUT=600
COMPUTE_PROC_TIMEOUT=900
COMPUTE_OUTPUT_VALUES_COPY=false
EOF

echo "完成。上述 .env 参数需要在项目目录执行 docker compose up -d --force-recreate 后生效。"

