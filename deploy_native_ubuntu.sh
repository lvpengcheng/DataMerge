#!/usr/bin/env bash
# DataMerge Ubuntu 原生部署（不使用 Docker）
# 用法：sudo ./deploy_native_ubuntu.sh install|update|restart|status|logs

set -Eeuo pipefail

SERVICE_NAME="${SERVICE_NAME:-datamerge}"
APP_DIR="${APP_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)}"
APP_USER="${APP_USER:-${SUDO_USER:-$(id -un)}}"
APP_GROUP="${APP_GROUP:-$(id -gn "${APP_USER}" 2>/dev/null || echo "${APP_USER}")}"
APP_PORT="${APP_PORT:-8000}"
PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="${VENV_DIR:-${APP_DIR}/venv}"
MEMORY_MAX="${MEMORY_MAX:-6G}"
MEMORY_HIGH="${MEMORY_HIGH:-5G}"
MEMORY_SWAP_MAX="${MEMORY_SWAP_MAX:-0}"
CPU_QUOTA="${CPU_QUOTA:-200%}"
TASKS_MAX="${TASKS_MAX:-256}"
PIP_INDEX_URL="${PIP_INDEX_URL:-}"
ACTION="${1:-install}"
UNIT_PATH="/etc/systemd/system/${SERVICE_NAME}.service"

log() { printf '[DataMerge] %s\n' "$*"; }
die() { printf '[DataMerge][错误] %s\n' "$*" >&2; exit 1; }

require_root() {
    [[ "${EUID}" -eq 0 ]] || die "请使用 sudo 运行：sudo $0 ${ACTION}"
}

validate_paths() {
    [[ -f "${APP_DIR}/requirements.txt" ]] || die "APP_DIR 不是 DataMerge 项目目录：${APP_DIR}"
    [[ -f "${APP_DIR}/backend/app/main.py" ]] || die "缺少 backend/app/main.py"
    [[ -f "${APP_DIR}/libs/Aspose.Cells.dll" ]] || die "缺少 libs/Aspose.Cells.dll"
    [[ -f "${APP_DIR}/libs/libSkiaSharp.so" ]] || die "缺少 libs/libSkiaSharp.so"
    id "${APP_USER}" >/dev/null 2>&1 || die "运行用户不存在：${APP_USER}"
}

install_system_packages() {
    log "安装 Ubuntu 系统依赖"
    export DEBIAN_FRONTEND=noninteractive
    apt-get update
    apt-get install -y --no-install-recommends \
        ca-certificates curl wget gnupg tzdata \
        "${PYTHON_BIN}" "${PYTHON_BIN}-venv" "${PYTHON_BIN}-dev" \
        build-essential gcc g++ pkg-config default-libmysqlclient-dev \
        libicu-dev libssl-dev libfontconfig1 fontconfig libfreetype6 libgdiplus \
        redis-server
}

install_dotnet_runtime() {
    if command -v dotnet >/dev/null 2>&1 && dotnet --list-runtimes | grep -q '^Microsoft.NETCore.App 9\.'; then
        log ".NET 9 Runtime 已安装"
        return
    fi
    log "安装 Microsoft .NET 9 Runtime"
    # shellcheck disable=SC1091
    source /etc/os-release
    [[ "${ID:-}" == "ubuntu" ]] || die "当前仅支持 Ubuntu，检测到：${ID:-unknown}"
    local repo_deb="/tmp/packages-microsoft-prod.deb"
    wget -q "https://packages.microsoft.com/config/ubuntu/${VERSION_ID}/packages-microsoft-prod.deb" -O "${repo_deb}"
    dpkg -i "${repo_deb}"
    rm -f "${repo_deb}"
    apt-get update
    apt-get install -y --no-install-recommends dotnet-runtime-9.0
}

prepare_runtime() {
    log "创建 Python 虚拟环境并安装依赖"
    if [[ ! -x "${VENV_DIR}/bin/python" ]]; then
        sudo -u "${APP_USER}" "${PYTHON_BIN}" -m venv "${VENV_DIR}"
    fi
    sudo -u "${APP_USER}" "${VENV_DIR}/bin/python" -m pip install --upgrade pip wheel setuptools
    local pip_args=()
    [[ -z "${PIP_INDEX_URL}" ]] || pip_args+=(--index-url "${PIP_INDEX_URL}")
    sudo -u "${APP_USER}" "${VENV_DIR}/bin/pip" install "${pip_args[@]}" -r "${APP_DIR}/requirements.txt"

    install -m 0755 "${APP_DIR}/libs/libSkiaSharp.so" /usr/local/lib/libSkiaSharp.so
    ldconfig

    local runtime_version
    runtime_version="$(dotnet --list-runtimes | awk '/^Microsoft.NETCore.App 9\./ {gsub(/[\[\]]/, "", $2); print $2}' | sort -V | tail -1)"
    [[ -n "${runtime_version}" ]] || die "未找到 Microsoft.NETCore.App 9.x Runtime"
    cat > "${APP_DIR}/libs/runtimeconfig.json" <<EOF
{
  "runtimeOptions": {
    "tfm": "net9.0",
    "framework": {"name": "Microsoft.NETCore.App", "version": "${runtime_version}"},
    "rollForward": "LatestPatch",
    "additionalProbingPaths": ["${APP_DIR}/libs"]
  }
}
EOF
    chown "${APP_USER}:${APP_GROUP}" "${APP_DIR}/libs/runtimeconfig.json"

    if [[ -d "${APP_DIR}/fonts" ]]; then
        mkdir -p /usr/local/share/fonts/datamerge
        find "${APP_DIR}/fonts" -maxdepth 1 -type f -exec install -m 0644 {} /usr/local/share/fonts/datamerge/ \;
        fc-cache -f >/dev/null
    fi

    mkdir -p "${APP_DIR}"/{tenants,data,logs,output,compare_results,temp,global_assets}
    chown -R "${APP_USER}:${APP_GROUP}" \
        "${APP_DIR}"/{tenants,data,logs,output,compare_results,temp,global_assets} "${VENV_DIR}"

    if [[ ! -f "${APP_DIR}/.env" ]]; then
        [[ -f "${APP_DIR}/.env.example" ]] || die "缺少 .env 和 .env.example"
        sudo -u "${APP_USER}" cp "${APP_DIR}/.env.example" "${APP_DIR}/.env"
        log "已生成 .env，请在启动前填写数据库和 AI 配置：${APP_DIR}/.env"
    fi
    chmod 0600 "${APP_DIR}/.env"
    chown "${APP_USER}:${APP_GROUP}" "${APP_DIR}/.env"

    if grep -Eq '^REDIS_URL=redis://redis([:/]|$)' "${APP_DIR}/.env"; then
        log "警告：.env 仍使用 Docker 主机名 redis；原生部署请改为 redis://127.0.0.1:6379/0"
    fi
}

verify_aspose() {
    log "验证 Aspose.Cells / .NET 运行环境"
    sudo -u "${APP_USER}" env \
        PYTHONPATH="${APP_DIR}" DOTNET_ROOT=/usr/share/dotnet DOTNET_CLI_HOME=/tmp \
        DOTNET_ROLL_FORWARD=LatestMajor DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false \
        LD_LIBRARY_PATH="${APP_DIR}/libs:/usr/local/lib:/usr/lib" \
        "${VENV_DIR}/bin/python" -c \
        "import aspose_init; assert aspose_init.is_initialized(); from Aspose.Cells import Workbook; w=Workbook(); w.Dispose(); print('Aspose.Cells OK')"
}

write_systemd_unit() {
    log "写入 systemd 服务：${UNIT_PATH}"
    cat > "${UNIT_PATH}" <<EOF
[Unit]
Description=DataMerge FastAPI Service
After=network-online.target redis-server.service
Wants=network-online.target redis-server.service

[Service]
Type=simple
User=${APP_USER}
Group=${APP_GROUP}
WorkingDirectory=${APP_DIR}
Environment=PYTHONPATH=${APP_DIR}
Environment=PYTHONUNBUFFERED=1
Environment=PYTHONUTF8=1
Environment=LANG=C.UTF-8
Environment=LC_ALL=C.UTF-8
Environment=TZ=Asia/Shanghai
Environment=DOTNET_ROOT=/usr/share/dotnet
Environment=DOTNET_CLI_HOME=/tmp
Environment=DOTNET_ROLL_FORWARD=LatestMajor
Environment=DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false
Environment=LD_LIBRARY_PATH=${APP_DIR}/libs:/usr/local/lib:/usr/lib
ExecStartPre=${VENV_DIR}/bin/python -m backend.database.init_db
ExecStart=${VENV_DIR}/bin/uvicorn backend.app.main:app --host 0.0.0.0 --port ${APP_PORT} --workers 1 --timeout-keep-alive 120
Restart=on-failure
RestartSec=5s
TimeoutStartSec=120s
TimeoutStopSec=30s
KillMode=control-group
SendSIGKILL=yes
OOMPolicy=stop
MemoryHigh=${MEMORY_HIGH}
MemoryMax=${MEMORY_MAX}
MemorySwapMax=${MEMORY_SWAP_MAX}
CPUQuota=${CPU_QUOTA}
TasksMax=${TASKS_MAX}
LimitNOFILE=65535

[Install]
WantedBy=multi-user.target
EOF
    systemctl daemon-reload
}

install_service() {
    require_root; validate_paths
    install_system_packages
    install_dotnet_runtime
    prepare_runtime
    verify_aspose
    write_systemd_unit
    systemctl enable --now redis-server
    systemctl enable --now "${SERVICE_NAME}"
    local healthy=0
    for _ in $(seq 1 20); do
        if curl -fsS --max-time 2 "http://127.0.0.1:${APP_PORT}/api/health" >/dev/null 2>&1; then
            healthy=1
            break
        fi
        sleep 1
    done
    if [[ "${healthy}" -ne 1 ]]; then
        systemctl --no-pager --full status "${SERVICE_NAME}" || true
        die "服务已安装但健康检查未通过；请运行 journalctl -u ${SERVICE_NAME} -n 200 查看原因"
    fi
    log "原生部署完成：http://$(hostname -I | awk '{print $1}'):${APP_PORT}"
    log "日志：sudo journalctl -u ${SERVICE_NAME} -f"
}

update_service() {
    require_root; validate_paths
    install_dotnet_runtime
    prepare_runtime
    verify_aspose
    write_systemd_unit
    systemctl restart "${SERVICE_NAME}"
    systemctl --no-pager --full status "${SERVICE_NAME}" || true
}

case "${ACTION}" in
    install) install_service ;;
    update) update_service ;;
    restart) require_root; systemctl restart "${SERVICE_NAME}"; systemctl --no-pager --full status "${SERVICE_NAME}" || true ;;
    status) systemctl --no-pager --full status "${SERVICE_NAME}" || true ;;
    logs) journalctl -u "${SERVICE_NAME}" -f ;;
    *) die "未知操作：${ACTION}；支持 install|update|restart|status|logs" ;;
esac
