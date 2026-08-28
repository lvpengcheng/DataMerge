# syntax=docker/dockerfile:1
# ========== Stage 1: 从微软官方镜像获取 .NET 9 运行时 ==========
FROM mcr.microsoft.com/dotnet/runtime:9.0 AS dotnet-runtime

# ========== Stage 2: Python 主镜像 ==========
FROM python:3.11-slim

WORKDIR /app

# ========== 1. 系统依赖 ==========
# libgdiplus: Aspose.Cells 渲染图片/PDF 的核心依赖
# libfontconfig1: 字体配置
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    g++ \
    default-libmysqlclient-dev \
    pkg-config \
    curl \
    tzdata \
    libicu-dev \
    libssl-dev \
    libfontconfig1 \
    fontconfig \
    libfreetype6 \
    libgdiplus \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# ========== 时区：上海（Asia/Shanghai） ==========
# 让容器内 datetime.now()/日志时间戳/报表生成时间均为北京时间
ENV TZ=Asia/Shanghai
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime \
    && echo $TZ > /etc/timezone

# ========== 2. 从 Stage 1 复制 .NET 9 运行时（免去手动安装） ==========
COPY --from=dotnet-runtime /usr/share/dotnet /usr/share/dotnet
RUN ln -s /usr/share/dotnet/dotnet /usr/local/bin/dotnet \
    && dotnet --list-runtimes

ENV DOTNET_ROOT=/usr/share/dotnet
ENV PATH="${DOTNET_ROOT}:${PATH}"
ENV DOTNET_CLI_HOME=/tmp

# ========== 3. 安装中文字体（解决 Excel 中文乱码） ==========
COPY ./fonts /usr/share/fonts/win-fonts
RUN fc-cache -fv

# ========== 4. Python 依赖 ==========
# 用 BuildKit 缓存挂载持久化 pip 下载缓存：跨构建复用已下载的 wheel，
# requirements 有改动时只下增量；缓存挂载不写进镜像层，镜像体积与 --no-cache-dir 时一致。
# 可选：构建时传国内镜像源加速下载（清华源示例）：
#   DOCKER_BUILDKIT=1 docker build --build-arg PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple .
ARG PIP_INDEX_URL=""
COPY requirements.txt .
RUN --mount=type=cache,target=/root/.cache/pip \
    pip install ${PIP_INDEX_URL:+-i $PIP_INDEX_URL} -r requirements.txt

# ========== 5. 复制应用代码 ==========
COPY libs/ ./libs/

# ========== 6. Linux 专用处理 ==========
# 删除 Windows 专用文件（Linux 不需要）
RUN rm -f /app/libs/libSkiaSharp.dll \
    && rm -f /app/libs/System.Text.Encoding.CodePages.dll

# libSkiaSharp.so 同时放到系统库目录，确保 .NET P/Invoke 能找到
RUN cp /app/libs/libSkiaSharp.so /usr/lib/libSkiaSharp.so \
    && ldconfig \
    && echo "OK: libSkiaSharp.so installed"

# 生成 Linux 专用 runtimeconfig.json（含 additionalProbingPaths）
RUN printf '{\n  "runtimeOptions": {\n    "tfm": "net9.0",\n    "framework": {\n      "name": "Microsoft.NETCore.App",\n      "version": "9.0.11"\n    },\n    "additionalProbingPaths": ["/app/libs"]\n  }\n}\n' > /app/libs/runtimeconfig.json

COPY excel_parser.py aspose_init.py run.py split_by_banner.py ./

# ========== 7. 验证（放在 COPY backend/frontend 之前：只依赖 libs + aspose_init，
#            这样改业务代码时该层仍命中缓存，不再每次重跑 .NET/Aspose 初始化测试） ==========
# 文件检查 + ldd（仅显示信息）
RUN echo "=== File check ===" \
    && test -f /app/libs/SkiaSharp.dll && echo "OK: SkiaSharp.dll" \
    && test -f /app/libs/Aspose.Cells.dll && echo "OK: Aspose.Cells.dll" \
    && test -f /app/libs/libSkiaSharp.so && echo "OK: libSkiaSharp.so" \
    && echo "=== ldd libSkiaSharp.so ===" \
    && (ldd /app/libs/libSkiaSharp.so || echo "WARNING: ldd had issues") \
    && echo "=== libs/ ===" \
    && ls -la /app/libs/

# Python 实际初始化 Aspose.Cells 测试
RUN python -c "\
import sys, os; \
sys.path.insert(0, '/app'); \
os.environ['LD_LIBRARY_PATH'] = '/app/libs:/usr/lib'; \
print('>>> Testing Aspose.Cells initialization...'); \
import aspose_init; \
ok = aspose_init.is_initialized(); \
print('>>> Init result:', ok); \
assert ok, 'Aspose.Cells initialization FAILED'; \
from Aspose.Cells import Workbook; \
wb = Workbook(); \
wb.Worksheets[0].Cells[0,0].PutValue('Docker build test'); \
print('>>> SUCCESS: Aspose.Cells + Workbook OK on Linux'); \
"

# ========== 业务代码（放在最后：改代码只重跑这几步轻量 COPY，
#            前面的 apt/dotnet/字体/pip/libs/Aspose 验证全部命中缓存） ==========
COPY backend/ ./backend/
COPY frontend/ ./frontend/

# ========== 8. 运行时目录 ==========
# global_assets 与 tenants/data/... 同为 docker-compose 卷挂载的运行时目录（发布时不上传），
# 运行时由宿主机目录遮盖镜像内容，故只需 mkdir 建目录，不 COPY（避免全新环境构建时源目录不存在而失败）。
RUN mkdir -p tenants data logs output compare_results temp global_assets

# ========== 9. 环境变量 ==========
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1
# 进程级强制 UTF-8：让所有 open() 文本读写默认 utf-8（不随系统 locale），
# 避免读脚本/CSV/文本时中文乱码。覆盖沙箱与 importlib 两条执行路径。
ENV PYTHONUTF8=1
ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8
ENV DATABASE_URL=sqlite:///./data/data.db
ENV LD_LIBRARY_PATH=/app/libs:/usr/lib:${LD_LIBRARY_PATH}
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false
ENV DOTNET_ROLL_FORWARD=LatestMajor

# ========== 10. 暴露端口 ==========
EXPOSE 8000

# ========== 11. 健康检查 ==========
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# ========== 12. 启动 ==========
CMD ["sh", "-c", "python -m backend.database.init_db && uvicorn backend.app.main:app --host 0.0.0.0 --port 8000 --workers 1 --timeout-keep-alive 600"]
