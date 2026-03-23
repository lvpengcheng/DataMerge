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
    libicu-dev \
    libssl-dev \
    libfontconfig1 \
    fontconfig \
    libfreetype6 \
    libgdiplus \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

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
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

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

COPY excel_parser.py aspose_init.py run.py ./
COPY backend/ ./backend/
COPY frontend/ ./frontend/
COPY global_assets/ ./global_assets/

# ========== 7. 验证 ==========
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

# ========== 8. 运行时目录 ==========
RUN mkdir -p tenants data logs output compare_results temp

# ========== 9. 环境变量 ==========
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1
ENV DATABASE_URL=sqlite:///./data/data.db
ENV LD_LIBRARY_PATH=/app/libs:/usr/lib:${LD_LIBRARY_PATH}
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false
ENV DOTNET_ROLL_FORWARD=LatestMajor

# ========== 10. 暴露端口 ==========
EXPOSE 8000

# ========== 11. 健康检查 ==========
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:8000/api/health || exit 1

# ========== 12. 启动 ==========
CMD ["sh", "-c", "python -m backend.database.init_db && uvicorn backend.app.main:app --host 0.0.0.0 --port 8000 --workers 2"]
