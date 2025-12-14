FROM python:3.11-slim-bookworm AS builder

WORKDIR /app
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONPATH=/app

# build 阶段只安装编译工具
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
    build-essential \
 && rm -rf /var/lib/apt/lists/*

# 编译 wheel
RUN pip install --no-cache-dir -U pip \
 && pip wheel --no-cache-dir -w /wheels \
    fastmcp mcp pandas numpy openpyxl xlsxwriter xlrd chardet \
    plotly matplotlib seaborn requests psutil formulas tabulate python-magic

# ---------------- runtime ----------------
FROM python:3.11-slim-bookworm AS runtime

WORKDIR /app
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONPATH=/app

# 安装运行时依赖 libmagic1
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
    libmagic1 \
 && rm -rf /var/lib/apt/lists/*

# 安装 Python wheel
COPY --from=builder /wheels /wheels
RUN pip install --no-cache-dir -U pip \
 && pip install --no-cache-dir /wheels/* \
 && rm -rf /wheels

# 复制程序文件
COPY server.py /app/server.py
COPY config.py config_manager.py error_codes.py error_handlers.py interface_standards.py /app/
COPY data_verification.py data_quality_tools.py excel_* formulas_tools.py formula_processor.py /app/
COPY df_processed_error_handler.py column_checker.py comprehensive_data_verification.py /app/
COPY core /app/core
COPY services /app/services
COPY service_management /app/service_management
COPY security /app/security
COPY monitoring /app/monitoring
COPY utils /app/utils
COPY templates /app/templates
COPY config /app/config

CMD ["python", "server.py"]
