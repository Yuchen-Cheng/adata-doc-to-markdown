# ── Stage 1: build deps in a clean layer ────────────────────────────────────
FROM python:3.11-slim AS builder

WORKDIR /install

# System build tools needed by some Python packages (e.g. Wand C headers)
RUN apt-get update && apt-get install -y --no-install-recommends \
        gcc \
        libmagickwand-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install/pkg -r requirements.txt


# ── Stage 2: runtime image ───────────────────────────────────────────────────
FROM python:3.11-slim

LABEL maintainer="yuchen_cheng@omniamr.ai"
LABEL description="Document-to-Markdown conversion API"

# Runtime system libraries
# - imagemagick   : required by Wand (EMF/WMF → PNG conversion in markitdown_to_md.py)
# - libmagickwand : shared library for Wand Python binding
# NOTE: win32com (COM automation for .doc / .xls) is Windows-only and is NOT
#       available in this Linux image. Those formats will return an error at
#       runtime. All other formats (docx, xlsx, pdf, pptx) work normally.
RUN apt-get update && apt-get install -y --no-install-recommends \
        imagemagick \
        libmagickwand-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy installed Python packages from builder
COPY --from=builder /install/pkg /usr/local

WORKDIR /app

# Copy application source
COPY main.py                        ./
COPY markitdown_to_md.py            ./
COPY pdf_to_markdown.py             ./
COPY ppt_to_markdown.py             ./
COPY excel_sop_kb_to_markdown.py    ./
COPY excel_ISO_to_markdown.py       ./

# .env is NOT baked into the image — mount it at runtime or pass env vars.
# See the "Running" section in API_REPORT.md.

EXPOSE 8000

# Uvicorn: 4 workers, bind to all interfaces
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]
