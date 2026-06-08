"""
Document to Markdown — FastAPI server

Endpoints:
  POST /convert/docx   - Convert DOCX or DOC → Markdown
  POST /convert/pdf    - Convert PDF → Markdown
  POST /convert/pptx   - Convert PPTX → Markdown
  POST /convert/excel  - Convert XLS or XLSX → Markdown (auto-selects converter)
  POST /convert/auto   - Auto-detect file type and convert

All endpoints accept multipart/form-data:
  file  (required) — the document to convert

Response JSON:
  {
    "filename": "report.xlsx",
    "markdown": "<all pages/sheets joined by double newline, images removed>",
    "markdown_with_descriptions": "<all pages/sheets joined by double newline, images replaced by descriptions>"
  }

Configuration (environment variables):
  CONVERTERS_DIR — directory that holds the converter scripts
                   defaults to the directory of this file

Run:
    uvicorn main:app --reload
"""

import base64
import concurrent.futures
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import requests as _requests
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse

load_dotenv()

# ---------------------------------------------------------------------------
# Image-to-description via Dify vision workflow
# ---------------------------------------------------------------------------

_DIFY_API_KEY      = os.environ["DIFY_API_KEY"]
_DIFY_WORKFLOW_ID  = os.environ["DIFY_WORKFLOW_ID"]
_DIFY_USER_ID      = os.environ.get("DIFY_USER_ID", "converter_api")
_DIFY_BASE_URL     = os.environ.get("DIFY_BASE_URL", "https://ai-dev.adata.com")
_DIFY_UPLOAD_URL   = f"{_DIFY_BASE_URL}/v1/files/upload"
_DIFY_WORKFLOW_URL = f"{_DIFY_BASE_URL}/v1/workflows/{_DIFY_WORKFLOW_ID}/run"
_DIFY_AUTH_HEADER  = {"Authorization": f"Bearer {_DIFY_API_KEY}"}


def image_to_description(image_bytes: bytes, mime_type: str) -> str:
    """
    Upload image bytes to the Dify vision workflow and return its text description.
    Falls back to an error string if either API call fails.
    """
    ext = mime_type.split("/")[-1] if "/" in mime_type else "png"
    try:
        upload_resp = _requests.post(
            _DIFY_UPLOAD_URL,
            headers=_DIFY_AUTH_HEADER,
            files={
                "file": (f"image.{ext}", image_bytes, mime_type),
                "user": (None, _DIFY_USER_ID),
            },
            timeout=30,
        )
        upload_resp.raise_for_status()
        file_id = upload_resp.json()["id"]
    except Exception as exc:
        return f"[Image description unavailable: upload failed — {exc}]"

    payload = {
        "inputs": {
            "img": {
                "transfer_method": "local_file",
                "upload_file_id": file_id,
                "type": "image",
            }
        },
        "response_mode": "blocking",
        "user": _DIFY_USER_ID,
    }
    try:
        wf_resp = _requests.post(
            _DIFY_WORKFLOW_URL,
            headers={**_DIFY_AUTH_HEADER, "Content-Type": "application/json"},
            json=payload,
            timeout=120,
        )
        wf_resp.raise_for_status()
        return wf_resp.json()["data"]["outputs"]["result"]
    except Exception as exc:
        return f"[Image description unavailable: workflow failed — {exc}]"


# ---------------------------------------------------------------------------
# Markdown post-processing: split one data-URI markdown into two variants
# ---------------------------------------------------------------------------

_DATA_URI_RE = re.compile(
    r'!\[([^\]]*)\]\(data:(?P<mime>[^;]+);base64,(?P<b64>[A-Za-z0-9+/=\n]+)\)'
)


def _split_markdown(raw: str) -> tuple[str, str]:
    """
    Return (markdown_no_images, markdown_with_descriptions).
    - markdown_no_images        : data URI image tags removed entirely
    - markdown_with_descriptions: data URI image tags replaced by image_to_description() text

    All image_to_description() calls run in parallel (up to 3 threads) so that a
    document with many images doesn't pay the Dify round-trip latency serially.
    """
    matches = list(_DATA_URI_RE.finditer(raw))

    if not matches:
        # Fast path: no images at all
        return raw, raw

    # Decode every image's bytes up-front (cheap, CPU-bound)
    decoded: list[tuple[bytes, str]] = []
    for m in matches:
        try:
            img_bytes = base64.b64decode(m.group("b64"))
        except Exception:
            img_bytes = b""
        decoded.append((img_bytes, m.group("mime")))

    # Fan out all Dify calls in parallel, max 5 concurrent threads
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [
            executor.submit(image_to_description, img_bytes, mime)
            for img_bytes, mime in decoded
        ]
        descriptions = [f.result() for f in futures]  # preserve order

    # Reassemble both output strings
    no_img: list[str] = []
    with_desc: list[str] = []
    last = 0
    for m, desc in zip(matches, descriptions):
        segment = raw[last:m.start()]
        no_img.append(segment)
        with_desc.append(segment)
        with_desc.append(desc)
        last = m.end()
    no_img.append(raw[last:])
    with_desc.append(raw[last:])
    return "".join(no_img), "".join(with_desc)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

CONVERTERS_DIR = Path(os.environ.get("CONVERTERS_DIR", Path(__file__).parent))

app = FastAPI(title="Document to Markdown API", version="1.0.0")


# ---------------------------------------------------------------------------
# Converter selection helpers
# ---------------------------------------------------------------------------

def get_file_type(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    return {
        ".doc": "doc", ".docx": "docx",
        ".xls": "xls", ".xlsx": "xlsx",
        ".pdf": "pdf", ".pptx": "pptx",
    }.get(ext, "unknown")


def choose_excel_converter(filename: str) -> str:
    if Path(filename).name.startswith("TW-"):
        return "excel_ISO_to_markdown.py"
    return "excel_sop_kb_to_markdown.py"


CONVERTER_MAP = {
    "doc":  "markitdown_to_md.py",
    "docx": "markitdown_to_md.py",
    "pdf":  "pdf_to_markdown.py",
    "pptx": "ppt_to_markdown.py",
}


# ---------------------------------------------------------------------------
# Subprocess runner
# ---------------------------------------------------------------------------

def _run_script(script_name: str, input_file: Path, output_dir: Path) -> tuple[bool, str]:
    """Run a converter script as a subprocess. Returns (success, error_message)."""
    script_path = CONVERTERS_DIR / script_name
    if not script_path.exists():
        return False, f"Converter script not found: {script_path}"

    # excel_ISO_to_markdown.py derives its output dir from the input file path
    if script_name == "excel_ISO_to_markdown.py":
        cmd = [sys.executable, str(script_path), str(input_file)]
    else:
        cmd = [sys.executable, str(script_path), str(input_file), str(output_dir)]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=300,
        )
        if result.returncode == 0:
            return True, ""
        error = result.stderr.strip() or (result.stdout.strip().splitlines() or ["unknown error"])[-1]
        return False, error
    except subprocess.TimeoutExpired:
        return False, "Conversion timed out (300 s)"
    except Exception as exc:
        return False, str(exc)


# ---------------------------------------------------------------------------
# Output collection
# ---------------------------------------------------------------------------

def _collect_output(output_dir: Path) -> tuple[str, str]:
    """
    Read all converted markdown files and return (markdown, markdown_with_descriptions).
    Multiple pages/sheets are joined with a double newline separator.

    Supported layouts:
      Flat   : output_dir/index.md                   (docx/pdf/pptx)
      Sharded: output_dir/<SheetName>/index.md        (both excel converters)
    """
    # Case 1: per-sheet subdirectories (both Excel converters write this layout)
    sheet_files = sorted(
        output_dir.glob("*/index.md"),
        key=lambda p: p.parent.name,
    )
    if sheet_files:
        all_md, all_desc = [], []
        for f in sheet_files:
            raw = f.read_text(encoding="utf-8")
            md, md_desc = _split_markdown(raw)
            all_md.append(md)
            all_desc.append(md_desc)
        return "\n\n".join(all_md), "\n\n".join(all_desc)

    # Case 2: flat index.md (docx / pdf / pptx), with .md glob fallback
    md_file = output_dir / "index.md"
    if not md_file.exists():
        candidates = list(output_dir.glob("*.md"))
        if not candidates:
            raise HTTPException(status_code=500, detail="Converter produced no Markdown output")
        md_file = candidates[0]

    raw = md_file.read_text(encoding="utf-8")
    return _split_markdown(raw)


def _build_response(filename: str, markdown: str, markdown_with_descriptions: str,
                    extra: dict | None = None) -> JSONResponse:
    """Assemble the final API response dict."""
    body = {
        "filename": filename,
        "markdown": markdown,
        "markdown_with_descriptions": markdown_with_descriptions,
        **(extra or {}),
    }
    return JSONResponse(body)


# ---------------------------------------------------------------------------
# Shared conversion driver  (docx / pdf / pptx)
# ---------------------------------------------------------------------------

def _convert(script: str, upload: UploadFile) -> JSONResponse:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        input_file = tmp_path / upload.filename
        with input_file.open("wb") as f:
            shutil.copyfileobj(upload.file, f)

        out_dir = tmp_path / input_file.stem
        out_dir.mkdir(parents=True, exist_ok=True)

        # markitdown_to_md.py creates its own sub-folder named after the input file
        if script == "markitdown_to_md.py":
            run_output_dir  = out_dir.parent
            final_output_dir = out_dir.parent / input_file.stem
        else:
            run_output_dir  = out_dir
            final_output_dir = out_dir

        success, error = _run_script(script, input_file, run_output_dir)
        if not success:
            raise HTTPException(status_code=500, detail=f"Conversion failed: {error}")

        md, md_desc = _collect_output(final_output_dir)

    return _build_response(upload.filename, md, md_desc)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.post("/convert/docx", summary="Convert DOCX or DOC to Markdown")
async def convert_docx(file: UploadFile = File(...)):
    if Path(file.filename).suffix.lower() not in (".docx", ".doc"):
        raise HTTPException(status_code=400, detail="Only .docx and .doc files are accepted")
    return _convert("markitdown_to_md.py", file)


@app.post("/convert/pdf", summary="Convert PDF to Markdown")
async def convert_pdf(file: UploadFile = File(...)):
    if Path(file.filename).suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="Only .pdf files are accepted")
    return _convert("pdf_to_markdown.py", file)


@app.post("/convert/pptx", summary="Convert PPTX to Markdown")
async def convert_pptx(file: UploadFile = File(...)):
    if Path(file.filename).suffix.lower() != ".pptx":
        raise HTTPException(status_code=400, detail="Only .pptx files are accepted")
    return _convert("ppt_to_markdown.py", file)


@app.post("/convert/excel", summary="Convert XLS or XLSX to Markdown")
async def convert_excel(file: UploadFile = File(...)):
    if Path(file.filename).suffix.lower() not in (".xls", ".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xls and .xlsx files are accepted")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        input_file = tmp_path / file.filename
        with input_file.open("wb") as f:
            shutil.copyfileobj(file.file, f)

        script  = choose_excel_converter(file.filename)
        out_dir = tmp_path / input_file.stem
        out_dir.mkdir(parents=True, exist_ok=True)

        success, error = _run_script(script, input_file, out_dir)
        if not success:
            raise HTTPException(status_code=500, detail=f"Conversion failed: {error}")

        md, md_desc = _collect_output(out_dir)

    return _build_response(file.filename, md, md_desc)


@app.post("/convert/auto", summary="Auto-detect file type and convert to Markdown")
async def convert_auto(file: UploadFile = File(...)):
    file_type = get_file_type(file.filename)
    if file_type == "unknown":
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported extension '{Path(file.filename).suffix}'. "
                "Accepted: .doc .docx .xls .xlsx .pdf .pptx"
            ),
        )

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        input_file = tmp_path / file.filename
        with input_file.open("wb") as f:
            shutil.copyfileobj(file.file, f)

        if file_type in ("xls", "xlsx"):
            script = choose_excel_converter(file.filename)
        else:
            script = CONVERTER_MAP[file_type]

        if script == "markitdown_to_md.py":
            out_dir          = tmp_path / input_file.stem
            out_dir.mkdir(parents=True, exist_ok=True)
            run_output_dir   = out_dir.parent
            final_output_dir = out_dir.parent / input_file.stem
        else:
            out_dir          = tmp_path / input_file.stem
            out_dir.mkdir(parents=True, exist_ok=True)
            run_output_dir   = out_dir
            final_output_dir = out_dir

        success, error = _run_script(script, input_file, run_output_dir)
        if not success:
            raise HTTPException(status_code=500, detail=f"Conversion failed: {error}")

        md, md_desc = _collect_output(final_output_dir)

    return _build_response(file.filename, md, md_desc, extra={"detected_type": file_type})
