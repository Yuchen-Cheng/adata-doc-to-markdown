"""
Document to Markdown — FastAPI server (local development / Docker)

For Azure Functions deployment, use function_app.py instead.

Endpoints:
  POST /convert/docx   - DOCX / DOC → Markdown
  POST /convert/pdf    - PDF → Markdown
  POST /convert/pptx   - PPTX → Markdown
  POST /convert/excel  - XLS / XLSX → Markdown
  POST /convert/auto   - Auto-detect and convert

All endpoints accept multipart/form-data with a single 'file' field.

Run locally:
    uvicorn main:app --reload
"""

from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse

load_dotenv()  # reads .env for local dev; Azure Functions uses App Settings

from converter_utils import convert_file, get_file_type  # noqa: E402

app = FastAPI(title="Document to Markdown API", version="1.0.0")


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _run(filename: str, file_bytes: bytes, allowed_exts: tuple[str, ...] | None = None,
         extra: dict | None = None) -> JSONResponse:
    if allowed_exts and Path(filename).suffix.lower() not in allowed_exts:
        raise HTTPException(
            status_code=400,
            detail=f"Only {', '.join(allowed_exts)} files are accepted.",
        )
    try:
        result = convert_file(filename, file_bytes)
    except ValueError as e:
        raise HTTPException(status_code=500, detail=str(e))
    if extra:
        result.update(extra)
    return JSONResponse(result)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.post("/convert/docx", summary="Convert DOCX or DOC to Markdown")
async def convert_docx(file: UploadFile = File(...)):
    return _run(file.filename, await file.read(), (".docx", ".doc"))


@app.post("/convert/pdf", summary="Convert PDF to Markdown")
async def convert_pdf(file: UploadFile = File(...)):
    return _run(file.filename, await file.read(), (".pdf",))


@app.post("/convert/pptx", summary="Convert PPTX to Markdown")
async def convert_pptx(file: UploadFile = File(...)):
    return _run(file.filename, await file.read(), (".pptx",))


@app.post("/convert/excel", summary="Convert XLS or XLSX to Markdown")
async def convert_excel(file: UploadFile = File(...)):
    return _run(file.filename, await file.read(), (".xls", ".xlsx"))


@app.post("/convert/auto", summary="Auto-detect file type and convert to Markdown")
async def convert_auto(file: UploadFile = File(...)):
    filename = file.filename
    file_type = get_file_type(filename)
    if file_type == "unknown":
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported extension '{Path(filename).suffix}'. "
                "Accepted: .doc .docx .xls .xlsx .pdf .pptx"
            ),
        )
    return _run(filename, await file.read(), extra={"detected_type": file_type})
