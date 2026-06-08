# Document to Markdown API

A self-contained FastAPI service that converts documents to Markdown by invoking the converter scripts as subprocesses.

## Setup

```bash
cd converter_api
pip install -r requirements.txt
```

## Run

```bash
uvicorn main:app --reload
```

Interactive docs: http://127.0.0.1:8000/docs

## Configuration

| Environment variable | Default | Purpose |
|---|---|---|
| `CONVERTERS_DIR` | parent directory of `converter_api/` | Path to the folder containing the converter scripts |

Example — if the API lives somewhere other than next to the scripts:

```bash
CONVERTERS_DIR=/path/to/scripts uvicorn main:app --reload
```

## Endpoints

| Method | Path | Accepts |
|---|---|---|
| POST | `/convert/docx` | `.docx`, `.doc` |
| POST | `/convert/pdf` | `.pdf` |
| POST | `/convert/pptx` | `.pptx` |
| POST | `/convert/excel` | `.xls`, `.xlsx` (auto-selects converter) |
| POST | `/convert/auto` | any supported format |

### Request

`multipart/form-data`:

| Field | Type | Required |
|---|---|---|
| `file` | file | yes |
| `output_dir` | string (path) | no — uses a temp dir if omitted |

### Response

```json
{
  "markdown":   "# Document title\n...",
  "images":     ["images/fig1.png"],
  "output_dir": "/absolute/path/to/output"
}
```

`/convert/auto` also includes `"detected_type": "pdf"` (or whichever type was detected).
