import requests
import json
import os
from pathlib import Path

BASE_URL = "http://127.0.0.1:8000"
INPUT_DIR = Path(__file__).parent / "test_input"
OUTPUT_DIR = Path(__file__).parent / "test_output"
OUTPUT_DIR.mkdir(exist_ok=True)

ENDPOINT_MAP = {
    ".pdf": "/convert/pdf",
    ".docx": "/convert/docx",
    ".doc": "/convert/docx",
    ".pptx": "/convert/pptx",
    ".xlsx": "/convert/excel",
    ".xls": "/convert/excel",
}

def test_file(file_path: Path):
    ext = file_path.suffix.lower()
    endpoint = ENDPOINT_MAP.get(ext)
    if endpoint is None:
        print(f"[SKIP] Unsupported extension: {file_path.name}")
        return

    print(f"[SEND] {file_path.name}  →  {endpoint}")
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(
                f"{BASE_URL}{endpoint}",
                files={"file": (file_path.name, f)},
                timeout=300,
            )
        out_name = file_path.stem + ".txt"
        out_path = OUTPUT_DIR / out_name
        with open(out_path, "w", encoding="utf-8") as out:
            out.write(f"File   : {file_path.name}\n")
            out.write(f"Endpoint: {BASE_URL}{endpoint}\n")
            out.write(f"Status : {resp.status_code}\n")
            out.write("=" * 80 + "\n\n")
            try:
                data = resp.json()
                out.write(json.dumps(data, ensure_ascii=False, indent=2))
            except Exception:
                out.write(resp.text)
        print(f"  [OK] Saved → {out_path}")
    except Exception as e:
        out_name = file_path.stem + ".txt"
        out_path = OUTPUT_DIR / out_name
        with open(out_path, "w", encoding="utf-8") as out:
            out.write(f"File   : {file_path.name}\n")
            out.write(f"ERROR  : {e}\n")
        print(f"  [ERR] {e}")

if __name__ == "__main__":
    files = sorted(INPUT_DIR.iterdir())
    for f in files:
        # skip temp Office lock files
        if f.name.startswith("~$"):
            print(f"[SKIP] Temp file: {f.name}")
            continue
        if f.is_file():
            test_file(f)
    print("\nDone. Results saved in:", OUTPUT_DIR)
