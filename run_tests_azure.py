import requests
import json
import os
import time
from pathlib import Path

BASE_URL = "http://localhost:7071"
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

def test_file(file_path: Path) -> float | None:
    """Send one file to the Azure Functions local endpoint. Returns elapsed seconds, or None if skipped."""
    ext = file_path.suffix.lower()
    endpoint = ENDPOINT_MAP.get(ext)
    if endpoint is None:
        print(f"[SKIP] Unsupported extension: {file_path.name}")
        return None

    print(f"[SEND] {file_path.name}  →  {endpoint}")
    out_name = file_path.stem + ".txt"
    out_path = OUTPUT_DIR / out_name
    t0 = time.perf_counter()
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(
                f"{BASE_URL}{endpoint}",
                files={"file": (file_path.name, f)},
                timeout=300,
            )
        elapsed = time.perf_counter() - t0
        with open(out_path, "w", encoding="utf-8") as out:
            out.write(f"File    : {file_path.name}\n")
            out.write(f"Endpoint: {BASE_URL}{endpoint}\n")
            out.write(f"Status  : {resp.status_code}\n")
            out.write(f"Elapsed : {elapsed:.2f}s\n")
            out.write("=" * 80 + "\n\n")
            try:
                data = resp.json()
                out.write(json.dumps(data, ensure_ascii=False, indent=2))
            except Exception:
                out.write(resp.text)
        status_tag = "[OK] " if resp.status_code == 200 else f"[{resp.status_code}]"
        print(f"  {status_tag} {elapsed:.2f}s  →  {out_path}")
        return elapsed
    except Exception as e:
        elapsed = time.perf_counter() - t0
        with open(out_path, "w", encoding="utf-8") as out:
            out.write(f"File    : {file_path.name}\n")
            out.write(f"Elapsed : {elapsed:.2f}s\n")
            out.write(f"ERROR   : {e}\n")
        print(f"  [ERR] {elapsed:.2f}s  {e}")
        return elapsed

if __name__ == "__main__":
    files = sorted(INPUT_DIR.iterdir())
    timings: list[tuple[str, float]] = []
    total_start = time.perf_counter()

    for f in files:
        if f.name.startswith("~$"):
            print(f"[SKIP] Temp file: {f.name}")
            continue
        if f.is_file():
            elapsed = test_file(f)
            if elapsed is not None:
                timings.append((f.name, elapsed))

    total_elapsed = time.perf_counter() - total_start

    print("\n" + "=" * 60)
    print(f"{'FILE':<40} {'TIME':>8}")
    print("-" * 60)
    for name, secs in timings:
        print(f"{name:<40} {secs:>7.2f}s")
    if timings:
        print("-" * 60)
        avg = sum(s for _, s in timings) / len(timings)
        slowest = max(timings, key=lambda x: x[1])
        print(f"{'Total wall-clock':<40} {total_elapsed:>7.2f}s")
        print(f"{'Average per file':<40} {avg:>7.2f}s")
        print(f"{'Slowest':<40} {slowest[0]} ({slowest[1]:.2f}s)")
    print("=" * 60)
    print("Results saved in:", OUTPUT_DIR)
