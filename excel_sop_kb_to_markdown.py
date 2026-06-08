"""
Excel/XLS/XLSX to Markdown Converter with Image Extraction

Usage:
    python excel_sop_kb_to_markdown.py <input.xls|input.xlsx> [output_dir]

Images are embedded as base64 data URIs directly in the markdown — no image
files are written to disk.
"""

import sys
import os
import re
import base64
import hashlib
import tempfile
import time
from openpyxl import load_workbook
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def convert_xls_to_xlsx(xls_path):
    """Convert .xls to a temp .xlsx via COM. Returns temp path (caller must delete)."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Cannot convert .xls files without pywin32. Install: pip install pywin32")

    print("  Converting .xls to .xlsx (COM automation)...")
    temp_xlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_xlsx_path = temp_xlsx.name
    temp_xlsx.close()

    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(xls_path), ReadOnly=True)
        wb.SaveAs(temp_xlsx_path, FileFormat=51)
        print("  Conversion successful")
        return temp_xlsx_path
    except Exception as e:
        try:
            os.remove(temp_xlsx_path)
        except Exception:
            pass
        raise RuntimeError(f"Failed to convert .xls: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(0.5)


def clean_cell_text(text):
    if text is None:
        return ""
    text = str(text).strip()
    return re.sub(r"\s+", " ", text)


def extract_text_from_sheet(ws):
    items = []
    merged_ranges = {cell.coord for cell in ws.merged_cells.ranges}
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
        for col_idx, cell in enumerate(row, 1):
            if cell.coordinate in str(merged_ranges):
                continue
            if cell.value is None:
                continue
            text = clean_cell_text(cell.value)
            if text:
                items.append((row_idx, col_idx, "text", text))
    return items


def extract_images_from_sheet(ws, seen_hashes=None):
    """Extract images from worksheet. Returns list of (row, col, data_uri_markdown)."""
    if seen_hashes is None:
        seen_hashes = {}

    image_refs = []
    if not hasattr(ws, '_images') or not ws._images:
        return image_refs, seen_hashes

    image_count = 0
    for image in ws._images:
        try:
            if not hasattr(image, 'anchor') or not hasattr(image.anchor, '_from'):
                continue
            col_idx = image.anchor._from.col + 1
            row_idx = image.anchor._from.row + 1

            if not hasattr(image, 'ref'):
                continue
            image.ref.seek(0)
            image_bytes = image.ref.read()

            fmt = getattr(image, 'format', 'png').lower()
            mime_map = {'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg',
                        'gif': 'image/gif', 'bmp': 'image/bmp'}
            mime = mime_map.get(fmt, f'image/{fmt}')

            img_hash = hashlib.md5(image_bytes).hexdigest()
            if img_hash in seen_hashes:
                image_refs.append((row_idx, col_idx, f"![]({seen_hashes[img_hash]})"))
                continue

            b64 = base64.b64encode(image_bytes).decode('ascii')
            data_uri = f"data:{mime};base64,{b64}"
            seen_hashes[img_hash] = data_uri
            image_count += 1
            image_refs.append((row_idx, col_idx, f"![]({data_uri})"))

        except Exception as e:
            print(f"  Warning: Failed to extract image: {e}")

    return image_refs, seen_hashes


def worksheet_to_markdown(ws, ws_name, seen_hashes=None):
    if seen_hashes is None:
        seen_hashes = {}

    print(f"  Processing worksheet: {ws_name}...", end="", flush=True)

    text_items = extract_text_from_sheet(ws)
    image_refs, seen_hashes = extract_images_from_sheet(ws, seen_hashes)

    all_items = [(r, c, "text", v) for r, c, _, v in text_items]
    all_items += [(r, c, "image", ref) for r, c, ref in image_refs]
    all_items.sort(key=lambda x: (x[0], x[1]))

    lines = []
    current_row = -1
    row_buffer = []

    for row_idx, col_idx, item_type, content in all_items:
        if row_idx != current_row:
            if row_buffer:
                lines.append("\n\n".join(row_buffer))
            current_row = row_idx
            row_buffer = []
        row_buffer.append(content)

    if row_buffer:
        lines.append("\n\n".join(row_buffer))

    markdown = "\n\n".join(lines)
    print(" done")
    return markdown, seen_hashes


def _safe_dirname(name):
    """Convert a sheet name to a safe directory name."""
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    return safe.strip('. ') or "sheet"


def excel_to_markdown(xlsx_path, output_dir):
    """
    Convert XLSX to Markdown. Each worksheet is written to its own subdirectory:
        output_dir/<SheetName>/index.md
    Images are embedded as base64 data URIs.
    """
    output_dir = str(Path(output_dir))
    os.makedirs(output_dir, exist_ok=True)

    wb = load_workbook(xlsx_path)
    try:
        seen_hashes = {}
        print(f"Converting: {xlsx_path}")
        for ws_name in wb.sheetnames:
            ws = wb[ws_name]
            ws_markdown, seen_hashes = worksheet_to_markdown(ws, ws_name, seen_hashes)
            if not ws_markdown:
                continue
            ws_markdown = f"# {ws_name}\n\n{ws_markdown}"
            sheet_dir = os.path.join(output_dir, _safe_dirname(ws_name))
            os.makedirs(sheet_dir, exist_ok=True)
            md_path = os.path.join(sheet_dir, "index.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(ws_markdown + "\n")
            print(f"  -> {md_path}")
    finally:
        try:
            wb.close()
        except Exception:
            pass


def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {os.path.basename(__file__)} <input.xls|input.xlsx> [output_dir]")
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_dir = os.path.join(os.path.dirname(input_path), base_name)

    xlsx_path = input_path
    temp_xlsx_path = None

    if input_path.lower().endswith('.xls'):
        try:
            temp_xlsx_path = convert_xls_to_xlsx(input_path)
            xlsx_path = temp_xlsx_path
        except RuntimeError as e:
            print(f"Error: {e}")
            sys.exit(1)

    try:
        excel_to_markdown(xlsx_path, output_dir)
        print(f"Output dir: {output_dir}")
    finally:
        if temp_xlsx_path and os.path.isfile(temp_xlsx_path):
            try:
                os.remove(temp_xlsx_path)
            except Exception:
                pass


if __name__ == "__main__":
    main()
