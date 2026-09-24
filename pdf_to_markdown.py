"""
PDF to Markdown Converter with Image Extraction

Usage:
    python pdf_to_markdown.py <input.pdf> [output_dir]

Images are embedded as base64 data URIs directly in the markdown — no image
files are written to disk.
"""

import sys
import os
import re
import base64
import hashlib
import io
import pdfplumber
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    import io as _io
    sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = _io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def clean_cell(cell):
    if cell is None:
        return ""
    text = str(cell).strip()
    text = re.sub(r"\s*\n\s*", " ", text)
    text = text.replace("|", "\\|")
    return text


def table_to_markdown(table):
    if not table:
        return ""
    cleaned = [[clean_cell(c) for c in row] for row in table]
    max_cols = max(len(row) for row in cleaned)
    for row in cleaned:
        while len(row) < max_cols:
            row.append("")
    if all(all(c == "" for c in row) for row in cleaned):
        return ""
    lines = []
    header = cleaned[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for row in cleaned[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


# Chars drawn under a slightly scaled matrix are flagged non-upright; read them as normal LTR text.
TEXT_DIR_SETTINGS = {"line_dir_rotated": "ttb", "char_dir_rotated": "ltr"}


def _snapped_edges(page, tol=1.0, min_len=3):
    """Rebuild near-axis-aligned lines/thin curves as exact h/v edges.

    pdfplumber marks a line as vertical unless top == bottom exactly, so
    horizontal rules with float jitter (e.g. inside a scaled Form XObject)
    are misclassified and the table is never detected.
    """
    horizontal, vertical = [], []
    for obj in page.lines + page.curves:
        w = obj["x1"] - obj["x0"]
        h = obj["bottom"] - obj["top"]
        if h < tol and w >= min_len:
            y = (obj["top"] + obj["bottom"]) / 2
            horizontal.append({"object_type": "line", "x0": obj["x0"], "x1": obj["x1"],
                               "top": y, "bottom": y, "width": w, "height": 0})
        elif w < tol and h >= min_len:
            x = (obj["x0"] + obj["x1"]) / 2
            vertical.append({"object_type": "line", "x0": x, "x1": x,
                             "top": obj["top"], "bottom": obj["bottom"], "width": 0, "height": h})
    return horizontal, vertical


def extract_page(page, page_num, seen_images=None, duplicate_hashes=None):
    """Extract content from a single page. Images returned as base64 data URIs."""
    if seen_images is None:
        seen_images = {}
    if duplicate_hashes is None:
        duplicate_hashes = set()

    parts = []
    image_refs = []

    for i, img in enumerate(page.images or []):
        try:
            img_crop = page.crop((img["x0"], img["top"], img["x1"], img["bottom"]))
            img_data = img_crop.to_image(resolution=200)

            buf = io.BytesIO()
            img_data.save(buf, format='PNG')
            img_bytes = buf.getvalue()
            img_hash = hashlib.md5(img_bytes).hexdigest()

            if img_hash in duplicate_hashes:
                continue

            if img_hash in seen_images:
                # Reuse the already-encoded data URI
                image_refs.append(f"![Image {page_num}-{i+1}]({seen_images[img_hash]})")
            else:
                b64 = base64.b64encode(img_bytes).decode('ascii')
                data_uri = f"data:image/png;base64,{b64}"
                seen_images[img_hash] = data_uri
                image_refs.append(f"![Image {page_num}-{i+1}]({data_uri})")
        except Exception as e:
            print(f"Warning: Failed to extract image from page {page_num}: {e}")

    text = page.extract_text(**TEXT_DIR_SETTINGS) or ""
    h_edges, v_edges = _snapped_edges(page)
    table_settings = {
        "explicit_horizontal_lines": h_edges,
        "explicit_vertical_lines": v_edges,
        **{f"text_{k}": v for k, v in TEXT_DIR_SETTINGS.items()},
    }
    tables = page.extract_tables(table_settings) or []

    table_cell_texts = set()
    for table in tables:
        for row in table:
            for cell in row:
                if cell:
                    table_cell_texts.add(cell.strip())
                    table_cell_texts.add(re.sub(r"\s+", " ", cell.strip()))

    if text:
        text = re.sub(
            r"P\.\s*\d+\s*\n?"
            r"This document is the proprietary property of NTT DATA TAIWAN\..*?(?:NTT DATA TAIWAN\.|$)",
            "",
            text,
            flags=re.DOTALL,
        )
        text = text.strip()

        if tables and table_cell_texts:
            filtered_lines = []
            for line in text.split("\n"):
                stripped = line.strip()
                normalized = re.sub(r"\s+", " ", stripped)
                if stripped and normalized not in table_cell_texts:
                    filtered_lines.append(line)
                elif not stripped:
                    filtered_lines.append(line)
            text = "\n".join(filtered_lines).strip()

        if text:
            parts.append(text)

    if image_refs:
        parts.append("\n".join(image_refs))

    for table in tables:
        md_table = table_to_markdown(table)
        if md_table:
            parts.append(md_table)

    return "\n\n".join(parts)


def pdf_to_markdown(pdf_path, output_dir):
    """Convert PDF to Markdown. Images are embedded as base64 data URIs."""
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # First pass: find duplicate images
    print("  Scanning for duplicate images...", end="", flush=True)
    image_hash_count = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for img in (page.images or []):
                try:
                    img_crop = page.crop((img["x0"], img["top"], img["x1"], img["bottom"]))
                    buf = io.BytesIO()
                    img_crop.to_image(resolution=200).save(buf, format='PNG')
                    h = hashlib.md5(buf.getvalue()).hexdigest()
                    image_hash_count[h] = image_hash_count.get(h, 0) + 1
                except Exception:
                    pass

    duplicate_hashes = {h for h, count in image_hash_count.items() if count > 1}
    print(f"\n  Found {len(duplicate_hashes)} duplicate image(s) to remove")

    # Second pass: extract content
    pages_md = []
    seen_images = {}
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            page_md = extract_page(page, i + 1, seen_images, duplicate_hashes)
            if page_md:
                pages_md.append(page_md)
            print(f"\r  Processing page {i + 1}/{total}...", end="", flush=True)

    print()
    return "\n\n---\n\n".join(pages_md)


def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {os.path.basename(__file__)} <input.pdf> [output_dir]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.join(os.path.dirname(pdf_path), base_name)

    try:
        print(f"Converting: {pdf_path}")
        markdown = pdf_to_markdown(pdf_path, output_dir)
        md_path = os.path.join(output_dir, "index.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown + "\n")
        print(f"Output: {md_path}")
        sys.exit(0)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
