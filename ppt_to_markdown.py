"""
PPT/PPTX to Markdown Converter with Image Extraction

Usage:
    python ppt_to_markdown.py <input.pptx> [output_dir]

Images are embedded as base64 data URIs directly in the markdown — no image
files are written to disk.
"""

import sys
import os
import re
import base64
import hashlib
import mimetypes
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def clean_cell(cell_text):
    if cell_text is None:
        return ""
    text = str(cell_text).strip()
    text = re.sub(r"\s*\n\s*", " ", text)
    text = text.replace("|", "\\|")
    return text


def table_to_markdown(table):
    rows = []
    for row in table.rows:
        rows.append([clean_cell(cell.text) for cell in row.cells])
    if not rows:
        return ""
    max_cols = max(len(row) for row in rows)
    for row in rows:
        while len(row) < max_cols:
            row.append("")
    if all(all(c == "" for c in row) for row in rows):
        return ""
    lines = []
    lines.append("| " + " | ".join(rows[0]) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def extract_text_from_shape(shape):
    if not shape.has_text_frame:
        return ""
    paragraphs = []
    for para in shape.text_frame.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        level = para.level if para.level else 0
        if level > 0:
            paragraphs.append("  " * (level - 1) + f"- {text}")
        else:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def get_shape_sort_key(shape, slide_shapes=None):
    top = shape.top if shape.top is not None else 0
    left = shape.left if shape.left is not None else 0
    col_tolerance = 1228800
    return (left // col_tolerance, top)


def _image_to_data_uri(image_bytes, content_type):
    """Encode image bytes as a base64 data URI."""
    ext_to_mime = {
        "image/png": "image/png",
        "image/jpeg": "image/jpeg",
        "image/gif": "image/gif",
        "image/bmp": "image/bmp",
        "image/tiff": "image/tiff",
        "image/svg+xml": "image/svg+xml",
    }
    mime = ext_to_mime.get(content_type, content_type)
    b64 = base64.b64encode(image_bytes).decode('ascii')
    return f"data:{mime};base64,{b64}"


def extract_image_from_shape(shape, slide_num, img_counter, seen_hashes):
    """Extract image from a shape, returning a data URI markdown reference."""
    try:
        image = shape.image
        content_type = image.content_type
        image_bytes = image.blob

        # Skip EMF/WMF vector images
        if content_type in ("image/x-emf", "image/x-wmf"):
            return None, img_counter

        img_hash = hashlib.md5(image_bytes).hexdigest()

        if img_hash in seen_hashes:
            img_counter += 1
            return f"![]({seen_hashes[img_hash]})", img_counter

        data_uri = _image_to_data_uri(image_bytes, content_type)
        img_counter += 1
        seen_hashes[img_hash] = data_uri
        return f"![]({data_uri})", img_counter

    except Exception as e:
        print(f"  Warning: Failed to extract image from slide {slide_num}: {e}")
        return None, img_counter


def process_group_shape(group_shape, slide_num, img_counter, seen_hashes):
    items = []
    for shape in sorted(group_shape.shapes, key=get_shape_sort_key):
        sort_key = get_shape_sort_key(shape)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            group_items, img_counter = process_group_shape(shape, slide_num, img_counter, seen_hashes)
            items.extend(group_items)
        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            md_table = table_to_markdown(shape.table)
            if md_table:
                items.append((sort_key, "table", md_table))
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE or (
            hasattr(shape, "image") and shape.image is not None
        ):
            try:
                ref, img_counter = extract_image_from_shape(shape, slide_num, img_counter, seen_hashes)
                if ref:
                    items.append((sort_key, "image", ref))
            except Exception:
                pass
        if shape.has_text_frame:
            text = extract_text_from_shape(shape)
            if text:
                items.append((sort_key, "text", text))
    return items, img_counter


def process_slide(slide, slide_num, seen_hashes):
    img_counter = 0
    items = []

    for shape in sorted(slide.shapes, key=get_shape_sort_key):
        sort_key = get_shape_sort_key(shape)

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            group_items, img_counter = process_group_shape(shape, slide_num, img_counter, seen_hashes)
            items.extend(group_items)
            continue

        if shape.has_table:
            md_table = table_to_markdown(shape.table)
            if md_table:
                items.append((sort_key, "table", md_table))
            continue

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            ref, img_counter = extract_image_from_shape(shape, slide_num, img_counter, seen_hashes)
            if ref:
                items.append((sort_key, "image", ref))
            continue

        if hasattr(shape, "image"):
            try:
                ref, img_counter = extract_image_from_shape(shape, slide_num, img_counter, seen_hashes)
                if ref:
                    items.append((sort_key, "image", ref))
            except Exception:
                pass

        if shape.has_text_frame:
            text = extract_text_from_shape(shape)
            if text:
                items.append((sort_key, "text", text))

    return "\n\n".join(content for _, _, content in items)


def pptx_to_markdown(pptx_path, output_dir):
    """Convert PPTX to Markdown. Images embedded as base64 data URIs."""
    output_dir = str(Path(output_dir))
    os.makedirs(output_dir, exist_ok=True)

    prs = Presentation(pptx_path)
    total_slides = len(prs.slides)
    seen_hashes = {}
    slides_md = []

    for i, slide in enumerate(prs.slides):
        print(f"\r  Processing slide {i + 1}/{total_slides}...", end="", flush=True)
        slide_md = process_slide(slide, i + 1, seen_hashes)
        if slide_md:
            slides_md.append(slide_md)

    print()
    return "\n\n---\n\n".join(slides_md)


def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {os.path.basename(__file__)} <input.pptx> [output_dir]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    if not os.path.isfile(pptx_path):
        print(f"Error: File not found: {pptx_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        base_name = os.path.splitext(os.path.basename(pptx_path))[0]
        output_dir = os.path.join(os.path.dirname(pptx_path), base_name)

    print(f"Converting: {pptx_path}")
    markdown = pptx_to_markdown(pptx_path, output_dir)
    md_path = os.path.join(output_dir, "index.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(markdown + "\n")
    print(f"Output: {md_path}")


if __name__ == "__main__":
    main()
