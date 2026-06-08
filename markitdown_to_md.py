"""
MarkItDown DOCX/DOC to Markdown Converter

Usage:
    python markitdown_to_md.py <input.docx|input.doc> [output_dir]

Images are embedded as base64 data URIs directly in the markdown — no image
files are written to disk.

Supported formats:
  - .docx (Office Open XML)
  - .doc  (Office 97-2003, requires Windows + pywin32)
"""

import sys
import os
import re
import base64
import tempfile
from pathlib import Path
from markitdown import MarkItDown

CONVERT_FORMATS = {'x-emf', 'emf', 'x-wmf', 'wmf'}

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def convert_doc_to_docx(doc_path):
    """Convert a .doc file to a temp .docx using Word COM. Returns temp path."""
    import time
    try:
        import win32com.client as win32
    except ImportError:
        print("Error: pywin32 is required for .doc file support")
        sys.exit(1)

    temp_fd, temp_docx_path = tempfile.mkstemp(suffix='.docx')
    os.close(temp_fd)

    word = None
    doc = None
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs(temp_docx_path, FileFormat=16)
        return temp_docx_path
    except Exception as e:
        if os.path.exists(temp_docx_path):
            try:
                os.remove(temp_docx_path)
            except Exception:
                pass
        raise Exception(f"Failed to convert .doc to .docx: {e}")
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        time.sleep(0.5)


def convert_emf_to_png_data_uri(raw_bytes):
    """Convert EMF/WMF bytes to a PNG base64 data URI in memory."""
    # Try wand (ImageMagick)
    try:
        import io as _io
        from wand.image import Image as WandImage
        with WandImage(blob=raw_bytes, resolution=600) as img:
            img.format = 'png'
            buf = _io.BytesIO()
            img.save(file=buf)
            png_bytes = buf.getvalue()
        b64 = base64.b64encode(png_bytes).decode('ascii')
        return f"data:image/png;base64,{b64}"
    except Exception:
        pass

    # Fallback: Pillow
    try:
        import io as _io
        from PIL import Image
        img = Image.open(_io.BytesIO(raw_bytes))
        buf = _io.BytesIO()
        img.save(buf, 'PNG', compress_level=0)
        b64 = base64.b64encode(buf.getvalue()).decode('ascii')
        return f"data:image/png;base64,{b64}"
    except Exception:
        pass

    return None  # conversion failed, will keep original data URI


def process_data_uri(alt_text, ext, base64_data):
    """
    Keep images as data URIs. For EMF/WMF, attempt in-memory conversion to PNG.
    Returns a markdown image tag with a data URI.
    """
    if ext.lower() in CONVERT_FORMATS:
        try:
            raw_bytes = base64.b64decode(base64_data)
            converted = convert_emf_to_png_data_uri(raw_bytes)
            if converted:
                return f"![{alt_text}]({converted})"
        except Exception:
            pass
        # Fall through: keep original data URI (even if it's EMF/WMF)

    return f"![{alt_text}](data:image/{ext};base64,{base64_data})"


def markitdown_to_md(input_file, output_dir):
    """
    Convert DOCX/DOC to Markdown. Images are kept as base64 data URIs inline.
    """
    output_dir = str(Path(output_dir))
    os.makedirs(output_dir, exist_ok=True)

    print(f"Converting: {os.path.basename(input_file)}")

    md = MarkItDown()
    result = md.convert(input_file, keep_data_uris=True)
    content = result.text_content

    # Replace data URI image tags, optionally converting EMF/WMF to PNG in memory
    pattern = r'!\[([^\]]*)\]\(data:image\/(?P<ext>[^;]+);base64,(?P<data>[^)]+)\)'

    def handle_image(match):
        alt_text = match.group(1)
        ext = match.group('ext')
        base64_data = match.group('data')
        return process_data_uri(alt_text, ext, base64_data)

    final_md = re.sub(pattern, handle_image, content)
    return final_md


def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {os.path.basename(__file__)} <input.docx|input.doc> [output_dir]")
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    if len(sys.argv) >= 3:
        output_dir = os.path.join(sys.argv[2], base_name)
    else:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(input_path)), base_name)

    # .doc → temp .docx
    temp_docx = None
    actual_input = input_path
    if input_path.lower().endswith('.doc') and not input_path.lower().endswith('.docx'):
        print(f"Converting .doc to .docx first...")
        temp_docx = convert_doc_to_docx(input_path)
        actual_input = temp_docx

    try:
        final_md = markitdown_to_md(actual_input, output_dir)
        md_path = os.path.join(output_dir, "index.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(final_md + "\n")
        print(f"Output: {md_path}")
    finally:
        if temp_docx and os.path.exists(temp_docx):
            os.remove(temp_docx)


if __name__ == "__main__":
    main()
