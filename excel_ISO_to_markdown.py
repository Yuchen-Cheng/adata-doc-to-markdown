"""
Excel to Markdown Converter (ISO / TW- files)
===============================================
Converts each worksheet in an Excel file into Markdown with Mermaid flowcharts.
Images are embedded as base64 data URIs — no image files are written to disk.

Supported Formats:
  - .xlsx (Office Open XML) - Recommended
  - .xls (Office 97-2003) - Also supported via COM automation

Requirements:
  - Windows (COM automation)
  - Python 3.8+
  - pywin32  (pip install pywin32)

Usage:
  python excel_ISO_to_markdown.py <input.xlsx|input.xls> [--sheet <name>]
"""

import argparse
import base64
import math
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

DECISION_TYPES = {4, 63, 110}
ROUNDED_TYPES = {5, 62, 116}

# MIME types for common image extensions
_EXT_MIME = {
    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
    '.gif': 'image/gif', '.bmp': 'image/bmp', '.tiff': 'image/tiff',
    '.svg': 'image/svg+xml', '.webp': 'image/webp',
}


def safe_get_text(shape):
    try:
        txt = shape.DrawingObject.Text
        if txt:
            return txt
    except Exception:
        pass
    try:
        txt = shape.TextFrame.Characters().Text
        if txt:
            return txt
    except Exception:
        pass
    try:
        tf2 = shape.TextFrame2
        if tf2.HasText:
            return tf2.TextRange.Text
    except Exception:
        pass
    return ""


def make_node_id(index):
    return f"n{index}"


def sanitize_text(text):
    text = text.strip()
    text = text.replace('"', "'")
    text = text.replace('\r\n', '<br>').replace('\r', '<br>').replace('\n', '<br>')
    text = re.sub(r'(<br>){3,}', '<br><br>', text)
    return text


def sanitize_filename(name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    safe = safe.strip('. ')
    return safe or "sheet"


def mermaid_node_def(node_id, text, auto_type):
    safe = sanitize_text(text)
    q = f'"{safe}"'
    if auto_type in DECISION_TYPES:
        return f'{node_id}{{{{{q}}}}}'
    elif auto_type in ROUNDED_TYPES:
        return f'{node_id}({q})'
    else:
        return f'{node_id}[{q}]'


def extract_images_from_sheet(ws, wb_path):
    """
    Extract images for a sheet from the xlsx zip and return list of data URI markdown strings.
    Only works for .xlsx files.
    """
    image_markdowns = []
    if not wb_path or not wb_path.lower().endswith('.xlsx'):
        return image_markdowns

    try:
        sheet_index = None
        wb_com = ws.Parent
        for i in range(1, wb_com.Sheets.Count + 1):
            if wb_com.Sheets(i).Name == ws.Name:
                sheet_index = i
                break
        if sheet_index is None:
            return image_markdowns

        with zipfile.ZipFile(wb_path, 'r') as zf:
            rels_path = f'xl/worksheets/_rels/sheet{sheet_index}.xml.rels'
            try:
                rels_data = zf.read(rels_path)
            except KeyError:
                return image_markdowns

            root = ET.fromstring(rels_data)
            ns_uri = 'http://schemas.openxmlformats.org/package/2006/relationships'
            images_to_extract = set()

            for rel in root.findall(f'{{{ns_uri}}}Relationship'):
                if 'drawing' not in rel.get('Type', ''):
                    continue
                drawing_target = rel.get('Target', '')
                if not drawing_target:
                    continue
                drawing_file = drawing_target.replace('../', '')
                drawing_abs = 'xl/' + drawing_file
                drawing_rels_path = drawing_abs.replace('drawings/', 'drawings/_rels/').replace('.xml', '.xml.rels')
                try:
                    d_rels_data = zf.read(drawing_rels_path)
                    d_root = ET.fromstring(d_rels_data)
                    for d_rel in d_root.findall(f'{{{ns_uri}}}Relationship'):
                        if 'image' in d_rel.get('Type', ''):
                            target = d_rel.get('Target', '')
                            if target:
                                images_to_extract.add('xl/' + target.replace('../', ''))
                except Exception:
                    pass

            for media_path in sorted(images_to_extract):
                try:
                    file_data = zf.read(media_path)
                    ext = os.path.splitext(media_path)[1].lower()
                    mime = _EXT_MIME.get(ext, 'image/png')
                    b64 = base64.b64encode(file_data).decode('ascii')
                    data_uri = f"data:{mime};base64,{b64}"
                    image_markdowns.append(f"![]({data_uri})")
                    print(f"  Embedded image: {os.path.basename(media_path)}")
                except Exception:
                    pass

    except Exception:
        pass

    return image_markdowns


def convert_sheet(ws, wb_path=None):
    """Convert one worksheet to Markdown. Images embedded as base64 data URIs."""
    sheet_name = ws.Name
    shape_count = ws.Shapes.Count

    # Phase 1: Classify connectors
    connector_indices = set()
    for i in range(1, shape_count + 1):
        shape = ws.Shapes.Item(i)
        try:
            cf = shape.ConnectorFormat
            _ = cf.BeginConnected
            connector_indices.add(i)
        except Exception:
            pass

    # Phase 2: Collect text nodes
    nodes = {}
    idx_by_name = {}
    for i in range(1, shape_count + 1):
        if i in connector_indices:
            continue
        shape = ws.Shapes.Item(i)
        text = safe_get_text(shape)
        if not text.strip():
            continue
        auto_type = -1
        try:
            auto_type = shape.AutoShapeType
        except Exception:
            pass
        nodes[i] = {
            "id": make_node_id(i), "text": text, "auto_type": auto_type,
            "top": shape.Top, "left": shape.Left,
            "width": shape.Width, "height": shape.Height, "name": shape.Name,
        }
        idx_by_name.setdefault(shape.Name, []).append(i)

    # Phase 3: Extract edges
    edges = []
    for i in range(1, shape_count + 1):
        if i not in connector_indices:
            continue
        shape = ws.Shapes.Item(i)
        try:
            cf = shape.ConnectorFormat
        except Exception:
            continue

        begin_idx = end_idx = None

        try:
            if cf.BeginConnected:
                bs = cf.BeginConnectedShape
                candidates = idx_by_name.get(bs.Name, [])
                if len(candidates) == 1:
                    begin_idx = candidates[0]
                elif candidates:
                    begin_idx = min(candidates,
                                    key=lambda ci: abs(nodes[ci]["top"] - bs.Top) + abs(nodes[ci]["left"] - bs.Left))
        except Exception:
            pass

        try:
            if cf.EndConnected:
                es = cf.EndConnectedShape
                candidates = idx_by_name.get(es.Name, [])
                if len(candidates) == 1:
                    end_idx = candidates[0]
                elif candidates:
                    end_idx = min(candidates,
                                  key=lambda ci: abs(nodes[ci]["top"] - es.Top) + abs(nodes[ci]["left"] - es.Left))
        except Exception:
            pass

        if begin_idx is None or end_idx is None:
            s_top, s_left = shape.Top, shape.Left
            s_w, s_h = shape.Width, shape.Height
            corner_a = (s_left, s_top)
            corner_b = (s_left + s_w, s_top + s_h)
            _label_words = {"yes", "no", "y", "n", "是", "否"}
            _label_indices = {nidx for nidx, ninfo in nodes.items()
                              if ninfo["text"].strip().lower() in _label_words}

            def _nearest_node(px, py, exclude_idx=None, threshold=60):
                best, best_dist = None, threshold
                for nidx, ninfo in nodes.items():
                    if nidx == exclude_idx or nidx in _label_indices:
                        continue
                    n_cx = ninfo["left"] + ninfo.get("width", 50) / 2
                    n_cy = ninfo["top"] + ninfo.get("height", 30) / 2
                    dx = max(0, abs(px - n_cx) - ninfo.get("width", 50) / 2)
                    dy = max(0, abs(py - n_cy) - ninfo.get("height", 30) / 2)
                    dist = (dx**2 + dy**2) ** 0.5
                    if dist < best_dist:
                        best_dist, best = dist, nidx
                return best

            def _node_center(nidx):
                n = nodes[nidx]
                return (n["left"] + n.get("width", 50) / 2, n["top"] + n.get("height", 30) / 2)

            if begin_idx is not None and end_idx is None:
                bcx, bcy = _node_center(begin_idx)
                da = abs(corner_a[0] - bcx) + abs(corner_a[1] - bcy)
                db = abs(corner_b[0] - bcx) + abs(corner_b[1] - bcy)
                ep = corner_b if db > da else corner_a
                end_idx = _nearest_node(ep[0], ep[1], exclude_idx=begin_idx)
            elif end_idx is not None and begin_idx is None:
                ecx, ecy = _node_center(end_idx)
                da = abs(corner_a[0] - ecx) + abs(corner_a[1] - ecy)
                db = abs(corner_b[0] - ecx) + abs(corner_b[1] - ecy)
                bp = corner_b if db > da else corner_a
                begin_idx = _nearest_node(bp[0], bp[1], exclude_idx=end_idx)
            else:
                na = _nearest_node(corner_a[0], corner_a[1])
                nb = _nearest_node(corner_b[0], corner_b[1], exclude_idx=na)
                if na and nb and na != nb:
                    begin_idx, end_idx = na, nb

        if begin_idx and end_idx and begin_idx in nodes and end_idx in nodes:
            label = safe_get_text(shape).strip()
            edges.append((begin_idx, end_idx, sanitize_text(label) if label else ""))

    has_flowchart = len(edges) > 0

    # Phase 4: Yes/No label matching
    label_nodes = set()
    for idx, info in list(nodes.items()):
        txt = info["text"].strip().lower()
        if txt in ("yes", "no", "y", "n", "是", "否"):
            if not any(idx == e[0] or idx == e[1] for e in edges):
                label_nodes.add(idx)

    decision_labels = defaultdict(list)
    for label_idx in label_nodes:
        li = nodes[label_idx]
        l_cx = li["left"] + li.get("width", 50) / 2
        l_cy = li["top"] + li.get("height", 20) / 2
        decision_sources = {bi for bi, _, _ in edges if nodes[bi]["auto_type"] in DECISION_TYPES}
        best_dec, best_dist = None, float("inf")
        for dec_idx in decision_sources:
            d = nodes[dec_idx]
            dist = math.hypot(l_cx - (d["left"] + d.get("width", 50) / 2),
                              l_cy - (d["top"] + d.get("height", 50) / 2))
            if dist < best_dist:
                best_dist, best_dec = dist, dec_idx
        if best_dec is not None and best_dist < 200:
            decision_labels[best_dec].append((label_idx, li["text"].strip(), l_cx, l_cy))

    for dec_idx, labels in decision_labels.items():
        d = nodes[dec_idx]
        d_cx = d["left"] + d.get("width", 50) / 2
        d_cy = d["top"] + d.get("height", 50) / 2
        dec_edges = [(ei, math.atan2(
            nodes[end_i]["top"] + nodes[end_i].get("height", 30) / 2 - d_cy,
            nodes[end_i]["left"] + nodes[end_i].get("width", 50) / 2 - d_cx))
            for ei, (bi, end_i, elabel) in enumerate(edges) if bi == dec_idx and not elabel]
        label_angles = [(li_idx, ltxt, math.atan2(l_cy - d_cy, l_cx - d_cx))
                        for li_idx, ltxt, l_cx, l_cy in labels]
        pairs = sorted(
            (abs(la - ea) if abs(la - ea) <= math.pi else 2 * math.pi - abs(la - ea),
             li_idx, ltxt, ei)
            for li_idx, ltxt, la in label_angles
            for ei, ea in dec_edges
        )
        used_e, used_l = set(), set()
        for diff, li_idx, ltxt, ei in pairs:
            if li_idx in used_l or ei in used_e:
                continue
            bi, end_i, _ = edges[ei]
            edges[ei] = (bi, end_i, ltxt)
            used_e.add(ei)
            used_l.add(li_idx)

    for idx in label_nodes:
        del nodes[idx]

    # Phase 5: Swim-lane detection
    connected_nodes = {bi for bi, ei, _ in edges} | {ei for bi, ei, _ in edges}
    candidate_headers = {idx: info for idx, info in nodes.items() if idx not in connected_nodes}

    header_nodes = {}
    if candidate_headers:
        min_top = min(info["top"] for info in candidate_headers.values())
        top_row = {idx: info for idx, info in candidate_headers.items()
                   if info["top"] <= min_top + 30}
        if len(top_row) >= 2:
            header_nodes = top_row

    flow_nodes = {idx: info for idx, info in nodes.items() if idx not in header_nodes}
    sorted_headers = sorted(header_nodes.items(), key=lambda x: x[1]["left"])
    lane_assignments = {}

    if sorted_headers:
        header_centers = [(idx, info["left"] + info.get("width", 50) / 2) for idx, info in sorted_headers]
        lane_ranges = []
        for i, (hidx, hcenter) in enumerate(header_centers):
            x_min = -float("inf") if i == 0 else (header_centers[i - 1][1] + hcenter) / 2
            x_max = float("inf") if i == len(header_centers) - 1 else (hcenter + header_centers[i + 1][1]) / 2
            lane_ranges.append((hidx, x_min, x_max))
        for fidx, finfo in flow_nodes.items():
            f_center = finfo["left"] + finfo.get("width", 50) / 2
            assigned = next((hidx for hidx, x_min, x_max in lane_ranges if x_min <= f_center <= x_max), None)
            if assigned is None:
                assigned = min(header_centers, key=lambda hc: abs(hc[1] - f_center))[0]
            lane_assignments[fidx] = assigned

    decorative = {fidx for fidx in list(flow_nodes) if fidx not in connected_nodes and fidx not in header_nodes}
    for fidx in decorative:
        del flow_nodes[fidx]
        lane_assignments.pop(fidx, None)

    # Phase 6: Read table data
    table_data = []
    try:
        used = ws.UsedRange
        if used:
            nrows = min(used.Rows.Count, 200)
            max_col_to_check = 50
            try:
                for test_col in range(1, min(100, used.Columns.Count + 1)):
                    if ws.Cells(1, test_col).Value:
                        max_col_to_check = max(max_col_to_check, test_col)
            except Exception:
                pass
            max_col_to_check = min(max_col_to_check + 5, 100)
            for r in range(1, nrows + 1):
                try:
                    row = []
                    last_col = 0
                    for c in range(1, max_col_to_check + 1):
                        try:
                            v = ws.Cells(r, c).Value
                            cell_str = str(v).strip() if v else ""
                            row.append(cell_str)
                            if cell_str:
                                last_col = c
                        except Exception:
                            row.append("")
                    if last_col > 0 and any(row[:last_col]):
                        table_data.append(row[:last_col])
                except Exception:
                    continue
    except Exception:
        pass

    # Phase 7: Build Markdown
    md = [f"# {sheet_name}", "", "---", ""]

    if has_flowchart and flow_nodes:
        md += ["# 第一部分：流程圖與圖形資訊", ""]
        md += ["## 1.1 流程圖", ""]
        if sorted_headers:
            md.append(f"> {len(sorted_headers)} 個部門以 subgraph 表示歸屬。")
        md += ["", "```mermaid", "flowchart TD"]

        sorted_flow = sorted(flow_nodes.items(), key=lambda x: (x[1]["top"], x[1]["left"]))
        if sorted_headers and lane_assignments:
            for hidx, hinfo in sorted_headers:
                lane_name = sanitize_text(hinfo["text"]).replace("<br>", " ")
                lane_id = hinfo["id"].replace("-", "_")
                md.append(f'    subgraph {lane_id}["{lane_name}"]')
                for fidx, finfo in sorted_flow:
                    if lane_assignments.get(fidx) == hidx:
                        md.append(f"        {mermaid_node_def(finfo['id'], finfo['text'], finfo['auto_type'])}")
                md.append("    end")
            md.append("")
        else:
            for idx, info in sorted_flow:
                md.append(f"    {mermaid_node_def(info['id'], info['text'], info['auto_type'])}")
            md.append("")

        for bi, ei, label in edges:
            if bi not in flow_nodes or ei not in flow_nodes:
                continue
            fid, tid = flow_nodes[bi]["id"], flow_nodes[ei]["id"]
            md.append(f'    {fid} -->|"{label}"| {tid}' if label else f"    {fid} --> {tid}")
        md += ["```", ""]

        md += ["## 1.2 節點清單", "", "| ID | 文字 | Mermaid 符號 | 所屬部門 |", "| --- | --- | --- | --- |"]
        for idx, info in sorted_flow:
            txt = info["text"].replace("\n", " ").replace("\r", " ").replace("|", "\\|").strip()
            at = info["auto_type"]
            symbol = "菱形 (決策)" if at in DECISION_TYPES else "圓角矩形" if at in ROUNDED_TYPES else "矩形"
            dept = header_nodes.get(lane_assignments.get(idx), {}).get("text", "").replace("\n", " ").strip()
            md.append(f"| {info['id']} | {txt} | {symbol} | {dept} |")
        md.append("")

        md += ["## 1.3 連線清單", "", "| # | 起點 | 終點 | 標籤 |", "| --- | --- | --- | --- |"]
        edge_num = 0
        for bi, ei, label in edges:
            if bi not in flow_nodes or ei not in flow_nodes:
                continue
            edge_num += 1
            md.append(f"| {edge_num} | {flow_nodes[bi]['id']} ({flow_nodes[bi]['text'][:20]}) "
                       f"| {flow_nodes[ei]['id']} ({flow_nodes[ei]['text'][:20]}) | {label} |")
        md += ["", "---", ""]

    if table_data:
        max_cols = max(len(r) for r in table_data)
        col_has_data = [any(r[ci].strip() for r in table_data if ci < len(r)) for ci in range(max_cols)]
        keep_cols = [i for i, has in enumerate(col_has_data) if has]
        if keep_cols:
            filtered = [[row[ci] if ci < len(row) else "" for ci in keep_cols] for row in table_data]
            hdr = [h.replace("\n", " ").replace("\r", " ").strip() for h in filtered[0]]
            hdr = [h if h else f"欄{i+1}" for i, h in enumerate(hdr)]
            md.append("| " + " | ".join(hdr) + " |")
            md.append("| " + " | ".join(["---"] * len(keep_cols)) + " |")
            for row in filtered[1:]:
                cells = [c.replace("\n", " ").replace("\r", " ").replace("|", "\\|").strip() for c in row]
                if any(cells):
                    md.append("| " + " | ".join(cells) + " |")
        md.append("")
    else:
        md += ["*此工作表無表格資料*", ""]

    # Phase 8: Embed images
    if wb_path is None:
        try:
            wb_path = ws.Parent.FullName
        except Exception:
            wb_path = None

    image_markdowns = extract_images_from_sheet(ws, wb_path)
    if image_markdowns:
        md += ["", "---", "", "# 圖片附件", ""]
        md.extend(image_markdowns)
        md.append("")

    return "\n".join(md), {
        "sheet": sheet_name, "nodes": len(flow_nodes), "edges": len(edges),
        "lanes": len(header_nodes), "has_flowchart": has_flowchart,
        "images": len(image_markdowns),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert Excel worksheets to Markdown (with Mermaid flowcharts)."
    )
    parser.add_argument("filepath", help="Path to the Excel file (.xlsx/.xls)")
    parser.add_argument("--sheet", default=None, help="Sheet name or substring. Omit to convert ALL sheets.")
    args = parser.parse_args()

    excel_path = os.path.abspath(args.filepath)
    if not os.path.exists(excel_path):
        print(f"Error: file not found: {excel_path}")
        sys.exit(1)

    file_dir = os.path.dirname(excel_path)
    file_name_no_ext = os.path.splitext(os.path.basename(excel_path))[0]
    output_dir = str(Path(os.path.join(file_dir, file_name_no_ext)))
    os.makedirs(output_dir, exist_ok=True)

    print(f"Opening Excel: {excel_path}")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(excel_path, ReadOnly=True)
        if wb is None:
            raise RuntimeError("Workbooks.Open returned None")

        target_sheets = [
            wb.Sheets(i) for i in range(1, wb.Sheets.Count + 1)
            if not args.sheet or args.sheet in wb.Sheets(i).Name
        ]

        if not target_sheets:
            print(f"No matching sheets found")
            sys.exit(1)

        print(f"Sheets to convert: {len(target_sheets)}")
        results = []

        for ws in target_sheets:
            sheet_dir = os.path.join(output_dir, sanitize_filename(ws.Name))
            os.makedirs(sheet_dir, exist_ok=True)

            print(f"\n--- Converting: {ws.Name} (shapes: {ws.Shapes.Count}) ---")
            try:
                md_content, stats = convert_sheet(ws, excel_path)
                out_path = os.path.join(sheet_dir, "index.md")
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(md_content)
                print(f"  -> {out_path}")
                print(f"     Lanes: {stats['lanes']}  Nodes: {stats['nodes']}  "
                      f"Edges: {stats['edges']}  Images: {stats['images']}")
                results.append((ws.Name, out_path, stats))
            except Exception as e:
                print(f"  Error converting {ws.Name}: {e}")

        print(f"\n{'='*50}")
        print(f"Converted {len(results)}/{len(target_sheets)} sheets:")
        for name, path, stats in results:
            fc = " (with flowchart)" if stats["has_flowchart"] else ""
            print(f"  {name} -> {os.path.basename(path)}{fc}")
        print(f"Output directory: {output_dir}")
        sys.exit(0 if len(results) == len(target_sheets) else 1)

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        print("Excel application closed.")


if __name__ == "__main__":
    main()
