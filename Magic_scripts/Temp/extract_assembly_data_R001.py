"""
Extract Assembly Data from Excel Work Instruction Files
========================================================
Usage:
    pip install openpyxl
    python extract_assembly_data.py

Extracts from each Excel file:
  1. Assembly Steps  — step-by-step instructions with cautions/notes
  2. Parts Referenced — part names with part numbers
  3. Comments — cell comments, threaded comments, and notes

Output folder: _extracted/
  - Individual JSON per file  → _extracted/<filename>.json
  - Combined JSON for all     → _extracted/_all_assembly_data.json
  - Summary CSV               → _extracted/_parts_summary.csv
  - Comments CSV              → _extracted/_comments_summary.csv
"""

import zipfile
import xml.etree.ElementTree as ET
import json
import csv
import re
import os
import glob


# ─── Configuration ───────────────────────────────────────
OUTPUT_FOLDER_NAME = "_extracted"

VIEW_LABELS = {"top view", "bottom view", "front view", "back view",
               "side view", "left view", "right view", "isometric view",
               "top", "bottom", "front", "back", "left", "right"}

PART_NUMBER_PATTERN = re.compile(
    r'[A-Z]{2,5}-\d{5,10}-\d+-\d+'   # e.g. GPMP-0300002-0-0
    r'|[A-Z]{2,5}-\d{3,}'             # e.g. ABC-12345
    r'|\d{3,}-\d{3,}'                 # e.g. 12345-67890
)

STEP_PATTERN = re.compile(r'^(\d+)\)\s*')  # e.g. "1) instruction..."
# ─────────────────────────────────────────────────────────

NS_XDR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def get_base_dir():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    work_dir = os.getcwd()
    if work_dir != script_dir:
        base_dir = work_dir
    else:
        base_dir = script_dir
    return base_dir


# ─── Comment Extraction ─────────────────────────────────

def extract_cell_comments(zip_file):
    """
    Extract classic cell comments (yellow sticky notes).
    Stored in xl/comments*.xml files.
    """
    comments = []
    all_files = zip_file.namelist()

    comment_files = [f for f in all_files
                     if f.startswith('xl/comments') and f.endswith('.xml')]

    for cf in comment_files:
        xml_bytes = zip_file.read(cf)
        root = ET.fromstring(xml_bytes)

        ns = ''
        if root.tag.startswith('{'):
            ns = root.tag.split('}')[0] + '}'

        authors_elem = root.find(f'{ns}authors')
        authors = []
        if authors_elem is not None:
            for a in authors_elem.findall(f'{ns}author'):
                authors.append(a.text if a.text else "")

        comment_list = root.find(f'{ns}commentList')
        if comment_list is not None:
            for c in comment_list.findall(f'{ns}comment'):
                cell_ref = c.get('ref', '?')
                author_id = int(c.get('authorId', 0))
                author = authors[author_id] if author_id < len(authors) else ""

                text_parts = []
                for t in c.findall(f'.//{ns}t'):
                    if t.text:
                        text_parts.append(t.text)
                text = "".join(text_parts).strip()

                if text:
                    comments.append({
                        "type": "cell_comment",
                        "cell": cell_ref,
                        "author": author,
                        "text": text
                    })

    return comments


def extract_threaded_comments(zip_file):
    """
    Extract modern threaded comments (reply chains).
    Stored in xl/threadedComments/*.xml
    """
    comments = []
    all_files = zip_file.namelist()

    threaded_files = [f for f in all_files
                      if 'threadedComments' in f and f.endswith('.xml')]

    for tf in threaded_files:
        xml_bytes = zip_file.read(tf)
        root = ET.fromstring(xml_bytes)

        ns = ''
        if root.tag.startswith('{'):
            ns = root.tag.split('}')[0] + '}'

        for tc in root.findall(f'{ns}threadedComment'):
            cell_ref = tc.get('ref', '?')
            text_elem = tc.find(f'{ns}text')
            text = text_elem.text.strip() if text_elem is not None and text_elem.text else ""
            parent_id = tc.get('parentId', '')

            if text:
                comments.append({
                    "type": "threaded_comment",
                    "cell": cell_ref,
                    "text": text,
                    "is_reply": bool(parent_id)
                })

    return comments


def extract_note_shapes(zip_file):
    """
    Extract notes embedded as VML shapes (legacy notes).
    Stored in xl/drawings/vmlDrawing*.vml
    """
    comments = []
    all_files = zip_file.namelist()

    vml_files = [f for f in all_files
                 if 'vmlDrawing' in f and f.endswith('.vml')]

    for vf in vml_files:
        xml_bytes = zip_file.read(vf)
        try:
            text_content = xml_bytes.decode('utf-8', errors='replace')
            div_texts = re.findall(r'<div[^>]*>(.*?)</div>', text_content, re.DOTALL)
            for dt in div_texts:
                clean = re.sub(r'<[^>]+>', '', dt).strip()
                if clean:
                    comments.append({
                        "type": "vml_note",
                        "text": clean
                    })
        except Exception:
            pass

    return comments


def extract_all_comments(zip_file):
    """Extract comments from all possible sources."""
    all_comments = []
    all_comments.extend(extract_cell_comments(zip_file))
    all_comments.extend(extract_threaded_comments(zip_file))
    all_comments.extend(extract_note_shapes(zip_file))
    return all_comments


# ─── Drawing Text Extraction ────────────────────────────

def get_full_text(elem):
    """Extract all text from an XML element, joining paragraphs with newline."""
    paragraphs = []
    for p in elem.findall(f'.//{NS_A}p'):
        runs = []
        for r in p.findall(f'{NS_A}r'):
            t = r.find(f'{NS_A}t')
            if t is not None and t.text:
                runs.append(t.text)
        if runs:
            paragraphs.append("".join(runs))
    return "\n".join(paragraphs)


def extract_all_texts(elem):
    """Recursively extract all text content from shapes in a drawing element."""
    texts = []
    tag = elem.tag.split('}')[-1]

    if tag == 'sp':
        text = get_full_text(elem).strip()
        if text:
            texts.append(text)

    elif tag in ('grpSp', 'oneCellAnchor', 'twoCellAnchor'):
        for child in elem:
            texts.extend(extract_all_texts(child))

    return texts


# ─── Classification ─────────────────────────────────────

def classify_texts(all_texts):
    """
    Classify extracted texts into:
      - assembly_steps: numbered instructions with optional cautions
      - parts: part names with part numbers
    """
    assembly_steps = []
    parts = []
    unclassified_notes = []

    current_step = None

    for text in all_texts:
        text = text.strip()
        if not text:
            continue

        if text.lower() in VIEW_LABELS:
            continue

        step_match = STEP_PATTERN.match(text)
        if step_match:
            if current_step:
                if unclassified_notes:
                    current_step["cautions"] = list(set(unclassified_notes))
                    unclassified_notes = []
                assembly_steps.append(current_step)

            step_num = int(step_match.group(1))
            instruction = STEP_PATTERN.sub('', text).strip().rstrip('.')
            current_step = {
                "step": step_num,
                "instruction": instruction,
            }
            continue

        pn_match = PART_NUMBER_PATTERN.search(text)
        if pn_match:
            lines = text.split('\n')
            part_number = None
            part_name_parts = []
            for line in lines:
                line = line.strip()
                if PART_NUMBER_PATTERN.search(line):
                    part_number = PART_NUMBER_PATTERN.search(line).group()
                else:
                    part_name_parts.append(line)

            part_name = " ".join(part_name_parts).strip() if part_name_parts else None
            if part_number:
                parts.append({
                    "name": part_name,
                    "part_number": part_number
                })
            continue

        is_likely_part_name = (
            len(text) < 50
            and not any(jp in text for jp in ['ください', 'してください', 'ること', 'します', 'ません'])
            and not text.endswith('。')
            and not text.endswith('.')
            and not STEP_PATTERN.match(text)
            and text.lower() not in VIEW_LABELS
        )

        if is_likely_part_name and '\n' not in text and not any(c in text for c in ['。', '、']):
            existing_names = {p.get("name", "").lower() for p in parts}
            if text.lower() not in existing_names:
                parts.append({
                    "name": text,
                    "part_number": None
                })
            continue

        if current_step:
            if "cautions" not in current_step:
                current_step["cautions"] = []
            if text not in current_step.get("cautions", []):
                current_step["cautions"].append(text)
        else:
            if text not in unclassified_notes:
                unclassified_notes.append(text)

    if current_step:
        if unclassified_notes:
            if "cautions" not in current_step:
                current_step["cautions"] = []
            current_step["cautions"].extend(unclassified_notes)
        assembly_steps.append(current_step)

    # Extract part names mentioned in quotes within instructions
    all_instruction_text = " ".join(
        s.get("instruction", "") + " ".join(s.get("cautions", []))
        for s in assembly_steps
    )
    # Match all quote styles: straight "...", curly "...", Japanese 「...」
    quoted_parts = re.findall(r'"([^"]+)"', all_instruction_text)
    quoted_parts += re.findall(r'\u201c([^\u201d]+)\u201d', all_instruction_text)
    quoted_parts += re.findall(r'\u300c([^\u300d]+)\u300d', all_instruction_text)

    existing_names_lower = {p["name"].lower() for p in parts if p.get("name")}
    for qp in quoted_parts:
        qp = qp.strip()
        if qp and qp.lower() not in existing_names_lower and len(qp) < 60:
            parts.append({
                "name": qp,
                "part_number": None
            })
            existing_names_lower.add(qp.lower())

    # Deduplicate parts
    seen = set()
    unique_parts = []
    for p in parts:
        key = (p.get("name", ""), p.get("part_number", ""))
        if key not in seen:
            seen.add(key)
            unique_parts.append(p)

    return assembly_steps, unique_parts


# ─── Main Processing ────────────────────────────────────

def process_excel(filepath):
    """Extract assembly steps, parts, and comments from a single Excel file."""
    filename = os.path.basename(filepath)

    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            all_files = z.namelist()

            drawing_files = [f for f in all_files
                             if f.startswith('xl/drawings/')
                             and f.endswith('.xml')
                             and '_rels' not in f]

            all_texts = []
            for df in drawing_files:
                drawing_xml = z.read(df).decode("utf-8")
                root = ET.fromstring(drawing_xml)
                for anchor in root:
                    anchor_tag = anchor.tag.split('}')[-1]
                    if 'Anchor' not in anchor_tag:
                        continue
                    all_texts.extend(extract_all_texts(anchor))

            comments = extract_all_comments(z)
            image_count = len([f for f in all_files if f.startswith('xl/media/')])

        assembly_steps, parts = classify_texts(all_texts) if all_texts else ([], [])

        result = {
            "file": filename,
            "assembly_steps": assembly_steps,
            "parts": parts,
            "comments": comments,
            "image_count": image_count,
        }

        return result

    except zipfile.BadZipFile:
        print(f"    ✗ Not a valid xlsx file")
        return None
    except Exception as e:
        print(f"    ✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    base_dir = get_base_dir()
    print(f"Working directory: {base_dir}")

    # Output goes to a subfolder named _extracted
    output_dir = os.path.join(base_dir, OUTPUT_FOLDER_NAME)

    # Find all .xlsx files in working directory, ONLY skip files inside the output folder
    output_dir_normalized = os.path.normpath(output_dir)
    excel_files = []
    for f in sorted(glob.glob(os.path.join(base_dir, "*.xlsx"))):
        # Only exclude files that are literally inside the output subfolder
        if os.path.normpath(os.path.dirname(f)) == output_dir_normalized:
            continue
        excel_files.append(f)

    if not excel_files:
        print(f"\nNo Excel files (.xlsx) found in: {base_dir}")
        print("Make sure .xlsx files are in the same folder as this script,")
        print("or run the script from the folder containing your Excel files.")
        exit(1)

    print(f"Found {len(excel_files)} Excel file(s)\n")

    os.makedirs(output_dir, exist_ok=True)

    all_results = []
    all_parts_rows = []

    for filepath in excel_files:
        filename = os.path.basename(filepath)
        print(f"  Processing: {filename}")

        result = process_excel(filepath)
        if result is None:
            continue

        json_path = os.path.join(output_dir, os.path.splitext(filename)[0] + ".json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        steps_count = len(result["assembly_steps"])
        parts_count = len(result["parts"])
        comments_count = len(result["comments"])
        print(f"    → {steps_count} steps, {parts_count} parts, {comments_count} comments, {result['image_count']} images")
        if comments_count > 0:
            for c in result["comments"]:
                preview = c["text"][:60] + "..." if len(c["text"]) > 60 else c["text"]
                cell = c.get("cell", "")
                author = c.get("author", "")
                label = f"[{cell}]" if cell else ""
                label += f" ({author})" if author else ""
                print(f"       💬 {label} {preview}")
        print(f"    → Saved: {os.path.basename(json_path)}")

        all_results.append(result)

        for p in result["parts"]:
            all_parts_rows.append({
                "source_file": filename,
                "part_name": p.get("name", ""),
                "part_number": p.get("part_number", ""),
            })

    # Save combined JSON
    if all_results:
        combined_path = os.path.join(output_dir, "_all_assembly_data.json")
        with open(combined_path, 'w', encoding='utf-8') as f:
            json.dump(all_results, f, ensure_ascii=False, indent=2)
        print(f"\n  Combined JSON: {combined_path}")

    # Save parts CSV
    if all_parts_rows:
        csv_path = os.path.join(output_dir, "_parts_summary.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=["source_file", "part_name", "part_number"])
            writer.writeheader()
            writer.writerows(all_parts_rows)
        print(f"  Parts CSV:     {csv_path}")

    # Save comments CSV
    all_comments_rows = []
    for r in all_results:
        for c in r.get("comments", []):
            all_comments_rows.append({
                "source_file": r["file"],
                "type": c.get("type", ""),
                "cell": c.get("cell", ""),
                "author": c.get("author", ""),
                "text": c.get("text", ""),
                "is_reply": c.get("is_reply", ""),
            })
    if all_comments_rows:
        csv_path = os.path.join(output_dir, "_comments_summary.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=["source_file", "type", "cell", "author", "text", "is_reply"])
            writer.writeheader()
            writer.writerows(all_comments_rows)
        print(f"  Comments CSV:  {csv_path}")

    print(f"\n{'='*50}")
    print(f"Done! {len(all_results)} file(s) processed.")
    print(f"Output saved to: {output_dir}")


if __name__ == "__main__":
    main()
