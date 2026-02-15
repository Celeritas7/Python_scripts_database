"""
Extract a clean, indented Table of Contents (TOC) from a Word (.docx) file.
✓ Keeps page numbers aligned
✓ Removes dotted leaders
✓ Saves as a plain .txt file (ready for Notepad)
"""

import os
import re
import sys
from docx import Document

# === 🪄 EDIT THIS PATH ONLY ===
WORD_FILE = r"C:\Users\manga\OneDrive\####Mind_Palace\####Technical\##AI\3_Machine_learning\Chapter_wise\Chapter_4_5_machine_learning.docx"

# === Output path ===
OUTPUT_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "TOC_Indented.txt")

# === Spaces per indent level ===
INDENT_SPACES = 4

# === Which mode to use: "auto", "headings", or "toc" ===
MODE = "auto"  # "auto" tries both Headings and TOC levels

# === Regex patterns for matching Word styles ===
HEADING_RE = re.compile(r"^Heading\s+(\d+)$", re.IGNORECASE)
TOC_RE = re.compile(r"^TOC\s+(\d+)$", re.IGNORECASE)


def clean_toc_joined(text: str) -> str:
    """
    Clean a TOC line AFTER joining runs.
    Removes leader dots but keeps trailing page numbers.
    Also collapses multiple spaces.
    """
    text = re.sub(r"\.{3,}", "", text)     # remove leader dots
    text = re.sub(r"\s{2,}", " ", text)    # collapse spaces
    return text.strip()


def detect_level(style_name: str, mode: str):
    """Return heading or TOC level number."""
    style_name = style_name or ""
    if mode in ("auto", "headings"):
        m = HEADING_RE.match(style_name)
        if m:
            return int(m.group(1))
    if mode in ("auto", "toc"):
        m = TOC_RE.match(style_name)
        if m:
            return int(m.group(1))
    return None


def extract_toc(docx_path: str, out_path: str, indent_spaces: int, mode: str):
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"❌ File not found:\n{docx_path}")

    try:
        doc = Document(docx_path)
    except Exception as e:
        raise RuntimeError(f"⚠️ Could not open Word file.\n{e}")

    lines = []
    for p in doc.paragraphs:
        style = getattr(p.style, "name", "")
        level = detect_level(style, mode)

        # Treat TOC 2 as top level if your doc uses only TOC 2
        if "TOC" in style and level == 2:
            level = 1

        if level is None:
            continue

        # --- Build text from RUNS to preserve page numbers (critical fix) ---
        run_parts = []
        for r in p.runs:
            t = r.text
            if t:
                run_parts.append(t.strip())
        joined = " ".join(run_parts).strip()
        if not joined:
            continue

        text = clean_toc_joined(joined)
        if not text:
            continue

        indent = " " * ((max(level, 1) - 1) * indent_spaces)
        lines.append(f"{indent}{text}")

    if not lines:
        raise RuntimeError("⚠️ No TOC or Heading lines found in this file.")

    # --- Write to file ---
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

    print(f"✅ TOC exported successfully:\n{out_path}")


if __name__ == "__main__":
    try:
        extract_toc(WORD_FILE, OUTPUT_FILE, INDENT_SPACES, MODE)
    except Exception as e:
        print(f"\n{e}")
        input("\nPress Enter to close...")
        sys.exit(1)

    print("\n🎉 Done! TOC_Indented.txt created on Desktop.")
    input("\nPress Enter to close...")
