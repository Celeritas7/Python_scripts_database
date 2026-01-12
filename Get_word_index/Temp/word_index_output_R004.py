"""
Extract every visible paragraph from a Word (.docx) file
✓ Keeps #, punctuation, and order exactly as in Word
✓ Removes empty lines only
✓ Outputs plain text (ready for Notepad)
"""

import os
import sys
from docx import Document

# === 🪄 EDIT THIS PATH ===
WORD_FILE = r"C:\Users\manga\Desktop\Here.docx"

# === Output path ===
OUTPUT_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "TOC_Exact.txt")


def extract_visible_text(docx_path: str, out_path: str):
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"❌ File not found:\n{docx_path}")

    try:
        doc = Document(docx_path)
    except Exception as e:
        raise RuntimeError(f"⚠️ Could not open Word file.\n{e}")

    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        # Keep the paragraph exactly as-is
        lines.append(text)

    if not lines:
        raise RuntimeError("⚠️ No visible text found in the document.")

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

    print(f"✅ Exported exact visible text to:\n{out_path}")


if __name__ == "__main__":
    try:
        extract_visible_text(WORD_FILE, OUTPUT_FILE)
    except Exception as e:
        print(f"\n{e}")
        input("\nPress Enter to close...")
        sys.exit(1)

    print("\n🎉 Done! TOC_Exact.txt created on Desktop.")
    input("\nPress Enter to close...")
