from docx import Document
import os, re

WORD_FILE = r"C:\Users\manga\OneDrive\#Coding_project_files\Python_own_projects\Get_word_index\Here.docx"

OUTPUT_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "TOC_Indented.txt")

def extract_toc_indented(docx_path, out_path, indent_spaces=4):
    doc = Document(docx_path)
    lines = []

    for p in doc.paragraphs:
        style = getattr(p.style, "name", "")
        text = p.text.strip()
        if not text:
            continue

        # Match both TOC and Heading levels (e.g. "TOC 2", "Heading 3")
        match = re.match(r"^(TOC|Heading)\s*(\d+)", style, re.IGNORECASE)
        if match:
            level = int(match.group(2))
        else:
            level = 1  # default if style not numbered

        indent = " " * ((level - 1) * indent_spaces)
        lines.append(f"{indent}{text}")

    if not lines:
        raise RuntimeError("No TOC or Heading entries found.")

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

if __name__ == "__main__":
    extract_toc_indented(WORD_FILE, OUTPUT_FILE)
