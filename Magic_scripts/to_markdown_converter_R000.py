"""
Convert to Markdown (Auto-Converter)
====================================
Automatically converts ALL supported files in the current folder to clean Markdown.
Markdown output is meant for feeding documents to LLMs (Claude, etc.) — clean text,
not pretty layout. Each .md is saved RIGHT NEXT TO its source file.

Usage:
  Just run it (via your launcher, or double-click). No input needed.
  After files are done, you can optionally paste a YouTube/web URL to convert too.

Supports:
  Documents : .pdf .docx .docm .pptx .xlsx .xls .epub
  Web/data  : .html .htm .csv .json .xml
  Plus      : any YouTube / web URL pasted at the prompt

Requirements:
  pip install "markitdown[all]"

Notes:
  - PDFs must have a real text layer (selectable text). markitdown has NO built-in
    OCR, so scanned/photo PDFs come out empty — OCR them first (e.g. ocrmypdf).
  - Images are intentionally NOT converted here: markitdown only pulls EXIF from
    them, not content. For images, just upload them straight to Claude.
"""

import re
from pathlib import Path
from markitdown import MarkItDown

SUPPORTED = {
    ".pdf", ".docx", ".docm", ".pptx", ".xlsx", ".xls", ".epub",
    ".html", ".htm", ".csv", ".json", ".xml",
}


def get_base_dir():
    """
    Run-anywhere detection (same pattern as your other scripts):
    - Launched from another folder (via launcher/tool) -> operate on that folder.
    - Run directly from its own folder -> operate on the script's folder.
    """
    script_dir = Path(__file__).parent.resolve()
    work_dir = Path.cwd()
    return work_dir if work_dir != script_dir else script_dir


def url_filename(url):
    """Pick a sensible .md filename for a URL."""
    m = re.search(r"[?&]v=([\w-]+)", url) or re.search(r"youtu\.be/([\w-]+)", url)
    if m:
        return f"youtube_{m.group(1)}.md"
    m = re.search(r"https?://([^/]+)", url)
    domain = m.group(1).replace("www.", "").replace(".", "_") if m else "url"
    return f"{domain}.md"


def main():
    print("=" * 50)
    print("    Convert to Markdown (for LLMs)")
    print("=" * 50)
    print()

    base_dir = get_base_dir()
    print(f"Scanning: {base_dir}\n")

    files = sorted(
        f for f in base_dir.iterdir()
        if f.is_file()
        and f.suffix.lower() in SUPPORTED
        and not f.name.startswith("~")
        and f.suffix.lower() != ".md"
    )

    md = MarkItDown(enable_plugins=False)
    success = failed = 0

    if not files:
        print("[INFO] No convertible files (.pdf .docx .pptx .xlsx ...) in this folder.")
    else:
        print(f"Found {len(files)} file(s)")
        print("-" * 50)
        for f in files:
            print(f"Converting: {f.name}")
            try:
                text = md.convert(str(f)).text_content
                out = f.with_suffix(".md")
                out.write_text(text, encoding="utf-8")
                if len(text) == 0:
                    print(f"  [WARN] {out.name} is EMPTY — likely a scanned PDF (no text layer)")
                else:
                    print(f"  -> {out.name} ({len(text):,} chars)")
                success += 1
            except Exception as e:
                print(f"  [ERROR] {e}")
                failed += 1
        print("-" * 50)
        print(f"Done! Success: {success} | Failed: {failed}")

    # Optional: convert a YouTube / web URL (no file needed)
    print()
    url = input("Paste a YouTube/web URL to convert (or press Enter to exit): ").strip()
    if url:
        print(f"\nConverting: {url}")
        try:
            text = md.convert(url).text_content
            out = base_dir / url_filename(url)
            out.write_text(text, encoding="utf-8")
            print(f"  -> {out.name} ({len(text):,} chars)")
        except Exception as e:
            print(f"  [ERROR] {e}")

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
