"""
PDF Merger
==========
Combines all PDFs in a folder into a single PDF, ordered by filename.

Usage:
    python pdf_merge.py                             # All PDFs in folder
    python pdf_merge.py file1.pdf file2.pdf         # Specific files
    python pdf_merge.py -o combined.pdf             # Custom output name

Output:
    combined_output.pdf in the same folder.

Requirements:
    pip install pypdf
"""

import re
import argparse
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)


def natural_sort_key(filepath):
    """Natural sort: Notes (2) before Notes (10)."""
    parts = re.split(r'(\d+)', Path(filepath).name.lower())
    return [int(p) if p.isdigit() else p for p in parts]


def merge_pdfs(files, output_path):
    """Merge multiple PDFs into one."""
    writer = PdfWriter()
    total_pages = 0

    for f in files:
        f = Path(f)
        print(f"  + {f.name}", end="")
        try:
            reader = PdfReader(str(f))
            pages = len(reader.pages)
            for page in reader.pages:
                writer.add_page(page)
            total_pages += pages
            print(f"  ({pages} pages)")
        except Exception as e:
            print(f"  [ERROR] {e}")

    if total_pages == 0:
        print("\nNo pages to merge.")
        return False

    output_path = Path(output_path)
    with open(str(output_path), "wb") as f:
        writer.write(f)

    size_kb = output_path.stat().st_size / 1024
    print(f"\n  ✓ {output_path.name} ({total_pages} pages, {size_kb:.0f} KB)")
    return True


def main():
    parser = argparse.ArgumentParser(description="Merge multiple PDFs into one")
    parser.add_argument("files", nargs="*", help="PDF files to merge (in order)")
    parser.add_argument("-o", "--output", default=None,
                        help="Output filename (default: combined_output.pdf)")
    args = parser.parse_args()

    # If no files, auto-find PDFs in working directory
    if not args.files:
        script_dir = Path(__file__).parent.resolve()
        work_dir = Path.cwd()
        if work_dir != script_dir:
            search_dir = work_dir
        else:
            search_dir = script_dir

        args.files = sorted(
            [str(f) for f in search_dir.glob("*.pdf")
             if f.name != "combined_output.pdf"],
            key=natural_sort_key
        )
        if not args.files:
            print(f"No PDF files found in: {search_dir}")
            return

        if args.output is None:
            args.output = str(search_dir / "combined_output.pdf")

        print(f"Found {len(args.files)} PDF(s) in: {search_dir}")
    else:
        if args.output is None:
            args.output = str(Path(args.files[0]).parent / "combined_output.pdf")

    print(f"Output: {Path(args.output).name}")
    print("-" * 50)

    merge_pdfs(args.files, args.output)

    print("-" * 50)
    print("Done!")


if __name__ == "__main__":
    main()
