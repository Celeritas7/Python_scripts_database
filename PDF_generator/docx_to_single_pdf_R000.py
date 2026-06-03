"""
Word Files to Single PDF
========================
Converts all .docx/.docm files in a folder to PDFs, then merges
them into a single combined PDF, naturally sorted by filename.

Usage:
    python docx_to_single_pdf.py                    # All Word files in folder
    python docx_to_single_pdf.py -o combined.pdf    # Custom output name

Output:
    combined_output.pdf in the same folder.
    Individual PDFs saved in a temp subfolder (deleted after merge).

Requirements:
    - LibreOffice installed (https://www.libreoffice.org/download/)
      OR MS Word installed + pip install docx2pdf
    - pip install pypdf
"""

import re
import argparse
import subprocess
import shutil
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)


def natural_sort_key(filepath):
    """Natural sort: Chapter 2 before Chapter 10."""
    parts = re.split(r'(\d+)', Path(filepath).name.lower())
    return [int(p) if p.isdigit() else p for p in parts]


def find_libreoffice():
    """Find LibreOffice executable."""
    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in common_paths:
        if Path(p).exists():
            return p
    try:
        result = subprocess.run(["soffice", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            return "soffice"
    except FileNotFoundError:
        pass
    return None


def convert_with_libreoffice(docx_files, output_dir, soffice_path):
    """Convert Word files to PDF using LibreOffice."""
    pdf_paths = []
    for docx_file in docx_files:
        print(f"  Converting: {docx_file.name}")
        try:
            subprocess.run(
                [soffice_path, "--headless", "--convert-to", "pdf",
                 "--outdir", str(output_dir), str(docx_file)],
                capture_output=True, text=True, timeout=120
            )
            pdf_path = output_dir / docx_file.with_suffix(".pdf").name
            if pdf_path.exists():
                pdf_paths.append(pdf_path)
                print(f"    ✓ done")
            else:
                print(f"    ✗ failed")
        except Exception as e:
            print(f"    ✗ {e}")
    return pdf_paths


def convert_with_word(docx_files, output_dir):
    """Convert Word files to PDF using MS Word."""
    try:
        from docx2pdf import convert
    except ImportError:
        print("ERROR: docx2pdf not installed. Run: pip install docx2pdf")
        return []

    pdf_paths = []
    for docx_file in docx_files:
        print(f"  Converting: {docx_file.name}")
        try:
            pdf_path = output_dir / docx_file.with_suffix(".pdf").name
            # Convert to temp location first, then move
            temp_pdf = docx_file.with_suffix(".pdf")
            convert(str(docx_file), str(temp_pdf))
            shutil.move(str(temp_pdf), str(pdf_path))
            pdf_paths.append(pdf_path)
            print(f"    ✓ done")
        except Exception as e:
            print(f"    ✗ {e}")
    return pdf_paths


def merge_pdfs(pdf_files, output_path):
    """Merge multiple PDFs into one."""
    writer = PdfWriter()
    total_pages = 0

    for f in pdf_files:
        reader = PdfReader(str(f))
        pages = len(reader.pages)
        for page in reader.pages:
            writer.add_page(page)
        total_pages += pages

    with open(str(output_path), "wb") as f:
        writer.write(f)

    return total_pages


def main():
    print("=" * 50)
    print("    Word Files → Single PDF")
    print("=" * 50)
    print()

    parser = argparse.ArgumentParser(description="Convert all Word files to a single merged PDF")
    parser.add_argument("-o", "--output", default=None, help="Output PDF name")
    parser.add_argument("--keep-individual", action="store_true",
                        help="Keep individual PDFs in PDF_Outputs folder")
    args = parser.parse_args()

    # Dual-mode
    script_dir = Path(__file__).parent.resolve()
    work_dir = Path.cwd()
    if work_dir != script_dir:
        search_dir = work_dir
    else:
        search_dir = script_dir

    # Find all Word files
    word_files = sorted(
        [f for f in search_dir.iterdir()
         if f.suffix.lower() in (".docx", ".docm") and not f.name.startswith("~")],
        key=natural_sort_key
    )

    if not word_files:
        print(f"No .docx/.docm files found in: {search_dir}")
        input("\nPress Enter to exit...")
        return

    # Output path
    if args.output:
        output_path = search_dir / args.output
    else:
        output_path = search_dir / "combined_output.pdf"

    print(f"Found {len(word_files)} Word file(s) in: {search_dir}")
    print(f"Output: {output_path.name}\n")

    # File list
    for i, f in enumerate(word_files, 1):
        print(f"  {i:2}) {f.name}")
    print()

    # Step 1: Convert to individual PDFs
    temp_dir = search_dir / "_temp_pdf_conversion"
    temp_dir.mkdir(exist_ok=True)

    soffice = find_libreoffice()
    if soffice:
        print(f"Step 1: Converting to PDF (LibreOffice)")
    else:
        print(f"Step 1: Converting to PDF (MS Word)")
    print("-" * 50)

    if soffice:
        pdf_files = convert_with_libreoffice(word_files, temp_dir, soffice)
    else:
        pdf_files = convert_with_word(word_files, temp_dir)

    if not pdf_files:
        print("\nNo PDFs were generated. Cannot merge.")
        shutil.rmtree(temp_dir, ignore_errors=True)
        input("\nPress Enter to exit...")
        return

    # Sort PDFs in same natural order
    pdf_files = sorted(pdf_files, key=natural_sort_key)

    # Step 2: Merge
    print()
    print(f"Step 2: Merging {len(pdf_files)} PDFs")
    print("-" * 50)

    total_pages = merge_pdfs(pdf_files, output_path)
    size_kb = output_path.stat().st_size / 1024

    # Cleanup or keep individual PDFs
    if args.keep_individual:
        final_dir = search_dir / "PDF_Outputs"
        if temp_dir != final_dir:
            if final_dir.exists():
                shutil.rmtree(final_dir)
            temp_dir.rename(final_dir)
        print(f"  Individual PDFs kept in: PDF_Outputs/")
    else:
        shutil.rmtree(temp_dir, ignore_errors=True)

    # Summary
    print()
    print("=" * 50)
    print(f"  ✓ {output_path.name}")
    print(f"    {len(pdf_files)} files | {total_pages} pages | {size_kb:.0f} KB")
    print("=" * 50)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
