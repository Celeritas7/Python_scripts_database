"""
DOCX to PDF Auto-Converter
===========================
Automatically converts ALL .docx files in the current folder.

Usage:
  Just double-click to run. No input needed.
  PDFs are saved in a PDF_Outputs/ subfolder.

Supports two backends:
  1. LibreOffice (free) — preferred
  2. MS Word via docx2pdf — fallback if Word is installed

Requirements:
  - LibreOffice installed (https://www.libreoffice.org/download/)
  - OR Windows with MS Word installed + pip install docx2pdf
"""

import subprocess
import sys
import shutil
from pathlib import Path


def find_libreoffice():
    """Find LibreOffice executable on Windows."""
    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in common_paths:
        if Path(p).exists():
            return p

    # Try PATH
    try:
        result = subprocess.run(["soffice", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            return "soffice"
    except FileNotFoundError:
        pass

    return None


def convert_with_libreoffice(docx_files, output_dir, soffice_path):
    """Convert docx files to PDF using LibreOffice."""
    success = 0
    failed = 0

    for docx_file in docx_files:
        print(f"Converting: {docx_file.name}")
        try:
            result = subprocess.run(
                [soffice_path, "--headless", "--convert-to", "pdf",
                 "--outdir", str(output_dir), str(docx_file)],
                capture_output=True, text=True, timeout=120
            )
            pdf_name = docx_file.with_suffix(".pdf").name
            pdf_path = output_dir / pdf_name

            if pdf_path.exists():
                size_kb = pdf_path.stat().st_size / 1024
                print(f"  -> {pdf_name} ({size_kb:.0f} KB)")
                success += 1
            else:
                print(f"  [ERROR] PDF not created")
                if result.stderr:
                    print(f"  {result.stderr.strip()}")
                failed += 1
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout — file took too long")
            failed += 1
        except Exception as e:
            print(f"  [ERROR] {e}")
            failed += 1

    return success, failed


def convert_with_word(docx_files, output_dir):
    """Convert docx files to PDF using MS Word (docx2pdf)."""
    try:
        from docx2pdf import convert
    except ImportError:
        print("ERROR: docx2pdf not installed. Run: pip install docx2pdf")
        return 0, len(docx_files)

    success = 0
    failed = 0

    for docx_file in docx_files:
        temp_pdf = docx_file.with_suffix(".pdf")
        final_pdf = output_dir / temp_pdf.name
        print(f"Converting: {docx_file.name}")

        try:
            convert(str(docx_file), str(temp_pdf))
            shutil.move(str(temp_pdf), str(final_pdf))
            size_kb = final_pdf.stat().st_size / 1024
            print(f"  -> {final_pdf.name} ({size_kb:.0f} KB)")
            success += 1
        except Exception as e:
            print(f"  [ERROR] {e}")
            if temp_pdf.exists():
                try:
                    shutil.move(str(temp_pdf), str(final_pdf))
                except:
                    pass
            failed += 1

    return success, failed


def main():
    print("=" * 50)
    print("    DOCX to PDF Auto-Converter")
    print("=" * 50)
    print()

    # Dual-mode: tools launcher or direct run
    script_dir = Path(__file__).parent.resolve()
    work_dir = Path.cwd()
    if work_dir != script_dir:
        script_dir = work_dir

    print(f"Scanning: {script_dir}\n")

    # Find all .docx files
    docx_files = sorted(script_dir.glob("*.docx"))

    if not docx_files:
        print("[INFO] No .docx files found in this folder.")
        input("\nPress Enter to exit...")
        return

    # Create output folder
    output_dir = script_dir / "PDF_Outputs"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Found {len(docx_files)} Word file(s)")
    print(f"Output: {output_dir}\n")

    # Choose backend
    soffice = find_libreoffice()
    if soffice:
        print(f"Using: LibreOffice")
    else:
        print(f"Using: MS Word (LibreOffice not found)")

    print("-" * 50)

    if soffice:
        success, failed = convert_with_libreoffice(docx_files, output_dir, soffice)
    else:
        success, failed = convert_with_word(docx_files, output_dir)

    # Summary
    print("-" * 50)
    print(f"Done! Success: {success} | Failed: {failed}")
    print(f"PDFs saved in: {output_dir}")
    print("=" * 50)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
