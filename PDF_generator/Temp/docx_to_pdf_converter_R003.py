"""
DOCX to PDF Auto-Converter (Windows)
====================================
Automatically converts ALL .docx files in the same folder as this script.

Usage:
  Just double-click to run. No input needed.
  PDFs are saved in a PDF_Outputs/ subfolder.

Requirements:
  - Windows with MS Word installed
  - pip install docx2pdf
"""

import sys
import shutil
from pathlib import Path

try:
    from docx2pdf import convert
except ImportError:
    print("=" * 50)
    print("ERROR: docx2pdf not installed!")
    print("Run: pip install docx2pdf")
    print("=" * 50)
    input("Press Enter to exit...")
    sys.exit(1)


def main():
    print("=" * 50)
    print("    DOCX to PDF Auto-Converter")
    print("=" * 50)
    print()

    # Get script's folder
    script_dir = Path(__file__).parent.resolve()
    work_dir = Path.cwd()
    if work_dir != script_dir:
        script_dir = work_dir

    print(f"Scanning: {script_dir}\n")

    # Find all .docx files
    docx_files = list(script_dir.glob("*.docx"))

    if not docx_files:
        print("[INFO] No .docx files found in this folder.")
        input("\nPress Enter to exit...")
        return

    # Create output folder
    output_dir = script_dir / "PDF_Outputs"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Found {len(docx_files)} Word file(s)")
    print(f"Output: {output_dir}\n")
    print("-" * 50)

    # Convert each file
    # NOTE: docx2pdf uses Word COM which fails on SaveAs to a different folder.
    # Workaround: convert in-place first, then move the PDF to output folder.
    success = 0
    failed = 0

    for docx_file in docx_files:
        temp_pdf = docx_file.with_suffix('.pdf')       # same folder as .docx
        final_pdf = output_dir / temp_pdf.name          # PDF_Outputs/ folder
        print(f"Converting: {docx_file.name}")

        try:
            convert(str(docx_file), str(temp_pdf))
            shutil.move(str(temp_pdf), str(final_pdf))
            print(f"  -> {final_pdf.name} done")
            success += 1
        except Exception as e:
            print(f"  [ERROR] {e}")
            # Clean up temp PDF if it was created but move failed
            if temp_pdf.exists():
                try:
                    shutil.move(str(temp_pdf), str(final_pdf))
                except:
                    pass
            failed += 1
        print()

    # Summary
    print("-" * 50)
    print(f"Done! Success: {success} | Failed: {failed}")
    print(f"PDFs saved in: {output_dir}")
    print("=" * 50)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
