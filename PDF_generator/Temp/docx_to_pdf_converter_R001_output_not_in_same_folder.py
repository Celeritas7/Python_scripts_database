"""
DOCX to PDF Auto-Converter (Windows)
====================================
Automatically converts ALL .docx files in the same folder as this script.

Usage:
  Just double-click to run. No input needed.
  PDFs are saved alongside the original Word files.

Requirements:
  - Windows with MS Word installed
  - pip install docx2pdf
"""

import sys
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
    script_folder = Path(__file__).parent.resolve()
    print(f"Scanning: {script_folder}\n")
    
    # Find all .docx files
    docx_files = list(script_folder.glob("*.docx"))
    
    if not docx_files:
        print("[INFO] No .docx files found in this folder.")
        input("\nPress Enter to exit...")
        return
    
    print(f"Found {len(docx_files)} Word file(s)\n")
    print("-" * 50)
    
    # Convert each file
    success = 0
    failed = 0
    
    for docx_file in docx_files:
        output_path = docx_file.with_suffix('.pdf')
        print(f"Converting: {docx_file.name}")
        
        try:
            convert(str(docx_file), str(output_path))
            print(f"  → {output_path.name} ✓")
            success += 1
        except Exception as e:
            print(f"  [ERROR] {e}")
            failed += 1
        print()
    
    # Summary
    print("-" * 50)
    print(f"Done! Success: {success} | Failed: {failed}")
    print("=" * 50)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
