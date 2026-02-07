"""
DOCX to PDF Converter (Windows)
===============================
Uses MS Word via docx2pdf for perfect margin/layout preservation.

Usage:
  1. Drag & drop: Drag .docx file(s) onto this script
  2. Command line: python docx_to_pdf_converter.py file.docx
  3. Batch folder: python docx_to_pdf_converter.py "C:/path/to/folder"
  4. Interactive: Double-click to run and enter path manually

Requirements:
  - Windows with MS Word installed
  - pip install docx2pdf
"""

import sys
import os
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


def convert_file(input_path: Path, output_path: Path = None):
    """Convert a single DOCX file to PDF."""
    if not input_path.exists():
        print(f"  [ERROR] File not found: {input_path}")
        return False
    
    if input_path.suffix.lower() != '.docx':
        print(f"  [SKIP] Not a .docx file: {input_path.name}")
        return False
    
    if output_path is None:
        output_path = input_path.with_suffix('.pdf')
    
    print(f"  Converting: {input_path.name}")
    print(f"  Output:     {output_path.name}")
    
    try:
        convert(str(input_path), str(output_path))
        print(f"  [SUCCESS] Created: {output_path}")
        return True
    except Exception as e:
        print(f"  [ERROR] Conversion failed: {e}")
        return False


def convert_folder(folder_path: Path, output_folder: Path = None):
    """Convert all DOCX files in a folder."""
    if not folder_path.exists():
        print(f"[ERROR] Folder not found: {folder_path}")
        return 0, 0
    
    docx_files = list(folder_path.glob("*.docx"))
    
    if not docx_files:
        print(f"[INFO] No .docx files found in: {folder_path}")
        return 0, 0
    
    print(f"\nFound {len(docx_files)} DOCX file(s) in: {folder_path}\n")
    
    if output_folder is None:
        output_folder = folder_path / "PDF_Output"
    
    output_folder.mkdir(exist_ok=True)
    print(f"Output folder: {output_folder}\n")
    
    success_count = 0
    fail_count = 0
    
    for docx_file in docx_files:
        output_path = output_folder / docx_file.with_suffix('.pdf').name
        if convert_file(docx_file, output_path):
            success_count += 1
        else:
            fail_count += 1
        print()
    
    return success_count, fail_count


def main():
    print("=" * 50)
    print("       DOCX to PDF Converter (Windows)")
    print("       Preserves margins & formatting")
    print("=" * 50)
    print()
    
    # Get input paths from command line args or user input
    if len(sys.argv) > 1:
        # Command line / drag-drop mode
        input_paths = [Path(arg) for arg in sys.argv[1:]]
    else:
        # Interactive mode
        print("Enter path to DOCX file or folder:")
        print("(You can also drag & drop files onto this script)\n")
        user_input = input("Path: ").strip().strip('"').strip("'")
        
        if not user_input:
            print("[ERROR] No path provided.")
            input("\nPress Enter to exit...")
            return
        
        input_paths = [Path(user_input)]
    
    total_success = 0
    total_fail = 0
    
    for input_path in input_paths:
        print(f"\nProcessing: {input_path}\n")
        
        if input_path.is_dir():
            # Batch convert folder
            success, fail = convert_folder(input_path)
            total_success += success
            total_fail += fail
        elif input_path.is_file():
            # Single file conversion
            if convert_file(input_path):
                total_success += 1
            else:
                total_fail += 1
        else:
            print(f"[ERROR] Path not found: {input_path}")
            total_fail += 1
    
    # Summary
    print("\n" + "=" * 50)
    print("CONVERSION COMPLETE")
    print(f"  Success: {total_success}")
    print(f"  Failed:  {total_fail}")
    print("=" * 50)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
