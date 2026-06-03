"""
PDF Compressor (Ghostscript)
============================
Compresses PDFs by optimizing images while keeping text searchable.

Usage:
    python pdf_compress.py                          # Auto-compress all PDFs in folder
    python pdf_compress.py input.pdf                # Single file
    python pdf_compress.py input.pdf -p screen      # Aggressive compression

Presets:
    screen   → 72 DPI  (smallest, good for screen viewing)
    ebook    → 150 DPI (good balance — default)
    printer  → 300 DPI (high quality, moderate compression)

Output:
    Original moved to "Old" folder, compressed version keeps original name.

Requirements:
    Ghostscript installed (choco install ghostscript)
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_ghostscript():
    """Find Ghostscript executable."""
    for cmd in ["gswin64c", "gswin32c", "gs"]:
        try:
            subprocess.run([cmd, "--version"], capture_output=True, check=True)
            return cmd
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
    return None


def compress_pdf(input_path, preset="ebook", backup=True, gs_cmd="gswin64c"):
    """Compress a PDF using Ghostscript."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    original_size = input_path.stat().st_size / 1024  # KB

    # Write to temp file first (Ghostscript can't overwrite input)
    temp_fd, temp_path = tempfile.mkstemp(suffix=".pdf")
    os.close(temp_fd)  # Close file descriptor so Ghostscript can write to it

    try:
        args = [
            gs_cmd,
            "-sDEVICE=pdfwrite",
            f"-dPDFSETTINGS=/{preset}",
            "-dNOPAUSE",
            "-dBATCH",
            "-dQUIET",
            "-dCompatibilityLevel=1.4",
            f"-sOutputFile={temp_path}",
            str(input_path),
        ]

        result = subprocess.run(args, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"  Error: {input_path.name} — {result.stderr.strip()}")
            return False

        new_size = Path(temp_path).stat().st_size / 1024

        # Only replace if actually smaller
        if new_size >= original_size:
            print(f"  ⊘ {input_path.name}")
            print(f"    {original_size:.0f} KB → {new_size:.0f} KB (no improvement, skipped)")
            return True

        # Backup original
        if backup:
            old_folder = input_path.parent / "Old"
            old_folder.mkdir(exist_ok=True)
            backup_path = old_folder / input_path.name
            counter = 1
            while backup_path.exists():
                backup_path = old_folder / f"{input_path.stem}_{counter}{input_path.suffix}"
                counter += 1
            shutil.move(str(input_path), str(backup_path))

        # Replace with compressed version
        shutil.copy2(temp_path, str(input_path))

        reduction = ((original_size - new_size) / original_size * 100)
        print(f"  ✓ {input_path.name}")
        print(f"    {original_size:.0f} KB → {new_size:.0f} KB ({reduction:.0f}% smaller)")

    finally:
        Path(temp_path).unlink(missing_ok=True)

    return True


def main():
    parser = argparse.ArgumentParser(description="Compress PDF files using Ghostscript")
    parser.add_argument("files", nargs="*", help="PDF file(s) to compress")
    parser.add_argument("-p", "--preset", choices=["screen", "ebook", "printer"],
                        default="ebook",
                        help="Quality preset (default: ebook)")
    parser.add_argument("--no-backup", action="store_true",
                        help="Don't backup originals to Old folder")
    args = parser.parse_args()

    # Find Ghostscript
    gs_cmd = find_ghostscript()
    if not gs_cmd:
        print("ERROR: Ghostscript not found.")
        print("Install: choco install ghostscript")
        sys.exit(1)

    # If no files, auto-find PDFs in working directory
    if not args.files:
        script_dir = Path(__file__).parent.resolve()
        work_dir = Path.cwd()
        if work_dir != script_dir:
            search_dir = work_dir
        else:
            search_dir = script_dir

        args.files = sorted(str(f) for f in search_dir.glob("*.pdf"))
        if not args.files:
            print(f"No PDF files found in: {search_dir}")
            return
        print(f"Found {len(args.files)} PDF(s) in: {search_dir}")

    print(f"Preset: {args.preset} (screen=smallest, ebook=balanced, printer=best quality)")
    print("-" * 50)

    for f in args.files:
        compress_pdf(f, preset=args.preset, backup=not args.no_backup, gs_cmd=gs_cmd)

    print("-" * 50)
    print("Done! Originals saved in 'Old' folder.")


if __name__ == "__main__":
    main()
