"""
PDF Compressor
==============
Compresses PDFs by re-rendering pages as optimized images.
Works reliably on any PDF — scanned docs, image-heavy notes, textbooks.

Usage:
    python pdf_compress.py                          # Auto-compress all PDFs in folder
    python pdf_compress.py input.pdf                # Single file
    python pdf_compress.py input.pdf -q 50          # Custom quality (1-95, default: 60)
    python pdf_compress.py input.pdf -d 150         # Custom DPI (default: 150)

Output:
    Original moved to "Old" folder, compressed version keeps original name.

Requirements:
    pip install pymupdf Pillow
"""

import argparse
import shutil
import sys
from pathlib import Path
from io import BytesIO

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF not installed. Run: pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow not installed. Run: pip install Pillow")
    sys.exit(1)


def compress_pdf(input_path, quality=60, dpi=150, backup=True):
    """Compress a PDF by re-rendering pages as JPEG images."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    original_size = input_path.stat().st_size / 1024  # KB

    # Open with PyMuPDF
    doc = fitz.open(str(input_path))
    num_pages = len(doc)

    # Render each page to image, then rebuild as PDF
    zoom = dpi / 72  # 72 is default PDF DPI
    mat = fitz.Matrix(zoom, zoom)

    pdf_images = []
    for i, page in enumerate(doc):
        # Render page to pixmap (image)
        pix = page.get_pixmap(matrix=mat)

        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Compress as JPEG in memory
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        buf.seek(0)

        # Reload as PIL image for PDF saving
        compressed_img = Image.open(buf)
        pdf_images.append(compressed_img)

        print(f"    Page {i+1}/{num_pages} rendered", end="\r")

    doc.close()

    # Backup original
    if backup:
        old_folder = input_path.parent / "Old"
        old_folder.mkdir(exist_ok=True)
        backup_path = old_folder / input_path.name
        # If backup already exists, add number
        counter = 1
        while backup_path.exists():
            backup_path = old_folder / f"{input_path.stem}_{counter}{input_path.suffix}"
            counter += 1
        shutil.move(str(input_path), str(backup_path))

    # Save as PDF using Pillow
    if pdf_images:
        pdf_images[0].save(
            str(input_path),
            "PDF",
            save_all=True,
            append_images=pdf_images[1:],
            resolution=dpi,
        )

    new_size = input_path.stat().st_size / 1024  # KB
    reduction = ((original_size - new_size) / original_size * 100) if original_size > 0 else 0

    print(f"  ✓ {input_path.name}" + " " * 20)
    print(f"    {original_size:.0f} KB → {new_size:.0f} KB ({reduction:.0f}% smaller, {num_pages} pages)")
    return True


def main():
    parser = argparse.ArgumentParser(description="Compress PDF files")
    parser.add_argument("files", nargs="*", help="PDF file(s) to compress")
    parser.add_argument("-q", "--quality", type=int, default=60,
                        help="JPEG quality 1-95 (default: 60, lower = smaller)")
    parser.add_argument("-d", "--dpi", type=int, default=150,
                        help="Resolution in DPI (default: 150, lower = smaller)")
    parser.add_argument("--no-backup", action="store_true",
                        help="Don't backup originals to Old folder")
    args = parser.parse_args()

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

    print(f"Quality: {args.quality} | DPI: {args.dpi}")
    print("-" * 50)

    for f in args.files:
        compress_pdf(f, quality=args.quality, dpi=args.dpi, backup=not args.no_backup)

    print("-" * 50)
    print("Done! Originals saved in 'Old' folder.")


if __name__ == "__main__":
    main()
