"""
PDF to Images Extractor
=======================
Extracts each page of a PDF as a JPG/PNG image into a folder
named after the PDF file.

Usage:
    python pdf_to_images.py                         # All PDFs in folder
    python pdf_to_images.py input.pdf               # Single file
    python pdf_to_images.py input.pdf -d 200        # Custom DPI (default: 150)
    python pdf_to_images.py input.pdf -f png        # PNG instead of JPG
    python pdf_to_images.py input.pdf -q 80         # JPEG quality (default: 85)

Output:
    input.pdf → input/
                  page_001.jpg
                  page_002.jpg
                  ...

Requirements:
    pip install pymupdf Pillow
"""

import argparse
import sys
from pathlib import Path

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


def extract_pages(input_path, dpi=150, fmt="jpg", quality=85):
    """Extract PDF pages as images into a folder named after the file."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    # Create output folder (same name as PDF without extension)
    output_folder = input_path.parent / input_path.stem
    output_folder.mkdir(exist_ok=True)

    doc = fitz.open(str(input_path))
    num_pages = len(doc)
    pad = len(str(num_pages))  # zero-pad width

    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)

    print(f"  {input_path.name} → {output_folder.name}/  ({num_pages} pages, {dpi} DPI)")

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        filename = f"page_{str(i+1).zfill(pad)}.{fmt}"
        out_path = output_folder / filename

        if fmt == "jpg":
            img.save(str(out_path), "JPEG", quality=quality, optimize=True)
        else:
            img.save(str(out_path), "PNG", optimize=True)

        print(f"    Page {i+1}/{num_pages}", end="\r")

    doc.close()

    total_size = sum(f.stat().st_size for f in output_folder.glob(f"*.{fmt}")) / 1024
    print(f"    ✓ {num_pages} images saved ({total_size:.0f} KB total)" + " " * 20)
    return True


def main():
    parser = argparse.ArgumentParser(description="Extract PDF pages as images")
    parser.add_argument("files", nargs="*", help="PDF file(s) to extract")
    parser.add_argument("-d", "--dpi", type=int, default=150,
                        help="Resolution in DPI (default: 150)")
    parser.add_argument("-f", "--format", choices=["jpg", "png"], default="jpg",
                        help="Image format (default: jpg)")
    parser.add_argument("-q", "--quality", type=int, default=85,
                        help="JPEG quality 1-95 (default: 85)")
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

    print(f"DPI: {args.dpi} | Format: {args.format} | Quality: {args.quality}")
    print("-" * 50)

    for f in args.files:
        extract_pages(f, dpi=args.dpi, fmt=args.format, quality=args.quality)

    print("-" * 50)
    print("Done!")


if __name__ == "__main__":
    main()
