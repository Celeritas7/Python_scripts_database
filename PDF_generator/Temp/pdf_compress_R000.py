"""
PDF Image Compressor
====================
Compresses images inside PDF files to reduce file size.

Usage:
    python pdf_compress.py                          # Auto-compress all PDFs in folder
    python pdf_compress.py input.pdf                # Single file
    python pdf_compress.py input.pdf -q 50          # Custom quality (1-95, default: 60)

Output:
    Single PDF  → same name (original moved to "Old" folder)
    Multiple    → same names (originals moved to "Old" folder)

Requirements:
    pip install pypdf Pillow
"""

import argparse
import shutil
import sys
from pathlib import Path
from io import BytesIO

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow not installed. Run: pip install Pillow")
    sys.exit(1)


def compress_pdf(input_path, quality=60, backup=True):
    """Compress images in a PDF file."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    original_size = input_path.stat().st_size / 1024  # KB

    reader = PdfReader(str(input_path))
    writer = PdfWriter()

    # Copy all pages
    for page in reader.pages:
        writer.add_page(page)

    # Compress images
    image_count = 0
    for page in writer.pages:
        if "/XObject" not in page.get("/Resources", {}):
            continue
        x_objects = page["/Resources"]["/XObject"].get_object()
        for obj_name in x_objects:
            obj = x_objects[obj_name].get_object()
            if obj.get("/Subtype") == "/Image":
                try:
                    # Get image dimensions
                    width = obj["/Width"]
                    height = obj["/Height"]

                    # Extract image data
                    data = obj.get_data()

                    # Determine mode
                    color_space = obj.get("/ColorSpace", "/DeviceRGB")
                    if isinstance(color_space, list):
                        color_space = str(color_space[0])
                    else:
                        color_space = str(color_space)

                    if "Gray" in color_space:
                        mode = "L"
                    elif "CMYK" in color_space:
                        mode = "CMYK"
                    else:
                        mode = "RGB"

                    # Try to create image from raw data
                    try:
                        img = Image.frombytes(mode, (width, height), data)
                    except Exception:
                        continue

                    # Compress as JPEG
                    buf = BytesIO()
                    if img.mode == "CMYK":
                        img = img.convert("RGB")
                    if img.mode == "RGBA":
                        img = img.convert("RGB")
                    img.save(buf, format="JPEG", quality=quality, optimize=True)
                    buf.seek(0)

                    # Replace image data
                    obj._data = buf.getvalue()
                    obj.update({
                        "/Filter": "/DCTDecode",
                        "/Length": len(buf.getvalue()),
                    })
                    image_count += 1
                except Exception:
                    continue

    # Backup original to Old folder
    if backup:
        old_folder = input_path.parent / "Old"
        old_folder.mkdir(exist_ok=True)
        backup_path = old_folder / input_path.name
        shutil.move(str(input_path), str(backup_path))

    # Write compressed PDF
    with open(str(input_path), "wb") as f:
        writer.write(f)

    new_size = input_path.stat().st_size / 1024  # KB
    reduction = ((original_size - new_size) / original_size * 100) if original_size > 0 else 0

    print(f"  ✓ {input_path.name}")
    print(f"    {original_size:.0f} KB → {new_size:.0f} KB ({reduction:.0f}% smaller, {image_count} images)")
    return True


def main():
    parser = argparse.ArgumentParser(description="Compress images inside PDF files")
    parser.add_argument("files", nargs="*", help="PDF file(s) to compress")
    parser.add_argument("-q", "--quality", type=int, default=60,
                        help="JPEG quality 1-95 (default: 60, lower = smaller)")
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

    print(f"Quality: {args.quality} (lower = smaller file)")
    print("-" * 50)

    for f in args.files:
        compress_pdf(f, quality=args.quality, backup=not args.no_backup)

    print("-" * 50)
    print("Done! Originals saved in 'Old' folder.")


if __name__ == "__main__":
    main()
