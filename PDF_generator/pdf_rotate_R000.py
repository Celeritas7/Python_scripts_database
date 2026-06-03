"""
PDF Page Rotator
================
Rotates all pages in a PDF by a specified angle and saves in-place.

Usage:
    python pdf_rotate.py                    # Auto-rotate all PDFs in folder (asks angle)
    python pdf_rotate.py input.pdf          # Rotate single file (asks angle)
    python pdf_rotate.py input.pdf -a 90    # Rotate 90° clockwise
    python pdf_rotate.py input.pdf -a -90   # Rotate 90° counter-clockwise

Requirements:
    pip install pypdf
"""

import argparse
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)


def rotate_pdf(input_path, angle=90):
    """Rotate all pages in a PDF and overwrite the original file."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    reader = PdfReader(str(input_path))
    writer = PdfWriter()

    for page in reader.pages:
        page.rotate(angle)
        writer.add_page(page)

    with open(str(input_path), "wb") as f:
        writer.write(f)

    print(f"  ✓ {input_path.name} — {len(reader.pages)} pages rotated {angle}°")
    return True


def main():
    parser = argparse.ArgumentParser(description="Rotate PDF pages")
    parser.add_argument("files", nargs="*", help="PDF file(s) to rotate")
    parser.add_argument("-a", "--angle", type=int, default=None,
                        help="Rotation angle: 90, 180, -90 (default: asks interactively)")
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

    # Ask angle if not provided
    angle = args.angle
    if angle is None:
        print("\nRotation options:")
        print("   90  = clockwise")
        print("  -90  = counter-clockwise")
        print("  180  = upside down")
        raw = input("\nEnter angle (default 90): ").strip()
        angle = int(raw) if raw else 90

    print(f"\nRotating {angle}°...")
    print("-" * 40)
    for f in args.files:
        rotate_pdf(f, angle)
    print("-" * 40)
    print("Done!")


if __name__ == "__main__":
    main()
