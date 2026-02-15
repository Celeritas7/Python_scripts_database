"""
Images to PDF Converter (Windows / Mac / Linux)
================================================
Automatically combines ALL images in the same folder as this script
into a single PDF, ordered alphabetically by filename.

Usage:
  Just double-click to run. No input needed.
  PDF is saved in the same folder as "combined_output.pdf".

Supported formats:
  .png, .jpg, .jpeg, .bmp, .tiff, .tif, .webp, .gif

Requirements:
  - pip install Pillow
"""

import re
import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("=" * 50)
    print("ERROR: Pillow not installed!")
    print("Run: pip install Pillow")
    print("=" * 50)
    input("Press Enter to exit...")
    sys.exit(1)


# Supported image extensions
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp", ".gif"}


def main():
    print("=" * 50)
    print("    Images to PDF Converter")
    print("=" * 50)
    print()

    # Get script's folder
    script_folder = Path(__file__).parent.resolve()
    print(f"Scanning: {script_folder}\n")

    # Natural sort: treats numbers inside filenames as actual numbers
    # e.g. Notes (1), Notes (2), ... Notes (10), Notes (11)
    def natural_sort_key(filepath):
        parts = re.split(r'(\d+)', filepath.name.lower())
        return [int(p) if p.isdigit() else p for p in parts]

    image_files = sorted(
        [f for f in script_folder.iterdir()
         if f.suffix.lower() in IMAGE_EXTENSIONS],
        key=natural_sort_key
    )

    if not image_files:
        print("[INFO] No image files found in this folder.")
        input("\nPress Enter to exit...")
        return

    print(f"Found {len(image_files)} image(s):\n")
    print("-" * 50)
    for img_file in image_files:
        print(f"  {img_file.name}")
    print("-" * 50)
    print()

    # Convert images to RGB (PDF requires RGB, not RGBA/P/etc.)
    pdf_pages = []
    failed = 0

    for img_file in image_files:
        try:
            img = Image.open(img_file)
            # Handle animated GIFs — take first frame only
            img = img.copy()
            # Convert to RGB for PDF compatibility
            if img.mode != "RGB":
                img = img.convert("RGB")
            pdf_pages.append(img)
            print(f"  Loaded: {img_file.name}")
        except Exception as e:
            print(f"  [ERROR] {img_file.name}: {e}")
            failed += 1

    if not pdf_pages:
        print("\n[ERROR] No images could be loaded.")
        input("\nPress Enter to exit...")
        return

    # Save as single PDF
    output_path = script_folder / "combined_output.pdf"
    print(f"\nCreating PDF with {len(pdf_pages)} page(s)...")

    try:
        pdf_pages[0].save(
            str(output_path),
            "PDF",
            save_all=True,
            append_images=pdf_pages[1:],
            resolution=150.0,
        )
        print(f"\n  -> {output_path.name} ✔")
    except Exception as e:
        print(f"\n  [ERROR] Failed to create PDF: {e}")
        failed += 1

    # Summary
    print()
    print("-" * 50)
    print(f"Done! Pages: {len(pdf_pages)} | Failed: {failed}")
    print("=" * 50)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
