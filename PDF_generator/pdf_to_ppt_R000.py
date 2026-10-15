"""
PDF to PowerPoint Converter
===========================
Converts each page of a PDF into a PowerPoint slide. Each page is
rendered as a full-slide image, so the layout looks exactly like the PDF.
The text of each page is also copied into the slide's speaker notes
(so it stays searchable / copyable).

Slide size is matched to the PDF page shape (e.g. 16:9 slides stay 16:9,
A4 pages become A4-shaped slides).

Usage:
    python pdf_to_ppt.py                         # All PDFs in folder
    python pdf_to_ppt.py input.pdf               # Single file
    python pdf_to_ppt.py input.pdf -d 200        # Custom DPI (default: 150)
    python pdf_to_ppt.py input.pdf -f png        # PNG instead of JPG (sharper, bigger)
    python pdf_to_ppt.py input.pdf -q 80         # JPEG quality (default: 85)
    python pdf_to_ppt.py input.pdf --no-notes    # Don't copy page text into notes

Output:
    input.pdf → input.pptx  (same folder)

Requirements:
    pip install pymupdf python-pptx
"""

import argparse
import io
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF (newer versions)
except ImportError:
    try:
        import fitz  # PyMuPDF (older versions)
    except ImportError:
        print("ERROR: PyMuPDF not installed. Run: pip install pymupdf")
        sys.exit(1)

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    print("ERROR: python-pptx not installed. Run: pip install python-pptx")
    sys.exit(1)

EMU_PER_POINT = 12700  # 1 pt = 12700 EMU
MAX_SLIDE_EMU = 51206400  # PowerPoint max slide side (56 in)
MIN_SLIDE_EMU = 914400  # PowerPoint min slide side (1 in)


def slide_size_for(page_rect):
    """Return (width, height) in EMU matching the page aspect ratio."""
    w = page_rect.width * EMU_PER_POINT
    h = page_rect.height * EMU_PER_POINT

    # Scale into PowerPoint's allowed range while keeping the aspect ratio
    scale = min(1.0, MAX_SLIDE_EMU / max(w, h))
    w, h = w * scale, h * scale
    scale = max(1.0, MIN_SLIDE_EMU / min(w, h))
    return int(w * scale), int(h * scale)


def convert_pdf(input_path, dpi=150, fmt="jpg", quality=85, notes=True):
    """Convert a PDF into a .pptx with one slide per page."""
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return False

    output_path = input_path.with_suffix(".pptx")

    doc = fitz.open(str(input_path))
    num_pages = len(doc)
    if num_pages == 0:
        print(f"  Skipped (no pages): {input_path.name}")
        doc.close()
        return False

    prs = Presentation()
    slide_w, slide_h = slide_size_for(doc[0].rect)
    prs.slide_width = Emu(slide_w)
    prs.slide_height = Emu(slide_h)
    blank_layout = prs.slide_layouts[6]

    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)

    print(f"  {input_path.name} → {output_path.name}  ({num_pages} pages, {dpi} DPI)")

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        if fmt == "jpg":
            img_bytes = pix.tobytes("jpg", jpg_quality=quality)
        else:
            img_bytes = pix.tobytes("png")

        slide = prs.slides.add_slide(blank_layout)

        # Fit the page inside the slide (centered) in case page sizes differ
        scale = min(slide_w / pix.width, slide_h / pix.height)
        pic_w, pic_h = int(pix.width * scale), int(pix.height * scale)
        left, top = (slide_w - pic_w) // 2, (slide_h - pic_h) // 2
        slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, pic_w, pic_h)

        if notes:
            text = page.get_text().strip()
            if text:
                slide.notes_slide.notes_text_frame.text = text

        print(f"    Page {i+1}/{num_pages}", end="\r")

    doc.close()
    prs.save(str(output_path))

    size_kb = output_path.stat().st_size / 1024
    print(f"    ✓ {num_pages} slides saved ({size_kb:.0f} KB)" + " " * 20)
    return True


def main():
    parser = argparse.ArgumentParser(description="Convert PDF pages to PowerPoint slides")
    parser.add_argument("files", nargs="*", help="PDF file(s) to convert")
    parser.add_argument("-d", "--dpi", type=int, default=150,
                        help="Resolution in DPI (default: 150)")
    parser.add_argument("-f", "--format", choices=["jpg", "png"], default="jpg",
                        help="Image format inside slides (default: jpg)")
    parser.add_argument("-q", "--quality", type=int, default=85,
                        help="JPEG quality 1-95 (default: 85)")
    parser.add_argument("--no-notes", action="store_true",
                        help="Don't copy page text into speaker notes")
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
        convert_pdf(f, dpi=args.dpi, fmt=args.format, quality=args.quality,
                    notes=not args.no_notes)

    print("-" * 50)
    print("Done!")


if __name__ == "__main__":
    main()
