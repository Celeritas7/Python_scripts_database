"""
PDF Booklet Creator
===================
Rearranges PDF pages into booklet (saddle-stitch) order for printing.

When printed double-sided (flip on short edge) and folded in half,
the pages read in the correct order — like a real booklet.

Usage:
    python pdf_booklet.py                          # Auto-convert all PDFs in folder
    python pdf_booklet.py input.pdf                # Single file → input_booklet.pdf
    python pdf_booklet.py input.pdf -o booklet.pdf # Custom output name

Features:
    - Auto-pads to multiple of 4 pages (blank pages added at end)
    - Supports A4 and Letter source pages
    - Two source pages placed side-by-side on landscape sheet
    - Proper scaling to fit without distortion
    - Works with tools launcher (dual-mode: current folder or script folder)

Print settings:
    - Paper size: A4 or Letter (landscape)
    - Double-sided: Flip on SHORT edge
    - Scaling: Actual size / 100%

Requirements:
    pip install pypdf reportlab
"""

import argparse
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter, Transformation, PageObject
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
except ImportError:
    print("ERROR: reportlab not installed. Run: pip install reportlab")
    sys.exit(1)

from io import BytesIO


def get_booklet_page_order(num_pages):
    """
    Calculate the page arrangement for booklet printing.
    
    For a booklet, pages are paired so that when the sheets are
    printed double-sided, folded, and stacked, they read in order.
    
    Returns list of (left_page, right_page) tuples for each half-sheet.
    Page numbers are 0-indexed. None means blank page.
    """
    # Pad to multiple of 4
    padded = num_pages
    while padded % 4 != 0:
        padded += 1
    
    pairs = []
    lo = 0
    hi = padded - 1
    
    while lo < hi:
        # Front of sheet: (last, first) — right side is the earlier page
        pairs.append((hi, lo))       # Front: left=hi, right=lo
        lo += 1
        hi -= 1
        # Back of sheet: (first, last)
        pairs.append((lo, hi))       # Back:  left=lo, right=hi
        lo += 1
        hi -= 1
    
    # Replace padded page numbers (>= num_pages) with None (blank)
    result = []
    for left, right in pairs:
        l = left if left < num_pages else None
        r = right if right < num_pages else None
        result.append((l, r))
    
    return result, padded


def create_blank_page(width, height):
    """Create a blank PDF page using reportlab."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def create_booklet(input_path, output_path=None):
    """
    Convert a PDF into booklet layout.
    
    Args:
        input_path: Path to source PDF
        output_path: Path for output (default: input_booklet.pdf)
    
    Returns:
        Path to generated booklet PDF
    """
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"  Error: File not found: {input_path}")
        return None
    
    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}_booklet.pdf"
    else:
        output_path = Path(output_path).resolve()
    
    # Read source PDF
    reader = PdfReader(str(input_path))
    num_pages = len(reader.pages)
    
    if num_pages == 0:
        print("  Error: PDF has no pages.")
        return None
    
    print(f"  Source: {input_path.name} ({num_pages} pages)")
    
    # Get source page dimensions (from first page)
    src_page = reader.pages[0]
    src_width = float(src_page.mediabox.width)
    src_height = float(src_page.mediabox.height)
    
    print(f"  Source page size: {src_width:.0f} x {src_height:.0f} pts "
          f"({src_width/72:.1f}\" x {src_height/72:.1f}\")")
    
    # Output page: landscape, two source pages side by side
    # Scale source pages to fit half the landscape sheet
    out_height = src_height                    # same height as source
    out_width = src_width * 2                  # double width (two pages side by side)
    half_width = src_width                     # each half = one source page width
    
    print(f"  Output page size: {out_width:.0f} x {out_height:.0f} pts "
          f"({out_width/72:.1f}\" x {out_height/72:.1f}\") landscape")
    
    # Calculate booklet page order
    pairs, padded_count = get_booklet_page_order(num_pages)
    num_sheets = padded_count // 4
    
    print(f"  Padded to {padded_count} pages → {num_sheets} sheet(s)")
    if padded_count > num_pages:
        print(f"  ({padded_count - num_pages} blank page(s) added)")
    
    # Build booklet PDF
    writer = PdfWriter()
    
    for i, (left_idx, right_idx) in enumerate(pairs):
        # Create a blank landscape page
        new_page = create_blank_page(out_width, out_height)
        
        # Place LEFT page (at x=0)
        if left_idx is not None:
            src = reader.pages[left_idx]
            # Merge at left position (x=0)
            new_page.merge_transformed_page(
                src,
                Transformation().translate(tx=0, ty=0)
            )
        
        # Place RIGHT page (at x=half_width)
        if right_idx is not None:
            src = reader.pages[right_idx]
            # Merge at right position
            new_page.merge_transformed_page(
                src,
                Transformation().translate(tx=half_width, ty=0)
            )
        
        writer.add_page(new_page)
    
    # Write output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(str(output_path), "wb") as f:
        writer.write(f)
    
    size_kb = output_path.stat().st_size / 1024
    print(f"  Output: {output_path.name} ({size_kb:.0f} KB)")
    print(f"  Print: Double-sided, flip on SHORT edge, actual size")
    
    return str(output_path)


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to booklet layout for saddle-stitch printing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    %(prog)s                              Auto-convert all PDFs in current folder
    %(prog)s document.pdf                 Single file → document_booklet.pdf
    %(prog)s document.pdf -o booklet.pdf  Custom output name
    
Printing instructions:
    1. Print the booklet PDF
    2. Paper: same size as source (A4/Letter)
    3. Double-sided: Flip on SHORT edge
    4. Scaling: Actual size / 100%%
    5. Fold all sheets in half together and staple at the spine
        """
    )
    
    parser.add_argument("files", nargs="*", help="PDF file(s) to convert")
    parser.add_argument("-o", "--output", help="Output PDF path (single file only)")
    
    args = parser.parse_args()
    
    # If no files provided, auto-find all PDFs in working directory
    if not args.files:
        script_dir = Path(__file__).parent.resolve()
        work_dir = Path.cwd()
        if work_dir != script_dir:
            search_dir = work_dir
        else:
            search_dir = script_dir
        
        pdf_files = sorted(
            f for f in search_dir.glob("*.pdf")
            if "_booklet" not in f.stem  # skip already-converted files
        )
        
        if not pdf_files:
            print(f"No PDF files found in: {search_dir}")
            return
        
        print(f"Auto-converting {len(pdf_files)} PDF(s) in: {search_dir}")
        print("=" * 60)
        
        for pdf_file in pdf_files:
            print(f"\nProcessing: {pdf_file.name}")
            create_booklet(str(pdf_file))
        
        print("\n" + "=" * 60)
        print(f"Done! {len(pdf_files)} booklet(s) created.")
    
    elif len(args.files) == 1:
        print(f"Processing: {args.files[0]}")
        create_booklet(args.files[0], output_path=args.output)
    
    else:
        if args.output:
            print("Warning: -o/--output ignored in batch mode.")
        for pdf_file in args.files:
            print(f"\nProcessing: {pdf_file}")
            create_booklet(pdf_file)


if __name__ == "__main__":
    main()
