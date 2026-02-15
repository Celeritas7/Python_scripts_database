"""
HTML to PDF converter using Playwright (headless Chrome).

Uses the same rendering engine as Ctrl+P, so ALL @media print CSS works:
  - columns / column-span / column-rule
  - page-break-inside / break-inside / break-after
  - @page margins
  - thead repeat across pages

Usage:
    python html_to_pdf.py input.html                     # → input.pdf (same folder)
    python html_to_pdf.py input.html -o output.pdf       # → specific output path
    python html_to_pdf.py *.html                         # → batch convert all
    python html_to_pdf.py input.html --landscape         # → landscape orientation
    python html_to_pdf.py input.html --format Letter     # → Letter instead of A4

Requirements:
    pip install playwright
    playwright install chromium
"""

import argparse
import sys
from pathlib import Path


def convert_html_to_pdf(
    html_path: str,
    output_path: str = None,
    format: str = "A4",
    landscape: bool = False,
    margin_top: str = "10mm",
    margin_bottom: str = "10mm",
    margin_left: str = "10mm",
    margin_right: str = "10mm",
    print_background: bool = True,
    scale: float = 1.0,
    header_template: str = "",
    footer_template: str = "",
    display_header_footer: bool = False,
):
    """
    Convert a local HTML file to PDF using headless Chrome via Playwright.

    Args:
        html_path:       Path to the HTML file
        output_path:     Output PDF path (default: same name with .pdf extension)
        format:          Page size - "A4", "Letter", "Legal", "A3", etc.
        landscape:       True for landscape orientation
        margin_*:        Page margins (CSS units: mm, cm, in, px)
        print_background: Include background colors/images
        scale:           Scale factor (0.1 to 2.0)
        header_template: HTML for page header (use classes: date, title, url, pageNumber, totalPages)
        footer_template: HTML for page footer
        display_header_footer: Show header/footer
    """
    from playwright.sync_api import sync_playwright

    html_path = Path(html_path).resolve()
    if not html_path.exists():
        print(f"Error: File not found: {html_path}")
        return None

    if output_path is None:
        output_path = html_path.with_suffix(".pdf")
    else:
        output_path = Path(output_path).resolve()

    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    file_url = html_path.as_uri()

    print(f"Converting: {html_path.name}")
    print(f"  Format: {format} {'(landscape)' if landscape else '(portrait)'}")
    print(f"  Margins: {margin_top} / {margin_right} / {margin_bottom} / {margin_left}")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Navigate and wait for full render
        page.goto(file_url, wait_until="networkidle")

        # Optional: wait for any lazy-loaded content
        page.wait_for_timeout(500)

        # Generate PDF - this triggers @media print CSS
        pdf_options = {
            "path": str(output_path),
            "format": format,
            "landscape": landscape,
            "print_background": print_background,
            "scale": scale,
            "margin": {
                "top": margin_top,
                "bottom": margin_bottom,
                "left": margin_left,
                "right": margin_right,
            },
            "display_header_footer": display_header_footer,
        }

        if display_header_footer:
            if header_template:
                pdf_options["header_template"] = header_template
            if footer_template:
                pdf_options["footer_template"] = footer_template

        page.pdf(**pdf_options)
        browser.close()

    size_kb = output_path.stat().st_size / 1024
    print(f"  Output: {output_path} ({size_kb:.0f} KB)")
    return str(output_path)


def batch_convert(
    html_files: list,
    output_dir: str = None,
    **kwargs,
):
    """Convert multiple HTML files to PDF."""
    results = []
    for html_file in html_files:
        html_path = Path(html_file)
        if output_dir:
            out = Path(output_dir) / html_path.with_suffix(".pdf").name
        else:
            out = None
        result = convert_html_to_pdf(html_file, output_path=out, **kwargs)
        if result:
            results.append(result)
        print()
    return results


def main():
    parser = argparse.ArgumentParser(
        description="Convert HTML to PDF using headless Chrome (Playwright)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s notes.html                        Single file → notes.pdf
  %(prog)s notes.html -o ~/Desktop/out.pdf   Custom output path
  %(prog)s *.html                            Batch convert all HTML files
  %(prog)s *.html --output-dir ./pdfs        Batch with output directory
  %(prog)s notes.html --landscape            Landscape orientation
  %(prog)s notes.html --format Letter        US Letter instead of A4
  %(prog)s notes.html --scale 0.9            Slightly shrink content
  %(prog)s notes.html --margin 7mm           Set all margins to 7mm
  %(prog)s notes.html --page-numbers         Add page numbers in footer
        """,
    )
    parser.add_argument("files", nargs="*", help="HTML file(s) to convert")
    parser.add_argument("-o", "--output", help="Output PDF path (single file only)")
    parser.add_argument("--output-dir", help="Output directory (batch mode)")
    parser.add_argument("--format", default="A4",
                        help="Page format: A4, Letter, Legal, A3 (default: A4)")
    parser.add_argument("--landscape", action="store_true",
                        help="Landscape orientation")
    parser.add_argument("--scale", type=float, default=1.0,
                        help="Scale factor 0.1-2.0 (default: 1.0)")
    parser.add_argument("--margin", default=None,
                        help="Set all margins (e.g., 7mm, 0.5in)")
    parser.add_argument("--margin-top", default="10mm")
    parser.add_argument("--margin-bottom", default="10mm")
    parser.add_argument("--margin-left", default="10mm")
    parser.add_argument("--margin-right", default="10mm")
    parser.add_argument("--no-background", action="store_true",
                        help="Exclude background colors/images")
    parser.add_argument("--page-numbers", action="store_true",
                        help="Add page numbers in footer")

    args = parser.parse_args()

    # If no files provided, auto-find all .html in working directory
    if not args.files:
        script_dir = Path(__file__).parent.resolve()
        work_dir = Path.cwd()
        if work_dir != script_dir:
            search_dir = work_dir
        else:
            search_dir = script_dir

        args.files = sorted(str(f) for f in search_dir.glob("*.html"))
        if not args.files:
            print(f"No .html files found in: {search_dir}")
            sys.exit(0)
        print(f"Auto-converting {len(args.files)} HTML file(s) in: {search_dir}")
        print("=" * 60)

    # Uniform margin shorthand
    if args.margin:
        args.margin_top = args.margin
        args.margin_bottom = args.margin
        args.margin_left = args.margin
        args.margin_right = args.margin

    # Page number footer template
    footer = ""
    display_hf = False
    if args.page_numbers:
        display_hf = True
        footer = (
            '<div style="font-size:8px; width:100%; text-align:center; color:#888;">'
            '<span class="pageNumber"></span> / <span class="totalPages"></span>'
            "</div>"
        )

    common_kwargs = dict(
        format=args.format,
        landscape=args.landscape,
        scale=args.scale,
        margin_top=args.margin_top,
        margin_bottom=args.margin_bottom,
        margin_left=args.margin_left,
        margin_right=args.margin_right,
        print_background=not args.no_background,
        display_header_footer=display_hf,
        footer_template=footer,
        # Empty header needed when display_header_footer is True
        header_template='<span></span>' if display_hf else "",
    )

    if len(args.files) == 1 and not args.output_dir:
        convert_html_to_pdf(args.files[0], output_path=args.output, **common_kwargs)
    else:
        if args.output:
            print("Warning: -o/--output ignored in batch mode. Use --output-dir instead.")
        batch_convert(args.files, output_dir=args.output_dir, **common_kwargs)


if __name__ == "__main__":
    main()
