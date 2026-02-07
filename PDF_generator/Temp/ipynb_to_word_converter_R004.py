"""
IPYNB to Formatted Word Converter (Windows)
============================================
Converts notebooks to Word with proper table rendering.

Pipeline: ipynb → HTML (via nbconvert) → DOCX (via pypandoc)

This approach preserves DataFrame tables because nbconvert renders
them as HTML tables, unlike direct pandoc conversion which dumps
the raw text/plain output.

Post-processing:
  - Centers all images (like VBA script)
  - Applies custom Title/Heading styles

Usage:
  Double-click to run. Processes all .ipynb in script's folder.

Output:
  Word_Outputs/ folder with formatted .docx files

Requirements:
  - pip install nbconvert pypandoc python-docx
"""

import sys
import tempfile
from pathlib import Path

# Check dependencies
missing = []
try:
    import nbconvert
    from nbconvert import HTMLExporter
except ImportError:
    missing.append("nbconvert")

try:
    import pypandoc
except ImportError:
    missing.append("pypandoc")

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    missing.append("python-docx")

try:
    import nbformat
except ImportError:
    missing.append("nbformat")

if missing:
    print("=" * 50)
    print("ERROR: Missing packages!")
    print(f"Run: pip install {' '.join(missing)}")
    print("=" * 50)
    input("Press Enter to exit...")
    sys.exit(1)


# =============================================================================
# HEADING STYLE CONFIGURATION (from your sample document)
# =============================================================================
HEADING_STYLES = {
    'Title': {
        'size': 26,
        'bold': False,
        'color': (0x17, 0x36, 0x5D),  # Dark navy #17365D
    },
    'Heading 1': {
        'size': 14,
        'bold': True,
        'color': (0x36, 0x5F, 0x91),  # Medium blue #365F91
    },
    'Heading 2': {
        'size': 13,
        'bold': True,
        'color': (0x4F, 0x81, 0xBD),  # Light blue #4F81BD
    },
    'Heading 3': {
        'size': 11,
        'bold': True,
        'color': (0x4F, 0x81, 0xBD),  # Light blue #4F81BD
    },
}


def apply_heading_styles(doc):
    """Apply custom formatting to heading styles."""
    for style_name, fmt in HEADING_STYLES.items():
        try:
            style = doc.styles[style_name]
            font = style.font
            font.size = Pt(fmt['size'])
            font.bold = fmt['bold']
            font.color.rgb = RGBColor(*fmt['color'])
        except KeyError:
            print(f"      [WARN] Style '{style_name}' not found")


def center_all_images(doc):
    """Center all images (replicates VBA LockAndCenterAllImages)."""
    centered = 0
    for para in doc.paragraphs:
        # Check if paragraph contains drawings/images
        if para._element.xpath('.//w:drawing') or para._element.xpath('.//w:pict'):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            centered += 1
    return centered


def format_docx(docx_path: Path) -> bool:
    """Apply formatting: heading styles + center images."""
    try:
        doc = Document(str(docx_path))
        apply_heading_styles(doc)
        img_count = center_all_images(doc)
        doc.save(str(docx_path))
        print(f"      Styled headings, centered {img_count} image(s)")
        return True
    except Exception as e:
        print(f"      [ERROR] Formatting failed: {e}")
        return False


def convert_notebook(ipynb_path: Path, output_folder: Path) -> bool:
    """
    Convert ipynb → HTML → DOCX with formatting.
    
    Going through HTML preserves DataFrame tables properly.
    """
    print(f"\n  Processing: {ipynb_path.name}")
    
    docx_path = output_folder / f"{ipynb_path.stem}.docx"
    
    try:
        # Step 1: Read notebook
        print("    [1/3] Reading notebook...")
        with open(ipynb_path, 'r', encoding='utf-8') as f:
            notebook = nbformat.read(f, as_version=4)
        
        # Step 2: Convert to HTML (this renders DataFrames as proper tables)
        print("    [2/3] Converting to HTML → Word...")
        html_exporter = HTMLExporter()
        html_exporter.template_name = 'basic'  # Clean output without extra CSS
        (html_content, resources) = html_exporter.from_notebook_node(notebook)
        
        # Write HTML to temp file, then convert to DOCX
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', 
                                          delete=False, encoding='utf-8') as tmp:
            tmp.write(html_content)
            tmp_html_path = tmp.name
        
        try:
            pypandoc.convert_file(tmp_html_path, 'docx', outputfile=str(docx_path))
        finally:
            # Clean up temp file
            Path(tmp_html_path).unlink(missing_ok=True)
        
        # Step 3: Apply formatting
        print("    [3/3] Applying formatting...")
        format_docx(docx_path)
        
        print(f"    [SUCCESS] → {docx_path.name}")
        return True
        
    except Exception as e:
        print(f"    [ERROR] Conversion failed: {e}")
        return False


def main():
    print("=" * 55)
    print("    IPYNB → Formatted Word Converter")
    print("    (Preserves tables & images)")
    print("=" * 55)
    
    # Get script's folder
    script_folder = Path(__file__).parent.resolve()
    print(f"\nScanning: {script_folder}")
    
    # Find notebooks (exclude checkpoints)
    notebooks = [f for f in script_folder.glob("*.ipynb") 
                 if ".ipynb_checkpoints" not in str(f)]
    
    if not notebooks:
        print("\n[INFO] No .ipynb files found in this folder.")
        input("\nPress Enter to exit...")
        return
    
    print(f"Found {len(notebooks)} notebook(s)")
    
    # Create output folder
    output_folder = script_folder / "Word_Outputs"
    output_folder.mkdir(exist_ok=True)
    print(f"Output: {output_folder}")
    print("-" * 55)
    
    # Process each notebook
    success = 0
    failed = 0
    
    for nb in notebooks:
        if convert_notebook(nb, output_folder):
            success += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "=" * 55)
    print(f"COMPLETE! Success: {success} | Failed: {failed}")
    print("=" * 55)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
