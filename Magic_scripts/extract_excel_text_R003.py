"""
Remove Images from ALL Excel Files (Keep Text Shapes & Formatting)
===================================================================
Usage:
    pip install openpyxl
    python extract_excel_text.py

This script surgically removes ONLY embedded images (photos, CAD views, etc.)
while preserving:
  - All cell data and formatting (borders, colors, merges)
  - All text shapes / text boxes in the drawing layer
  - Annotations (red circles, arrows, connectors)
  - Step numbers, labels, part numbers

Cleaned files are saved in a 'cleaned' subfolder.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import os
import glob
import io


def get_base_dir():
    """
    Smart directory detection:
    - If called from a different folder (e.g. via tools, IDE, or another script),
      use the current working directory.
    - If run directly from its own folder, use the script's directory.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    work_dir = os.getcwd()

    if work_dir != script_dir:
        base_dir = work_dir
    else:
        base_dir = script_dir

    return base_dir


def register_namespaces(xml_content):
    """Register all XML namespaces found in the content to preserve them on output."""
    ns_pattern = re.findall(r'xmlns:?(\w*)="([^"]+)"', xml_content)
    for prefix, uri in ns_pattern:
        if prefix:
            try:
                ET.register_namespace(prefix, uri)
            except ValueError:
                pass


def remove_pics_from_element(element, ns_xdr):
    """Recursively remove <xdr:pic> elements but keep shapes with text."""
    removed = 0
    pics_to_remove = element.findall(f'{ns_xdr}pic')
    for pic in pics_to_remove:
        element.remove(pic)
        removed += 1

    # Recurse into group shapes
    for grp in element.findall(f'{ns_xdr}grpSp'):
        removed += remove_pics_from_element(grp, ns_xdr)

    return removed


def process_drawing_xml(drawing_xml_bytes):
    """
    Parse drawing XML, remove all <xdr:pic> (image) elements,
    keep all <xdr:sp> (text shapes), <xdr:cxnSp> (connectors), etc.
    Returns modified XML bytes and count of removed images.
    """
    drawing_xml = drawing_xml_bytes.decode("utf-8")
    register_namespaces(drawing_xml)

    root = ET.fromstring(drawing_xml)

    ns_xdr = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
    ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

    total_removed = 0
    anchors_to_remove = []

    for anchor in list(root):
        if 'Anchor' not in anchor.tag:
            continue

        has_pic = anchor.find(f'.//{ns_xdr}pic') is not None
        has_shape = anchor.find(f'.//{ns_xdr}sp') is not None
        has_group = anchor.find(f'.//{ns_xdr}grpSp') is not None
        has_connector = anchor.find(f'.//{ns_xdr}cxnSp') is not None
        texts = [t.text for t in anchor.findall(f'.//{ns_a}t') if t.text]

        if has_pic and not has_shape and not has_group and not has_connector and not texts:
            # Pure image anchor — remove entirely
            anchors_to_remove.append(anchor)
            total_removed += len(anchor.findall(f'.//{ns_xdr}pic'))
        elif has_pic:
            # Mixed anchor — remove pics but keep shapes/text
            total_removed += remove_pics_from_element(anchor, ns_xdr)
            # Also recurse into groups
            for grp in anchor.findall(f'.//{ns_xdr}grpSp'):
                total_removed += remove_pics_from_element(grp, ns_xdr)

    for anchor in anchors_to_remove:
        root.remove(anchor)

    # Write modified XML
    output = io.BytesIO()
    tree = ET.ElementTree(root)
    tree.write(output, xml_declaration=True, encoding='UTF-8')

    return output.getvalue(), total_removed


def remove_images_from_excel(input_path, output_path):
    """Remove only images from an Excel file while preserving everything else."""
    print(f"\n  Loading: {os.path.basename(input_path)} ...")

    with zipfile.ZipFile(input_path, 'r') as zin:
        all_files = zin.namelist()

        # Find all drawing XML files
        drawing_files = [f for f in all_files if f.startswith('xl/drawings/') and f.endswith('.xml') and '_rels' not in f]
        image_files = [f for f in all_files if f.startswith('xl/media/')]

        if not image_files:
            print(f"    No images found — skipping")
            return False

        print(f"    Found {len(image_files)} image file(s), {len(drawing_files)} drawing(s)")

        # Process each drawing XML to remove pic elements
        modified_drawings = {}
        total_removed = 0

        for drawing_file in drawing_files:
            drawing_bytes = zin.read(drawing_file)
            modified_xml, removed = process_drawing_xml(drawing_bytes)
            modified_drawings[drawing_file] = modified_xml
            total_removed += removed
            if removed > 0:
                print(f"    {drawing_file}: removed {removed} image reference(s)")

        # Rebuild xlsx: skip image files, use modified drawing XMLs
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in all_files:
                if item in image_files:
                    continue  # Skip image files
                elif item in modified_drawings:
                    zout.writestr(item, modified_drawings[item])
                else:
                    zout.writestr(item, zin.read(item))

    input_size = os.path.getsize(input_path) / (1024 * 1024)
    output_size = os.path.getsize(output_path) / (1024 * 1024)
    print(f"    {input_size:.2f} MB --> {output_size:.2f} MB  ({total_removed} images removed)  ✓")
    return True


if __name__ == "__main__":
    base_dir = get_base_dir()
    print(f"Working directory: {base_dir}")

    # Find all Excel files in the folder
    excel_files = sorted(
        glob.glob(os.path.join(base_dir, "*.xlsx"))
    )

    # Exclude files already in 'cleaned' subfolder
    excel_files = [f for f in excel_files if os.sep + "cleaned" + os.sep not in f]

    if not excel_files:
        print("\nNo Excel files (.xlsx) found in this folder.")
        print("Please run this script from the folder containing your Excel files.")
        exit(1)

    print(f"Found {len(excel_files)} Excel file(s)")

    # Create 'cleaned' subfolder for output
    output_dir = os.path.join(base_dir, "cleaned")
    os.makedirs(output_dir, exist_ok=True)

    success = 0
    skipped = 0
    failed = 0

    for filepath in excel_files:
        filename = os.path.basename(filepath)
        output_path = os.path.join(output_dir, filename)

        try:
            result = remove_images_from_excel(filepath, output_path)
            if result:
                success += 1
            else:
                skipped += 1
        except Exception as e:
            print(f"\n  ✗ Failed: {filename} — {e}")
            failed += 1

    print(f"\n{'='*50}")
    print(f"Complete! {success} cleaned, {skipped} skipped (no images), {failed} failed.")
    if success > 0:
        print(f"Cleaned files saved to: {output_dir}")
