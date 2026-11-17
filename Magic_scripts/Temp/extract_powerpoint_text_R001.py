#!/usr/bin/env python3
"""PowerPoint text extractor and picture remover (R001).

Creates TWO independent outputs for each .pptx:
  1. cleaned/<name>.pptx -- removes editable slide picture shapes but retains
     slide text, shapes, charts, tables, and other slide elements.
  2. cleaned/<name>.txt -- extracts text from slides in presentation order.

Original presentations are NEVER changed. Nested folders are processed by default,
and each folder gets its own cleaned/ directory. The script works directly on
PowerPoint's ZIP/XML format, without loading whole decks into python-pptx.

Important limitations: This does not remove images used as slide backgrounds,
master/layout art, or inside chart/diagram data. Text converted to a bitmap
cannot be recovered. Some complex PowerPoint drawing markup may be altered by
XML serialization, so inspect a cleaned copy before deleting any original.

Usage:
    python extract_powerpoint_text_R001.py
    python extract_powerpoint_text_R001.py --folder "D:\\Presentations"
    python extract_powerpoint_text_R001.py --text-only
    python extract_powerpoint_text_R001.py --overwrite
No external Python packages required.
"""

import argparse
import os
from pathlib import Path
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Register widely used OOXML namespaces; may still normalize some serialization.
for prefix, uri in {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
}.items():
    ET.register_namespace(prefix, uri)

SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")


def slide_paths(archive):
    return sorted((n for n in archive.namelist() if SLIDE_RE.fullmatch(n)),
                  key=lambda n: int(SLIDE_RE.fullmatch(n).group(1)))


def remove_pictures(root):
    """Remove picture shapes everywhere in the slide shape tree, including groups."""
    count = 0
    for parent in root.iter():
        for child in list(parent):
            if child.tag == P + "pic":
                parent.remove(child)
                count += 1
    return count


def slide_text(root):
    """All editable DrawingML text, including grouped shapes and tables."""
    lines = []
    for paragraph in root.iter(A + "p"):
        chunks = [node.text or "" for node in paragraph.iter(A + "t")]
        if chunks and "".join(chunks).strip():
            lines.append("".join(chunks))
    return lines


def xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def pruned_image_relationships(archive, slide_name, new_slide_root):
    """Remove image relationships unreferenced by the edited slide XML."""
    dirname, filename = posixpath.split(slide_name)
    relname = posixpath.join(dirname, "_rels", filename + ".rels")
    if relname not in archive.namelist():
        return None
    relroot = ET.fromstring(archive.read(relname))
    referenced_ids = {value for elem in new_slide_root.iter()
                      for key, value in elem.attrib.items()
                      if key.startswith(R)}
    changed = False
    for rel in list(relroot):
        if rel.attrib.get("Type", "").endswith("/image") and rel.attrib.get("Id") not in referenced_ids:
            relroot.remove(rel)
            changed = True
    return (relname, xml_bytes(relroot)) if changed else None


def media_still_used(names, data_for):
    """Find all internal targets of remaining OOXML relationship files."""
    referenced = set()
    for relfile in names:
        if not relfile.endswith(".rels"):
            continue
        try:
            relroot = ET.fromstring(data_for(relfile))
        except ET.ParseError:
            continue  # Be conservative: don't delete media if rel parsing failed.
        # Example: ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
        folder, basename = posixpath.split(relfile)
        source_folder = posixpath.dirname(posixpath.dirname(relfile))
        for rel in relroot:
            if rel.attrib.get("TargetMode") == "External":
                continue
            target = rel.attrib.get("Target", "").split("#", 1)[0]
            if not target:
                continue
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join(source_folder, target))
            referenced.add(resolved)
    return referenced


def process_pptx(source, destination_pptx, destination_txt, text_only=False):
    with zipfile.ZipFile(source, "r") as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise ValueError("Corrupt ZIP member: " + bad_file)
        slides = slide_paths(archive)
        if not slides:
            raise ValueError("No PowerPoint slide XML found")

        replacements = {}
        report_lines = ["Source: " + source.name, ""]
        total_images = 0
        for slide_number, slide_name in enumerate(slides, 1):
            root = ET.fromstring(archive.read(slide_name))
            paragraphs = slide_text(root)
            report_lines.append("========== Slide {} ==========".format(slide_number))
            report_lines.extend(paragraphs or ["(No editable slide text found)"])
            report_lines.append("")

            if not text_only:
                removed = remove_pictures(root)
                total_images += removed
                if removed:
                    replacements[slide_name] = xml_bytes(root)
                    rel_change = pruned_image_relationships(archive, slide_name, root)
                    if rel_change:
                        replacements[rel_change[0]] = rel_change[1]

        # Write the text report atomically (avoid half-written output files).
        destination_txt.parent.mkdir(parents=True, exist_ok=True)
        def atomic_write_bytes(dest, data):
            fd, tempname = tempfile.mkstemp(prefix=".tmp_", suffix=dest.suffix, dir=dest.parent)
            try:
                with os.fdopen(fd, "wb") as f:
                    f.write(data)
                os.replace(tempname, dest)
            finally:
                if os.path.exists(tempname):
                    os.unlink(tempname)

        if not text_only:
            all_names = set(archive.namelist())
            def read_current(n):
                return replacements.get(n, archive.read(n))
            # Only drop media if there are no relationships still referencing it.
            # This avoids stripping charts/master media that other parts share.
            used = media_still_used(all_names, read_current)
            unused_media = {n for n in all_names if n.startswith("ppt/media/") and n not in used}
            fd, tempname = tempfile.mkstemp(prefix=".tmp_", suffix=".pptx", dir=destination_pptx.parent)
            os.close(fd)
            try:
                with zipfile.ZipFile(tempname, "w") as output:
                    for info in archive.infolist():
                        if info.filename in unused_media:
                            continue
                        output.writestr(info, read_current(info.filename))
                with zipfile.ZipFile(tempname) as verify:
                    if verify.testzip():
                        raise ValueError("Cleaned PPTX failed ZIP verification")
                os.replace(tempname, destination_pptx)
            finally:
                if os.path.exists(tempname):
                    os.unlink(tempname)
        atomic_write_bytes(destination_txt, "\n".join(report_lines).encode("utf-8-sig"))
        return len(slides), total_images


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--folder", type=Path, default=Path.cwd(), help="Folder to scan (default: current working directory)")
    parser.add_argument("--text-only", action="store_true", help="Only export TXT files, do not create cleaned PPTX files")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing output files")
    parser.add_argument("--no-recursive", action="store_true", help="Only scan the selected folder")
    args = parser.parse_args()
    base = args.folder.expanduser().resolve()
    if not base.is_dir():
        parser.error("Not a folder: " + str(base))

    candidates = []
    walker = base.rglob("*.pptx") if not args.no_recursive else base.glob("*.pptx")
    for path in walker:
        if any(p.lower() == "cleaned" for p in path.relative_to(base).parts[:-1]):
            continue
        if path.name.startswith("~$"):
            continue
        candidates.append(path)
    print("Found {} PowerPoint file(s) in {}".format(len(candidates), base))
    success = skipped = failed = 0
    for path in sorted(candidates):
        output = path.parent / "cleaned"
        output.mkdir(exist_ok=True)
        pptx_path = output / path.name
        txt_path = output / (path.stem + ".txt")
        required = [txt_path] if args.text_only else [txt_path, pptx_path]
        if not args.overwrite and any(p.exists() for p in required):
            print("SKIP (output exists):", path)
            skipped += 1
            continue
        try:
            slide_count, image_count = process_pptx(path, pptx_path, txt_path, args.text_only)
            print("OK: {} | {} slides | {} picture shapes removed".format(path, slide_count, image_count))
            success += 1
        except Exception as exc:
            print("FAILED: {} | {}".format(path, exc))
            failed += 1
    print("Finished: {} succeeded, {} skipped, {} failed".format(success, skipped, failed))


if __name__ == "__main__":
    main()
