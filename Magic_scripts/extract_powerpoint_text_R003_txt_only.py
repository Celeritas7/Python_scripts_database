#!/usr/bin/env python3
"""PowerPoint Text Extractor R003 — TXT only, with recursive folders.

One standalone Python file. Python 3.9+; no pip packages or PowerPoint required.
Only reads .pptx files. Does not modify them or create .pptx/.json outputs.

By default, scans the current folder AND all its subfolders. For each file,
creates one UTF-8 .txt file in a sibling 'extracted_text' folder.

Examples (PowerShell / Command Prompt):
    python extract_powerpoint_text_R003_txt_only.py
    python extract_powerpoint_text_R003_txt_only.py --folder "D:\\Presentations"
    python extract_powerpoint_text_R003_txt_only.py --file "D:\\Presentations\\One.pptx"
    python extract_powerpoint_text_R003_txt_only.py --folder . --include-notes
    python extract_powerpoint_text_R003_txt_only.py --folder . --overwrite

Uses true presentation slide order; includes editable slide/table text and
supported cached chart/SmartArt text. Speaker notes are optional.
Images are not OCR'd; legacy .ppt files must be converted to .pptx first.
"""

import argparse
import os
from pathlib import Path
import posixpath
import sys
import tempfile
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

VERSION = "R003"
DEFAULT_XML_LIMIT_MB = 32
OUTPUT_FOLDER = "extracted_text"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def tag(node):
    return node.tag.rsplit("}", 1)[-1] if isinstance(node.tag, str) else ""


def ns(node):
    return node.tag[1:].split("}", 1)[0] if isinstance(node.tag, str) and node.tag.startswith("{") else ""


def resolve_target(owner, value):
    uri = urlsplit(value)
    if uri.scheme or uri.netloc or uri.query:
        raise ValueError(f"Unsupported internal relationship target: {value}")
    target = unquote(uri.path)
    if not target or "\\" in target or "\x00" in target:
        raise ValueError(f"Invalid internal relationship target: {value}")
    if target.startswith("/"):
        result = posixpath.normpath(target.lstrip("/"))
    else:
        result = posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
    if result in ("", ".", "..") or result.startswith("../"):
        raise ValueError(f"Unsafe internal relationship target: {value}")
    return result


def rels_part(owner):
    folder, name = posixpath.split(owner)
    return posixpath.join(folder, "_rels", name + ".rels") if owner else "_rels/.rels"


class Presentation:
    def __init__(self, archive, xml_limit):
        self.archive = archive
        self.xml_limit = xml_limit
        self.members = {}
        seen_names = set()
        for entry in archive.infolist():
            name = entry.filename
            if (name.startswith(("/", "\\")) or "\\" in name
                    or any(piece in (".", "..") for piece in name.split("/"))):
                raise ValueError(f"Unsafe package member: {name}")
            if name in self.members or name.casefold() in seen_names:
                raise ValueError(f"Duplicate package member: {name}")
            if entry.flag_bits & 1:
                raise ValueError("Encrypted PowerPoint files are unsupported")
            seen_names.add(name.casefold())
            self.members[name] = entry
        if "[Content_Types].xml" not in self.members or "_rels/.rels" not in self.members:
            raise ValueError("Not a valid .pptx Open XML package")

    def xml(self, name):
        member = self.members.get(name)
        if member is None:
            raise ValueError(f"Missing PowerPoint component: {name}")
        if member.file_size > self.xml_limit:
            raise ValueError(f"XML component too large ({member.file_size} bytes): {name}; use --max-xml-mb")
        with self.archive.open(member) as src:
            data = src.read(self.xml_limit + 1)
        if len(data) > self.xml_limit:
            raise ValueError(f"XML component exceeded safety size: {name}")
        # Entity declarations are unnecessary in Office Open XML; reject them.
        if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
            raise ValueError(f"Unsupported XML entity/DOCTYPE in {name}")
        try:
            return ET.fromstring(data)
        except ET.ParseError as exc:
            raise ValueError(f"Malformed XML in {name}: {exc}") from exc

    def relationships(self, owner):
        name = rels_part(owner)
        if name not in self.members:
            return {}
        root = self.xml(name)
        if ns(root) != PKG or tag(root) != "Relationships":
            raise ValueError(f"Invalid relationship part: {name}")
        result = {}
        for item in root:
            if ns(item) != PKG or tag(item) != "Relationship":
                raise ValueError(f"Unexpected relationship element in {name}")
            rid, kind, target = (item.get(k) for k in ("Id", "Type", "Target"))
            external = item.get("TargetMode") == "External"
            if not rid or not kind or not target or rid in result:
                raise ValueError(f"Missing or duplicate relationship in {name}")
            result[rid] = (kind, target if external else resolve_target(owner, target), external)
        return result

    def ordered_slides(self):
        roots = [v for v in self.relationships("").values()
                 if v[0].endswith("/officeDocument") and not v[2]]
        if len(roots) != 1:
            raise ValueError("Could not identify the PowerPoint main presentation part")
        presentation = roots[0][1]
        doc = self.xml(presentation)
        if tag(doc) != "presentation" or ns(doc) != P:
            raise ValueError("The document is not a standard PowerPoint presentation")
        links = self.relationships(presentation)
        id_list = doc.find(f"{{{P}}}sldIdLst")
        output = []
        for item in id_list if id_list is not None else []:
            rid = item.get(f"{{{R}}}id")
            link = links.get(rid)
            if not link or not link[0].endswith("/slide") or link[2] or link[1] in output:
                raise ValueError("Missing or duplicate slide relationship")
            output.append(link[1])
        return output


def paragraph(node):
    pieces = []
    for item in node.iter():
        if ns(item) == A:
            if tag(item) == "t":
                pieces.append(item.text or "")
            elif tag(item) == "br":
                pieces.append("\n")
            elif tag(item) == "tab":
                pieces.append("\t")
    return "".join(pieces)


def editable_lines(node):
    """Traverse in XML order, treating tables as tab-delimited blocks."""
    table_counter = [0]

    def walk(current):
        if ns(current) == A and tag(current) == "tbl":
            table_counter[0] += 1
            yield f"[Table {table_counter[0]} - columns separated by tabs; cell breaks shown as \\n]"
            for row in current:
                if ns(row) != A or tag(row) != "tr":
                    continue
                values = []
                for cell in row:
                    if ns(cell) == A and tag(cell) == "tc":
                        ps = [paragraph(p) for p in cell.iter() if ns(p) == A and tag(p) == "p"]
                        content = "\n".join(ps)
                        values.append(content.replace("\\", "\\\\").replace("\t", "\\t")
                                      .replace("\r", "\\r").replace("\n", "\\n"))
                yield "\t".join(values)
            yield "[End table]"
        elif ns(current) == A and tag(current) == "p":
            value = paragraph(current)
            if value.strip():
                yield value
        else:
            for child in current:
                yield from walk(child)

    yield from walk(node)


def related_text(pkg, slide_doc, slide_rels):
    observed = set()
    for item in slide_doc.iter():
        role = rid = expected = None
        if ns(item) == C and tag(item) == "chart":
            role, rid, expected = "Chart", item.get(f"{{{R}}}id"), "/chart"
        elif ns(item) == DGM and tag(item) == "relIds":
            role, rid, expected = "SmartArt", item.get(f"{{{R}}}dm"), "/diagramData"
        if not role or (role, rid) in observed:
            continue
        observed.add((role, rid))
        link = slide_rels.get(rid)
        if not link or not link[0].endswith(expected):
            raise ValueError(f"Slide contains an unresolved {role} reference: {rid}")
        if link[2]:
            yield f"[{role}: external content skipped]"
            continue
        root = pkg.xml(link[1])
        yield f"[{role}: stored editable text / cached labels]"
        seen = set()
        lines = list(editable_lines(root))
        if role == "Chart":
            for container in root.iter():
                if ns(container) == C and tag(container) in ("strCache", "strLit"):
                    for val in container.iter():
                        if ns(val) == C and tag(val) == "v" and (val.text or "").strip():
                            lines.append(val.text)
        for line in lines:
            if line not in seen:
                seen.add(line)
                yield line
        if not seen:
            yield "(No supported stored text found.)"


def slide_text(pkg, slide, index, include_notes):
    doc = pkg.xml(slide)
    if ns(doc) != P or tag(doc) != "sld":
        raise ValueError(f"Not a slide XML component: {slide}")
    rels = pkg.relationships(slide)
    yield f"========== Slide {index} =========="
    if doc.get("show") in ("0", "false"):
        yield "[Hidden slide - included in presentation order]"
    lines = list(editable_lines(doc))
    yield from lines or ["(No editable text in the slide's own shapes.)"]
    yield from related_text(pkg, doc, rels)
    if include_notes:
        for typ, dest, external in rels.values():
            if external or not typ.endswith("/notesSlide"):
                continue
            yield "[Speaker notes]"
            note_root = pkg.xml(dest)
            found = False
            for shape in note_root.iter():
                if ns(shape) != P or tag(shape) != "sp":
                    continue
                placeholders = [p.get("type") for p in shape.iter() if ns(p) == P and tag(p) == "ph"]
                if any(p in ("dt", "hdr", "ftr", "sldNum", "sldImg") for p in placeholders):
                    continue
                for line in editable_lines(shape):
                    yield line
                    found = True
            if not found:
                yield "(No supported speaker-note text found.)"
    yield ""


def is_link(path):
    return path.is_symlink()


def find_pptx(folder, recursive=True):
    files = []
    errors = []
    unsupported = 0

    def scan_err(exc):
        errors.append(str(exc))

    for base, folders, names in os.walk(folder, followlinks=False, onerror=scan_err):
        here = Path(base)
        folders[:] = sorted(f for f in folders
                            if f.casefold() not in {"cleaned", OUTPUT_FOLDER.casefold()}
                            and not f.startswith(".r002-stage-")
                            and not is_link(here / f))
        if not recursive:
            folders.clear()
        for name in sorted(names, key=str.casefold):
            path = here / name
            if name.startswith("~$") or is_link(path):
                continue
            if path.suffix.casefold() == ".pptx":
                files.append(path)
            elif path.suffix.casefold() in (".ppt", ".pptm"):
                unsupported += 1
    return files, errors, unsupported


def convert_one(source, overwrite, include_notes, max_xml_bytes):
    if not source.is_file() or source.suffix.casefold() != ".pptx" or is_link(source):
        raise ValueError("Source must be a regular .pptx file, not a shortcut or symlink")
    destination_dir = source.parent / OUTPUT_FOLDER
    destination = destination_dir / f"{source.stem}.txt"
    if destination.exists() and not overwrite:
        return "skipped", destination, 0
    destination_dir.mkdir(exist_ok=True)
    temp_path = None
    old_stat = source.stat()
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8-sig", newline="\n",
                                         prefix=".extract-", suffix=".tmp",
                                         dir=destination_dir, delete=False) as handle:
            temp_path = Path(handle.name)
            with zipfile.ZipFile(source, "r") as archive:
                pkg = Presentation(archive, max_xml_bytes)
                slides = pkg.ordered_slides()
                handle.write(f"Source: {source.name}\nExtractor: {VERSION} (TXT only)\n")
                handle.write("Slides follow true presentation order, including hidden slides.\n")
                handle.write("Editable text + supported chart labels / SmartArt text.\n")
                handle.write(f"Speaker notes: {'included' if include_notes else 'not requested'}\n")
                handle.write("Images are not OCR'd. Text within a slide follows XML order.\n\n")
                for index, slide in enumerate(slides, 1):
                    for line in slide_text(pkg, slide, index, include_notes):
                        handle.write(line + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        new_stat = source.stat()
        if (old_stat.st_size, old_stat.st_mtime_ns) != (new_stat.st_size, new_stat.st_mtime_ns):
            raise ValueError("Source changed during extraction; please rerun when closed")
        # No overwrite unless explicitly requested, including if the destination appeared mid-run.
        if destination.exists() and not overwrite:
            raise ValueError(f"Output already exists: {destination}. Use --overwrite")
        os.replace(temp_path, destination)
        temp_path = None
        return "saved", destination, len(slides)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--folder", type=Path, help="Folder to scan (default: current working directory)")
    choice.add_argument("--file", type=Path, help="Extract just one .pptx presentation")
    parser.add_argument("--include-notes", action="store_true", help="Include speaker notes")
    parser.add_argument("--no-recursive", action="store_true", help="Scan only the top-level folder")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing TXT files")
    parser.add_argument("--max-xml-mb", type=int, default=DEFAULT_XML_LIMIT_MB,
                        help="Maximum size of any XML component in MiB; default 32")
    args = parser.parse_args(argv)
    if args.max_xml_mb < 1:
        parser.error("--max-xml-mb must be positive")
    if args.file:
        sources = [args.file.expanduser().absolute()]
        errors = []
        unsupported = 0
    else:
        folder = (args.folder or Path.cwd()).expanduser().resolve()
        if not folder.is_dir():
            parser.error(f"Folder not found: {folder}")
        if folder.name.casefold() in {"cleaned", OUTPUT_FOLDER.casefold()}:
            parser.error("Select the source folder, not an output folder")
        sources, errors, unsupported = find_pptx(folder, recursive=not args.no_recursive)
    if unsupported:
        print(f"NOTE: Ignored {unsupported} .ppt/.pptm files; use Save As .pptx first")
    for error in errors:
        print(f"FOLDER SCAN ERROR: {error}", file=sys.stderr)
    if not sources:
        print("No .pptx files found. No changes made.")
        return 1 if errors else 2
    print(f"Found {len(sources)} PowerPoint file(s). Only TXT will be generated.")
    counts = {"saved": 0, "skipped": 0, "failed": len(errors)}
    for index, source in enumerate(sources, 1):
        print(f"[{index}/{len(sources)}] {source}", flush=True)
        try:
            status, target, slides = convert_one(
                source, args.overwrite, args.include_notes,
                args.max_xml_mb * 1024 * 1024)
            counts[status] += 1
            if status == "saved":
                print(f"  OK: {slides} slide(s) -> {target}")
            else:
                print(f"  SKIPPED: {target} already exists; add --overwrite to update")
        except (OSError, ValueError, zipfile.BadZipFile, RuntimeError, ET.ParseError) as exc:
            print(f"  FAILED: {exc}", file=sys.stderr)
            counts["failed"] += 1
    print(f"DONE: {counts['saved']} TXT saved, {counts['skipped']} skipped, {counts['failed']} failed.")
    return 1 if counts["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
