#!/usr/bin/env python3
"""PowerPoint Text Extractor / Safe Picture Remover R002.

Python 3.9+; standard library only. No pip install or Microsoft Office needed.

For every .pptx in the selected folder (including subfolders), create:
  cleaned/<name>.pptx       conservative picture-cleaned copy
  cleaned/<name>.txt        editable text, in presentation slide order
  cleaned/<name>.r002.json  source/output fingerprints and processing report

Originals are never modified. --text-only creates TXT and the report, not PPTX.
Existing R002 outputs are checked before being reused. Missing managed outputs
are recreated. Use --overwrite to replace older/unverified cleaned outputs or
refresh outputs after editing a source. It NEVER overwrites source presentations.

Examples:
  python extract_powerpoint_text_R002.py --folder "D:\\Presentations"
  python extract_powerpoint_text_R002.py --file "D:\\Presentations\\Assembly.pptx"
  python extract_powerpoint_text_R002.py --text-only --include-notes
  python extract_powerpoint_text_R002.py --overwrite

Supported: .pptx only, not legacy .ppt, .pptm, encrypted or signed cleanup.
Text in images is not recognized (no OCR). Standard cached chart labels and
stored SmartArt text are included; notes are opt-in. Visual reading order within
a slide is not reconstructed. Background/master/layout pictures, interactive
pictures, multimedia, and pictures with animation/connector references are kept.
The resulting deck still needs inspection in desktop PowerPoint.
"""

import argparse
from contextlib import contextmanager
import copy
from dataclasses import dataclass, field
import hashlib
import json
import os
from pathlib import Path
import posixpath
import shutil
import stat
import sys
import tempfile
from typing import Dict, Iterator, List, Optional, Set, Tuple
from urllib.parse import unquote, urlsplit
import xml.dom.minidom as DOM
from xml.dom import Node
from xml.parsers import expat
import zipfile

VERSION = "R002"
CHUNK = 1024 * 1024
DEFAULT_XML_LIMIT = 32 * 1024 * 1024
P_NS = {"http://schemas.openxmlformats.org/presentationml/2006/main",
        "http://purl.oclc.org/ooxml/presentationml/main"}
A_NS = {"http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main"}
R_NS = {"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships"}
C_NS = {"http://schemas.openxmlformats.org/drawingml/2006/chart",
        "http://purl.oclc.org/ooxml/drawingml/chart"}
DGM_NS = {"http://schemas.openxmlformats.org/drawingml/2006/diagram",
          "http://purl.oclc.org/ooxml/drawingml/diagram"}
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XMLNS_NS = "http://www.w3.org/2000/xmlns/"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


class ProcessingError(Exception):
    """A controlled, user-readable processing failure."""


class RecoveryError(ProcessingError):
    """Rollback failed: retain the staging directory for manual recovery."""


@dataclass
class Options:
    text_only: bool = False
    overwrite: bool = False
    include_notes: bool = False
    skip_crc_check: bool = False
    max_xml_bytes: int = DEFAULT_XML_LIMIT


@dataclass
class Stats:
    slides: int = 0
    removed_pictures: int = 0
    protected_pictures: int = 0
    removed_media: int = 0
    removed_media_bytes: int = 0
    protection_reasons: Dict[str, int] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)

    def warn(self, text: str) -> None:
        if text not in self.warnings:
            self.warnings.append(text)


def elements(node) -> Iterator:
    """Element descendants, including node itself, in document order."""
    stack = [node]
    while stack:
        current = stack.pop()
        if current.nodeType == Node.ELEMENT_NODE:
            yield current
        stack.extend(reversed([c for c in current.childNodes
                               if c.nodeType == Node.ELEMENT_NODE]))


def children(node) -> List:
    return [c for c in node.childNodes if c.nodeType == Node.ELEMENT_NODE]


def matches(node, namespaces, name: str) -> bool:
    return (node.nodeType == Node.ELEMENT_NODE
            and node.namespaceURI in namespaces and node.localName == name)


def attrs(node) -> Iterator:
    if node.nodeType == Node.ELEMENT_NODE:
        for index in range(node.attributes.length):
            yield node.attributes.item(index)


def rel_attr(node, name: str = "id") -> str:
    for attr in attrs(node):
        if attr.namespaceURI in R_NS and attr.localName == name:
            return attr.value
    return ""


def text_content(node) -> str:
    return "".join(c.data if c.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE)
                   else text_content(c) for c in node.childNodes)


def namespace_is_bound(node, prefix: str) -> bool:
    if prefix == "xml":
        return True
    while node is not None and node.nodeType == Node.ELEMENT_NODE:
        value = node.getAttribute("xmlns:" + prefix)
        if value:
            return True
        node = node.parentNode
    return False


def validate_compatibility_namespaces(doc, label: str) -> None:
    """Check prefix-valued MC attributes; a normal XML parse misses these."""
    for node in elements(doc.documentElement):
        for attr in attrs(node):
            bare_prefixes = (attr.namespaceURI == MC_NS and
                             attr.localName in {"Ignorable", "MustUnderstand"})
            bare_prefixes = bare_prefixes or (
                node.namespaceURI == MC_NS and node.localName == "Choice"
                and attr.name == "Requires")
            qnames = (attr.namespaceURI == MC_NS and attr.localName in {
                "ProcessContent", "PreserveElements", "PreserveAttributes"})
            if not (bare_prefixes or qnames):
                continue
            for token in attr.value.split():
                prefix = token if bare_prefixes else token.split(":", 1)[0]
                if not bare_prefixes and ":" not in token:
                    continue
                if prefix != "*" and not namespace_is_bound(node, prefix):
                    raise ProcessingError(
                        f"{label}: undeclared compatibility prefix {prefix!r}")


def parse_xml(data: bytes, label: str):
    """Reject DTD/entity input; minidom retains explicit namespace declarations.

    Namespace declarations are kept as attributes, including ones used only in
    MC attribute VALUES. Do not replace this with plain ElementTree serialization.
    """
    guard = expat.ParserCreate()
    depth = 0

    def reject_doctype(*_args):
        raise ProcessingError(f"{label}: DTD/entity declarations are not supported")

    def start(_name, _attributes):
        nonlocal depth
        depth += 1
        if depth > 256:
            raise ProcessingError(f"{label}: XML nesting exceeds the safety limit")

    def end(_name):
        nonlocal depth
        depth -= 1

    guard.StartDoctypeDeclHandler = reject_doctype
    guard.StartElementHandler = start
    guard.EndElementHandler = end
    try:
        guard.Parse(data, True)
        doc = DOM.parseString(data)
    except (expat.ExpatError, ValueError) as exc:
        raise ProcessingError(f"{label}: invalid XML ({exc})") from exc
    try:
        validate_compatibility_namespaces(doc, label)
    except BaseException:
        doc.unlink()
        raise
    return doc


@contextmanager
def xml_document(data: bytes, label: str):
    doc = parse_xml(data, label)
    try:
        yield doc
    finally:
        doc.unlink()


def serialize_xml(doc, label: str) -> bytes:
    data = doc.toxml(encoding="UTF-8")
    # Verify the edited serialization, not only the original DOM.
    with xml_document(data, label):
        pass
    return data


def relationship_part(owner: str) -> str:
    if not owner:
        return "_rels/.rels"
    folder, name = posixpath.split(owner)
    return posixpath.join(folder, "_rels", name + ".rels")


def relationship_owner(part: str) -> str:
    if part == "_rels/.rels":
        return ""
    folder, name = posixpath.split(part)
    if not name.endswith(".rels") or posixpath.basename(folder) != "_rels":
        raise ProcessingError(f"Invalid relationship-part path: {part}")
    return posixpath.join(posixpath.dirname(folder), name[:-5])


def resolve_target(owner: str, target: str) -> str:
    uri = urlsplit(target)
    if uri.scheme or uri.netloc or uri.query:
        raise ProcessingError(f"Unsupported internal relationship target: {target}")
    path = unquote(uri.path, encoding="utf-8", errors="strict")
    if "\\" in path or "\x00" in path:
        raise ProcessingError(f"Unsafe relationship target: {target}")
    if not path:
        resolved = owner  # A fragment-only link to the owning part.
    elif path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner), path))
    if not resolved or resolved in {".", ".."} or resolved.startswith("../"):
        raise ProcessingError(f"Relationship target escapes the package: {target}")
    return resolved


@dataclass(frozen=True)
class Relationship:
    ident: str
    kind: str
    target: str
    external: bool


def read_relationship_nodes(doc, owner: str, label: str) -> Dict[str, Relationship]:
    root = doc.documentElement
    if root.namespaceURI != PKG_NS or root.localName != "Relationships":
        raise ProcessingError(f"{label}: not an OPC relationship document")
    result = {}
    for node in children(root):
        if node.namespaceURI != PKG_NS or node.localName != "Relationship":
            raise ProcessingError(f"{label}: unsupported relationship element")
        ident, kind, target = (node.getAttribute(k) for k in ("Id", "Type", "Target"))
        mode = node.getAttribute("TargetMode")
        if not ident or not kind or not target or ident in result:
            raise ProcessingError(f"{label}: missing or duplicate relationship fields")
        if mode not in {"", "Internal", "External"}:
            raise ProcessingError(f"{label}: unknown TargetMode {mode!r}")
        external = mode == "External"
        result[ident] = Relationship(ident, kind, target if external else
                                    resolve_target(owner, target), external)
    return result


class Package:
    """One source ZIP. Edited XML is spooled to disk, not held for the whole deck."""

    def __init__(self, archive: zipfile.ZipFile, stage: Path, xml_limit: int):
        self.archive = archive
        self.stage = stage
        self.xml_limit = xml_limit
        self.replacements: Dict[str, Path] = {}
        self.info = {}
        normalized = set()
        for item in archive.infolist():
            name = item.filename
            if (name.startswith(("/", "\\")) or "\\" in name
                    or any(p in {".", ".."} for p in name.split("/"))):
                raise ProcessingError(f"Unsafe ZIP member name: {name}")
            if name in self.info or name.casefold() in normalized:
                raise ProcessingError(f"Duplicate ZIP member: {name}")
            if item.flag_bits & 1:
                raise ProcessingError("Encrypted presentations are not supported")
            self.info[name] = item
            normalized.add(name.casefold())
        if "[Content_Types].xml" not in self.info or "_rels/.rels" not in self.info:
            raise ProcessingError("Not an Open XML presentation package")

    def read_xml(self, name: str) -> bytes:
        if name not in self.info:
            raise ProcessingError(f"Missing XML part: {name}")
        replacement = self.replacements.get(name)
        size = replacement.stat().st_size if replacement else self.info[name].file_size
        if size > self.xml_limit:
            raise ProcessingError(
                f"XML part {name} is {size / 1024**2:.1f} MiB; limit is "
                f"{self.xml_limit / 1024**2:.0f} MiB. Raise --max-xml-mb only with enough RAM.")
        if replacement:
            return replacement.read_bytes()
        with self.archive.open(name) as stream:
            data = stream.read(self.xml_limit + 1)
        if len(data) > self.xml_limit:
            raise ProcessingError(f"XML part exceeds the size limit: {name}")
        return data

    def save_xml(self, name: str, doc) -> None:
        data = serialize_xml(doc, name)
        if len(data) > self.xml_limit:
            raise ProcessingError(f"Edited XML part exceeds the size limit: {name}")
        destination = self.stage / ("part_%06d.xml" % len(self.replacements))
        destination.write_bytes(data)
        self.replacements[name] = destination

    def relationships(self, owner: str) -> Dict[str, Relationship]:
        name = relationship_part(owner)
        if name not in self.info:
            return {}
        with xml_document(self.read_xml(name), name) as doc:
            return read_relationship_nodes(doc, owner, name)

    def validate_relationships(self, deleted: Optional[Set[str]] = None) -> Set[str]:
        deleted = deleted or set()
        used = set()
        for name in self.info:
            if not name.endswith(".rels") or name in deleted:
                continue
            owner = relationship_owner(name)
            if owner and (owner not in self.info or owner in deleted):
                raise ProcessingError(f"Relationship owner is missing: {owner}")
            for rel in self.relationships(owner).values():
                if not rel.external:
                    if rel.target not in self.info or rel.target in deleted:
                        raise ProcessingError(f"{name}: target is missing: {rel.target}")
                    used.add(rel.target)
        return used

    def slide_order(self) -> List[str]:
        offices = [r for r in self.relationships("").values()
                   if r.kind.endswith("/officeDocument") and not r.external]
        if len(offices) != 1:
            raise ProcessingError("Expected one internal officeDocument relationship")
        presentation = offices[0].target
        rels = self.relationships(presentation)
        with xml_document(self.read_xml(presentation), presentation) as doc:
            root = doc.documentElement
            if not matches(root, P_NS, "presentation"):
                raise ProcessingError("This package is not a supported PowerPoint presentation")
            slide_lists = [n for n in children(root) if matches(n, P_NS, "sldIdLst")]
            if len(slide_lists) > 1:
                raise ProcessingError("Multiple presentation slide lists")
            order = []
            seen_ids = set()
            for node in children(slide_lists[0]) if slide_lists else []:
                if not matches(node, P_NS, "sldId"):
                    raise ProcessingError("Unsupported item in presentation slide list")
                ident = node.getAttribute("id")
                rel = rels.get(rel_attr(node))
                if (not ident or ident in seen_ids or not rel or rel.external
                        or not rel.kind.endswith("/slide") or rel.target in order):
                    raise ProcessingError("Invalid/duplicate presentation slide relationship")
                seen_ids.add(ident)
                order.append(rel.target)
            return order

    def content_types(self) -> Tuple[Dict[str, str], Dict[str, str]]:
        name = "[Content_Types].xml"
        with xml_document(self.read_xml(name), name) as doc:
            root = doc.documentElement
            if root.namespaceURI != CT_NS or root.localName != "Types":
                raise ProcessingError("Invalid content-type manifest")
            defaults, overrides = {}, {}
            for node in children(root):
                if node.namespaceURI != CT_NS:
                    raise ProcessingError("Unsupported content-type element")
                mime = node.getAttribute("ContentType")
                if node.localName == "Default":
                    key = node.getAttribute("Extension").casefold()
                    table = defaults
                elif node.localName == "Override":
                    key = resolve_target("", node.getAttribute("PartName"))
                    table = overrides
                else:
                    raise ProcessingError("Unsupported content-type element")
                if not key or not mime or key in table:
                    raise ProcessingError("Missing or duplicate content-type fields")
                table[key] = mime
            return defaults, overrides

    def remove_content_type_overrides(self, deleted: Set[str]) -> None:
        if not deleted:
            return
        name = "[Content_Types].xml"
        with xml_document(self.read_xml(name), name) as doc:
            changed = False
            for node in list(children(doc.documentElement)):
                if (node.localName == "Override"
                        and resolve_target("", node.getAttribute("PartName")) in deleted):
                    node.parentNode.removeChild(node)
                    node.unlink()
                    changed = True
            if changed:
                self.save_xml(name, doc)


def paragraph_text(node) -> str:
    result = []
    for part in elements(node):
        if matches(part, A_NS, "t"):
            result.append(text_content(part))
        elif matches(part, A_NS, "br"):
            result.append("\n")
        elif matches(part, A_NS, "tab"):
            result.append("\t")
    return "".join(result)


def editable_text(root, table_count: Optional[List[int]] = None) -> Iterator[str]:
    """XML order; table cells stay in rows rather than becoming loose paragraphs."""
    if table_count is None:
        table_count = [0]
    if matches(root, A_NS, "tbl"):
        table_count[0] += 1
        yield f"[Table {table_count[0]} - columns separated by tabs; cell breaks shown as \\n]"
        for row in children(root):
            if not matches(row, A_NS, "tr"):
                continue
            values = []
            for cell in children(row):
                if matches(cell, A_NS, "tc"):
                    paragraphs = [paragraph_text(p) for p in elements(cell)
                                  if matches(p, A_NS, "p")]
                    value = "\n".join(paragraphs)
                    values.append(value.replace("\\", "\\\\").replace("\t", "\\t")
                                  .replace("\r", "\\r").replace("\n", "\\n"))
            yield "\t".join(values)
        yield "[End table]"
    elif matches(root, A_NS, "p"):
        value = paragraph_text(root)
        if value.strip():
            yield value
    else:
        for child in children(root):
            yield from editable_text(child, table_count)


def related_text_parts(root, rels) -> List[Tuple[str, Relationship]]:
    parts = []
    seen = set()
    for node in elements(root):
        role, rid = None, ""
        if matches(node, C_NS, "chart"):
            role, rid = "Chart", rel_attr(node)
        elif matches(node, DGM_NS, "relIds"):
            role, rid = "SmartArt", rel_attr(node, "dm")
        if not role:
            continue
        rel = rels.get(rid)
        expected = "/chart" if role == "Chart" else "/diagramData"
        if not rel or not rel.kind.endswith(expected):
            raise ProcessingError(f"Invalid {role} relationship: {rid}")
        if (role, rid) not in seen:
            parts.append((role, rel))
            seen.add((role, rid))
    return parts


def write_slide_text(handle, pkg: Package, slide: str, doc, number: int,
                     rels, include_notes: bool, stats: Stats) -> None:
    handle.write(f"========== Slide {number} ==========\n")
    hidden = doc.documentElement.getAttribute("show") in {"0", "false"}
    if hidden:
        handle.write("[Hidden slide - included in presentation order]\n")
    found = False
    for line in editable_text(doc.documentElement):
        found = True
        handle.write(line + "\n")
    if not found:
        handle.write("(No editable text in the slide's own shapes.)\n")
    for role, rel in related_text_parts(doc.documentElement, rels):
        if rel.external:
            stats.warn(f"Slide {number}: external {role} text was not fetched.")
            continue
        handle.write(f"\n[{role}: stored editable text / cached labels]\n")
        with xml_document(pkg.read_xml(rel.target), rel.target) as related:
            texts = []
            for line in editable_text(related.documentElement):
                if line not in texts:
                    texts.append(line)
            if role == "Chart":
                for node in elements(related.documentElement):
                    if (node.namespaceURI in C_NS and
                            node.localName in {"strCache", "strLit"}):
                        for value in elements(node):
                            if matches(value, C_NS, "v"):
                                text = text_content(value)
                                if text.strip() and text not in texts:
                                    texts.append(text)
            for line in texts:
                handle.write(line + "\n")
            if not texts:
                handle.write("(No supported stored text found.)\n")
    if include_notes:
        for rel in rels.values():
            if rel.external or not rel.kind.endswith("/notesSlide"):
                continue
            handle.write("\n[Speaker notes]\n")
            with xml_document(pkg.read_xml(rel.target), rel.target) as notes:
                any_note = False
                for shape in elements(notes.documentElement):
                    if not matches(shape, P_NS, "sp"):
                        continue
                    placeholders = [n.getAttribute("type") for n in elements(shape)
                                    if matches(n, P_NS, "ph")]
                    if any(p in {"dt", "hdr", "ftr", "sldNum", "sldImg"}
                           for p in placeholders):
                        continue
                    for line in editable_text(shape):
                        any_note = True
                        handle.write(line + "\n")
                if not any_note:
                    handle.write("(No supported speaker-note text found.)\n")
    handle.write("\n")


def shape_id(node) -> str:
    for child in children(node):
        if child.namespaceURI in P_NS and child.localName.startswith("nv"):
            for item in children(child):
                if matches(item, P_NS, "cNvPr"):
                    return item.getAttribute("id")
    return ""


def referenced_shape_ids(root) -> Set[str]:
    result = set()
    for node in elements(root):
        for attr in attrs(node):
            name = (attr.localName or attr.name).lower()
            if name in {"spid", "shapeid", "shaperef", "objectid"}:
                result.add(attr.value)
            elif name == "id" and not matches(node, P_NS, "cNvPr"):
                # Includes connector stCxn/endCxn and conservative unknown uses.
                # Timing-node IDs can overprotect a similarly numbered picture.
                result.add(attr.value)
    return result


def picture_protection_reason(pic, referenced: Set[str], rels) -> Optional[str]:
    ident = shape_id(pic)
    if not ident:
        return "missing shape identity"
    ancestor = pic
    while ancestor is not None and ancestor.nodeType == Node.ELEMENT_NODE:
        if ancestor.namespaceURI == MC_NS:
            return "compatibility/alternate-content branch"
        if ancestor.namespaceURI in P_NS and ancestor.localName in {
                "oleObj", "contentPart", "control", "graphicFrame"}:
            return "embedded object or object preview"
        if shape_id(ancestor) in referenced:
            return "animation, connector, or other shape reference"
        ancestor = ancestor.parentNode
    if (pic.parentNode.namespaceURI not in P_NS
            or pic.parentNode.localName not in {"spTree", "grpSp"}):
        return "unsupported picture container"
    has_image = False
    for node in elements(pic):
        local = (node.localName or "").lower()
        if local in {"videofile", "audiofile", "wavaudiofile", "media", "oleobj",
                     "control", "contentpart", "hlinkclick", "hlinkmouseover"}:
            return "multimedia or interactive object"
        for attr in attrs(node):
            if attr.namespaceURI in R_NS:
                rel = rels.get(attr.value)
                if not rel:
                    return "unresolved object relationship"
                if not rel.kind.endswith("/image"):
                    return "non-image object relationship"
                has_image = True
        known_namespace = (node.namespaceURI in P_NS or node.namespaceURI in A_NS
                           or node.namespaceURI == MC_NS)
        known_extension = ((node.namespaceURI, node.localName) in {
            (A14_NS, "useLocalDpi"), (A16_NS, "creationId"), (SVG_NS, "svgBlip")})
        if not known_namespace and not known_extension:
            return "unrecognized picture extension"
    if not has_image:
        return "no identifiable image relationship"
    return None


def clean_slide(pkg: Package, slide: str, doc, rels,
                stats: Stats) -> Set[str]:
    root = doc.documentElement
    identities = {}
    for node in elements(root):
        if matches(node, P_NS, "cNvPr"):
            ident = node.getAttribute("id")
            if not ident:
                raise ProcessingError(f"{slide}: a shape has no ID; cleanup refused")
            identities.setdefault(ident, []).append(node)

    def compatibility_branches(node):
        selected = {}
        while node is not None and node.nodeType == Node.ELEMENT_NODE:
            parent = node.parentNode
            if (node.namespaceURI == MC_NS and node.localName in {"Choice", "Fallback"}
                    and parent.nodeType == Node.ELEMENT_NODE
                    and parent.namespaceURI == MC_NS and parent.localName == "AlternateContent"):
                selected[id(parent)] = id(node)
            node = parent
        return selected

    for nodes in identities.values():
        # Choice and Fallback can legitimately define the same logical shape ID.
        branches = [compatibility_branches(node) for node in nodes]
        for left in range(len(nodes)):
            for right in range(left + 1, len(nodes)):
                mutually_exclusive = any(
                    key in branches[right] and choice != branches[right][key]
                    for key, choice in branches[left].items())
                if not mutually_exclusive:
                    raise ProcessingError(f"{slide}: duplicate shape IDs outside alternative branches; cleanup refused")
    referenced = referenced_shape_ids(root)
    removed_ids = set()
    removed_count = 0
    for pic in [n for n in elements(root) if matches(n, P_NS, "pic")]:
        reason = picture_protection_reason(pic, referenced, rels)
        if reason:
            stats.protected_pictures += 1
            stats.protection_reasons[reason] = stats.protection_reasons.get(reason, 0) + 1
            continue
        for node in elements(pic):
            for attr in attrs(node):
                if attr.namespaceURI in R_NS:
                    removed_ids.add(attr.value)
        pic.parentNode.removeChild(pic)
        pic.unlink()
        removed_count += 1
    if not removed_count:
        return set()
    stats.removed_pictures += removed_count
    pkg.save_xml(slide, doc)
    # Consider *all* remaining attribute values, including legacy relationship
    # attributes, before removing a relationship used by a removed picture.
    still_used = {attr.value for n in elements(root) for attr in attrs(n)}
    removable = {rid for rid in removed_ids if rid not in still_used
                 and rid in rels and rels[rid].kind.endswith("/image")}
    candidates = {rels[rid].target for rid in removable if not rels[rid].external}
    if removable:
        part = relationship_part(slide)
        with xml_document(pkg.read_xml(part), part) as rel_doc:
            for node in list(children(rel_doc.documentElement)):
                if node.getAttribute("Id") in removable:
                    node.parentNode.removeChild(node)
                    node.unlink()
            pkg.save_xml(part, rel_doc)
    return candidates


def copy_package(pkg: Package, destination: Path, deleted: Set[str]) -> None:
    """Bounded binary copying: retained media never uses ZipFile.read/writestr."""
    with zipfile.ZipFile(destination, "w", allowZip64=True) as output:
        output.comment = pkg.archive.comment
        for info in pkg.archive.infolist():
            if info.filename in deleted:
                continue
            # Reusing ZipInfo preserves compression method, timestamps and attrs.
            replacement = pkg.replacements.get(info.filename)
            out_info = copy.copy(info)
            if replacement:
                out_info.file_size = replacement.stat().st_size
            with output.open(out_info, "w", force_zip64=out_info.file_size >= zipfile.ZIP64_LIMIT) as target:
                if replacement:
                    with replacement.open("rb") as source:
                        shutil.copyfileobj(source, target, CHUNK)
                elif not info.is_dir():
                    with pkg.archive.open(info.filename, "r") as source:
                        shutil.copyfileobj(source, target, CHUNK)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(CHUNK), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_record(path: Path, recipe: str) -> dict:
    return {"sha256": sha256_file(path), "bytes": path.stat().st_size, "recipe": recipe}


def linklike(path: Path) -> bool:
    try:
        value = path.lstat()
        return stat.S_ISLNK(value.st_mode) or bool(
            getattr(value, "st_file_attributes", 0) & 0x400)
    except FileNotFoundError:
        return False


def publish_files(pairs: List[Tuple[Path, Path]], stage: Path) -> None:
    """Stage the whole set first; atomic per-file replace with error rollback.

    No portable filesystem operation can atomically replace two/three files at
    once. On ordinary errors (or Ctrl+C) undo published files. A process kill or
    power loss may still leave a partial set, which must not be called complete.
    """
    backups = {}
    published = []
    try:
        for index, (_prepared, destination) in enumerate(pairs):
            if destination.exists():
                backup = stage / f"previous_{index:03d}.bak"
                # Hardlinks do not duplicate a large existing presentation.
                try:
                    os.link(destination, backup)
                except OSError:
                    shutil.copy2(destination, backup)
                backups[destination] = backup
        for prepared, destination in pairs:
            os.replace(prepared, destination)
            published.append(destination)
    except BaseException as exc:
        failures = []
        for destination in reversed(published):
            try:
                if destination in backups:
                    os.replace(backups[destination], destination)
                else:
                    destination.unlink(missing_ok=True)
            except OSError as rollback_exc:
                failures.append(f"{destination}: {rollback_exc}")
        if failures:
            raise RecoveryError(
                f"Output publication failed and rollback was incomplete. Retained recovery "
                f"files: {stage}. Details: {'; '.join(failures)}") from exc
        raise


def output_recipes(options: Options) -> dict:
    return {"txt": f"R002-text-1;notes={int(options.include_notes)};charts=1;smartart=1",
            "pptx": "R002-conservative-clean-1"}


def load_report(path: Path) -> Optional[dict]:
    if not path.exists():
        return None
    if path.stat().st_size > 4 * 1024 * 1024:
        raise ProcessingError("Existing processing report is unexpectedly large")
    try:
        report = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeError) as exc:
        raise ProcessingError("Unreadable existing R002 report; use --overwrite to rebuild") from exc
    if (not isinstance(report, dict) or report.get("tool") != "PowerPoint Text Extractor"
            or report.get("version") != VERSION or not isinstance(report.get("outputs"), dict)):
        raise ProcessingError("Unrecognized processing report; use --overwrite to rebuild")
    return report


def process_file(source: Path, options: Options, log=print) -> dict:
    source = source.absolute()
    if source.suffix.lower() != ".pptx":
        raise ProcessingError("Only .pptx is supported; use Save As in PowerPoint to convert .ppt")
    if not source.is_file() or linklike(source):
        raise ProcessingError("Source must be a regular, non-linked .pptx file")
    output = source.parent / "cleaned"
    if linklike(output):
        raise ProcessingError("Refusing a linked/junction cleaned folder")
    output.mkdir(exist_ok=True)
    paths = {"txt": output / (source.stem + ".txt"), "pptx": output / source.name}
    report_path = output / (source.stem + ".r002.json")
    for path in list(paths.values()) + [report_path]:
        if linklike(path) or (path.exists() and not path.is_file()):
            raise ProcessingError(f"Output is not a regular file: {path}")
        if path.exists() and os.path.samefile(source, path):
            raise ProcessingError("An output points to the original presentation")
    before = source.stat()
    log("  Reading source fingerprint...")
    source_hash = sha256_file(source)
    recipes = output_recipes(options)
    required = ["txt"] if options.text_only else ["txt", "pptx"]
    report = None if options.overwrite else load_report(report_path)
    if report and report.get("source", {}).get("sha256") != source_hash:
        raise ProcessingError("Source changed since the saved outputs. Use --overwrite to refresh them")
    to_make = []
    records = dict(report.get("outputs", {})) if report else {}
    for key in required:
        path = paths[key]
        if options.overwrite or not path.exists():
            to_make.append(key)
            continue
        old = records.get(key)
        if not old:
            raise ProcessingError(
                f"Existing output is not verified by R002: {path.name}. "
                "Use --overwrite to regenerate cleaned outputs; originals are not changed")
        if old.get("recipe") != recipes[key]:
            raise ProcessingError("Extraction settings changed; use --overwrite to refresh outputs")
        if path.stat().st_size != old.get("bytes") or sha256_file(path) != old.get("sha256"):
            raise ProcessingError(f"Output was changed/incomplete: {path.name}. Use --overwrite to rebuild")
    if not to_make:
        return {"status": "skipped", "source": str(source), "report": str(report_path)}
    stage = Path(tempfile.mkdtemp(prefix=".r002-stage-", dir=output))
    keep_stage = False
    stats = Stats()
    prepared = {}
    try:
        with zipfile.ZipFile(source, "r") as archive:
            pkg = Package(archive, stage, options.max_xml_bytes)
            log("  Validating relationships and reading slide order...")
            pkg.validate_relationships()  # Malformed input must fail BEFORE deletion.
            defaults, overrides = pkg.content_types()
            if "pptx" in to_make:
                if any(n.lower().startswith("_xmlsignatures/") for n in pkg.info):
                    raise ProcessingError("Digitally signed package: use --text-only; cleanup would invalidate its signature")
                if any(n.lower().endswith("vbaproject.bin") for n in pkg.info):
                    raise ProcessingError("Macro-bearing package: cleanup is unsupported; use --text-only")
            slides = pkg.slide_order()
            stats.slides = len(slides)
            txt_path = stage / "text.txt"
            # Writing even when reusing TXT avoids holding extracted text in RAM;
            # only requested outputs will be published.
            with txt_path.open("w", encoding="utf-8-sig", newline="\n") as handle:
                handle.write(f"Source: {source.name}\nExtractor: {VERSION}\n")
                handle.write("Slide order: presentation order (hidden slides included).\n")
                handle.write("Slide text + standard cached chart labels + stored SmartArt text.\n")
                handle.write("Speaker notes: " + ("included" if options.include_notes else "not requested") + ".\n")
                handle.write("No OCR. Text within a slide follows XML order, not a reconstructed visual layout.\n\n")
                if not slides:
                    handle.write("(This presentation contains no slides.)\n")
                candidates = set()
                for index, slide in enumerate(slides, 1):
                    if index == 1 or index == len(slides) or index % 25 == 0:
                        log(f"  Slide {index}/{len(slides)}...")
                    rels = pkg.relationships(slide)
                    with xml_document(pkg.read_xml(slide), slide) as doc:
                        if not matches(doc.documentElement, P_NS, "sld"):
                            raise ProcessingError(f"{slide}: not a supported slide document")
                        for node in elements(doc.documentElement):
                            for attr in attrs(node):
                                if (attr.namespaceURI in R_NS and attr.value
                                        and attr.value not in rels):
                                    raise ProcessingError(
                                        f"{slide}: unresolved relationship ID {attr.value}")
                        if any(n.namespaceURI == MC_NS and n.localName == "AlternateContent"
                               for n in elements(doc.documentElement)):
                            stats.warn(f"Slide {index}: compatibility branches are exported in XML "
                                       "order; some text may repeat.")
                        write_slide_text(handle, pkg, slide, doc, index, rels,
                                         options.include_notes, stats)
                        if "pptx" in to_make:
                            candidates.update(clean_slide(pkg, slide, doc, rels, stats))
                handle.flush()
                os.fsync(handle.fileno())
            if "txt" in to_make:
                prepared["txt"] = txt_path
            if "pptx" in to_make:
                used = pkg.validate_relationships()
                deleted = set()
                for name in candidates - used:
                    mime = overrides.get(name, defaults.get(posixpath.splitext(name)[1][1:].lower(), ""))
                    # Only image parts detached by this run are eligible. Pre-existing
                    # orphan media and unknown media types are left alone.
                    if mime.lower().startswith("image/"):
                        deleted.add(name)
                    else:
                        stats.warn(f"Retained detached part with unknown/non-image content type: {name}")
                stats.removed_media = len(deleted)
                stats.removed_media_bytes = sum(pkg.info[n].file_size for n in deleted)
                pkg.remove_content_type_overrides(deleted)
                pkg.validate_relationships(deleted)
                pptx_path = stage / "presentation.pptx"
                log("  Copying retained content in 1 MiB chunks...")
                copy_package(pkg, pptx_path, deleted)
                with pptx_path.open("rb+") as handle:
                    os.fsync(handle.fileno())
                log("  Verifying cleaned package" + (" (CRC pass skipped)..." if options.skip_crc_check else " and ZIP CRC..."))
                with zipfile.ZipFile(pptx_path, "r") as verify:
                    checked = Package(verify, stage, options.max_xml_bytes)
                    checked.validate_relationships()
                    checked.content_types()
                    if checked.slide_order() != slides:
                        raise ProcessingError("Slide order changed during output creation")
                    for name in pkg.replacements:
                        with xml_document(checked.read_xml(name), name):
                            pass
                    if not options.skip_crc_check:
                        bad = verify.testzip()
                        if bad:
                            raise ProcessingError(f"Output ZIP integrity check failed: {bad}")
                prepared["pptx"] = pptx_path
        after = source.stat()
        if (before.st_size, before.st_mtime_ns, before.st_ino, before.st_dev) != (
                after.st_size, after.st_mtime_ns, after.st_ino, after.st_dev):
            raise ProcessingError("Source changed while processing; no output was published. Close PowerPoint and retry")
        for key, path in prepared.items():
            records[key] = file_record(path, recipes[key])
        # Do not claim an old unrequested file belongs to a newly overwritten run.
        if options.overwrite:
            records = {key: records[key] for key in prepared}
        if stats.protected_pictures:
            stats.warn("Protected pictures were intentionally retained. See protection_reasons.")
        cleanup_stats = vars(stats).copy() if "pptx" in prepared else (
            report.get("cleanup") if report else None)
        report_data = {
            "tool": "PowerPoint Text Extractor", "version": VERSION,
            "source": {"filename": source.name, "sha256": source_hash, "bytes": before.st_size},
            "outputs": records, "slides": stats.slides,
            "last_run": {"generated": list(prepared), "notes_requested": options.include_notes,
                         "output_crc_verified": "pptx" in prepared and not options.skip_crc_check},
            "cleanup": cleanup_stats,
            "warnings": stats.warnings,
            "limitations": ["No OCR; .pptx only", "Background/master/layout art is retained",
                            "Text order within slides is XML order", "Notes are opt-in",
                            "Chart and SmartArt extraction covers supported stored text, not all embedded content",
                            "No desktop PowerPoint compatibility guarantee"]}
        prepared_report = stage / "report.json"
        with prepared_report.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(report_data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        for key in prepared:
            if not options.overwrite and paths[key].exists():
                raise ProcessingError(
                    f"An output appeared while processing: {paths[key].name}. "
                    "It was not overwritten; avoid running two copies at once.")
        pairs = [(path, paths[key]) for key, path in prepared.items()]
        pairs.append((prepared_report, report_path))  # Completion record is published last.
        publish_files(pairs, stage)
        return {"status": "ok", "source": str(source), "generated": list(prepared),
                "stats": vars(stats), "output": str(output), "report": str(report_path)}
    except RecoveryError:
        keep_stage = True
        raise
    finally:
        if not keep_stage:
            shutil.rmtree(stage, ignore_errors=True)


def discover(folder: Path, recursive: bool) -> Tuple[List[Path], List[str], int]:
    files, errors = [], []
    unsupported = 0

    def scan_error(exc):
        errors.append(str(exc))

    for root, dirs, names in os.walk(folder, followlinks=False, onerror=scan_error):
        here = Path(root)
        dirs[:] = sorted(d for d in dirs if d.casefold() != "cleaned"
                         and not d.startswith(".r002-stage-") and not linklike(here / d))
        if not recursive:
            dirs[:] = []
        for name in sorted(names, key=str.casefold):
            path = here / name
            if name.startswith("~$") or linklike(path):
                continue
            if path.suffix.casefold() == ".pptx":
                files.append(path)
            elif path.suffix.casefold() in {".ppt", ".pptm"}:
                unsupported += 1
    return files, errors, unsupported


def main(argv: Optional[List[str]] = None) -> int:
    # Redirected Windows terminals can use legacy encodings. Keep console logging
    # from aborting a batch on a multilingual filename; TXT remains full UTF-8.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--folder", type=Path, help="Folder to scan; default is the current working folder")
    choice.add_argument("--file", type=Path, help="Process one .pptx instead of scanning folders")
    parser.add_argument("--text-only", action="store_true", help="Create TXT/report only, no cleaned presentation")
    parser.add_argument("--include-notes", action="store_true", help="Include speaker notes in TXT (off by default)")
    parser.add_argument("--overwrite", action="store_true", help="Replace cleaned outputs, never source presentations")
    parser.add_argument("--no-recursive", action="store_true", help="Do not scan subfolders")
    parser.add_argument("--skip-crc-check", action="store_true", help="Skip the final full output CRC read; faster but less verification")
    parser.add_argument("--max-xml-mb", type=int, default=32, help="Maximum uncompressed XML part size (MiB), default 32")
    parser.add_argument("--version", action="version", version="PowerPoint Text Extractor " + VERSION)
    args = parser.parse_args(argv)
    if args.max_xml_mb < 1:
        parser.error("--max-xml-mb must be at least 1")
    if args.file:
        if linklike(args.file.expanduser()):
            parser.error("Linked source files are not followed")
        files, errors, unsupported = [args.file.expanduser().absolute()], [], 0
    else:
        folder = (args.folder or Path.cwd()).expanduser().resolve()
        if not folder.is_dir():
            parser.error(f"Not a folder: {folder}")
        if folder.name.casefold() == "cleaned":
            parser.error("Select the source folder, not its cleaned output folder")
        files, errors, unsupported = discover(folder, not args.no_recursive)
    options = Options(args.text_only, args.overwrite, args.include_notes,
                      args.skip_crc_check, args.max_xml_mb * 1024 * 1024)
    if unsupported:
        print(f"NOTE: {unsupported} .ppt/.pptm file(s) ignored. Convert legacy .ppt with PowerPoint Save As; do not just rename it.")
    for error in errors:
        print("SCAN FAILED:", error, file=sys.stderr)
    if not files:
        print("No .pptx presentations found. Nothing was changed.")
        return 1 if errors else 2
    print(f"Found {len(files)} presentation(s). Originals will not be modified.", flush=True)
    success = skipped = 0
    failed = len(errors)
    for index, source in enumerate(files, 1):
        print(f"\n[{index}/{len(files)}] {source}", flush=True)
        try:
            result = process_file(source, options, log=lambda message: print(message, flush=True))
            if result["status"] == "skipped":
                skipped += 1
                print("  SKIP: requested R002 outputs are present and fingerprints match.")
            else:
                success += 1
                stats = result["stats"]
                print(f"  OK: {stats['slides']} slide(s); created {', '.join(result['generated'])} + report.")
                if "pptx" in result["generated"]:
                    print(f"  Removed {stats['removed_pictures']} picture shape(s), retained "
                          f"{stats['protected_pictures']} protected picture object(s).")
                    print(f"  Removed {stats['removed_media']} unreferenced image part(s) "
                          f"({stats['removed_media_bytes'] / 1024**2:.2f} MiB uncompressed).")
                for warning in stats["warnings"]:
                    print("  NOTE:", warning)
                print("  Output:", result["output"])
        except KeyboardInterrupt:
            print("\nCancelled. Originals are unchanged; do not treat an interrupted set as complete.", file=sys.stderr)
            return 130
        except Exception as exc:
            failed += 1
            print(f"  FAILED: {exc}", file=sys.stderr, flush=True)
    print(f"\nFinished: {success} succeeded, {skipped} skipped, {failed} failed.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
