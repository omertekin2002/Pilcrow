"""Read WordprocessingML without modifying the package or requiring Word.

This is a prose view, not a Word renderer. All omissions are reported in metadata.
"""

from __future__ import annotations

import hashlib
import json
import posixpath
import re
from collections import Counter
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypeAlias
from zipfile import BadZipFile, ZipFile

from defusedxml import ElementTree as ET

from .prose import normalize, sentences

if TYPE_CHECKING:
    # Only the node type; all parsing goes through defusedxml.
    from xml.etree.ElementTree import Element

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W}
# Shape content that a rendered text box still leaves out of the prose.
UNRENDERED_SHAPE_PARTS = frozenset({"pic", "chart", "relIds", "imagedata"})
TEXT_BOX_NOTE = ("Text box text is shown in [Text box] blocks after its anchoring paragraph; "
                 "position, anchoring, wrapping and shape formatting are not compared.")
LIST_NOTE = ("List boundaries and indentation are retained; bullet and number formats are recorded in "
             "metadata, but counter values and bullet glyphs are not reconstructed.")
UNKNOWN_LIST_NOTE = ("Some list numbering formats could not be resolved (no numbering part, an unknown "
                     "list or level, or style-linked numbering); they are recorded as unknown.")
# Relationship Id -> (type suffix, target, is external).
Relationships: TypeAlias = dict[str | None, tuple[str, str, bool]]


def w(name: str) -> str:
    return f"{{{W}}}{name}"


def attr(node: Element, name: str, default: str = "") -> str:
    return node.get(w(name), default)


class DocumentError(ValueError):
    """An input is unavailable or cannot be safely interpreted."""


@dataclass
class Extraction:
    prose: str
    metadata: dict[str, Any]

    def outputs(self) -> dict[str, bytes]:
        return {
            ".prose.md": self.prose.encode("utf-8"),
            ".meta.json": (json.dumps(self.metadata, ensure_ascii=False, indent=2,
                                      sort_keys=True) + "\n").encode("utf-8"),
        }


class Reader:
    def __init__(self, package: ZipFile) -> None:
        self.package = package
        self.names = set(package.namelist())
        self.warnings: set[str] = {
            "Formatting and page layout are not compared; unchanged prose is not an identical document."
        }
        self.review: list[dict[str, str]] = []
        self.links: list[dict[str, str]] = []
        self.fields: list[dict[str, str]] = []
        self.styles: dict[str, Element] = {}
        self.numbering: Element | None = None
        # Raw numId -> list entry; insertion order is first use, so ordinals ignore raw IDs.
        self.lists: dict[str, dict[str, Any]] = {}
        self.field_stack: list[list[str]] = []
        self.part = "word/document.xml"
        self.rels: Relationships = {}
        self.features: Counter[str] = Counter()
        # Text boxes found while reading a paragraph, rendered as blocks after it.
        self.anchored: list[list[Element]] = []

    def xml(self, name: str) -> Element:
        try:
            return ET.fromstring(self.package.read(name))
        except Exception as exc:
            raise DocumentError(f"cannot parse {name}: {exc}") from exc

    def relationships(self, part: str) -> Relationships:
        path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        result: Relationships = {}
        if path not in self.names:
            return result
        for rel in self.xml(path):
            target = rel.get("Target", "")
            external = rel.get("TargetMode") == "External"
            if not external:
                target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
                target = target.lstrip("/")
                if target.startswith("../"):
                    raise DocumentError(f"relationship escapes package: {path}")
            result[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], target, external)
        return result

    def warn(self, message: str) -> None:
        self.warnings.add(message)

    def raw_text(self, node: Element) -> str:
        return normalize("".join(
            (n.text or "") if n.tag in {w("t"), w("delText")} else " "
            for n in self.walk(node) if n.tag in {w("t"), w("delText"), w("tab"), w("br"), w("cr")}
        ))

    def walk(self, node: Element) -> Iterator[Element]:
        """Like ``node.iter()``, but read only the rendered copy of a text box."""
        yield node
        branch = self.branch(node) if node.tag == f"{{{MC}}}AlternateContent" else None
        for child in (branch if branch is not None else node):
            yield from self.walk(child)

    def branch(self, alternate: Element) -> Element | None:
        """The one copy of an mc:AlternateContent text box to render, or None if it has none.

        Word stores a modern text box twice: DrawingML in mc:Choice and VML in mc:Fallback.
        """
        def has_box(node: Element) -> bool:
            return any(d.tag == w("txbxContent") for d in node.iter())
        choice = next((c for c in alternate if c.tag == f"{{{MC}}}Choice" and has_box(c)), None)
        if choice is not None:
            return choice
        fallback = alternate.find(f"{{{MC}}}Fallback")
        return fallback if fallback is not None and has_box(fallback) else None

    def text_boxes(self, node: Element) -> list[Element]:
        """Outermost rendered w:txbxContent elements in document order; nested ones render inside them."""
        if node.tag == w("txbxContent"):
            return [node]
        if node.tag == f"{{{MC}}}AlternateContent":
            branch = self.branch(node)
            return self.text_boxes(branch) if branch is not None else []
        return [box for child in node for box in self.text_boxes(child)]

    def unrendered_shape_parts(self, node: Element) -> bool:
        """Whether a shape holding text boxes also holds pictures, charts or diagrams outside them."""
        if node.tag == w("txbxContent"):
            return False
        if node.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
            return self.unrendered_shape_parts(branch)
        # Alternatives without a text box (e.g. a grouped picture) are checked in full.
        return (node.tag.rsplit("}", 1)[-1] in UNRENDERED_SHAPE_PARTS
                or any(self.unrendered_shape_parts(child) for child in node))

    def anchor(self, node: Element) -> tuple[str, list[str]]:
        """Read ``node`` inline and render the text boxes it anchors as separate blocks."""
        text, boxes = self.anchor_boxes(node)
        return text, [self.text_box(box) for box in boxes]

    def anchor_boxes(self, node: Element) -> tuple[str, list[Element]]:
        """Read ``node`` inline and return the text boxes it anchors, not yet rendered."""
        self.anchored.append([])
        text = self.inline(node)
        return text, self.anchored.pop()

    def text_box(self, box: Element) -> str:
        self.features["text boxes"] += 1
        self.warn(TEXT_BOX_NOTE)
        return "\n\n".join(("[Text box]", self.blocks(box) or "(empty)", "[/Text box]"))

    def inline(self, node: Element) -> str:
        tag = node.tag
        local = tag.rsplit("}", 1)[-1]
        if tag in {w("del"), w("ins"), w("moveFrom"), w("moveTo")}:
            self.review.append({"kind": local, "author": attr(node, "author"),
                                "text": self.raw_text(node), "part": self.part})
            if local in {"del", "moveFrom"}:
                return ""
        if tag == w("t"):
            # XML source line wrapping is whitespace, unlike Word's explicit w:br.
            return re.sub(r"\s", " ", node.text or "")
        if tag == w("delText"):
            return ""
        if tag in {w("br"), w("cr")}:
            return "\n"
        if tag == w("tab"):
            return " "
        if tag == w("noBreakHyphen"):
            return "\u2011"
        if tag == w("softHyphen"):
            return "\u00ad"
        if tag in {w("footnoteReference"), w("endnoteReference")}:
            kind = "footnote" if tag == w("footnoteReference") else "endnote"
            self.features[kind + " references"] += 1
            return f"[^{kind}-{attr(node, 'id')}]"
        if tag == w("hyperlink"):
            rid = node.get(f"{{{R}}}id")
            rel = self.rels.get(rid)
            target = rel[1] if rel else "#" + attr(node, "anchor")
            if rid and not rel:
                self.warn("Some hyperlink relationships are missing; destinations may be incomplete.")
            visible = "".join(self.inline(child) for child in node)
            self.links.append({"text": normalize(visible), "target": target, "part": self.part})
            return visible
        if tag == w("fldChar"):
            kind = attr(node, "fldCharType")
            if kind == "begin":
                self.field_stack.append([])
            elif kind == "end" and self.field_stack:
                self.record_field("".join(self.field_stack.pop()))
            self.warn("Fields use saved display text, which can be stale; codes are fingerprinted, not evaluated (including Zotero).")
            return ""
        if tag == w("instrText"):
            if self.field_stack:
                self.field_stack[-1].append(node.text or "")
            else:
                self.record_field(node.text or "")
            return ""
        if tag == w("fldSimple"):
            self.record_field(attr(node, "instr"))
            self.warn("Fields use saved display text, which can be stale; codes are fingerprinted, not evaluated (including Zotero).")
        if tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
            return "".join(self.inline(child) for child in branch)
        if tag in {w("drawing"), w("pict"), w("txbxContent")} and self.anchored and (boxes := self.text_boxes(node)):
            self.anchored[-1].extend(boxes)
            if not self.unrendered_shape_parts(node):
                return ""
        if local in {"drawing", "pict", "object", "txbxContent", "oMath", "oMathPara", "AlternateContent", "altChunk", "sym"}:
            self.warn(f"Unsupported {local}: content is not rendered in prose; inspect the DOCX in Word.")
            self.features[local] += 1
            return f"[Unsupported {local}]"
        if tag in {w("pPr"), w("rPr"), w("sdtPr"), w("sdtEndPr")}:
            return ""
        return "".join(self.inline(child) for child in node)

    def record_field(self, instruction: str) -> None:
        value = normalize(instruction)
        self.fields.append({"command": value.split(" ", 1)[0],
                            "sha256": hashlib.sha256(value.encode()).hexdigest(), "part": self.part})

    def properties(self, paragraph: Element) -> tuple[list[Element], str]:
        direct = paragraph.find("w:pPr", NS)
        style = direct.find("w:pStyle", NS) if direct is not None else None
        style_id = attr(style, "val") if style is not None else ""
        properties = []
        if direct is not None:
            properties.append(direct)
        seen = set()
        while style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            current = self.styles[style_id]
            prop = current.find("w:pPr", NS)
            if prop is not None:
                properties.append(prop)
            based = current.find("w:basedOn", NS)
            style_id = attr(based, "val") if based is not None else ""
        return properties, attr(style, "val") if style is not None else ""

    def paragraph(self, node: Element) -> str:
        text, boxes = self.anchor_boxes(node)
        # Render the paragraph before its boxes so list first use follows prose order.
        own = self.paragraph_text(node, text)
        return "\n\n".join(part for part in (own, *(self.text_box(box) for box in boxes)) if part)

    def paragraph_text(self, node: Element, text: str) -> str:
        lines = sentences(text)
        if not lines:
            return ""
        props, style_id = self.properties(node)
        outline = next((p.find("w:outlineLvl", NS) for p in props
                        if p.find("w:outlineLvl", NS) is not None), None)
        heading = re.fullmatch(r"Heading([1-9])", style_id, flags=re.I)
        if outline is not None and attr(outline, "val").isdigit() and int(attr(outline, "val")) < 9:
            level = min(6, int(attr(outline, "val")) + 1)
        else:
            level = min(6, int(heading[1])) if heading else 0
        num = next((p.find("w:numPr", NS) for p in props if p.find("w:numPr", NS) is not None), None)
        if level:
            return "#" * level + " " + normalize(text)
        if num is not None:
            num_id = num.find("w:numId", NS)
            if num_id is None or attr(num_id, "val") != "0":
                self.warn(LIST_NOTE)
                ilvl = num.find("w:ilvl", NS)
                depth = attr(ilvl, "val", "0") if ilvl is not None else "0"
                self.record_list(props, num, int(depth) if depth.isdigit() else 0)
                indent = "  " * min(20, int(depth)) if depth.isdigit() else ""
                return indent + "- " + ("\n" + indent + "  ").join(lines)
        return "\n".join(lines)

    def record_list(self, props: list[Element], num: Element, level: int) -> None:
        """Fingerprint the effective format of one used list level, keyed by first use."""
        # A numPr may carry only ilvl and inherit numId from its style chain.
        found = num.find("w:numId", NS)
        if found is None:
            found = next((n for p in props if (n := p.find("w:numPr/w:numId", NS)) is not None), None)
        num_id = attr(found, "val") if found is not None else None
        key = num_id if num_id is not None else ""
        entry = self.lists.setdefault(key, {"list": len(self.lists) + 1, "part": self.part, "levels": {}})
        if level not in entry["levels"]:
            entry["levels"][level] = self.list_level(num_id, level)

    def list_level(self, num_id: str | None, level: int) -> dict[str, Any]:
        """Resolve w:num -> w:abstractNum -> w:lvl, applying w:lvlOverride."""
        unknown = {"level": level, "format": "unknown", "text": "", "start": 0}
        numbering = self.numbering
        num = None if numbering is None or num_id is None else next(
            (n for n in numbering.findall("w:num", NS) if attr(n, "numId") == num_id), None)
        abstract_id = num.find("w:abstractNumId", NS) if num is not None else None
        abstract = None if numbering is None or abstract_id is None else next(
            (a for a in numbering.findall("w:abstractNum", NS)
             if attr(a, "abstractNumId") == attr(abstract_id, "val")), None)
        if num is None or abstract is None or abstract.find("w:numStyleLink", NS) is not None:
            self.warn(UNKNOWN_LIST_NOTE)
            return unknown
        override = next((o for o in num.findall("w:lvlOverride", NS) if attr(o, "ilvl") == str(level)), None)
        lvl = override.find("w:lvl", NS) if override is not None else None
        if lvl is None:
            lvl = next((v for v in abstract.findall("w:lvl", NS) if attr(v, "ilvl") == str(level)), None)
        if lvl is None:
            self.warn(UNKNOWN_LIST_NOTE)
            return unknown
        fmt = lvl.find("w:numFmt", NS)
        text = lvl.find("w:lvlText", NS)
        start = lvl.find("w:start", NS)
        start_override = override.find("w:startOverride", NS) if override is not None else None
        if start_override is not None:
            start = start_override
        start_value = attr(start, "val") if start is not None else ""
        return {"level": level,
                # OOXML defaults an absent w:numFmt to decimal.
                "format": attr(fmt, "val", "decimal") if fmt is not None else "decimal",
                "text": attr(text, "val") if text is not None else "",
                "start": int(start_value) if re.fullmatch(r"-?[0-9]+", start_value) else 0}

    def blocks(self, node: Element) -> str:
        parts = []
        for child in node:
            if child.tag == w("p"):
                value = self.paragraph(child)
            elif child.tag == w("tbl"):
                value = self.table(child)
            elif child.tag in {w("sectPr"), w("tcPr"), w("trPr"), w("tblPr"), w("tblGrid"), w("sdtPr"), w("sdtEndPr")}:
                continue
            elif child.tag in {w("del"), w("moveFrom")}:
                self.inline(child)  # record review, omit deleted blocks
                continue
            elif child.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(child)) is not None:
                value = self.blocks(branch)  # one copy of block-level text box content
            elif child.tag in {w("altChunk"), w("drawing"), w("pict"), w("object")}:
                text, boxes = self.anchor(child)
                value = "\n\n".join(part for part in (text, *boxes) if part)
            else:
                if child.tag in {w("ins"), w("moveTo")}:
                    self.review.append({"kind": child.tag.rsplit("}", 1)[-1], "author": attr(child, "author"),
                                        "text": self.raw_text(child), "part": self.part})
                value = self.blocks(child)
            if value:
                parts.append(value)
        return "\n\n".join(parts)

    def table_items(self, node: Element, tag: str) -> Iterator[Element]:
        """Unwrap content controls without descending into nested tables."""
        for child in node:
            if child.tag == w(tag):
                yield child
            elif child.tag in {w("sdt"), w("sdtContent"), w("customXml"), w("ins"), w("moveTo")}:
                if child.tag in {w("ins"), w("moveTo")}:
                    self.warn("Tracked table structure changes are only partially represented; inspect the DOCX.")
                yield from self.table_items(child, tag)
            elif child.tag in {w("del"), w("moveFrom")}:
                self.inline(child)

    def table(self, node: Element) -> str:
        self.features["tables"] += 1
        lines = ["[Table]"]
        # Explicit cells avoid Markdown pipe escaping and ambiguity for multi-paragraph cells.
        for row in self.table_items(node, "tr"):
            if row.find("w:trPr/w:del", NS) is not None:
                self.warn("Tracked table structure changes are only partially represented; inspect the DOCX.")
                continue
            lines.append("[Row]")
            for cell in self.table_items(row, "tc"):
                span = cell.find("w:tcPr/w:gridSpan", NS)
                merge = cell.find("w:tcPr/w:vMerge", NS)
                label = "[Cell"
                if span is not None:
                    label += " span=" + attr(span, "val")
                if merge is not None:
                    label += " vertical-merge=" + attr(merge, "val", "continue")
                lines.extend([label + "]", self.blocks(cell) or "(empty)"])
            lines.append("[/Row]")
        lines.append("[/Table]")
        return "\n\n".join(lines)

    def audit(self, tree: Element) -> None:
        # This small allowlist is intentionally conservative: unknown constructs warn.
        known = set("document body p r t delText tab br cr noBreakHyphen softHyphen pPr rPr sectPr tbl tr tc tblPr trPr tcPr tblGrid gridCol sdt sdtContent sdtPr sdtEndPr customXml smartTag ins del moveFrom moveTo hyperlink bookmarkStart bookmarkEnd proofErr permStart permEnd commentRangeStart commentRangeEnd commentReference footnoteReference endnoteReference footnotes footnote endnotes endnote footnoteRef endnoteRef comments comment hdr ftr fldChar instrText fldSimple lastRenderedPageBreak moveFromRangeStart moveFromRangeEnd moveToRangeStart moveToRangeEnd".split())
        def visit(node: Element) -> None:
            local = node.tag.rsplit("}", 1)[-1]
            if node.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
                # Only the copy holding the rendered text box is read.
                for child in branch:
                    visit(child)
                return
            if node.tag in {w("drawing"), w("pict"), w("txbxContent")}:
                # A rendered text box: audit its Word content, not the shape markup around it.
                # inline() warns for pictures or charts that share the shape.
                boxes = self.text_boxes(node)
                if boxes:
                    for box in boxes:
                        for child in box:
                            visit(child)
                    return
            if node.tag.startswith("{" + W + "}") and local not in known:
                self.warn(f"Unsupported Word construct {local}; its semantics are not fully compared.")
            if local.endswith("Pr") or local in {"tblGrid", "sdtEndPr"}:
                for descendant in node.iter():
                    name = descendant.tag.rsplit("}", 1)[-1]
                    if name in {"vanish", "webHidden", "specVanish"}:
                        self.warn("Hidden text is included; visibility formatting is not applied.")
                    if name.endswith("Change") or name in {"ins", "del", "cellIns", "cellDel", "cellMerge"}:
                        self.warn("Tracked formatting or structure changes are not fully represented; inspect the DOCX.")
                return
            if not node.tag.startswith("{" + W + "}"):
                self.warn(f"Unsupported XML construct {local}; content may be omitted or approximate.")
            for child in node:
                visit(child)
        visit(tree)

    def render_part(self, part: str, tree: Element) -> str:
        self.part, self.rels = part, self.relationships(part)
        self.audit(tree)
        value = self.blocks(tree)
        if self.field_stack:
            self.warn("Unclosed field codes found; saved display text may be incomplete.")
            while self.field_stack:
                self.record_field("".join(self.field_stack.pop()))
        return value

    def extract(self) -> Extraction:
        if "word/document.xml" not in self.names:
            raise DocumentError("missing word/document.xml; expected a standard DOCX package")
        main = self.xml("word/document.xml")
        body = main.find("w:body", NS)
        if main.tag != w("document") or body is None:
            raise DocumentError("unsupported Word XML dialect or missing document body (Strict OOXML is not supported)")
        main_rels = self.relationships("word/document.xml")
        supported = {"word/document.xml"}
        for kind, target, external in main_rels.values():
            if kind == "styles" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                self.styles = {attr(s, "styleId"): s for s in self.xml(target)}
                supported.add(target)
            elif kind == "numbering" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                self.numbering = self.xml(target)
                supported.add(target)
        prose = [self.render_part("word/document.xml", body)]
        # All supported related stories are distinct; never combine headers with body prose.
        stories = sorted({(kind, target) for kind, target, external in main_rels.values()
                          if kind in {"footnotes", "endnotes", "comments", "header", "footer"} and not external})
        for kind, part in stories:
            if part not in self.names:
                raise DocumentError(f"missing related part {part}")
            supported.add(part)
            tree = self.xml(part)
            if kind == "comments":
                self.audit(tree)
                for comment in tree:
                    self.review.append({"kind": "comment", "author": attr(comment, "author"),
                                        "text": self.raw_text(comment), "part": part})
                self.warn("Comment text and authors are recorded; anchors, threads, dates, and resolution status are not compared.")
            elif kind in {"footnotes", "endnotes"}:
                for note in tree:
                    if attr(note, "type") in {"separator", "continuationSeparator", "continuationNotice"}:
                        continue
                    singular = kind[:-1]
                    prose.append(f"[^{singular}-{attr(note, 'id')}]:\n\n" + self.render_part(part, note))
            else:
                prose.append(f"## {kind.title()} ({part})\n\n" + self.render_part(part, tree))
                self.warn("Headers and footers are shown by part; page/section placement is not reconstructed.")
        for name in sorted(self.names):
            if name in supported or name.endswith("/") or name.endswith(".rels") or name == "[Content_Types].xml" or name.startswith("docProps/"):
                continue
            if name in {"word/settings.xml", "word/webSettings.xml", "word/fontTable.xml"} or name.startswith("word/theme/"):
                continue
            self.warn(f"Unrendered package part: {name}; changes to its content are not compared.")
        if self.features["footnote references"] and not any(k == "footnotes" for k, _ in stories):
            self.warn("Footnote references have no related footnotes part.")
        if self.features["endnote references"] and not any(k == "endnotes" for k, _ in stories):
            self.warn("Endnote references have no related endnotes part.")
        metadata = {"schema": 1, "view": "accept-insertions-omit-deletions",
                    "coverage": "partial-prose", "warnings": sorted(self.warnings),
                    "review": self.review, "hyperlinks": self.links, "fields": self.fields,
                    "lists": [{**entry, "levels": [entry["levels"][k] for k in sorted(entry["levels"])]}
                              for entry in self.lists.values()]}
        return Extraction("\n\n".join(p for p in prose if p).rstrip() + "\n", metadata)


def extract(path: Path) -> Extraction:
    try:
        before = path.stat()
        with ZipFile(path) as package:
            infos = package.infolist()
            if len(infos) > 4096 or sum(i.file_size for i in infos) > 64 * 1024 * 1024:
                raise DocumentError("package exceeds safety limits (4096 entries / 64 MiB uncompressed)")
            if any(i.file_size > 16 * 1024 * 1024 for i in infos):
                raise DocumentError("package part exceeds 16 MiB safety limit")
            if len({i.filename for i in infos}) != len(infos):
                raise DocumentError("duplicate ZIP entries")
            result = Reader(package).extract()
        after = path.stat()
        if (before.st_ino, before.st_size, before.st_mtime_ns) != (after.st_ino, after.st_size, after.st_mtime_ns):
            raise DocumentError("document changed while reading; finish saving and retry")
        return result
    except (OSError, BadZipFile, RuntimeError, ValueError) as exc:
        if isinstance(exc, DocumentError):
            raise
        raise DocumentError(f"cannot read DOCX: {exc}; finish saving, check permissions, and retry") from exc
