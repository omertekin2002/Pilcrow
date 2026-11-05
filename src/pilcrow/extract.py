"""Read WordprocessingML without modifying the package or requiring Word.

This is a prose view, not a Word renderer. All omissions are reported in metadata.
"""

from collections import Counter
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import posixpath
import re
from zipfile import BadZipFile, ZipFile

from defusedxml import ElementTree as ET

from .prose import normalize, sentences

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W}


def w(name):
    return f"{{{W}}}{name}"


def attr(node, name, default=""):
    return node.get(w(name), default)


class DocumentError(ValueError):
    """An input is unavailable or cannot be safely interpreted."""


@dataclass
class Extraction:
    prose: str
    metadata: dict

    def outputs(self):
        return {
            ".prose.md": self.prose.encode("utf-8"),
            ".meta.json": (json.dumps(self.metadata, ensure_ascii=False, indent=2,
                                      sort_keys=True) + "\n").encode("utf-8"),
        }


class Reader:
    def __init__(self, package):
        self.package = package
        self.names = set(package.namelist())
        self.warnings = {
            "Formatting and page layout are not compared; unchanged prose is not an identical document."
        }
        self.review = []
        self.links = []
        self.fields = []
        self.styles = {}
        self.field_stack = []
        self.part = "word/document.xml"
        self.rels = {}
        self.features = Counter()

    def xml(self, name):
        try:
            return ET.fromstring(self.package.read(name))
        except Exception as exc:
            raise DocumentError(f"cannot parse {name}: {exc}") from exc

    def relationships(self, part):
        path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        result = {}
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

    def warn(self, message):
        self.warnings.add(message)

    def raw_text(self, node):
        return normalize("".join(
            (n.text or "") if n.tag in {w("t"), w("delText")} else " "
            for n in node.iter() if n.tag in {w("t"), w("delText"), w("tab"), w("br"), w("cr")}
        ))

    def inline(self, node):
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
        if local in {"drawing", "pict", "object", "txbxContent", "oMath", "oMathPara", "AlternateContent", "altChunk", "sym"}:
            self.warn(f"Unsupported {local}: content is not rendered in prose; inspect the DOCX in Word.")
            self.features[local] += 1
            return f"[Unsupported {local}]"
        if tag in {w("pPr"), w("rPr"), w("sdtPr"), w("sdtEndPr")}:
            return ""
        return "".join(self.inline(child) for child in node)

    def record_field(self, instruction):
        value = normalize(instruction)
        self.fields.append({"command": value.split(" ", 1)[0],
                            "sha256": hashlib.sha256(value.encode()).hexdigest(), "part": self.part})

    def properties(self, paragraph):
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

    def paragraph(self, node):
        text = self.inline(node)
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
                self.warn("List boundaries and indentation are retained; automatic numbering and bullet glyphs are not reconstructed.")
                ilvl = num.find("w:ilvl", NS)
                depth = attr(ilvl, "val", "0") if ilvl is not None else "0"
                indent = "  " * min(20, int(depth)) if depth.isdigit() else ""
                return indent + "- " + ("\n" + indent + "  ").join(lines)
        return "\n".join(lines)

    def blocks(self, node):
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
            elif child.tag in {w("altChunk"), w("drawing"), w("pict"), w("object")}:
                value = self.inline(child)
            else:
                if child.tag in {w("ins"), w("moveTo")}:
                    self.review.append({"kind": child.tag.rsplit("}", 1)[-1], "author": attr(child, "author"),
                                        "text": self.raw_text(child), "part": self.part})
                value = self.blocks(child)
            if value:
                parts.append(value)
        return "\n\n".join(parts)

    def table_items(self, node, tag):
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

    def table(self, node):
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

    def audit(self, tree):
        # This small allowlist is intentionally conservative: unknown constructs warn.
        known = set("document body p r t delText tab br cr noBreakHyphen softHyphen pPr rPr sectPr tbl tr tc tblPr trPr tcPr tblGrid gridCol sdt sdtContent sdtPr sdtEndPr customXml smartTag ins del moveFrom moveTo hyperlink bookmarkStart bookmarkEnd proofErr permStart permEnd commentRangeStart commentRangeEnd commentReference footnoteReference endnoteReference footnotes footnote endnotes endnote footnoteRef endnoteRef comments comment hdr ftr fldChar instrText fldSimple lastRenderedPageBreak moveFromRangeStart moveFromRangeEnd moveToRangeStart moveToRangeEnd".split())
        def visit(node):
            local = node.tag.rsplit("}", 1)[-1]
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

    def render_part(self, part, tree):
        self.part, self.rels = part, self.relationships(part)
        self.audit(tree)
        value = self.blocks(tree)
        if self.field_stack:
            self.warn("Unclosed field codes found; saved display text may be incomplete.")
            while self.field_stack:
                self.record_field("".join(self.field_stack.pop()))
        return value

    def extract(self):
        if "word/document.xml" not in self.names:
            raise DocumentError("missing word/document.xml; expected a standard DOCX package")
        main = self.xml("word/document.xml")
        if main.tag != w("document") or main.find("w:body", NS) is None:
            raise DocumentError("unsupported Word XML dialect or missing document body (Strict OOXML is not supported)")
        main_rels = self.relationships("word/document.xml")
        supported = {"word/document.xml"}
        for kind, target, external in main_rels.values():
            if kind == "styles" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                self.styles = {attr(s, "styleId"): s for s in self.xml(target)}
                supported.add(target)
        prose = [self.render_part("word/document.xml", main.find("w:body", NS))]
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
            if name in {"word/settings.xml", "word/webSettings.xml", "word/fontTable.xml", "word/numbering.xml"} or name.startswith("word/theme/"):
                continue
            self.warn(f"Unrendered package part: {name}; changes to its content are not compared.")
        if self.features["footnote references"] and not any(k == "footnotes" for k, _ in stories):
            self.warn("Footnote references have no related footnotes part.")
        if self.features["endnote references"] and not any(k == "endnotes" for k, _ in stories):
            self.warn("Endnote references have no related endnotes part.")
        metadata = {"schema": 1, "view": "accept-insertions-omit-deletions",
                    "coverage": "partial-prose", "warnings": sorted(self.warnings),
                    "review": self.review, "hyperlinks": self.links, "fields": self.fields}
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
