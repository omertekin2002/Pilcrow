"""Read WordprocessingML without modifying the package or requiring Word.

This is a prose view, not a Word renderer. All omissions are reported in metadata.
"""

from __future__ import annotations

import hashlib
import io
import json
import posixpath
import re
from bisect import bisect_right
from collections import Counter
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO, TypeAlias
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile

from defusedxml import ElementTree as ET

from .prose import normalize, sentences

if TYPE_CHECKING:
    # Only the node type; all parsing goes through defusedxml.
    from xml.etree.ElementTree import Element

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# Office Math (OMML) equations.
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# Word 2010 paragraph IDs and the Word 2012 comments-extended part (reply threads, resolved state).
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS = {"w": W}
# Image references: DrawingML blips (r:embed internal, r:link external), the SVG copy
# Word stores in a blip's extension list, and legacy VML image data.
BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
SVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"
VML = "urn:schemas-microsoft-com:vml"
IMAGEDATA = f"{{{VML}}}imagedata"
O_RELID = "{urn:schemas-microsoft-com:office:office}relid"
# Where an image's alt text lives: the drawing's wp:docPr, else the picture's own pic:cNvPr,
# and a VML shape's alt attribute. A group's wp:docPr describes the group, not its members.
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
DRAWINGS = frozenset({f"{{{WP}}}inline", f"{{{WP}}}anchor"})
PICTURE = f"{{{PIC}}}pic"
GROUPS = frozenset({"{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}wgp",
                    "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}grpSp",
                    f"{{{VML}}}group"})
# Legacy VML WordArt: the text a v:shape draws along its path.
TEXTPATH = f"{{{VML}}}textpath"
# Objects whose content lives in a package part: a DrawingML chart, a SmartArt diagram's
# data part (r:dm; its layout, quick-style and colour parts are formatting), and an OLE
# object. w:altChunk is the fourth, an imported chunk.
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CHART = f"{{{CHART_NS}}}chart"
DIAGRAM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"
OLE_OBJECT = "{urn:schemas-microsoft-com:office:office}OLEObject"
# Shape content that a rendered text box still leaves out of the prose.
UNRENDERED_SHAPE_PARTS = frozenset({"pic", "chart", "relIds", "imagedata"})
TEXT_BOX_NOTE = ("Text box text is shown in [Text box] blocks after its anchoring paragraph; "
                 "position, anchoring, wrapping and shape formatting are not compared.")
# A SmartArt diagram's data model (the r:dm part): its points, and the parOf connections
# that make its outline. Only node and assistant points hold prose.
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
DIAGRAM_NOTE = ("SmartArt diagram text is shown as an outline in [Diagram] blocks after its anchoring paragraph; "
                "layout, shapes and formatting are not shown, and the diagram's data part is fingerprinted in objects.")
DIAGRAM_UNREACHED_NOTE = ("Some SmartArt diagram text is not connected to the diagram's outline; "
                          "it is shown at the top level after the outline.")
# Outline levels deeper than this are indented as this level, as list items are.
MAX_DIAGRAM_INDENT = 20
# A chart part's labels (c:chartSpace/c:chart): its title, axis titles, series names and string
# categories, from rich text or the cached values Word displays. Data values are not prose.
CHART_NOTE = ("Chart titles, axis titles, series names and text categories are shown in [Chart] blocks after "
              "their anchoring paragraph, from the chart's cached values; data values, numeric categories, "
              "data labels and formatting are not shown, and the chart part is fingerprinted in objects.")
CHART_AXES = frozenset(f"{{{CHART_NS}}}{name}" for name in ("catAx", "valAx", "dateAx", "serAx"))
# Category lines shown per chart; the rest are counted in one "(N more)" line.
MAX_CHART_CATEGORIES = 200
LIST_NOTE = ("List boundaries and indentation are retained; bullet and number formats are recorded in "
             "metadata, but counter values and bullet glyphs are not reconstructed.")
UNKNOWN_LIST_NOTE = ("Some list numbering formats could not be resolved (no numbering part, an unknown "
                     "list or level, or style-linked numbering); they are recorded as unknown.")
IMAGE_NOTE = "Images are compared by content hash only; size, position, cropping, and alt text are not compared."
# Replaces IMAGE_NOTE only when an image has alt text, so other documents keep their companions.
IMAGE_ALT_NOTE = "Images are compared by content hash and alt text; size, position, and cropping are not compared."
MISSING_IMAGE_NOTE = ("Some image relationships are missing or point to missing parts; "
                      "those images are recorded without a hash.")
MISSING_OBJECT_NOTE = ("Some chart, diagram, embedded object or imported chunk relationships are missing "
                       "or point to missing parts; those objects are recorded without a hash.")
COMMENT_NOTE = ("Comment text, authors, anchored text, replies, and resolved state are recorded; "
                "dates and rich formatting are not compared.")
UNANCHORED_COMMENT_NOTE = ("Some comments are not anchored in any rendered part; "
                           "their anchor is recorded as null.")
UNCLOSED_COMMENT_NOTE = ("Some comment ranges are missing a start or an end; their anchors run "
                         "from the start or to the end of the story part.")
FORM_FIELD_NOTE = ("Legacy form field checkbox and dropdown state is recorded in field metadata; "
                   "text form field settings are not compared.")
MIB = 1024 * 1024
# Package safety limits. Parsed parts (XML and .rels) are read whole into memory, so they
# have tight caps; media is hashed in fixed-size chunks, so only the package-wide total
# applies to it, which bounds decompression time. Word's own documented ceiling for a
# document with graphics is 512 MB.
MAX_ENTRIES = 4096
MAX_PACKAGE_BYTES = 1024 * MIB
MAX_PARSED_PART_BYTES = 16 * MIB
MAX_PARSED_TOTAL_BYTES = 64 * MIB
HASH_CHUNK_BYTES = 1 * MIB
# The only ZIP methods OPC allows, and the only ones zipfile decompresses with bounded
# output: it inflates a whole BZIP2, LZMA or Zstandard chunk at once, so a small
# compressed part could expand to gigabytes in memory before any limit applies.
BOUNDED_COMPRESSION = frozenset({ZIP_STORED, ZIP_DEFLATED})
# Relationship Id -> (type suffix, target, is external).
Relationships: TypeAlias = dict[str | None, tuple[str, str, bool]]


def w(name: str) -> str:
    return f"{{{W}}}{name}"


COMMENT_MARKERS = frozenset({w("commentRangeStart"), w("commentRangeEnd"), w("commentReference")})
# Legacy form field data (w:ffData) by element: the children whose meaning is recorded, or that
# only set behaviour (macros, help text, enabled state). Anything else, such as w:textInput and
# its children, keeps the unsupported construct warning.
FORM_DATA = {
    w("ffData"): frozenset(map(w, ("name", "enabled", "calcOnExit", "entryMacro", "exitMacro", "helpText",
                                   "statusText", "checkBox", "ddList"))),
    w("checkBox"): frozenset(map(w, ("size", "sizeAuto", "default", "checked"))),
    w("ddList"): frozenset(map(w, ("result", "default", "listEntry"))),
}
EQUATIONS = frozenset({f"{{{M}}}oMathPara", f"{{{M}}}oMath"})
# Markup an equation's fingerprint leaves out: deleted content, formatting, and markers whose
# IDs Word renumbers on save. Proofing, bookmark, permission and move range markers carry no
# content; comment markers are anchors, not equation content.
EQUATION_DROPPED = frozenset({
    w("del"), w("moveFrom"), w("rPr"), f"{{{M}}}ctrlPr", w("proofErr"), w("bookmarkStart"), w("bookmarkEnd"),
    w("permStart"), w("permEnd"), w("moveFromRangeStart"), w("moveFromRangeEnd"), w("moveToRangeStart"),
    w("moveToRangeEnd"), *COMMENT_MARKERS,
})
EQUATION_UNWRAPPED = frozenset({w("ins"), w("moveTo")})
# Attributes that change on every save without changing the equation.
EQUATION_IGNORED_ATTRIBUTES = frozenset({f"{{{W14}}}paraId", f"{{{W14}}}textId"})


def om(name: str) -> str:
    return f"{{{M}}}{name}"


MATH_NOTE = ("Equations are shown as linear [Math: …] text; layout, fonts and alignment are not compared, "
             "and equation fingerprints are in objects.")
SYMBOL_WARNING = "Unsupported sym: content is not rendered in prose; inspect the DOCX in Word."
# An equation's text is held between these marks inside a paragraph until its sentences are split.
# XML 1.0 text cannot contain either character, so document text never holds them.
MATH_START, MATH_END = "\x01", "\x02"
MATH_SPAN = re.compile("\x01([^\x02]*)\x02")
MATH_HELD = re.compile("\x01([0-9]+)\x02")
# Math elements whose linear text is their content's text, one after another.
MATH_CONTAINERS = frozenset(map(om, "oMathPara oMath e num den sub sup deg lim fName mr r".split()))
# Math structures with a linear form of their own (m:m, a matrix, is read row by row).
MATH_STRUCTURES = frozenset(map(om, "f sSup sSub sSubSup sPre rad d nary func limLow limUpp acc bar groupChr "
                                    "box borderBox phant eqArr".split()))
# A single atom needs no parentheses as an argument: only letters, a number, or one character.
MATH_ATOM = re.compile(r"[^\W\d_]+|[0-9]+(?:\.[0-9]+)?|.", re.DOTALL)


def plain(text: str) -> str:
    """``text`` without the marks around equation text."""
    return text.replace(MATH_START, "").replace(MATH_END, "")


def math_sentences(text: str) -> list[str]:
    """Sentence lines of ``text`` that never split inside an equation's [Math: …] text.

    Each marked equation is swapped for a short token while sentences are split, then put
    back without its marks, so splitting stays linear in the paragraph's length.
    """
    held: list[str] = []

    def hold(match: re.Match[str]) -> str:
        held.append(match[1])
        return f"{MATH_START}{len(held) - 1}{MATH_END}"

    lines = sentences(MATH_SPAN.sub(hold, text))
    if not held:
        return lines
    return [MATH_HELD.sub(lambda match: held[int(match[1])], line) for line in lines]


def math_children(node: Element) -> Iterator[Element]:
    """A math element's children in the accepted view, as ``equation_form`` reads them, without
    property elements (``m:*Pr``): those are read only by the structure they belong to."""
    for child in node:
        if child.tag in EQUATION_UNWRAPPED:
            yield from math_children(child)
        elif child.tag not in EQUATION_DROPPED and not (child.tag.startswith(f"{{{M}}}") and child.tag.endswith("Pr")):
            yield child


# A structure's arguments by local name (e, num, sub, …): each element and its normalised text.
MathArgs: TypeAlias = dict[str, list[tuple["Element", str]]]


def math_arg(args: MathArgs, name: str) -> str:
    """The first ``name`` argument's text, "" when there is none."""
    return args[name][0][1] if name in args else ""


def math_group(args: MathArgs, name: str) -> str:
    """An argument as an operand: parenthesised unless it is empty, a single atom, or one visible
    delimiter (``(a+b)`` is not wrapped again)."""
    if name not in args:
        return ""
    child, text = args[name][0]
    if not text or MATH_ATOM.fullmatch(text):
        return text
    content = list(math_children(child))
    if len(content) == 1 and content[0].tag == om("d") and all(delimiters(content[0])[::2]):
        return text
    return f"({text})"


def math_scripts(node: Element, args: MathArgs, hide: bool = False) -> str:
    """``_sub`` and ``^sup`` of a structure, each left out when empty (or hidden, for m:nary)."""
    sub = "" if hide and math_flag(node, "subHide") else math_group(args, "sub")
    sup = "" if hide and math_flag(node, "supHide") else math_group(args, "sup")
    return ("_" + sub if sub else "") + ("^" + sup if sup else "")


def delimiters(node: Element) -> tuple[str, str, str]:
    """An m:d's beginning, separator and ending characters."""
    return math_char(node, "begChr", "("), math_char(node, "sepChr", "|"), math_char(node, "endChr", ")")


def math_property(node: Element, name: str) -> Element | None:
    """The ``name`` property of a math structure, from its own m:<structure>Pr element."""
    props = node.find(node.tag + "Pr")
    return None if props is None else props.find(om(name))


def math_char(node: Element, name: str, default: str) -> str:
    """A character property (m:chr, m:begChr, …); an empty m:val means no character."""
    prop = math_property(node, name)
    value = None if prop is None else prop.get(om("val"))
    return default if value is None else value


def math_flag(node: Element, name: str) -> bool:
    """An on/off property (m:degHide, m:subHide, …); present without m:val means on."""
    prop = math_property(node, name)
    return prop is not None and on_off(prop.get(om("val"), "1"))


def equation_form(node: Element) -> list[Any]:
    """The accepted view of an equation as nested lists: tag, sorted attributes, text, children.

    Tracked deletions are dropped and insertions unwrapped; run formatting, control
    character formatting, revision IDs and markers are left out, so a formatting-only
    re-save keeps the same form. m:t text is kept exactly; other text is XML layout.
    """
    def children(parent: Element) -> Iterator[Element]:
        for child in parent:
            if child.tag in EQUATION_UNWRAPPED:
                yield from children(child)
            elif child.tag not in EQUATION_DROPPED:
                yield child
    attributes = sorted([key, value] for key, value in node.attrib.items()
                        if key not in EQUATION_IGNORED_ATTRIBUTES
                        and not key.rsplit("}", 1)[-1].startswith("rsid"))
    value = node.text or ""
    text = value if node.tag.rsplit("}", 1)[-1] == "t" else value.strip()
    return [node.tag, attributes, text, [equation_form(child) for child in children(node)]]


def equation_hash(node: Element) -> str:
    form = json.dumps(equation_form(node), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(form.encode("utf-8")).hexdigest()


def ch(name: str) -> str:
    return f"{{{CHART_NS}}}{name}"


def drawing_text(paragraphs: Iterable[Element]) -> str:
    """DrawingML text: each a:p's runs, fields and breaks (as a space), normalised, joined by a space."""
    found = []
    for para in paragraphs:
        parts: list[str] = []
        for child in para:
            if child.tag in {f"{{{A}}}r", f"{{{A}}}fld"}:
                parts.extend(t.text or "" for t in child.findall(f"{{{A}}}t"))
            elif child.tag == f"{{{A}}}br":
                parts.append(" ")
        found.append(normalize("".join(parts)))
    return " ".join(p for p in found if p)


def chart_points(cache: Element) -> dict[int, str]:
    """A string cache's (or literal's) point texts by index, normalised; the first point per index wins."""
    points: dict[int, str] = {}
    for point in cache.findall(ch("pt")):
        index = decimal(point.get("idx", ""))
        value = point.find(ch("v"))
        if index is not None:
            points.setdefault(index, normalize("" if value is None else value.text or ""))
    return points


def chart_text(tx: Element | None) -> str:
    """A chart label's text (c:tx): rich text, else a cached string reference, else a literal value."""
    if tx is None:
        return ""
    rich = tx.find(ch("rich"))
    if rich is not None:
        return drawing_text(rich.findall(f"{{{A}}}p"))
    cache = tx.find(f"{ch('strRef')}/{ch('strCache')}")
    if cache is not None:
        points = chart_points(cache)
        return normalize(" ".join(points[i] for i in sorted(points)))
    value = tx.find(ch("v"))
    return "" if value is None else normalize(value.text or "")


def chart_categories(cat: Element | None) -> list[str]:
    """A series's string category labels (c:cat) in index order; numeric and date categories have none.

    A multi-level category joins its levels from the outermost to the innermost with " / ";
    an outer label covers every index from its own up to the next outer label.
    """
    if cat is None:
        return []
    for path in (f"{ch('strRef')}/{ch('strCache')}", ch("strLit")):
        cache = cat.find(path)
        if cache is not None:
            points = chart_points(cache)
            return [points[i] for i in sorted(points) if points[i]]
    cache = cat.find(f"{ch('multiLvlStrRef')}/{ch('multiLvlStrCache')}")
    if cache is None:
        return []
    # c:lvl elements run from the innermost level outwards.
    levels = [chart_points(level) for level in cache.findall(ch("lvl"))]
    starts = [sorted(level) for level in levels]
    labels = []
    for index in sorted({i for level in levels for i in level}):
        parts = []
        for depth in reversed(range(len(levels))):
            if depth == 0:
                text = levels[0].get(index, "")
            else:
                at = bisect_right(starts[depth], index)
                text = levels[depth][starts[depth][at - 1]] if at else ""
            if text:
                parts.append(text)
        if parts:
            labels.append(" / ".join(parts))
    return labels


def size_label(size: int) -> str:
    """A byte limit in the unit it is written in: "16 MiB", "1 GiB"."""
    return f"{size // 1024 // MIB} GiB" if size % (1024 * MIB) == 0 else f"{size // MIB} MiB"


def attr(node: Element, name: str, default: str = "") -> str:
    return node.get(w(name), default)


def decimal(value: str) -> int | None:
    """Read an ST_DecimalNumber (xsd:integer): a sign and leading zeroes are valid.

    At most ten significant digits, so a huge value cannot exceed int()'s digit limit.
    """
    match = re.fullmatch(r"([+-]?)0*([0-9]{1,10})", value.strip(" \t\r\n"))
    if match is None:
        return None
    return -int(match[2]) if match[1] == "-" else int(match[2])


def para_id(value: str) -> str:
    """One spelling per w14:paraId (ST_LongHexNumber), so "00ab" and "00AB" match."""
    return value.strip(" \t\r\n").upper()


def on_off(value: str) -> bool:
    """Read an ST_OnOff value; anything other than a true spelling is false."""
    return value.strip(" \t\r\n") in {"1", "true", "on"}


def number_id(value: str) -> str:
    """One spelling per numbering ID, so "1", "+1" and "01" name the same definition."""
    number = decimal(value)
    return value if number is None else str(number)


# An image's (alt text, alt title), each whitespace-normalised and "" when absent.
AltText: TypeAlias = tuple[str, str]


def alt_text(node: Element, inherited: AltText) -> AltText:
    """The alt text images inside ``node`` have, given what its ancestors set.

    A drawing's wp:docPr comes first, then the picture's own pic:cNvPr; a VML shape's
    alt attribute; group members only have their own.
    """
    if node.tag in DRAWINGS:
        props = node.find(f"{{{WP}}}docPr")
        return ("", "") if props is None else (normalize(props.get("descr", "")), normalize(props.get("title", "")))
    if node.tag in GROUPS:
        return "", ""
    if node.tag == PICTURE:
        props = node.find(f"{{{PIC}}}nvPicPr/{{{PIC}}}cNvPr")
        if props is None:
            return inherited
        return (inherited[0] or normalize(props.get("descr", "")), inherited[1] or normalize(props.get("title", "")))
    if node.tag.startswith(f"{{{VML}}}") and "alt" in node.attrib:
        return normalize(node.get("alt", "")), ""
    return inherited


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
        self.review: list[dict[str, str | int | bool | None]] = []
        self.links: list[dict[str, str]] = []
        self.fields: list[dict[str, Any]] = []
        self.styles: dict[str, Element] = {}
        # w:num by numId and w:abstractNum by abstractNumId; the first definition wins.
        self.nums: dict[str, Element] = {}
        self.abstract_nums: dict[str, Element] = {}
        # Raw numId -> list entry; insertion order is first use, so ordinals ignore raw IDs.
        self.lists: dict[str, dict[str, Any]] = {}
        # Which list each list item belongs to, run-length encoded in traversal order.
        self.list_runs: list[dict[str, Any]] = []
        # Image references in traversal order, and hashed part (media, or the part a chart,
        # diagram, embedded object or imported chunk points to) -> (sha256, size).
        self.images: list[dict[str, Any]] = []
        self.media: dict[str, tuple[str, int]] = {}
        # Equation, symbol and object part fingerprints in traversal order.
        self.objects: list[dict[str, str | None]] = []
        # Inside an unselected mc:AlternateContent copy whose prose is still read.
        self.unselected = 0
        # Open fields: instruction text so far, and the form state of a legacy form field.
        self.field_stack: list[tuple[list[str], dict[str, Any] | None]] = []
        self.part = "word/document.xml"
        self.rels: Relationships = {}
        self.features: Counter[str] = Counter()
        # Text boxes found while reading a paragraph, rendered as blocks after it, and the
        # [Chart] and [Diagram] blocks of charts and SmartArt diagrams, already rendered.
        self.anchored: list[list[Element | str]] = []
        # Diagram data part -> its outline lines, or None when it has no text-bearing points.
        self.diagram_outlines: dict[str, list[str] | None] = {}
        # Chart part -> its label lines; empty when it has none.
        self.chart_labels: dict[str, list[str]] = {}
        # The current story part's rendered text in fragments, in the accepted view, so a
        # comment range is the fragments between the offsets of its start and end markers.
        self.story: list[str] = []
        # Open comment range ID -> fragment offset of its start, for the current part.
        self.open_ranges: dict[str, int] = {}
        # Comment ID -> (anchored text, part holding the range start); the first range wins.
        self.comment_ranges: dict[str, tuple[str, str]] = {}
        # Comment ID -> part of its first w:commentReference, for comments without a range.
        self.comment_refs: dict[str, str] = {}
        # Uncompressed bytes of every part parsed so far, bounded by MAX_PARSED_TOTAL_BYTES.
        self.parsed_bytes = 0
        # Story part -> its relationships, parsed once however many notes render from it.
        self.part_rels: dict[str, Relationships] = {}

    def xml(self, name: str) -> Element:
        # Reading at most the declared size bounds memory, so check that size first. Reading
        # exactly that many bytes, not to the end, also bounds each inflate step: ZipFile.read
        # inflates a whole deflate stream in one step, so a part declaring a small size could
        # expand to gigabytes before being cut to its size and failing its CRC check.
        size = self.package.getinfo(name).file_size if name in self.names else 0
        if size > MAX_PARSED_PART_BYTES:
            raise DocumentError(f"{name} is {size} bytes uncompressed; parsed parts are limited to "
                                f"{size_label(MAX_PARSED_PART_BYTES)} each")
        self.parsed_bytes += size
        if self.parsed_bytes > MAX_PARSED_TOTAL_BYTES:
            raise DocumentError(f"cannot parse {name}: parsed parts exceed the "
                                f"{size_label(MAX_PARSED_TOTAL_BYTES)} total safety limit")
        try:
            with self.package.open(name) as stream:
                data = stream.read(size)
            return ET.fromstring(data)
        except Exception as exc:
            raise DocumentError(f"cannot parse {name}: {exc}") from exc

    def hash_part(self, name: str) -> tuple[str, int]:
        """SHA-256 and size of a part, read in fixed-size chunks so memory stays bounded."""
        digest = hashlib.sha256()
        size = 0
        try:
            with self.package.open(name) as stream:
                while chunk := stream.read(HASH_CHUNK_BYTES):
                    digest.update(chunk)
                    size += len(chunk)
        except Exception as exc:  # e.g. zlib.error for corrupt compressed data
            raise DocumentError(f"cannot read {name}: {exc}") from exc
        return digest.hexdigest(), size

    def relationships(self, part: str) -> Relationships:
        if part in self.part_rels:
            return self.part_rels[part]
        path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        result: Relationships = self.part_rels.setdefault(part, {})
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

    def last_paragraph_id(self, comment: Element) -> str:
        """The w14:paraId of a comment's last paragraph, which w15:commentEx refers to.

        Paragraphs inside text boxes are not the comment's own paragraphs.
        """
        boxed = {p for box in comment.iter(w("txbxContent")) for p in box.iter(w("p"))}
        own = [p for p in comment.iter(w("p")) if p not in boxed]
        return para_id(own[-1].get(f"{{{W14}}}paraId", "")) if own else ""

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
        """The one copy of an mc:AlternateContent text box or equation to render, or None.

        Word stores a modern text box (DrawingML in mc:Choice, VML in mc:Fallback) and a
        compatibility-wrapped equation twice; text boxes take precedence when present.
        """
        def pick(tags: frozenset[str]) -> Element | None:
            def has(node: Element) -> bool:
                return any(d.tag in tags for d in node.iter())
            choice = next((c for c in alternate if c.tag == f"{{{MC}}}Choice" and has(c)), None)
            if choice is not None:
                return choice
            fallback = alternate.find(f"{{{MC}}}Fallback")
            return fallback if fallback is not None and has(fallback) else None
        found = pick(frozenset({w("txbxContent")}))
        return found if found is not None else pick(EQUATIONS)

    def text_boxes(self, node: Element) -> list[Element]:
        """Outermost rendered w:txbxContent elements in document order; nested ones render inside them."""
        if node.tag == w("txbxContent"):
            return [node]
        if node.tag == f"{{{MC}}}AlternateContent":
            branch = self.branch(node)
            return self.text_boxes(branch) if branch is not None else []
        return [box for child in node for box in self.text_boxes(child)]

    def unrendered_shape_parts(self, node: Element, rendered: frozenset[Element] = frozenset()) -> bool:
        """Whether a shape holding text boxes or ``rendered`` charts and diagrams also holds
        pictures, charts or diagrams outside them."""
        if node.tag == w("txbxContent") or node in rendered:
            return False
        if node.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
            return self.unrendered_shape_parts(branch, rendered)
        # Alternatives without a text box (e.g. a grouped picture) are checked in full.
        return (node.tag.rsplit("}", 1)[-1] in UNRENDERED_SHAPE_PARTS
                or any(self.unrendered_shape_parts(child, rendered) for child in node))

    def image_refs(self, node: Element, alt: AltText = ("", "")) -> Iterator[tuple[list[str | None], AltText]]:
        """Relationship IDs and alt text of each image in ``node``, in document order.

        Text box content is skipped (it is read when the box renders), and only one
        copy of an mc:AlternateContent is read: the text box branch, else the first
        mc:Choice, else mc:Fallback.
        """
        if node.tag == w("txbxContent"):
            return
        if node.tag == f"{{{MC}}}AlternateContent":
            branch = self.image_branch(node)
            if branch is not None:
                yield from self.image_refs(branch, alt)
            return
        alt = alt_text(node, alt)
        if node.tag in {BLIP, SVG_BLIP}:
            yield [node.get(f"{{{R}}}embed"), node.get(f"{{{R}}}link")], alt
        elif node.tag == IMAGEDATA:
            yield [node.get(f"{{{R}}}id") or node.get(O_RELID)], alt
        for child in node:
            yield from self.image_refs(child, alt)

    def image_branch(self, alternate: Element) -> Element | None:
        """The one copy of an mc:AlternateContent whose images are recorded."""
        branch = self.branch(alternate)
        if branch is None:
            branch = next((c for c in alternate if c.tag == f"{{{MC}}}Choice"), alternate.find(f"{{{MC}}}Fallback"))
        return branch

    def record_images(self, node: Element) -> None:
        """Record each image in ``node`` by content hash (internal) or target (external), and its alt text."""
        if self.unselected:
            return
        for ids, (description, title) in self.image_refs(node):
            entry: dict[str, Any] = {"part": self.part}
            for rid in filter(None, ids):
                rel = self.rels.get(rid)
                if rel is not None and rel[2]:
                    entry["target"] = rel[1]  # never fetched, like hyperlinks
                elif rel is not None and rel[1] in self.names:
                    if rel[1] not in self.media:
                        self.media[rel[1]] = self.hash_part(rel[1])
                    entry["sha256"], entry["bytes"] = self.media[rel[1]]
                else:
                    self.warn(MISSING_IMAGE_NOTE)
                    entry.setdefault("sha256", None)
            if len(entry) == 1:  # no relationship ID at all
                self.warn(MISSING_IMAGE_NOTE)
                entry["sha256"] = None
            # Only when present, so documents without alt text keep their companions.
            if description:
                entry["alt"] = description
            if title:
                entry["alt_title"] = title
            self.images.append(entry)
            self.warn(IMAGE_NOTE)

    def object_nodes(self, node: Element) -> Iterator[Element]:
        """Outermost equations, w:sym, charts, diagrams, OLE objects, alt chunks and VML WordArt in ``node``, in document order.

        Like ``image_refs``: text box content is read when the box renders, only one
        copy of an mc:AlternateContent is read, and tracked deletions are skipped.
        """
        if node.tag in EQUATIONS or node.tag in {w("sym"), w("altChunk"), CHART, DIAGRAM, OLE_OBJECT}:
            yield node
            return
        if node.tag == TEXTPATH and normalize(node.get("string", "")):
            yield node  # WordArt; a v:textpath without text (e.g. in a v:shapetype) is only a style
            return
        if node.tag in {w("txbxContent"), w("del"), w("moveFrom")}:
            return
        if node.tag == f"{{{MC}}}AlternateContent":
            branch = self.image_branch(node)
            if branch is not None:
                yield from self.object_nodes(branch)
            return
        for child in node:
            yield from self.object_nodes(child)

    def record_objects(self, node: Element) -> None:
        """Fingerprint each object in ``node``: an equation by its accepted-view hash, a symbol by
        font and character, a chart, diagram, embedded object or imported chunk by its part, and
        VML WordArt by its text."""
        if self.unselected:
            return
        for found in self.object_nodes(node):
            if found.tag == w("sym"):
                self.objects.append({"kind": "symbol", "part": self.part, "font": attr(found, "font"),
                                     "char": attr(found, "char").upper()})
            elif found.tag == CHART:
                self.objects.append(self.object_part("chart", found.get(f"{{{R}}}id")))
            elif found.tag == DIAGRAM:
                self.objects.append(self.object_part("diagram", found.get(f"{{{R}}}dm")))
            elif found.tag == OLE_OBJECT:
                entry = self.object_part("embedded object", found.get(f"{{{R}}}id"))
                entry["prog_id"] = found.get("ProgID", "")
                self.objects.append(entry)
            elif found.tag == w("altChunk"):
                self.objects.append(self.object_part("imported chunk", found.get(f"{{{R}}}id")))
            elif found.tag == TEXTPATH:
                self.objects.append({"kind": "WordArt", "part": self.part, "text": normalize(found.get("string", ""))})
            else:
                self.objects.append({"kind": "equation", "part": self.part, "sha256": equation_hash(found)})

    def object_part(self, kind: str, rid: str | None) -> dict[str, str | None]:
        """An object entry with the hash of the part ``rid`` points to, or its target if external.

        Part names are not recorded, so renaming a part with the same content is no change.
        """
        entry: dict[str, str | None] = {"kind": kind, "part": self.part}
        rel = self.rels.get(rid) if rid else None
        if rel is not None and rel[2]:
            entry["target"] = rel[1]  # never fetched, like linked images
        elif rel is not None and rel[1] in self.names:
            if rel[1] not in self.media:
                self.media[rel[1]] = self.hash_part(rel[1])
            entry["sha256"] = self.media[rel[1]][0]
        else:
            self.warn(MISSING_OBJECT_NOTE)
            entry["sha256"] = None
        return entry

    def equation_text(self, node: Element) -> str:
        """An equation as linear text: ``[Math: …]`` per m:oMath, separated by a space.

        It reads the accepted view, like the fingerprint: tracked deletions are left out,
        insertions kept, and formatting adds no text.
        """
        self.warn(MATH_NOTE)
        if node.tag == om("oMathPara"):
            maths = [normalize(self.math(child)) for child in math_children(node)] or [""]
        else:
            maths = [normalize(self.math(node))]
        return " ".join(f"[Math: {text or '(empty)'}]" for text in maths)

    def math(self, node: Element) -> str:
        """The linear text of one math element; arguments of structures are whitespace-normalised.

        Each level of nesting recurses straight into its children, so an equation can nest
        about as deeply as its fingerprint allows before failing.
        """
        tag = node.tag
        if tag in {om("t"), w("t")}:
            return re.sub(r"\s", " ", node.text or "")
        if tag in {w("tab"), w("br"), w("cr")}:
            return " "
        if tag == w("sym"):
            self.warn(SYMBOL_WARNING)
            return "[Unsupported sym]"
        if tag == f"{{{MC}}}AlternateContent":
            branch = self.image_branch(node)
            return "" if branch is None else "".join([self.math(child) for child in math_children(branch)])
        if tag == om("m"):
            rows = [", ".join([normalize(self.math(cell)) for cell in math_children(row) if cell.tag == om("e")])
                    for row in math_children(node) if row.tag == om("mr")]
            return "[" + "; ".join(rows) + "]"
        children = [(child, self.math(child)) for child in math_children(node)]
        if tag not in MATH_STRUCTURES:
            if tag.startswith(f"{{{M}}}") and tag not in MATH_CONTAINERS:
                self.warn(f"Unsupported math construct {tag.rsplit('}', 1)[-1]}; rendered as its text, structure not shown.")
            return "".join(text for _, text in children)
        args: MathArgs = {}
        for child, text in children:
            args.setdefault(child.tag.rsplit("}", 1)[-1], []).append((child, normalize(text)))
        if tag == om("f"):
            bar = "¦" if math_char(node, "type", "bar") == "noBar" else "/"
            return math_group(args, "num") + bar + math_group(args, "den")
        if tag in {om("sSup"), om("sSub"), om("sSubSup")}:
            return math_group(args, "e") + math_scripts(node, args)
        if tag == om("sPre"):
            return " ".join(part for part in (math_scripts(node, args), math_group(args, "e")) if part)
        if tag == om("rad"):
            degree = "" if math_flag(node, "degHide") else math_arg(args, "deg")
            return f"√({degree}&{math_arg(args, 'e')})" if degree else f"√({math_arg(args, 'e')})"
        if tag == om("d"):
            begin, separator, end = delimiters(node)
            return begin + separator.join(text for _, text in args.get("e", [])) + end
        if tag == om("nary"):
            scripts = math_scripts(node, args, hide=True)
            operand = math_arg(args, "e")
            return math_char(node, "chr", "∫") + scripts + (" " if scripts and operand else "") + operand
        if tag == om("func"):
            # The space already sets the operand apart: "sin x", "lim_(n→∞) a_n".
            return " ".join(part for part in (math_arg(args, "fName"), math_arg(args, "e")) if part)
        if tag in {om("limLow"), om("limUpp")}:
            limit = math_group(args, "lim")
            mark = "_" if tag == om("limLow") else "^"
            return math_group(args, "e") + (mark + limit if limit else "")
        if tag == om("acc"):
            return math_group(args, "e") + math_char(node, "chr", "\u0302")
        if tag == om("bar"):
            return math_group(args, "e") + ("\u0305" if math_char(node, "pos", "bot") == "top" else "\u0332")
        if tag == om("groupChr"):
            return math_group(args, "e") + math_char(node, "chr", "\u23df")
        if tag == om("eqArr"):
            return "; ".join(text for _, text in args.get("e", []))
        return math_arg(args, "e")  # m:box, m:borderBox, m:phant

    def emit(self, text: str, anchored: str | None = None) -> str:
        """Return prose ``text`` and add it (or ``anchored``) to the story text comment ranges cover."""
        if not self.unselected:
            self.story.append(text if anchored is None else anchored)
        return text

    def comment_marker(self, node: Element) -> None:
        """Track a comment range start/end or reference by its ID; the ID itself is not recorded."""
        if self.unselected:
            return
        comment = number_id(attr(node, "id"))
        if node.tag == w("commentReference"):
            self.comment_refs.setdefault(comment, self.part)
        elif node.tag == w("commentRangeStart"):
            self.open_ranges.setdefault(comment, len(self.story))
        else:
            start = self.open_ranges.pop(comment, None)
            if start is None:
                self.warn(UNCLOSED_COMMENT_NOTE)
            self.close_range(comment, start or 0)

    def close_range(self, comment: str, start: int) -> None:
        self.comment_ranges.setdefault(comment, (normalize("".join(self.story[start:])), self.part))

    def close_open_ranges(self) -> None:
        """Close the ranges still open at the end of a story at that end, with a warning."""
        if self.open_ranges:
            self.warn(UNCLOSED_COMMENT_NOTE)
            for comment, start in self.open_ranges.items():
                self.close_range(comment, start)
            self.open_ranges = {}

    def skip_markers(self, node: Element) -> None:
        """Track the comment markers in content that is not rendered, such as tracked deletions."""
        for descendant in node.iter():
            if descendant.tag in COMMENT_MARKERS:
                self.comment_marker(descendant)

    def anchor(self, node: Element) -> tuple[str, list[str]]:
        """Read ``node`` inline and render the text boxes, charts and diagrams it anchors as separate blocks."""
        text, boxes = self.anchor_boxes(node)
        return text, [self.anchored_block(box) for box in boxes]

    def anchor_boxes(self, node: Element) -> tuple[str, list[Element | str]]:
        """Read ``node`` inline and return the text boxes it anchors, not yet rendered, and its
        rendered [Chart] and [Diagram] blocks."""
        self.anchored.append([])
        text = self.inline(node)
        return text, self.anchored.pop()

    def anchored_block(self, item: Element | str) -> str:
        return item if isinstance(item, str) else self.text_box(item)

    def diagram(self, node: Element) -> str | None:
        """A SmartArt diagram's [Diagram] block, or None when its data part is missing, external,
        or holds no text-bearing points, so the drawing keeps its placeholder."""
        rid = node.get(f"{{{R}}}dm")
        rel = self.rels.get(rid) if rid else None
        if rel is None or rel[2] or rel[1] not in self.names:
            return None
        if rel[1] not in self.diagram_outlines:
            self.diagram_outlines[rel[1]] = self.diagram_outline(self.xml(rel[1]))
        lines = self.diagram_outlines[rel[1]]
        if lines is None:
            return None
        self.features["diagrams"] += 1
        self.warn(DIAGRAM_NOTE)
        return "\n\n".join(("[Diagram]", "\n".join(lines) or "(empty)", "[/Diagram]"))

    def diagram_outline(self, model: Element) -> list[str] | None:
        """The outline lines of a dgm:dataModel, or None when it has no text-bearing points.

        The parOf connections are walked depth-first from the doc point, siblings by srcOrd and
        then connection order, each point at most once, without recursion. Text-bearing points
        the walk never reaches follow at the top level, in point-list order, with a warning.
        """
        if model.tag != f"{{{DGM}}}dataModel":
            return None
        points: dict[str, tuple[str, str]] = {}  # modelId -> (type, text); the first point wins
        for point in model.findall(f"{{{DGM}}}ptLst/{{{DGM}}}pt"):
            points.setdefault(point.get("modelId", ""), (point.get("type", "node"), self.diagram_text(point)))
        bearing = [key for key, (kind, text) in points.items() if kind == "node" or (kind == "asst" and text)]
        if not bearing:
            return None
        children: dict[str, list[tuple[int, int, str]]] = {}
        for order, connection in enumerate(model.findall(f"{{{DGM}}}cxnLst/{{{DGM}}}cxn")):
            if connection.get("type", "parOf") != "parOf":
                continue
            rank = decimal(connection.get("srcOrd", "0"))
            children.setdefault(connection.get("srcId", ""), []).append(
                (rank if rank is not None else 0, order, connection.get("destId", "")))
        root = next((key for key, (kind, _) in points.items() if kind == "doc"), None)
        lines: list[str] = []
        visited: set[str] = set() if root is None else {root}
        # (point, level) pairs still to visit, the next one last.
        stack: list[tuple[str, int]] = [] if root is None else [
            (child, 0) for *_, child in sorted(children.get(root, []), reverse=True)]
        while stack:
            key, depth = stack.pop()
            if key in visited or key not in points:
                continue
            visited.add(key)
            self.diagram_line(lines, points[key], depth)
            stack.extend((child, depth + 1) for *_, child in sorted(children.get(key, []), reverse=True))
        unreached = [key for key in bearing if key not in visited]
        if any(points[key][1] for key in unreached):
            self.warn(DIAGRAM_UNREACHED_NOTE)
        for key in unreached:
            self.diagram_line(lines, points[key], 0)
        return lines

    def diagram_line(self, lines: list[str], point: tuple[str, str], depth: int) -> None:
        """Add a point's outline line, split into sentences like a list item; empty points add none."""
        kind, text = point
        found = sentences(text) if kind in {"node", "asst"} else []
        if found:
            indent = "  " * min(MAX_DIAGRAM_INDENT, depth)
            lines.append(indent + "- " + ("\n" + indent + "  ").join(found))

    def diagram_text(self, point: Element) -> str:
        """A point's text: each dgm:t paragraph's runs, fields and breaks, normalised, joined by a space."""
        return drawing_text(point.findall(f"{{{DGM}}}t/{{{A}}}p"))

    def chart(self, node: Element) -> str | None:
        """A chart's [Chart] block, or None when its part is missing, external, or holds no
        labels, so the drawing keeps its placeholder."""
        rid = node.get(f"{{{R}}}id")
        rel = self.rels.get(rid) if rid else None
        if rel is None or rel[2] or rel[1] not in self.names:
            return None
        if rel[1] not in self.chart_labels:
            self.chart_labels[rel[1]] = self.chart_lines(self.xml(rel[1]))
        lines = self.chart_labels[rel[1]]
        if not lines:
            return None
        self.features["charts"] += 1
        self.warn(CHART_NOTE)
        return "\n\n".join(("[Chart]", "\n".join(lines), "[/Chart]"))

    def chart_lines(self, space: Element) -> list[str]:
        """A chart part's label lines: the title, axis titles in document order, then series
        names in document order with their string categories, at most MAX_CHART_CATEGORIES.

        Categories every series shares are listed once after the series; otherwise each
        series's categories follow its name.
        """
        chart = space.find(ch("chart")) if space.tag == ch("chartSpace") else None
        if chart is None:
            return []
        lines = []
        title, deleted = chart.find(ch("title")), chart.find(ch("autoTitleDeleted"))
        # CT_Boolean: a missing val means true. Word's automatic title is never inferred.
        removed = deleted is not None and on_off(deleted.get("val", "1"))
        if title is not None and not removed and (text := chart_text(title.find(ch("tx")))):
            lines.append(f"Title: {text}")
        plot = chart.find(ch("plotArea"))
        if plot is None:
            return lines
        for axis in plot:
            heading = axis.find(ch("title")) if axis.tag in CHART_AXES else None
            if heading is not None and (text := chart_text(heading.find(ch("tx")))):
                lines.append(f"Axis title: {text}")
        # Every chart-type group (c:barChart, c:lineChart, c:pieChart, ...) holds its series.
        series = [(chart_text(ser.find(ch("tx"))), chart_categories(ser.find(ch("cat"))))
                  for group in plot if group.tag.startswith(f"{{{CHART_NS}}}") and group.tag.endswith("Chart")
                  for ser in group.findall(ch("ser"))]
        lists = [categories for _, categories in series if categories]
        shared = all(categories == lists[0] for categories in lists)
        labels: list[tuple[str, str]] = []  # (kind, text)
        for name, categories in series:
            if name:
                labels.append(("Series", name))
            if not shared:
                labels.extend(("Category", category) for category in categories)
        if shared and lists:
            labels.extend(("Category", category) for category in lists[0])
        counted = 0
        hidden = sum(kind == "Category" for kind, _ in labels) - MAX_CHART_CATEGORIES
        for kind, text in labels:
            if kind == "Category":
                counted += 1
                if counted == MAX_CHART_CATEGORIES + 1:
                    lines.append(f"Category: … ({hidden} more)")
                if counted > MAX_CHART_CATEGORIES:
                    continue
            lines.append(f"{kind}: {text}")
        return lines

    def text_box(self, box: Element) -> str:
        self.features["text boxes"] += 1
        self.warn(TEXT_BOX_NOTE)
        # A text box is its own story: a range open around its anchor does not cover its
        # text, and a range inside it runs at most to the box's edges.
        outer = self.story, self.open_ranges
        self.story, self.open_ranges = [], {}
        try:
            value = self.blocks(box)
            self.close_open_ranges()
        finally:
            self.story, self.open_ranges = outer
        return "\n\n".join(("[Text box]", value or "(empty)", "[/Text box]"))

    def inline(self, node: Element) -> str:
        tag = node.tag
        local = tag.rsplit("}", 1)[-1]
        if tag in {w("del"), w("ins"), w("moveFrom"), w("moveTo")}:
            self.review.append({"kind": local, "author": attr(node, "author"),
                                "text": self.raw_text(node), "part": self.part})
            if local in {"del", "moveFrom"}:
                self.skip_markers(node)
                return ""
        if tag in COMMENT_MARKERS:
            self.comment_marker(node)
            return ""
        if tag == w("t"):
            # XML source line wrapping is whitespace, unlike Word's explicit w:br.
            return self.emit(re.sub(r"\s", " ", node.text or ""))
        if tag == w("delText"):
            return ""
        if tag in {w("br"), w("cr")}:
            return self.emit("\n", " ")
        if tag == w("tab"):
            return self.emit(" ")
        if tag == w("noBreakHyphen"):
            return self.emit("\u2011")
        if tag == w("softHyphen"):
            return self.emit("\u00ad")
        if tag in {w("footnoteReference"), w("endnoteReference")}:
            kind = "footnote" if tag == w("footnoteReference") else "endnote"
            self.features[kind + " references"] += 1
            return self.emit(f"[^{kind}-{attr(node, 'id')}]")
        if tag == w("hyperlink"):
            rid = node.get(f"{{{R}}}id")
            rel = self.rels.get(rid)
            target = rel[1] if rel else "#" + attr(node, "anchor")
            if rid and not rel:
                self.warn("Some hyperlink relationships are missing; destinations may be incomplete.")
            visible = "".join(self.inline(child) for child in node)
            self.links.append({"text": normalize(plain(visible)), "target": target, "part": self.part})
            return visible
        if tag == w("fldChar"):
            kind = attr(node, "fldCharType")
            if kind == "begin":
                data = node.find("w:ffData", NS)
                self.field_stack.append(([], None if data is None else self.form_field(data)))
            elif kind == "end" and self.field_stack:
                self.close_field()
            self.warn("Fields use saved display text, which can be stale; codes are fingerprinted, not evaluated (including Zotero).")
            return ""
        if tag == w("instrText"):
            if self.field_stack:
                self.field_stack[-1][0].append(node.text or "")
            else:
                self.record_field(node.text or "")
            return ""
        if tag == w("fldSimple"):
            self.record_field(attr(node, "instr"))
            self.warn("Fields use saved display text, which can be stale; codes are fingerprinted, not evaluated (including Zotero).")
        if tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
            return "".join(self.inline(child) for child in branch)
        if tag in EQUATIONS:
            self.record_objects(node)
            # Marked, so sentence splitting never breaks inside it; comment anchors get it unmarked.
            text = self.equation_text(node)
            return self.emit(MATH_START + text + MATH_END, text)
        if tag in {w("drawing"), w("pict"), f"{{{MC}}}AlternateContent"}:
            self.record_images(node)
        if tag in {w("drawing"), w("pict"), w("object"), w("altChunk"), f"{{{MC}}}AlternateContent", w("sym")}:
            self.record_objects(node)
        if tag in {w("drawing"), w("pict"), w("txbxContent")} and self.anchored and (boxes := self.text_boxes(node)):
            self.anchored[-1].extend(boxes)
            if not self.unrendered_shape_parts(node, self.anchor_objects(node)):
                return ""
        elif (tag in {w("drawing"), w("pict"), f"{{{MC}}}AlternateContent"} and self.anchored
              and (rendered := self.anchor_objects(node)) and not self.unrendered_shape_parts(node, rendered)):
            return ""
        if local in {"drawing", "pict", "object", "txbxContent", "AlternateContent", "altChunk", "sym"}:
            self.warn(f"Unsupported {local}: content is not rendered in prose; inspect the DOCX in Word.")
            self.features[local] += 1
            return self.emit(f"[Unsupported {local}]")
        if tag in {w("pPr"), w("rPr"), w("sdtPr"), w("sdtEndPr")}:
            return ""
        return "".join(self.inline(child) for child in node)

    def anchor_objects(self, node: Element) -> frozenset[Element]:
        """Render the charts and SmartArt diagrams in a drawing as anchored blocks, in document
        order; return those that rendered.

        They are found as objects are, so those in text boxes render with the box, and none
        render from an unselected mc:AlternateContent copy.
        """
        if self.unselected:
            return frozenset()
        rendered = []
        for found in self.object_nodes(node):
            block = self.chart(found) if found.tag == CHART else self.diagram(found) if found.tag == DIAGRAM else None
            if block is not None:
                self.anchored[-1].append(block)
                rendered.append(found)
        return frozenset(rendered)

    def close_field(self) -> None:
        instruction, form = self.field_stack.pop()
        self.record_field("".join(instruction), form)

    def record_field(self, instruction: str, form: dict[str, Any] | None = None) -> None:
        value = normalize(instruction)
        entry: dict[str, Any] = {"command": value.split(" ", 1)[0],
                                 "sha256": hashlib.sha256(value.encode()).hexdigest(), "part": self.part}
        if form is not None:
            # Only for checkbox and dropdown form fields, so other fields keep their metadata.
            entry["form"] = form
            self.warn(FORM_FIELD_NOTE)
        self.fields.append(entry)

    def form_field(self, data: Element) -> dict[str, Any] | None:
        """The state of a legacy checkbox or dropdown form field from its w:ffData, else None.

        A text input's visible value is already prose, so it records nothing here.
        """
        name = data.find("w:name", NS)
        label = attr(name, "val") if name is not None else ""
        box = data.find("w:checkBox", NS)
        if box is not None:
            # The current state, else the default state; an absent w:val is on, as for any ST_OnOff.
            state = box.find("w:checked", NS)
            if state is None:
                state = box.find("w:default", NS)
            return {"type": "checkbox", "name": label,
                    "checked": state is not None and on_off(attr(state, "val", "1"))}
        dropdown = data.find("w:ddList", NS)
        if dropdown is None:
            return None
        entries = [attr(entry, "val") for entry in dropdown.findall("w:listEntry", NS)]
        # w:result is the zero-based selected entry; without one, w:default applies, which defaults to 0.
        chosen = dropdown.find("w:result", NS)
        if chosen is None:
            chosen = dropdown.find("w:default", NS)
        index = 0 if chosen is None else decimal(attr(chosen, "val"))
        # An index that names no entry is recorded as null, never guessed.
        selected = entries[index] if index is not None and 0 <= index < len(entries) else None
        return {"type": "dropdown", "name": label, "entries": entries, "selected": selected}

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
        self.emit("", " ")  # a paragraph boundary is a space in comment anchors
        # Render the paragraph before its boxes so list first use follows prose order.
        own = self.paragraph_text(node, text)
        return "\n\n".join(part for part in (own, *(self.anchored_block(box) for box in boxes)) if part)

    def paragraph_text(self, node: Element, text: str) -> str:
        lines = math_sentences(text)
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
            return "#" * level + " " + normalize(plain(text))
        if num is not None:
            # Prose keeps the rule it had before list metadata: the nearest numPr alone decides
            # whether this is a list item and how far it is indented. The inherited numId and
            # the style-associated level are resolved for the metadata only.
            num_id = num.find("w:numId", NS)
            if num_id is None or attr(num_id, "val") != "0":
                self.warn(LIST_NOTE)
                self.record_list(props, num, self.numbering_level(node, props, num, style_id))
                ilvl = num.find("w:ilvl", NS)
                depth = attr(ilvl, "val", "0") if ilvl is not None else "0"
                indent = "  " * min(20, int(depth)) if depth.isdigit() else ""
                return indent + "- " + ("\n" + indent + "  ").join(lines)
        return "\n".join(lines)

    def list_id(self, props: list[Element], num: Element) -> str | None:
        """The nearest numId, including one inherited by a partial direct numPr."""
        # A numPr may carry only ilvl and inherit numId from its style chain.
        found = num.find("w:numId", NS)
        if found is None:
            found = next((n for p in props if (n := p.find("w:numPr/w:numId", NS)) is not None), None)
        return number_id(attr(found, "val")) if found is not None else None

    def numbering_level(self, paragraph: Element, props: list[Element], num: Element, style_id: str) -> int:
        """The level whose format is recorded: direct ilvl, else the level associated with the style."""
        ilvl = paragraph.find("w:pPr/w:numPr/w:ilvl", NS)
        if ilvl is None:
            num_id = self.list_id(props, num)
            definition = self.nums.get(num_id) if num_id is not None else None
            abstract_id = definition.find("w:abstractNumId", NS) if definition is not None else None
            abstract = None if abstract_id is None else self.abstract_nums.get(number_id(attr(abstract_id, "val")))
            levels = {} if definition is None or abstract is None else self.effective_levels(definition, abstract)
            seen = set()
            # A derived paragraph style may inherit its numbering from a basedOn style.
            # Match the nearest style first, independent of numbering-level XML order.
            while levels and style_id in self.styles and style_id not in seen:
                seen.add(style_id)
                for number, lvl in levels.items():
                    associated = lvl.find("w:pStyle", NS)
                    if associated is not None and attr(associated, "val") == style_id:
                        return number
                based = self.styles[style_id].find("w:basedOn", NS)
                style_id = attr(based, "val") if based is not None else ""
            ilvl = num.find("w:ilvl", NS)
        level = decimal(attr(ilvl, "val")) if ilvl is not None else None
        return level if level is not None and level >= 0 else 0

    def effective_levels(self, num: Element, abstract: Element) -> dict[int, Element]:
        """Each level's w:lvl once w:lvlOverride is applied, abstract levels first in document order."""
        levels: dict[int, Element] = {}
        for lvl in abstract.findall("w:lvl", NS):
            number = decimal(attr(lvl, "ilvl"))
            if number is not None and number >= 0:
                levels.setdefault(number, lvl)
        overridden = set()
        for override in num.findall("w:lvlOverride", NS):
            number = decimal(attr(override, "ilvl"))
            if number is None or number < 0 or number in overridden:
                continue  # the first override of a level is the one that applies
            overridden.add(number)
            replacement = override.find("w:lvl", NS)
            if replacement is not None:  # replaces the whole level, style association included
                levels[number] = replacement
        return levels

    def record_list(self, props: list[Element], num: Element, level: int) -> None:
        """Fingerprint the effective format of one used list level, keyed by first use."""
        num_id = self.list_id(props, num)
        if num_id == "0":
            return  # numbering switched off by a style; nothing to fingerprint
        key = num_id if num_id is not None else ""
        entry = self.lists.setdefault(key, {"list": len(self.lists) + 1, "part": self.part, "levels": {}})
        if level not in entry["levels"]:
            entry["levels"][level] = self.list_level(num_id, level)
        # The level is part of a run: items of one list can take different formats from their levels
        # without any difference in prose, when the level comes from the paragraph style.
        run = {"list": entry["list"], "part": self.part, "level": level}
        last = self.list_runs[-1] if self.list_runs else None
        if last is not None and all(last[key] == value for key, value in run.items()):
            last["items"] += 1
        else:
            self.list_runs.append({**run, "items": 1})

    def list_level(self, num_id: str | None, level: int) -> dict[str, Any]:
        """Resolve w:num -> w:abstractNum -> w:lvl, applying w:lvlOverride."""
        unknown = {"level": level, "format": "unknown", "text": "", "start": 0}
        num = self.nums.get(num_id) if num_id is not None else None
        abstract_id = num.find("w:abstractNumId", NS) if num is not None else None
        abstract = None if abstract_id is None else self.abstract_nums.get(number_id(attr(abstract_id, "val")))
        if num is None or abstract is None or abstract.find("w:numStyleLink", NS) is not None:
            self.warn(UNKNOWN_LIST_NOTE)
            return unknown
        override = next((o for o in num.findall("w:lvlOverride", NS) if decimal(attr(o, "ilvl")) == level), None)
        lvl = self.effective_levels(num, abstract).get(level)
        if lvl is None:
            self.warn(UNKNOWN_LIST_NOTE)
            return unknown
        fmt = lvl.find("w:numFmt", NS)
        alternate = lvl.find(f"{{{MC}}}AlternateContent")
        if fmt is None and alternate is not None:
            # Word stores custom formats such as "001, 002, ..." in mc:Choice, with a Fallback.
            fmt = next((f for branch in alternate for f in branch.findall("w:numFmt", NS)), None)
        # OOXML defaults an absent w:numFmt to decimal.
        format_name = attr(fmt, "val", "decimal") if fmt is not None else "decimal"
        if fmt is not None and attr(fmt, "format"):
            format_name += ": " + attr(fmt, "format")
        text = lvl.find("w:lvlText", NS)
        start = lvl.find("w:start", NS)
        start_override = override.find("w:startOverride", NS) if override is not None else None
        if start_override is not None:
            start = start_override
        return {"level": level, "format": format_name,
                "text": attr(text, "val") if text is not None else "",
                # Absent, non-numeric and oversized values record 0.
                "start": (decimal(attr(start, "val")) or 0) if start is not None else 0}

    def blocks(self, node: Element) -> str:
        parts = []
        for child in node:
            if child.tag == w("p"):
                value = self.paragraph(child)
            elif child.tag == w("tbl"):
                value = self.table(child)
            elif child.tag in {w("sectPr"), w("tcPr"), w("trPr"), w("tblPr"), w("tblGrid"), w("sdtPr"), w("sdtEndPr")}:
                continue
            elif child.tag in COMMENT_MARKERS:
                self.comment_marker(child)  # block-level range markers, e.g. between paragraphs
                continue
            elif child.tag in {w("del"), w("moveFrom")}:
                self.inline(child)  # record review, omit deleted blocks
                continue
            elif child.tag in EQUATIONS:
                # Body and cell content may hold math directly: it is a paragraph of its own.
                self.record_objects(child)
                value = self.emit(self.equation_text(child))
                self.emit("", " ")  # a paragraph boundary, as after a w:p
            elif child.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(child)) is not None:
                value = self.blocks(branch)  # one copy of block-level text box content
            elif child.tag == f"{{{MC}}}AlternateContent":
                # Without a text box every copy's prose is still read, but images come from one copy.
                chosen, values = self.image_branch(child), []
                for copy in child:
                    self.unselected += copy is not chosen
                    try:
                        values.append(self.blocks(copy))
                    finally:
                        self.unselected -= copy is not chosen
                value = "\n\n".join(v for v in values if v)
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
            elif child.tag in COMMENT_MARKERS:
                self.comment_marker(child)

    def table(self, node: Element) -> str:
        self.features["tables"] += 1
        lines = ["[Table]"]
        # Explicit cells avoid Markdown pipe escaping and ambiguity for multi-paragraph cells.
        for row in self.table_items(node, "tr"):
            if row.find("w:trPr/w:del", NS) is not None:
                self.warn("Tracked table structure changes are only partially represented; inspect the DOCX.")
                self.skip_markers(row)
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
        def visit_form(node: Element) -> None:
            for child in node:
                if child.tag not in FORM_DATA[node.tag]:
                    visit(child)
                elif child.tag in FORM_DATA:
                    visit_form(child)

        def visit(node: Element, math: bool = False) -> None:
            local = node.tag.rsplit("}", 1)[-1]
            # Math is rendered as linear text, which warns for math elements it does not know.
            math = math or node.tag in EQUATIONS
            if node.tag == w("ffData"):
                visit_form(node)  # form field data whose meaning is recorded, see FORM_DATA
                return
            if node.tag == f"{{{MC}}}AlternateContent" and (branch := self.branch(node)) is not None:
                # Only the copy holding the rendered text box is read.
                for child in branch:
                    visit(child, math)
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
            if not node.tag.startswith("{" + W + "}") and not (math and node.tag.startswith("{" + M + "}")):
                self.warn(f"Unsupported XML construct {local}; content may be omitted or approximate.")
            for child in node:
                visit(child, math)
        visit(tree)

    def render_part(self, part: str, tree: Element) -> str:
        self.part, self.rels = part, self.relationships(part)
        self.story, self.open_ranges = [], {}
        self.audit(tree)
        value = self.blocks(tree)
        self.close_open_ranges()
        if self.field_stack:
            self.warn("Unclosed field codes found; saved display text may be incomplete.")
            while self.field_stack:
                self.close_field()
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
        # Paragraph ID -> w15:commentEx (reply parent and resolved state); the first entry wins.
        extended: dict[str, Element] = {}
        for kind, target, external in main_rels.values():
            if kind == "styles" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                self.styles = {attr(s, "styleId"): s for s in self.xml(target)}
                supported.add(target)
            elif kind == "numbering" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                numbering = self.xml(target)
                for node in numbering.findall("w:num", NS):
                    self.nums.setdefault(number_id(attr(node, "numId")), node)
                for node in numbering.findall("w:abstractNum", NS):
                    self.abstract_nums.setdefault(number_id(attr(node, "abstractNumId")), node)
                supported.add(target)
            elif kind == "commentsExtended" and not external:
                if target not in self.names:
                    raise DocumentError(f"missing related part {target}")
                for node in self.xml(target):
                    key = para_id(node.get(f"{{{W15}}}paraId", ""))
                    if node.tag == f"{{{W15}}}commentEx" and key:
                        extended.setdefault(key, node)
                supported.add(target)
        prose = [self.render_part("word/document.xml", body)]
        # All supported related stories are distinct; never combine headers with body prose.
        stories = sorted({(kind, target) for kind, target, external in main_rels.values()
                          if kind in {"footnotes", "endnotes", "comments", "header", "footer"} and not external})
        # Comment entries keep their place in review; anchors and threads are filled in once
        # every story is read. Each is (comment ID, entry, paraId of its last paragraph).
        comments: list[tuple[str, dict[str, str | int | bool | None], str]] = []
        for kind, part in stories:
            if part not in self.names:
                raise DocumentError(f"missing related part {part}")
            supported.add(part)
            tree = self.xml(part)
            if kind == "comments":
                self.audit(tree)
                for comment in tree:
                    entry: dict[str, str | int | bool | None] = {
                        "kind": "comment", "author": attr(comment, "author"),
                        "text": self.raw_text(comment), "part": part}
                    self.review.append(entry)
                    comments.append((number_id(attr(comment, "id")), entry, self.last_paragraph_id(comment)))
                self.warn(COMMENT_NOTE)
            elif kind in {"footnotes", "endnotes"}:
                for note in tree:
                    if attr(note, "type") in {"separator", "continuationSeparator", "continuationNotice"}:
                        continue
                    singular = kind[:-1]
                    prose.append(f"[^{singular}-{attr(note, 'id')}]:\n\n" + self.render_part(part, note))
            else:
                prose.append(f"## {kind.title()} ({part})\n\n" + self.render_part(part, tree))
                self.warn("Headers and footers are shown by part; page/section placement is not reconstructed.")
        # A comment's position among all comments, by the paraId that w15:commentEx refers to.
        # Raw comment IDs and paraIds change when Word re-saves, so neither is recorded.
        positions: dict[str, int] = {}
        for position, (_, _, key) in enumerate(comments):
            if key:
                positions.setdefault(key, position)
        for position, (comment_id, entry, key) in enumerate(comments):
            ex = extended.get(key) if key and positions.get(key) == position else None
            # Without a matching w15:commentEx the state is unknown, never "not resolved".
            entry["resolved"] = None if ex is None else on_off(ex.get(f"{{{W15}}}done", ""))
            parent = None if ex is None else positions.get(para_id(ex.get(f"{{{W15}}}paraIdParent", "")))
            entry["reply_to"] = parent if parent != position else None
            if comment_id in self.comment_ranges:
                entry["anchor"], entry["anchor_part"] = self.comment_ranges[comment_id]
            elif comment_id in self.comment_refs:  # a point comment
                entry["anchor"], entry["anchor_part"] = "", self.comment_refs[comment_id]
            else:
                entry["anchor"] = entry["anchor_part"] = None
                self.warn(UNANCHORED_COMMENT_NOTE)
        supported.update(self.media)
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
        if any("alt" in image or "alt_title" in image for image in self.images):
            self.warnings.discard(IMAGE_NOTE)
            self.warn(IMAGE_ALT_NOTE)
        metadata = {"schema": 1, "view": "accept-insertions-omit-deletions",
                    "coverage": "partial-prose", "warnings": sorted(self.warnings),
                    "review": self.review, "hyperlinks": self.links, "fields": self.fields,
                    "lists": [{**entry, "levels": [entry["levels"][k] for k in sorted(entry["levels"])]}
                              for entry in self.lists.values()],
                    "list_runs": self.list_runs, "images": self.images}
        if self.objects:
            # Only when present, so documents without any of these objects keep their companions.
            metadata["objects"] = self.objects
        return Extraction("\n\n".join(p for p in prose if p).rstrip() + "\n", metadata)


def read_package(file: Path | BinaryIO) -> Extraction:
    with ZipFile(file) as package:
        infos = package.infolist()
        if len(infos) > MAX_ENTRIES:
            raise DocumentError(f"package has {len(infos)} entries; the safety limit is {MAX_ENTRIES}")
        total = sum(i.file_size for i in infos)
        if total > MAX_PACKAGE_BYTES:
            raise DocumentError(f"package is {total} bytes uncompressed; the safety limit is "
                                f"{size_label(MAX_PACKAGE_BYTES)}")
        for info in infos:
            if info.compress_type not in BOUNDED_COMPRESSION:
                raise DocumentError(f"{info.filename} uses ZIP compression method {info.compress_type}; "
                                    "DOCX packages may use only stored or deflate")
        if len({i.filename for i in infos}) != len(infos):
            raise DocumentError("duplicate ZIP entries")
        return Reader(package).extract()


def extract(source: Path | bytes) -> Extraction:
    """Extract a DOCX file, or a DOCX already read into memory, such as a Git blob.

    The same package limits apply to both. Only a file can change while it is
    read, so only a file is checked for that.
    """
    try:
        if isinstance(source, bytes):
            return read_package(io.BytesIO(source))
        before = source.stat()
        result = read_package(source)
        after = source.stat()
        if (before.st_ino, before.st_size, before.st_mtime_ns) != (after.st_ino, after.st_size, after.st_mtime_ns):
            raise DocumentError("document changed while reading; finish saving and retry")
        return result
    except (OSError, BadZipFile, RuntimeError, ValueError) as exc:
        if isinstance(exc, DocumentError):
            raise
        raise DocumentError(f"cannot read DOCX: {exc}; finish saving, check permissions, and retry") from exc
