"""Black-box regression tests built from small, generated OOXML packages.

The fixtures deliberately do not use Word, python-docx, or checked-in DOCX files.
Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any
from unittest.mock import patch
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROJECT = Path(__file__).resolve().parents[1]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"


def paragraph(text: str, properties: str = "") -> str:
    return f'<w:p>{properties}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}</w:body></w:document>'
    )


MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V = "urn:schemas-microsoft-com:vml"
TEXT_BOX_NOTE = (
    "Text box text is shown in [Text box] blocks after its anchoring paragraph; "
    "position, anchoring, wrapping and shape formatting are not compared."
)


def drawing_text_box(content: str, *, picture: bool = False) -> str:
    """A DrawingML text box shape (wps:wsp) as Word writes it, optionally grouped with a picture."""
    shape = (
        f'<wps:wsp xmlns:wps="{WPS}"><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr>'
        f'<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp>'
    )
    uri = WPS
    if picture:
        uri = WPG
        shape = f'<wpg:wgp xmlns:wpg="{WPG}"><wpg:cNvGrpSpPr/><wpg:grpSpPr/>{shape}<pic:pic xmlns:pic="{PIC}"><pic:blipFill><a:blip r:embed="image1"/></pic:blipFill></pic:pic></wpg:wgp>'
    return (
        f'<w:drawing><wp:anchor xmlns:wp="{WP}" distT="0" distB="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="{A}"><a:graphicData uri="{uri}">'
        f'{shape}</a:graphicData></a:graphic></wp:anchor></w:drawing>'
    )


def vml_text_box(content: str) -> str:
    """A legacy VML text box, as in older files and in mc:Fallback."""
    return (
        f'<w:pict><v:shapetype xmlns:v="{V}" id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
        f'</v:shapetype><v:shape xmlns:v="{V}" id="Text Box 1" type="#_x0000_t202" style="position:absolute;width:144pt;height:72pt">'
        f'<v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict>'
    )


def modern_text_box(content: str, *, fallback: str | None = None, picture: bool = False) -> str:
    """mc:AlternateContent with a DrawingML Choice and a VML Fallback copy of the same text."""
    return (
        f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="wps">{drawing_text_box(content, picture=picture)}</mc:Choice>'
        f'<mc:Fallback>{vml_text_box(content if fallback is None else fallback)}</mc:Fallback></mc:AlternateContent>'
    )


def anchored(text: str, shape: str) -> str:
    """A paragraph with its own text and one shape in a following run."""
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r><w:r>{shape}</w:r></w:p>'


NUMBERING_REL = f'<Relationship Id="numbering" Type="{R}/numbering" Target="numbering.xml"/>'
LIST_NOTE = (
    "List boundaries and indentation are retained; bullet and number formats are recorded in "
    "metadata, but counter values and bullet glyphs are not reconstructed."
)
UNKNOWN_LIST_NOTE = (
    "Some list numbering formats could not be resolved (no numbering part, an unknown "
    "list or level, or style-linked numbering); they are recorded as unknown."
)


def list_item(text: str, num_id: str, ilvl: int = 0) -> str:
    return paragraph(text, f'<w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>')


def level(ilvl: int = 0, fmt: str = "bullet", text: str = "•", start: str | None = "1", *,
          style: str | None = None) -> str:
    start_xml = f'<w:start w:val="{start}"/>' if start is not None else ""
    style_xml = f'<w:pStyle w:val="{style}"/>' if style is not None else ""
    return f'<w:lvl w:ilvl="{ilvl}">{start_xml}<w:numFmt w:val="{fmt}"/>{style_xml}<w:lvlText w:val="{text}"/></w:lvl>'


def abstract_num(abstract_id: str, *levels: str) -> str:
    return f'<w:abstractNum w:abstractNumId="{abstract_id}">{"".join(levels)}</w:abstractNum>'


def num(num_id: str, abstract_id: str, overrides: str = "") -> str:
    return f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/>{overrides}</w:num>'


def numbering(*definitions: str) -> str:
    return f'<w:numbering xmlns:w="{W}">{"".join(definitions)}</w:numbering>'


IMAGE_NOTE = "Images are compared by content hash only; size, position, cropping, and alt text are not compared."
IMAGE_ALT_NOTE = "Images are compared by content hash and alt text; size, position, and cropping are not compared."
MISSING_IMAGE_NOTE = (
    "Some image relationships are missing or point to missing parts; "
    "those images are recorded without a hash."
)


def picture(rid: str, attribute: str = "embed") -> str:
    """An inline DrawingML picture referencing an image relationship."""
    return (
        f'<w:drawing><wp:inline xmlns:wp="{WP}"><a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}">'
        f'<pic:pic xmlns:pic="{PIC}"><pic:blipFill><a:blip r:{attribute}="{rid}"/></pic:blipFill></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    )


def vml_picture(rid: str) -> str:
    """A legacy VML picture, as in older files and in mc:Fallback."""
    return f'<w:pict><v:shape xmlns:v="{V}" id="Picture 1" style="width:72pt;height:72pt"><v:imagedata r:id="{rid}"/></v:shape></w:pict>'


def quoted(value: str) -> str:
    """``value`` escaped for a double-quoted XML attribute."""
    return escape(value, {'"': "&quot;", "\n": "&#10;"})


def described_picture(rid: str, descr: str | None = None, *, own: str | None = None,
                      title: str | None = None) -> str:
    """An inline DrawingML picture with alt text on its wp:docPr (``descr``, ``title``) and on
    its pic:cNvPr (``own``), as Word writes it; None leaves an attribute out."""
    def attributes(**values: str | None) -> str:
        return "".join(f' {key}="{quoted(value)}"' for key, value in values.items() if value is not None)
    return (f'<w:drawing><wp:inline xmlns:wp="{WP}"><wp:extent cx="914400" cy="914400"/>'
            f'<wp:docPr id="1" name="Picture 1"{attributes(descr=descr, title=title)}/>'
            f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}">'
            f'<pic:nvPicPr><pic:cNvPr id="0" name="cat.png"{attributes(descr=own)}/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing>')


def vml_wordart(text: str) -> str:
    """Legacy VML WordArt: a v:shape drawing ``text`` along its v:textpath, with the shape type
    Word writes next to it (whose own v:textpath has no text)."""
    return (f'<w:pict><v:shapetype xmlns:v="{V}" id="_x0000_t136" coordsize="21600,21600" o:spt="136" '
            f'xmlns:o="urn:schemas-microsoft-com:office:office" path="m@7,l@8,m@5,21600l@6,21600e">'
            '<v:path textpathok="t" o:connecttype="custom"/><v:textpath on="t" fitshape="t"/></v:shapetype>'
            f'<v:shape xmlns:v="{V}" id="WordArt 1" type="#_x0000_t136" style="width:200pt;height:50pt" fillcolor="black">'
            f'<v:textpath style="font-family:&quot;Arial Black&quot;;v-text-kern:t" trim="t" fitpath="t" '
            f'string="{quoted(text)}"/></v:shape></w:pict>')


def wordart_entry(text: str, part: str = "word/document.xml") -> dict[str, Any]:
    return {"kind": "WordArt", "part": part, "text": text}


C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
VML_O = "urn:schemas-microsoft-com:office:office"


def graphic(uri: str, content: str) -> str:
    """An inline DrawingML graphic frame holding ``content``."""
    return (f'<w:drawing><wp:inline xmlns:wp="{WP}"><a:graphic xmlns:a="{A}"><a:graphicData uri="{uri}">'
            f'{content}</a:graphicData></a:graphic></wp:inline></w:drawing>')


def chart(rid: str) -> str:
    """An inline DrawingML chart referencing a chart part relationship."""
    return graphic(C, f'<c:chart xmlns:c="{C}" r:id="{rid}"/>')


def smartart(data: str, layout: str = "dgmLayout", style: str = "dgmStyle", colors: str = "dgmColors") -> str:
    """An inline SmartArt diagram referencing its data, layout, quick-style and colour parts."""
    return graphic(DGM, f'<dgm:relIds xmlns:dgm="{DGM}" r:dm="{data}" r:lo="{layout}" r:qs="{style}" r:cs="{colors}"/>')


def ole_object(rid: str, prog_id: str = "Excel.Sheet.12", kind: str = "Embed") -> str:
    """A w:object with its VML preview shape and an o:OLEObject referencing the embedding."""
    return (f'<w:object><v:shape xmlns:v="{V}" id="_x0000_i1025" style="width:300pt;height:150pt"/>'
            f'<o:OLEObject xmlns:o="{VML_O}" Type="{kind}" ProgID="{escape(prog_id, {chr(34): "&quot;"})}" '
            f'ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="{rid}"/></w:object>')


def alt_chunk(rid: str) -> str:
    return f'<w:altChunk r:id="{rid}"/>'


def rel(rid: str, kind: str, target: str, *, external: bool = False) -> str:
    """A relationship of type ``kind`` (e.g. chart, diagramData, oleObject, aFChunk)."""
    mode = ' TargetMode="External"' if external else ""
    return f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{target}"{mode}/>'


def part_entry(kind: str, data: str | None, part: str = "word/document.xml", **extra: str) -> dict[str, Any]:
    """The objects entry for an object whose target part holds ``data`` (None: missing)."""
    sha = None if data is None else hashlib.sha256(data.encode("utf-8")).hexdigest()
    return {"kind": kind, "part": part, "sha256": sha, **extra}


M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def math_run(text: str) -> str:
    return f"<m:r><m:t>{escape(text)}</m:t></m:r>"


def equation(content: str = "E=mc2", *, para: bool = False) -> str:
    """An OMML equation (m:oMath, or m:oMathPara around one); a plain string is one math run."""
    inner = content if content.startswith("<") else math_run(content)
    math = f'<m:oMath xmlns:m="{M}">{inner}</m:oMath>'
    return f'<m:oMathPara xmlns:m="{M}">{math}</m:oMathPara>' if para else math


def symbol(char: str = "F061", font: str = "Symbol") -> str:
    return f'<w:r><w:sym w:font="{font}" w:char="{char}"/></w:r>'


def equation_entry(form: list[Any], part: str = "word/document.xml") -> dict[str, Any]:
    """The objects entry for an equation whose canonical form is ``form``."""
    data = json.dumps(form, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return {"kind": "equation", "part": part, "sha256": hashlib.sha256(data).hexdigest()}


def simple_form(text: str = "E=mc2") -> list[Any]:
    """The canonical form of ``equation(text)``."""
    return [f"{{{M}}}oMath", [], "", [[f"{{{M}}}r", [], "", [[f"{{{M}}}t", [], text, []]]]]]


def math_form(text: str) -> list[Any]:
    """The canonical form of ``math_run(text)``."""
    return [f"{{{M}}}r", [], "", [[f"{{{M}}}t", [], text, []]]]


def structure(name: str, *args: tuple[str, str], props: str = "") -> str:
    """An OMML structure m:<name> with its m:<name>Pr properties and (argument, content) children;
    plain-string content is one math run."""
    pr = f"<m:{name}Pr>{props}</m:{name}Pr>" if props else ""
    inner = "".join(f"<m:{arg}>{content if content.startswith('<') or not content else math_run(content)}</m:{arg}>"
                    for arg, content in args)
    return f"<m:{name}>{pr}{inner}</m:{name}>"


def energy(power: str = "2") -> str:
    """E=mc² as Word writes it: a run and a superscript."""
    return math_run("E=m") + structure("sSup", ("e", "c"), ("sup", power))


def energy_form(power: str = "2") -> list[Any]:
    """The canonical form of ``equation(energy(power))``."""
    return [f"{{{M}}}oMath", [], "", [math_form("E=m"), [f"{{{M}}}sSup", [], "", [
        [f"{{{M}}}e", [], "", [math_form("c")]], [f"{{{M}}}sup", [], "", [math_form(power)]]]]]]


MATH_NOTE = ("Equations are shown as linear [Math: …] text; layout, fonts and alignment are not compared, "
             "and equation fingerprints are in objects.")


def image_rel(rid: str, target: str, *, external: bool = False) -> str:
    mode = ' TargetMode="External"' if external else ""
    return f'<Relationship Id="{rid}" Type="{R}/image" Target="{target}"{mode}/>'


def image_entry(data: str, part: str = "word/document.xml") -> dict[str, Any]:
    raw = data.encode("utf-8")
    return {"part": part, "sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)}


def form_field(command: str, data: str, display: str = "", *, name: str | None = "Field1", rsid: str = "") -> str:
    """A legacy form field as Word writes it: ``data`` is the type-specific part of its w:ffData.

    Checkboxes and dropdowns have no display text; a text input's ``display`` is its value.
    """
    named = "" if name is None else f'<w:name w:val="{escape(name)}"/>'
    runs = '<w:r{rsid}><w:fldChar w:fldCharType="begin"><w:ffData>{named}<w:enabled/><w:calcOnExit w:val="0"/>{data}'
    runs += '</w:ffData></w:fldChar></w:r><w:r{rsid}><w:instrText xml:space="preserve"> {command} </w:instrText></w:r>'
    if display:
        runs += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(display)
    runs += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    return runs.format(rsid=f' w:rsidR="{rsid}"' if rsid else "", named=named, data=data, command=command)


def checkbox(checked: str | None = "0", default: str | None = "0", *, name: str | None = "Check1", rsid: str = "") -> str:
    """A FORMCHECKBOX; "" for ``checked`` or ``default`` writes the element without w:val, None leaves it out."""
    def on_off(tag: str, value: str | None) -> str:
        return "" if value is None else f"<w:{tag}/>" if value == "" else f'<w:{tag} w:val="{value}"/>'
    data = f'<w:checkBox><w:sizeAuto/>{on_off("default", default)}{on_off("checked", checked)}</w:checkBox>'
    return form_field("FORMCHECKBOX", data, name=name, rsid=rsid)


def dropdown(*entries: str, result: str | None = "0", default: str | None = None, name: str | None = "Drop1") -> str:
    """A FORMDROPDOWN with ``entries``; None leaves w:result or w:default out."""
    data = "<w:ddList>" + ("" if result is None else f'<w:result w:val="{result}"/>')
    data += "" if default is None else f'<w:default w:val="{default}"/>'
    data += "".join(f'<w:listEntry w:val="{escape(entry)}"/>' for entry in entries) + "</w:ddList>"
    return form_field("FORMDROPDOWN", data, name=name)


def checkbox_entry(checked: bool, name: str = "Check1", part: str = "word/document.xml") -> dict[str, Any]:
    return {"command": "FORMCHECKBOX", "sha256": hashlib.sha256(b"FORMCHECKBOX").hexdigest(), "part": part,
            "form": {"type": "checkbox", "name": name, "checked": checked}}


def dropdown_form(entries: list[str], selected: str | None, name: str = "Drop1") -> dict[str, Any]:
    return {"type": "dropdown", "name": name, "entries": entries, "selected": selected}


FORM_FIELD_NOTE = ("Legacy form field checkbox and dropdown state is recorded in field metadata; "
                   "text form field settings are not compared.")


def make_docx(
    path: Path,
    body: str,
    *,
    extra: dict[str, str] | None = None,
    relationships: str = "",
    churn: bool = False,
) -> None:
    """Write enough OPC structure for a conventional, self-contained DOCX."""
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{PKG_R}"><Relationship Id="rId1" '
            f'Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>'
        ),
        "word/document.xml": document(body),
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{PKG_R}">{relationships}</Relationships>',
    }
    if extra:
        parts.update(extra)
    if churn:
        parts["docProps/core.xml"] = (
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dcterms="http://purl.org/dc/terms/"><dcterms:modified>2030-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>'
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    items = sorted(parts.items(), reverse=churn)
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in items:
            info = ZipInfo(name, date_time=(2030 if churn else 2020, 1, 2, 3, 4, 6))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, content.encode("utf-8"))


COMMENTS_REL = f'<Relationship Id="comments" Type="{R}/comments" Target="comments.xml"/>'
COMMENT_NOTE = (
    "Comment text, authors, anchored text, replies, and resolved state are recorded; "
    "dates and rich formatting are not compared."
)
UNANCHORED_COMMENT_NOTE = (
    "Some comments are not anchored in any rendered part; "
    "their anchor is recorded as null."
)
UNCLOSED_COMMENT_NOTE = (
    "Some comment ranges are missing a start or an end; their anchors run "
    "from the start or to the end of the story part."
)


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def start(comment: str) -> str:
    return f'<w:commentRangeStart w:id="{comment}"/>'


def end(comment: str) -> str:
    """A range end followed by the reference run, as Word writes it."""
    return f'<w:commentRangeEnd w:id="{comment}"/>' + reference(comment)


def reference(comment: str) -> str:
    return f'<w:r><w:commentReference w:id="{comment}"/></w:r>'


W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def comments_part(comments: dict[str, str], para_ids: dict[str, str] | None = None) -> str:
    """word/comments.xml with one comment by Ada per ID, in the given order.

    ``para_ids`` gives a comment ID's paragraph the w14:paraId that w15:commentEx refers to.
    """
    def comment_paragraph(comment: str, text: str) -> str:
        if para_ids is None or comment not in para_ids:
            return paragraph(text)
        return f'<w:p w14:paraId="{para_ids[comment]}" w14:textId="77777777">{run(text)}</w:p>'
    return f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}">' + "".join(
        f'<w:comment w:id="{comment}" w:author="Ada">{comment_paragraph(comment, text)}</w:comment>'
        for comment, text in comments.items()) + "</w:comments>"


COMMENTS_EXT_REL = ('<Relationship Id="commentsExtended" Type="http://schemas.microsoft.com/office/2011/'
                    'relationships/commentsExtended" Target="commentsExtended.xml"/>')


def comments_extended(*entries: tuple[str, str | None, str | None]) -> str:
    """word/commentsExtended.xml with one w15:commentEx per (paraId, parent paraId, done)."""
    def entry(para: str, parent: str | None, done: str | None) -> str:
        parent_xml = f' w15:paraIdParent="{parent}"' if parent is not None else ""
        done_xml = f' w15:done="{done}"' if done is not None else ""
        return f'<w15:commentEx w15:paraId="{para}"{parent_xml}{done_xml}/>'
    return (f'<w15:commentsEx xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15">'
            + "".join(entry(*e) for e in entries) + "</w15:commentsEx>")


def object_parts(*, value: str = "10", workbook: str = "workbook v1", node: str = "Plan",
                 chunk: str = "Chunk text") -> dict[str, str]:
    """The parts a chart, a SmartArt diagram, an embedded workbook and an imported chunk point to."""
    return {
        "word/charts/chart1.xml": (f'<c:chartSpace xmlns:c="{C}" xmlns:r="{R}"><c:chart><c:plotArea><c:barChart>'
                                   f'<c:ser><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>{value}</c:v></c:pt>'
                                   '</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>'
                                   '<c:externalData r:id="rId1"/></c:chartSpace>'),
        "word/charts/_rels/chart1.xml.rels": (f'<Relationships xmlns="{PKG_R}">'
                                              + rel("rId1", "package", "../embeddings/chart_workbook.xlsx")
                                              + "</Relationships>"),
        "word/embeddings/chart_workbook.xlsx": "chart workbook bytes",
        "word/diagrams/data1.xml": (f'<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst><dgm:pt modelId="1">'
                                    f"<dgm:t><a:p><a:r><a:t>{node}</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>"
                                    "</dgm:dataModel>"),
        "word/diagrams/layout1.xml": "<layout/>",
        "word/diagrams/quickStyle1.xml": "<quickStyle/>",
        "word/diagrams/colors1.xml": "<colors/>",
        "word/embeddings/Microsoft_Excel_Worksheet.xlsx": workbook,
        "word/afchunk.htm": f"<html><body><p>{chunk}</p></body></html>",
    }


# Relationships and a body for object_parts(): every object kind, inline and block-level.
OBJECT_RELS = (rel("chart", "chart", "charts/chart1.xml") + rel("dgmData", "diagramData", "diagrams/data1.xml")
               + rel("dgmLayout", "diagramLayout", "diagrams/layout1.xml")
               + rel("dgmStyle", "diagramQuickStyle", "diagrams/quickStyle1.xml")
               + rel("dgmColors", "diagramColors", "diagrams/colors1.xml")
               + rel("ole", "package", "embeddings/Microsoft_Excel_Worksheet.xlsx")
               + rel("chunk", "aFChunk", "afchunk.htm"))
OBJECT_BODY = (f'<w:p>{run("Results: ")}<w:r>{chart("chart")}</w:r><w:r>{smartart("dgmData")}</w:r>'
               f'<w:r>{ole_object("ole")}</w:r></w:p>{alt_chunk("chunk")}')


class CLITests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def cli(
        self, *args: str | Path, code: int = 0, cwd: Path | None = None, env: dict[str, str] | None = None,
    ) -> subprocess.CompletedProcess[str]:
        env = {**os.environ, **(env or {})}
        env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), str(PROJECT), env.get("PYTHONPATH", "")))
        result = subprocess.run(
            [sys.executable, "-m", "pilcrow", *(str(arg) for arg in args)],
            cwd=cwd or self.root,
            env=env,
            capture_output=True,
            text=True,
            timeout=20,
        )
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout}\nstderr={result.stderr}")
        return result

    def companions(self, relative: str = "essay.docx") -> tuple[Path, Path]:
        base = self.root / ".pilcrow" / relative
        return Path(str(base) + ".prose.md"), Path(str(base) + ".meta.json")

    def exported(self, relative: str = "essay.docx") -> tuple[bytes, bytes]:
        prose, meta = self.companions(relative)
        self.assertTrue(prose.is_file(), str(prose))
        self.assertTrue(meta.is_file(), str(meta))
        json.loads(meta.read_text(encoding="utf-8"))
        return prose.read_bytes(), meta.read_bytes()

    def git(self, *args: str) -> str:
        result = subprocess.run(["git", *args], cwd=self.root, capture_output=True, text=True, timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout

    def init_git(self) -> None:
        self.git("init", "-q")
        self.git("config", "user.name", "Pilcrow Test")
        self.git("config", "user.email", "pilcrow@example.invalid")
        self.git("config", "commit.gpgsign", "false")

    def test_explicit_selection_is_required(self) -> None:
        self.cli("export", "--root", self.root, code=2)
        self.cli("check", "--root", self.root, code=2)
        self.assertFalse((self.root / ".pilcrow").exists())

    def test_all_and_paths_cannot_be_combined(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("A sentence."))
        self.cli("export", "essay.docx", "--all", "--root", self.root, code=2)
        self.assertFalse((self.root / ".pilcrow").exists())

    def test_export_and_check_preserve_read_only_source(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("A complete sentence."))
        source.chmod(0o444)
        original = source.read_bytes()
        before = source.stat()
        self.cli("export", source, "--root", self.root)
        self.cli("check", source, "--root", self.root)
        prose, _ = self.exported()
        self.assertIn(b"A complete sentence.", prose)
        after = source.stat()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual((after.st_mode, after.st_mtime_ns), (before.st_mode, before.st_mtime_ns))

    def test_repeat_exports_are_byte_identical(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("Café, İstanbul, and naïve prose."))
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        self.cli("export", source, "--root", self.root)
        self.assertEqual(self.exported(), before)

    def test_run_splits_and_save_churn_do_not_change_companions(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("A single uninterrupted sentence."))
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        original_zip = source.read_bytes()
        split = (
            '<w:p w:rsidR="12345678"><w:r w:rsidRPr="87654321"><w:t xml:space="preserve">A single </w:t></w:r>'
            '<w:r><w:t>uninterrupted</w:t></w:r><w:r><w:t xml:space="preserve"> sentence.</w:t></w:r></w:p>'
        )
        make_docx(source, split, churn=True)
        self.assertNotEqual(source.read_bytes(), original_zip)
        self.cli("check", source, "--root", self.root)
        self.cli("export", source, "--root", self.root)
        self.assertEqual(self.exported(), before)

    def test_localized_prose_edit_is_visible_and_check_is_read_only(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("First paragraph.") + paragraph("Second paragraph.") + paragraph("Third paragraph."))
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        make_docx(source, paragraph("First paragraph.") + paragraph("Revised second paragraph.") + paragraph("Third paragraph."))
        self.cli("check", source, "--root", self.root, code=1)
        self.assertEqual(self.exported(), before, "check must not regenerate stale files")
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(after[0], before[0].replace(b"Second paragraph.", b"Revised second paragraph."))
        self.cli("check", source, "--root", self.root)

    def test_sentences_get_lines_while_paragraph_boundaries_remain(self) -> None:
        make_docx(
            self.root / "essay.docx",
            paragraph("Dr. Smith cites Art. 12. The value is 3.14. Is this clear? Yes!")
            + paragraph("A separate paragraph."),
        )
        self.cli("export", "essay.docx", "--root", self.root)
        prose = self.exported()[0].decode("utf-8")
        self.assertIn("The value is 3.14.\nIs this clear?\nYes!\n\nA separate paragraph.", prose)
        self.assertIn("Dr. Smith", prose)
        self.assertNotIn("Dr.\n", prose)

    def test_missing_or_edited_companions_are_stale(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("A sentence."))
        self.cli("check", source, "--root", self.root, code=1)
        self.assertFalse((self.root / ".pilcrow").exists())
        self.cli("export", source, "--root", self.root)
        prose, meta = self.companions()
        prose.write_text("Hand-edited companion.\n", encoding="utf-8")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        meta.unlink()
        self.cli("check", source, "--root", self.root, code=1)

    def test_paths_are_relative_to_cwd_and_constrained_by_root(self) -> None:
        nested = self.root / "nested"
        nested.mkdir()
        make_docx(nested / "essay.docx", paragraph("Nested prose."))
        self.cli("export", "essay.docx", "--root", self.root, cwd=nested)
        self.exported("nested/essay.docx")
        outside = self.root / "outside.docx"
        make_docx(outside, paragraph("Outside."))
        self.cli("export", outside, "--root", nested, code=2)

    def test_symlink_source_cannot_escape_root(self) -> None:
        nested = self.root / "restricted"
        nested.mkdir()
        make_docx(self.root / "outside.docx", paragraph("Outside."))
        try:
            (nested / "escape.docx").symlink_to(self.root / "outside.docx")
        except OSError:
            self.skipTest("symlink creation unavailable on this platform")
        self.cli("export", "escape.docx", "--root", nested, cwd=nested, code=2)
        self.assertFalse((nested / ".pilcrow").exists())

    def test_paths_through_an_alias_of_root_are_inside_root(self) -> None:
        # macOS spells temp directories as /var -> /private/var, and Windows as
        # 8.3 short names; --root is resolved, so the alias must still match.
        real = self.root / "real"
        real.mkdir()
        alias = self.root / "alias"
        try:
            alias.symlink_to(real, target_is_directory=True)
        except OSError:
            self.skipTest("symlink creation unavailable on this platform")
        make_docx(real / "essay.docx", paragraph("Aliased root."))
        self.cli("export", alias / "essay.docx", "--root", alias)
        self.cli("check", "alias/essay.docx", "--root", alias)
        self.assertIn(b"Aliased root.", (real / ".pilcrow/essay.docx.prose.md").read_bytes())
        # A symlink inside the root is still rejected when reached through the alias.
        make_docx(self.root / "outside.docx", paragraph("Outside."))
        (real / "escape.docx").symlink_to(self.root / "outside.docx")
        self.cli("export", alias / "escape.docx", "--root", alias, code=2)
        self.assertFalse((real / ".pilcrow/escape.docx.prose.md").exists())

    def test_unencodable_names_do_not_fail_export(self) -> None:
        # Windows pipes stdout through the ANSI code page, which lacks many letters,
        # with the surrogateescape error handler by default.
        make_docx(self.root / "İstanbul résumé.docx", paragraph("Encoded name."))
        for encoding in ("ascii:strict", "ascii:surrogateescape"):
            with self.subTest(encoding=encoding):
                result = self.cli("export", "İstanbul résumé.docx", env={"PYTHONIOENCODING": encoding})
                self.assertIn("\\u0130stanbul r\\xe9sum\\xe9.docx", result.stdout)
        self.assertIn(b"Encoded name.", self.exported("İstanbul résumé.docx")[0])

    def test_default_root_is_current_directory(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Default root."))
        self.cli("export", "essay.docx")
        self.exported()

    def test_git_discovery_handles_spaces_unicode_ignore_and_lock_files(self) -> None:
        self.init_git()
        (self.root / ".gitignore").write_text("ignored/\n", encoding="utf-8")
        paths = ("essay.docx", "nested with spaces/İstanbul résumé.docx", "ignored/tracked.docx")
        for relative in paths:
            make_docx(self.root / relative, paragraph(relative))
        self.git("add", "essay.docx", "nested with spaces/İstanbul résumé.docx")
        self.git("add", "-f", "ignored/tracked.docx")
        make_docx(self.root / "ignored/untracked.docx", paragraph("Ignored."))
        make_docx(self.root / "untracked.docx", paragraph("Included untracked."))
        make_docx(self.root / "~$essay.docx", paragraph("Office lock file."))
        self.cli("export", "--all", "--root", self.root)
        for relative in (*paths, "untracked.docx"):
            self.exported(relative)
        self.assertFalse(self.companions("ignored/untracked.docx")[0].exists())
        self.assertFalse(self.companions("~$essay.docx")[0].exists())
        self.cli("check", "--all", "--root", self.root)

    def test_all_discovers_under_root_outside_git(self) -> None:
        for relative in ("essay.docx", "nested/another document.docx"):
            make_docx(self.root / relative, paragraph(relative))
        (self.root / "unrelated.txt").write_text("Not a document.", encoding="utf-8")
        self.cli("export", "--all", "--root", self.root)
        self.exported()
        self.exported("nested/another document.docx")

    def test_deleted_source_reports_orphans_without_deleting_them(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("Deleted later."))
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        source.unlink()
        result = self.cli("check", "--all", "--root", self.root, code=1)
        self.assertIn("essay.docx", result.stdout + result.stderr)
        self.cli("export", "--all", "--root", self.root, code=1)
        self.assertEqual(self.exported(), before)

    def test_missing_unreadable_and_malformed_sources_are_errors(self) -> None:
        self.cli("export", "missing.docx", "--root", self.root, code=2)
        malformed = self.root / "malformed.docx"
        malformed.write_bytes(b"This is not a ZIP archive.")
        self.cli("export", malformed, "--root", self.root, code=2)
        with ZipFile(malformed, "w") as archive:
            archive.writestr("word/document.xml", "<broken>")
        self.cli("export", malformed, "--root", self.root, code=2)
        if os.name != "posix" or os.geteuid() == 0:
            return  # chmod cannot simulate denied access on Windows or as root
        unreadable = self.root / "unreadable.docx"
        make_docx(unreadable, paragraph("Cannot read."))
        unreadable.chmod(0)
        self.addCleanup(unreadable.chmod, 0o600)
        self.cli("export", unreadable, "--root", self.root, code=2)

    def test_dtd_and_entity_expansion_are_rejected(self) -> None:
        source = self.root / "essay.docx"
        with ZipFile(source, "w") as archive:
            archive.writestr(
                "word/document.xml",
                f'<!DOCTYPE w:document [<!ENTITY injected "untrusted expansion">]><w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&injected;</w:t></w:r></w:p></w:body></w:document>',
            )
        self.cli("export", source, "--root", self.root, code=2)
        self.assertFalse(self.companions()[0].exists())

    def test_headings_lists_and_tables_are_readable(self) -> None:
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>'
        numbering = (
            f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
            '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
            '</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
        )
        body = paragraph("A section", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')
        body += paragraph("A list item", '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>')
        body += '<w:tbl><w:tr><w:tc>' + paragraph("Left cell") + '</w:tc><w:tc>' + paragraph("Right cell") + '</w:tc></w:tr></w:tbl>'
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/><Relationship Id="numbering" Type="{R}/numbering" Target="numbering.xml"/>'
        make_docx(self.root / "essay.docx", body, extra={"word/styles.xml": styles, "word/numbering.xml": numbering}, relationships=relationships)
        self.cli("export", "essay.docx", "--root", self.root)
        prose = self.exported()[0].decode("utf-8")
        self.assertIn("# A section", prose)
        self.assertRegex(prose, r"(?m)^\s*[-*+] A list item")
        self.assertIn("Left cell", prose)
        self.assertIn("Right cell", prose)
        self.assertLess(prose.index("A section"), prose.index("A list item"))
        self.assertLess(prose.index("Left cell"), prose.index("Right cell"))

    def write_lists(self, body: str, numbering_xml: str | None, *, extra: dict[str, str] | None = None,
                    relationships: str = "") -> None:
        parts = dict(extra or {})
        if numbering_xml is not None:
            parts["word/numbering.xml"] = numbering_xml
            relationships += NUMBERING_REL
        make_docx(self.root / "essay.docx", body, extra=parts, relationships=relationships)

    def exported_lists(self) -> list[dict[str, Any]]:
        lists: list[dict[str, Any]] = json.loads(self.exported()[1])["lists"]
        return lists

    def test_list_format_changes_are_stale_in_metadata_only(self) -> None:
        body = paragraph("Intro.") + list_item("First item.", "1") + list_item("Second item.", "1")

        def write(definition: str, overrides: str = "") -> None:
            self.write_lists(body, numbering(abstract_num("0", definition), num("1", "0", overrides)))

        write(level(fmt="bullet", text="•"))
        result = self.cli("export", "essay.docx")
        self.assertIn(LIST_NOTE, result.stderr)
        before = self.exported()
        self.assertEqual(self.exported_lists(), [
            {"list": 1, "part": "word/document.xml",
             "levels": [{"level": 0, "format": "bullet", "text": "•", "start": 1}]},
        ])
        self.assertIn(LIST_NOTE, json.loads(before[1])["warnings"])
        variants = [
            (level(fmt="decimal", text="%1."), "", {"format": "decimal", "text": "%1.", "start": 1}),
            (level(fmt="decimal", text="%1)"), "", {"format": "decimal", "text": "%1)", "start": 1}),
            (level(fmt="decimal", text="%1)", start="5"), "", {"format": "decimal", "text": "%1)", "start": 5}),
            (level(fmt="decimal", text="%1)", start="5"),
             '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>',
             {"format": "decimal", "text": "%1)", "start": 3}),
            (level(fmt="decimal", text="%1)", start="5"),
             f'<w:lvlOverride w:ilvl="0">{level(fmt="lowerLetter", text="%1)", start="2")}</w:lvlOverride>',
             {"format": "lowerLetter", "text": "%1)", "start": 2}),
            (level(fmt="decimal", text="%1)", start=None), "", {"format": "decimal", "text": "%1)", "start": 0}),
        ]
        previous = before
        for definition, overrides, expected in variants:
            with self.subTest(expected=expected):
                write(definition, overrides)
                self.cli("check", "essay.docx", code=1)
                self.cli("export", "essay.docx")
                after = self.exported()
                self.assertEqual(after[0], before[0], "list format changes must not alter the prose")
                self.assertNotEqual(after[1], previous[1])
                self.assertEqual(self.exported_lists()[0]["levels"], [{"level": 0, **expected}])
                self.cli("check", "essay.docx")
                previous = after

    def test_style_associated_list_formats_make_metadata_stale(self) -> None:
        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="BaseList">'
            '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="NestedList"><w:basedOn w:val="BaseList"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="DerivedList"><w:basedOn w:val="NestedList"/></w:style>'
            '</w:styles>'
        )
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
        for style in ("NestedList", "DerivedList"):
            for direct in ("", '<w:numPr><w:numId w:val="1"/></w:numPr>'):
                with self.subTest(style=style, direct=direct):
                    body = paragraph("Styled item.", f'<w:pPr><w:pStyle w:val="{style}"/>{direct}</w:pPr>')

                    def write(fmt: str, body: str = body) -> None:
                        # BaseList and NestedList both match ancestors; the nearer style selects level 1.
                        self.write_lists(body, numbering(abstract_num("0", level(0, style="BaseList"),
                                                                   level(1, fmt, style="NestedList")),
                                                         num("1", "0")),
                                         extra={"word/styles.xml": styles}, relationships=relationships)

                    write("bullet")
                    self.cli("export", "essay.docx")
                    before = self.exported()
                    self.assertEqual(before[0], b"- Styled item.\n",
                                     "the associated level is recorded in metadata; prose indentation is unchanged")
                    self.assertEqual(self.exported_lists()[0]["levels"], [
                        {"level": 1, "format": "bullet", "text": "•", "start": 1},
                    ])
                    write("decimal")
                    self.cli("check", "essay.docx", code=1)
                    self.cli("export", "essay.docx")
                    after = self.exported()
                    self.assertEqual(after[0], before[0])
                    self.assertNotEqual(after[1], before[1])
                    self.assertEqual(self.exported_lists()[0]["levels"][0]["format"], "decimal")

    def test_swapping_styles_associated_with_different_levels_is_stale(self) -> None:
        # Both styles share one numId and carry no ilvl, so prose and the used levels stay the same;
        # only which item takes which level's format changes.
        styles = f'<w:styles xmlns:w="{W}">' + "".join(
            f'<w:style w:type="paragraph" w:styleId="{style}">'
            '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' for style in ("Bulleted", "Numbered")
        ) + "</w:styles>"
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'

        def write(*applied: str) -> None:
            body = "".join(paragraph(text, f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>')
                           for text, style in zip(("Alpha.", "Beta.", "Gamma."), applied, strict=True))
            self.write_lists(body, numbering(abstract_num("0", level(0, style="Bulleted"),
                                                          level(1, "decimal", "%2.", style="Numbered")),
                                             num("1", "0")),
                             extra={"word/styles.xml": styles}, relationships=relationships)

        def runs() -> list[dict[str, object]]:
            return [{"level": run["level"], "items": run["items"]} for run in json.loads(self.exported()[1])["list_runs"]]

        write("Bulleted", "Numbered", "Numbered")
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual(before[0], b"- Alpha.\n\n- Beta.\n\n- Gamma.\n")
        self.assertEqual(runs(), [{"level": 0, "items": 1}, {"level": 1, "items": 2}])
        for applied in (("Numbered", "Bulleted", "Numbered"), ("Numbered", "Numbered", "Bulleted")):
            with self.subTest(applied=applied):
                write(*applied)
                self.cli("check", "essay.docx", code=1)
                self.cli("export", "essay.docx")
                after = self.exported()
                self.assertEqual(after[0], before[0], "the swap must not alter the prose")
                self.assertEqual(json.loads(after[1])["lists"], json.loads(before[1])["lists"])
                self.assertNotEqual(after[1], before[1])
                write("Bulleted", "Numbered", "Numbered")
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported(), before)

    def test_style_associations_follow_level_overrides(self) -> None:
        # A w:lvl inside w:lvlOverride replaces the abstract level, its w:pStyle association included.
        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Clause">'
            '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>'
        )
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
        body = paragraph("Clause item.", '<w:pPr><w:pStyle w:val="Clause"/></w:pPr>')

        def override(ilvl: int, content: str) -> str:
            return f'<w:lvlOverride w:ilvl="{ilvl}">{content}</w:lvlOverride>'

        cases = {
            # The association exists only on the override of level 1.
            "override only": (abstract_num("0", level(0), level(1, "decimal", "%2.")),
                              lambda fmt: override(1, level(1, fmt, "%2.", style="Clause"))),
            # The override of level 0 drops the abstract association; the override of level 1 adds one.
            "moved by overrides": (abstract_num("0", level(0, style="Clause"), level(1, "decimal", "%2.")),
                                   lambda fmt: override(0, level(0)) + override(1, level(1, fmt, "%2.", style="Clause"))),
            # An override without w:lvl leaves the abstract level, and its association, in place.
            "start override keeps it": (abstract_num("0", level(0), level(1, "decimal", "%2.", style="Clause")),
                                        lambda fmt: override(1, '<w:startOverride w:val="1"/>')
                                        + override(2, level(2, fmt, "%3."))),
        }
        for name, (abstract, overrides) in cases.items():
            with self.subTest(name):
                def write(fmt: str, abstract: str = abstract, overrides: Any = overrides) -> None:
                    self.write_lists(body, numbering(abstract, num("1", "0", overrides(fmt))),
                                     extra={"word/styles.xml": styles}, relationships=relationships)

                write("lowerRoman")
                self.cli("export", "essay.docx")
                before = self.exported()
                self.assertEqual(before[0], b"- Clause item.\n")
                recorded = self.exported_lists()[0]["levels"]
                self.assertEqual([entry["level"] for entry in recorded], [1])
                write("upperRoman")
                if name == "start override keeps it":
                    # Level 2 is unused, so its format is not recorded.
                    self.assertEqual(recorded[0]["format"], "decimal")
                    self.cli("check", "essay.docx")
                    continue
                self.assertEqual(recorded[0]["format"], "lowerRoman")
                self.cli("check", "essay.docx", code=1)
                self.cli("export", "essay.docx")
                after = self.exported()
                self.assertEqual(after[0], before[0])
                self.assertEqual(self.exported_lists()[0]["levels"][0]["format"], "upperRoman")

    def test_direct_list_level_overrides_style_association(self) -> None:
        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="NestedList">'
            '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>'
        )
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
        body = paragraph("Explicit item.", '<w:pPr><w:pStyle w:val="NestedList"/>'
                         '<w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr>')

        def write(fmt: str) -> None:
            self.write_lists(body, numbering(abstract_num("0", level(0), level(1, fmt, style="NestedList")),
                                             num("1", "0")),
                             extra={"word/styles.xml": styles}, relationships=relationships)

        write("bullet")
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual(before[0], b"- Explicit item.\n")
        self.assertEqual(self.exported_lists()[0]["levels"][0]["level"], 0)
        write("decimal")
        self.cli("check", "essay.docx")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), before, "unused style-associated level must not churn metadata")

    def test_style_resolution_for_metadata_leaves_list_prose_unchanged(self) -> None:
        # Prose follows the nearest numPr alone, as it did before list metadata was recorded.
        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="NestedList">'
            '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="NoList">'
            '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>'
        )
        relationships = f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
        body = (
            # The style's own ilvl indents the prose; the level associated with the style is recorded.
            paragraph("Styled item.", '<w:pPr><w:pStyle w:val="NestedList"/></w:pPr>')
            # Own numPr with only ilvl while the style switches numbering off: a prose item, no list entry.
            + paragraph("Unnumbered by style.", '<w:pPr><w:pStyle w:val="NoList"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>')
        )
        self.write_lists(body, numbering(abstract_num("0", level(0), level(1, "decimal", "%2.", style="NestedList"),
                                                      level(2, "lowerRoman", "%3.")),
                                         num("1", "0")),
                         extra={"word/styles.xml": styles}, relationships=relationships)
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        self.assertEqual(prose, b"    - Styled item.\n\n  - Unnumbered by style.\n")
        self.assertEqual(self.exported_lists(), [
            {"list": 1, "part": "word/document.xml",
             "levels": [{"level": 1, "format": "decimal", "text": "%2.", "start": 1}]},
        ])
        self.assertEqual(json.loads(meta)["list_runs"],
                         [{"list": 1, "part": "word/document.xml", "level": 1, "items": 1}])

    def test_equivalent_start_spellings_do_not_churn_and_changed_starts_are_stale(self) -> None:
        body = list_item("Item.", "1")

        def write(start: str, override: str | None = None) -> None:
            overrides = "" if override is None else (
                f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="{override}"/></w:lvlOverride>')
            self.write_lists(body, numbering(abstract_num("0", level(0, "decimal", "%1.", start=start)),
                                             num("1", "0", overrides)))

        write("1")
        self.cli("export", "essay.docx")
        before = self.exported()
        # A sign, leading zeroes and surrounding whitespace are valid xsd:integer spellings.
        for start, override in (("+1", None), ("01", None), ("+001", None), (" 1 ", None), ("0" * 5000 + "1", None),
                                ("9", "+1"), ("9", "001")):
            with self.subTest(start=start[-8:], override=override):
                write(start, override)
                self.cli("check", "essay.docx")
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported(), before)
        previous = 1
        steps: list[tuple[str, str | None, int]] = [
            ("+5", None, 5), ("05", None, 5), ("1", "+3", 3), ("1", "+7", 7), ("1", "007", 7),
            ("-05", None, -5), ("-5", None, -5), ("-0", None, 0), ("+", None, 0), ("+-1", None, 0),
            ("1" + "0" * 10, None, 0), ("0" * 20 + "9" * 10, None, 9999999999),
        ]
        for start, override, expected in steps:
            with self.subTest(start=start, override=override):
                write(start, override)
                self.cli("check", "essay.docx", code=1 if expected != previous else 0)
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported()[0], before[0])
                self.assertEqual(self.exported_lists()[0]["levels"][0]["start"], expected)
                previous = expected

    def test_equivalent_level_and_id_spellings_resolve_the_same_definitions(self) -> None:
        def write(*, ilvl: str = "1", num_id: str = "1", defined_num: str = "1", abstract_ref: str = "0",
                  defined_level: str = "1", override_level: str = "1") -> None:
            body = paragraph("Item.", f'<w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>')
            definition = f'<w:lvl w:ilvl="{defined_level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl>'
            override = f'<w:lvlOverride w:ilvl="{override_level}"><w:startOverride w:val="4"/></w:lvlOverride>'
            self.write_lists(body, numbering(abstract_num("0", level(0), definition), num(defined_num, abstract_ref, override)))

        write()
        result = self.cli("export", "essay.docx")
        self.assertNotIn(UNKNOWN_LIST_NOTE, result.stderr)
        before = self.exported()
        resolved = [{"list": 1, "part": "word/document.xml",
                     "levels": [{"level": 1, "format": "decimal", "text": "%2.", "start": 4}]}]
        self.assertEqual(self.exported_lists(), resolved)
        variants: list[dict[str, str]] = [
            {"ilvl": "01"}, {"num_id": "01"}, {"num_id": "+1"}, {"defined_num": "+01"}, {"abstract_ref": "+0"},
            {"abstract_ref": "00"}, {"defined_level": "01"}, {"defined_level": "+1"}, {"override_level": "+01"},
        ]
        for variant in variants:
            with self.subTest(variant=variant):
                write(**variant)
                self.cli("check", "essay.docx")
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported(), before)
        # Prose indentation is not re-read here: only the recorded level follows the signed spelling.
        write(ilvl="+1")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported_lists(), resolved)

    def test_renumbered_ids_and_unused_definitions_do_not_churn(self) -> None:
        body = list_item("Bullet.", "1") + list_item("Nested.", "1", ilvl=1) + list_item("Numbered.", "2")
        self.write_lists(body, numbering(
            abstract_num("0", level(0), level(1, "bullet", "o")),
            abstract_num("1", level(0, "decimal", "%1.")),
            num("1", "0"), num("2", "1"),
        ))
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual([entry["list"] for entry in self.exported_lists()], [1, 2])
        renumbered = list_item("Bullet.", "12") + list_item("Nested.", "12", ilvl=1) + list_item("Numbered.", "4")
        # New IDs, reordered definitions, an unused list, and a changed level no item uses.
        self.write_lists(renumbered, numbering(
            abstract_num("9", level(0, "decimal", "%1.")),
            abstract_num("7", level(0), level(1, "bullet", "o"), level(2, "upperRoman", "%3:", "8")),
            abstract_num("3", level(0, "lowerLetter", "%1)")),
            num("4", "9"), num("5", "3"), num("12", "7"),
        ))
        self.cli("check", "essay.docx")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), before)

    def test_list_ordinals_follow_first_use_across_parts(self) -> None:
        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
            '<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style></w:styles>'
        )
        extra = {
            "word/styles.xml": styles,
            "word/header1.xml": f'<w:hdr xmlns:w="{W}">{list_item("Header item.", "7")}{list_item("Header only.", "5")}</w:hdr>',
        }
        relationships = (
            f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
        )
        body = (
            list_item("Second list first.", "7", ilvl=2)
            + paragraph("Styled item.", '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>')
            # Own numPr with only ilvl: numId comes from the style.
            + paragraph("Styled sub-item.", '<w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>')
            + list_item("Back to the first.", "7")
            + paragraph("", '<w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr>')  # empty: not rendered
            + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>'
        )
        self.write_lists(body, numbering(
            abstract_num("0", level(0, "decimal", "%1."), level(2, "lowerRoman", "%3.")),
            abstract_num("1", level(0), level(1, "bullet", "o")),
            num("3", "1"), num("5", "1"), num("7", "0"), num("9", "1"),
        ), extra=extra, relationships=relationships)
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported_lists(), [
            {"list": 1, "part": "word/document.xml", "levels": [
                {"level": 0, "format": "decimal", "text": "%1.", "start": 1},
                {"level": 2, "format": "lowerRoman", "text": "%3.", "start": 1},
            ]},
            {"list": 2, "part": "word/document.xml", "levels": [
                {"level": 0, "format": "bullet", "text": "•", "start": 1},
                {"level": 1, "format": "bullet", "text": "o", "start": 1},
            ]},
            {"list": 3, "part": "word/header1.xml",
             "levels": [{"level": 0, "format": "bullet", "text": "•", "start": 1}]},
        ])
        self.assertNotIn("word/numbering.xml", self.exported()[1].decode("utf-8"))
        self.assertEqual(json.loads(self.exported()[1])["list_runs"], [
            {"list": 1, "part": "word/document.xml", "level": 2, "items": 1},
            {"list": 2, "part": "word/document.xml", "level": 0, "items": 1},
            {"list": 2, "part": "word/document.xml", "level": 1, "items": 1},
            {"list": 1, "part": "word/document.xml", "level": 0, "items": 1},
            {"list": 1, "part": "word/header1.xml", "level": 0, "items": 1},
            {"list": 3, "part": "word/header1.xml", "level": 0, "items": 1},
        ])

    def test_list_in_a_text_box_counts_after_its_anchoring_item(self) -> None:
        outer = list_item("Anchor item.", "4").replace("</w:r></w:p>", f"</w:r><w:r>{modern_text_box(list_item('Boxed item.', '2'))}</w:r></w:p>")
        self.write_lists(outer, numbering(abstract_num("0", level(0)), abstract_num("1", level(0, "decimal", "%1.")),
                                          num("2", "1"), num("4", "0")))
        self.cli("export", "essay.docx")
        prose = self.exported()[0].decode("utf-8")
        self.assertLess(prose.index("Anchor item."), prose.index("Boxed item."))
        self.assertEqual([entry["levels"][0]["format"] for entry in self.exported_lists()], ["bullet", "decimal"])

    def test_moving_an_item_between_existing_lists_is_stale(self) -> None:
        definitions = numbering(abstract_num("0", level(0)), abstract_num("1", level(0, "decimal", "%1.")),
                                num("1", "0"), num("2", "1"))

        def write(*ids: str) -> None:
            texts = ("Alpha.", "Beta.", "Gamma.", "Delta.")
            body = paragraph("Intro.") + "".join(list_item(text, num_id) for text, num_id in zip(texts, ids, strict=True))
            self.write_lists(body, definitions)

        write("1", "2", "1", "2")
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual(json.loads(before[1])["list_runs"], [
            {"list": 1, "part": "word/document.xml", "level": 0, "items": 1},
            {"list": 2, "part": "word/document.xml", "level": 0, "items": 1},
            {"list": 1, "part": "word/document.xml", "level": 0, "items": 1},
            {"list": 2, "part": "word/document.xml", "level": 0, "items": 1},
        ])
        for ids in (("1", "2", "2", "2"), ("1", "2", "2", "1")):
            with self.subTest(ids=ids):
                write(*ids)
                self.cli("check", "essay.docx", code=1)
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported()[0], before[0])
                write("1", "2", "1", "2")
                self.cli("export", "essay.docx")
                self.assertEqual(self.exported(), before)

    def test_custom_formats_notes_cells_and_odd_values(self) -> None:
        def custom(pattern: str) -> str:
            return (
                f'<w:lvl w:ilvl="0"><w:start w:val="1"/><mc:AlternateContent xmlns:mc="{MC}">'
                f'<mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="{pattern}"/></mc:Choice>'
                '<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>'
                '<w:lvlText w:val="%1."/></w:lvl>'
            )

        styles = (
            f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="NoList"><w:name w:val="No List"/>'
            '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>'
        )
        extra = {
            "word/styles.xml": styles,
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1">{list_item("Noted item.", "3")}</w:footnote></w:footnotes>',
        }
        relationships = (
            f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>'
        )
        body = (
            '<w:tbl><w:tr><w:tc>' + list_item("Cell item.", "1") + '</w:tc></w:tr></w:tbl>'
            + paragraph("Unnumbered by style.", '<w:pPr><w:pStyle w:val="NoList"/><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr>')
            + list_item("Odd start.", "2")
            + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
        )

        def write(pattern: str) -> None:
            self.write_lists(body, numbering(
                abstract_num("0", custom(pattern)),
                abstract_num("1", level(0, "decimal", "%1.", start="9" * 5000)),
                abstract_num("2", level(0, "upperLetter", "%1.", start="two")),
                num("1", "0"), num("2", "1"), num("3", "2"),
            ), extra=extra, relationships=relationships)

        write("001, 002, 003, ...")
        result = self.cli("export", "essay.docx")
        self.assertNotIn(UNKNOWN_LIST_NOTE, result.stderr)
        self.assertEqual(self.exported_lists(), [
            {"list": 1, "part": "word/document.xml",
             "levels": [{"level": 0, "format": "custom: 001, 002, 003, ...", "text": "%1.", "start": 1}]},
            {"list": 2, "part": "word/document.xml",
             "levels": [{"level": 0, "format": "decimal", "text": "%1.", "start": 0}]},
            {"list": 3, "part": "word/footnotes.xml",
             "levels": [{"level": 0, "format": "upperLetter", "text": "%1.", "start": 0}]},
        ])
        write("01, 02, 03, ...")
        self.cli("check", "essay.docx", code=1)

    def test_numbering_part_without_a_relationship_is_unrendered(self) -> None:
        make_docx(self.root / "essay.docx", list_item("Item.", "1"),
                  extra={"word/numbering.xml": numbering(abstract_num("0", level(0)), num("1", "0"))})
        result = self.cli("export", "essay.docx")
        self.assertIn("Unrendered package part: word/numbering.xml", result.stderr)
        self.assertIn(UNKNOWN_LIST_NOTE, result.stderr)

    def test_unresolvable_lists_are_unknown_with_a_warning(self) -> None:
        unknown = {"level": 0, "format": "unknown", "text": "", "start": 0}
        cases = {
            "no numbering part": (list_item("Item.", "1"), None),
            "unknown numId": (list_item("Item.", "8"), numbering(abstract_num("0", level(0)), num("1", "0"))),
            "unknown abstractNum": (list_item("Item.", "1"), numbering(abstract_num("0", level(0)), num("1", "6"))),
            "unknown level": (list_item("Item.", "1", ilvl=3), numbering(abstract_num("0", level(0)), num("1", "0"))),
            "numStyleLink": (list_item("Item.", "1"), numbering(
                '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="Bullets"/></w:abstractNum>', num("1", "0"))),
        }
        for name, (body, numbering_xml) in cases.items():
            with self.subTest(name):
                self.write_lists(body, numbering_xml)
                result = self.cli("export", "essay.docx")
                self.assertIn(UNKNOWN_LIST_NOTE, result.stderr)
                level_no = 3 if name == "unknown level" else 0
                self.assertEqual(self.exported_lists(), [
                    {"list": 1, "part": "word/document.xml", "levels": [{**unknown, "level": level_no}]},
                ])
        self.write_lists(list_item("Item.", "1"), None, relationships=NUMBERING_REL)
        self.cli("export", "essay.docx", code=2)

    def test_documents_without_lists_record_an_empty_array(self) -> None:
        self.write_lists(paragraph("Plain prose."), None)
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported_lists(), [])
        self.assertEqual(json.loads(self.exported()[1])["list_runs"], [])
        removed = paragraph("No longer a list.", '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>')
        self.write_lists(removed, numbering(abstract_num("0", level(0)), num("1", "0")))
        result = self.cli("export", "essay.docx")
        self.assertEqual(self.exported_lists(), [])
        self.assertNotIn(LIST_NOTE, result.stderr)
        self.assertNotIn("numbering.xml", result.stderr)

    def test_notes_headers_and_footers_are_exported(self) -> None:
        extra = {
            "word/header1.xml": f'<w:hdr xmlns:w="{W}">{paragraph("Header words.")}</w:hdr>',
            "word/footer1.xml": f'<w:ftr xmlns:w="{W}">{paragraph("Footer words.")}</w:ftr>',
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1">{paragraph("Footnote words.")}</w:footnote></w:footnotes>',
            "word/endnotes.xml": f'<w:endnotes xmlns:w="{W}"><w:endnote w:id="2">{paragraph("Endnote words.")}</w:endnote></w:endnotes>',
        }
        relationships = (
            f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
            f'<Relationship Id="footerRef" Type="{R}/footer" Target="footer1.xml"/>'
            f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>'
            f'<Relationship Id="endnotes" Type="{R}/endnotes" Target="endnotes.xml"/>'
        )
        body = paragraph("Body words.") + '<w:p><w:r><w:footnoteReference w:id="1"/><w:endnoteReference w:id="2"/></w:r></w:p>'
        body += '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/><w:footerReference w:type="default" r:id="footerRef"/></w:sectPr>'
        make_docx(self.root / "essay.docx", body, extra=extra, relationships=relationships)
        self.cli("export", "essay.docx", "--root", self.root)
        prose = self.exported()[0].decode("utf-8")
        for text in ("Body words.", "Header words.", "Footer words.", "Footnote words.", "Endnote words."):
            self.assertIn(text, prose)

    def test_accepted_tracked_changes_and_review_metadata(self) -> None:
        body = paragraph("Unchanged.") + (
            '<w:p><w:ins w:id="1" w:author="Ada" w:date="2024-01-02T03:04:05Z"><w:r><w:t>Inserted words.</w:t></w:r></w:ins>'
            '<w:del w:id="2" w:author="Ben" w:date="2024-01-03T03:04:05Z"><w:r><w:delText>Deleted words.</w:delText></w:r></w:del>'
            '<w:moveTo w:id="3" w:author="Cara"><w:r><w:t>Moved here.</w:t></w:r></w:moveTo>'
            '<w:moveFrom w:id="4" w:author="Dan"><w:r><w:delText>Moved away.</w:delText></w:r></w:moveFrom></w:p>'
        )
        make_docx(self.root / "essay.docx", body)
        self.cli("export", "essay.docx", "--root", self.root)
        prose_bytes, meta_bytes = self.exported()
        prose, meta = prose_bytes.decode("utf-8"), meta_bytes.decode("utf-8")
        self.assertIn("Inserted words.", prose)
        self.assertIn("Moved here.", prose)
        self.assertNotIn("Deleted words.", prose)
        self.assertNotIn("Moved away.", prose)
        for text in ("Inserted words.", "Deleted words.", "Moved here.", "Moved away.", "Ada", "Ben", "Cara", "Dan"):
            self.assertIn(text, meta)
        self.assertEqual({item["kind"] for item in json.loads(meta)["review"]}, {"ins", "del", "moveTo", "moveFrom"})
        self.assertNotIn("2024-01-", meta)

    def test_review_dates_are_ignored_but_author_changes_are_stale(self) -> None:
        source = self.root / "essay.docx"

        def write_review(author: str, date: str) -> None:
            make_docx(source, f'<w:p><w:ins w:id="5" w:author="{author}" w:date="{date}"><w:r><w:t>Reviewed prose.</w:t></w:r></w:ins></w:p>')

        write_review("Ada", "2024-01-02T03:04:05Z")
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        write_review("Ada", "2030-12-31T23:59:59Z")
        self.cli("check", source, "--root", self.root)
        self.cli("export", source, "--root", self.root)
        self.assertEqual(before, self.exported())
        write_review("Grace", "2030-12-31T23:59:59Z")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertNotEqual(before[1], after[1])

    def test_saved_field_display_is_prose_and_instruction_changes_are_stale(self) -> None:
        source = self.root / "essay.docx"

        def write_field(identifier: str) -> None:
            body = (
                '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                f'<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {identifier} </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                '<w:r><w:t>(Smith, 2024)</w:t></w:r>'
                '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            )
            make_docx(source, body)

        write_field("citation-one")
        result = self.cli("export", source, "--root", self.root)
        self.assertRegex(result.stdout + result.stderr, r"(?i)field|saved display|zotero")
        before = self.exported()
        self.assertIn(b"(Smith, 2024)", before[0])
        self.assertNotIn(b"ADDIN ZOTERO_ITEM", before[0])
        write_field("citation-two")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertNotEqual(before[1], after[1])

    def export_form(self, body: str, **options: Any) -> dict[str, Any]:
        source = self.root / "essay.docx"
        make_docx(source, body, **options)
        self.cli("export", source, "--root", self.root)
        meta: dict[str, Any] = json.loads(self.exported()[1])
        return meta

    def test_toggling_a_form_checkbox_is_stale_in_metadata_only(self) -> None:
        source = self.root / "essay.docx"
        meta = self.export_form(f"<w:p>{run('The Tenant may sublet: ')}{checkbox('0')}</w:p>")
        self.assertEqual(meta["fields"], [checkbox_entry(False)])
        self.assertIn(FORM_FIELD_NOTE, meta["warnings"])
        before = self.exported()
        make_docx(source, f"<w:p>{run('The Tenant may sublet: ')}{checkbox('1')}</w:p>")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertEqual(json.loads(after[1])["fields"], [checkbox_entry(True)])

    def test_checkbox_state_follows_checked_then_default(self) -> None:
        cases = [
            (checkbox("1", "0"), True), (checkbox("0", "1"), False), (checkbox("", "0"), True),
            (checkbox("true", None), True), (checkbox("off", None), False),
            (checkbox(None, "1"), True), (checkbox(None, "0"), False), (checkbox(None, ""), True),
            (checkbox(None, None), False), (checkbox("0", name=None), False),
        ]
        for xml, checked in cases:
            with self.subTest(xml=xml):
                meta = self.export_form(f"<w:p>{xml}</w:p>")
                self.assertEqual(meta["fields"][0]["form"]["checked"], checked)
        self.assertEqual(meta["fields"][0]["form"]["name"], "")
        # w:default is the state when w:checked is absent, so changing it is stale too.
        source = self.root / "essay.docx"
        make_docx(source, f"<w:p>{checkbox(None, '1')}</w:p>")
        self.cli("check", source, "--root", self.root, code=1)

    def test_choosing_another_dropdown_entry_is_stale_in_metadata_only(self) -> None:
        source = self.root / "essay.docx"
        meta = self.export_form(f"<w:p>{run('Sublet allowed: ')}{dropdown('Yes', 'No', result='0')}</w:p>")
        self.assertEqual(meta["fields"][0]["form"], dropdown_form(["Yes", "No"], "Yes"))
        self.assertIn(FORM_FIELD_NOTE, meta["warnings"])
        before = self.exported()
        make_docx(source, f"<w:p>{run('Sublet allowed: ')}{dropdown('Yes', 'No', result='1')}</w:p>")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertEqual(json.loads(after[1])["fields"][0]["form"], dropdown_form(["Yes", "No"], "No"))
        # Changing the entries, even with the same selection, is stale too.
        make_docx(source, f"<w:p>{run('Sublet allowed: ')}{dropdown('Yes', 'No', 'Maybe', result='1')}</w:p>")
        self.cli("check", source, "--root", self.root, code=1)

    def test_dropdown_selection_is_null_when_it_names_no_entry(self) -> None:
        cases = [
            (dropdown("Yes", "No", result=None), "Yes"),  # no w:result: the default entry, 0
            (dropdown("Yes", "No", result=None, default="1"), "No"),  # no w:result: w:default
            (dropdown("Yes", "No", result="1", default="0"), "No"),  # w:result wins over w:default
            (dropdown("Yes", "No", result="+01"), "No"),
            (dropdown("Yes", "No", result="2"), None),
            (dropdown("Yes", "No", result="-1"), None),
            (dropdown("Yes", "No", result="first"), None),
            (dropdown("Yes", "No", result="99999999999999999999"), None),
            (dropdown("Yes", "No", result=None, default="5"), None),
            (dropdown(result=None), None),
            (dropdown(result="0"), None),
        ]
        for xml, selected in cases:
            with self.subTest(xml=xml):
                form = self.export_form(f"<w:p>{xml}</w:p>")["fields"][0]["form"]
                self.assertEqual(form["selected"], selected)
        self.assertEqual(form["entries"], [])

    def test_form_field_save_churn_keeps_companions_byte_identical(self) -> None:
        source = self.root / "essay.docx"
        self.export_form(f"<w:p>{checkbox('1')}{dropdown('Yes', 'No', result='1')}</w:p>")
        before = self.exported()
        # Reordered w:ffData children, an instruction split across runs, rsid attributes,
        # proofing marks and a bookmark between the field's runs.
        reordered = checkbox("1", rsid="00AB12CD").replace(
            '<w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/>',
            '<w:calcOnExit w:val="0"/><w:enabled/><w:name w:val="Check1"/>')
        reordered = reordered.replace(
            '<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="1"/></w:checkBox>',
            '<w:checkBox><w:checked w:val="1"/><w:default w:val="0"/><w:sizeAuto/></w:checkBox>')
        split = dropdown("Yes", "No", result="1").replace(
            '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>',
            '<w:r><w:instrText xml:space="preserve"> FORM</w:instrText></w:r><w:proofErr w:type="spellStart"/>'
            '<w:bookmarkStart w:id="3" w:name="Drop1"/>'
            '<w:r w:rsidR="0011AA22"><w:instrText>DROPDOWN </w:instrText></w:r>')
        split = split.replace('<w:result w:val="1"/><w:listEntry', '<w:listEntry').replace(
            '<w:listEntry w:val="No"/>', '<w:listEntry w:val="No"/><w:result w:val="1"/>')
        self.assertNotEqual(split, dropdown("Yes", "No", result="1"))
        make_docx(source, f'<w:p w:rsidR="00AB12CD">{reordered}{split}</w:p>', churn=True)
        self.cli("check", source, "--root", self.root)
        self.cli("export", source, "--root", self.root)
        self.assertEqual(before, self.exported())

    def test_fields_without_form_data_keep_their_metadata(self) -> None:
        # Pins the metadata of ordinary fields (a PAGE field and a Zotero citation): it must not
        # change, so documents without legacy form fields keep their companions.
        body = (
            '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            '<w:p><w:fldSimple w:instr=" ADDIN ZOTERO_ITEM CSL_CITATION {} "><w:r><w:t>(Smith, 2024)</w:t></w:r>'
            '</w:fldSimple></w:p>'
        )
        meta = self.export_form(body)
        expected = {
            "schema": 1, "view": "accept-insertions-omit-deletions", "coverage": "partial-prose",
            "warnings": [
                "Fields use saved display text, which can be stale; codes are fingerprinted, not evaluated (including Zotero).",
                "Formatting and page layout are not compared; unchanged prose is not an identical document.",
            ],
            "review": [], "hyperlinks": [], "lists": [], "list_runs": [], "images": [],
            "fields": [
                {"command": "PAGE", "sha256": hashlib.sha256(b"PAGE").hexdigest(), "part": "word/document.xml"},
                {"command": "ADDIN", "sha256": hashlib.sha256(b"ADDIN ZOTERO_ITEM CSL_CITATION {}").hexdigest(),
                 "part": "word/document.xml"},
            ],
        }
        self.assertEqual(meta, expected)
        self.assertEqual(self.exported()[1], (json.dumps(expected, ensure_ascii=False, indent=2,
                                                         sort_keys=True) + "\n").encode("utf-8"))

    def test_form_fields_in_other_stories_and_tracked_changes(self) -> None:
        extra = {
            "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p>{checkbox("1", name="InHeader")}</w:p></w:hdr>',
            "word/footnotes.xml": (f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>{checkbox("0", name="InNote")}'
                                   '</w:p></w:footnote></w:footnotes>'),
        }
        relationships = (f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
                         f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        boxed = f"<w:p>{checkbox('1', name='InBox')}</w:p>"
        body = (anchored("Boxed:", modern_text_box(boxed))
                + f'<w:p><w:ins w:id="1" w:author="Ben">{checkbox("1", name="Inserted")}</w:ins></w:p>'
                + f'<w:p>{run("Kept.")}<w:del w:id="2" w:author="Ben">{checkbox("1", name="Deleted")}</w:del></w:p>'
                + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
                + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>')
        meta = self.export_form(body, extra=extra, relationships=relationships)
        # Like every field, one inside a tracked deletion is not recorded; an insertion's is.
        self.assertEqual(meta["fields"], [
            checkbox_entry(True, "InBox"), checkbox_entry(True, "Inserted"),
            checkbox_entry(False, "InNote", "word/footnotes.xml"), checkbox_entry(True, "InHeader", "word/header1.xml"),
        ])

    def test_recorded_form_data_no_longer_warns_but_text_inputs_still_do(self) -> None:
        behaviour = ('<w:entryMacro w:val="Go"/><w:exitMacro w:val="Done"/><w:helpText w:type="text" w:val="Help"/>'
                     '<w:statusText w:type="text" w:val="Status"/>')
        text_input = form_field("FORMTEXT", '<w:textInput><w:type w:val="number"/><w:default w:val="0"/>'
                                '<w:maxLength w:val="4"/><w:format w:val="0"/></w:textInput>', "1200", name="Rent")
        body = (f"<w:p>{checkbox('1', '0').replace('<w:enabled/>', '<w:enabled/>' + behaviour)}"
                f"{dropdown('Yes', 'No', result='1', default='0')}</w:p><w:p>{run('Rent: ')}{text_input}</w:p>")
        meta = self.export_form(body)
        unsupported = sorted(w for w in meta["warnings"] if w.startswith("Unsupported Word construct"))
        self.assertEqual(unsupported, [f"Unsupported Word construct {name}; its semantics are not fully compared."
                                       for name in ("default", "format", "maxLength", "textInput", "type")])
        # The text input's value is prose; it records no form state.
        self.assertIn(b"Rent: 1200", self.exported()[0])
        self.assertNotIn("form", meta["fields"][2])
        self.assertEqual([f.get("form", {}).get("type") for f in meta["fields"]], ["checkbox", "dropdown", None])

    def test_hyperlink_target_changes_are_metadata_changes(self) -> None:
        body = '<w:p><w:hyperlink r:id="link"><w:r><w:t>Visible link label</w:t></w:r></w:hyperlink></w:p>'
        source = self.root / "essay.docx"

        def write_link(target: str) -> None:
            make_docx(source, body, relationships=f'<Relationship Id="link" Type="{R}/hyperlink" Target="{target}" TargetMode="External"/>')

        write_link("https://example.invalid/first")
        self.cli("export", source, "--root", self.root)
        before = self.exported()
        self.assertIn(b"https://example.invalid/first", before[1])
        write_link("https://example.invalid/second")
        self.cli("check", source, "--root", self.root, code=1)
        self.cli("export", source, "--root", self.root)
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertNotEqual(before[1], after[1])
        self.assertIn(b"https://example.invalid/second", after[1])

    def test_known_drawing_omission_warns_but_export_succeeds(self) -> None:
        body = paragraph("Surrounding prose.") + '<w:p><w:r><w:drawing/></w:r></w:p>'
        make_docx(self.root / "essay.docx", body)
        result = self.cli("export", "essay.docx", "--root", self.root)
        self.assertRegex(result.stdout + result.stderr, r"(?i)drawing")
        self.assertIn(b"Surrounding prose.", self.exported()[0])
        self.cli("check", "essay.docx", "--root", self.root)


    def export_text_box(self, body: str, **options: Any) -> tuple[str, dict[str, Any]]:
        make_docx(self.root / "essay.docx", body, **options)
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        return prose.decode("utf-8"), json.loads(meta)

    def assert_only_text_box_warning(self, meta: dict[str, Any], *extra: str) -> None:
        self.assertEqual(meta["warnings"], sorted([
            "Formatting and page layout are not compared; unchanged prose is not an identical document.",
            TEXT_BOX_NOTE, *extra,
        ]))

    def test_modern_text_box_text_is_rendered_once_after_its_anchor(self) -> None:
        body = paragraph("Before.") + anchored("Anchor text.", modern_text_box(paragraph("Boxed sentence. Second one."))) + paragraph("After.")
        prose, meta = self.export_text_box(body)
        self.assertEqual(prose, "Before.\n\nAnchor text.\n\n[Text box]\n\nBoxed sentence.\nSecond one.\n\n[/Text box]\n\nAfter.\n")
        # One small, stable warning set rather than one per DrawingML/VML element.
        self.assert_only_text_box_warning(meta)

    def test_vml_text_box_is_rendered_like_a_modern_one(self) -> None:
        body = paragraph("Before.") + anchored("Anchor text.", vml_text_box(paragraph("Boxed sentence. Second one."))) + paragraph("After.")
        prose, meta = self.export_text_box(body)
        self.assertEqual(prose, "Before.\n\nAnchor text.\n\n[Text box]\n\nBoxed sentence.\nSecond one.\n\n[/Text box]\n\nAfter.\n")
        self.assert_only_text_box_warning(meta)

    def test_choice_copy_is_preferred_and_fallback_used_only_without_one(self) -> None:
        body = anchored("", modern_text_box(paragraph("Choice copy."), fallback=paragraph("Fallback copy.")))
        prose, _ = self.export_text_box(body)
        self.assertIn("Choice copy.", prose)
        self.assertNotIn("Fallback copy.", prose)
        # A Choice holding only a picture defers to the Fallback text box.
        picture = f'<w:drawing><wp:inline xmlns:wp="{WP}"><a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}"/></a:graphicData></a:graphic></wp:inline></w:drawing>'
        alternate = f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="wps">{picture}</mc:Choice><mc:Fallback>{vml_text_box(paragraph("Fallback copy."))}</mc:Fallback></mc:AlternateContent>'
        prose, meta = self.export_text_box(anchored("", alternate))
        self.assertEqual(prose, "[Text box]\n\nFallback copy.\n\n[/Text box]\n")
        self.assert_only_text_box_warning(meta)

    def test_text_box_edit_is_stale_until_exported_again(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, anchored("Stable anchor.", modern_text_box(paragraph("Original callout."))))
        self.cli("export", "essay.docx")
        before = self.exported()
        make_docx(source, anchored("Stable anchor.", modern_text_box(paragraph("Revised callout."))))
        self.cli("check", "essay.docx", code=1)
        self.assertEqual(self.exported(), before, "check must not regenerate stale files")
        self.cli("export", "essay.docx")
        after = self.exported()
        self.assertEqual(after[0], before[0].replace(b"Original callout.", b"Revised callout."))
        self.cli("check", "essay.docx")

    def test_tracked_changes_in_a_text_box_are_reviewed_once(self) -> None:
        content = (
            '<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>Inserted words.</w:t></w:r></w:ins>'
            '<w:del w:id="2" w:author="Ben"><w:r><w:delText>Deleted words.</w:delText></w:r></w:del></w:p>'
        )
        body = f'<w:p><w:ins w:id="3" w:author="Cara"><w:r>{modern_text_box(content)}</w:r></w:ins></w:p>'
        prose, meta = self.export_text_box(body)
        self.assertEqual(prose, "[Text box]\n\nInserted words.\n\n[/Text box]\n")
        review = [(item["kind"], item["author"], item["text"]) for item in meta["review"]]
        self.assertEqual(review, [
            ("ins", "Cara", "Inserted words.Deleted words."),
            ("ins", "Ada", "Inserted words."),
            ("del", "Ben", "Deleted words."),
        ])

    def test_text_boxes_in_cells_headers_and_nested_boxes(self) -> None:
        inner = modern_text_box(paragraph("Inner box."))
        outer = modern_text_box(paragraph("Outer box.") + anchored("Outer anchor.", inner))
        body = '<w:tbl><w:tr><w:tc>' + anchored("Cell anchor.", modern_text_box(paragraph("Cell box."))) + '</w:tc></w:tr></w:tbl>'
        body += anchored("Body anchor.", outer) + anchored("Two boxes.", modern_text_box(paragraph("First.")) + vml_text_box(paragraph("Second.")))
        body += '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>'
        header = f'<w:hdr xmlns:w="{W}" xmlns:r="{R}">{anchored("Header anchor.", vml_text_box(paragraph("Header box.")))}</w:hdr>'
        prose, meta = self.export_text_box(
            body, extra={"word/header1.xml": header},
            relationships=f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>',
        )
        self.assertIn("[Cell]\n\nCell anchor.\n\n[Text box]\n\nCell box.\n\n[/Text box]\n\n[/Row]", prose)
        self.assertIn(
            "Body anchor.\n\n[Text box]\n\nOuter box.\n\nOuter anchor.\n\n[Text box]\n\nInner box.\n\n[/Text box]\n\n[/Text box]",
            prose,
        )
        self.assertIn("Two boxes.\n\n[Text box]\n\nFirst.\n\n[/Text box]\n\n[Text box]\n\nSecond.\n\n[/Text box]", prose)
        self.assertIn("## Header (word/header1.xml)\n\nHeader anchor.\n\n[Text box]\n\nHeader box.\n\n[/Text box]", prose)
        for text in ("Cell box.", "Outer box.", "Inner box.", "First.", "Second.", "Header box."):
            self.assertEqual(prose.count(text), 1, text)
        self.assertEqual(prose.count("[Text box]"), prose.count("[/Text box]"))
        self.assertNotIn("Unsupported", prose)
        self.assert_only_text_box_warning(meta, "Headers and footers are shown by part; page/section placement is not reconstructed.")

    def test_shapes_without_text_keep_placeholders(self) -> None:
        picture = f'<w:drawing><wp:inline xmlns:wp="{WP}"><a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}"/></a:graphicData></a:graphic></wp:inline></w:drawing>'
        prose, meta = self.export_text_box(anchored("Picture anchor.", picture))
        self.assertEqual(prose, "Picture anchor.[Unsupported drawing]\n")
        self.assertIn("Unsupported drawing: content is not rendered in prose; inspect the DOCX in Word.", meta["warnings"])
        self.assertNotIn(TEXT_BOX_NOTE, meta["warnings"])
        # A picture grouped with a text box is still reported; the text box text is shown.
        prose, meta = self.export_text_box(anchored("Group anchor.", modern_text_box(paragraph("Grouped box."), picture=True)))
        self.assertEqual(prose, "Group anchor.[Unsupported drawing]\n\n[Text box]\n\nGrouped box.\n\n[/Text box]\n")
        self.assertIn("Unsupported drawing: content is not rendered in prose; inspect the DOCX in Word.", meta["warnings"])
        self.assertIn(TEXT_BOX_NOTE, meta["warnings"])

    def test_block_level_alternate_content_is_rendered_once(self) -> None:
        choice = anchored("Choice anchor.", drawing_text_box(paragraph("Choice box.")))
        fallback = anchored("Fallback anchor.", vml_text_box(paragraph("Fallback box.")))
        body = paragraph("Before.") + f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="wps">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>'
        prose, meta = self.export_text_box(body)
        self.assertEqual(prose, "Before.\n\nChoice anchor.\n\n[Text box]\n\nChoice box.\n\n[/Text box]\n")
        self.assert_only_text_box_warning(meta)

    def test_grouped_picture_in_a_compatibility_wrapper_keeps_its_placeholder(self) -> None:
        picture = f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="pic"><pic:pic xmlns:pic="{PIC}"/></mc:Choice><mc:Fallback/></mc:AlternateContent>'
        shape = drawing_text_box(paragraph("Grouped box.")).replace("</wps:wsp>", "</wps:wsp>" + picture)
        prose, meta = self.export_text_box(anchored("Group anchor.", shape))
        self.assertEqual(prose, "Group anchor.[Unsupported drawing]\n\n[Text box]\n\nGrouped box.\n\n[/Text box]\n")
        self.assertIn("Unsupported drawing: content is not rendered in prose; inspect the DOCX in Word.", meta["warnings"])

    def export_images(self, body: str, **options: Any) -> tuple[str, dict[str, Any], str]:
        make_docx(self.root / "essay.docx", body, **options)
        result = self.cli("export", "essay.docx")
        prose, meta = self.exported()
        return prose.decode("utf-8"), json.loads(meta), result.stderr

    def test_image_swaps_are_stale_in_metadata_only_and_renames_are_not(self) -> None:
        body = paragraph("Caption above.") + anchored("Figure:", picture("rId5"))

        def write(name: str, data: str, rid: str = "rId5") -> None:
            make_docx(self.root / "essay.docx", body.replace("rId5", rid),
                      extra={f"word/media/{name}": data}, relationships=image_rel(rid, f"media/{name}"))

        write("image1.png", "first picture bytes")
        self.cli("export", "essay.docx")
        before = self.exported()
        meta = json.loads(before[1])
        self.assertEqual(before[0], b"Caption above.\n\nFigure:[Unsupported drawing]\n")
        self.assertEqual(meta["images"], [image_entry("first picture bytes")])
        self.assertIn(IMAGE_NOTE, meta["warnings"])
        self.assertFalse([wn for wn in meta["warnings"] if "Unrendered package part" in wn], meta["warnings"])
        self.assertNotIn(b"image1.png", before[1], "media file names churn on save and are not recorded")
        write("image1.png", "second picture bytes")
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        after = self.exported()
        self.assertEqual(after[0], before[0], "image swaps must not alter the prose")
        changed = json.loads(after[1])
        self.assertEqual(changed["images"], [image_entry("second picture bytes")])
        self.assertEqual({k: v for k, v in changed.items() if k != "images"},
                         {k: v for k, v in meta.items() if k != "images"})
        # Word renumbers media parts and relationship IDs on save; identical bytes stay fresh.
        write("image7.png", "second picture bytes", rid="rId12")
        self.cli("check", "essay.docx")

    def test_images_in_headers_notes_text_boxes_and_vml_are_recorded_in_order(self) -> None:
        header = f'<w:hdr xmlns:w="{W}" xmlns:r="{R}">{anchored("Logo:", picture("logo"))}</w:hdr>'
        footnotes = (f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}"><w:footnote w:id="1">'
                     f'{anchored("Note figure:", picture("noteImage"))}</w:footnote></w:footnotes>')
        boxed = paragraph("Boxed caption.") + anchored("", picture("boxed"))
        # A compatibility wrapper with no text box: only the mc:Choice copy is recorded.
        wrapped = f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="wps">{picture("choice")}</mc:Choice><mc:Fallback>{vml_picture("fallback")}</mc:Fallback></mc:AlternateContent>'
        body = (anchored("Body:", picture("body") + modern_text_box(boxed)) + anchored("Legacy:", vml_picture("legacy"))
                + anchored("Wrapped:", wrapped) + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
                + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>')
        media = {name: f"{name} bytes" for name in ("body", "boxed", "legacy", "choice", "fallback", "logo", "note")}
        extra = {f"word/media/{name}.png": data for name, data in media.items()}
        extra.update({
            "word/header1.xml": header, "word/footnotes.xml": footnotes,
            "word/_rels/header1.xml.rels": f'<Relationships xmlns="{PKG_R}">{image_rel("logo", "media/logo.png")}</Relationships>',
            "word/_rels/footnotes.xml.rels": f'<Relationships xmlns="{PKG_R}">{image_rel("noteImage", "media/note.png")}</Relationships>',
        })
        relationships = "".join(image_rel(rid, f"media/{rid}.png") for rid in ("body", "boxed", "legacy", "choice", "fallback"))
        relationships += (f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
                          f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        prose, meta, stderr = self.export_images(body, extra=extra, relationships=relationships)
        self.assertEqual(meta["images"], [
            image_entry(media["body"]), image_entry(media["boxed"]), image_entry(media["legacy"]),
            image_entry(media["choice"]), image_entry(media["note"], "word/footnotes.xml"),
            image_entry(media["logo"], "word/header1.xml"),
        ])
        self.assertIn("[Text box]\n\nBoxed caption.\n\n[Unsupported drawing]\n\n[/Text box]", prose)
        self.assertIn("Legacy:[Unsupported pict]", prose)
        self.assertIn("Wrapped:[Unsupported AlternateContent]", prose)
        # Only the unread Fallback copy is still an unrendered part.
        unrendered = [wn for wn in meta["warnings"] if "Unrendered package part" in wn]
        self.assertEqual(unrendered, ["Unrendered package part: word/media/fallback.png; changes to its content are not compared."])
        self.assertNotIn(MISSING_IMAGE_NOTE, stderr)

    def test_block_level_compatibility_copies_record_images_once(self) -> None:
        def write(fallback: str) -> None:
            choice = anchored("Choice:", picture("choice"))
            backup = anchored("Fallback:", vml_picture("fallback"))
            body = paragraph("Before.") + f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="wps">{choice}</mc:Choice><mc:Fallback>{backup}</mc:Fallback></mc:AlternateContent>'
            make_docx(self.root / "essay.docx", body,
                      extra={"word/media/choice.png": "choice bytes", "word/media/fallback.png": fallback},
                      relationships=image_rel("choice", "media/choice.png") + image_rel("fallback", "media/fallback.png"))

        write("fallback bytes")
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        # Prose still reads both copies, as before; only the mc:Choice image is recorded.
        self.assertEqual(prose, b"Before.\n\nChoice:[Unsupported drawing]\n\nFallback:[Unsupported pict]\n")
        self.assertEqual(json.loads(meta)["images"], [image_entry("choice bytes")])
        write("changed fallback bytes")
        self.cli("check", "essay.docx")

    def test_external_and_dangling_images_are_recorded_without_reading_bytes(self) -> None:
        body = (anchored("Linked:", picture("linked", "link")) + anchored("No relationship:", picture("nowhere"))
                + anchored("No part:", picture("lost")))
        relationships = (image_rel("linked", "https://example.invalid/figure.png", external=True)
                         + image_rel("lost", "media/lost.png"))
        prose, meta, stderr = self.export_images(body, relationships=relationships)
        self.assertEqual(meta["images"], [
            {"part": "word/document.xml", "target": "https://example.invalid/figure.png"},
            {"part": "word/document.xml", "sha256": None},
            {"part": "word/document.xml", "sha256": None},
        ])
        self.assertIn(MISSING_IMAGE_NOTE, meta["warnings"])
        self.assertIn(MISSING_IMAGE_NOTE, stderr)
        self.assertEqual(prose.count("[Unsupported drawing]"), 3)
        self.cli("check", "essay.docx")

    def test_corrupt_compressed_media_is_a_document_error(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, anchored("Figure:", picture("rId5")), extra={"word/media/image1.png": "picture " * 64},
                  relationships=image_rel("rId5", "media/image1.png"))
        with ZipFile(source) as archive:
            info = archive.getinfo("word/media/image1.png")
        data = bytearray(source.read_bytes())
        # Overwrite the start of the DEFLATE stream with an invalid block type.
        start = info.header_offset + 30 + len(info.filename.encode()) + len(info.extra)
        data[start:start + 8] = b"\xff" * 8
        source.write_bytes(bytes(data))
        result = self.cli("export", "essay.docx", code=2)
        self.assertIn("error: essay.docx:", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        self.assertFalse(self.companions()[0].exists())

    def test_unreferenced_media_still_warns_and_imageless_documents_record_an_empty_array(self) -> None:
        prose, meta, stderr = self.export_images(paragraph("Plain prose."), extra={"word/media/image9.png": "orphan"})
        self.assertEqual(meta["images"], [])
        self.assertIn("Unrendered package part: word/media/image9.png; changes to its content are not compared.", meta["warnings"])
        self.assertNotIn(IMAGE_NOTE, meta["warnings"])
        prose, meta, stderr = self.export_images(paragraph("Plain prose."))
        self.assertEqual(meta["images"], [])

    def review_report(self, old: str, new: str) -> str:
        """``pilcrow review old new``'s report, which must find a change; read as the UTF-8 it is
        written in, not the locale encoding ``cli`` decodes with (cp1252 on Windows)."""
        env = {**os.environ, "PYTHONPATH": os.pathsep.join((str(PROJECT / "src"), str(PROJECT)))}
        result = subprocess.run([sys.executable, "-m", "pilcrow", "review", old, new], cwd=self.root, env=env,
                                capture_output=True, timeout=20)
        self.assertEqual(result.returncode, 1, result)
        return result.stdout.decode("utf-8")

    def write_described(self, body: str, path: str = "essay.docx") -> None:
        make_docx(self.root / path, body, extra={"word/media/image1.png": "picture bytes"},
                  relationships=image_rel("rId5", "media/image1.png"))

    def test_alt_text_edits_make_check_stale_and_review_names_them(self) -> None:
        def figure(descr: str | None) -> str:
            # Word writes the same description on the drawing and on the picture.
            return anchored("Figure:", described_picture("rId5", descr, own=descr))

        self.write_described(figure("A red cat"))
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        data = json.loads(meta)
        self.assertEqual(data["images"], [{**image_entry("picture bytes"), "alt": "A red cat"}])
        self.assertIn(IMAGE_ALT_NOTE, data["warnings"])
        self.assertNotIn(IMAGE_NOTE, data["warnings"])
        self.assertEqual(prose, b"Figure:[Unsupported drawing]\n", "alt text is not rendered in prose")
        self.write_described(figure("A red cat"), "old.docx")
        for descr, line in (
            ("A blue dog", '~ image 1 (word/document.xml) alt text: "A red cat" → "A blue dog"'),
            (None, '~ image 1 (word/document.xml) alt text: "A red cat" → (none)'),
            ("", '~ image 1 (word/document.xml) alt text: "A red cat" → (none)'),
        ):
            with self.subTest(descr=descr):
                self.write_described(figure(descr))
                self.cli("check", "essay.docx", code=1)
                report = self.review_report("old.docx", "essay.docx")
                self.assertIn(f"Images:\n{line}\n", report)
                self.assertIn("Prose: no changes", report)
        # Adding alt text to an image that had none is a change too.
        self.write_described(figure(None))
        self.cli("export", "essay.docx")
        after = self.exported()
        self.assertEqual(after[0], prose)
        self.assertEqual(json.loads(after[1])["images"], [image_entry("picture bytes")])
        self.assertIn(IMAGE_NOTE, json.loads(after[1])["warnings"])
        self.write_described(figure(None), "old.docx")
        self.write_described(figure("A red cat"))
        self.cli("check", "essay.docx", code=1)
        report = self.review_report("old.docx", "essay.docx")
        self.assertIn('Images:\n~ image 1 (word/document.xml) alt text: (none) → "A red cat"\n', report)

    def test_alt_text_reads_the_same_from_either_attribute(self) -> None:
        def alt(body: str) -> list[dict[str, Any]]:
            self.write_described(body)
            self.cli("export", "essay.docx")
            images: list[dict[str, Any]] = json.loads(self.exported()[1])["images"]
            return images

        expected = [{**image_entry("picture bytes"), "alt": "A red cat"}]
        # On the drawing only, on the picture only, or spread over lines: one recorded value.
        self.assertEqual(alt(anchored("", described_picture("rId5", "A red cat"))), expected)
        self.assertEqual(alt(anchored("", described_picture("rId5", own="A red cat"))), expected)
        self.assertEqual(alt(anchored("", described_picture("rId5", " A red\n  cat ", own="A red cat"))), expected)
        self.cli("check", "essay.docx")
        self.write_described(anchored("", described_picture("rId5", own="A red cat")))
        self.cli("check", "essay.docx")
        # The drawing's description wins over the picture's own, as in Word's alt text pane.
        self.assertEqual(alt(anchored("", described_picture("rId5", "A red cat", own="A cat"))), expected)
        # A title is recorded apart from the description, and only when present.
        self.assertEqual(alt(anchored("", described_picture("rId5", title="Cat"))),
                         [{**image_entry("picture bytes"), "alt_title": "Cat"}])
        self.assertEqual(alt(anchored("", described_picture("rId5", "", own=" ", title=""))),
                         [image_entry("picture bytes")])
        # A legacy VML picture keeps its alt text on the shape.
        vml = vml_picture("rId5").replace('id="Picture 1"', 'id="Picture 1" alt="A red cat"')
        self.assertEqual(alt(anchored("", vml)), expected)

    def test_grouped_pictures_keep_their_own_alt_text(self) -> None:
        def member(rid: str, descr: str | None) -> str:
            props = "" if descr is None else f' descr="{descr}"'
            return (f'<pic:pic xmlns:pic="{PIC}"><pic:nvPicPr><pic:cNvPr id="2" name="p"{props}/></pic:nvPicPr>'
                    f'<pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic>')
        group = (f'<w:drawing><wp:anchor xmlns:wp="{WP}"><wp:docPr id="1" name="Group 1" descr="Two pets"/>'
                 f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{WPG}"><wpg:wgp xmlns:wpg="{WPG}">'
                 f'{member("rId5", "A red cat")}{member("rId6", None)}</wpg:wgp></a:graphicData></a:graphic>'
                 '</wp:anchor></w:drawing>')
        make_docx(self.root / "essay.docx", anchored("Group:", group),
                  extra={"word/media/image1.png": "cat bytes", "word/media/image2.png": "dog bytes"},
                  relationships=image_rel("rId5", "media/image1.png") + image_rel("rId6", "media/image2.png"))
        self.cli("export", "essay.docx")
        self.assertEqual(json.loads(self.exported()[1])["images"],
                         [{**image_entry("cat bytes"), "alt": "A red cat"}, image_entry("dog bytes")])

    def test_documents_without_alt_text_or_wordart_keep_their_companions(self) -> None:
        # Companions exported by agent/improvements before alt text and WordArt were recorded
        # (87a23c6): DrawingML, described-but-empty and VML pictures, and a WordArt shape type
        # with no text, must export byte for byte the same.
        shape_type = vml_wordart("unused").split("<v:shape ")[0] + "</w:pict>"
        body = (paragraph("Caption above.") + anchored("Figure:", picture("rId5"))
                + anchored("Described:", described_picture("rId6")) + anchored("Legacy:", vml_picture("rId7"))
                + anchored("Shape type:", shape_type))
        make_docx(self.root / "essay.docx", body, extra={f"word/media/image{i}.png": f"picture {i}" for i in (5, 6, 7)},
                  relationships="".join(image_rel(f"rId{i}", f"media/image{i}.png") for i in (5, 6, 7)))
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        self.assertEqual(hashlib.sha256(prose).hexdigest(),
                         "37997d7537bbac3bd80770076952b402b2538528a7fd7efef5d99ea6f9ff8959", prose)
        self.assertEqual(hashlib.sha256(meta).hexdigest(),
                         "bf59117831f52348f3221f581f11cef4d2ebcd040d2dc6c0f19a750b89a928e8", meta)
        self.assertNotIn("objects", json.loads(meta))

    def test_wordart_text_edits_make_check_stale_and_review_names_them(self) -> None:
        def write(text: str, path: str = "essay.docx") -> None:
            make_docx(self.root / path, anchored("Banner:", vml_wordart(text)))

        write("SALE")
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        self.assertEqual(prose, b"Banner:[Unsupported pict]\n", "WordArt is still not rendered in prose")
        data = json.loads(meta)
        self.assertEqual(data["objects"], [wordart_entry("SALE")])
        self.assertIn("Unsupported pict: content is not rendered in prose; inspect the DOCX in Word.", data["warnings"])
        self.assertEqual(data["images"], [])
        write("SALE", "old.docx")
        write("CLOSED")
        self.cli("check", "essay.docx", code=1)
        report = self.review_report("old.docx", "essay.docx")
        self.assertIn('Charts, diagrams and embedded objects:\n~ WordArt 1 (word/document.xml): "SALE" → "CLOSED"\n',
                      report)
        # Whitespace is normalised like other recorded text.
        write("  SALE\n ")
        self.cli("check", "essay.docx")

    def test_wordart_follows_tracked_deletions_and_compatibility_copies(self) -> None:
        def objects(body: str) -> list[dict[str, Any]] | None:
            _, meta = self.export_objects(body)
            return meta.get("objects")

        deleted = f'<w:p>{run("Kept.")}<w:del w:id="1" w:author="Ada"><w:r>{vml_wordart("Gone")}</w:r></w:del></w:p>'
        self.assertIsNone(objects(deleted))
        moved = f'<w:p>{run("Kept.")}<w:moveFrom w:id="1" w:author="Ada"><w:r>{vml_wordart("Gone")}</w:r></w:moveFrom></w:p>'
        self.assertIsNone(objects(moved))
        inserted = f'<w:p>{run("Kept.")}<w:ins w:id="1" w:author="Ada"><w:r>{vml_wordart("New")}</w:r></w:ins></w:p>'
        self.assertEqual(objects(inserted), [wordart_entry("New")])
        both = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="v">{vml_wordart("SALE")}</mc:Choice>'
                f'<mc:Fallback>{vml_wordart("SALE")}</mc:Fallback></mc:AlternateContent>')
        self.assertEqual(objects(f"<w:p><w:r>{both}</w:r></w:p>"), [wordart_entry("SALE")])
        block = both.replace(vml_wordart("SALE"), f'<w:p><w:r>{vml_wordart("SALE")}</w:r></w:p>')
        self.assertEqual(objects(block), [wordart_entry("SALE")])
        # In document order with other objects, and by story part.
        header = f'<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r>{vml_wordart("Head")}</w:r></w:p></w:hdr>'
        body = (f'<w:p><w:r>{vml_wordart("One")}</w:r>{symbol()}<w:r>{vml_wordart("Two")}</w:r></w:p>'
                '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>')
        _, meta = self.export_objects(body, extra={"word/header1.xml": header},
                                      relationships=f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>')
        self.assertEqual(meta["objects"], [
            wordart_entry("One"), {"kind": "symbol", "part": "word/document.xml", "font": "Symbol", "char": "F061"},
            wordart_entry("Two"), wordart_entry("Head", "word/header1.xml")])

    def export_objects(self, body: str, **options: Any) -> tuple[str, dict[str, Any]]:
        make_docx(self.root / "essay.docx", body, **options)
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        return prose.decode("utf-8"), json.loads(meta)

    def test_equation_and_symbol_edits_make_check_stale(self) -> None:
        def body(power: str = "2", char: str = "F061", font: str = "Symbol") -> str:
            return f'<w:p>{run("Energy: ")}{equation(energy(power))}{run(" and ")}{symbol(char, font)}</w:p>'

        prose, meta = self.export_objects(body())
        self.assertEqual(prose, "Energy: [Math: E=mc^2] and [Unsupported sym]\n")
        # The fingerprints are what they were before equations had prose.
        self.assertEqual(meta["objects"], [
            equation_entry(energy_form()),
            {"kind": "symbol", "part": "word/document.xml", "font": "Symbol", "char": "F061"},
        ])
        self.assertIn(MATH_NOTE, meta["warnings"])
        self.assertFalse([w for w in meta["warnings"] if "oMath" in w or "XML construct" in w], meta["warnings"])
        self.assertIn("Unsupported sym: content is not rendered in prose; inspect the DOCX in Word.", meta["warnings"])
        companions = self.exported()
        # Exporting the same document again gives byte-identical companions.
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), companions)
        for changed in (body("3"), body(char="F062"), body(font="Wingdings")):
            make_docx(self.root / "essay.docx", changed)
            result = self.cli("check", "essay.docx", code=1)
            self.assertIn("STALE", result.stdout + result.stderr)
        # w:char is a hex number: its letter case is not a change.
        make_docx(self.root / "essay.docx", body(char="f061"))
        self.cli("check", "essay.docx")
        make_docx(self.root / "essay.docx", body("3"))
        self.cli("export", "essay.docx")
        self.cli("check", "essay.docx")
        after = self.exported()
        self.assertEqual(after[0], b"Energy: [Math: E=mc^3] and [Unsupported sym]\n")
        self.assertEqual(json.loads(after[1])["objects"][0], equation_entry(energy_form("3")))

    def test_documents_without_equations_or_symbols_have_no_objects_key(self) -> None:
        _, meta = self.export_objects(paragraph("Plain prose."))
        self.assertNotIn("objects", meta)
        # A tracked deletion is not in the accepted view: its equation and symbol are not recorded.
        deleted = f'<w:p>{run("Kept.")}<w:del w:id="1" w:author="Ada"><w:r>{equation()}</w:r>{symbol()}</w:del></w:p>'
        _, meta = self.export_objects(deleted)
        self.assertNotIn("objects", meta)
        inserted = f'<w:p>{run("Kept.")}<w:ins w:id="1" w:author="Ada">{equation()}{symbol()}</w:ins></w:p>'
        _, meta = self.export_objects(inserted)
        self.assertEqual([entry["kind"] for entry in meta["objects"]], ["equation", "symbol"])

    def test_equations_in_notes_headers_and_text_boxes_are_recorded_by_part(self) -> None:
        header = f'<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p>{equation("h")}</w:p></w:hdr>'
        footnotes = (f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}"><w:footnote w:id="1">'
                     f'<w:p>{equation("f")}</w:p></w:footnote></w:footnotes>')
        boxed = f'<w:p>{equation("t")}</w:p>'
        # An equation paragraph: one entry for the m:oMathPara, not one per m:oMath inside it.
        display = (f'<m:oMathPara xmlns:m="{M}"><m:oMath>{math_run("a")}</m:oMath>'
                   f'<m:oMath>{math_run("b")}</m:oMath></m:oMathPara>')
        body = (f'<w:p>{equation("b")}</w:p>' + anchored("Box:", modern_text_box(boxed)) + f"<w:p>{display}</w:p>"
                + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
                + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>')
        relationships = (f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
                         f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        _, meta = self.export_objects(body, extra={"word/header1.xml": header, "word/footnotes.xml": footnotes},
                                      relationships=relationships)
        para_form = [f"{{{M}}}oMathPara", [], "", [simple_form("a"), simple_form("b")]]
        self.assertEqual(meta["objects"], [
            equation_entry(simple_form("b")), equation_entry(simple_form("t")), equation_entry(para_form),
            equation_entry(simple_form("f"), "word/footnotes.xml"),
            equation_entry(simple_form("h"), "word/header1.xml"),
        ])

    def test_equation_fingerprint_is_the_accepted_view_without_formatting(self) -> None:
        def write(inner: str) -> None:
            make_docx(self.root / "essay.docx", f"<w:p>{equation(inner)}</w:p>")

        base = (f'<m:sSup><m:e>{math_run("E")}</m:e><m:sup>{math_run("2")}</m:sup></m:sSup>'
                f'<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub/><m:sup/><m:e>{math_run("x")}</m:e></m:nary>')
        write(base)
        self.cli("export", "essay.docx")
        fresh = [
            # Run and control character formatting, revision IDs, proofing marks, bookmarks, comment markers.
            base.replace("<m:r>", '<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr>'),
            base.replace("<m:naryPr>", '<m:naryPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr></m:ctrlPr>'),
            base.replace("<m:r>", '<m:r w:rsidR="00A1B2C3" w:rsidRPr="00D4E5F6">'),
            base.replace("<m:sSup>", '<w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="0" w:name="_Ref1"/><m:sSup>')
            .replace("</m:nary>", '</m:nary><w:bookmarkEnd w:id="0"/><w:proofErr w:type="spellEnd"/>'),
            base.replace("<m:sSup>", '<w:commentRangeStart w:id="3"/><m:sSup>')
            .replace("</m:sSup>", '</m:sSup><w:commentRangeEnd w:id="3"/><w:commentReference w:id="3"/>'),
            # A tracked deletion is not in the accepted view.
            base + f'<w:del w:id="9" w:author="Ada" w:date="2026-01-01T00:00:00Z">{math_run("+1")}</w:del>',
            # A tracked insertion of the same content is the same accepted view.
            base.replace(math_run("x"), f'<w:ins w:id="7" w:author="Ada">{math_run("x")}</w:ins>'),
        ]
        for variant in fresh:
            with self.subTest(variant=variant):
                write(variant)
                self.cli("check", "essay.docx")
        stale = [
            base + f'<w:ins w:id="9" w:author="Ada">{math_run("+1")}</w:ins>',
            base.replace(math_run("2"), math_run("3")),
            base.replace(math_run("x"), math_run(" x")),  # m:t text is kept exactly
            base.replace('m:val="∑"', 'm:val="∏"'),  # structure properties carry meaning
            base.replace("<m:sSup>", "<m:sSub>").replace("</m:sSup>", "</m:sSub>").replace("m:sup>", "m:sub>", 2),
        ]
        for variant in stale:
            with self.subTest(variant=variant):
                write(variant)
                self.cli("check", "essay.docx", code=1)

    def test_block_level_equations_are_recorded(self) -> None:
        # OOXML allows math directly in the body, a table cell, a block-level insertion, and a
        # block-level compatibility copy; deleted block-level math is not in the accepted view.
        cell = f"<w:tbl><w:tr><w:tc>{equation('c', para=True)}</w:tc></w:tr></w:tbl>"
        inserted = f'<w:ins w:id="1" w:author="Ada">{equation("i", para=True)}</w:ins>'
        deleted = f'<w:del w:id="2" w:author="Ada">{equation("d", para=True)}</w:del>'
        compatible = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="m">{equation("m", para=True)}'
                      f'</mc:Choice><mc:Fallback>{equation("f", para=True)}</mc:Fallback></mc:AlternateContent>')
        body = paragraph("Before.") + equation("b", para=True) + cell + inserted + deleted + compatible

        def para_entry(text: str) -> dict[str, Any]:
            return equation_entry([f"{{{M}}}oMathPara", [], "", [simple_form(text)]])

        prose, meta = self.export_objects(body)
        self.assertEqual(meta["objects"], [para_entry(t) for t in "bcim"])
        # Block-level math is a paragraph of its own; without a text box, every compatibility
        # copy's prose is read, as for paragraphs, while only the chosen copy is fingerprinted.
        self.assertEqual(prose, "Before.\n\n[Math: b]\n\n[Table]\n\n[Row]\n\n[Cell]\n\n[Math: c]\n\n[/Row]\n\n"
                                "[/Table]\n\n[Math: i]\n\n[Math: m]\n\n[Math: f]\n")
        make_docx(self.root / "essay.docx", body.replace(equation("b", para=True), equation("B", para=True)))
        self.cli("check", "essay.docx", code=1)

    def test_compatibility_copies_record_one_equation(self) -> None:
        inline = (f'<w:p>{run("Formula: ")}<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="m">'
                  f'{equation(para=True)}</mc:Choice><mc:Fallback><w:r>{vml_picture("fallback")}</w:r></mc:Fallback>'
                  "</mc:AlternateContent></w:p>")
        block = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="m"><w:p>{equation()}</w:p></mc:Choice>'
                 f'<mc:Fallback>{anchored("", vml_picture("fallback"))}</mc:Fallback></mc:AlternateContent>')
        media = {"word/media/fallback.png": "fallback bytes"}
        para_form = [f"{{{M}}}oMathPara", [], "", [simple_form()]]
        for body, form in ((inline, para_form), (block, simple_form())):
            with self.subTest(body=body):
                _, meta = self.export_objects(body, extra=media, relationships=image_rel("fallback", "media/fallback.png"))
                self.assertEqual(meta["objects"], [equation_entry(form)])
                self.assertEqual(meta["images"], [])

    def test_chart_diagram_embedded_object_and_chunk_edits_make_check_stale(self) -> None:
        def write(body: str = OBJECT_BODY, **parts: str) -> None:
            make_docx(self.root / "essay.docx", body, extra=object_parts(**parts), relationships=OBJECT_RELS)

        write()
        self.cli("export", "essay.docx")
        prose, meta = self.exported()
        data = json.loads(meta)
        parts = object_parts()
        self.assertEqual(data["objects"], [
            part_entry("chart", parts["word/charts/chart1.xml"]),
            part_entry("diagram", parts["word/diagrams/data1.xml"]),
            part_entry("embedded object", parts["word/embeddings/Microsoft_Excel_Worksheet.xlsx"], prog_id="Excel.Sheet.12"),
            part_entry("imported chunk", parts["word/afchunk.htm"]),
        ])
        self.assertEqual(prose.decode("utf-8"), "Results: [Unsupported drawing][Unsupported drawing][Unsupported object]"
                                                "\n\n[Unsupported altChunk]\n")
        # Hashed parts are compared now; formatting parts and the chart's workbook still are not.
        unrendered = sorted(wn.split(";")[0].removeprefix("Unrendered package part: ")
                            for wn in data["warnings"] if wn.startswith("Unrendered package part"))
        self.assertEqual(unrendered, ["word/diagrams/colors1.xml", "word/diagrams/layout1.xml",
                                      "word/diagrams/quickStyle1.xml", "word/embeddings/chart_workbook.xlsx"])
        self.assertIn("Unsupported object: content is not rendered in prose; inspect the DOCX in Word.", data["warnings"])
        # Exporting the same document again gives byte-identical companions.
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), (prose, meta))
        edits: list[dict[str, Any]] = [{"value": "99"}, {"workbook": "workbook v2"}, {"node": "Build"},
                                       {"chunk": "Other text"}]
        for edit in edits:
            with self.subTest(edit=edit):
                write(**edit)
                result = self.cli("check", "essay.docx", code=1)
                self.assertIn("STALE", result.stdout + result.stderr)
        # Neither workbook a chart reads its values from nor diagram formatting is fingerprinted.
        changed = object_parts()
        changed["word/embeddings/chart_workbook.xlsx"] = "edited chart workbook"
        changed["word/diagrams/colors1.xml"] = "<colors edited='1'/>"
        make_docx(self.root / "essay.docx", OBJECT_BODY, extra=changed, relationships=OBJECT_RELS)
        self.cli("check", "essay.docx")
        # Adding or removing an object is a change too.
        added = OBJECT_BODY.replace("</w:p>", f'<w:r>{chart("chart")}</w:r></w:p>', 1)
        removed = OBJECT_BODY.replace(alt_chunk("chunk"), "")
        for body in (added, removed):
            with self.subTest(body=body):
                write(body)
                self.cli("check", "essay.docx", code=1)
        write(value="99")
        self.cli("export", "essay.docx")
        self.cli("check", "essay.docx")
        after = self.exported()
        self.assertEqual(after[0], prose, "object edits must not alter the prose")
        self.assertNotEqual(after[1], meta)

    def test_objects_in_headers_notes_and_text_boxes_are_recorded_by_part(self) -> None:
        chart_rels = f'<Relationships xmlns="{PKG_R}">{rel("hc", "chart", "charts/chart2.xml")}</Relationships>'
        header = f'<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r>{chart("hc")}</w:r></w:p></w:hdr>'
        footnotes = (f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}"><w:footnote w:id="1">'
                     f'<w:p><w:r>{chart("fc")}</w:r></w:p></w:footnote></w:footnotes>')
        footnote_rels = f'<Relationships xmlns="{PKG_R}">{rel("fc", "chart", "charts/chart3.xml")}</Relationships>'
        boxed = f'<w:p><w:r>{chart("chart")}</w:r></w:p>'
        # A chart in a tracked deletion is not in the accepted view.
        deleted = f'<w:p>{run("Kept.")}<w:del w:id="1" w:author="Ada"><w:r>{chart("chart")}</w:r></w:del></w:p>'
        body = (anchored("Box:", modern_text_box(boxed)) + deleted
                + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
                + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>')
        relationships = (rel("chart", "chart", "charts/chart1.xml") + rel("headerRef", "header", "header1.xml")
                         + rel("footnotes", "footnotes", "footnotes.xml"))
        extra = {"word/header1.xml": header, "word/_rels/header1.xml.rels": chart_rels,
                 "word/footnotes.xml": footnotes, "word/_rels/footnotes.xml.rels": footnote_rels,
                 "word/charts/chart1.xml": "<body chart/>", "word/charts/chart2.xml": "<header chart/>",
                 "word/charts/chart3.xml": "<footnote chart/>"}
        _, meta = self.export_objects(body, extra=extra, relationships=relationships)
        self.assertEqual(meta["objects"], [
            part_entry("chart", "<body chart/>"),
            part_entry("chart", "<footnote chart/>", "word/footnotes.xml"),
            part_entry("chart", "<header chart/>", "word/header1.xml"),
        ])
        self.assertFalse([wn for wn in meta["warnings"] if "Unrendered package part" in wn], meta["warnings"])

    def test_renamed_object_part_with_the_same_content_stays_fresh(self) -> None:
        body = f'<w:p><w:r>{chart("chart")}</w:r></w:p>'
        make_docx(self.root / "essay.docx", body, extra={"word/charts/chart1.xml": "<chart/>"},
                  relationships=rel("chart", "chart", "charts/chart1.xml"))
        self.cli("export", "essay.docx")
        make_docx(self.root / "essay.docx", body, extra={"word/charts/chart2.xml": "<chart/>"},
                  relationships=rel("chart", "chart", "charts/chart2.xml"))
        self.cli("check", "essay.docx")

    def test_linked_and_missing_object_parts(self) -> None:
        missing_note = ("Some chart, diagram, embedded object or imported chunk relationships are missing "
                        "or point to missing parts; those objects are recorded without a hash.")
        linked = f'<w:p><w:r>{ole_object("link", "Excel.Sheet.12", "Link")}</w:r></w:p>'
        _, meta = self.export_objects(linked, relationships=rel("link", "oleObject", "file:///C:/data/budget.xlsx", external=True))
        self.assertEqual(meta["objects"], [{"kind": "embedded object", "part": "word/document.xml",
                                            "prog_id": "Excel.Sheet.12", "target": "file:///C:/data/budget.xlsx"}])
        self.assertNotIn(missing_note, meta["warnings"])
        # A relationship that is missing, one whose part is missing, and a chart with no r:id at all.
        dangling = (f'<w:p><w:r>{chart("nowhere")}</w:r><w:r>{chart("gone")}</w:r>'
                    f'<w:r>{graphic(C, f"<c:chart xmlns:c={chr(34)}{C}{chr(34)}/>")}</w:r></w:p>{alt_chunk("gone")}')
        _, meta = self.export_objects(dangling, relationships=rel("gone", "chart", "charts/chart9.xml"))
        self.assertEqual(meta["objects"], [part_entry("chart", None)] * 3 + [part_entry("imported chunk", None)])
        self.assertIn(missing_note, meta["warnings"])

    def test_compatibility_copies_record_one_chart(self) -> None:
        # Word stores a chart in mc:Choice with a picture of it in mc:Fallback; an object only the
        # unselected copy references is not hashed, and keeps its unrendered-part warning.
        fallback = f'<w:r>{vml_picture("fallback")}</w:r><w:r>{ole_object("ole")}</w:r>'
        inline = (f'<w:p>{run("Chart: ")}<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="c14">'
                  f'<w:r>{chart("chart")}</w:r></mc:Choice><mc:Fallback>{fallback}</mc:Fallback>'
                  "</mc:AlternateContent></w:p>")
        block = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="c14"><w:p><w:r>{chart("chart")}</w:r></w:p>'
                 f'</mc:Choice><mc:Fallback><w:p>{fallback}</w:p></mc:Fallback></mc:AlternateContent>')
        extra = {"word/charts/chart1.xml": "<chart/>", "word/media/fallback.png": "fallback bytes",
                 "word/embeddings/oleObject1.bin": "ole bytes"}
        relationships = (rel("chart", "chart", "charts/chart1.xml") + image_rel("fallback", "media/fallback.png")
                         + rel("ole", "oleObject", "embeddings/oleObject1.bin"))
        for body in (inline, block):
            with self.subTest(body=body):
                _, meta = self.export_objects(body, extra=extra, relationships=relationships)
                self.assertEqual(meta["objects"], [part_entry("chart", "<chart/>")])
                self.assertEqual(meta["images"], [])
                unrendered = [wn for wn in meta["warnings"] if "Unrendered package part" in wn]
                self.assertEqual(unrendered, [
                    "Unrendered package part: word/embeddings/oleObject1.bin; changes to its content are not compared.",
                    "Unrendered package part: word/media/fallback.png; changes to its content are not compared.",
                ])

    def test_an_imported_chunk_inside_a_paragraph_is_recorded(self) -> None:
        body = f'<w:p>{run("Before.")}{alt_chunk("chunk")}</w:p>'
        _, meta = self.export_objects(body, extra={"word/afchunk.htm": "<p>x</p>"},
                                      relationships=rel("chunk", "aFChunk", "afchunk.htm"))
        self.assertEqual(meta["objects"], [part_entry("imported chunk", "<p>x</p>")])

    def test_deeply_nested_text_boxes_fail_cleanly(self) -> None:
        content = paragraph("Innermost.")
        for _ in range(400):
            # The Fallback copy is left empty: duplicating each level would double the fixture.
            content = anchored("", modern_text_box(content, fallback=""))
        make_docx(self.root / "essay.docx", content)
        result = self.cli("export", "essay.docx", code=2)
        self.assertIn("error: essay.docx:", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        self.assertFalse(self.companions()[0].exists())

    def test_setup_and_export_leave_git_index_and_history_untouched(self) -> None:
        self.init_git()
        make_docx(self.root / "essay.docx", paragraph("Original text."))
        (self.root / ".gitattributes").write_text("*.png binary\n", encoding="utf-8")
        self.git("add", ".")
        self.git("commit", "-qm", "Fixture baseline")
        index_before = hashlib.sha256((self.root / ".git/index").read_bytes()).digest()
        head_before = self.git("rev-parse", "HEAD")
        self.cli("setup", "--root", self.root)
        self.cli("export", "--all", "--root", self.root)
        attributes = (self.root / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("*.png binary", attributes)
        self.assertIn(".pilcrow", attributes)
        self.assertEqual(self.git("check-attr", "merge", "--", "essay.docx"), "essay.docx: merge: unset\n")
        self.assertIn("text: set", self.git("check-attr", "text", "--", ".pilcrow/essay.docx.prose.md"))
        self.assertEqual(hashlib.sha256((self.root / ".git/index").read_bytes()).digest(), index_before)
        self.assertEqual(self.git("rev-parse", "HEAD"), head_before)
        setup_before = attributes
        self.cli("setup", "--root", self.root)
        self.assertEqual((self.root / ".gitattributes").read_text(encoding="utf-8"), setup_before)
        self.git("add", ".")
        self.git("commit", "-qm", "Fixture companions")
        index_before = hashlib.sha256((self.root / ".git/index").read_bytes()).digest()
        head_before = self.git("rev-parse", "HEAD")
        make_docx(self.root / "essay.docx", paragraph("Revised text."))
        self.cli("export", "--all", "--root", self.root)
        diff = self.git("diff", "--", ".pilcrow/essay.docx.prose.md")
        self.assertIn("-Original text.", diff)
        self.assertIn("+Revised text.", diff)
        self.assertEqual(hashlib.sha256((self.root / ".git/index").read_bytes()).digest(), index_before)
        self.assertEqual(self.git("rev-parse", "HEAD"), head_before)


    def test_identical_names_and_excluded_directories(self) -> None:
        for name in ("chapter-a/report.docx", "chapter-b/report.docx"):
            make_docx(self.root / name, paragraph(name))
        for folder in (".git", ".venv", "venv", "env", "node_modules", "__pycache__", ".tox", ".nox", "build", "dist", ".pilcrow"):
            if folder == ".git":
                continue  # a fake .git is intentionally an error, not a standalone directory
            make_docx(self.root / folder / "skip.docx", paragraph("Excluded."))
        self.cli("export", "--all")
        self.assertIn(b"chapter-a/report.docx", self.exported("chapter-a/report.docx")[0])
        self.assertIn(b"chapter-b/report.docx", self.exported("chapter-b/report.docx")[0])
        companions = list((self.root / ".pilcrow").rglob("*.prose.md"))
        self.assertEqual(len(companions), 2)

    def test_discovery_from_git_subdirectory_does_not_ascend(self) -> None:
        self.init_git()
        make_docx(self.root / "outside.docx", paragraph("Outside."))
        nested = self.root / "nested"
        make_docx(nested / "inside.docx", paragraph("Inside."))
        self.cli("export", "--all", cwd=nested)
        self.assertTrue((nested / ".pilcrow/inside.docx.prose.md").exists())
        self.assertFalse((self.root / ".pilcrow").exists())

    def test_noop_export_keeps_mtimes_and_formatting_changes_are_ignored(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Same words."))
        self.cli("export", "essay.docx")
        for path in self.companions():
            os.utime(path, (1_600_000_000, 1_600_000_000))
        mtimes = [p.stat().st_mtime_ns for p in self.companions()]
        make_docx(self.root / "essay.docx", '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Same words.</w:t></w:r></w:p>')
        self.cli("check", "essay.docx")
        self.cli("export", "essay.docx")
        self.assertEqual([p.stat().st_mtime_ns for p in self.companions()], mtimes)

    def test_zip_compression_does_not_affect_freshness(self) -> None:
        source = self.root / "essay.docx"
        make_docx(source, paragraph("Same text."))
        self.cli("export", "essay.docx")
        with ZipFile(source) as archive:
            parts = [(name, archive.read(name)) for name in archive.namelist()]
        with ZipFile(source, "w") as archive:
            for name, content in reversed(parts):
                archive.writestr(name, content)  # ZIP_STORED rather than ZIP_DEFLATED
        self.cli("check", "essay.docx")

    def test_comment_changes_are_metadata_only(self) -> None:
        source = self.root / "essay.docx"
        relationships = f'<Relationship Id="comments" Type="{R}/comments" Target="comments.xml"/>'
        def write_comment(text: str) -> None:
            comments = f'<w:comments xmlns:w="{W}"><w:comment w:id="1" w:author="Ada">{paragraph(text)}</w:comment></w:comments>'
            make_docx(source, paragraph("Stable body."), extra={"word/comments.xml": comments}, relationships=relationships)
        write_comment("Please clarify.")
        self.cli("export", "essay.docx")
        before = self.exported()
        write_comment("Clarified, thank you.")
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        after = self.exported()
        self.assertEqual(before[0], after[0])
        self.assertIn(b"Clarified, thank you.", after[1])
        self.assertNotEqual(before[1], after[1])

    def write_commented(self, body: str, comments: dict[str, str] | None = None, *,
                        extra: dict[str, str] | None = None, relationships: str = "") -> None:
        make_docx(self.root / "essay.docx", body,
                  extra={"word/comments.xml": comments_part(comments or {"1": "Note."}), **(extra or {})},
                  relationships=COMMENTS_REL + relationships)

    def write_threaded(self, comments: dict[str, str], para_ids: dict[str, str],
                       *entries: tuple[str, str | None, str | None], extended: bool = True) -> None:
        """A body anchoring every comment, with a comments-extended part holding ``entries``."""
        body = "".join(f"<w:p>{start(c)}{run(f'Passage {n}.')}{end(c)}</w:p>" for n, c in enumerate(comments))
        make_docx(self.root / "essay.docx", body,
                  extra={"word/comments.xml": comments_part(comments, para_ids),
                         **({"word/commentsExtended.xml": comments_extended(*entries)} if extended else {})},
                  relationships=COMMENTS_REL + (COMMENTS_EXT_REL if extended else ""))

    def comment_threads(self) -> list[tuple[str, bool | None, int | None]]:
        """Each exported comment's (text, resolved, reply_to), in review order."""
        review = json.loads(self.exported()[1])["review"]
        return [(item["text"], item["resolved"], item["reply_to"]) for item in review if item["kind"] == "comment"]

    def comment_anchors(self) -> dict[str, tuple[str | None, str | None]]:
        """Each exported comment's text -> (anchor, anchor_part)."""
        review = json.loads(self.exported()[1])["review"]
        return {item["text"]: (item["anchor"], item["anchor_part"]) for item in review if item["kind"] == "comment"}

    def test_moving_or_resizing_a_comment_range_is_stale_in_metadata_only(self) -> None:
        def write(before: str, anchored: str, after: str) -> None:
            self.write_commented(f"<w:p>{run(before)}{start('1')}{run(anchored)}{end('1')}{run(after)}</w:p>")
        write("The ", "first sentence.", " The second sentence.")
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {"Note.": ("first sentence.", "word/document.xml")})
        prose = self.exported()[0]
        for text in [("The first sentence. The ", "second sentence.", ""),  # moved
                     ("The ", "first sentence. The second", " sentence."),  # widened
                     ("The first ", "sentence.", " The second sentence.")]:  # narrowed
            before = self.exported()[1]
            write(*text)
            self.cli("check", "essay.docx", code=1)
            self.cli("export", "essay.docx")
            self.cli("check", "essay.docx")
            self.assertEqual(self.exported()[0], prose)
            self.assertNotEqual(self.exported()[1], before)
            self.assertEqual(self.comment_anchors()["Note."][0], text[1].strip())

    def test_renumbered_comment_ids_and_run_splits_do_not_churn(self) -> None:
        self.write_commented(f"<w:p>{run('Keep ')}{start('1')}{run('this phrase')}{end('1')}{run(' here.')}</w:p>"
                             f"<w:p>{start('2')}{run('Second.')}{end('2')}</w:p>",
                             {"1": "First note.", "2": "Second note."})
        self.cli("export", "essay.docx")
        before = self.exported()
        # Word renumbers comments on save; split runs, proofing marks and bookmarks are save churn.
        split = (f"<w:p>{run('Keep ')}{start('41')}<w:proofErr w:type=\"spellStart\"/>{run('thi')}{run('s ph')}"
                 f'<w:bookmarkStart w:id="0" w:name="mark"/>{run("rase")}<w:bookmarkEnd w:id="0"/>'
                 f"<w:proofErr w:type=\"spellEnd\"/>{end('41')}{run(' here.')}</w:p>"
                 f"<w:p>{start('007')}<w:r><w:t>Sec</w:t><w:t>ond.</w:t></w:r>{end('7')}</w:p>")
        self.write_commented(split, {"41": "First note.", "7": "Second note."})
        self.cli("check", "essay.docx")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), before)
        meta = json.loads(before[1])
        self.assertNotIn('"41"', before[1].decode())
        self.assertIn(COMMENT_NOTE, meta["warnings"])
        self.assertNotIn(UNANCHORED_COMMENT_NOTE, meta["warnings"])

    def test_comment_anchors_use_the_accepted_view(self) -> None:
        body = (f"<w:p>{start('1')}{run('Kept')}"
                '<w:ins w:id="10" w:author="Ben"><w:r><w:t xml:space="preserve"> added</w:t></w:r></w:ins>'
                '<w:del w:id="11" w:author="Ben"><w:r><w:delText xml:space="preserve"> removed</w:delText></w:r></w:del>'
                '<w:r><w:tab/><w:t>after</w:t><w:br/><w:t>a break</w:t></w:r>'
                f"{end('1')}</w:p>"
                # A range that ends inside deleted text still closes.
                f"<w:p>{start('2')}{run('Before deletion')}"
                f'<w:del w:id="12" w:author="Ben"><w:r><w:delText>gone</w:delText></w:r><w:commentRangeEnd w:id="2"/></w:del>'
                f"{reference('2')}{run(' outside.')}</w:p>")
        self.write_commented(body, {"1": "Accepted.", "2": "Ends in a deletion."})
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {"Accepted.": ("Kept added after a break", "word/document.xml"),
                                                  "Ends in a deletion.": ("Before deletion", "word/document.xml")})
        self.assertNotIn(UNCLOSED_COMMENT_NOTE, json.loads(self.exported()[1])["warnings"])

    def test_overlapping_nested_multi_paragraph_cell_and_note_ranges(self) -> None:
        footnotes = (f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>{run("Note ")}{start("6")}'
                     f'{run("anchored text")}{end("6")}</w:p></w:footnote></w:footnotes>')
        body = (f"<w:p>{start('1')}{run('Alpha ')}{start('2')}{run('beta ')}{end('1')}{run('gamma')}{end('2')}</w:p>"
                f"<w:p>{start('3')}{run('Outer ')}{start('4')}{run('inner')}{end('4')}{run(' outer.')}{end('3')}</w:p>"
                f"<w:p>{run('One ')}{start('5')}{run('paragraph')}</w:p>{paragraph('Next paragraph')}"
                f"<w:p>{run('ends')}{end('5')}</w:p>"
                f"<w:tbl><w:tr><w:tc><w:p>{run('Cell ')}{start('7')}{run('one')}</w:p></w:tc>"
                f"<w:tc><w:p>{run('cell two')}{end('7')}</w:p></w:tc></w:tr></w:tbl>"
                # Block-level markers between paragraphs, as Word writes for whole-paragraph ranges.
                f"{start('8')}{paragraph('Whole paragraph.')}<w:commentRangeEnd w:id=\"8\"/>"
                f"<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r>{reference('8')}</w:p>")
        comments = {str(n): f"Comment {n}." for n in range(1, 9)}
        self.write_commented(body, comments, extra={"word/footnotes.xml": footnotes},
                             relationships=f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        self.cli("export", "essay.docx")
        body_part = "word/document.xml"
        self.assertEqual(self.comment_anchors(), {
            "Comment 1.": ("Alpha beta", body_part),
            "Comment 2.": ("beta gamma", body_part),
            "Comment 3.": ("Outer inner outer.", body_part),
            "Comment 4.": ("inner", body_part),
            "Comment 5.": ("paragraph Next paragraph ends", body_part),
            "Comment 6.": ("anchored text", "word/footnotes.xml"),
            "Comment 7.": ("one cell two", body_part),
            "Comment 8.": ("Whole paragraph.", body_part),
        })
        # Moving a comment from the body to a note with the same anchored text changes anchor_part.
        footnotes = (f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>{run("Note ")}{start("1")}'
                     f'{run("Alpha beta")}{end("1")}</w:p></w:footnote></w:footnotes>')
        body = (f"<w:p>{run('Alpha ')}{run('beta')}</w:p>"
                f"<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>")
        self.write_commented(body, {"1": "Comment 1."}, extra={"word/footnotes.xml": footnotes},
                             relationships=f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {"Comment 1.": ("Alpha beta", "word/footnotes.xml")})

    def test_point_unreferenced_and_unclosed_comments(self) -> None:
        body = (f"<w:p>{run('Point here')}{reference('1')}{run(' and ')}{start('2')}{end('2')}{run('empty.')}</w:p>"
                f"<w:p>{run('Unclosed ')}{start('4')}{run('to the end')}</w:p>{paragraph('of the body.')}")
        self.write_commented(body, {"1": "Point.", "2": "Empty.", "3": "Nowhere.", "4": "Unclosed."})
        result = self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {
            "Point.": ("", "word/document.xml"),
            "Empty.": ("", "word/document.xml"),
            "Nowhere.": (None, None),
            "Unclosed.": ("to the end of the body.", "word/document.xml"),
        })
        warnings = json.loads(self.exported()[1])["warnings"]
        self.assertIn(UNANCHORED_COMMENT_NOTE, warnings)
        self.assertIn(UNCLOSED_COMMENT_NOTE, warnings)
        self.assertIn("null", result.stderr)
        # An end without a start runs from the start of its part.
        self.write_commented(f"{paragraph('From the start')}<w:p>{run('to here')}{end('1')}{run(' only.')}</w:p>")
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {"Note.": ("From the start to here", "word/document.xml")})
        self.assertIn(UNCLOSED_COMMENT_NOTE, json.loads(self.exported()[1])["warnings"])
        self.assertNotIn(UNANCHORED_COMMENT_NOTE, json.loads(self.exported()[1])["warnings"])

    def test_text_box_text_is_a_separate_story_for_comment_anchors(self) -> None:
        def write(box: str) -> None:
            # Comment 1 spans two body paragraphs; the first anchors a box holding comment 2.
            self.write_commented(f"<w:p>{start('1')}{run('Body first.')}<w:r>{modern_text_box(box)}</w:r></w:p>"
                                 f"<w:p>{run('Body last.')}{end('1')}</w:p>",
                                 {"1": "Body.", "2": "Box."})
        write(f"<w:p>{run('Box ')}{start('2')}{run('story.')}{end('2')}</w:p>")
        self.cli("export", "essay.docx")
        self.assertIn(b"[Text box]\n\nBox story.\n\n[/Text box]", self.exported()[0])
        self.assertEqual(self.comment_anchors(), {"Body.": ("Body first. Body last.", "word/document.xml"),
                                                  "Box.": ("story.", "word/document.xml")})
        self.assertNotIn(UNCLOSED_COMMENT_NOTE, json.loads(self.exported()[1])["warnings"])
        before = self.exported()[1]
        # Editing box text outside every range leaves both anchors, and the metadata, unchanged.
        write(f"<w:p>{run('Edited box ')}{start('2')}{run('story.')}{end('2')}</w:p>")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported()[1], before)
        # A range left open in a box runs to the end of the box, not into the body.
        write(f"<w:p>{start('2')}{run('Open in the box')}</w:p>")
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_anchors(), {"Body.": ("Body first. Body last.", "word/document.xml"),
                                                  "Box.": ("Open in the box", "word/document.xml")})
        self.assertIn(UNCLOSED_COMMENT_NOTE, json.loads(self.exported()[1])["warnings"])

    def test_resolving_a_comment_is_stale_in_metadata_only(self) -> None:
        comments, ids = {"1": "First.", "2": "Second."}, {"1": "1A2B3C4D", "2": "0F0F0F0F"}
        self.write_threaded(comments, ids, ("1A2B3C4D", None, "0"), ("0F0F0F0F", None, "0"))
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual(self.comment_threads(), [("First.", False, None), ("Second.", False, None)])
        meta = json.loads(before[1])
        self.assertIn(COMMENT_NOTE, meta["warnings"])
        self.assertFalse([wn for wn in meta["warnings"] if "Unrendered package part" in wn], meta["warnings"])
        self.assertNotIn("1A2B3C4D", before[1].decode())
        # Resolving one comment in Word only flips w15:done.
        self.write_threaded(comments, ids, ("1A2B3C4D", None, "1"), ("0F0F0F0F", None, "0"))
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        self.cli("check", "essay.docx")
        self.assertEqual(self.exported()[0], before[0])
        self.assertEqual(self.comment_threads(), [("First.", True, None), ("Second.", False, None)])
        # Other ST_OnOff spellings read the same; reopening is stale again.
        self.write_threaded(comments, ids, ("1A2B3C4D", None, "true"), ("0F0F0F0F", None, "off"))
        self.cli("check", "essay.docx")
        self.write_threaded(comments, ids, ("1A2B3C4D", None, "0"), ("0F0F0F0F", None, "0"))
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), before)

    def test_replies_record_their_parent_position(self) -> None:
        comments, ids = {"1": "First.", "2": "Second."}, {"1": "11111111", "2": "22222222"}
        self.write_threaded(comments, ids, ("11111111", None, "0"), ("22222222", None, "0"))
        self.cli("export", "essay.docx")
        # Word writes a reply as a new comment whose commentEx names its parent's paraId.
        comments, ids = {**comments, "3": "Reply."}, {**ids, "3": "33333333"}
        self.write_threaded(comments, ids, ("11111111", None, "0"), ("22222222", None, "0"),
                            ("33333333", "11111111", "0"))
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_threads(),
                         [("First.", False, None), ("Second.", False, None), ("Reply.", False, 0)])
        # Re-parenting the reply is stale too.
        self.write_threaded(comments, ids, ("11111111", None, "0"), ("22222222", None, "0"),
                            ("33333333", "22222222", "0"))
        self.cli("check", "essay.docx", code=1)
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_threads()[2], ("Reply.", False, 1))
        self.cli("check", "essay.docx")

    def test_renumbered_comment_and_paragraph_ids_do_not_churn(self) -> None:
        self.write_threaded({"1": "First.", "2": "Reply."}, {"1": "11111111", "2": "22222222"},
                            ("11111111", None, "1"), ("22222222", "11111111", "0"))
        self.cli("export", "essay.docx")
        before = self.exported()
        self.assertEqual(self.comment_threads(), [("First.", True, None), ("Reply.", False, 0)])
        # Word renumbers w:id and regenerates paraIds consistently across parts on re-save.
        self.write_threaded({"40": "First.", "7": "Reply."}, {"40": "5e6f7a8b", "7": "0000ABCD"},
                            ("5E6F7A8B", None, "1"), ("0000abcd", "5e6f7a8b", "0"))
        self.cli("check", "essay.docx")
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported(), before)
        for raw in ("11111111", "22222222", "5E6F7A8B", "0000ABCD", "durableId"):
            self.assertNotIn(raw, before[1].decode())

    def test_comment_threads_use_the_last_paragraph_and_tolerate_gaps(self) -> None:
        # Comments without a comments-extended part record unknown state, with no new warning.
        self.write_threaded({"1": "First.", "2": "Second."}, {"1": "11111111", "2": "22222222"}, extended=False)
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_threads(), [("First.", None, None), ("Second.", None, None)])
        plain = json.loads(self.exported()[1])["warnings"]
        # A commentEx refers to the last paragraph of a multi-paragraph comment, not the first.
        comments = (f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}">'
                    f'<w:comment w:id="1" w:author="Ada"><w:p w14:paraId="0000000A">{run("Opening.")}</w:p>'
                    f'<w:p w14:paraId="0000000B">{run("Closing.")}</w:p></w:comment>'
                    f'<w:comment w:id="2" w:author="Ada"><w:p w14:paraId="0000000C">{run("Only.")}</w:p></w:comment>'
                    f'<w:comment w:id="3" w:author="Ada">{paragraph("No paraId.")}</w:comment>'
                    f'<w:comment w:id="4" w:author="Ada"><w:p w14:paraId="0000000D">{run("Self.")}</w:p></w:comment>'
                    f'<w:comment w:id="5" w:author="Ada"><w:p w14:paraId="0000000E">{run("No done.")}</w:p></w:comment>'
                    "</w:comments>")
        extended = comments_extended(
            ("0000000A", None, "1"),  # the first paragraph: no comment's last paragraph
            ("0000000B", None, "0"),
            ("0000000C", "FFFFFFFF", "1"),  # a dangling parent
            ("0000000C", "0000000B", "0"),  # a duplicate entry: the first one wins
            ("12345678", "0000000B", "1"),  # matches no comment
            ("0000000D", "0000000D", "0"),  # its own parent
            ("0000000E", "0000000B", None),
        )
        body = "".join(f"<w:p>{start(c)}{run(f'Passage {c}.')}{end(c)}</w:p>" for c in "12345")
        make_docx(self.root / "essay.docx", body,
                  extra={"word/comments.xml": comments, "word/commentsExtended.xml": extended},
                  relationships=COMMENTS_REL + COMMENTS_EXT_REL)
        self.cli("export", "essay.docx")
        self.assertEqual(self.comment_threads(), [
            ("Opening.Closing.", False, None), ("Only.", True, None), ("No paraId.", None, None),
            ("Self.", False, None), ("No done.", False, 0)])
        self.assertEqual(json.loads(self.exported()[1])["warnings"], plain)

    def test_comments_extended_part_coverage_and_errors(self) -> None:
        body = paragraph("No comments here.")
        extended = {"word/commentsExtended.xml": comments_extended()}
        # Not reached through the document's relationships, the part is still reported.
        make_docx(self.root / "essay.docx", body, extra=extended)
        self.cli("export", "essay.docx")
        self.assertIn("Unrendered package part: word/commentsExtended.xml; changes to its content are not compared.",
                      json.loads(self.exported()[1])["warnings"])
        make_docx(self.root / "essay.docx", body, extra=extended, relationships=COMMENTS_EXT_REL)
        self.cli("export", "essay.docx")
        self.assertFalse([wn for wn in json.loads(self.exported()[1])["warnings"] if "Unrendered" in wn])
        # A relationship to a missing or malformed part is an error, as for styles and numbering.
        make_docx(self.root / "essay.docx", body, relationships=COMMENTS_EXT_REL)
        result = self.cli("export", "essay.docx", code=2)
        self.assertIn("missing related part word/commentsExtended.xml", result.stderr)
        make_docx(self.root / "essay.docx", body, extra={"word/commentsExtended.xml": "<broken"},
                  relationships=COMMENTS_EXT_REL)
        result = self.cli("export", "essay.docx", code=2)
        self.assertIn("cannot parse word/commentsExtended.xml", result.stderr)

    def test_table_content_controls_and_cell_boundaries(self) -> None:
        body = '<w:tbl><w:sdt><w:sdtContent><w:tr><w:sdt><w:sdtContent><w:tc>'
        body += paragraph("One cell. Another sentence.")
        body += '</w:tc></w:sdtContent></w:sdt><w:tc>' + paragraph("Another cell.")
        body += '</w:tc></w:tr></w:sdtContent></w:sdt></w:tbl>'
        make_docx(self.root / "essay.docx", body)
        self.cli("export", "essay.docx")
        prose = self.exported()[0].decode()
        self.assertEqual(prose.count("[Cell]"), 2)
        self.assertIn("One cell.\nAnother sentence.", prose)
        self.assertIn("Another cell.", prose)

    def test_one_bad_input_does_not_silently_skip_other_documents(self) -> None:
        (self.root / "bad.docx").write_bytes(b"not a ZIP")
        make_docx(self.root / "good.docx", paragraph("Good text."))
        result = self.cli("export", "--all", code=2)
        self.assertIn("bad.docx", result.stderr)
        self.assertIn(b"Good text.", self.exported("good.docx")[0])

    def test_output_symlink_is_rejected(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Text."))
        target = self.root / "elsewhere"
        target.mkdir()
        try:
            (self.root / ".pilcrow").symlink_to(target, target_is_directory=True)
        except OSError:
            self.skipTest("symlink creation unavailable on this platform")
        self.cli("export", "essay.docx", code=2)
        self.assertEqual(list(target.iterdir()), [])

    def test_failed_atomic_replace_keeps_previous_companion(self) -> None:
        from pilcrow.cli import main
        source = self.root / "essay.docx"
        make_docx(source, paragraph("Before."))
        self.cli("export", source)
        before = self.exported()
        make_docx(source, paragraph("After."))
        stderr = io.StringIO()
        with (
            patch("pilcrow.files.os.replace", side_effect=OSError("simulated write failure")),
            redirect_stdout(io.StringIO()),
            redirect_stderr(stderr),
        ):
            status = main(["export", str(source), "--root", str(self.root)])
        self.assertEqual(status, 2)
        self.assertIn("simulated write failure", stderr.getvalue())
        self.assertEqual(self.exported(), before)
        self.assertEqual(list((self.root / ".pilcrow").glob(".pilcrow-tmp-*")), [])



    def test_setup_forces_conflict_for_both_sided_companion_edits(self) -> None:
        self.init_git()
        make_docx(self.root / "essay.docx", paragraph("First sentence.") + paragraph("Last sentence."))
        self.cli("export", "essay.docx")
        self.cli("setup")
        self.git("add", ".")
        self.git("commit", "-qm", "Baseline")
        original_branch = self.git("branch", "--show-current").strip()
        self.git("checkout", "-qb", "collaborator")
        make_docx(self.root / "essay.docx", paragraph("Collaborator sentence.") + paragraph("Last sentence."))
        self.cli("export", "essay.docx")
        self.git("add", ".")
        self.git("commit", "-qm", "Collaborator change")
        self.git("checkout", original_branch)
        make_docx(self.root / "essay.docx", paragraph("First sentence.") + paragraph("My sentence."))
        self.cli("export", "essay.docx")
        self.git("add", ".")
        self.git("commit", "-qm", "My change")
        result = subprocess.run(["git", "merge", "--no-edit", "collaborator"], cwd=self.root, capture_output=True, text=True)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        conflicts = self.git("diff", "--name-only", "--diff-filter=U")
        self.assertIn(".pilcrow/essay.docx.prose.md", conflicts)
        self.assertIn("essay.docx", conflicts)
        # Regenerating from the chosen authoritative DOCX resolves companion content.
        self.cli("export", "essay.docx")
        self.cli("check", "essay.docx")

    def test_hidden_text_is_reported(self) -> None:
        make_docx(self.root / "essay.docx", '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden words.</w:t></w:r></w:p>')
        result = self.cli("export", "essay.docx")
        self.assertIn("Hidden text is included", result.stderr)
        self.assertIn(b"Hidden words.", self.exported()[0])


class EquationTextTests(unittest.TestCase):
    """Equations as linear [Math: …] text in prose, read in process."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-math-")
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "essay.docx"

    def extract(self, body: str, **options: Any) -> tuple[str, dict[str, Any]]:
        from pilcrow.extract import extract

        make_docx(self.path, body, **options)
        result = extract(self.path)
        return result.prose, result.metadata

    def math(self, content: str) -> str:
        """The prose of one paragraph holding one equation with ``content``."""
        return self.extract(f"<w:p>{equation(content)}</w:p>")[0]

    def test_each_structure_has_its_linear_text(self) -> None:
        r = math_run
        cases = [
            # Text runs, whitespace collapsed; w:r inside math; formatting adds no text.
            (r(" a  +\n b "), "a + b"),
            (f'{r("x")}<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> if </w:t></w:r>{r("y")}', "x if y"),
            ('<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>z</m:t></m:r>', "z"),
            # Fractions: an atom (one character, letters only, a number) is not wrapped.
            (structure("f", ("num", "a"), ("den", "b")), "a/b"),
            (structure("f", ("num", "a+b"), ("den", "2c")), "(a+b)/(2c)"),
            (structure("f", ("num", "dy"), ("den", "3.5")), "dy/3.5"),
            (structure("f", ("num", "n"), ("den", "k"), props='<m:type m:val="noBar"/>'), "n¦k"),
            (structure("f", ("num", "a"), ("den", "b"), props='<m:type m:val="skw"/>'), "a/b"),
            # Scripts.
            (structure("sSup", ("e", "x"), ("sup", "2")), "x^2"),
            (structure("sSup", ("e", "x"), ("sup", "n+1")), "x^(n+1)"),
            (structure("sSub", ("e", "x"), ("sub", "i")), "x_i"),
            (structure("sSubSup", ("e", "x"), ("sub", "i"), ("sup", "2")), "x_i^2"),
            (structure("sPre", ("sub", "i"), ("sup", "2"), ("e", "x")), "_i^2 x"),
            # Radicals: a degree shows unless it is hidden or empty.
            (structure("rad", ("deg", ""), ("e", "x"), props='<m:degHide m:val="1"/>'), "√(x)"),
            (structure("rad", ("deg", "3"), ("e", "x")), "√(3&x)"),
            (structure("rad", ("deg", "3"), ("e", "x+1"), props="<m:degHide/>"), "√(x+1)"),
            (structure("rad", ("deg", ""), ("e", "x"), props='<m:degHide m:val="off"/>'), "√(x)"),
            # Delimiters: defaults, custom characters, and an empty m:val for none.
            (structure("d", ("e", "a"), ("e", "b")), "(a|b)"),
            (structure("d", ("e", "0"), ("e", "1"), props='<m:begChr m:val="["/><m:sepChr m:val=","/>'), "[0,1)"),
            (structure("d", ("e", "x"), props='<m:begChr m:val="|"/><m:endChr m:val="|"/>'), "|x|"),
            (structure("d", ("e", "x"), props='<m:begChr m:val=""/><m:endChr m:val="}"/>'), "x}"),
            # N-ary operators.
            (structure("nary", ("sub", "i=1"), ("sup", "n"), ("e", structure("sSub", ("e", "x"), ("sub", "i"))),
                       props='<m:chr m:val="∑"/>'), "∑_(i=1)^n x_i"),
            (structure("nary", ("sub", "0"), ("sup", "1"), ("e", "f")), "∫_0^1 f"),
            (structure("nary", ("sub", ""), ("sup", ""), ("e", "x"),
                       props='<m:chr m:val="∏"/><m:subHide m:val="1"/><m:supHide/>'), "∏x"),
            (structure("nary", ("sub", "S"), ("sup", "T"), ("e", "x"), props="<m:supHide/>"), "∫_S x"),
            # Functions and limits.
            (structure("func", ("fName", "sin"), ("e", "x")), "sin x"),
            (structure("func", ("fName", structure("limLow", ("e", "lim"), ("lim", "n→∞"))),
                       ("e", structure("sSub", ("e", "a"), ("sub", "n")))), "lim_(n→∞) a_n"),
            (structure("limUpp", ("e", "max"), ("lim", "k")), "max^k"),
            # Accents, bars, group characters.
            (structure("acc", ("e", "x")), "x̂"),
            (structure("acc", ("e", "x+y"), props='<m:chr m:val="̃"/>'), "(x+y)̃"),
            (structure("bar", ("e", "x"), props='<m:pos m:val="top"/>'), "x̅"),
            (structure("bar", ("e", "x")), "x̲"),
            (structure("groupChr", ("e", "a+b")), "(a+b)⏟"),
            (structure("groupChr", ("e", "a"), props='<m:chr m:val="⏞"/><m:pos m:val="top"/>'), "a⏞"),
            # Boxes and phantoms show their content.
            (structure("box", ("e", "a+b")), "a+b"),
            (structure("borderBox", ("e", "E")), "E"),
            (structure("phant", ("e", "x"), props='<m:show m:val="0"/>'), "x"),
            # Matrices and equation arrays.
            ("<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"2\"/></m:mcPr></m:mc></m:mcs></m:mPr>"
             + "".join(f"<m:mr><m:e>{r(a)}</m:e><m:e>{r(b)}</m:e></m:mr>" for a, b in (("a", "b"), ("c", "d")))
             + "</m:m>", "[a, b; c, d]"),
            (structure("eqArr", ("e", "x=1"), ("e", "y=2")), "x=1; y=2"),
            # Nesting: a fraction inside a superscript inside a delimiter.
            (structure("d", ("e", structure("sSup", ("e", "x"), ("sup", structure("f", ("num", "a+b"), ("den", "2")))))),
             "(x^((a+b)/2))"),
            # A delimiter is not wrapped again as an operand.
            (structure("sSup", ("e", structure("d", ("e", "a+b"))), ("sup", "2")), "(a+b)^2"),
            (structure("sSup", ("e", structure("d", ("e", "a+b"), props='<m:begChr m:val=""/>')), ("sup", "2")),
             "(a+b))^2"),
        ]
        for content, expected in cases:
            with self.subTest(expected=expected):
                self.assertEqual(self.math(content), f"[Math: {expected}]\n")

    def test_an_equation_paragraph_shows_each_equation(self) -> None:
        display = (f'<m:oMathPara xmlns:m="{M}"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
                   f'<m:oMath>{math_run("a=1")}</m:oMath><m:oMath>{energy()}</m:oMath></m:oMathPara>')
        prose, meta = self.extract(f"<w:p>{run('So ')}{display}</w:p>")
        self.assertEqual(prose, "So [Math: a=1] [Math: E=mc^2]\n")
        self.assertEqual([entry["kind"] for entry in meta["objects"]], ["equation"])
        self.assertEqual(self.math(""), "[Math: (empty)]\n")
        self.assertEqual(self.extract(f'<w:p><m:oMathPara xmlns:m="{M}"/></w:p>')[0], "[Math: (empty)]\n")

    def test_block_level_math_is_its_own_paragraph(self) -> None:
        cell = f"<w:tbl><w:tr><w:tc>{paragraph('Cell.')}{equation('c', para=True)}</w:tc></w:tr></w:tbl>"
        body = paragraph("Before.") + equation(energy()) + cell + paragraph("After.")
        prose, _ = self.extract(body)
        self.assertEqual(prose, "Before.\n\n[Math: E=mc^2]\n\n[Table]\n\n[Row]\n\n[Cell]\n\nCell.\n\n[Math: c]\n\n"
                                "[/Row]\n\n[/Table]\n\nAfter.\n")

    def test_math_text_is_the_accepted_view(self) -> None:
        content = (math_run("a") + f'<w:ins w:id="1" w:author="Ada">{math_run("+b")}</w:ins>'
                   + f'<w:del w:id="2" w:author="Ada">{math_run("+c")}</w:del>'
                   + f'<w:moveFrom w:id="3" w:author="Ada">{math_run("+d")}</w:moveFrom>'
                   + f'<w:moveTo w:id="4" w:author="Ada">{math_run("+e")}</w:moveTo>')
        prose, meta = self.extract(f"<w:p>{equation(content)}</w:p>")
        self.assertEqual(prose, "[Math: a+b+e]\n")
        # Tracked changes inside an equation are not listed with the document's, as before.
        self.assertEqual(meta["review"], [])

    def test_an_unknown_math_element_shows_its_text_and_warns_once(self) -> None:
        unknown = f"<m:future>{math_run('q')}</m:future>"
        prose, meta = self.extract(f"<w:p>{equation(unknown + math_run('+') + unknown)}</w:p>")
        self.assertEqual(prose, "[Math: q+q]\n")
        self.assertEqual([w for w in meta["warnings"] if "future" in w],
                         ["Unsupported math construct future; rendered as its text, structure not shown."])
        self.assertFalse([w for w in meta["warnings"] if "XML construct" in w], meta["warnings"])
        # A symbol inside an equation keeps its placeholder.
        prose, meta = self.extract(f"<w:p>{equation(math_run('a') + symbol())}</w:p>")
        self.assertEqual(prose, "[Math: a[Unsupported sym]]\n")
        self.assertIn("Unsupported sym: content is not rendered in prose; inspect the DOCX in Word.", meta["warnings"])

    def test_sentences_never_split_inside_math(self) -> None:
        body = (f'<w:p>{run("We have ")}{equation("x. Y")}{run(" here. Next one.")}</w:p>'
                f'<w:p>{run("Count ")}{equation("n! Then")}{run(" ways? Yes.")}</w:p>'
                f'<w:p>{run("Ends with ")}{equation(math_run("y") + structure("f", ("num", "1"), ("den", "2")))}'
                f'{run(". Then more.")}</w:p>')
        prose, _ = self.extract(body)
        self.assertEqual(prose, "We have [Math: x. Y] here.\nNext one.\n\nCount [Math: n! Then] ways?\nYes.\n\n"
                                "Ends with [Math: y1/2].\nThen more.\n")
        # Headings, list items and hyperlink text show the same math text.
        heading = f'<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>{run("Proof of ")}{equation("a. B")}</w:p>'
        item = f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>{run("Item ")}{equation("p! Q")}</w:p>'
        link = f'<w:p><w:hyperlink w:anchor="eq">{run("See ")}{equation("z")}</w:hyperlink></w:p>'
        prose, meta = self.extract(heading + item + link)
        self.assertEqual(prose, "## Proof of [Math: a. B]\n\n- Item [Math: p! Q]\n\nSee [Math: z]\n")
        self.assertEqual(meta["hyperlinks"], [{"text": "See [Math: z]", "target": "#eq", "part": "word/document.xml"}])

    def test_a_comment_anchor_over_math_holds_its_text(self) -> None:
        body = f'<w:p>{start("1")}{run("Energy: ")}{equation(energy())}{end("1")}</w:p>'
        _, meta = self.extract(body, extra={"word/comments.xml": comments_part({"1": "Check units."})},
                               relationships=COMMENTS_REL)
        (comment,) = meta["review"]
        self.assertEqual(comment["anchor"], "Energy: [Math: E=mc^2]")

    def test_deeply_nested_math_fails_cleanly(self) -> None:
        from pilcrow.extract import DocumentError

        # Rendering costs no more stack per level than the fingerprint: 200 nested fractions
        # exported before equations had prose, and still do.
        nested = "x"
        for _ in range(200):
            nested = structure("f", ("num", nested), ("den", "y"))
        prose, _ = self.extract(f"<w:p>{equation(nested)}</w:p>")
        self.assertTrue(prose.startswith("[Math: " + "(" * 199 + "x/y)/y)/y"), prose[:300])
        for _ in range(2800):
            nested = structure("f", ("num", nested), ("den", "y"))
        with self.assertRaises(DocumentError):
            self.extract(f"<w:p>{equation(nested)}</w:p>")
        # As a block of its own, too.
        with self.assertRaises(DocumentError):
            self.extract(equation(nested, para=True))

    def test_documents_without_math_keep_their_companions(self) -> None:
        # Pinned from the companions Pilcrow wrote before equations had prose.
        body = (paragraph("Title", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')
                + paragraph("Dr. Smith wrote this. Art. 5 applies! Is it x. Y? Yes.")
                + list_item("First item. Second sentence.", "1") + list_item("Nested.", "1", 1)
                + "<w:tbl><w:tr><w:tc>" + paragraph("Cell text.") + "</w:tc><w:tc/></w:tr></w:tbl>"
                + f'<w:p>{start("1")}{run("Commented passage.")}{end("1")}</w:p>'
                + f'<w:p>{run("Kept ")}<w:ins w:id="5" w:author="Ada">{run("inserted")}</w:ins>'
                + '<w:del w:id="6" w:author="Ada"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>'
                + f'<w:p><w:hyperlink r:id="link">{run("A link")}</w:hyperlink>{symbol()}<w:r><w:br/></w:r>'
                + f'{run("After break.")}</w:p>'
                + anchored("Box:", modern_text_box(paragraph("Inside the box.")))
                + anchored("Figure:", picture("img")) + f"<w:p>{checkbox('1')}</w:p>")
        relationships = (COMMENTS_REL + NUMBERING_REL + image_rel("img", "media/image1.png")
                         + rel("link", "hyperlink", "https://example.invalid/", external=True))
        extra = {"word/comments.xml": comments_part({"1": "Please check."}), "word/media/image1.png": "png bytes",
                 "word/numbering.xml": numbering(abstract_num("0", level(0), level(1, "decimal", "%2.")), num("1", "0"))}
        from pilcrow.extract import extract

        make_docx(self.path, body, extra=extra, relationships=relationships)
        outputs = extract(self.path).outputs()
        self.assertEqual({name: hashlib.sha256(data).hexdigest() for name, data in outputs.items()}, {
            ".prose.md": "8a6483a884064a67cfe3ac935b295872d8ad1d864b2b945332068c8189c038b5",
            ".meta.json": "8fd177d4171cfe566c5adda90807756c5fdd5596c408381f3b495d36f21208d2",
        })


if __name__ == "__main__":
    unittest.main()
