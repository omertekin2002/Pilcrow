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


def level(ilvl: int = 0, fmt: str = "bullet", text: str = "•", start: str | None = "1") -> str:
    start_xml = f'<w:start w:val="{start}"/>' if start is not None else ""
    return f'<w:lvl w:ilvl="{ilvl}">{start_xml}<w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/></w:lvl>'


def abstract_num(abstract_id: str, *levels: str) -> str:
    return f'<w:abstractNum w:abstractNumId="{abstract_id}">{"".join(levels)}</w:abstractNum>'


def num(num_id: str, abstract_id: str, overrides: str = "") -> str:
    return f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/>{overrides}</w:num>'


def numbering(*definitions: str) -> str:
    return f'<w:numbering xmlns:w="{W}">{"".join(definitions)}</w:numbering>'


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
            + list_item("Back to the first.", "7")
            + paragraph("", '<w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr>')  # empty: not rendered
            + '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>'
        )
        self.write_lists(body, numbering(
            abstract_num("0", level(0, "decimal", "%1."), level(2, "lowerRoman", "%3.")),
            abstract_num("1", level(0)),
            num("3", "1"), num("5", "1"), num("7", "0"), num("9", "1"),
        ), extra=extra, relationships=relationships)
        self.cli("export", "essay.docx")
        self.assertEqual(self.exported_lists(), [
            {"list": 1, "part": "word/document.xml", "levels": [
                {"level": 0, "format": "decimal", "text": "%1.", "start": 1},
                {"level": 2, "format": "lowerRoman", "text": "%3.", "start": 1},
            ]},
            {"list": 2, "part": "word/document.xml",
             "levels": [{"level": 0, "format": "bullet", "text": "•", "start": 1}]},
            {"list": 3, "part": "word/header1.xml",
             "levels": [{"level": 0, "format": "bullet", "text": "•", "start": 1}]},
        ])
        self.assertNotIn("word/numbering.xml", self.exported()[1].decode("utf-8"))
        self.assertEqual(json.loads(self.exported()[1])["list_runs"], [
            {"list": 1, "part": "word/document.xml", "items": 1},
            {"list": 2, "part": "word/document.xml", "items": 1},
            {"list": 1, "part": "word/document.xml", "items": 1},
            {"list": 1, "part": "word/header1.xml", "items": 1},
            {"list": 3, "part": "word/header1.xml", "items": 1},
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
            {"list": 1, "part": "word/document.xml", "items": 1},
            {"list": 2, "part": "word/document.xml", "items": 1},
            {"list": 1, "part": "word/document.xml", "items": 1},
            {"list": 2, "part": "word/document.xml", "items": 1},
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


if __name__ == "__main__":
    unittest.main()
