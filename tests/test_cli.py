"""Black-box regression tests built from small, generated OOXML packages.

The fixtures deliberately do not use Word, python-docx, or checked-in DOCX files.
Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
from contextlib import redirect_stderr, redirect_stdout
import io
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

    def cli(self, *args: str | Path, code: int = 0, cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
        env = os.environ.copy()
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
        with patch("pilcrow.files.os.replace", side_effect=OSError("simulated write failure")):
            with redirect_stdout(io.StringIO()), redirect_stderr(stderr):
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
