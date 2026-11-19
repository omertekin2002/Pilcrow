"""`pilcrow textconv` and the opt-in `setup --textconv` Git diff driver.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

try:
    from .test_cli import (
        COMMENTS_REL,
        NUMBERING_REL,
        PROJECT,
        R,
        W,
        abstract_num,
        anchored,
        comments_part,
        end,
        level,
        list_item,
        make_docx,
        modern_text_box,
        num,
        numbering,
        paragraph,
        run,
        start,
        vml_text_box,
    )
except ImportError:  # discovered with -s tests: the directory itself is on sys.path
    from test_cli import (  # type: ignore[no-redef]
        COMMENTS_REL,
        NUMBERING_REL,
        PROJECT,
        R,
        W,
        abstract_num,
        anchored,
        comments_part,
        end,
        level,
        list_item,
        make_docx,
        modern_text_box,
        num,
        numbering,
        paragraph,
        run,
        start,
        vml_text_box,
    )

BEGIN = "# BEGIN Pilcrow generated review rules"
FINISH = "# END Pilcrow generated review rules"
DIFF_LINE = "*.[dD][oO][cC][xX] diff=pilcrow"
CONFIG_COMMAND = 'git config diff.pilcrow.textconv "pilcrow textconv"'


def rich_documents() -> dict[str, tuple[str, dict[str, str], str]]:
    """Name -> (body, extra parts, relationships), covering the constructs prose renders."""
    styles = (f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Heading1">'
              '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>')
    lists = numbering(abstract_num("0", level(0, "decimal", "%1."), level(1, "bullet", "•")), num("1", "0"))
    structure = (
        paragraph("Methods", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')
        + list_item("First step. It has two sentences.", "1") + list_item("Nested step.", "1", 1)
        + '<w:tbl><w:tr><w:tc>' + paragraph("Left cell.") + '</w:tc><w:tc></w:tc><w:tc>'
        + paragraph("Right cell.") + '</w:tc></w:tr></w:tbl>'
        + paragraph("Café, naïve, 東京 — “quoted”.")
    )
    notes = {
        "word/header1.xml": f'<w:hdr xmlns:w="{W}">{paragraph("Header words.")}</w:hdr>',
        "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1">{paragraph("Footnote words.")}</w:footnote></w:footnotes>',
        "word/endnotes.xml": f'<w:endnotes xmlns:w="{W}"><w:endnote w:id="2">{paragraph("Endnote words.")}</w:endnote></w:endnotes>',
    }
    notes_rels = (
        f'<Relationship Id="headerRef" Type="{R}/header" Target="header1.xml"/>'
        f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>'
        f'<Relationship Id="endnotes" Type="{R}/endnotes" Target="endnotes.xml"/>'
    )
    notes_body = (
        paragraph("Body words.") + '<w:p><w:r><w:footnoteReference w:id="1"/><w:endnoteReference w:id="2"/></w:r></w:p>'
        '<w:sectPr><w:headerReference w:type="default" r:id="headerRef"/></w:sectPr>'
    )
    tracked = paragraph("Unchanged.") + (
        '<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>Inserted words.</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="Ben"><w:r><w:delText>Deleted words.</w:delText></w:r></w:del>'
        '<w:moveTo w:id="3" w:author="Cara"><w:r><w:t> Moved here.</w:t></w:r></w:moveTo></w:p>'
    )
    boxes = (
        paragraph("Before.") + anchored("Anchor.", modern_text_box(paragraph("Boxed sentence. Second one.")))
        + anchored("Legacy anchor.", vml_text_box(paragraph("Legacy box.")))
    )
    commented = f'<w:p>{start("1")}{run("Commented text.")}{end("1")}</w:p>'
    return {
        "structure.docx": (structure, {"word/styles.xml": styles, "word/numbering.xml": lists},
                           f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>' + NUMBERING_REL),
        "notes.docx": (notes_body, notes, notes_rels),
        "tracked.docx": (tracked, {}, ""),
        "boxes.docx": (boxes, {}, ""),
        "commented.docx": (commented, {"word/comments.xml": comments_part({"1": "A remark."})}, COMMENTS_REL),
    }


def tree(path: Path) -> list[tuple[str, int, int]]:
    """Every file and directory below ``path`` with its size and mtime."""
    listing = []
    for item in sorted(path.rglob("*")):
        info = item.stat()
        listing.append((item.relative_to(path).as_posix(), info.st_size, info.st_mtime_ns))
    return listing


class TextconvTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-textconv-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "repo"
        self.root.mkdir()
        self.outside = Path(self.temp.name) / "elsewhere"
        self.outside.mkdir()
        self.env = {**os.environ}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))

    def pilcrow(self, *args: str | Path, code: int = 0, cwd: Path | None = None) -> subprocess.CompletedProcess[bytes]:
        """Run Pilcrow, keeping stdout as raw bytes."""
        result = subprocess.run(
            [sys.executable, "-m", "pilcrow", *(str(arg) for arg in args)],
            cwd=cwd or self.root, env=self.env, capture_output=True, timeout=20,
        )
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout!r}\nstderr={result.stderr!r}")
        return result

    def git(self, *args: str) -> str:
        result = subprocess.run(["git", *args], cwd=self.root, env=self.env, capture_output=True,
                                text=True, encoding="utf-8", timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout

    def init_git(self) -> None:
        self.git("init", "-q")
        self.git("config", "user.name", "Pilcrow Test")
        self.git("config", "user.email", "pilcrow@example.invalid")
        self.git("config", "commit.gpgsign", "false")
        self.git("config", "core.autocrlf", "false")

    def prose_companion(self, name: str) -> bytes:
        return (self.root / ".pilcrow" / (name + ".prose.md")).read_bytes()

    def test_output_is_byte_identical_to_the_prose_companion(self) -> None:
        for name, (body, extra, relationships) in rich_documents().items():
            with self.subTest(document=name):
                make_docx(self.root / name, body, extra=extra, relationships=relationships)
                self.pilcrow("export", name)
                result = self.pilcrow("textconv", name)
                self.assertEqual(result.stdout, self.prose_companion(name))
                self.assertNotIn(b"warning", result.stdout)
                self.assertNotIn(b'"schema"', result.stdout)
                self.assertIn(f"warning: {name}: ".encode(), result.stderr)
        self.assertIn("東京".encode(), self.prose_companion("structure.docx"))

    def test_reads_a_temporary_copy_without_extension_outside_any_root(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("A sentence. Another sentence."))
        self.pilcrow("export", "essay.docx")
        # Git passes old revisions as temporary files with arbitrary names.
        copy = self.outside / "XXXXXX_essay"
        shutil.copyfile(self.root / "essay.docx", copy)
        result = self.pilcrow("textconv", copy, cwd=self.root)
        self.assertEqual(result.stdout, self.prose_companion("essay.docx"))
        self.assertEqual(self.pilcrow("textconv", copy.name, cwd=self.outside).stdout, result.stdout)
        # A lock-file name and a path inside the output directory are not excluded here either.
        lock = self.outside / ".pilcrow" / "~$essay.docx"
        lock.parent.mkdir()
        shutil.copyfile(copy, lock)
        self.assertEqual(self.pilcrow("textconv", lock).stdout, result.stdout)

    def test_unreadable_inputs_exit_2_with_nothing_on_stdout(self) -> None:
        not_zip = self.outside / "not-zip.docx"
        not_zip.write_bytes(b"This is not a ZIP archive.")
        over_limit = self.outside / "over-limit"
        with ZipFile(over_limit, "w") as archive:
            for index in range(4097):
                archive.writestr(f"part{index}.xml", "")
        no_body = self.outside / "no-body.docx"
        with ZipFile(no_body, "w") as archive:
            archive.writestr("word/other.xml", "<x/>")
        for path in (self.outside / "missing.docx", not_zip, over_limit, no_body, self.outside):
            with self.subTest(path=path.name):
                result = self.pilcrow("textconv", path, code=2)
                self.assertEqual(result.stdout, b"")
                self.assertTrue(result.stderr.startswith(f"error: {path}: ".encode()), result.stderr)
                self.assertNotIn(b"Traceback", result.stderr)

    def test_writes_no_files(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("A sentence."))
        broken = self.root / "broken.docx"
        broken.write_bytes(b"not a zip")
        before = tree(Path(self.temp.name))
        self.pilcrow("textconv", "essay.docx")
        self.pilcrow("textconv", broken, code=2)
        self.pilcrow("textconv", self.root / "essay.docx", cwd=self.outside)
        self.assertEqual(tree(Path(self.temp.name)), before)

    def test_textconv_takes_one_path_and_no_root(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("A sentence."))
        self.pilcrow("textconv", code=2)
        self.pilcrow("textconv", "essay.docx", "other.docx", code=2)
        self.pilcrow("textconv", "essay.docx", "--root", self.root, code=2)

    def attributes(self) -> str:
        return (self.root / ".gitattributes").read_bytes().decode("utf-8")

    def block(self) -> list[str]:
        text = self.attributes()
        return text[text.index(BEGIN):text.index(FINISH)].splitlines()

    def test_setup_textconv_adds_the_line_inside_the_managed_block(self) -> None:
        (self.root / ".gitattributes").write_bytes(b"*.png binary\n")
        result = self.pilcrow("setup", "--textconv")
        self.assertIn(DIFF_LINE, self.block())
        self.assertEqual(self.attributes().count(DIFF_LINE), 1)
        self.assertTrue(self.attributes().startswith("*.png binary\n"))
        self.assertIn(CONFIG_COMMAND.encode(), result.stdout)
        self.assertIn(b"cachetextconv", result.stdout)
        first = self.attributes()
        repeat = self.pilcrow("setup", "--textconv")
        self.assertEqual(self.attributes(), first)
        self.assertIn(b"Unchanged", repeat.stdout)
        # A later plain setup keeps the opt-in instead of silently dropping it.
        plain = self.pilcrow("setup")
        self.assertEqual(self.attributes(), first)
        self.assertIn(CONFIG_COMMAND.encode(), plain.stdout)

    def test_plain_setup_does_not_opt_in(self) -> None:
        result = self.pilcrow("setup")
        self.assertNotIn("diff=", self.attributes())
        self.assertNotIn(b"git config", result.stdout)
        plain = self.attributes()
        self.pilcrow("setup", "--textconv")
        self.assertIn(DIFF_LINE, self.block())
        # Everything else in the block is unchanged; the line is only added.
        self.assertEqual(self.attributes().replace(DIFF_LINE + "\n", ""), plain)

    def test_setup_textconv_upgrades_an_existing_block_and_keeps_user_rules(self) -> None:
        self.pilcrow("setup")
        old = self.attributes()
        (self.root / ".gitattributes").write_bytes(("*.png binary\r\n" + old + "*.pdf -diff\r\n").encode())
        self.pilcrow("setup", "--textconv")
        text = self.attributes()
        self.assertTrue(text.startswith("*.png binary\r\n"))
        self.assertTrue(text.endswith("*.pdf -diff\r\n"))
        self.assertIn(DIFF_LINE, self.block())
        self.assertEqual(text.count(BEGIN), 1)
        # A block with the line written with CRLF endings is still recognised as opted in.
        (self.root / ".gitattributes").write_bytes(text.replace("\n", "\r\n").replace("\r\r\n", "\r\n").encode())
        self.pilcrow("setup")
        self.assertIn(DIFF_LINE, self.block())

    def test_git_diff_shows_prose_only_once_configured(self) -> None:
        self.init_git()
        source = self.root / "essay.docx"
        make_docx(source, paragraph("Opening sentence. The cat sat on the mat.") + paragraph("Closing line."))
        config_before = (self.root / ".git" / "config").read_bytes()
        setup = self.pilcrow("setup", "--textconv")
        self.assertIn(CONFIG_COMMAND.encode(), setup.stdout)
        # Pilcrow prints the command; it never runs it.
        self.assertEqual((self.root / ".git" / "config").read_bytes(), config_before)
        self.assertEqual(self.git("check-attr", "diff", "--", "essay.docx"), "essay.docx: diff: pilcrow\n")
        self.git("add", ".")
        self.git("commit", "-qm", "Add essay without companions")
        make_docx(source, paragraph("Opening sentence. The dog sat on the mat.") + paragraph("Closing line."))

        # Without the Git config, the attribute alone changes nothing.
        diff = self.git("diff", "--", "essay.docx")
        self.assertIn("Binary files a/essay.docx and b/essay.docx differ", diff)
        self.assertNotIn("sat on the mat", diff)

        # Quote the interpreter for Git's shell; forward slashes also work on Windows.
        python = Path(sys.executable).as_posix()
        self.git("config", "diff.pilcrow.textconv", f"'{python}' -m pilcrow textconv")
        diff = self.git("diff", "--", "essay.docx")
        self.assertIn("-The cat sat on the mat.\n", diff)
        self.assertIn("+The dog sat on the mat.\n", diff)
        self.assertNotIn("Opening sentence.\n", [line[1:] + "\n" for line in diff.splitlines() if line[:1] in "+-"])
        self.git("commit", "-qam", "Edit essay")
        log = self.git("log", "-p", "--", "essay.docx")
        self.assertIn("+The dog sat on the mat.\n", log)
        self.assertIn("-The cat sat on the mat.\n", log)
        # The first revision, committed without companions, is readable too.
        self.assertIn("+Opening sentence.\n", log)
        self.assertIn("+Closing line.\n", log)
        self.assertFalse((self.root / ".pilcrow").exists())

        self.git("config", "--unset", "diff.pilcrow.textconv")
        self.assertIn("Binary files", self.git("show", "--", "essay.docx"))


if __name__ == "__main__":
    unittest.main()
