"""`pilcrow review`: a read-only, plain-language report of what changed between two DOCX files.

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

try:
    from .test_cli import (
        COMMENTS_EXT_REL,
        COMMENTS_REL,
        NUMBERING_REL,
        PROJECT,
        R,
        abstract_num,
        anchored,
        comments_extended,
        comments_part,
        end,
        image_rel,
        level,
        list_item,
        make_docx,
        num,
        numbering,
        paragraph,
        picture,
        run,
        start,
    )
    from .test_textconv import tree
except ImportError:  # discovered with -s tests: the directory itself is on sys.path
    from test_cli import (  # type: ignore[no-redef]
        COMMENTS_EXT_REL,
        COMMENTS_REL,
        NUMBERING_REL,
        PROJECT,
        R,
        abstract_num,
        anchored,
        comments_extended,
        comments_part,
        end,
        image_rel,
        level,
        list_item,
        make_docx,
        num,
        numbering,
        paragraph,
        picture,
        run,
        start,
    )
    from test_textconv import tree  # type: ignore[no-redef]

CAVEAT = "No supported prose or review changes detected (not a complete document comparison)."


class ReviewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-review-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.old = self.root / "old.docx"
        self.new = self.root / "new.docx"
        self.env = {**os.environ}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))

    def pilcrow(self, *args: str | Path, code: int = 0, cwd: Path | None = None) -> subprocess.CompletedProcess[bytes]:
        result = subprocess.run(
            [sys.executable, "-m", "pilcrow", *(str(arg) for arg in args)],
            cwd=cwd or self.root, env=self.env, capture_output=True, timeout=30,
        )
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout!r}\nstderr={result.stderr!r}")
        return result

    def review(self, code: int = 1) -> str:
        """Review old.docx -> new.docx by relative name and return the report."""
        result = self.pilcrow("review", "old.docx", "new.docx", code=code)
        self.assertEqual(result.stderr, b"")
        report = result.stdout.decode("utf-8")
        self.assertTrue(report.startswith("Review: old.docx → new.docx\n\n"), report)
        return report

    def section(self, report: str, title: str) -> list[str]:
        """The lines of one metadata section, without its title."""
        blocks = report.split("\n\n")
        block = next((b for b in blocks if b.startswith(title + ":")), None)
        self.assertIsNotNone(block, f"no {title} section in:\n{report}")
        assert block is not None
        return block.split("\n")[1:]

    # Prose and the no-change caveat

    def test_identical_documents_and_ignored_churn_report_no_changes(self) -> None:
        body = paragraph("A sentence. Another sentence.")
        make_docx(self.old, body)
        make_docx(self.new, body)
        report = self.review(code=0)
        self.assertEqual(report, f"Review: old.docx → new.docx\n\nProse: no changes\n\n{CAVEAT}\n")
        # Word save churn (part order, timestamps, docProps) is already ignored by extraction.
        make_docx(self.new, body, churn=True)
        self.assertNotEqual(self.old.read_bytes(), self.new.read_bytes())
        self.assertEqual(self.review(code=0), report)

    def test_a_sentence_edit_is_a_unified_diff(self) -> None:
        make_docx(self.old, paragraph("Opening sentence. The cat sat on the mat. Closing sentence."))
        make_docx(self.new, paragraph("Opening sentence. The dog sat on the mat. Closing sentence."))
        report = self.review()
        prose = self.section(report, "Prose")
        self.assertEqual(prose[:2], ["--- OLD", "+++ NEW"])
        self.assertIn("-The cat sat on the mat.", prose)
        self.assertIn("+The dog sat on the mat.", prose)
        self.assertIn(" Opening sentence.", prose)
        self.assertNotIn(CAVEAT, report)
        self.assertIn("not a complete document comparison", report)

    def test_a_very_large_prose_change_is_cut_at_the_cap(self) -> None:
        make_docx(self.old, "".join(paragraph(f"Old sentence {n}.") for n in range(1500)))
        make_docx(self.new, "".join(paragraph(f"New sentence {n}.") for n in range(1500)))
        prose = self.section(self.review(), "Prose")
        # Paragraphs are separated by blank lines: 6001 diff lines in all, 2000 shown.
        self.assertEqual(len(prose), 2001)
        self.assertEqual(prose[-1], "… 4001 more diff lines omitted")

    # Comments

    def threaded(self, path: Path, comments: dict[str, str], para_ids: dict[str, str],
                 *entries: tuple[str, str | None, str | None], body: str | None = None) -> None:
        if body is None:
            body = "".join(f"<w:p>{start(c)}{run(f'Passage {n}.')}{end(c)}</w:p>" for n, c in enumerate(comments))
        make_docx(path, body,
                  extra={"word/comments.xml": comments_part(comments, para_ids),
                         "word/commentsExtended.xml": comments_extended(*entries)},
                  relationships=COMMENTS_REL + COMMENTS_EXT_REL)

    def test_resolving_and_reopening_a_comment(self) -> None:
        comments, ids = {"1": "First.", "2": "Second."}, {"1": "11111111", "2": "22222222"}
        self.threaded(self.old, comments, ids, ("11111111", None, "0"), ("22222222", None, "0"))
        self.threaded(self.new, comments, ids, ("11111111", None, "1"), ("22222222", None, "0"))
        report = self.review()
        self.assertIn("Prose: no changes", report)
        self.assertEqual(self.section(report, "Comments"),
                         ['~ resolved comment by Ada on "Passage 0.": "First."'])
        self.threaded(self.old, comments, ids, ("11111111", None, "1"), ("22222222", None, "0"))
        self.threaded(self.new, comments, ids, ("11111111", None, "0"), ("22222222", None, "0"))
        self.assertEqual(self.section(self.review(), "Comments"),
                         ['~ reopened comment by Ada on "Passage 0.": "First."'])

    def test_indistinguishable_comments_never_claim_resolved_or_reopened(self) -> None:
        # Same author, text and anchor: the metadata cannot say which one Word resolved.
        body = "".join(f"<w:p>{start(c)}{run('Same passage.')}{end(c)}</w:p>" for c in ("1", "2"))
        comments, ids = {"1": "Same note.", "2": "Same note."}, {"1": "11111111", "2": "22222222"}
        self.threaded(self.old, comments, ids, ("11111111", None, "0"), ("22222222", None, "0"), body=body)
        self.threaded(self.new, comments, ids, ("11111111", None, "1"), ("22222222", None, "0"), body=body)
        lines = self.section(self.review(), "Comments")
        self.assertFalse([line for line in lines if line.startswith("~")], lines)
        self.assertFalse([line for line in lines if "resolved comment" in line or "reopened" in line], lines)
        self.assertTrue(lines, lines)
        for line in lines:
            self.assertTrue(line.startswith(("- removed comment", "+ added comment")), lines)

    def test_unknown_resolved_state_is_never_claimed(self) -> None:
        comments, ids = {"1": "First."}, {"1": "11111111"}
        # No w15:done in OLD, so the old state is unknown, not "open".
        self.threaded(self.old, comments, ids, ("99999999", None, None))
        self.threaded(self.new, comments, ids, ("11111111", None, "1"))
        lines = self.section(self.review(), "Comments")
        self.assertEqual(lines, ['- removed comment by Ada on "Passage 0.": "First."',
                                 '+ added comment by Ada on "Passage 0.": "First." (resolved)'])

    def test_reply_moved_range_and_edited_text_are_described(self) -> None:
        comments, ids = {"1": "First.", "2": "Second."}, {"1": "11111111", "2": "22222222"}
        self.threaded(self.old, comments, ids, ("11111111", None, "0"), ("22222222", None, "0"))
        replied = {**comments, "3": "A reply."}
        self.threaded(self.new, replied, {**ids, "3": "33333333"},
                      ("11111111", None, "0"), ("22222222", None, "0"), ("33333333", "11111111", "0"))
        self.assertEqual(self.section(self.review(), "Comments"),
                         ['+ added reply by Ada to Ada\'s comment "First.": "A reply." (open)'])

        def ranged(path: Path, before: str, anchored: str, after: str, text: str = "Note.") -> None:
            make_docx(path, f"<w:p>{run(before)}{start('1')}{run(anchored)}{end('1')}{run(after)}</w:p>",
                      extra={"word/comments.xml": comments_part({"1": text})}, relationships=COMMENTS_REL)
        ranged(self.old, "The ", "first sentence.", " The second sentence.")
        ranged(self.new, "The first sentence. The ", "second sentence.", "")
        report = self.review()
        self.assertIn("Prose: no changes", report)
        self.assertEqual(self.section(report, "Comments"),
                         ['~ re-anchored comment by Ada on "first sentence.": "Note." → now on "second sentence."'])
        ranged(self.new, "The ", "first sentence.", " The second sentence.", text="Please clarify.")
        self.assertEqual(self.section(self.review(), "Comments"),
                         ['~ edited text of comment by Ada on "first sentence.": "Note." → "Please clarify."'])

    def test_inserting_an_earlier_comment_does_not_rewrite_later_replies(self) -> None:
        def body(comments: dict[str, str]) -> str:
            return "".join(f"<w:p>{start(c)}{run(f'Passage {c}.')}{end(c)}</w:p>" for c in comments)
        comments, ids = {"1": "First.", "2": "Reply."}, {"1": "11111111", "2": "22222222"}
        self.threaded(self.old, comments, ids, ("11111111", None, "0"), ("22222222", "11111111", "0"),
                      body=body(comments))
        # A new comment before the thread shifts the reply's raw reply_to position from 0 to 1.
        comments, ids = {"0": "Earlier.", **comments}, {"0": "00000000", **ids}
        self.threaded(self.new, comments, ids, ("00000000", None, "0"), ("11111111", None, "0"),
                      ("22222222", "11111111", "0"), body=body(comments))
        self.assertEqual(self.section(self.review(), "Comments"),
                         ['+ added comment by Ada on "Passage 0.": "Earlier." (open)'])

    def test_long_comment_text_is_shortened(self) -> None:
        long = "word " * 40
        make_docx(self.old, paragraph("Body."))
        make_docx(self.new, f"<w:p>{start('1')}{run('Body.')}{end('1')}</w:p>",
                  extra={"word/comments.xml": comments_part({"1": long})}, relationships=COMMENTS_REL)
        (line,) = self.section(self.review(), "Comments")
        quoted = line.split(": ", 1)[1]
        self.assertEqual(len(quoted), 82)  # 80 characters and two quotes
        self.assertTrue(quoted.endswith('…"'), quoted)

    # Other review metadata

    def test_tracked_insertion_by_a_new_author(self) -> None:
        make_docx(self.old, paragraph("Unchanged."))
        make_docx(self.new, paragraph("Unchanged.")
                  + '<w:p><w:ins w:id="1" w:author="Ben"><w:r><w:t>Inserted words.</w:t></w:r></w:ins></w:p>')
        report = self.review()
        self.assertIn("+Inserted words.", self.section(report, "Prose"))
        self.assertEqual(self.section(report, "Tracked changes"),
                         ['+ added tracked insertion by Ben: "Inserted words."'])

    def test_hyperlink_target_change(self) -> None:
        body = '<w:p><w:hyperlink r:id="link"><w:r><w:t>Visible label</w:t></w:r></w:hyperlink></w:p>'
        for path, target in ((self.old, "https://example.invalid/first"), (self.new, "https://example.invalid/second")):
            make_docx(path, body, relationships=f'<Relationship Id="link" Type="{R}/hyperlink" Target="{target}" TargetMode="External"/>')
        report = self.review()
        self.assertIn("Prose: no changes", report)
        self.assertEqual(self.section(report, "Hyperlinks"),
                         ['~ link "Visible label": https://example.invalid/first → https://example.invalid/second'])

    def test_image_swap(self) -> None:
        for path, data in ((self.old, "first picture"), (self.new, "second picture")):
            make_docx(path, anchored("Figure:", picture("rId5")), extra={"word/media/image1.png": data},
                      relationships=image_rel("rId5", "media/image1.png"))
        report = self.review()
        self.assertIn("Prose: no changes", report)
        self.assertEqual(self.section(report, "Images"), ["~ image 1 (word/document.xml) replaced"])

    def test_bullet_to_number_list_format_change(self) -> None:
        body = list_item("First.", "1") + list_item("Second.", "1")
        for path, lvl in ((self.old, level(0, "bullet", "•")), (self.new, level(0, "decimal", "%1."))):
            make_docx(path, body, extra={"word/numbering.xml": numbering(abstract_num("0", lvl), num("1", "0"))},
                      relationships=NUMBERING_REL)
        report = self.review()
        self.assertIn("Prose: no changes", report)
        self.assertEqual(self.section(report, "Lists"),
                         ['~ list 1 level 0: bullet → decimal; level text "•" → "%1."'])

    def test_a_warning_only_in_new_is_listed(self) -> None:
        make_docx(self.old, paragraph("Figure:"))
        make_docx(self.new, anchored("Figure:", picture("rId5")), extra={"word/media/image1.png": "picture"},
                  relationships=image_rel("rId5", "media/image1.png"))
        report = self.review()
        warnings = self.section(report, "Warnings")
        self.assertTrue(warnings, report)
        self.assertTrue(all(line.startswith("+ added: ") for line in warnings), warnings)
        self.assertIn("+ added: Images are compared by content hash only; size, position, cropping, and alt text are not compared.",
                      warnings)
        footer = self.section(report, "Coverage warnings (either side)")
        self.assertEqual(len(footer), len(set(footer)))

    # Inputs, errors and side effects

    def test_missing_or_malformed_input_exits_2_with_no_report(self) -> None:
        make_docx(self.old, paragraph("Prose."))
        result = self.pilcrow("review", "old.docx", "absent.docx", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertIn(b"error: absent.docx: ", result.stderr)
        (self.root / "broken.docx").write_bytes(b"not a zip")
        result = self.pilcrow("review", "old.docx", "broken.docx", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertIn(b"error: broken.docx: ", result.stderr)
        self.assertNotIn(b"Traceback", result.stderr)
        self.pilcrow("review", "old.docx", code=2)

    def test_review_writes_nothing_and_ignores_scope(self) -> None:
        inputs = self.root / "inputs"
        make_docx(inputs / "a.tmp", paragraph("Before."))  # git difftool names: any extension
        make_docx(inputs / "b", paragraph("After."))
        before = tree(self.root)
        elsewhere = self.root / "cwd"
        elsewhere.mkdir()
        listing = tree(self.root)
        self.pilcrow("review", inputs / "a.tmp", inputs / "b", code=1, cwd=elsewhere)
        self.pilcrow("review", inputs / "a.tmp", inputs / "a.tmp", code=0, cwd=elsewhere)
        self.assertEqual(tree(self.root), listing)
        self.assertNotEqual(before, listing)  # only the cwd directory was added
        self.assertFalse(any(path.name == ".pilcrow" for path in self.root.rglob("*")))

    def test_output_is_deterministic_and_independent_of_file_names(self) -> None:
        make_docx(self.old, paragraph("One sentence."))
        make_docx(self.new, paragraph("Another sentence."), relationships=image_rel("rId5", "missing.png"))
        first = self.review()
        self.assertEqual(self.review(), first)
        shutil.copy(self.old, self.root / "x.docx")
        shutil.copy(self.new, self.root / "y.docx")
        renamed = self.pilcrow("review", "x.docx", "y.docx", code=1).stdout.decode("utf-8")
        self.assertEqual(renamed.split("\n", 1)[1], first.split("\n", 1)[1])

    def test_git_difftool_usage(self) -> None:
        if shutil.which("git") is None:
            self.skipTest("git is not installed")
        repo = self.root / "repo"
        repo.mkdir()

        def git(*args: str) -> subprocess.CompletedProcess[bytes]:
            return subprocess.run(["git", *args], cwd=repo, env=self.env, capture_output=True, timeout=60)
        for args in (("init", "-q"), ("config", "user.name", "Pilcrow Test"),
                     ("config", "user.email", "pilcrow@example.invalid"), ("config", "commit.gpgsign", "false")):
            self.assertEqual(git(*args).returncode, 0)
        make_docx(repo / "report.docx", paragraph("The cat sat."))
        git("add", "report.docx")
        self.assertEqual(git("commit", "-q", "-m", "first").returncode, 0)
        make_docx(repo / "report.docx", paragraph("The dog sat."))
        git("commit", "-q", "-am", "second")
        command = f'"{Path(sys.executable).as_posix()}" -m pilcrow review'
        result = git("difftool", "-y", "-x", command, "HEAD~1", "--", "report.docx")
        output = result.stdout.decode("utf-8")
        self.assertIn("-The cat sat.", output)
        self.assertIn("+The dog sat.", output)
        self.assertFalse((repo / ".pilcrow").exists())


if __name__ == "__main__":
    unittest.main()
