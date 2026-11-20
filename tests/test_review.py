"""`pilcrow review`: a read-only, plain-language report of what changed between two DOCX files.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time
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
                         ['+ added reply by Ada to Ada\'s comment "First." on "Passage 0.": "A reply." (open)'])

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

    def test_moving_a_reply_between_look_alike_parents_is_reported(self) -> None:
        def write(path: Path, parent: str, other: str, same_anchor: bool) -> None:
            comments = {"1": "Same.", "2": "Same.", "3": "Reply.", "4": other}
            ids = {"1": "11111111", "2": "22222222", "3": "33333333", "4": "44444444"}
            body = "".join(f"<w:p>{start(c)}{run('Same passage.' if same_anchor and c in '12' else f'Passage {c}.')}"
                           f"{end(c)}</w:p>" for c in comments)
            self.threaded(path, comments, ids, ("11111111", None, "0"), ("22222222", None, "0"),
                          ("33333333", parent, "0"), ("44444444", None, "0"), body=body)
        # Parents with the same author and text but different anchors; another comment also changes.
        write(self.old, "11111111", "Other.", same_anchor=False)
        write(self.new, "22222222", "Other, edited.", same_anchor=False)
        lines = self.section(self.review(), "Comments")
        self.assertIn('- removed reply by Ada to Ada\'s comment "Same." on "Passage 1.": "Reply." (open)', lines)
        self.assertIn('+ added reply by Ada to Ada\'s comment "Same." on "Passage 2.": "Reply." (open)', lines)
        self.assertIn('~ edited text of comment by Ada on "Passage 4.": "Other." → "Other, edited."', lines)
        # Parents the metadata cannot tell apart at all: their position tells the replies apart.
        write(self.old, "11111111", "Other.", same_anchor=True)
        write(self.new, "22222222", "Other, edited.", same_anchor=True)
        lines = self.section(self.review(), "Comments")
        self.assertIn('- removed reply by Ada to Ada\'s comment "Same." on "Same passage." (comment 1): "Reply." (open)', lines)
        self.assertIn('+ added reply by Ada to Ada\'s comment "Same." on "Same passage." (comment 2): "Reply." (open)', lines)

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

    def test_an_image_with_embedded_and_linked_copies(self) -> None:
        # Word can keep an embedded copy (r:embed) and an external link (r:link) for one image.
        both = picture("rId5").replace('r:embed="rId5"', 'r:embed="rId5" r:link="rId6"')

        def write(path: Path, data: str, target: str) -> None:
            make_docx(path, anchored("Figure:", both), extra={"word/media/image1.png": data},
                      relationships=image_rel("rId5", "media/image1.png") + image_rel("rId6", target, external=True))
        write(self.old, "first picture", "https://example.invalid/a.png")
        write(self.new, "second picture", "https://example.invalid/a.png")
        self.assertEqual(self.section(self.review(), "Images"), ["~ image 1 (word/document.xml) replaced"])
        write(self.new, "first picture", "https://example.invalid/b.png")
        self.assertEqual(self.section(self.review(), "Images"),
                         ["~ image 1 (word/document.xml) link changed: https://example.invalid/a.png → https://example.invalid/b.png"])
        write(self.new, "second picture", "https://example.invalid/b.png")
        self.assertEqual(self.section(self.review(), "Images"),
                         ["~ image 1 (word/document.xml) replaced; link changed: https://example.invalid/a.png → https://example.invalid/b.png"])
        # A dangling embedded reference next to the link: no hash to compare, so both sides are listed.
        dangling = picture("rId9").replace('r:embed="rId9"', 'r:embed="rId9" r:link="rId6"')
        make_docx(self.old, anchored("Figure:", picture("rId6", "link")),
                  relationships=image_rel("rId6", "https://example.invalid/b.png", external=True))
        make_docx(self.new, anchored("Figure:", dangling),
                  relationships=image_rel("rId6", "https://example.invalid/b.png", external=True))
        self.assertEqual(self.section(self.review(), "Images"),
                         ["- removed image 1 (word/document.xml): linked to https://example.invalid/b.png",
                          "+ added image 1 (word/document.xml): with a missing media part, linked to https://example.invalid/b.png"])
        write(self.new, "second picture", "https://example.invalid/b.png")
        make_docx(self.old, paragraph("Figure:"))
        (line,) = self.section(self.review(), "Images")
        self.assertRegex(line, r"^\+ added image 1 \(word/document\.xml\): 14 bytes, sha256 [0-9a-f]{12}, "
                               r"linked to https://example\.invalid/b\.png$")

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

    def test_long_repetitive_metadata_is_aligned_quickly(self) -> None:
        # 4,000 identical field records: unchanged ones must not be aligned pairwise.
        field = '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'
        make_docx(self.old, field * 4000 + paragraph("Old ending."))
        make_docx(self.new, field * 4000 + paragraph("New ending."))
        started = time.monotonic()
        report = self.review()
        self.assertLess(time.monotonic() - started, 15)
        self.assertNotIn("Fields:", report)
        make_docx(self.new, field * 3999 + paragraph("Old ending."))
        self.assertEqual(self.section(self.review(), "Fields"), ["- removed field 4000 (PAGE) in word/document.xml"])

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


class RevisionReviewTests(unittest.TestCase):
    """`pilcrow review --rev`: the old side, or both, read from Git with read-only plumbing."""

    def setUp(self) -> None:
        if shutil.which("git") is None:
            self.skipTest("git is not installed")
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-review-rev-")
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.repo = base / "repo"
        self.repo.mkdir()
        self.outside = base / "outside"  # not in any repository: GIT_CEILING_DIRECTORIES stops the search
        self.outside.mkdir()
        # An empty temporary directory for every process, to show nothing is written there.
        self.tmp = base / "tmp"
        self.tmp.mkdir()
        self.env = {**os.environ, "TMPDIR": str(self.tmp), "TEMP": str(self.tmp), "TMP": str(self.tmp),
                    "GIT_CEILING_DIRECTORIES": str(base)}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))
        for args in (("init", "-q"), ("config", "user.name", "Pilcrow Test"),
                     ("config", "user.email", "pilcrow@example.invalid"), ("config", "commit.gpgsign", "false")):
            self.git(*args)

    def git(self, *args: str) -> str:
        result = subprocess.run(["git", *args], cwd=self.repo, env=self.env, capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.decode("utf-8")

    def commit(self, message: str) -> None:
        self.git("add", "-A")
        self.git("commit", "-q", "-m", message)

    def pilcrow(self, *args: str | Path, code: int, cwd: Path | None = None,
                limit: int | None = None) -> subprocess.CompletedProcess[bytes]:
        command = [sys.executable, "-m", "pilcrow"]
        if limit is not None:  # lower the package size limit in the child process
            command = [sys.executable, "-c", "import sys; import pilcrow.cli as cli; "
                       f"cli.MAX_PACKAGE_BYTES = {limit}; sys.exit(cli.main(sys.argv[1:]))"]
        result = subprocess.run([*command, *(str(arg) for arg in args)], cwd=cwd or self.repo, env=self.env,
                                capture_output=True, timeout=60)
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout!r}\nstderr={result.stderr!r}")
        return result

    def state(self) -> tuple[object, ...]:
        """The work tree, the index and the temporary directory, read without refreshing the index."""
        index = self.repo / ".git" / "index"
        status = subprocess.run(["git", "--no-optional-locks", "status", "--porcelain"], cwd=self.repo,
                                env=self.env, capture_output=True, timeout=60).stdout
        return tree(self.repo), index.stat().st_mtime_ns, status, tree(self.tmp)

    def threaded(self, path: Path, text: str, done: str) -> None:
        comments, ids = {"1": "Please check."}, {"1": "11111111"}
        body = paragraph(text) + f"<w:p>{start('1')}{run('A commented passage.')}{end('1')}</w:p>"
        make_docx(path, body,
                  extra={"word/comments.xml": comments_part(comments, ids),
                         "word/commentsExtended.xml": comments_extended(("11111111", None, done))},
                  relationships=COMMENTS_REL + COMMENTS_EXT_REL)

    def test_an_unchanged_work_tree_reports_no_changes(self) -> None:
        make_docx(self.repo / "report.docx", paragraph("A sentence."))
        self.commit("first")
        before = self.state()
        result = self.pilcrow("review", "--rev", "HEAD", "report.docx", code=0)
        self.assertEqual(result.stderr, b"")
        self.assertEqual(result.stdout.decode("utf-8"),
                         f"Review: HEAD:report.docx → report.docx\n\nProse: no changes\n\n{CAVEAT}\n")
        self.assertEqual(self.state(), before)

    def test_work_tree_edits_match_a_review_of_two_files(self) -> None:
        document = self.repo / "report.docx"
        self.threaded(document, "The cat sat. It was warm.", "0")
        self.commit("first")
        committed = Path(self.temp.name) / "old.docx"
        shutil.copy(document, committed)
        self.threaded(document, "The dog sat. It was warm.", "1")  # edit the prose and resolve the comment
        before = self.state()
        result = self.pilcrow("review", "--rev", "HEAD", "report.docx", code=1)
        self.assertEqual(self.state(), before)
        header, body = result.stdout.split(b"\n", 1)
        self.assertEqual(header.decode("utf-8"), "Review: HEAD:report.docx → report.docx")
        files = self.pilcrow("review", committed, "report.docx", code=1)
        self.assertEqual(body, files.stdout.split(b"\n", 1)[1])
        self.assertIn(b"-The cat sat.", body)
        self.assertIn(b"~ resolved comment by Ada", body)

    def test_two_revisions_ignore_the_work_tree(self) -> None:
        document = self.repo / "report.docx"
        make_docx(document, paragraph("The cat sat."))
        self.commit("first")
        make_docx(document, paragraph("The dog sat."))
        self.commit("second")
        make_docx(document, paragraph("An uncommitted edit."))
        before = self.state()
        report = self.pilcrow("review", "--rev", "HEAD~1", "--rev", "HEAD", "report.docx", code=1).stdout.decode("utf-8")
        self.assertTrue(report.startswith("Review: HEAD~1:report.docx → HEAD:report.docx\n\n"), report)
        self.assertIn("-The cat sat.", report)
        self.assertIn("+The dog sat.", report)
        self.assertNotIn("uncommitted", report)
        self.assertEqual(self.state(), before)
        # The work-tree file is not read at all, so it need not exist.
        document.unlink()
        self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", "report.docx", code=0)
        branch = self.git("rev-parse", "--abbrev-ref", "HEAD").strip()
        self.git("checkout", "-q", "-b", "feature", "HEAD~1")
        report = self.pilcrow("review", "--rev", branch, "--rev", "feature", "report.docx", code=1).stdout.decode("utf-8")
        self.assertTrue(report.startswith(f"Review: {branch}:report.docx → feature:report.docx\n\n"), report)
        self.assertIn("-The dog sat.", report)

    def test_subdirectories_spaces_and_unicode(self) -> None:
        folder = self.repo / "drafts and notes" / "İstanbul"
        document = folder / "résumé final.docx"
        make_docx(document, paragraph("Old wording."))
        make_docx(self.repo / "résumé final.docx", paragraph("A different file at the top."))
        self.commit("first")
        make_docx(document, paragraph("New wording."))
        relative = Path("drafts and notes") / "İstanbul" / "résumé final.docx"
        for cwd, path in ((self.repo, relative), (folder, Path("résumé final.docx")),
                          (folder.parent, Path("İstanbul") / "résumé final.docx"),
                          (self.outside, document)):
            with self.subTest(cwd=cwd.name, path=str(path)):
                report = self.pilcrow("review", "--rev", "HEAD", path, cwd=cwd, code=1).stdout.decode("utf-8")
                self.assertTrue(report.startswith(f"Review: HEAD:{path} → {path}\n\n"), report)
                self.assertIn("-Old wording.", report)
                self.assertIn("+New wording.", report)
        # From a subdirectory, a bare name means the file there, not the one at the top of the repository.
        self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", "résumé final.docx", cwd=folder, code=0)
        report = self.pilcrow("review", "--rev", "HEAD", "résumé final.docx", cwd=folder, code=1).stdout
        self.assertNotIn(b"A different file", report)

    def test_errors_name_the_input_print_no_report_and_write_nothing(self) -> None:
        document = self.repo / "report.docx"
        make_docx(document, paragraph("First."))
        self.commit("first")
        make_docx(self.repo / "added.docx", paragraph("Added later."))
        (self.repo / "folder.docx").mkdir()
        make_docx(self.repo / "folder.docx" / "inner.docx", paragraph("Inside."))
        self.commit("second")
        shutil.copy(document, self.outside / "report.docx")
        size = int(self.git("cat-file", "-s", "HEAD:report.docx"))
        before = self.state()
        cases: list[tuple[tuple[str, ...], Path, int | None, bytes]] = [
            (("--rev", "HEAD", "report.docx"), self.outside, None, b"error: HEAD:report.docx: not a git repository"),
            (("--rev", "no-such-revision", "report.docx"), self.repo, None,
             b"error: no-such-revision:report.docx: unknown revision"),
            (("--rev", "HEAD~1", "added.docx"), self.repo, None, b"error: HEAD~1:added.docx: no such file at HEAD~1"),
            (("--rev", "HEAD", "folder.docx"), self.repo, None, b"error: HEAD:folder.docx: not a file at HEAD"),
            (("--rev=-c", "report.docx"), self.repo, None, b'error: -c:report.docx: a revision must not start with "-"'),
            (("--rev=--output=x", "report.docx"), self.repo, None, b"error: --output=x:report.docx: a revision must"),
            (("--rev", "HEAD", "report.docx"), self.repo, size - 1,
             f"error: HEAD:report.docx: file is {size} bytes at HEAD; the safety limit is {size - 1} bytes".encode()),
            (("--rev", "HEAD", "report.docx", "report.docx"), self.repo, None, b"--rev takes one PATH, not 2"),
            (("--rev", "HEAD", "--rev", "HEAD", "--rev", "HEAD", "report.docx"), self.repo, None,
             b"give --rev at most twice"),
        ]
        for args, cwd, limit, message in cases:
            with self.subTest(args=args, limit=limit):
                result = self.pilcrow("review", *args, cwd=cwd, limit=limit, code=2)
                self.assertEqual(result.stdout, b"")
                self.assertIn(message, result.stderr)
                self.assertNotIn(b"Traceback", result.stderr)
                self.assertEqual(self.state(), before)
        self.assertFalse((self.repo / "x").exists())
        # A malformed committed DOCX is an error too, with the same label.
        (self.repo / "broken.docx").write_bytes(b"not a zip")
        self.commit("third")
        result = self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", "broken.docx", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertIn(b"error: HEAD:broken.docx: cannot read DOCX", result.stderr)
        # The limit is on the committed blob's size: a blob exactly at the limit still reads.
        self.pilcrow("review", "--rev", "HEAD", "report.docx", limit=size, code=0)


if __name__ == "__main__":
    unittest.main()
