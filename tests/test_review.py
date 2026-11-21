"""`pilcrow review`: a read-only, plain-language report of what changed between two DOCX files.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import os
import random
import re
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
MARKERS = {"- ", "~ ", "+ "}


def assert_same_report(test: unittest.TestCase, text: str, markdown: str) -> None:
    """Check that a Markdown report says what the text report says, block for block.

    The header becomes a heading, the prose diff sits unchanged in a fenced diff block,
    and each metadata section is a subheading with one bullet per text line, its
    -/~/+ marker in a code span.
    """
    blocks = text.rstrip("\n").split("\n\n")
    header = blocks.pop(0)
    old, new = header.removeprefix("Review: ").split(" → ")
    md = markdown.rstrip("\n").split("\n\n")
    test.assertEqual(md.pop(0), f"## Review: `{old}` → `{new}`")
    prose = blocks.pop(0)
    if prose == "Prose: no changes":
        test.assertEqual(md.pop(0), prose)
    else:
        # Unified diff lines are never empty (a blank context line is " "), so the
        # diff and its fence are one block, like the text's "Prose:" block.
        test.assertEqual(md.pop(0), "### Prose")
        fenced = md.pop(0).split("\n")
        fence = fenced[-1]
        test.assertEqual(set(fence), {"`"})
        test.assertGreaterEqual(len(fence), 3)
        test.assertGreater(len(fence), max((len(run) for run in re.findall("`+", prose)), default=0))
        test.assertEqual(fenced, [f"{fence}diff", *prose.split("\n")[1:], fence])
    test.assertEqual(md.pop(), blocks.pop())  # the closing caveat has nothing to escape
    for block in blocks:
        first, *lines = block.split("\n")
        if lines:
            title = first.removesuffix(":")
        else:  # a one-line section such as "Metadata: key changed"
            title, inline = first.split(": ", 1)
            lines = [inline]
        test.assertEqual(md.pop(0), f"### {title}")
        bullets = md.pop(0).split("\n")
        test.assertEqual(len(bullets), len(lines), block)
        for line, item in zip(lines, bullets, strict=True):
            prefix = f"- `{line[0]}` " if line[:2] in MARKERS else "- "
            test.assertTrue(item.startswith(prefix), item)
    test.assertEqual(md, [])


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
        """Review old.docx -> new.docx by relative name and return the report.

        Every review is also printed as Markdown, which must reach the same verdict
        and say the same things (see ``assert_same_report``).
        """
        result = self.pilcrow("review", "old.docx", "new.docx", code=code)
        self.assertEqual(result.stderr, b"")
        report = result.stdout.decode("utf-8")
        self.assertTrue(report.startswith("Review: old.docx → new.docx\n\n"), report)
        markdown = self.pilcrow("review", "--format", "markdown", "old.docx", "new.docx", code=code)
        self.assertEqual(markdown.stderr, b"")
        assert_same_report(self, report, markdown.stdout.decode("utf-8"))
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

    # Markdown

    def markdown(self, old: str = "old.docx", new: str = "new.docx", code: int = 1) -> str:
        result = self.pilcrow("review", "--format", "markdown", old, new, code=code)
        self.assertEqual(result.stderr, b"")
        return result.stdout.decode("utf-8")

    def test_text_is_the_default_format_and_unchanged(self) -> None:
        make_docx(self.old, paragraph("Opening sentence. The cat sat on the mat."))
        make_docx(self.new, paragraph("Opening sentence. The dog sat on the mat."))
        changed = ("Review: old.docx → new.docx\n\nProse:\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n"
                   " Opening sentence.\n-The cat sat on the mat.\n+The dog sat on the mat.\n\n"
                   "Coverage warnings (either side):\n"
                   "  Formatting and page layout are not compared; unchanged prose is not an identical document.\n\n"
                   "Only the prose and review metadata Pilcrow extracts are compared"
                   " (not a complete document comparison).\n")
        unchanged = f"Review: old.docx → old.docx\n\nProse: no changes\n\n{CAVEAT}\n"
        for new, code, expected in (("new.docx", 1, changed), ("old.docx", 0, unchanged)):
            default = self.pilcrow("review", "old.docx", new, code=code).stdout
            self.assertEqual(default, expected.encode("utf-8"))
            self.assertEqual(self.pilcrow("review", "--format", "text", "old.docx", new, code=code).stdout, default)
            self.assertEqual(self.pilcrow("review", "old.docx", new, "--format=text", code=code).stdout, default)

    def test_markdown_report_shape(self) -> None:
        make_docx(self.old, paragraph("Opening sentence. The cat sat on the mat."))
        make_docx(self.new, paragraph("Opening sentence. The dog sat on the mat."))
        self.assertEqual(self.markdown(), (
            "## Review: `old.docx` → `new.docx`\n\n### Prose\n\n```diff\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n"
            " Opening sentence.\n-The cat sat on the mat.\n+The dog sat on the mat.\n```\n\n"
            "### Coverage warnings (either side)\n\n"
            "- Formatting and page layout are not compared; unchanged prose is not an identical document.\n\n"
            "Only the prose and review metadata Pilcrow extracts are compared (not a complete document comparison).\n"))
        self.assertEqual(self.markdown("old.docx", "old.docx", code=0),
                         f"## Review: `old.docx` → `old.docx`\n\nProse: no changes\n\n{CAVEAT}\n")

    def test_an_invalid_format_is_a_usage_error(self) -> None:
        make_docx(self.old, paragraph("Prose."))
        result = self.pilcrow("review", "--format", "html", "old.docx", "old.docx", code=2)
        self.assertEqual(result.stdout, b"")
        for word in (b"--format", b"html", b"text", b"markdown", b"usage:"):
            self.assertIn(word, result.stderr)

    def test_the_diff_fence_is_longer_than_any_backtick_run_in_the_prose(self) -> None:
        make_docx(self.old, paragraph("Code: ``` here.") + paragraph("Before."))
        make_docx(self.new, paragraph("Code: ```` here.") + paragraph("```") + paragraph("````"))
        text = self.review()  # also checks the fenced block holds exactly the text diff
        report = self.markdown()
        lines = report.split("\n")
        opening = lines.index("`````diff")
        closing = lines.index("`````", opening)
        self.assertEqual(lines[opening + 1:closing], self.section(text, "Prose"))
        self.assertIn("+````", lines[opening + 1:closing])
        # Nothing inside could close a fence of five backticks.
        self.assertFalse(any(set(line.strip()) == {"`"} and len(line.strip()) >= 5 for line in lines[opening + 1:closing]))

    def test_hostile_document_text_is_escaped_outside_the_diff(self) -> None:
        hostile = "# heading *emphasis* [link](x) ![img](x) <script> &amp; | `tick"
        link = f'<w:p><w:hyperlink r:id="link">{run(hostile)}</w:hyperlink></w:p>'
        make_docx(self.old, paragraph("Body."))
        make_docx(self.new, f"<w:p>{start('1')}{run('Body.')}{end('1')}</w:p>{link}"
                  '<w:p><w:ins w:id="7" w:author="Mallory &lt;b&gt;|*x*"><w:r><w:t>Inserted.</w:t></w:r></w:ins></w:p>',
                  extra={"word/comments.xml": comments_part({"1": hostile})},
                  relationships=COMMENTS_REL + f'<Relationship Id="link" Type="{R}/hyperlink" '
                                               'Target="https://example.invalid/a_b*c" TargetMode="External"/>')
        self.review()
        report = self.markdown()
        fence = report.index("```diff\n")
        outside = report[:fence] + report[report.index("\n```\n", fence) + 5:]
        escaped = r'"\# heading \*emphasis\* \[link\](x) \!\[img\](x) &lt;script&gt; &amp;amp; \| \`tick"'
        for title, marker in (("Comments", "comment by Ada"), ("Hyperlinks", "link")):
            section = outside.split(f"### {title}\n\n", 1)[1].split("\n\n", 1)[0]
            self.assertTrue(section.startswith(f"- `+` added {marker}"), section)
            self.assertIn(escaped, section)
        self.assertIn(r"→ https\://example.invalid/a\_b\*c", outside)
        self.assertIn(r"by Mallory &lt;b&gt;\|\*x\*\: ", outside)
        # Nothing from the document survives unescaped outside the fence.
        for raw in ("<script>", "[link](x)", "*emphasis*", "| `", "&amp; ", "<b>"):
            self.assertNotIn(raw, outside)
        self.assertEqual([line for line in outside.split("\n") if line.startswith("#")],
                         ["## Review: `old.docx` → `new.docx`", "### Prose", "### Comments", "### Tracked changes",
                          "### Hyperlinks", "### Warnings", "### Coverage warnings (either side)"])
        # The prose diff itself is verbatim.
        self.assertIn(f"+{hostile}\n", report[fence:])

    def test_markdown_header_labels_are_code_spans(self) -> None:
        make_docx(self.old, paragraph("Same."))
        shutil.copy(self.old, self.root / "a `b` #.docx")
        report = self.markdown("a `b` #.docx", "old.docx", code=0)
        self.assertTrue(report.startswith("## Review: ``a `b` #.docx`` → `old.docx`\n\n"), report)

    # Word marks (--word-diff)

    def test_word_diff_marks_the_changed_words_of_a_sentence(self) -> None:
        make_docx(self.old, paragraph("The supplier shall deliver the goods within 30 days. Payment is due on receipt."))
        make_docx(self.new, paragraph("The supplier may deliver the goods within 60 days. Payment is due on receipt."))
        tail = ("Coverage warnings (either side):\n"
                "  Formatting and page layout are not compared; unchanged prose is not an identical document.\n\n"
                "Only the prose and review metadata Pilcrow extracts are compared (not a complete document comparison).\n")
        result = self.pilcrow("review", "--word-diff", "old.docx", "new.docx", code=1)
        self.assertEqual(result.stderr, b"")
        self.assertEqual(result.stdout.decode("utf-8"), (
            "Review: old.docx → new.docx\n\nProse:\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n"
            "~ The supplier [-shall-]{+may+} deliver the goods within [-30-]{+60+} days.\n"
            " Payment is due on receipt.\n\n" + tail))
        # Without the flag, the same documents give the whole-line diff, as before.
        plain = self.pilcrow("review", "old.docx", "new.docx", code=1).stdout.decode("utf-8")
        self.assertEqual(plain, (
            "Review: old.docx → new.docx\n\nProse:\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n"
            "-The supplier shall deliver the goods within 30 days.\n"
            "+The supplier may deliver the goods within 60 days.\n"
            " Payment is due on receipt.\n\n" + tail))
        markdown = self.pilcrow("review", "--format", "markdown", "--word-diff", "old.docx", "new.docx", code=1)
        self.assertEqual(markdown.stderr, b"")
        self.assertEqual(markdown.stdout.decode("utf-8"), (
            "## Review: `old.docx` → `new.docx`\n\n### Prose\n\n```diff\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n"
            "~ The supplier [-shall-]{+may+} deliver the goods within [-30-]{+60+} days.\n"
            " Payment is due on receipt.\n```\n\n"
            "### Coverage warnings (either side)\n\n"
            "- Formatting and page layout are not compared; unchanged prose is not an identical document.\n\n"
            "Only the prose and review metadata Pilcrow extracts are compared (not a complete document comparison).\n"))
        assert_same_report(self, result.stdout.decode("utf-8"), markdown.stdout.decode("utf-8"))
        # The flag changes no exit code: 0 for no changes, 2 for an input error.
        for args in (("--word-diff", "old.docx", "old.docx"), ("--format", "markdown", "--word-diff", "old.docx", "old.docx")):
            unchanged = self.pilcrow("review", *args, code=0).stdout
            self.assertEqual(unchanged, self.pilcrow("review", *(a for a in args if a != "--word-diff"), code=0).stdout)
        result = self.pilcrow("review", "--word-diff", "old.docx", "absent.docx", code=2)
        self.assertEqual(result.stdout, b"")

    def test_word_diff_falls_back_to_whole_lines(self) -> None:
        long = " ".join(f"word{n}" for n in range(200))
        cases = {
            "unequal line counts": ("Opening. The cat sat on the mat.", "Opening. The cat sat. It liked the mat."),
            "a dissimilar pair": ("Opening. The cat sat on the mat.", "Opening. Rain fell hard all night."),
            "one dissimilar pair in the block": ("Opening. The cat sat. The mat was red.",
                                                 "Opening. The dog sat. Rain fell hard all night."),
            "a line over the token cap": (f"Opening. {long} old.", f"Opening. {long} new."),
            "a mark sequence": ("Opening. Write [-x-] here.", "Opening. Write [-y-] here."),
        }
        for name, (old, new) in cases.items():
            with self.subTest(name):
                make_docx(self.old, paragraph(old))
                make_docx(self.new, paragraph(new))
                for form in ("text", "markdown"):
                    plain = self.pilcrow("review", "--format", form, "old.docx", "new.docx", code=1).stdout
                    marked = self.pilcrow("review", "--format", form, "--word-diff", "old.docx", "new.docx", code=1).stdout
                    self.assertEqual(marked, plain)
                    self.assertNotIn(b"\n~ ", marked)

    def test_word_diff_still_caps_a_very_large_change(self) -> None:
        make_docx(self.old, "".join(paragraph(f"Old sentence {n}.") for n in range(1500)))
        make_docx(self.new, "".join(paragraph(f"New sentence {n}.") for n in range(1500)))
        result = self.pilcrow("review", "--word-diff", "old.docx", "new.docx", code=1)
        prose = self.section(result.stdout.decode("utf-8"), "Prose")
        # Every sentence pairs, and the blank lines between paragraphs are unchanged
        # pairs shown as context: 3002 diff lines in all, 2000 shown.
        self.assertEqual(len(prose), 2001)
        self.assertEqual(prose[3:6], ["~ [-Old-]{+New+} sentence 0.", " ", "~ [-Old-]{+New+} sentence 1."])
        self.assertEqual(prose[-1], "… 1002 more diff lines omitted")

    def test_help_describes_word_diff(self) -> None:
        help_text = b" ".join(self.pilcrow("review", "--help").stdout.split())
        for words in (b"[--word-diff] OLD NEW", b"[--word-diff] --rev REV", b"[-...-]", b"{+...+}", b"whole"):
            self.assertIn(words, help_text)

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

    def test_markdown_with_revisions(self) -> None:
        document = self.repo / "report.docx"
        make_docx(document, paragraph("The cat sat."))
        self.commit("first")
        make_docx(document, paragraph("The dog sat."))
        self.commit("second")
        before = self.state()
        for args, header in ((("--rev", "HEAD~1", "--rev", "HEAD"), "`HEAD~1:report.docx` → `HEAD:report.docx`"),
                             (("--rev", "HEAD~1"), "`HEAD~1:report.docx` → `report.docx`")):
            text = self.pilcrow("review", *args, "report.docx", code=1)
            result = self.pilcrow("review", "--format", "markdown", *args, "report.docx", code=1)
            self.assertEqual(result.stderr, b"")
            assert_same_report(self, text.stdout.decode("utf-8"), result.stdout.decode("utf-8"))
            self.assertTrue(result.stdout.decode("utf-8").startswith(f"## Review: {header}\n\n"), result.stdout)
        self.assertEqual(self.pilcrow("review", "--format", "markdown", "--rev", "HEAD", "report.docx", code=0).stdout,
                         f"## Review: `HEAD:report.docx` → `report.docx`\n\nProse: no changes\n\n{CAVEAT}\n".encode())
        # Errors print nothing on stdout, whatever the format.
        result = self.pilcrow("review", "--format", "markdown", "--rev", "no-such-rev", "report.docx", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertEqual(self.state(), before)

    def test_word_diff_with_revisions(self) -> None:
        document = self.repo / "report.docx"
        make_docx(document, paragraph("The supplier shall deliver the goods. Payment is due on receipt."))
        self.commit("first")
        make_docx(document, paragraph("The supplier may deliver the goods. Payment is due on receipt."))
        self.commit("second")
        make_docx(document, paragraph("The supplier may deliver all the goods. Payment is due on receipt."))
        before = self.state()
        marked = {
            ("--rev", "HEAD"): "~ The supplier may deliver {+all +}the goods.",
            ("--rev", "HEAD~1"): "~ The supplier [-shall-]{+may+} deliver {+all +}the goods.",
            ("--rev", "HEAD~1", "--rev", "HEAD"): "~ The supplier [-shall-]{+may+} deliver the goods.",
        }
        for args, line in marked.items():
            with self.subTest(args=args):
                text = self.pilcrow("review", "--word-diff", *args, "report.docx", code=1).stdout.decode("utf-8")
                self.assertIn(f"\n@@ -1,2 +1,2 @@\n{line}\n Payment is due on receipt.\n\n", text)
                result = self.pilcrow("review", "--format", "markdown", "--word-diff", *args, "report.docx", code=1)
                self.assertEqual(result.stderr, b"")
                assert_same_report(self, text, result.stdout.decode("utf-8"))
                self.assertIn(f"```diff\n--- OLD\n+++ NEW\n@@ -1,2 +1,2 @@\n{line}\n", result.stdout.decode("utf-8"))
        self.pilcrow("review", "--word-diff", "--rev", "HEAD~1", "--rev", "HEAD~1", "report.docx", code=0)
        self.assertEqual(self.pilcrow("review", "--word-diff", "--rev", "no-such-rev", "report.docx", code=2).stdout, b"")
        self.assertEqual(self.state(), before)

    def test_two_revisions_need_no_folder_in_the_work_tree(self) -> None:
        # A later commit or a sparse checkout may have removed the document's folders.
        nested = Path("old notes") / "drafts" / "report.docx"
        make_docx(self.repo / nested, paragraph("The cat sat."))
        self.commit("first")
        make_docx(self.repo / nested, paragraph("The dog sat."))
        self.commit("second")
        shutil.rmtree(self.repo / "old notes")
        for path in (nested, self.repo / nested):
            with self.subTest(path=str(path)):
                report = self.pilcrow("review", "--rev", "HEAD~1", "--rev", "HEAD", path, code=1).stdout.decode("utf-8")
                self.assertTrue(report.startswith(f"Review: HEAD~1:{path} → HEAD:{path}\n\n"), report)
                self.assertIn("-The cat sat.", report)
                self.assertIn("+The dog sat.", report)
        result = self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", Path("old notes") / "gone.docx", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertIn(f"error: HEAD:{Path('old notes') / 'gone.docx'}: no such file at HEAD".encode(), result.stderr)

    def test_a_folder_now_symlinked_elsewhere_keeps_its_own_history(self) -> None:
        for relative in (Path("docs") / "report.docx", Path("docs") / "sub" / "report.docx"):
            make_docx(self.repo / relative, paragraph("The committed docs version."))
            make_docx(self.repo / "other" / relative.relative_to("docs"), paragraph("The other folder's version."))
        self.commit("first")
        shutil.rmtree(self.repo / "docs")
        try:
            (self.repo / "docs").symlink_to("other", target_is_directory=True)
        except OSError as exc:  # Windows without the symlink privilege
            self.skipTest(f"cannot create a symlink: {exc}")
        for path in (Path("docs") / "report.docx", Path("docs") / "sub" / "report.docx",
                     self.repo / "docs" / "sub" / "report.docx"):
            with self.subTest(path=str(path)):
                # Old: docs/... as committed. New: the working copy, read through the symlink.
                report = self.pilcrow("review", "--rev", "HEAD", path, code=1).stdout.decode("utf-8")
                self.assertIn("-The committed docs version.", report)
                self.assertIn("+The other folder's version.", report)

    def link(self, link: Path, target: Path | str) -> None:
        try:
            link.symlink_to(target, target_is_directory=True)
        except OSError as exc:  # Windows without the symlink privilege
            self.skipTest(f"cannot create a symlink: {exc}")

    def test_a_folder_now_symlinked_to_the_top_keeps_its_own_history(self) -> None:
        for relative in (Path("report.docx"), Path("sub") / "report.docx"):
            make_docx(self.repo / relative, paragraph("The root file's version."))
            make_docx(self.repo / "docs" / relative, paragraph("The committed docs version."))
        self.commit("first")
        for relative in (Path("report.docx"), Path("sub") / "report.docx"):
            make_docx(self.repo / "docs" / relative, paragraph("The second docs version."))
        self.commit("second")
        shutil.rmtree(self.repo / "docs")
        self.link(self.repo / "docs", ".")
        before = self.state()
        for path in (Path("docs") / "report.docx", Path("docs") / "sub" / "report.docx",
                     self.repo / "docs" / "report.docx", self.repo / "docs" / "sub" / "report.docx"):
            with self.subTest(path=str(path)):
                # Old: docs/... as committed, not the root file the link now leads to. New: the working copy.
                report = self.pilcrow("review", "--rev", "HEAD", path, code=1).stdout.decode("utf-8")
                self.assertIn("-The second docs version.", report)
                self.assertIn("+The root file's version.", report)
                # Both sides from Git read docs/..., not the root file.
                self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", path, code=0)
                report = self.pilcrow("review", "--rev", "HEAD~1", "--rev", "HEAD", path, code=1).stdout.decode("utf-8")
                self.assertIn("-The committed docs version.", report)
                self.assertIn("+The second docs version.", report)
                self.assertNotIn("root file", report)
        self.assertEqual(self.state(), before)

    def test_a_folder_now_symlinked_outside_any_repository_keeps_its_own_history(self) -> None:
        for relative in (Path("report.docx"), Path("sub") / "report.docx"):
            make_docx(self.repo / "docs" / relative, paragraph("The committed docs version."))
            make_docx(self.outside / "docs" / relative, paragraph("The outside folder's version."))
        self.commit("first")
        shutil.rmtree(self.repo / "docs")
        self.link(self.repo / "docs", self.outside / "docs")
        before = self.state()
        for path in (Path("docs") / "report.docx", Path("docs") / "sub" / "report.docx",
                     self.repo / "docs" / "report.docx", self.repo / "docs" / "sub" / "report.docx"):
            with self.subTest(path=str(path)):
                report = self.pilcrow("review", "--rev", "HEAD", path, code=1).stdout.decode("utf-8")
                self.assertIn("-The committed docs version.", report)
                self.assertIn("+The outside folder's version.", report)
                self.pilcrow("review", "--rev", "HEAD", "--rev", "HEAD", path, code=0)
        self.assertEqual(self.state(), before)

    def test_symlinks_above_the_repository_are_followed(self) -> None:
        base = Path(self.temp.name)
        make_docx(self.repo / "report.docx", paragraph("The cat sat."))
        make_docx(self.repo / "docs" / "report.docx", paragraph("The committed docs version."))
        self.commit("first")
        make_docx(self.repo / "report.docx", paragraph("The dog sat."))
        shutil.rmtree(self.repo / "docs")
        self.link(self.repo / "docs", ".")
        self.link(base / "alias", base)
        # A checkout reached through a link from inside another repository, such as a home folder kept in Git.
        home = base / "home"
        home.mkdir()
        subprocess.run(["git", "init", "-q", str(home)], env=self.env, capture_output=True, timeout=60, check=True)
        self.link(home / "code", base)
        before = self.state()
        for top in (base / "alias" / "repo", home / "code" / "repo"):
            with self.subTest(top=str(top)):
                report = self.pilcrow("review", "--rev", "HEAD", top / "report.docx", cwd=self.outside,
                                      code=1).stdout.decode("utf-8")
                self.assertIn("-The cat sat.", report)
                self.assertIn("+The dog sat.", report)
                # A link inside the repository is still not followed when its top is reached through one.
                report = self.pilcrow("review", "--rev", "HEAD", top / "docs" / "report.docx", cwd=self.outside,
                                      code=1).stdout.decode("utf-8")
                self.assertIn("-The committed docs version.", report)
                self.assertIn("+The dog sat.", report)
        self.assertEqual(self.state(), before)

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
            (("--rev", "HEAD", "--rev", "HEAD", "."), self.repo, None, b"error: HEAD:.: expected the path of a file"),
            (("--rev", "HEAD", "--rev", "HEAD", self.repo.anchor), self.repo, None,
             f"error: HEAD:{self.repo.anchor}: expected the path of a file".encode()),
            (("--rev", "HEAD", "."), self.repo, None, b"error: HEAD:.: expected the path of a file"),
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


class MarkdownEscapingTests(unittest.TestCase):
    """The helpers that keep document text literal in a Markdown report."""

    def test_escape_renders_literally_on_one_line(self) -> None:
        from pilcrow.review import escape
        self.assertEqual(escape("a\nb\r\nc\rd"), "a b c d")
        self.assertEqual(escape("- item"), "\\- item")
        self.assertEqual(escape("+ item"), "\\+ item")
        self.assertEqual(escape("12. item"), "12\\. item")
        self.assertEqual(escape("3) item"), "3\\) item")
        self.assertEqual(escape("=== x"), "\\=== x")
        self.assertEqual(escape("> quote"), "&gt; quote")
        self.assertEqual(escape("<b>&copy;"), "&lt;b&gt;&amp;copy;")
        self.assertEqual(escape("$x$ ~~s~~ _u_"), "\\$x\\$ \\~\\~s\\~\\~ \\_u\\_")
        self.assertEqual(escape("see www.example.invalid or a@b.invalid"),
                         "see www\\.example.invalid or a\\@b.invalid")
        self.assertEqual(escape("back\\slash"), "back\\\\slash")
        self.assertEqual(escape("plain words, 1.5 (ok)"), "plain words, 1.5 (ok)")

    def test_code_spans_outrun_backticks_inside(self) -> None:
        from pilcrow.review import code
        self.assertEqual(code("plain.docx"), "`plain.docx`")
        self.assertEqual(code("a`b"), "``a`b``")
        self.assertEqual(code("`a``"), "``` `a`` ```")
        self.assertEqual(code(" padded "), "`  padded  `")
        self.assertEqual(code("two\nlines"), "`two lines`")


def unmark(line: str, side: str) -> str:
    """The old (``side`` "old") or new line a marked line was made from."""
    if side == "old":
        return re.sub(r"\[-(.*?)-\]", r"\1", re.sub(r"\{\+.*?\+\}", "", line, flags=re.S), flags=re.S)
    return re.sub(r"\{\+(.*?)\+\}", r"\1", re.sub(r"\[-.*?-\]", "", line, flags=re.S), flags=re.S)


class WordMarksTests(unittest.TestCase):
    """The word marks of `pilcrow review --word-diff`, and when lines are not paired."""

    def assert_round_trip(self, old: str, new: str) -> str:
        from pilcrow.review import word_marks
        marked = word_marks(old, new)
        self.assertEqual(unmark(marked, "old"), old, marked)
        self.assertEqual(unmark(marked, "new"), new, marked)
        return marked

    def test_marks(self) -> None:
        cases = {
            ("The supplier shall deliver the goods within 30 days.", "The supplier may deliver the goods within 60 days."):
                "The supplier [-shall-]{+may+} deliver the goods within [-30-]{+60+} days.",
            # Changes separated only by whitespace are one mark.
            ("The supplier shall deliver the goods.", "The supplier may send the goods."):
                "The supplier [-shall deliver-]{+may send+} the goods.",
            ("Pay 1,000.50 EUR, net.", "Pay 1,200.50 USD; net!"): "Pay 1,[-000-]{+200+}.50 [-EUR,-]{+USD;+} net[-.-]{+!+}",
            ("the the cat sat", "the cat sat"): "the [-the -]cat sat",
            # An added or removed word keeps its following space inside the mark, wherever it is.
            ("deliver the goods", "deliver all the goods"): "deliver {+all +}the goods",
            ("a b c d", "a c d"): "a [-b -]c d",
            ("Die Größe ist groß.", "Die Größe ist klein."): "Die Größe ist [-groß-]{+klein+}.",
            ("a  b", "a b"): "a[-  -]{+ +}b",
            ("end", "end now"): "end{+ now+}",
            ("", "added"): "{+added+}",
        }
        for (old, new), expected in cases.items():
            with self.subTest(old=old, new=new):
                self.assertEqual(self.assert_round_trip(old, new), expected)

    def test_round_trip(self) -> None:
        pairs = [
            ("He said: “No.” (See §3.)", "She said: “Yes!” (See §4.)"),
            ("10 apples, 20 pears, 30 plums.", "10 apples, 25 pears, 30 plums and figs."),
            ("very very very good", "very good very very"),
            ("Ünïcödé café naïve 東京 test", "Unicode café naive 京都 test"),
            ("tabs\tand  spaces ", " tabs and\tspaces"),
            ("Brackets [ ] { } + - stay.", "Brackets [ ] { } - + stay."),
            ("x-", "x--"), ("{x", "{{x"), ("[", "[["), ("-", "+"), ("", ""), ("a", ""),
        ]
        for old, new in pairs:
            with self.subTest(old=old, new=new):
                self.assert_round_trip(old, new)
        pieces = ["the", "cat", "Größe", "42", "3.5", " ", "  ", "\t", ",", ".", "-", "+", "[", "]", "{", "}",
                  "“", "é", "_", "x"]
        rng = random.Random(43)
        for _ in range(2000):
            old = "".join(rng.choice(pieces) for _ in range(rng.randrange(12)))
            new = list(old)
            for _ in range(rng.randrange(4)):
                at = rng.randrange(len(new) + 1)
                new[at:at + rng.randrange(3)] = rng.choice(pieces)
            new_line = "".join(new)
            if any(mark in line for line in (old, new_line) for mark in ("[-", "-]", "{+", "+}")):
                continue  # never paired, so never marked
            self.assert_round_trip(old, new_line)

    def test_pairing_rules(self) -> None:
        from pilcrow.review import (
            MAX_WORD_TOKENS,
            MIN_WORD_SIMILARITY,
            pairable,
            paired,
        )
        self.assertEqual(MIN_WORD_SIMILARITY, 0.5)
        self.assertTrue(pairable("The cat sat.", "The dog sat."))  # 3 of 4 words: 0.75
        self.assertTrue(pairable("a b c d", "a b x y"))  # exactly 0.5
        self.assertFalse(pairable("a b c d", "a x y z"))  # 0.25
        words = " ".join(["w"] * MAX_WORD_TOKENS)
        self.assertTrue(pairable(words, words))
        self.assertFalse(pairable(words + " w", words + " v"))  # one line over the cap
        self.assertFalse(pairable(words + " w", words))
        for mark in ("[-", "-]", "{+", "+}"):
            self.assertFalse(pairable(f"The {mark} cat.", "The dog."), mark)
            self.assertFalse(pairable("The dog.", f"The {mark} cat."), mark)
        self.assertEqual(paired(["The cat sat.", "", "It was warm."], ["The dog sat.", "", "It was hot."]),
                         ["~ The [-cat-]{+dog+} sat.", " ", "~ It was [-warm-]{+hot+}."])
        self.assertIsNone(paired(["The cat sat."], ["The dog sat.", "More."]))
        self.assertIsNone(paired(["The cat sat.", "It was warm."], ["The dog sat.", "Rain fell all night."]))

    def test_whole_line_fallback_is_the_plain_diff(self) -> None:
        from pilcrow.review import prose_diff
        long = " ".join(f"w{n}" for n in range(201))
        for old, new in (("A.\nThe cat sat.\nB.", "A.\nThe cat sat.\nIt purred.\nB."),
                         ("A.\nThe cat sat.\nB.", "A.\nRain fell all night.\nB."),
                         (f"A.\n{long} x\nB.", f"A.\n{long} y\nB."),
                         ("A.\nWrite {+x+}.\nB.", "A.\nWrite {+y+}.\nB."),
                         ("A.\nGone.\nB.", "A.\nB.")):
            with self.subTest(old=old[:40], new=new[:40]):
                self.assertEqual(prose_diff(old, new, marks=True), prose_diff(old, new))
        self.assertEqual(prose_diff("same", "same", marks=True), [])

    def test_hunks_match_unified_diff(self) -> None:
        """With pairing switched off, the word diff is exactly difflib's unified diff, hunk headers included."""
        from unittest.mock import patch

        from pilcrow import review
        rng = random.Random(7)
        vocabulary = ["", "A.", "B.", "C.", "D.", "E.", "F."]
        with patch.object(review, "paired", return_value=None):
            for _ in range(500):
                old = "\n".join(rng.choice(vocabulary) for _ in range(rng.randrange(30)))
                new = "\n".join(rng.choice(vocabulary) for _ in range(rng.randrange(30)))
                with self.subTest(old=old, new=new):
                    self.assertEqual(review.prose_diff(old, new, marks=True), review.prose_diff(old, new))
