"""`python -m pilcrow.summary`: the GitHub Action's opt-in job summary of changed documents.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from pilcrow import summary

try:
    from .test_cli import (
        COMMENTS_EXT_REL,
        COMMENTS_REL,
        NUMBERING_REL,
        PROJECT,
        R,
        abstract_num,
        comments_extended,
        comments_part,
        end,
        level,
        list_item,
        make_docx,
        num,
        numbering,
        paragraph,
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
        comments_extended,
        comments_part,
        end,
        level,
        list_item,
        make_docx,
        num,
        numbering,
        paragraph,
        run,
        start,
    )
    from test_textconv import tree  # type: ignore[no-redef]


def rich(path: Path, word: str) -> None:
    """A DOCX whose prose, comment, tracked change, link, field and list all carry ``word``."""
    body = (paragraph(f"Prose{word} sentence.")
            + f"<w:p>{start('1')}{run(f'Anchor{word} passage.')}{end('1')}</w:p>"
            + f'<w:p><w:ins w:id="7" w:author="Author{word}"><w:r><w:t>Inserted{word}.</w:t></w:r></w:ins></w:p>'
            + f'<w:p><w:hyperlink r:id="link">{run(f"Label{word}")}</w:hyperlink></w:p>'
            + f'<w:p><w:fldSimple w:instr=" Field{word} "><w:r><w:t>Shown{word}</w:t></w:r></w:fldSimple></w:p>'
            + list_item(f"Item{word}", "1"))
    make_docx(path, body,
              extra={"word/comments.xml": comments_part({"1": f"Comment{word} text."}),
                     "word/numbering.xml": numbering(abstract_num("0", level(0, "decimal", f"{word}%1.")),
                                                     num("1", "0"))},
              relationships=COMMENTS_REL + NUMBERING_REL
              + f'<Relationship Id="link" Type="{R}/hyperlink" Target="https://{word.lower()}.example/"'
                ' TargetMode="External"/>')


class SummaryTests(unittest.TestCase):
    def setUp(self) -> None:
        if shutil.which("git") is None:
            self.skipTest("git is not installed")
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-summary-")
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.repo = base / "repo"
        self.repo.mkdir()
        self.output = base / "out" / "summary.md"
        self.output.parent.mkdir()
        self.env = {**os.environ, "GIT_CEILING_DIRECTORIES": str(base)}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))
        patcher = patch.dict(os.environ, {"GIT_CEILING_DIRECTORIES": str(base)})
        patcher.start()
        self.addCleanup(patcher.stop)
        # state() lists .git too, so stop the background maintenance and auto-gc that `git commit` can start
        # from creating or deleting files there while it is read (see work_tree() in tests/test_quickstart.py).
        for args in (("init", "-q"), ("config", "user.name", "Pilcrow Test"),
                     ("config", "user.email", "pilcrow@example.invalid"), ("config", "commit.gpgsign", "false"),
                     ("config", "maintenance.auto", "false"), ("config", "gc.auto", "0")):
            self.git(*args)

    def git(self, *args: str, cwd: Path | None = None) -> str:
        result = subprocess.run(["git", *args], cwd=cwd or self.repo, env=self.env, capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.decode("utf-8")

    def commit(self, message: str) -> str:
        self.git("add", "-A")
        self.git("commit", "-q", "--allow-empty", "-m", message)
        return self.git("rev-parse", "HEAD").strip()

    def summarize(self, base: str, mode: str = "outline", root: Path | None = None) -> tuple[str, str]:
        """Run the module in-process; return the summary and what it printed (its annotations)."""
        printed = io.StringIO()
        with redirect_stdout(printed):
            code = summary.main([f"--base={base}", f"--root={root or self.repo}", f"--mode={mode}",
                                 f"--output={self.output}"])
        self.assertEqual(code, 0, printed.getvalue())
        return self.output.read_text(encoding="utf-8"), printed.getvalue()

    def state(self) -> tuple[object, ...]:
        status = subprocess.run(["git", "--no-optional-locks", "status", "--porcelain"], cwd=self.repo,
                                env=self.env, capture_output=True, timeout=60).stdout
        return tree(self.repo), status

    def test_full_mode_prints_the_review_markdown_of_each_document(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("The cat sat."))
        make_docx(self.repo / "sub dir" / "b.docx", paragraph("First version."))
        base = self.commit("first")
        make_docx(self.repo / "a.docx", paragraph("The dog sat."))
        make_docx(self.repo / "sub dir" / "b.docx", paragraph("Second version."))
        head = self.commit("second")
        before = self.state()
        text, printed = self.summarize(base, "full")
        self.assertEqual(self.state(), before)
        self.assertEqual(printed, "")
        self.assertIn("**This summary includes document text.**", text)
        self.assertIn(f"from `{base[:7]}` (base) to `{head[:7]}` (head).", text)
        reports = re.split(r"^(?=## Review: )", text, flags=re.MULTILINE)[1:]
        self.assertEqual(len(reports), 2)
        for report, path in zip(reports, ("a.docx", "sub dir/b.docx"), strict=True):
            cli = subprocess.run([sys.executable, "-m", "pilcrow", "review", "--format", "markdown",
                                  "--rev", base, "--rev", "HEAD", path], cwd=self.repo, env=self.env,
                                 capture_output=True, timeout=60)
            self.assertEqual(cli.returncode, 1, cli.stderr)
            heading, body = report.split("\n", 1)
            self.assertEqual(heading, f"## Review: `{base[:7]}:{path}` → `{head[:7]}:{path}`")
            # The same report apart from the labels in the heading.
            self.assertEqual(body.rstrip("\n"), cli.stdout.decode("utf-8").split("\n", 1)[1].rstrip("\n"))
        self.assertIn("-The cat sat.", text)
        self.assertIn("### Prose", text)

    def test_outline_publishes_no_document_text(self) -> None:
        rich(self.repo / "report.docx", "Zebraold")
        base = self.commit("first")
        rich(self.repo / "report.docx", "Zebranew")
        self.commit("second")
        text, _ = self.summarize(base, "outline")
        self.assertNotIn("zebra", text.lower())
        self.assertNotIn("Ada", text)  # comment author
        self.assertIn("- `report.docx` — changed: Prose changed, Comments (", text)
        for title in ("Tracked changes (", "Hyperlinks (", "Fields (", "Lists ("):
            self.assertIn(title, text)
        self.assertNotIn("document text", text)  # no visibility note: outline publishes none
        # The same documents in full mode do publish their text.
        full, _ = self.summarize(base, "full")
        for word in ("ProseZebraold", "CommentZebranew", "InsertedZebranew", "zebranew.example"):
            self.assertIn(word, full)

    def test_outline_counts_one_comment_change_for_an_edited_parent(self) -> None:
        def thread(parent: str) -> None:
            body = f"<w:p>{start('1')}{start('2')}{run('The cat sat.')}{end('1')}{end('2')}</w:p>"
            make_docx(self.repo / "thread.docx", body,
                      extra={"word/comments.xml": comments_part({"1": parent, "2": "I agree."},
                                                                {"1": "11111111", "2": "22222222"}),
                             "word/commentsExtended.xml": comments_extended(("11111111", None, "0"),
                                                                            ("22222222", "11111111", "0"))},
                      relationships=COMMENTS_REL + COMMENTS_EXT_REL)
        thread("Please clarify.")
        base = self.commit("first")
        thread("Please clarify this.")
        self.commit("second")
        text, _ = self.summarize(base, "outline")
        self.assertIn("- `thread.docx` — changed: Comments (1)\n", text)

    def test_statuses_and_eligibility(self) -> None:
        make_docx(self.repo / "docs" / "kept.docx", paragraph("Same."))
        make_docx(self.repo / "docs" / "gone.docx", paragraph("Going."))
        make_docx(self.repo / "docs" / "same.docx", paragraph("Unchanged text."))
        make_docx(self.repo / "outside.docx", paragraph("Outside the root."))
        base = self.commit("first")
        make_docx(self.repo / "docs" / "kept.docx", paragraph("Changed."))
        (self.repo / "docs" / "gone.docx").unlink()
        make_docx(self.repo / "docs" / "new.docx", paragraph("New."))
        make_docx(self.repo / "docs" / "same.docx", paragraph("Unchanged text."), churn=True)  # bytes differ only
        make_docx(self.repo / "docs" / "~$kept.docx", paragraph("Lock file."))
        make_docx(self.repo / "docs" / ".pilcrow" / "x.docx", paragraph("Output folder."))
        (self.repo / "docs" / "notes.txt").write_text("not a document", encoding="utf-8")
        make_docx(self.repo / "outside.docx", paragraph("Changed outside the root."))
        self.commit("second")
        for mode in ("outline", "full"):
            text, printed = self.summarize(base, mode, root=self.repo / "docs")
            self.assertEqual(printed, "")
            self.assertIn(" in `", text)
            self.assertIn("- `gone.docx` — removed", text)
            self.assertIn("- `new.docx` — added", text)
            for absent in ("~$", ".pilcrow", "notes.txt", "outside"):
                self.assertNotIn(absent, text)
            if mode == "outline":
                self.assertEqual(text.split("\n\n", 2)[2], "\n".join([
                    "- `gone.docx` — removed", "- `kept.docx` — changed: Prose changed",
                    "- `new.docx` — added", "- `same.docx` — no supported changes"]) + "\n")
            else:
                self.assertIn("## Review: ", text)
                self.assertIn("+Changed.", text)
                self.assertIn(":same.docx`\n\nProse: no changes\n\nNo supported prose or review changes detected", text)

    def test_no_changed_documents(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("Same."))
        base = self.commit("first")
        (self.repo / "readme.txt").write_text("x", encoding="utf-8")
        self.commit("second")
        text, printed = self.summarize(base)
        self.assertTrue(text.endswith("\n\nNo DOCX documents changed.\n"), text)
        self.assertEqual(printed, "")

    def test_unusable_bases_explain_and_warn(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        self.commit("first")
        for base, reason in (("", "there is no base revision"), ("0" * 40, "there is no base revision"),
                             ("1234567890abcdef1234567890abcdef12345678", "is not in this checkout"),
                             ("-x", "is not in this checkout")):
            with self.subTest(base=base):
                text, printed = self.summarize(base, "full")
                self.assertIn("No review: ", text)
                self.assertIn(reason, text)
                if "checkout" in reason:
                    self.assertIn("fetch-depth: 0", text)
                self.assertTrue(printed.startswith("::warning::Pilcrow summary: no review: "), printed)
                self.assertEqual(printed.count("\n"), 1)
                self.assertNotIn("## Review", text)

    def test_a_shallow_checkout_suggests_fetching_more_history(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        base = self.commit("first")
        make_docx(self.repo / "a.docx", paragraph("Two."))
        self.commit("second")
        clone = Path(self.temp.name) / "clone"
        self.git("clone", "-q", "--depth", "1", self.repo.as_uri(), str(clone))
        text, printed = self.summarize(base, "outline", root=clone)
        self.assertIn("The checkout is shallow.", text)
        self.assertIn("`fetch-depth: 0`", text)
        self.assertIn("::warning::", printed)

    def fork(self) -> tuple[str, str, str]:
        """Main and branch ``pr`` fork at a commit; ``pr`` edits b.docx, then main edits a.docx and adds c.docx.

        Returns (fork, main's tip, pr's tip), with ``pr`` checked out."""
        make_docx(self.repo / "a.docx", paragraph("Alpha one."))
        make_docx(self.repo / "b.docx", paragraph("Bravo one."))
        fork = self.commit("fork")
        main = self.git("symbolic-ref", "--short", "HEAD").strip()
        self.git("checkout", "-q", "-b", "pr")
        make_docx(self.repo / "b.docx", paragraph("Bravo two."))
        pr = self.commit("pr edits b")
        self.git("checkout", "-q", main)
        make_docx(self.repo / "a.docx", paragraph("Alpha two."))
        make_docx(self.repo / "c.docx", paragraph("Charlie."))
        tip = self.commit("main edits a and adds c")
        self.git("checkout", "-q", "pr")
        return fork, tip, pr

    def test_a_pull_request_head_is_compared_from_the_merge_base(self) -> None:
        fork, tip, pr = self.fork()
        before = self.state()
        for mode in ("outline", "full"):
            with self.subTest(mode=mode):
                text, printed = self.summarize(tip, mode)
                self.assertEqual(self.state(), before)
                self.assertEqual(printed, "")
                self.assertIn(f"from `{fork[:7]}` (merge base of `{tip[:7]}` and `HEAD`) to `{pr[:7]}` (head).", text)
                for absent in ("a.docx", "c.docx", "Alpha", "Charlie", "removed"):
                    self.assertNotIn(absent, text)
                if mode == "outline":
                    self.assertTrue(text.endswith("\n\n- `b.docx` — changed: Prose changed\n"), text)
                else:
                    self.assertEqual(text.count("## Review: "), 1)
                    self.assertIn(f"## Review: `{fork[:7]}:b.docx` → `{pr[:7]}:b.docx`", text)
                    self.assertIn("-Bravo one.", text)
                    self.assertIn("+Bravo two.", text)
        # Checked out on main, the pull request's fork point is an ancestor: the usual header, main's own edits.
        self.git("checkout", "-q", "-")
        text, _ = self.summarize(fork, "outline")
        self.assertIn(f"from `{fork[:7]}` (base) to `{tip[:7]}` (head).", text)
        self.assertNotIn("merge base", text)
        self.assertIn("- `a.docx` — changed: Prose changed\n- `c.docx` — added\n", text)

    def test_a_force_pushed_before_is_compared_from_the_merge_base(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        make_docx(self.repo / "b.docx", paragraph("Kept."))
        parent = self.commit("first")
        make_docx(self.repo / "a.docx", paragraph("Two."))
        replaced = self.commit("pushed, then force-pushed away")
        self.git("reset", "-q", "--hard", parent)
        make_docx(self.repo / "b.docx", paragraph("Rewritten."))
        head = self.commit("the new push")
        before = self.state()
        for mode in ("outline", "full"):
            with self.subTest(mode=mode):
                text, printed = self.summarize(replaced, mode)
                self.assertEqual(self.state(), before)
                self.assertEqual(printed, "")
                self.assertIn(f"from `{parent[:7]}` (merge base of `{replaced[:7]}` and `HEAD`) to `{head[:7]}` (head).",
                              text)
                self.assertNotIn("a.docx", text)
                self.assertIn("b.docx", text)
                if mode == "full":
                    self.assertIn(f"## Review: `{parent[:7]}:b.docx` → `{head[:7]}:b.docx`", text)
                    self.assertIn("-Kept.", text)

    def test_unrelated_histories_explain_and_warn(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        other = self.commit("first")
        self.git("checkout", "-q", "--orphan", "unrelated")
        make_docx(self.repo / "a.docx", paragraph("Two."))
        self.commit("a separate history")
        before = self.state()
        for mode in ("outline", "full"):
            with self.subTest(mode=mode):
                text, printed = self.summarize(other, mode)
                self.assertEqual(self.state(), before)
                self.assertTrue(text.startswith(f"{summary.TITLE}\n\nNo review: "), text)
                self.assertIn("have no common ancestor", text)
                self.assertIn("unrelated histories", text)
                self.assertNotIn("fetch-depth", text)
                self.assertNotIn("## Review", text)
                self.assertTrue(printed.startswith("::warning::Pilcrow summary: no review: "), printed)
                self.assertEqual(printed.count("\n"), 1)

    def test_a_shallow_checkout_that_misses_the_merge_base_suggests_fetching_more(self) -> None:
        _, tip, _ = self.fork()
        clone = Path(self.temp.name) / "clone"
        self.git("clone", "-q", "--depth", "1", "--no-single-branch", "--branch", "pr", self.repo.as_uri(), str(clone))
        before = tree(clone)
        text, printed = self.summarize(tip, "full", root=clone)
        self.assertEqual(tree(clone), before)
        self.assertIn("No review: ", text)
        self.assertIn("have no common ancestor", text)
        self.assertIn("The checkout is shallow", text)
        self.assertIn("`fetch-depth: 0`", text)
        self.assertNotIn("## Review", text)
        self.assertIn("::warning::", printed)

    def test_outside_a_repository(self) -> None:
        outside = Path(self.temp.name) / "plain"
        outside.mkdir()
        text, printed = self.summarize("HEAD~1", root=outside)
        self.assertIn("No review: ", text)
        self.assertIn("::warning::", printed)

    def test_one_unreadable_document_does_not_stop_the_others(self) -> None:
        (self.repo / "bad.docx").write_bytes(b"not a zip")
        make_docx(self.repo / "good.docx", paragraph("Before."))
        base = self.commit("first")
        (self.repo / "bad.docx").write_bytes(b"still not a zip")
        make_docx(self.repo / "good.docx", paragraph("After."))
        self.commit("second")
        for mode in ("outline", "full"):
            text, printed = self.summarize(base, mode)
            if mode == "outline":  # the reason can quote package contents: only the warning has it
                self.assertIn("- `bad.docx` — could not be read\n", text)
            else:
                self.assertIn("- `bad.docx` — could not be read: ", text)
            self.assertTrue(printed.startswith("::warning::Pilcrow could not read bad.docx: "), printed)
            self.assertIn("good.docx", text)
            self.assertIn("Prose changed" if mode == "outline" else "+After.", text)

    def test_an_unexpected_error_is_reported_like_a_read_error(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        make_docx(self.repo / "b.docx", paragraph("One."))
        base = self.commit("first")
        make_docx(self.repo / "a.docx", paragraph("Two."))
        make_docx(self.repo / "b.docx", paragraph("Two."))
        self.commit("second")
        real = summary.build

        def flaky(old: object, new: object) -> object:
            if flaky.calls == 0:  # type: ignore[attr-defined]
                flaky.calls = 1  # type: ignore[attr-defined]
                raise KeyError("boom")
            return real(old, new)  # type: ignore[arg-type]
        flaky.calls = 0  # type: ignore[attr-defined]
        with patch.object(summary, "build", flaky):
            text, printed = self.summarize(base, "outline")
        self.assertIn("- `a.docx` — could not be read\n", text)
        self.assertIn("::warning::Pilcrow could not read a.docx: 'boom'", printed)
        self.assertIn("- `b.docx` — changed: Prose changed", text)

    @unittest.skipUnless(sys.platform.startswith("linux"), "needs a file system that accepts non-UTF-8 names")
    def test_a_non_utf8_file_name(self) -> None:
        name = os.fsdecode(b"bad\xff name.docx")
        make_docx(self.repo / name, paragraph("One."))
        base = self.commit("first")
        make_docx(self.repo / name, paragraph("Two."))
        self.commit("second")
        for mode in ("outline", "full"):
            text, _ = self.summarize(base, mode)
            self.assertIn("bad\\xff name.docx", text)
            self.assertIn("Prose changed" if mode == "outline" else "+Two.", text)

    def test_a_non_ascii_name_with_a_narrow_console_encoding(self) -> None:
        name = "café 😀.docx"
        (self.repo / name).write_bytes(b"not a zip")
        base = self.commit("first")
        (self.repo / name).write_bytes(b"still not a zip")
        self.commit("second")
        env = {**self.env, "PYTHONIOENCODING": "cp1252"}
        result = subprocess.run([sys.executable, "-m", "pilcrow.summary", f"--base={base}", "--root=.", "--mode=full",
                                 f"--output={self.output}"], cwd=self.repo, env=env, capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(b"::warning::Pilcrow could not read caf\xe9 \\U0001f600.docx: ", result.stdout)
        self.assertIn(f"- `{name}` — could not be read: ", self.output.read_text(encoding="utf-8"))

    def test_a_size_limit_omits_later_reports(self) -> None:
        for name in "abcd":
            make_docx(self.repo / f"{name}.docx", paragraph(f"Old {name}."))
        base = self.commit("first")
        for name in "abcd":
            make_docx(self.repo / f"{name}.docx", paragraph(f"New {name}. " + "Long text. " * 50))
        self.commit("second")
        full, _ = self.summarize(base, "full")
        one = len(re.split(r"^(?=## Review: )", full, flags=re.MULTILINE)[1].encode("utf-8"))
        header = len(full.encode("utf-8").split(b"## Review: ", 1)[0])
        with patch.object(summary, "BUDGET", header + 2 * one + 10):
            text, printed = self.summarize(base, "full")
        self.assertLessEqual(len(text.encode("utf-8")), header + 2 * one + 10 + 200)
        self.assertEqual(text.count("## Review: "), 2)
        self.assertIn("- `c.docx` — report omitted: summary size limit", text)
        self.assertIn("- `d.docx` — report omitted: summary size limit", text)
        self.assertIn("::warning::Pilcrow summary: some reports were omitted", printed)
        # Past the hard limit even the omitted documents stop being listed one by one.
        with patch.object(summary, "BUDGET", header + one + 10), patch.object(summary, "LIMIT", header + one + 60):
            text, _ = self.summarize(base, "full")
        self.assertEqual(text.count("## Review: "), 1)
        self.assertIn("more changed documents not listed: summary size limit", text)
        self.assertNotIn("d.docx", text)

    def test_paths_are_code_spans(self) -> None:
        make_docx(self.repo / "a`b _c_.docx", paragraph("One."))
        base = self.commit("first")
        make_docx(self.repo / "a`b _c_.docx", paragraph("Two."))
        self.commit("second")
        text, _ = self.summarize(base)
        self.assertIn("- ``a`b _c_.docx`` — changed", text)

    def test_runs_as_a_module(self) -> None:
        make_docx(self.repo / "a.docx", paragraph("One."))
        base = self.commit("first")
        make_docx(self.repo / "a.docx", paragraph("Two."))
        self.commit("second")
        command = [sys.executable, "-m", "pilcrow.summary", "--base", base, "--root", ".", "--output", str(self.output)]
        result = subprocess.run([*command, "--mode", "outline"], cwd=self.repo, env=self.env,
                                capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, b"")
        text = self.output.read_text(encoding="utf-8")
        self.assertIn("Changed DOCX documents from ", text)
        self.assertNotIn(" in `.`", text)
        self.assertIn("- `a.docx` — changed: Prose changed", text)
        result = subprocess.run([*command, "--mode", "Full"], cwd=self.repo, env=self.env,
                                capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 2)


class ActionManifestTests(unittest.TestCase):
    def test_summary_is_opt_in_and_runs_after_a_failed_check(self) -> None:
        text = (PROJECT / "action.yml").read_text(encoding="utf-8")
        self.assertRegex(text, r"\n  summary:\n(    .*\n)*?    default: \"none\"\n")
        self.assertIn("default: ${{ github.event.pull_request.base.sha || github.event.before }}", text)
        steps = text.split("\n    - ")
        (step,) = [s for s in steps if "pilcrow.summary" in s]
        self.assertIn("if: always() && steps.install.outcome == 'success' && inputs.summary != 'none'"
                      " && inputs.summary != ''", step)
        self.assertLess(text.index("name: Check companions"), text.index("pilcrow.summary"))
        self.assertIn("value: ${{ steps.summary.outputs.file }}", text)


if __name__ == "__main__":
    unittest.main()
