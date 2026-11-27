"""`pilcrow review --rev REV [--rev REV2] --all`: every DOCX a branch or revision range changed.

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
    from .test_cli import PROJECT, make_docx, paragraph
    from .test_textconv import tree
except ImportError:  # discovered with -s tests: the directory itself is on sys.path
    from test_cli import PROJECT, make_docx, paragraph  # type: ignore[no-redef]
    from test_textconv import tree  # type: ignore[no-redef]

NONE = "No DOCX documents changed.\n"


class ReviewAllTests(unittest.TestCase):
    def setUp(self) -> None:
        if shutil.which("git") is None:
            self.skipTest("git is not installed")
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-review-all-")
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.repo = base / "repo"
        self.repo.mkdir()
        self.outside = base / "outside"  # not in any repository: GIT_CEILING_DIRECTORIES stops the search
        self.outside.mkdir()
        self.tmp = base / "tmp"  # an empty temporary directory for every process, to show nothing is written there
        self.tmp.mkdir()
        self.env = {**os.environ, "TMPDIR": str(self.tmp), "TEMP": str(self.tmp), "TMP": str(self.tmp),
                    "GIT_CEILING_DIRECTORIES": str(base)}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))
        # state() lists .git too, so stop the background maintenance and auto-gc that `git commit` can start.
        for args in (("init", "-q", "-b", "main"), ("config", "user.name", "Pilcrow Test"),
                     ("config", "user.email", "pilcrow@example.invalid"), ("config", "commit.gpgsign", "false"),
                     ("config", "maintenance.auto", "false"), ("config", "gc.auto", "0")):
            self.git(*args)

    def git(self, *args: str, cwd: Path | None = None) -> str:
        result = subprocess.run(["git", *args], cwd=cwd or self.repo, env=self.env, capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.decode("utf-8")

    def commit(self, message: str) -> None:
        self.git("add", "-A")
        self.git("commit", "-q", "--allow-empty", "-m", message)

    def doc(self, path: str, *sentences: str) -> None:
        target = self.repo / path
        target.parent.mkdir(parents=True, exist_ok=True)
        make_docx(target, "".join(paragraph(s) for s in sentences))

    def pilcrow(self, *args: str | Path, code: int, cwd: Path | None = None) -> subprocess.CompletedProcess[bytes]:
        result = subprocess.run([sys.executable, "-m", "pilcrow", *(str(arg) for arg in args)], cwd=cwd or self.repo,
                                env=self.env, capture_output=True, timeout=120)
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout!r}\nstderr={result.stderr!r}")
        self.assertNotIn(b"Traceback", result.stderr)
        return result

    def out(self, *args: str | Path, code: int, cwd: Path | None = None) -> str:
        result = self.pilcrow("review", *args, code=code, cwd=cwd)
        return result.stdout.decode("utf-8")

    def state(self) -> tuple[object, ...]:
        """The work tree (with .git), the index and the temporary directory, read without refreshing the index."""
        index = self.repo / ".git" / "index"
        status = subprocess.run(["git", "--no-optional-locks", "status", "--porcelain"], cwd=self.repo,
                                env=self.env, capture_output=True, timeout=60).stdout
        return tree(self.repo), index.stat().st_mtime_ns, status, tree(self.tmp)

    def branch(self) -> None:
        """main has a, b and d under docs/; feature edits a, adds c, removes d and leaves b alone."""
        self.doc("docs/a.docx", "The supplier shall deliver the goods within 30 days.", "Payment is due on receipt.")
        self.doc("docs/b.docx", "Unchanged.")
        self.doc("docs/d.docx", "To be removed.")
        self.commit("base")
        self.git("checkout", "-q", "-b", "feature")
        self.doc("docs/a.docx", "The supplier may deliver the goods within 60 days.", "Payment is due on receipt.")
        self.doc("docs/c.docx", "A new document.")
        (self.repo / "docs" / "d.docx").unlink()
        self.commit("feature")

    def body(self, report: str) -> str:
        """A report without its header line, which names the sides as each form spells them."""
        return report.split("\n", 1)[1]

    def test_a_branch_reports_changed_added_and_removed_documents_in_path_order(self) -> None:
        self.branch()
        before = self.state()
        printed = self.out("--rev", "main", "--all", code=1)
        self.assertEqual(self.state(), before)
        self.assertFalse((self.repo / ".pilcrow").exists())
        self.assertEqual(printed.count("Review: "), 1)
        self.assertNotIn("b.docx", printed)
        # a's report is the one the single-PATH form prints, after its header; c and d are one line each.
        single = self.out("--rev", "main", Path("docs", "a.docx"), code=1)
        self.assertEqual(printed, "Review: main:docs/a.docx → docs/a.docx\n" + self.body(single)
                         + "\nAdded: docs/c.docx\n\nRemoved: main:docs/d.docx\n")
        self.assertNotIn("/dev/null", printed)
        # Two revisions name both sides by revision; paths stay relative to the repository from a subfolder.
        printed = self.out("--rev", "main", "--rev", "feature", "--all", code=1, cwd=self.repo / "docs")
        self.assertTrue(printed.startswith("Review: main:docs/a.docx → feature:docs/a.docx\n\n"), printed)
        self.assertTrue(printed.endswith("\nAdded: feature:docs/c.docx\n\nRemoved: main:docs/d.docx\n"), printed)
        self.assertEqual(self.state(), before)

    def test_root_limits_the_scope(self) -> None:
        self.branch()
        self.doc("other.docx", "Outside the scope.")
        self.commit("other")
        self.assertIn("other.docx", self.out("--rev", "main", "--all", code=1))
        printed = self.out("--rev", "main", "--all", "--root", "docs", code=1)
        self.assertNotIn("other.docx", printed)
        self.assertIn("Review: main:docs/a.docx → docs/a.docx", printed)
        self.assertEqual(self.out("--rev", "feature", "--all", "--root", self.repo / "docs", code=0), NONE)
        result = self.pilcrow("review", "--rev", "main", "--all", "--root", "missing", code=2)
        self.assertIn(b"--root is not an existing directory", result.stderr)

    def test_a_rename_is_removed_and_added(self) -> None:
        self.doc("old.docx", "Same text.")
        self.commit("first")
        self.git("mv", "old.docx", "new.docx")
        self.commit("rename")
        self.assertEqual(self.out("--rev", "HEAD~1", "--rev", "HEAD", "--all", code=1),
                         "Added: HEAD:new.docx\n\nRemoved: HEAD~1:old.docx\n")

    def test_one_revision_reads_the_work_tree(self) -> None:
        self.doc("a.docx", "Committed.")
        self.doc("b.docx", "Untouched.")
        (self.repo / ".gitignore").write_text("ignored.docx\n", encoding="utf-8")
        self.commit("first")
        self.doc("a.docx", "Uncommitted edit.")
        self.doc("new.docx", "Untracked.")
        self.doc("ignored.docx", "Ignored.")
        self.doc("~$a.docx", "A Word lock file.")
        self.doc(".pilcrow/companion.docx", "Under the output folder.")
        before = self.state()
        printed = self.out("--rev", "HEAD", "--all", code=1)
        self.assertEqual(self.state(), before)
        self.assertTrue(printed.startswith("Review: HEAD:a.docx → a.docx\n\n"), printed)
        self.assertIn("+Uncommitted edit.", printed)
        self.assertTrue(printed.endswith("\nAdded: new.docx\n"), printed)
        for name in ("b.docx", "ignored", "~$", ".pilcrow"):
            self.assertNotIn(name, printed)
        # A document removed from the index but still on disk exists on both sides.
        self.git("rm", "-q", "--cached", "b.docx")
        printed = self.out("--rev", "HEAD", "--all", code=1)
        self.assertIn("Review: HEAD:b.docx → b.docx\n\nProse: no changes\n", printed)
        self.assertNotIn("Added: b.docx", printed)
        # Two revisions never read the work tree.
        self.assertEqual(self.out("--rev", "HEAD", "--rev", "HEAD", "--all", code=0), NONE)

    def test_two_revisions_ignore_the_work_tree(self) -> None:
        self.branch()
        self.doc("docs/a.docx", "An uncommitted edit.")
        printed = self.out("--rev", "main", "--rev", "feature", "--all", code=1)
        self.assertIn("+The supplier may deliver", printed)
        self.assertNotIn("uncommitted", printed)

    def test_merge_base_leaves_out_edits_made_on_the_other_branch(self) -> None:
        self.branch()
        fork = self.git("merge-base", "main", "feature").strip()
        short = self.git("rev-parse", "--short", fork).strip()
        self.git("checkout", "-q", "main")
        self.doc("docs/b.docx", "Edited on main.")
        self.commit("main edit")
        self.git("checkout", "-q", "feature")
        before = self.state()
        direct = self.out("--rev", "main", "--all", code=1)
        self.assertIn("Review: main:docs/b.docx → docs/b.docx", direct)
        for args in (("--rev", "main"), ("--rev", "main", "--rev", "feature")):
            with self.subTest(args=args):
                printed = self.out(*args, "--all", "--merge-base", code=1)
                self.assertNotIn("b.docx", printed)
                self.assertTrue(printed.startswith(f"Review: {short}:docs/a.docx → "), printed)
                self.assertTrue(printed.endswith(f"\nRemoved: {short}:docs/d.docx\n"), printed)
        self.assertEqual(self.state(), before)
        # Unrelated histories have no merge base.
        self.git("checkout", "-q", "--orphan", "lonely")
        self.commit("unrelated")
        result = self.pilcrow("review", "--rev", "main", "--all", "--merge-base", code=2)
        self.assertEqual(result.stdout, b"")
        self.assertIn(b"error: main and HEAD have no common ancestor", result.stderr)

    def test_an_unreadable_document_does_not_stop_the_others(self) -> None:
        self.doc("a.docx", "First.")
        self.doc("b.docx", "Before.")
        self.commit("first")
        (self.repo / "a.docx").write_bytes(b"not a zip")
        self.doc("b.docx", "After.")
        self.commit("second")
        before = self.state()
        result = self.pilcrow("review", "--rev", "HEAD~1", "--rev", "HEAD", "--all", code=2)
        self.assertEqual(self.state(), before)
        self.assertIn(b"error: HEAD:a.docx: cannot read DOCX", result.stderr)
        self.assertTrue(result.stdout.startswith(b"Review: HEAD~1:b.docx \xe2\x86\x92 HEAD:b.docx\n"), result.stdout)
        self.assertIn(b"+After.", result.stdout)
        self.assertNotIn(b"a.docx", result.stdout)
        # The work-tree side of a document is read from the file.
        result = self.pilcrow("review", "--rev", "HEAD~1", "--all", code=2)
        self.assertIn(b"error: a.docx: cannot read DOCX", result.stderr)

    def test_no_changed_documents(self) -> None:
        self.doc("a.docx", "Text.")
        self.commit("first")
        (self.repo / "notes.txt").write_text("not a document\n", encoding="utf-8")
        self.commit("second")
        self.assertEqual(self.out("--rev", "HEAD", "--all", code=0), NONE)
        self.assertEqual(self.out("--rev", "HEAD~1", "--rev", "HEAD", "--all", code=0), NONE)
        self.assertEqual(self.out("--rev", "HEAD~1", "--all", "--format", "markdown", code=0), NONE)

    def test_markdown_and_word_diff_match_the_single_document_form(self) -> None:
        self.branch()
        for flags in (("--format", "markdown"), ("--word-diff",), ("--format", "markdown", "--word-diff")):
            with self.subTest(flags=flags):
                printed = self.out(*flags, "--rev", "main", "--rev", "feature", "--all", code=1)
                single = self.out(*flags, "--rev", "main", "--rev", "feature", Path("docs", "a.docx"), code=1)
                markdown = "--format" in flags
                header = ("## Review: `main:docs/a.docx` → `feature:docs/a.docx`" if markdown
                          else "Review: main:docs/a.docx → feature:docs/a.docx")
                added, removed = (("## Added: `feature:docs/c.docx`", "## Removed: `main:docs/d.docx`") if markdown
                                  else ("Added: feature:docs/c.docx", "Removed: main:docs/d.docx"))
                self.assertEqual(printed, f"{header}\n{self.body(single)}\n{added}\n\n{removed}\n")
                if "--word-diff" in flags:
                    self.assertIn("~ The supplier [-shall-]{+may+} deliver the goods within [-30-]{+60+} days.", printed)

    def test_markdown_escapes_document_text(self) -> None:
        self.doc("a.docx", "Plain.")
        self.commit("first")
        self.doc("a.docx", "``` # <b>bold</b> [link](https://x.example)")
        self.commit("second")
        printed = self.out("--format", "markdown", "--rev", "HEAD~1", "--rev", "HEAD", "--all", code=1)
        self.assertEqual(printed, self.out("--format", "markdown", "--rev", "HEAD~1", "--rev", "HEAD", "a.docx", code=1))
        self.assertIn("````diff", printed)  # the fence outruns the backticks in the document

    def test_a_nested_repository_does_not_hide_a_listed_document(self) -> None:
        self.doc("sub/a.docx", "Before.")
        self.commit("first")
        self.doc("sub/a.docx", "After.")
        self.commit("second")
        self.git("init", "-q", cwd=self.repo / "sub")  # a repository appears on the path later
        for args in (("--rev", "HEAD~1", "--rev", "HEAD"), ("--rev", "HEAD~1")):
            with self.subTest(args=args):
                result = self.pilcrow("review", *args, "--all", code=1)
                self.assertEqual(result.stderr, b"")
                self.assertIn(b"+After.", result.stdout)

    def test_usage_errors(self) -> None:
        self.branch()
        before = self.state()
        cases = [
            (("--rev", "main", "--all", "docs/a.docx"), b"--all reviews every changed DOCX; give no PATH"),
            (("--all",), b"--all needs --rev REV"),
            (("--merge-base", "--rev", "main", "docs/a.docx"), b"--merge-base needs --all"),
            (("--root", "docs", "--rev", "main", "docs/a.docx"), b"--root needs --all"),
            (("--rev", "main", "--rev", "main", "--rev", "main", "--all"), b"give --rev at most twice"),
        ]
        for args, message in cases:
            with self.subTest(args=args):
                result = self.pilcrow("review", *args, code=2)
                self.assertEqual(result.stdout, b"")
                self.assertIn(message, result.stderr)
                self.assertIn(b"usage:", result.stderr)
        self.assertEqual(self.state(), before)

    def test_revision_errors(self) -> None:
        self.branch()
        cases: list[tuple[tuple[str, ...], Path, bytes]] = [
            (("--rev", "no-such-revision", "--all"), self.repo, b"error: unknown revision 'no-such-revision'"),
            (("--rev", "main", "--rev", "nope", "--all"), self.repo, b"error: unknown revision 'nope'"),
            (("--rev=-c", "--all"), self.repo, b'error: -c: a revision must not start with "-"'),
            (("--rev", "main", "--all"), self.outside, b"error: not a git repository"),
        ]
        for args, cwd, message in cases:
            with self.subTest(args=args):
                result = self.pilcrow("review", *args, cwd=cwd, code=2)
                self.assertEqual(result.stdout, b"")
                self.assertIn(message, result.stderr)

    def test_help_shows_the_all_form(self) -> None:
        result = self.pilcrow("review", "--help", code=0)
        self.assertIn(b"--rev REV [--rev REV] --all", result.stdout)
        self.assertIn(b"--merge-base", result.stdout)


if __name__ == "__main__":
    unittest.main()
