"""The README quickstart, replayed command by command in a temporary Git repository.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

try:
    from .test_cli import PROJECT, make_docx, paragraph
except ImportError:  # discovered with -s tests: the directory itself is on sys.path
    from test_cli import PROJECT, make_docx, paragraph  # type: ignore[no-redef]

DOCUMENT = "report.docx"
PROSE = Path(".pilcrow") / f"{DOCUMENT}.prose.md"
META = Path(".pilcrow") / f"{DOCUMENT}.meta.json"


def work_tree(repo: Path) -> list[tuple[str, int, int]]:
    """Every file and directory in the work tree with its size and mtime, leaving out ``.git``.

    Git may still be tidying ``.git`` in the background after a commit, so it is not compared.
    """
    listing = []
    for directory, folders, files in os.walk(repo):
        if Path(directory) == repo:
            folders[:] = [folder for folder in folders if folder != ".git"]
        for name in folders + files:
            item = Path(directory) / name
            info = item.stat()
            listing.append((item.relative_to(repo).as_posix(), info.st_size, info.st_mtime_ns))
    return sorted(listing)


def quickstart_section() -> str:
    """The README text from ``## Quickstart`` up to the next second-level heading."""
    readme = (PROJECT / "README.md").read_text(encoding="utf-8")
    match = re.search(r"^## Quickstart\n(.*?)(?=^## )", readme, re.MULTILINE | re.DOTALL)
    assert match is not None, "README.md has no ## Quickstart section"
    return match.group(1)


def github_anchors(markdown: str) -> set[str]:
    """Heading anchors as GitHub generates them for the headings outside code fences."""
    anchors = set()
    fence: str | None = None
    for line in markdown.splitlines():
        marker = re.match(r"^(`{3,}|~{3,})", line)
        if marker:
            if fence is None:
                fence = marker.group(1)
            elif line.startswith(fence) and not line[len(fence):].strip():
                fence = None
            continue
        heading = re.match(r"^#{1,6} +(.+?) *$", line)
        if fence is None and heading:
            slug = re.sub(r"[^\w\- ]", "", heading.group(1).lower()).replace(" ", "-")
            anchors.add(slug)
    return anchors


class QuickstartTests(unittest.TestCase):
    """Each step of the README quickstart, run through the CLI entry point as a reader would."""

    def setUp(self) -> None:
        if shutil.which("git") is None:
            self.skipTest("git is not installed")
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-quickstart-")
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.repo = base / "repo"
        self.repo.mkdir()
        self.env = {**os.environ, "GIT_CEILING_DIRECTORIES": str(base)}
        self.env["PYTHONPATH"] = os.pathsep.join((str(PROJECT / "src"), self.env.get("PYTHONPATH", "")))
        for args in (("init", "-q"), ("config", "user.name", "Pilcrow Test"),
                     ("config", "user.email", "pilcrow@example.invalid"), ("config", "commit.gpgsign", "false"),
                     ("config", "maintenance.auto", "false"), ("config", "gc.auto", "0")):
            self.git(*args)

    def git(self, *args: str) -> str:
        result = subprocess.run(["git", *args], cwd=self.repo, env=self.env, capture_output=True, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.decode("utf-8")

    def pilcrow(self, *args: str, code: int) -> str:
        result = subprocess.run([sys.executable, "-m", "pilcrow", *args], cwd=self.repo, env=self.env,
                                capture_output=True, timeout=60)
        self.assertEqual(result.returncode, code, f"args={args!r}\nstdout={result.stdout!r}\nstderr={result.stderr!r}")
        return (result.stdout + result.stderr).decode("utf-8")

    def write(self, *sentences: str) -> None:
        make_docx(self.repo / DOCUMENT, "".join(paragraph(sentence) for sentence in sentences))

    def test_quickstart_steps(self) -> None:
        # A Word file already in Git, committed without companions.
        self.write("The cat sat on the mat.", "It was warm.")
        self.git("add", DOCUMENT)
        self.git("commit", "-q", "-m", "Add the report")

        # A copy of the committed file outside the repository, for the two-file form.
        old = self.repo.parent / "old.docx"
        shutil.copy(self.repo / DOCUMENT, old)

        # Step 1: review a committed change with no setup. It reports the edit and writes nothing.
        self.write("The dog sat on the mat.", "It was warm.")
        before = work_tree(self.repo)
        report = self.pilcrow("review", "--rev", "HEAD", DOCUMENT, code=1)
        self.assertIn(f"Review: HEAD:{DOCUMENT} → {DOCUMENT}", report)
        self.assertIn("-The cat sat on the mat.", report)
        self.assertIn("+The dog sat on the mat.", report)
        self.assertEqual(work_tree(self.repo), before)
        self.assertFalse((self.repo / ".pilcrow").exists())
        # The two-file form reports the same edit between the committed copy and the working copy.
        self.assertIn("+The dog sat on the mat.", self.pilcrow("review", str(old), DOCUMENT, code=1))
        self.assertEqual(work_tree(self.repo), before)

        # Step 2: export a baseline; check passes and names the document. Commit all three files.
        self.assertIn(f"EXPORTED {DOCUMENT}", self.pilcrow("export", DOCUMENT, "--root", ".", code=0))
        self.assertTrue((self.repo / PROSE).is_file())
        self.assertTrue((self.repo / META).is_file())
        self.assertIn(f"FRESH {DOCUMENT}", self.pilcrow("check", DOCUMENT, "--root", ".", code=0))
        self.git("add", DOCUMENT, ".pilcrow")
        self.git("commit", "-q", "-m", "Edit the report with companions")
        self.assertEqual(self.git("status", "--porcelain"), "")

        # Step 3: edit without exporting, and the check fails; export again and the prose diff shows the edit.
        self.write("The dog sat on the rug.", "It was warm.")
        self.assertIn(f"STALE/MISSING {DOCUMENT}", self.pilcrow("check", DOCUMENT, "--root", ".", code=1))
        self.pilcrow("export", DOCUMENT, "--root", ".", code=0)
        self.assertIn(f"FRESH {DOCUMENT}", self.pilcrow("check", DOCUMENT, "--root", ".", code=0))
        diff = self.git("diff", "--", ".pilcrow/")
        self.assertIn("-The dog sat on the mat.", diff)
        self.assertIn("+The dog sat on the rug.", diff)
        words = self.git("diff", "--word-diff", "--", PROSE.as_posix())
        self.assertIn("[-mat.-]{+rug.+}", words)

        # Step 4: the freshness gate runs the same check over every document.
        self.assertIn(f"FRESH {DOCUMENT}", self.pilcrow("check", "--all", "--root", ".", code=0))
        self.write("The dog sat on the sofa.", "It was warm.")
        self.pilcrow("check", "--all", "--root", ".", code=1)


class QuickstartReadmeTests(unittest.TestCase):
    """The quickstart section itself: placement, length and links."""

    def test_quickstart_follows_install_and_stays_short(self) -> None:
        readme = (PROJECT / "README.md").read_text(encoding="utf-8")
        headings = re.findall(r"^## (.+)$", readme, re.MULTILINE)
        install = headings.index("Install")
        self.assertEqual(headings[install + 1], "Quickstart")
        self.assertLessEqual(len(quickstart_section().splitlines()), 80)

    def test_quickstart_shows_the_documented_commands(self) -> None:
        section = quickstart_section()
        for command in (f'pilcrow review --rev HEAD "{DOCUMENT}"', "pilcrow review old.docx new.docx",
                        f'pilcrow export "{DOCUMENT}" --root .', f'pilcrow check "{DOCUMENT}" --root .',
                        "git diff --word-diff", "id: pilcrow-check", "uses: omertekin2002/Pilcrow@"):
            self.assertIn(command, section)
        self.assertIn("writes nothing", section)

    def test_quickstart_links_resolve(self) -> None:
        readme = (PROJECT / "README.md").read_text(encoding="utf-8")
        anchors = github_anchors(readme)
        links = re.findall(r"\]\(([^)\s]+)\)", quickstart_section())
        self.assertTrue(links)
        for link in links:
            with self.subTest(link=link):
                if link.startswith("#"):
                    self.assertIn(link[1:], anchors)
                elif not link.startswith("https://"):
                    self.assertTrue((PROJECT / link.split("#")[0]).is_file(), link)

    def test_release_pins_name_a_release_with_the_quickstart(self) -> None:
        """Every snippet pins one release, new enough for `--rev` and the action's `summary` input."""
        readme = (PROJECT / "README.md").read_text(encoding="utf-8")
        pins = set(re.findall(r"(?:rev: |omertekin2002/Pilcrow@)(v\d+\.\d+\.\d+)", readme))
        self.assertEqual(pins, {"v0.1.8"})


if __name__ == "__main__":
    unittest.main()
