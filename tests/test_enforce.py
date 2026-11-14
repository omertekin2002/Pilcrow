"""The pre-commit hook's file filter and the fixture repository CI self-tests use."""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]


def hook_files_pattern() -> re.Pattern[str]:
    text = (PROJECT / ".pre-commit-hooks.yaml").read_text(encoding="utf-8")
    patterns = re.findall(r"^\s*files: '([^']*)'$", text, flags=re.MULTILINE)
    if len(patterns) != 1:
        raise AssertionError("expected one single-quoted files: pattern")
    return re.compile(patterns[0])


class HookManifestTests(unittest.TestCase):
    def test_hook_runs_the_read_only_check_without_filenames(self) -> None:
        text = (PROJECT / ".pre-commit-hooks.yaml").read_text(encoding="utf-8")
        self.assertIn("- id: pilcrow-check\n", text)
        self.assertIn("  entry: pilcrow check --all --root .\n", text)
        self.assertIn("  pass_filenames: false\n", text)
        self.assertNotIn("export", re.sub(r"#.*", "", text))

    def test_files_pattern_matches_documents_and_companions(self) -> None:
        # pre-commit applies `files` with re.search to repository-relative,
        # forward-slash paths.
        pattern = hook_files_pattern()
        for path in [
            "report.docx", "REPORT.DOCX", "chapters/Methods and results.Docx",
            ".pilcrow/report.docx.prose.md", ".pilcrow/chapters/x.docx.meta.json",
        ]:
            with self.subTest(path=path):
                self.assertIsNotNone(pattern.search(path))
        for path in [
            "README.md", "report.docx.bak", "notes/report.doc", "docx/readme.txt",
            "sub/.pilcrow/report.docx.prose.md", "x.pilcrow/y",
        ]:
            with self.subTest(path=path):
                self.assertIsNone(pattern.search(path))


class FixtureRepoTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory(prefix="pilcrow-fixture-test-")
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name) / "fixture"
        self.env = {**os.environ}
        self.env["PYTHONPATH"] = os.pathsep.join(
            (str(PROJECT / "src"), self.env.get("PYTHONPATH", "")),
        )

    def run_python(self, *args: str | Path) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, *map(str, args)], capture_output=True, text=True,
            env=self.env, check=False,
        )

    def check(self) -> subprocess.CompletedProcess[str]:
        return self.run_python("-m", "pilcrow", "check", "--all", "--root", self.root)

    def test_fresh_fixture_passes_and_stale_fixture_fails(self) -> None:
        script = PROJECT / "scripts" / "make_fixture_repo.py"
        made = self.run_python(script, self.root)
        self.assertEqual(made.returncode, 0, made.stderr)
        status = subprocess.run(
            ["git", "-C", str(self.root), "status", "--porcelain"],
            capture_output=True, text=True, check=True,
        )
        self.assertEqual(status.stdout, "", "fresh fixture should be fully committed")
        result = self.check()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

        made = self.run_python(script, self.root, "--stale")
        self.assertEqual(made.returncode, 0, made.stderr)
        result = self.check()
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("report.docx", result.stdout + result.stderr)

    def test_refuses_a_non_empty_directory(self) -> None:
        self.root.mkdir()
        (self.root / "keep.txt").write_text("x", encoding="utf-8")
        made = self.run_python(PROJECT / "scripts" / "make_fixture_repo.py", self.root)
        self.assertNotEqual(made.returncode, 0)
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ["keep.txt"])


if __name__ == "__main__":
    unittest.main()
