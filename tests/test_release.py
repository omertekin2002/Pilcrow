"""Release version changes must agree with the installed CLI."""

from pathlib import Path
import tempfile
import unittest

from scripts.bump_version import bump_version


class ReleaseTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory(prefix="pilcrow-release-test-")
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.metadata = self.root / "pyproject.toml"
        self.module = self.root / "src/pilcrow/__init__.py"
        self.module.parent.mkdir(parents=True)

    def write_versions(self, package: str, cli: str) -> None:
        self.metadata.write_text(
            f'[project]\nname = "pilcrow-docx"\nversion = "{package}"\n',
            encoding="utf-8",
        )
        self.module.write_text(f'__version__ = "{cli}"\n', encoding="utf-8")

    def test_bumps_both_versions_and_preserves_other_metadata(self) -> None:
        self.write_versions("0.1.9", "0.1.9")
        self.assertEqual(bump_version(self.root), "0.1.10")
        self.assertEqual(
            self.metadata.read_text(encoding="utf-8"),
            '[project]\nname = "pilcrow-docx"\nversion = "0.1.10"\n',
        )
        self.assertEqual(
            self.module.read_text(encoding="utf-8"), '__version__ = "0.1.10"\n',
        )
        self.assertEqual(bump_version(self.root), "0.1.11")

    def test_rejects_mismatch_and_prerelease_without_writing(self) -> None:
        for package, cli in [("0.1.0", "0.2.0"), ("0.1.0rc1", "0.1.0rc1")]:
            with self.subTest(package=package, cli=cli):
                self.write_versions(package, cli)
                before = (self.metadata.read_bytes(), self.module.read_bytes())
                with self.assertRaises(ValueError):
                    bump_version(self.root)
                self.assertEqual(
                    (self.metadata.read_bytes(), self.module.read_bytes()), before,
                )
