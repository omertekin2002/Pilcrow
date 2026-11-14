"""Build a tiny document repository for self-testing the hook and the action.

Usage (Pilcrow must be importable by this Python, e.g. after ``pip install .``)::

    python scripts/make_fixture_repo.py DIR           # fresh: DOCX + exported companions, committed
    python scripts/make_fixture_repo.py DIR --stale   # edit and stage the DOCX without re-exporting

The DOCX is generated deterministically; no binary fixture is checked in.
"""

import argparse
import subprocess
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT = "report.docx"
FRESH_TEXT = "The fixture report says one thing."
STALE_TEXT = "The fixture report now says something else."
# Fixed identity and settings so commits work on clean CI runners and ignore user config.
GIT = [
    "git",
    "-c", "user.name=Pilcrow fixture",
    "-c", "user.email=fixture@example.invalid",
    "-c", "commit.gpgsign=false",
    "-c", "core.hooksPath=/dev/null",
]


def write_docx(path: Path, text: str) -> None:
    """Write a minimal, self-contained DOCX with one paragraph."""
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
        "word/document.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
            f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
            '</w:body></w:document>'
        ),
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{PKG_R}"></Relationships>',
    }
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in sorted(parts.items()):
            info = ZipInfo(name, date_time=(2020, 1, 2, 3, 4, 6))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, content.encode("utf-8"))


def run(*args: str, cwd: Path) -> None:
    subprocess.run(args, cwd=cwd, check=True)


def make_fresh(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    if any(root.iterdir()):
        raise SystemExit(f"{root}: expected a new or empty directory")
    run(*GIT, "init", "--quiet", cwd=root)
    write_docx(root / DOCUMENT, FRESH_TEXT)
    run(sys.executable, "-m", "pilcrow", "export", "--all", "--root", ".", cwd=root)
    run(*GIT, "add", "--all", cwd=root)
    run(*GIT, "commit", "--quiet", "-m", "Add fixture report with fresh companions", cwd=root)


def make_stale(root: Path) -> None:
    if not (root / DOCUMENT).is_file():
        raise SystemExit(f"{root}: run without --stale first")
    write_docx(root / DOCUMENT, STALE_TEXT)
    run(*GIT, "add", DOCUMENT, cwd=root)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--stale", action="store_true", help="edit and stage the DOCX without re-exporting")
    args = parser.parse_args(argv)
    (make_stale if args.stale else make_fresh)(args.directory)


if __name__ == "__main__":
    main()
