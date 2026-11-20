"""Write the GitHub Action's job summary: a review of each DOCX a push or pull request changed.

Internal to the Action (``action.yml``), not a ``pilcrow`` subcommand::

    python -m pilcrow.summary --base REV --root ROOT --mode outline|full --output FILE

Read-only: it lists changed documents and reads both sides with read-only Git
plumbing, and writes no file except ``--output``. Problems with the base revision
or a single document become an explanation in the summary and a ``::warning::``
annotation, never a failed job: the Action's outcome stays ``pilcrow check``'s.

``outline`` publishes no document text: only paths, statuses, and the titles of the
report sections that differ, with counts. ``full`` publishes each document's
``pilcrow review --format markdown`` report.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from .extract import MAX_PACKAGE_BYTES, extract
from .files import eligible, git_output, revision_blob
from .review import Report, build, code, escape, markdown

MODES = ("outline", "full")
SHORT = 7  # commit IDs in the summary are shortened to this many characters, as on GitHub
# GitHub refuses a step summary over 1 MiB. Reports stop at BUDGET bytes; the lines
# that list the documents left out stop at LIMIT, then one line counts the rest.
BUDGET = 900_000
LIMIT = 1_000_000
TITLE = "## Pilcrow review summary"
FULL_NOTE = ("**This summary includes document text.** Anyone who can view this workflow run can read it. "
             "Use `summary: outline` to publish only which documents and sections changed.")
OMITTED = "report omitted: summary size limit"


@dataclass
class Document:
    """One changed DOCX: its path relative to the root, and what the summary says about it."""

    path: str
    status: str  # changed, no supported changes, added, removed, could not be read
    report: Report | None = None
    error: str | None = None


def annotate(message: str) -> None:
    """Print a GitHub Actions warning annotation, escaped so it stays one command."""
    value = message.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
    print(f"::warning::{value}", flush=True)


def resolve(root: Path, revision: str) -> str | None:
    """The full commit ID ``revision`` names, or None if it names no commit here."""
    if revision.startswith("-"):
        return None
    found = git_output(root, "rev-parse", "--verify", "--quiet", f"{revision}^{{commit}}")
    return found.decode("ascii").strip() if found else None


def shallow(root: Path) -> bool:
    try:
        return (git_output(root, "rev-parse", "--is-shallow-repository") or b"").strip() == b"true"
    except ValueError:
        return False


def changed(root: Path, base: str, head: str) -> list[tuple[str, str]]:
    """(status letter, path relative to the root) of each eligible DOCX that differs, sorted by path."""
    # --relative limits the listing to the root and prints paths relative to it; no rename
    # detection, so a renamed document is listed as removed and added.
    output = git_output(root, "diff", "--no-ext-diff", "--name-status", "--no-renames", "--relative", "-z", base, head, "--", ".")
    fields = [os.fsdecode(f) for f in (output or b"").split(b"\0") if f]
    pairs = list(zip(fields[::2], fields[1::2], strict=True))
    return sorted(((status[:1], path) for status, path in pairs if eligible(Path(path))), key=lambda p: p[1])


def compare(root: Path, path: str, base: str, head: str) -> Document:
    """Read ``path`` at both revisions and compare them; a read error is reported, not raised."""
    try:
        old, new = (extract(revision_blob(root / path, rev, MAX_PACKAGE_BYTES)) for rev in (base, head))
    except (OSError, ValueError) as exc:  # DocumentError is a ValueError
        return Document(path, "could not be read", error=str(exc))
    report = build(old, new)
    return Document(path, "changed" if report.changed else "no supported changes", report)


def sections(report: Report) -> list[str]:
    """The titles of the sections that differ, each with its count of described lines; no document text."""
    counts: dict[str, int] = {}
    for section in report.sections:
        counts[section.title] = counts.get(section.title, 0) + len(section.lines)
    return (["Prose changed"] if report.diff else []) + [f"{title} ({n})" for title, n in counts.items()]


def outline(document: Document) -> str:
    line = f"- {code(document.path)} — {document.status}"
    if document.error is not None:
        return f"{line}: {escape(document.error)}"
    if document.report is not None and document.report.changed:
        return f"{line}: {', '.join(sections(document.report))}"
    return line


def block(document: Document, mode: str, base: str, head: str) -> str:
    if mode == "full" and document.report is not None:
        return markdown(document.report, f"{base}:{document.path}", f"{head}:{document.path}").rstrip("\n")
    return outline(document)


def header(mode: str, root: str, base: str, head: str) -> list[str]:
    where = f" in {code(root)}" if root != "." else ""
    lines = [TITLE, f"Changed DOCX documents{where} from {code(base)} (base) to {code(head)} (head)."]
    if mode == "full":
        lines.append(FULL_NOTE)
    return lines


def problem(reason: str, plain: str | None = None) -> str:
    """A summary with no reports that says why (``reason`` in Markdown); also a warning annotation."""
    annotate(f"Pilcrow summary: no review: {plain or reason}")
    return f"{TITLE}\n\nNo review: {reason}\n"


def summarize(root: Path, base: str, mode: str, label: str = ".") -> str:
    """The summary Markdown for ``root``, shown as ``label``. Never raises for a base or document problem."""
    base = base.strip()
    if not base or re.fullmatch(r"0+", base):
        return problem("there is no base revision to compare with, as on the first push of a new branch.")
    try:
        head = resolve(root, "HEAD")
        if head is None:
            return problem("HEAD names no commit in this repository.")
        old = resolve(root, base)
        if old is None:
            hint = "The checkout is shallow. " if shallow(root) else ""
            hint += "Fetch history that reaches the base, for example with {} on actions/checkout."
            return problem(f"the base revision {code(base)} is not in this checkout. " + hint.format(code("fetch-depth: 0")),
                           f"the base revision {base} is not in this checkout. " + hint.format("fetch-depth: 0"))
        listing = changed(root, old, head)
    except ValueError as exc:  # Git missing, not a repository, or a failed diff
        return problem(f"Git could not list the changes: {escape(str(exc))}", f"Git could not list the changes: {exc}")
    short_base, short_head = old[:SHORT], head[:SHORT]
    lines = header(mode, label, short_base, short_head)
    if not listing:
        return "\n\n".join([*lines, "No DOCX documents changed."]) + "\n"
    blocks: list[str] = []
    size = len("\n\n".join(lines).encode("utf-8"))
    over = False
    for index, (status, path) in enumerate(listing):
        if not over:
            if status == "A":
                document = Document(path, "added")
            elif status == "D":
                document = Document(path, "removed")
            else:
                document = compare(root, path, old, head)
                if document.error is not None:
                    annotate(f"Pilcrow could not read {path}: {document.error}")
            text = block(document, mode, short_base, short_head)
            cost = len(text.encode("utf-8")) + 2
            if size + cost <= BUDGET:
                blocks.append(text)
                size += cost
                continue
            over = True
        text = f"- {code(path)} — {OMITTED}"
        cost = len(text.encode("utf-8")) + 2
        if size + cost > LIMIT:
            blocks.append(f"- {len(listing) - index} more changed documents not listed: summary size limit")
            break
        blocks.append(text)
        size += cost
    if over:
        annotate("Pilcrow summary: some reports were omitted to stay under the job summary size limit.")
    # Outline lines form one list; full reports are separate blocks.
    return "\n\n".join([*lines, ("\n\n" if mode == "full" else "\n").join(blocks)]) + "\n"


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m pilcrow.summary",
                                     description="Write the Pilcrow Action's job summary (internal).")
    parser.add_argument("--base", required=True, help="the revision to compare from; the new side is HEAD")
    parser.add_argument("--root", default=".", help="document repository root (default: current directory)")
    parser.add_argument("--mode", choices=MODES, required=True)
    parser.add_argument("--output", type=Path, required=True, help="the Markdown file to write")
    args = parser.parse_args(argv)
    root = Path(os.path.abspath(args.root))
    if not root.is_dir():
        content = problem(f"--root is not an existing directory: {code(args.root)}",
                          f"--root is not an existing directory: {args.root}")
    else:
        content = summarize(root, args.base, args.mode, args.root)
    try:
        args.output.write_text(content, encoding="utf-8", newline="\n")
    except OSError as exc:
        print(f"error: {args.output}: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
