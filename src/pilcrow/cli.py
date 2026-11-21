"""Refresh files only. Git mutations intentionally do not exist here."""

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from . import __version__
from .extract import MAX_PACKAGE_BYTES, DocumentError, extract
from .files import (
    TEXTCONV_CONFIG,
    companions,
    no_symlinks,
    orphaned,
    revision_blob,
    select,
    setup,
    write_if_changed,
)
from .review import FORMATS, build, markdown, text


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description="Generate readable DOCX companions for GitHub Desktop.")
    result.add_argument("--version", action="version", version=__version__)
    commands = result.add_subparsers(dest="command", required=True)
    for name in ("export", "check"):
        command = commands.add_parser(name, help="refresh companions" if name == "export" else "read-only freshness check")
        command.add_argument("paths", nargs="*", help="DOCX paths, relative to current directory")
        command.add_argument("--root", type=Path, default=Path.cwd(), help="scope and output root (default: current directory)")
        command.add_argument("--all", action="store_true", help="discover DOCX files recursively within --root")
    command = commands.add_parser("setup", help="opt in to LF text and conservative merge rules in .gitattributes")
    command.add_argument("--root", type=Path, default=Path.cwd())
    command.add_argument("--textconv", action="store_true",
                         help="also mark DOCX files diff=pilcrow so git diff can show their prose (needs a one-time git config)")
    command = commands.add_parser("textconv", help="print one DOCX's prose to stdout (a Git textconv driver)")
    command.add_argument("path", type=Path, help="any DOCX file, such as the temporary file Git passes")
    command = commands.add_parser(
        "review", help="describe prose and review changes between two DOCX files in words (read-only)",
        usage="%(prog)s [--format {text,markdown}] [--word-diff] OLD NEW\n"
              "       %(prog)s [--format {text,markdown}] [--word-diff] --rev REV [--rev REV] PATH",
        description="Compare OLD with NEW; or PATH at REV with the PATH in the work tree; "
                    "or PATH at the first REV with PATH at the second, without reading the work tree.")
    command.add_argument("paths", type=Path, nargs="+", metavar="PATH",
                         help="OLD and NEW (any DOCX files, such as the temporary ones git difftool passes), "
                              "or with --rev one DOCX path relative to the current directory")
    command.add_argument("--rev", action="append", default=[], metavar="REV",
                         help="read PATH as committed at this Git revision; give it once or twice, as in git diff")
    command.add_argument("--format", choices=FORMATS, default="text",
                         help="text (the default) or markdown, for pasting into a pull request, issue or chat")
    command.add_argument("--word-diff", action="store_true",
                         help="print a changed sentence once, after '~ ', with removed words as [-...-] and added "
                              "words as {+...+}. Sentences pair in order only within a replaced block with as many "
                              "new lines as old, when at least half of each pair's words match; a block that does "
                              "not pair, or has a line over 200 words or already containing a mark, is shown as "
                              "whole - and + lines, as without this flag")
    command.set_defaults(usage_error=command.error)  # report usage errors with the review usage
    return result


def textconv(path: Path) -> int:
    """Print exactly the bytes of the prose companion; never write files.

    Git passes temporary copies of old revisions with arbitrary names outside
    any repository, so no scope, discovery, extension, or symlink rules apply.
    """
    try:
        result = extract(path)
    except (OSError, DocumentError) as exc:
        print(f"error: {path}: {exc}", file=sys.stderr)
        return 2
    for warning in result.metadata["warnings"]:
        print(f"warning: {path}: {warning}", file=sys.stderr)
    # Raw UTF-8 bytes, not the console encoding, so output is identical on Windows.
    sys.stdout.buffer.write(result.outputs()[".prose.md"])
    sys.stdout.buffer.flush()
    return 0


def review_command(paths: list[Path], revisions: list[str], output: str = "text", marks: bool = False) -> int:
    """Print a report of what changed from the old side to the new one; never write files.

    Each side is a DOCX file, or with ``--rev`` the DOCX blob committed at a
    revision, read with read-only Git plumbing. Like textconv, no scope,
    discovery, extension, or symlink rules apply, because git difftool passes
    temporary copies with arbitrary names. Exit codes follow diff: 0 no
    supported differences, 1 differences reported, 2 an input failed. ``output``
    chooses the format; it changes how the report is printed, never what it says.
    ``marks`` (``--word-diff``) marks the changed words of replaced prose lines that pair up.
    """
    sides: list[tuple[str, Path, str | None]]
    if revisions:
        (path,) = paths
        sides = [(f"{rev}:{path}", path, rev) for rev in revisions]
        if len(sides) == 1:
            sides.append((str(path), path, None))
    else:
        sides = [(str(path), path, None) for path in paths]
    results = []
    for label, path, rev in sides:
        try:
            results.append(extract(path if rev is None else revision_blob(path, rev, MAX_PACKAGE_BYTES)))
        except (OSError, ValueError) as exc:  # DocumentError is a ValueError
            print(f"error: {label}: {exc}", file=sys.stderr)
    if len(results) != 2:
        return 2  # no partial report
    report = build(results[0], results[1], marks)
    old, new = sides[0][0], sides[1][0]
    printed = markdown(report, old, new) if output == "markdown" else f"Review: {old} → {new}\n\n{text(report)}"
    # Raw UTF-8 bytes, not the console encoding, so output is identical on Windows.
    sys.stdout.buffer.write(printed.encode("utf-8", "backslashreplace"))
    sys.stdout.buffer.flush()
    return 1 if report.changed else 0


def main(argv: Sequence[str] | None = None) -> int:
    # A piped stdout on Windows uses the ANSI code page; escape file names it
    # cannot encode, as stderr already does, instead of failing the document.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="backslashreplace")
    args = parser().parse_args(argv)
    if args.command == "textconv":
        return textconv(args.path)
    if args.command == "review":
        if len(args.rev) > 2:
            args.usage_error("give --rev at most twice")
        if len(args.paths) != (1 if args.rev else 2):
            args.usage_error(f"{'--rev takes one PATH' if args.rev else 'give OLD and NEW, or --rev REV PATH'}, "
                             f"not {len(args.paths)}")
        return review_command(args.paths, args.rev, args.format, args.word_diff)
    root = args.root.resolve()
    if not root.is_dir():
        print(f"error: --root is not an existing directory: {root}", file=sys.stderr)
        return 2
    try:
        if args.command == "setup":
            updated, diff_driver = setup(root, args.textconv)
            print(f"{'Updated' if updated else 'Unchanged'} {root / '.gitattributes'}; review and commit through Git.")
            if diff_driver:
                print("DOCX files are marked diff=pilcrow. Pilcrow does not change Git config; to make git diff show"
                      f" their prose, run once in this repository:\n  {TEXTCONV_CONFIG}\n"
                      "Optionally cache the converted text: git config diff.pilcrow.cachetextconv true")
            return 0
        if bool(args.paths) == args.all:
            raise ValueError("specify DOCX paths OR --all; use --root to choose discovery/output scope")
        sources = select(root, args.paths, args.all)
        status = 0
        for source in sources:
            label = source.relative_to(root).as_posix()
            try:
                no_symlinks(source, root)
                result = extract(source)
                changed: list[str] = []
                for path, content in zip(companions(root, source), result.outputs().values(), strict=True):
                    no_symlinks(path, root)
                    if args.command == "export":
                        if write_if_changed(path, content, root):
                            changed.append(path.name)
                    elif not path.exists() or path.read_bytes() != content:
                        changed.append(path.name)
                if args.command == "check" and changed:
                    print(f"STALE/MISSING {label}: {', '.join(changed)}; run pilcrow export with the same scope")
                    status = max(status, 1)
                elif changed:
                    print(f"EXPORTED {label}: {', '.join(changed)}")
                else:
                    print(f"FRESH {label}: no normalized prose/review changes detected (not a complete document comparison)")
                for warning in result.metadata["warnings"]:
                    print(f"warning: {label}: {warning}", file=sys.stderr)
            except (OSError, ValueError, DocumentError) as exc:
                print(f"error: {label}: {exc}", file=sys.stderr)
                status = 2
        if args.all:
            for path in orphaned(root):
                print(f"ORPHAN {path.relative_to(root)}: source deleted or renamed; remove this companion manually if intentional", file=sys.stderr)
                status = max(status, 1)
        if not sources:
            print("No DOCX documents selected within this scope.")
        return status
    except (OSError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
