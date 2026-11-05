"""Refresh files only. Git mutations intentionally do not exist here."""

import argparse
from pathlib import Path
import sys

from . import __version__
from .extract import DocumentError, extract
from .files import companions, no_symlinks, orphaned, select, setup, write_if_changed


def parser():
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
    return result


def main(argv=None):
    args = parser().parse_args(argv)
    root = args.root.resolve()
    if not root.is_dir():
        print(f"error: --root is not an existing directory: {root}", file=sys.stderr)
        return 2
    try:
        if args.command == "setup":
            changed = setup(root)
            print(f"{'Updated' if changed else 'Unchanged'} {root / '.gitattributes'}; review and commit through Git.")
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
                changed = []
                for path, content in zip(companions(root, source), result.outputs().values()):
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
