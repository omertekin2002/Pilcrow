"""Explicit scope, Git-aware discovery, and atomic companion writes."""

import os
import subprocess
import tempfile
from collections.abc import Iterable
from pathlib import Path

OUTPUT = ".pilcrow"
EXCLUDED = {OUTPUT, ".git", ".venv", "venv", "env", "node_modules", "__pycache__", ".tox", ".nox", "build", "dist"}
SUFFIXES = (".prose.md", ".meta.json")


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes] | None:
    try:
        return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)
    except FileNotFoundError:
        return None


def eligible(path: Path) -> bool:
    return (path.suffix.lower() == ".docx" and not path.name.startswith("~$")
            and not any(p in EXCLUDED for p in path.parts))


def no_symlinks(path: Path, root: Path) -> None:
    """Do not follow inputs or outputs outside scope through symlink components."""
    relative = path.relative_to(root)
    current = root
    for component in relative.parts:
        current = current / component
        if current.is_symlink():
            raise ValueError(f"symlink paths are not supported: {current}")


def discover(root: Path) -> list[Path]:
    probe = git(root, "rev-parse", "--is-inside-work-tree")
    has_git = any((p / ".git").exists() for p in (root, *root.parents))
    unavailable = "Git discovery unavailable in this repository; install/check Git or select explicit paths"
    if has_git and (probe is None or probe.returncode != 0):
        raise ValueError(unavailable)
    if probe and probe.returncode == 0 and probe.stdout.strip() == b"true":
        listing = git(root, "ls-files", "-z", "--cached", "--others", "--exclude-standard", "--", ".")
        if listing is None:  # Git vanished after the probe above succeeded.
            raise ValueError(unavailable)
        if listing.returncode:
            raise ValueError("Git discovery failed: " + os.fsdecode(listing.stderr).strip())
        listed = {Path(os.fsdecode(p)) for p in listing.stdout.split(b"\0") if p}
        # Deleted tracked inputs are handled by the orphan scan, not resurrected.
        selected = []
        for path in sorted(listed):
            if not eligible(path):
                continue
            source = root / path
            no_symlinks(source, root)
            try:
                source.stat()
            except FileNotFoundError:
                continue
            selected.append(source)
        return selected
    paths: list[Path] = []
    def onerror(error: OSError) -> None:
        raise error
    for directory, dirs, files in os.walk(root, onerror=onerror, followlinks=False):
        dirs[:] = sorted(d for d in dirs if d not in EXCLUDED and not (Path(directory) / d).is_symlink())
        paths.extend(Path(directory) / f for f in files if eligible(Path(f)))
    return sorted(paths, key=str)


def within(root: Path, path: Path) -> Path | None:
    """Return ``path`` relative to the resolved ``root``, or None if outside it.

    ``root`` is resolved, but a user's path may spell it through an alias, such
    as macOS ``/var`` -> ``/private/var`` or a Windows 8.3 short name. Only the
    ancestor that names the root itself is resolved; components inside the root
    stay lexical so ``no_symlinks`` still checks them.
    """
    try:
        return path.relative_to(root)
    except ValueError:
        pass
    for ancestor in reversed(path.parents):  # shortest first
        try:
            resolved = ancestor.resolve()
        except (OSError, RuntimeError):  # RuntimeError: symlink loop before 3.13
            continue
        if resolved == root:
            return path.relative_to(ancestor)
    return None


def select(root: Path, paths: Iterable[str], all_documents: bool) -> list[Path]:
    if all_documents:
        return discover(root)
    selected = []
    for raw in paths:
        path = Path(os.path.abspath(raw))  # lexical path: do not resolve a hidden symlink
        relative = within(root, path)
        if relative is None:
            raise ValueError(f"{raw}: outside --root {root}")
        if not eligible(relative):
            raise ValueError(f"{raw}: expected a DOCX outside lock/output/environment paths")
        selected.append(root / relative)
    return sorted(set(selected), key=str)


def companions(root: Path, source: Path) -> list[Path]:
    stem = root / OUTPUT / source.relative_to(root)
    return [Path(str(stem) + suffix) for suffix in SUFFIXES]


def orphaned(root: Path) -> list[Path]:
    base = root / OUTPUT
    no_symlinks(base, root)
    if not base.exists():
        return []
    result = []
    for directory, dirs, files in os.walk(base, followlinks=False):
        for name in dirs:
            no_symlinks(Path(directory) / name, root)
        for name in files:
            path = Path(directory) / name
            for suffix in SUFFIXES:
                if name.endswith(suffix):
                    source = root / str(path.relative_to(base))[:-len(suffix)]
                    if not source.exists():
                        result.append(path)
    return sorted(result, key=str)


def write_if_changed(path: Path, content: bytes, root: Path) -> bool:
    no_symlinks(path, root)
    if path.exists() and path.read_bytes() == content:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".pilcrow-tmp-", delete=False) as handle:
            temporary = Path(handle.name)
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return True


START = b"# BEGIN Pilcrow generated review rules"
END = b"# END Pilcrow generated review rules"
RULES = b"\n".join([
    START,
    b"*.[dD][oO][cC][xX] -merge",
    b".pilcrow/** text eol=lf -merge",
    END,
]) + b"\n"


def setup(root: Path) -> bool:
    path = root / ".gitattributes"
    no_symlinks(path, root)
    old = path.read_bytes() if path.exists() else b""
    if START in old or END in old:
        if old.count(START) != 1 or old.count(END) != 1 or old.index(END) < old.index(START):
            raise ValueError("malformed Pilcrow block in .gitattributes; repair it before setup")
        start = old.index(START)
        end = old.index(END) + len(END)
        if old[end:end+2] == b"\r\n":
            end += 2
        elif old[end:end+1] == b"\n":
            end += 1
        new = old[:start] + RULES + old[end:]
    else:
        new = old + (b"\n" if old and not old.endswith(b"\n") else b"") + RULES
    return write_if_changed(path, new, root)
