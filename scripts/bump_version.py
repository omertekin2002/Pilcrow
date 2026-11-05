"""Keep package metadata and the CLI version in sync for patch releases."""

from pathlib import Path
import re


def bump_version(root: Path) -> str:
    metadata_path = root / "pyproject.toml"
    module_path = root / "src/pilcrow/__init__.py"
    metadata = metadata_path.read_text(encoding="utf-8")
    module = module_path.read_text(encoding="utf-8")
    versions = re.findall(r'^version = "([^"]+)"$', metadata, flags=re.MULTILINE)
    if len(versions) != 1:
        raise ValueError("Expected one package version declaration")
    current = versions[0]
    if not re.fullmatch(r"\d+\.\d+\.\d+", current):
        raise ValueError(f"Expected a stable major.minor.patch version, got {current!r}")

    major, minor, patch = map(int, current.split("."))
    version = f"{major}.{minor}.{patch + 1}"
    metadata, metadata_count = re.subn(
        rf'^version = "{re.escape(current)}"$',
        f'version = "{version}"', metadata, flags=re.MULTILINE,
    )
    module, module_count = re.subn(
        rf'^__version__ = "{re.escape(current)}"$',
        f'__version__ = "{version}"', module, flags=re.MULTILINE,
    )
    if metadata_count != 1 or module_count != 1:
        raise ValueError("Expected one matching version declaration in each file")

    metadata_path.write_text(metadata, encoding="utf-8")
    module_path.write_text(module, encoding="utf-8")
    return version


if __name__ == "__main__":
    print(bump_version(Path(__file__).resolve().parents[1]))
