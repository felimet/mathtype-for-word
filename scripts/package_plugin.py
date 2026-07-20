#!/usr/bin/env python3
"""Create a reproducible ZIP of the complete cross-agent plugin."""

from __future__ import annotations

import hashlib
import sys
import zipfile
from pathlib import Path


EXCLUDED_DIRS = {".git", "__pycache__", ".pytest_cache", "dist", "mathtype-for-word-workspace"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def should_include(path: Path, root: Path) -> bool:
    relative = path.relative_to(root)
    if any(part in EXCLUDED_DIRS for part in relative.parts) or path.suffix in EXCLUDED_SUFFIXES:
        return False
    return relative.parts[:2] != ("evals", "results")


def package(root: Path, output: Path) -> str:
    root = root.resolve()
    output = output.resolve()
    required = [
        root / ".codex-plugin" / "plugin.json",
        root / ".claude-plugin" / "plugin.json",
        root / ".mcp.json",
        root / "skills" / "mathtype-for-word" / "SKILL.md",
    ]
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError("Missing required plugin files: " + ", ".join(missing))
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(root.rglob("*")):
            if path.is_file() and should_include(path, root) and path.resolve() != output:
                archive.write(path, Path(root.name) / path.relative_to(root))
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix(output.suffix + ".sha256").write_text(
        f"{digest}  {output.name}\n", encoding="ascii"
    )
    return digest


def main() -> int:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[1]
    output = (
        Path(sys.argv[2])
        if len(sys.argv) > 2
        else root / "dist" / "mathtype-for-word-plugin.zip"
    )
    digest = package(root, output)
    print(f"Created: {output.resolve()}")
    print(f"SHA256: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
