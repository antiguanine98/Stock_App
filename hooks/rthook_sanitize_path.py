"""Strip third-party CRT DLLs from PATH to avoid Windows R6034 on startup.

Some installers (e.g. EaseUS Todo Backup) put msvcr90.dll on PATH. When a
PyInstaller app then loads the CRT without a matching activation context,
Windows shows Runtime Error R6034.
"""
from __future__ import annotations

import os
from pathlib import Path


_BAD_NAME_MARKERS = (
    "easeus",
    "iclient",
    "icls",
    "landesk",
)


def _is_system_dir(path: Path) -> bool:
    windir = Path(os.environ.get("WINDIR", r"C:\Windows")).resolve()
    try:
        resolved = path.resolve()
    except OSError:
        return False
    return resolved == windir or windir in resolved.parents


def _dir_looks_unsafe(path: Path) -> bool:
    name = str(path).lower()
    if any(marker in name for marker in _BAD_NAME_MARKERS):
        return True
    try:
        if (path / "msvcr90.dll").is_file() and not _is_system_dir(path):
            return True
    except OSError:
        return False
    return False


def _sanitize_path() -> None:
    raw = os.environ.get("PATH", "")
    kept: list[str] = []
    for part in raw.split(os.pathsep):
        part = part.strip()
        if not part:
            continue
        try:
            p = Path(part)
        except (OSError, ValueError):
            kept.append(part)
            continue
        if _dir_looks_unsafe(p):
            continue
        kept.append(part)
    os.environ["PATH"] = os.pathsep.join(kept)


_sanitize_path()
