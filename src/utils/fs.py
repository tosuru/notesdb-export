
"""
Minimal filesystem helpers.
- ensure_dir(path): create directory if missing
- read_text(path): read UTF-8 text
- write_text(path, text): write UTF-8 text (create parent dirs)
- write_bytes(path, data): write bytes (create parent dirs)
- safe_stem(path): filename stem without extension (no directories)
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Union

PathLike = Union[str, os.PathLike]


def ensure_dir(path: PathLike) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def read_text(path: PathLike) -> str:
    return Path(path).read_text(encoding="utf-8")


def write_text(path: PathLike, text: str) -> None:
    p = Path(path)
    if p.parent:
        p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")


def write_bytes(path: PathLike, data: bytes) -> None:
    p = Path(path)
    if p.parent:
        p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)


def safe_stem(path: PathLike) -> str:
    return Path(path).stem
