
from __future__ import annotations
import os
from pathlib import Path

try:
    from utils import config  # project-wide config
except Exception:
    # fallback minimal config substitute
    class _Dummy:
        FONT_PATH = None
        FONT_PATH_CANDIDATES = []
    config = _Dummy()


def resolve_font_path() -> str | None:
    # 1) explicit config.FONT_PATH
    p = getattr(config, "FONT_PATH", None)
    if p and os.path.exists(p):
        return p
    # 2) env override
    env = os.getenv("FONT_PATH")
    if env and os.path.exists(env):
        return env
    # 3) candidates
    for c in getattr(config, "FONT_PATH_CANDIDATES", []):
        if os.path.exists(c):
            return c
    return None

# ReportLab helpers


def register_reportlab_font(font_path: str | None) -> tuple[str | None, Exception | None]:
    try:
        if not font_path:
            return None, None
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        name = Path(font_path).stem.replace("-", "").replace(" ", "")
        # Register once
        try:
            pdfmetrics.getFont(name)
            return name, None
        except Exception:
            pass
        pdfmetrics.registerFont(TTFont(name, font_path))
        return name, None
    except Exception as e:
        return None, e

# HTML helpers


def font_file_url(font_path: str | None) -> str | None:
    if not font_path or not os.path.exists(font_path):
        return None
    p = Path(font_path).resolve().as_posix()
    if os.name == "nt":
        return "file:///" + p
    return "file://" + p

# DOCX helpers (python-docx cannot embed; set family name only)


def guess_family_name(font_path: str | None, default: str = "Noto Sans JP") -> str:
    if not font_path:
        return default
    name = Path(font_path).stem
    # map some common Noto naming to readable family
    if "NotoSansCJK" in name or "NotoSansJP" in name or "NotoSans" in name:
        return "Noto Sans JP"
    return default
