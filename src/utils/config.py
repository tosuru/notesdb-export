
"""Global config flags for rendering/parse behavior (Phase 2 additions)."""
from __future__ import annotations
import os as _os

# Body richtext detail emission
BODY_EMIT_RUNS: bool = True     # include doc["fields"]["Body"]["runs"]
BODY_EMIT_TABLES: bool = True   # include doc["fields"]["Body"]["tables"]

# Attachments meta emission
ATTACHMENTS_EMIT_META: bool = True


# --- Rendering fonts (Phase 3) ---
# Preferred font file path (set via environment variable to override)
FONT_PATH: str | None = _os.getenv(
    "FONT_PATH", r"C:\Windows\Fonts\NotoSansJP-VF.ttf")

# Fallback candidates for Windows (first existing will be used)
# フォントが見つからない場合、以下の候補を上から順に探します
FONT_PATH_CANDIDATES = [
    # Noto Sans (if installed)
    r"C:\Windows\Fonts\NotoSansJP-VF.ttf",
    r"C:\Windows\Fonts\NotoSansCJKjp-Regular.otf",
    # Windows standard fonts
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\msgothic.ttc",
    r"C:\Windows\Fonts\YuGothB.ttc",
]
