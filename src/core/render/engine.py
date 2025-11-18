"""
レンダリングエンジン (v1.2 -> v1.1 仕様反映) - 共通ハンドラ対応
'core.render' パッケージのメインエントリポイント。

v1.1 仕様:
- 引数を `doc_path` (JSONファイルのPath) に一本化。
- `attachment_dir` は各レンダラーが `doc_path.parent` から内部的に解決する。
- `DocxRenderer`, `PdfRenderer` が `BaseRenderer` を継承しているかチェックを追加。
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any, Dict, Optional

# 共通基底クラスと、各エンジンクラスのインポート
# 各レンダラーのコンストラクタは (doc_path: Path) を受け取るように修正されている前提
from .common import BaseRenderer
from .engines.html import HtmlRenderer
from .engines.md import MarkdownRenderer

# DOCX/PDF レンダラーも BaseRenderer を継承したクラスとしてリファクタリングされている前提
try:
    from .engines.docx import DocxRenderer
except ImportError:
    logging.warning(
        "DocxRenderer not found or failed to import. DOCX rendering will not be available.")
    DocxRenderer = None  # プレースホルダー

try:
    from .engines.pdf import PdfRenderer
except ImportError:
    logging.warning(
        "PdfRenderer not found or failed to import. PDF rendering will not be available.")
    PdfRenderer = None  # プレースホルダー

logger = logging.getLogger(__name__)


def render(doc_path: Path, fmt: str) -> bytes | str:
    """
    正規化されたJSONファイルパスを受け取り、指定されたフォーマットでレンダリングします。
    v1.1: attachment_dir はレンダラー内部で doc_path から解決します。

    @param doc_path: normalized.json の Path オブジェクト
    @param fmt: "html", "md", "docx", "pdf" のいずれか
    @return: str (HTML/MD) または bytes (DOCX/PDF)
    @raises FileNotFoundError: doc_path が存在しないファイルの場合
    @raises ValueError: サポートされていない fmt が指定された場合
    @raises NotImplementedError: DOCX/PDF レンダラーが BaseRenderer を継承していない場合
    """
    renderer: BaseRenderer  # 型ヒント
    logger.info(
        f"Render engine dispatched for format: {fmt}, input: {doc_path}")

    if not doc_path or not doc_path.is_file():
        raise FileNotFoundError(f"Normalized JSON file not found: {doc_path}")

    # fmt に応じて適切なレンダラークラスをインスタンス化
    if fmt == 'html':
        renderer = HtmlRenderer(doc_path)
    elif fmt == 'md':
        renderer = MarkdownRenderer(doc_path)
    elif fmt == 'docx':
        if not DocxRenderer:
            logger.error("DocxRenderer is not available.")
            raise NotImplementedError("DocxRenderer is not implemented or imported correctly.")
        # DocxRenderer が BaseRenderer を継承しているか確認 (v1.1 追加)
        if not issubclass(DocxRenderer, BaseRenderer):
             raise NotImplementedError("DocxRenderer must inherit from BaseRenderer.")
        renderer = DocxRenderer(doc_path)
    elif fmt == 'pdf':
        if not PdfRenderer:
            logger.error("PdfRenderer is not available.")
            raise NotImplementedError("PdfRenderer is not implemented or imported correctly.")
        # PdfRenderer が BaseRenderer を継承しているか確認 (v1.1 追加)
        if not issubclass(PdfRenderer, BaseRenderer):
             raise NotImplementedError("PdfRenderer must inherit from BaseRenderer.")
        renderer = PdfRenderer(doc_path)
    else:
        logger.error(f"Unsupported format specified: {fmt}")
        raise ValueError(f'unsupported fmt: {fmt}')

    # レンダリング実行
    # render() メソッド内のエラーは BaseRenderer でキャッチされ、
    # get_output_on_error が呼ばれる可能性がある
    try:
        result = renderer.render()
        logger.info(f"Rendering successful for format: {fmt}, input: {doc_path}")
        return result
    except Exception as e:
        logger.error(f"Error during rendering for {fmt} from {doc_path}: {e}", exc_info=True)
        # ここで再 raise するか、エラーを示す特定の値を返すかは呼び出し元の設計による
        raise # 例としてエラーを再 raise する

# test code
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    fmt = "html"  # or "md", "docx", "pdf"
    test_doc_path = Path(r"C:\Users\A512292\py2024\notes_to_box_v4_2_0\dist\output\manifest_run\TEMP_DB\Document\その他\Doc_20251023_TEST4\Doc_20251023_TEST4.normalized.json")
    try:
        output = render(test_doc_path, fmt)
        print(f"{fmt.upper()} Output:", output)
    except Exception as e:
        print(f"Rendering failed: {e}")