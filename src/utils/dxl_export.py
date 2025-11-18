from __future__ import annotations
from typing import Optional
import logging

logger = logging.getLogger(__name__)


def export_document_as_dxl(doc) -> Optional[str]:
    """
    Exports a NotesDocument as a DXL string using CreateDXLExporter + NotesStream.
    - 大容量DXLでもメモリピークを抑え安定化
    - _olerepr_ 参照を廃止（ノイズ削減）
    """
    session = None
    exporter = None
    stream = None

    logger.debug("--- DXL EXPORT DEBUG ---")
    logger.debug(f"[dxl] Received document object. Type: {type(doc)}")

    # _olerepr_ 参照は廃止（pywin32の内部実装依存＆ノイズ）
    # logger.debug(f"[dxl] Document COM type: {getattr(doc, '_olerepr_', 'N/A')}")

    logger.debug("[dxl] Starting DXL export based on successful sample...")

    try:
        # 親セッションを取得
        session = doc.ParentDatabase.Parent
        logger.debug(f"[dxl] 1. Got parent session. Type: {type(session)}")

        exporter = session.CreateDXLExporter()
        logger.debug("[dxl] 2. Created DXLExporter object.")

        # 画像はGIFへ変換（既存設定を踏襲）
        exporter.ConvertNotesbitmapsToGIF = True

        # 出力先ストリーム（メモリ）を作成
        stream = session.CreateStream()
        logger.debug("[dxl] 3. Created Stream object.")

        # 文字コードを明示（NotesStream）
        try:
            stream.Charset = "UTF-8"
        except Exception:
            # 古い環境で属性が無い場合はスキップ
            pass

        # --- 重要 ---
        # 巨大文字列を返す Export(doc) は避け、ストリーム宛にエクスポート
        logger.debug("[dxl] 4. Calling exporter.Export(doc, stream)...")
        # 環境によりメソッドシグネチャが異なる場合があるため try/except 二段構え
        try:
            # 1st: Export(note, stream) 形式
            exporter.Export(doc, stream)
            dxl_content = None  # ここではまだ文字列にしない
        except TypeError:
            # 2nd: Export(note) しかない場合は、やむを得ず文字列で受ける
            logger.debug(
                "[dxl] Export(doc, stream) unavailable. Fallback to Export(doc) returning string.")
            dxl_content = exporter.Export(doc)

        logger.debug("[dxl] 5. exporter.Export() method completed.")

        # ストリーム経由で取得（第一候補）
        if stream is not None:
            try:
                if getattr(stream, "Bytes", 0) and stream.Bytes > 0:
                    stream.Position = 0
                    text = stream.ReadText()
                    logger.debug(
                        f"[dxl] 6a. Read {len(text)} characters from stream.")
                    return text
            except Exception as e:
                logger.debug(f"[dxl] 6a-x. Failed to read from stream: {e}")

        # 文字列フォールバック（第二候補）
        if isinstance(dxl_content, str) and dxl_content:
            logger.debug(
                f"[dxl] 7. Successfully obtained DXL content (length: {len(dxl_content)}).")
            return dxl_content

        # どちらも取れなかった場合
        logger.debug("[dxl] X. DXL export returned no content.")
        return None

    except Exception as e:
        logger.exception(f"[dxl] X. An error occurred during DXL export: {e}")
        return None
    finally:
        if stream:
            try:
                stream.Close()
            except Exception:
                pass
