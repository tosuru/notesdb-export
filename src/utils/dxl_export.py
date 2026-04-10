from __future__ import annotations
from typing import Optional
import logging
import tempfile
import os
import traceback

logger = logging.getLogger(__name__)


def export_document_as_dxl(doc, fallback_att_dir: Optional[str] = None) -> Optional[str]:
    """
    Exports a NotesDocument as a DXL string using CreateDXLExporter + NotesStream.
    - 通常動作は高速なインメモリストリームを使用します。
    - 大容量ドキュメント（巨大添付ファイル等）によるメモリクラッシュ等の例外が発生した場合、
      物理ファイルベースのストリーム出力へフォールバックし、プロセスを保護します。
    - _olerepr_ 参照を廃止（ノイズ削減）
    """
    session = None
    exporter = None
    stream = None

    logger.debug("--- DXL EXPORT DEBUG ---")
    logger.debug(f"[dxl] Received document object. Type: {type(doc)}")

    try:
        session = doc.ParentDatabase.Parent
        logger.debug(f"[dxl] 1. Got parent session. Type: {type(session)}")

        exporter = session.CreateDXLExporter()
        logger.debug("[dxl] 2. Created DXLExporter object.")

        exporter.ConvertNotesbitmapsToGIF = True

        stream = session.CreateStream()
        logger.debug("[dxl] 3. Created in-memory Stream object.")

        try:
            stream.Charset = "UTF-8"
        except Exception:
            pass

        logger.debug("[dxl] 4. Calling exporter.Export(doc)...")
        # 0th attempt: Proactive COM extraction completely independent of MemoryError.
        # This acts as a global safety net because sometimes DXL misses attachments entirely without throwing errors.
        if fallback_att_dir:
            os.makedirs(fallback_att_dir, exist_ok=True)
            logger.debug(f"[dxl] Starting PROACTIVE direct COM extraction to: {fallback_att_dir}")
            extracted_names = set()
            try:
                for item in doc.Items:
                    # 1084 = ATTACHMENT type in Notes API
                    if getattr(item, "Type", 0) == 1084:
                        values = getattr(item, "Values", [])
                        if values:
                            att_name = values[0]
                            if att_name in extracted_names:
                                continue
                            extracted_names.add(att_name)
                            try:
                                att = doc.GetAttachment(att_name)
                                if att:
                                    save_path = os.path.join(fallback_att_dir, att_name)
                                    if os.path.exists(save_path):
                                        try:
                                            os.remove(save_path)
                                        except OSError:
                                            pass
                                    att.ExtractFile(save_path)
                                    logger.debug(f"[dxl] Proactively extracted: {att_name} ({os.path.getsize(save_path)} bytes)")
                            except Exception as ext_e:
                                logger.warning(f"[dxl] Failed to proactively extract {att_name}: {ext_e}")
            except Exception as e_pre:
                logger.warning(f"[dxl] Proactive extraction failed: {e_pre}")

        # 1st attempt: In-memory export
        dxl_text = None
        
        try:
            dxl_text = exporter.Export(doc)
            # Check for Empty DXL
            if not dxl_text or not dxl_text.strip():
                raise ValueError("In-memory export returned empty string.")
                
            logger.debug(f"[dxl] 5. In-memory export succeeded. ({len(dxl_text)} bytes)")
            
            # Suggest garbage collection for very large strings
            if len(dxl_text) > 50 * 1024 * 1024:
                import gc
                gc.collect()
                
        except Exception as memory_err:
            logger.warning(f"[dxl] In-memory export failed: {memory_err}. Proceeding with OmitRichTextAttachments fallback.")
            
            # COM references that will need cleanup
            temp_db = None
            temp_doc = None
            temp_nsf_path = None
            
            try:
                import uuid
                import time
                import gc
                
                # Create a local temporary database to hold the stripped copy
                temp_dir = tempfile.gettempdir()
                temp_nsf_name = f"temp_exporter_{uuid.uuid4().hex}.nsf"
                temp_nsf_path = os.path.join(temp_dir, temp_nsf_name)
                
                logger.debug(f"[dxl] Creating local transient DB at: {temp_nsf_path}")
                db_dir = session.GetDbDirectory("")
                temp_db = db_dir.CreateDatabase(temp_nsf_path, True)
                
                logger.debug("[dxl] Creating transient copy of document for fallback export...")
                temp_doc = doc.CopyToDatabase(temp_db)
                
                # Attachments extraction logic
                if fallback_att_dir:
                    os.makedirs(fallback_att_dir, exist_ok=True)
                    logger.debug(f"[dxl] Starting direct COM extraction to: {fallback_att_dir}")
                    
                    extracted_count = 0
                    extracted_names = set()
                    for item in doc.Items:
                        # 1084 = ATTACHMENT type in Notes API
                        if getattr(item, "Type", 0) == 1084:
                            values = getattr(item, "Values", [])
                            if values:
                                att_name = values[0]
                                if att_name in extracted_names:
                                    continue
                                extracted_names.add(att_name)
                                try:
                                    att = doc.GetAttachment(att_name)
                                    if att:
                                        save_path = os.path.join(fallback_att_dir, att_name)
                                        if os.path.exists(save_path):
                                            try:
                                                os.remove(save_path)
                                            except OSError:
                                                pass
                                        att.ExtractFile(save_path)
                                        extracted_count += 1
                                        logger.debug(f"[dxl] Extracted: {att_name} ({os.path.getsize(save_path)} bytes)")
                                except Exception as ext_e:
                                    logger.warning(f"[dxl] Failed to extract {att_name}: {ext_e}")
                    
                    logger.info(f"[dxl] Extracted {extracted_count} attachments directly via COM.")

                # Strip $FILE / 1084 items from the transient doc
                items_to_remove = []
                for item in temp_doc.Items:
                    if getattr(item, "Name", "").upper() == "$FILE" or getattr(item, "Type", 0) == 1084:
                        items_to_remove.append(item)
                
                for item in items_to_remove:
                    item.Remove()
                    
                logger.debug(f"[dxl] Stripped {len(items_to_remove)} attachment items from transient doc.")

                # Export the lightweight document
                dxl_text = exporter.Export(temp_doc)
                if not dxl_text or not dxl_text.strip():
                    raise ValueError("Fallback export returned empty string.")
                    
                logger.debug(f"[dxl] Fallback export succeeded. ({len(dxl_text)} bytes)")
                
                # Cleanup huge strings
                if len(dxl_text) > 50 * 1024 * 1024:
                    gc.collect()
                    
            except Exception as fallback_err:
                logger.error(f"[dxl] Fallback export also failed: {fallback_err}")
                return None
            finally:
                # Carefully release COM references
                try:
                    del db_dir
                except Exception:
                    pass
                try:
                    del temp_doc
                except Exception:
                    pass
                try:
                    if temp_db:
                        temp_db = None
                except Exception:
                    pass
                try:
                    del temp_db
                except Exception:
                    pass
                
                # Force garbage collection to free file locks
                import gc
                gc.collect()
                time.sleep(0.5)
                
                # Delete the temp local database
                if temp_nsf_path and os.path.exists(temp_nsf_path):
                    retry = 0
                    while retry < 3:
                        try:
                            os.remove(temp_nsf_path)
                            logger.debug(f"[dxl] Local temp DB deleted securely on try {retry + 1}.")
                            break
                        except Exception as e:
                            logger.debug(f"[dxl] Could not delete temp DB {temp_nsf_path} (try {retry+1}/3): {e}")
                            time.sleep(1)
                            gc.collect()
                            retry += 1


        if dxl_text:
            return dxl_text
            
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
