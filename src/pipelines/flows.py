# coding: utf-8
"""
NotesDB パイプライン実行フロー ライブラリ (旧 notes_to_outputs.py)
main.py から呼び出される純粋な関数群を提供します。

- Unified (Notes -> DXL -> Normalize -> Attach -> Render)
- Decoupled (Export, Process)
"""
from __future__ import annotations
import os
import json
import logging
import tempfile
from pathlib import Path
import re
from typing import Dict, List, Any, Optional
from datetime import datetime
from tqdm import tqdm

# --- Internal Imports ---
# main.py が src をパスに追加することを前提とします
from utils.progress_jsonl import JsonlProgress
from notes.client import NotesClient
from utils.dxl_export import export_document_as_dxl
from core.dxl.parser import parse_dxl_document_from_string
from core.attachments import extract_and_save_json_paths
from core.render.engine import render as render_from_path

# --- Constants ---
DEFAULT_VIEW_CANDIDATES: List[str] = [
    "($All)", "AllDocuments", "All Documents",
    "すべての文書", "全ての文書", "すべてのドキュメント"
]
DEFAULT_OUTPUT_FORMATS: List[str] = ["html", "md"]  # Default formats to render

# --- Configuration Flags ---
# main.py 側で制御する想定だが、旧ロジック互換のために残す
USE_LEGACY_RENDER_INTERFACE: bool = False

# --- Logger Setup ---
# ロガー設定は main.py で行う
# 各関数は __name__ でロガーを取得する
logger = logging.getLogger(__name__)

# --- Helper Functions ---


def _ensure_safe_name(s: str | None) -> str:
    """Creates a filesystem-safe name by replacing invalid characters and spaces."""
    if not s:
        return "NoName"
    bad_chars = '<>:\"/\\|?* \n\r\t' + ''.join(chr(i) for i in range(32))
    out = ''.join('_' if c in bad_chars else c for c in s).strip(' ._')
    if out:
        parts = [part for part in out.split('_') if part]
        out = '_'.join(parts)
    out = out[:120]
    return out if out else "_sanitized_"


def _compose_outdir(base_dir: Path, db_title: str, form: str | None, cats: List[str] | None, date_yyyymmdd: str, title: str) -> Path:
    """Composes the output directory path for a document based on its metadata."""
    p = base_dir / _ensure_safe_name(db_title) / \
        _ensure_safe_name(form or "NoForm")
    for c in (cats or [])[:2]:  # Limit category depth
        p = p / _ensure_safe_name(c)
    doc_dir_name = f"Doc_{date_yyyymmdd}_{_ensure_safe_name(title or 'NoTitle')}"
    p = p / doc_dir_name
    return p


def _compose_outdir_legacy_slashsplit(base_dir: Path, db_title: str, form: str | None, cats: List[str] | None, date_yyyymmdd: str, title: str) -> Path:
    """
    Legacy compatibility: when title contained '/' or '\\' and was treated as subfolders.
    We reproduce that layout to *detect* already-created directories, but we do not create new ones here.
    Example: title "A/B/C" -> ".../Doc_{date}_A/B/C"
    """
    # ▼ 追加：念のためここでも防御
    if isinstance(title, list):
        title = " ".join(str(x) for x in title if isinstance(x, (str, int, float)) and str(x).strip())
    elif not isinstance(title, str):
        title = str(title) if title is not None else "NoTitle"

    p = base_dir / _ensure_safe_name(db_title) / \
        _ensure_safe_name(form or "NoForm")
    for c in (cats or [])[:2]:
        p = p / _ensure_safe_name(c)
    import re as _re
    parts = _re.split(r"[\\/]+", title or "NoTitle")
    head = _ensure_safe_name(parts[0]) if parts else "NoTitle"
    p = p / f"Doc_{date_yyyymmdd}_{head}"
    for tail in parts[1:]:
        if tail:
            p = p / _ensure_safe_name(tail)
    return p

# --- Rendering Wrappers ---


def _render(json_filepath: Path, fmt: str) -> bytes | str:
    """Wrapper for the rendering engine call, using the JSON file path interface."""
    if not json_filepath.is_file():
        logger.error(
            f"Cannot render: Updated JSON file not found at {json_filepath}")
        return b"" if fmt in ["pdf", "docx"] else ""
    try:
        logger.debug(
            f"Calling render engine for '{fmt}' with JSON path: {json_filepath.name}")
        # render_from_path は core.render.engine から直接インポート
        return render_from_path(doc_path=json_filepath, fmt=fmt)
    except FileNotFoundError:
        logger.error(
            f"Render engine error: JSON file disappeared? {json_filepath}")
        return b"" if fmt in ["pdf", "docx"] else ""
    except NotImplementedError as e:
        logger.error(
            f"Render engine error: Format '{fmt}' not implemented correctly. {e}")
        return b"" if fmt in ["pdf", "docx"] else ""
    except Exception as render_err:
        logger.error(
            f"Error rendering format '{fmt}' from path '{json_filepath.name}': {render_err}", exc_info=True)
        return b"" if fmt in ["pdf", "docx"] else ""


def _render_legacy(updated_doc_dict: dict, fmt: str, **kwargs) -> bytes | str:
    """Legacy wrapper accepting dict and attachment_dir."""
    logger.warning(
        "Using legacy _render function (dict + attachment_dir) - DEPRECATED.")
    try:
        attachment_dir = kwargs.get("attachment_dir")
        if attachment_dir and not isinstance(attachment_dir, Path):
            attachment_dir = Path(attachment_dir)
        logger.error(
            "Legacy render function call requires specific legacy engine.")
        return b"" if fmt in ["pdf", "docx"] else ""
    except Exception as render_err:
        logger.error(
            f"Error in legacy render for '{fmt}': {render_err}", exc_info=True)
        return b"" if fmt in ["pdf", "docx"] else ""
# --- End Rendering Wrappers ---


# --- Core Pipeline Function (Integrated) ---
def process_one(
    dxl_path: Path,  # Takes DXL file path
    out_base: Path,
    db_title: str,
    formats: List[str]
) -> Optional[Path]:
    """
    Processes a single DXL file through the full pipeline.
    (Parse -> Attach -> Render)
    """
    # core モジュールはトップレベルでインポート済み
    logger.info(f"--- Starting pipeline for DXL: {dxl_path.name} ---")
    initial_json_dict: Optional[Dict] = None
    unid: Optional[str] = None
    doc_outdir: Optional[Path] = None
    attachment_dir: Optional[Path] = None
    title: str = "NoTitle"
    date_yyyymmdd: str = "NODATE"
    updated_json_dict: Optional[Dict] = None
    json_filepath: Optional[Path] = None

    try:
        # Phase 1: Parse DXL
        logger.debug(f"Reading DXL file: {dxl_path}")
        try:
            dxl_text = dxl_path.read_text(encoding='utf-8')
        except UnicodeDecodeError:
            dxl_text = dxl_path.read_text(
                encoding='utf-8-sig', errors='ignore')
        logger.debug(f"Calling parser for DXL: {dxl_path.name}")
        initial_json_dict = parse_dxl_document_from_string(
            dxl_text, db_title=db_title)
        if not initial_json_dict:
            raise ValueError("Parser returned None.")
        unid = initial_json_dict.get("meta", {}).get("unid")
        if not unid:
            raise ValueError("Parsed JSON missing UNID.")
        logger.info(f"Phase 1 Complete: Parsed DXL for UNID: {unid}")

        # Phase 2: Determine Output Directories
        meta = initial_json_dict.get("meta", {})
        fields = initial_json_dict.get("fields", {})
        title = fields.get("Subject", {}).get("value", f"NoTitle_{unid}")
        # ▼ ここから追加：list や非文字列を安全に文字列化
        if isinstance(title, list):
            # 文字列/数値だけ連結。他の型は無視（必要なら str(x) にしても良い）
            title = " ".join(str(x) for x in title if isinstance(x, (str, int, float)) and str(x).strip())
        elif not isinstance(title, str):
            title = str(title) if title is not None else f"NoTitle_{unid}"
        if not title.strip():
            title = f"NoTitle_{unid}"

        form = meta.get("form", "NoForm")
        created_iso = meta.get("created")
        try:
            date_yyyymmdd = datetime.fromisoformat(created_iso).strftime(
                '%Y%m%d') if created_iso else "NODATE"
        except (ValueError, TypeError):
            date_yyyymmdd = "INVALIDDATE"
            logger.warning(f"Invalid date: {created_iso}")
        cats_val = fields.get("Categories", {}).get("value")
        cats = []
        if isinstance(cats_val, list):
            cats = cats_val
        elif isinstance(cats_val, str):
            cats = [c.strip() for c in cats_val.split(';') if c.strip()]
            
        # まず通常の出力ディレクトリを決定
        doc_outdir = _compose_outdir(
            out_base, db_title, form, cats, date_yyyymmdd, title
        )

        # --- Legacy folder detection（タイトルに "/" が含まれ、既存がサブフォルダ構成で作られているケースを救う）---
        legacy_outdir = _compose_outdir_legacy_slashsplit(
            out_base, db_title, form, cats, date_yyyymmdd, title or "NoTitle"
        )
        if (not doc_outdir.exists()) and legacy_outdir.exists():
            logger.warning(
                f"Detected legacy slash-split directory for UNID {unid}. Using existing legacy path: {legacy_outdir}"
            )
            doc_outdir = legacy_outdir

        attachment_dir = doc_outdir / "attachments"
        initial_json_filename = f"Doc_{date_yyyymmdd}_{_ensure_safe_name(title)}.normalized.json"
        initial_json_filepath = doc_outdir / initial_json_filename
        logger.debug(f"Output directory: {doc_outdir}")
        logger.debug(f"Attachment directory: {attachment_dir}")
        doc_outdir.mkdir(parents=True, exist_ok=True)  # Create needed dirs

        # Save the initial JSON before attachment extraction
        logger.debug(f"Saving initial JSON to: {initial_json_filepath}")
        with initial_json_filepath.open("w", encoding="utf-8") as f:
            json.dump(initial_json_dict, f, ensure_ascii=False, indent=2)

        # Verify initial JSON exists (guard against odd path issues)
        if not initial_json_filepath.exists():
            logger.warning(
                f"Initial JSON missing right after write. Rewriting: {initial_json_filepath}")
            initial_json_filepath.parent.mkdir(parents=True, exist_ok=True)
            with initial_json_filepath.open("w", encoding="utf-8") as f:
                json.dump(initial_json_dict, f, ensure_ascii=False, indent=2)

        logger.info(
            f"Phase 2 Complete: Output directory ensured: {doc_outdir}")

        # Phase 3: Extract Attachments & Update JSON Dictionary
        logger.info(
            f"Phase 3 Starting: Attachment extraction for UNID: {unid}")

        # Alternative filename compatibility for legacy slash-split filenames
        if not initial_json_filepath.exists():
            title_last = re.split(r"[\\/]+", title or "NoTitle")[-1]
            alt_name = f"{_ensure_safe_name(title_last)}.normalized.json"
            alt_path = doc_outdir / alt_name
            if alt_path.exists():
                logger.warning(
                    f"Using legacy initial JSON filename: {alt_path.name}")
                initial_json_filepath = alt_path
        updated_json_dict = extract_and_save_json_paths(
            dxl_path=dxl_path,
            initial_json_path=initial_json_filepath,
            attachment_output_dir=attachment_dir
        )
        if updated_json_dict is None:
            raise RuntimeError("Attachment extraction returned None.")
        logger.info(f"Phase 3 Complete: Attachments extracted for {unid}.")
        if updated_json_dict.get("meta", {}).get("attachment_error"):
            logger.warning(
                f"Attachment processing reported error: {updated_json_dict['meta']['attachment_error']}")
        if any(att.get("extraction_error") for att in updated_json_dict.get("attachments", [])):
            logger.warning(
                f"One or more attachments failed extraction for {unid}.")

        # Phase 4: Save Updated Normalized JSON
        json_filename = f"Doc_{date_yyyymmdd}_{_ensure_safe_name(title)}.normalized.json"
        json_filepath = doc_outdir / json_filename
        logger.debug(f"Saving updated JSON to: {json_filepath}")
        with json_filepath.open("w", encoding="utf-8") as f:
            json.dump(updated_json_dict, f, ensure_ascii=False, indent=2)
        logger.info(
            f"Phase 4 Complete: Saved updated JSON: {json_filepath.name}")

        # Phase 5: Render Each Format
        logger.info(
            f"Phase 5 Starting: Rendering for {unid} formats: {formats}")
        rendering_successful = True
        can_render_new = not USE_LEGACY_RENDER_INTERFACE and json_filepath and json_filepath.exists()
        can_render_legacy = USE_LEGACY_RENDER_INTERFACE and updated_json_dict is not None
        if not (can_render_new or can_render_legacy):
            raise RuntimeError(
                "Cannot render: Required JSON data/file unavailable.")

        for fmt in formats:
            logger.debug(f"Rendering format '{fmt}' for {unid}")
            rendered_data: bytes | str | None = None
            if USE_LEGACY_RENDER_INTERFACE:
                render_kwargs = {"attachment_dir": attachment_dir}
                rendered_data = _render_legacy(
                    updated_json_dict, fmt, **render_kwargs)
            else:  # Use new interface
                if not json_filepath or not json_filepath.exists():  # Double check file exists
                    logger.error(
                        f"Cannot render '{fmt}', JSON file missing: {json_filepath}")
                    rendering_successful = False
                    continue
                rendered_data = _render(json_filepath, fmt)

            if not rendered_data:
                logger.warning(f"Rendering '{fmt}' produced no output.")
                continue  # Don't treat empty as failure necessarily
            ext, mode, encoding = fmt, "w", "utf-8"
            if fmt == "pdf":
                ext, mode, encoding = "pdf", "wb", None
            elif fmt == "docx":
                ext, mode, encoding = "docx", "wb", None
            render_filename = f"Doc_{date_yyyymmdd}_{_ensure_safe_name(title)}.{ext}"
            render_filepath = doc_outdir / render_filename
            logger.debug(f"Writing {fmt} to {render_filepath.name}")
            with render_filepath.open(mode, encoding=encoding) as f:
                f.write(rendered_data)
            logger.info(f"Rendered '{fmt}' to: {render_filepath.name}")

        if rendering_successful:
            logger.info(f"Phase 5 Complete: Rendering finished for {unid}.")
        else:
            logger.warning(
                f"Phase 5 Complete: Rendering for {unid} had errors.")

        return doc_outdir  # Success

    except Exception as e:
        phase = "Unknown"
        if initial_json_dict is None:
            phase = "Phase 1 (Parse)"
        elif doc_outdir is None:
            phase = "Phase 2 (Path/Dir)"
        elif updated_json_dict is None:
            phase = "Phase 3 (Attach)"
        elif json_filepath is None and not USE_LEGACY_RENDER_INTERFACE:
            phase = "Phase 4 (JSON Save)"
        elif updated_json_dict is not None or (json_filepath and json_filepath.exists()):
            phase = "Phase 5 (Render)"

        logger.error(
            f"{phase} FAILED for {dxl_path.name} (UNID: {unid or 'N/A'}): {e}", exc_info=True)
        # Attempt to save intermediate JSON if available and useful
        if phase != "Phase 1 (Parse)" and initial_json_dict and doc_outdir:
            try:
                fail_json_path = doc_outdir / \
                    f"Doc_{date_yyyymmdd}_{_ensure_safe_name(title)}.normalized.FAILED_{phase.split()[1]}.json"
                save_dict = updated_json_dict if updated_json_dict else initial_json_dict
                save_dict.setdefault("meta", {})[
                    "pipeline_error"] = f"{phase} FAILED: {repr(e)}"
                with fail_json_path.open("w", encoding="utf-8") as f:
                    json.dump(save_dict, f, ensure_ascii=False, indent=2)
                logger.info(
                    f"Saved intermediate JSON upon failure: {fail_json_path.name}")
            except Exception:
                pass
        return None  # Indicate failure

# --- Runner Functions (Exported) ---


def run_unified(
    db_config: Dict, out_base: Path, progress_path: Path, formats: List[str],
    keep_dxl: bool = True, limit: Optional[int] = None, retry_errors_only=False, retry_max=3
):
    """Runs the full pipeline (Notes -> DXL -> Normalize -> Attach -> Render)."""
    # 必要なモジュールはトップレベルでインポート済み
    logger.info(
        f"Starting unified run for DB: {db_config.get('title', db_config.get('db_file'))}")
    out_base.mkdir(parents=True, exist_ok=True)
    progress = JsonlProgress(progress_path)
    client: Optional[NotesClient] = None
    temp_dxl_files: List[Path] = []
    try:
        client = NotesClient(db_config.get("server", ""),
                             db_config["db_file"]).connect()
        if client is None:
            raise ConnectionError("Failed connection.")
        view_names = db_config.get("views") or (
            [db_config["view_name"]] if db_config.get("view_name") else DEFAULT_VIEW_CANDIDATES)
        all_unids = list(client.iter_unids(view_names))
        logger.info(f"Found {len(all_unids)} UNIDs.")
        if limit:
            all_unids = all_unids[:limit]
            logger.info(f"Limiting to first {len(all_unids)}.")
        pairs = [(db_config["db_file"], u) for u in all_unids]
        
        
        # -------------------------------------------------------------------------------------
        if retry_errors_only:
            # progress.jsonl から error のみを抽出（try < retry_max）
            st = progress.snapshot()
            iterable = []
            for dbf, unid in pairs:
                rec = st.get((dbf, unid))
                if not rec:
                    continue  # 未処理は対象外（= error 再実行モードだから）
                if rec.get("status") == "error" and int(rec.get("try", 0)) < retry_max:
                    iterable.append((dbf, unid, int(rec.get("try", 0))))
            logger.info(f"Retry-errors-only mode: {len(iterable)} items to process.")
        else:
            # 通常：未処理 + 再試行可能エラー
            iterable = list(progress.remaining_unids(pairs, retry_max=retry_max))
            logger.info(f"Normal mode: {len(iterable)} items to process.")
        # -------------------------------------------------------------------------------------
        
        db_title_safe = _ensure_safe_name(
            db_config.get("title", db_config["db_file"]))
        dxl_dir = out_base / "_dxl" / db_title_safe
        if keep_dxl:
            dxl_dir.mkdir(parents=True, exist_ok=True)
            
            
        processed_count = 0
        # for db_file, unid, try_count in progress.remaining_unids(pairs, retry_max=1):
        for db_file, unid, try_count in tqdm(iterable, desc=f"Processing DB {db_config.get('title')}"):
            processed_count += 1
            dxl_path: Optional[Path] = None
            try:
                progress.append(db_file, unid, "processing",
                                try_count=try_count + 1)
                doc = client.get_document_by_unid(unid)
                if not doc:
                    raise ValueError(f"No doc {unid}")
                dxl = export_document_as_dxl(doc)
                if not dxl:
                    raise RuntimeError("Empty DXL.")
                if keep_dxl:
                    dxl_path = dxl_dir / f"{unid}.xml"
                    try:
                        dxl_path.write_text(dxl, encoding="utf-8")
                        logger.info(f"Saved DXL: {dxl_path.name}")
                    except IOError as dxl_save_err:
                        logger.error(f"DXL save error: {dxl_save_err}")
                        dxl_path = None
                else:
                    # Correct use of tempfile for text mode
                    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', prefix=f'{unid}_', delete=False, encoding='utf-8') as temp_f:
                        temp_f.write(dxl)
                        dxl_path = Path(temp_f.name)
                    temp_dxl_files.append(dxl_path)
                    logger.debug(f"Temp DXL: {dxl_path}")

                if dxl_path is None or not dxl_path.exists():
                    raise RuntimeError(f"DXL path invalid {unid}")
                # Pass the correct db_title from config
                outdir = process_one(dxl_path, out_base, db_config.get(
                    "title", db_title_safe), formats)
                if outdir:
                    progress.append(db_file, unid, "done",
                                    try_count=try_count + 1, out=str(outdir))
                else:
                    # process_one logs details
                    raise RuntimeError(f"process_one failed for {unid}.")
            except Exception as e:
                logger.error(
                    f"Error processing UNID {unid}: {e}", exc_info=False)
                logger.debug("Traceback:", exc_info=True)
                progress.append(db_file, unid, "error",
                                try_count=try_count + 1, err=repr(e))
    except (ImportError, ConnectionError, Exception) as run_err:
        logger.error(f"Unified run failed: {run_err}", exc_info=True)
    finally:
        for temp_path in temp_dxl_files:
            try:
                temp_path.unlink()
                logger.debug(f"Cleaned temp DXL: {temp_path}")
            except OSError as unlink_err:
                logger.warning(
                    f"Could not delete temp DXL {temp_path}: {unlink_err}")
        if client:
            logger.info("Notes client finished.")


def run_from_manifest(manifest_path: Path, out_base: Path, state_base: Path, formats: List[str], keep_dxl: bool = True, limit: Optional[int] = None, retry_errors_only=False):
    """Runs unified pipeline for all DBs in a manifest."""
    # 必要なモジュールはトップレベルでインポート済み

    # --- ▼▼▼ 修正箇所 ▼▼▼ ---
    # 引数が文字列(str)の場合、グローバルな Path を使って変換する
    if not isinstance(manifest_path, Path):
        manifest_path = Path(manifest_path)

    # out_base と state_base にも同様の変換処理を追加
    if not isinstance(out_base, Path):
        out_base = Path(out_base)
    if not isinstance(state_base, Path):
        state_base = Path(state_base)
    # --- ▲▲▲ 修正箇所 ▲▲▲ ---

    logger.info(f"Running pipeline from manifest: {manifest_path}")
    try:
        with manifest_path.open("r", encoding="utf-8") as f:
            manifest = json.load(f)
        if not isinstance(manifest, list):
            raise TypeError("Manifest not list.")
        out_base.mkdir(parents=True, exist_ok=True)
        state_base.mkdir(parents=True, exist_ok=True)
        for i, cfg in enumerate(manifest):
            if not isinstance(cfg, dict) or "title" not in cfg or "db_file" not in cfg:
                logger.warning(f"Invalid entry #{i+1}.")
                continue
            title_safe = _ensure_safe_name(cfg["title"])
            progress_path = state_base / title_safe / "progress.jsonl"
            logger.info(f"\n--- Processing DB #{i+1}: {cfg['title']} ---")
            run_unified(cfg, out_base, progress_path, formats,
                        keep_dxl=keep_dxl, limit=limit, retry_errors_only=retry_errors_only)
    except (FileNotFoundError, json.JSONDecodeError, TypeError) as manifest_err:
        logger.error(
            f"Manifest error '{manifest_path}': {manifest_err}", exc_info=True)
    except Exception as e:
        logger.error(f"Error running from manifest: {e}", exc_info=True)


def run_decoupled_export(db_config: Dict, dxl_out: Path):
    """Exports DXL files only."""
    # 必要なモジュールはトップレベルでインポート済み
    logger.info(
        f"Starting DXL export for DB: {db_config.get('title', db_config.get('db_file'))}")
    client: Optional[NotesClient] = None
    try:
        dxl_out.mkdir(parents=True, exist_ok=True)
        client = NotesClient(db_config.get("server", ""),
                             db_config["db_file"]).connect()
        if client is None:
            raise ConnectionError("Failed connect.")
        view_names = db_config.get("views") or (
            [db_config["view_name"]] if db_config.get("view_name") else DEFAULT_VIEW_CANDIDATES)
        all_unids = list(client.iter_unids(view_names))
        logger.info(f"Found {len(all_unids)} UNIDs.")
        exported_count = 0
        for unid in tqdm(all_unids, desc=f"Exporting DXLs from {db_config.get('title')}"):
            try:
                doc = client.get_document_by_unid(unid)
                if not doc:
                    logger.warning(f"No doc {unid}")
                    continue
                dxl = export_document_as_dxl(doc)
                if dxl:
                    dxl_path = dxl_out / f"{unid}.xml"
                    dxl_path.write_text(dxl, encoding="utf-8")
                    exported_count += 1
                    logger.debug(f"Exported: {dxl_path.name}")
                else:
                    logger.warning(f"Empty DXL {unid}")
            except Exception as export_err:
                logger.error(
                    f"DXL export error {unid}: {export_err}", exc_info=False)
        logger.info(
            f"Finished DXL export. Exported {exported_count}/{len(all_unids)} docs.")
    except (ImportError, ConnectionError, Exception) as run_err:
        logger.error(f"DXL export failed: {run_err}", exc_info=True)
    finally:
        if client:
            logger.info("Notes client finished.")


def run_decoupled_process(dxl_path: Path, out_base: Path, db_title: str, formats: List[str]):
    """ Processes a single DXL file (parse, attach, render). """
    logger.info(f"Starting decoupled process for DXL: {dxl_path.name}")
    if not dxl_path.is_file():
        logger.error(f"Input DXL invalid: {dxl_path}")
        return
    try:
        # Calls the integrated process_one function
        output_dir = process_one(
            dxl_path=dxl_path, out_base=out_base, db_title=db_title, formats=formats)
        if output_dir:
            logger.info(
                f"Successfully processed DXL. Output(s) in: {output_dir}")
        else:
            logger.error(
                f"Failed to process DXL: {dxl_path.name} (Check previous logs)")
    except Exception as e:
        logger.error(
            f"Error during decoupled process for {dxl_path.name}: {e}", exc_info=True)
