# coding: utf-8
"""
core.attachments — DXL attachments extraction & JSON path updater
v1.5-dev (Feature)

- [v1.5-dev] (Feature) Support Environment Variables for Shared Icons:
  - `DXL_SHARED_ICONS_DIR`: (Optional) Specifies the *physical* directory to write shared icons.
    If unset, defaults to `(attachment_output_dir) / "icons"`.
  - `DXL_ICON_PATH_MODE`: (Optional) Controls the *relative path* written into the JSON `icon_path`.
    - `local` (default): "attachments/icons/<ext>.gif"
    - `shared`: "icons/<ext>.gif" (Assumes a shared root).
- [v1.5-dev] (Feature) `_update_runs_paths`: Add recursive walk for "section" tokens
  (title_runs, body_runs) to update paths within sections.
- [v1.5-dev] (Spec) `extract_and_save_json_paths`: Only add `icon_path` to
  `type:"file"` attachments, as per v1.5 schema (icons are no longer in metadata array).
- [v1.4.6] (Fix) Extract DXL-embedded icons (attachmentref>picture>gif) as shared icons.
- [v1.4.5] (Fix) Ensure img.src resolution by mapping alt text (stem) to image path.
"""
from __future__ import annotations
import base64
import os
import logging
import json
import copy  # For deep copying the input JSON
from pathlib import Path
from xml.etree import ElementTree as ET
from typing import Dict, Any, List, Optional, Tuple, Set
from datetime import datetime, timezone, timedelta  # Added timedelta

# --- Logger Setup ---

# --- XML サニタイズ（XML1.0 非許容制御文字の除去）---


def _sanitize_xml_text(s: str) -> str:
    # XML1.0で禁止される制御文字（0x00-0x08,0x0B-0x0C,0x0E-0x1F）を除去
    import re
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", s)


logger = logging.getLogger(__name__)

# --- Shared icon assets (very small transparent 1x1 GIF as fallback) ---
# レンダラで独自アイコンを同梱する場合は、このプレースホルダは実際には使われません。
_GIF_PLACEHOLDER = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00"
    b"\xff\xff\xff!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00"
    b"\x00\x00\x01\x00\x01\x00\x00\x02\x02L\x01\x00;"
)

ICON_DIR_NAME = "icons"  # attachments/icons/


def _strip_seq_suffix(fname: str) -> str:
    """
    '設計要件書_v4.1.4.002.md' -> '設計要件書_v4.1.4.md'
    Remove dot-number (.NNN) right before extension, if present.
    """
    p = Path(fname)
    stem, suf = p.stem, p.suffix
    import re
    m = re.match(r"^(?P<base>.+?)\.(?P<num>\d{3})$", stem)
    if m:
        return f"{m.group('base')}{suf}"
    return fname


def _allocate_saved_name(pretty: str, out_dir: Path, used: Set[str]) -> str:
    """
    Allocate a collision-free filename based on 'pretty'.
    If collides, append _2, _3, ... before extension.
    """
    base = Path(pretty)
    stem, ext = base.stem, base.suffix
    cand = pretty
    n = 2
    while cand in used or (out_dir / cand).exists():
        cand = f"{stem}_{n}{ext}"
        n += 1
    used.add(cand)
    return cand


def _collect_attachmentref_displaynames(json_doc: Dict[str, Any]) -> Dict[str, str]:
    """
    Collect attachmentref displaynames from runs; map key = stripped(displayname) -> displayname.
    Used to decide saved_name for $FILE entries.
    """
    result: Dict[str, str] = {}
    fields = json_doc.get("fields", {}) or {}

    def _walk(runs: Any):
        if not isinstance(runs, list):
            return
        for r in runs:
            if not isinstance(r, dict):
                continue
            t = r.get("t")
            if t == "attachmentref":
                name = r.get("name")
                disp = r.get("displayname") or r.get("name")
                if isinstance(name, str) and name and isinstance(disp, str) and disp:
                    result[name] = disp
            elif t == "table":
                for row in (r.get("rows") or []):
                    for cell in (row.get("cells") or []):
                        _walk(cell.get("runs") or [])
            # (B) v1.5-dev: section を再帰的に探索
            elif t == "section":
                _walk(r.get("title_runs") or [])
                _walk(r.get("body_runs") or [])

    for fd in fields.values():
        if isinstance(fd, dict) and fd.get("type") == "richtext":
            _walk(fd.get("runs") or [])
    return result


def _ensure_extension_icon(
    ext: str,
    # (A) v1.5-dev: 物理的な書き込み先
    icon_root_dir: Path,
    # (A) v1.5-dev: JSON に書き込むパスのモード
    icon_path_mode: Literal["shared", "local"],
    ext_icon_b64_map: Dict[str, str]
) -> str:
    """
    拡張子ごとの共有アイコン (e.g., <icon_root_dir>/md.gif) を用意し、
    モードに基づいた相対パス (e.g., "icons/md.gif" or "attachments/icons/md.gif") を返す。
    """
    safe_ext = (ext or "unknown").lstrip(".").lower()

    # (A) v1.5-dev: 物理的な書き込み先パス
    icon_root_dir.mkdir(parents=True, exist_ok=True)
    icon_physical_path = icon_root_dir / f"{safe_ext}.gif"

    if not icon_physical_path.exists():
        icon_data = None
        # ★ 修正 (v1.4.6): DXLから抽出したアイコンデータ (b64) を試す
        if safe_ext in ext_icon_b64_map:
            try:
                icon_data = base64.b64decode(ext_icon_b64_map[safe_ext])
                logger.debug(f"Found DXL icon data for .{safe_ext}")
            except Exception:
                logger.warning(
                    f"Failed to decode DXL icon data for .{safe_ext}")
                icon_data = None

        # データがなければプレースホルダーを使う
        if not icon_data:
            icon_data = _GIF_PLACEHOLDER
            logger.debug(f"Using placeholder icon for .{safe_ext}")

        try:
            icon_physical_path.write_bytes(icon_data)
            logger.debug(f"Wrote icon: {icon_physical_path}")
        except Exception as e:
            logger.warning(
                f"Could not create icon {icon_physical_path}: {e}")

    # (A) v1.5-dev: モードに基づいて JSON に書き込む相対パスを決定
    icon_rel_path: str
    if icon_path_mode == "shared":
        # "icons/md.gif"
        icon_rel_path = str(Path(ICON_DIR_NAME) /
                            f"{safe_ext}.gif").replace("\\", "/")
    else:  # "local" (default)
        # "attachments/icons/md.gif"
        icon_rel_path = str(Path("attachments") / ICON_DIR_NAME /
                            f"{safe_ext}.gif").replace("\\", "/")

    return icon_rel_path


# ---- Backward-compat context holders (for 3-arg callers) ----
_CTX_USED_SAVED_NAMES: Set[str] = set()
_CTX_DISPLAYNAME_MAP: Dict[str, str] = {}


# --- XML Namespace Helpers ---


def _ns_from_root(root: ET.Element) -> dict:
    """Gets the namespace map from the DXL root element."""
    if '}' in root.tag:
        # Extracts URI from Clark notation {uri}tag
        uri = root.tag.split('}', 1)[0][1:]
        return {'dxl': uri}
    logger.debug("No namespace found in DXL root tag.")
    return {}


def _q(tag: str, ns: dict) -> str:
    """Adds namespace prefix to tag name (Clark notation)."""
    return f"{{{ns['dxl']}}}{tag}" if ns else tag


def _localname(tag_with_ns: str) -> str:
    """Returns local name from '{uri}name' or 'name'."""
    return tag_with_ns.split('}')[-1]


def _safe_find(node: Optional[ET.Element], xpath: str, ns: Dict[str, str] | None = None) -> Optional[ET.Element]:
    """Safely finds an element."""
    if node is None:
        return None
    try:
        # Escape quotes within the attribute value if necessary for XPath
        if '@name="' in xpath:
            xpath = xpath.replace('"', '&quot;')  # Basic XML entity escaping
            logger.debug(f"Modified XPath for quotes: {xpath}")

        return node.find(xpath, ns or {})
    except Exception as e:
        logger.error(
            f"Error finding element with XPath '{xpath}': {e}", exc_info=False)
        return None


def _safe_findall(node: Optional[ET.Element], xpath: str, ns: Dict[str, str] | None = None) -> List[ET.Element]:
    """Safely finds all elements."""
    if node is None:
        return []
    try:
        if '@name="' in xpath:
            xpath = xpath.replace('"', '&quot;')
            logger.debug(f"Modified XPath for quotes: {xpath}")

        return node.findall(xpath, ns or {})
    except Exception as e:
        logger.error(
            f"Error finding elements with XPath '{xpath}': {e}", exc_info=False)
        return []

# --- File Handling Helpers ---


# --- Deduplication Helpers (サイズ→先頭1MB→BLAKE2b) ---
MANIFEST_NAME = ".attachments.manifest.json"
FIRST_N = 1 * 1024 * 1024  # 先頭1MB


def _blake2b_hex_bytes(data: bytes) -> str:
    import hashlib
    h = hashlib.blake2b()
    h.update(data)
    return h.hexdigest()


def _blake2b_hex_file(path: Path) -> str:
    import hashlib
    h = hashlib.blake2b()
    with path.open("rb") as rf:
        for chunk in iter(lambda: rf.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _sha256_hex_bytes(data: bytes) -> str:
    import hashlib
    h = hashlib.sha256()
    h.update(data)
    return h.hexdigest()


def _firstN_hash_bytes(data: bytes, n: int = FIRST_N) -> str:
    import hashlib
    h = hashlib.sha256()
    h.update(data[:n])
    return h.hexdigest()


def _firstN_hash_file(path: Path, n: int = FIRST_N) -> str:
    import hashlib
    h = hashlib.sha256()
    total = 0
    with path.open("rb") as rf:
        while total < n:
            need = min(1024 * 1024, n - total)
            chunk = rf.read(need)
            if not chunk:
                break
            h.update(chunk)
            total += len(chunk)
    return h.hexdigest()


def _load_manifest(out_dir: Path) -> dict:
    """Manifest loading is disabled."""
    return {}


def _save_manifest(out_dir: Path, manifest: dict) -> None:
    """Manifest saving is disabled."""
    pass


def _same_by_chain(existing: Path, size: int, first1_hash: str, full_b2b: str) -> bool:
    try:
        st = existing.stat()
    except FileNotFoundError:
        return False
    # ① サイズ一致
    if st.st_size != size:
        return False
    # ② 先頭1MBハッシュ一致
    if _firstN_hash_file(existing, FIRST_N) != first1_hash:
        return False
    # ③ 全体BLAKE2b一致
    if _blake2b_hex_file(existing) != full_b2b:
        return False
    return True


def _decide_and_maybe_write(pretty_name: str,
                            out_dir: Path,
                            *,
                            data_bytes: Optional[bytes] = None,
                            source_path: Optional[Path] = None) -> Tuple[Path, bool, dict]:
    """
    既存ファイル再利用 or 新規作成を決定して Path を返す
    Returns: (resolved_path, reused_existing, meta_dict)
    """
    logger.debug("[dedup] 開始: 出力ファイル重複判定")
    out_dir.mkdir(parents=True, exist_ok=True)
    # ファイル名をファイルシステムセーフに
    pretty_safe = _sanitize_filename(pretty_name)

    # 入力準備
    if data_bytes is None and source_path is None:
        raise ValueError("data_bytes か source_path のどちらかは必須です")
    if data_bytes is None:
        source_path = Path(source_path) if not isinstance(
            source_path, Path) else source_path
        data_bytes = source_path.read_bytes()

    size = len(data_bytes)
    first1_sha = _firstN_hash_bytes(data_bytes, FIRST_N)
    blake2b = _blake2b_hex_bytes(data_bytes)
    sha256 = _sha256_hex_bytes(data_bytes)
    meta = {"size": size, "first1_sha256": first1_sha,
            "blake2b": blake2b, "sha256": sha256}

    manifest = _load_manifest(out_dir)
    by_b2b = manifest.get("by_b2b", {})  # blake2b -> filename
    logger.debug(
        f"[dedup] 入力: size={size}, first1={first1_sha[:8]}..., b2b={blake2b[:8]}...")

    # 1) 同名が存在して同一なら再利用
    candidate_path = out_dir / pretty_safe
    if candidate_path.exists():
        logger.debug(f"[dedup] 既存: {candidate_path.name} が存在、同一性を確認します")
        if _same_by_chain(candidate_path, size, first1_sha, blake2b):
            logger.debug("[dedup] 同一判定OK -> 既存を再利用します")
            by_b2b.setdefault(blake2b, candidate_path.name)
            manifest["by_b2b"] = by_b2b
            _save_manifest(out_dir, manifest)
            return candidate_path, True, meta
        else:
            logger.debug("[dedup] 内容が異なります -> 新規名を検討します")
            # Find a new name with _2, _3, etc.
            base = Path(pretty_safe)
            stem, ext = base.stem, base.suffix
            n = 2
            while True:
                new_name = f"{stem}_{n}{ext}"
                new_path = out_dir / new_name
                if not new_path.exists():
                    logger.debug(f"[dedup] 衝突のため新規作成 -> {new_name}")
                    tmp = out_dir / (new_name + ".tmp")
                    tmp.write_bytes(data_bytes)
                    tmp.replace(new_path)
                    # マニフェスト更新
                    by_b2b[blake2b] = new_name
                    manifest["by_b2b"] = by_b2b
                    _save_manifest(out_dir, manifest)
                    return new_path, False, meta
                n += 1

    # 2) マニフェストに同一内容の別名が記録済みなら再利用
    reuse_name = by_b2b.get(blake2b)
    if reuse_name:
        reuse_path = out_dir / reuse_name
        if reuse_path.exists():
            logger.debug(f"[dedup] マニフェスト一致 -> {reuse_name} を再利用します")
            return reuse_path, True, meta

    # 3) 新規保存 (衝突なし、再利用なし)
    new_path = out_dir / pretty_safe
    logger.debug(f"[dedup] 新規作成 -> {pretty_safe}")
    tmp = out_dir / (pretty_safe + ".tmp")
    tmp.write_bytes(data_bytes)
    tmp.replace(new_path)

    # マニフェスト更新
    by_b2b[blake2b] = pretty_safe
    manifest["by_b2b"] = by_b2b
    _save_manifest(out_dir, manifest)

    return new_path, False, meta


def _sanitize_filename(filename: str, max_length: int = 200) -> str:
    """Replaces OS-forbidden characters and spaces with underscores and truncates."""
    if not filename:
        return "_no_name_"
    bad_chars = '<>:\"/\\|?* \n\r\t' + ''.join(chr(i) for i in range(32))
    safe_name = ''.join(
        '_' if c in bad_chars else c for c in filename).strip(' ._')
    if safe_name:
        parts = [part for part in safe_name.split('_') if part]
        safe_name = '_'.join(parts)

    if len(safe_name) > max_length:
        name_part, ext_part = os.path.splitext(safe_name)
        allowed_name_len = max_length - len(ext_part)
        if allowed_name_len < 1:
            safe_name = safe_name[:max_length]
        else:
            name_part = name_part[:allowed_name_len].rstrip('_')
            safe_name = name_part + ext_part
    return safe_name if safe_name else "_sanitized_"


def _set_file_timestamp(filepath: Path, att_meta: dict) -> None:
    """Sets file timestamp based on JSON metadata."""
    try:
        iso_timestamp_str = att_meta.get("modified") or att_meta.get("created")
        if not iso_timestamp_str:
            return

        dt = datetime.fromisoformat(iso_timestamp_str)
        mtime_unix = dt.timestamp()
        atime_unix = mtime_unix

        os.utime(filepath, (atime_unix, mtime_unix))
        logger.debug(
            f"Set timestamp for '{filepath.name}' to {iso_timestamp_str}")

    except ValueError:
        logger.warning(
            f"Could not parse ISO timestamp '{iso_timestamp_str}' for '{filepath.name}'.")
    except OSError as e:
        logger.warning(
            f"Failed to set timestamp for '{filepath.name}'. Error: {e}")
    except Exception as e:
        logger.warning(
            f"Unexpected error setting timestamp for '{filepath.name}' from '{iso_timestamp_str}'. Error: {e}", exc_info=False)


def _finalize_attachment_meta(att_meta: dict, attachment_output_dir: Path) -> None:
    """
    Post-process an attachment metadata dict *in place*:

    - Remove transient payload fields (bytes_b64/raw_bytes/source_path)
      to keep the final JSON slim.
    - Compute and store a SHA-256 hash of the saved file (if present) as
      `sha256` for traceability/dedup without keeping raw bytes.
    - Normalize/refresh `size` from the saved file (if present).

    This is intentionally idempotent and safe to run multiple times.
    """
    # 1) Remove transient payloads
    for k in ("bytes_b64", "raw_bytes", "source_path"):
        if k in att_meta:
            try:
                del att_meta[k]
            except Exception:
                att_meta.pop(k, None)

    # 2) Hash & size (only if a real file was saved)
    content_path = att_meta.get("content_path") or att_meta.get("saved_name")
    if not content_path:
        return

    target = attachment_output_dir / Path(str(content_path)).name
    try:
        if target.exists() and target.is_file():
            import hashlib
            h = hashlib.sha256()
            with target.open("rb") as rf:
                for chunk in iter(lambda: rf.read(8192), b""):
                    h.update(chunk)
            att_meta["sha256"] = h.hexdigest()
            try:
                att_meta["size"] = int(target.stat().st_size)
            except Exception:
                pass
    except Exception as e:
        logger.debug(
            f"_finalize_attachment_meta: skip hashing for {target}: {e}")


def extract_and_save(att_meta: Dict[str, Any],
                     attachment_output_dir: Path,
                     processed_attachment_map: Dict[str, Dict[str, Any]],
                     used_saved_names: Optional[Set[str]] = None,
                     displayname_map: Optional[Dict[str, str]] = None) -> None:
    """
    Persist a single attachment entry to disk (if not already persisted) and
    augment its metadata with content_path/saved_name. Dependencies are injected.
    Real bytes are required; will raise if none are available.

    (v1.4.4) This function supports both 3-arg and 5-arg calls for compatibility.
    """
    meta_name = att_meta.get("name")

    # --- (v1.4.4) 互換性維持ロジック ---
    if used_saved_names is None:
        used_saved_names = _CTX_USED_SAVED_NAMES
    if displayname_map is None:
        displayname_map = _CTX_DISPLAYNAME_MAP
    # --- (v1.4.4) 互換性維持ロジック終 ---

    # 1) Decide desired name: prefer displayname mapped by $FILE name, then by stripped($FILE), then fallback to $FILE@name
    preferred = None
    base_key = None
    if isinstance(meta_name, str) and meta_name:
        base_key = _strip_seq_suffix(meta_name)
        # v1.5: use $FILE name as the primary key into the map; fallback to stripped variant
        preferred = (displayname_map.get(meta_name)
                     or displayname_map.get(base_key))

    # If displayname lacks an extension, preserve original extension from meta_name
    if preferred and '.' not in Path(preferred).name and isinstance(meta_name, str) and meta_name:
        preferred = f"{preferred}{Path(meta_name).suffix}"

    # displayname が無ければ、元の name や saved_name をフォールバックとして使用
    desired = preferred or meta_name or att_meta.get(
        "saved_name") or "attachment"

    # picture: ensure .gif extension (または DXL から検出した拡張子)
    if (att_meta.get("type") == "image") and ((att_meta.get("ref") or {}).get("element") == "picture"):
        if desired and "." not in Path(desired).name:
            # (v1.4.6) _find_base64_data_node が file_ext を設定しているはず
            ext = att_meta.get("file_ext") or "gif"
            desired = f"{desired}.{ext}"
            att_meta["saved_name"] = desired  # メタデータにも反映

    # 2) Strip .NNN (v1.4) then allocate collision-free saved_name
    pretty = _strip_seq_suffix(desired)
    # _sanitize_filename を適用してファイルシステムセーフな名前にする
    pretty_safe = _sanitize_filename(pretty)

    # 2) Strip .NNN then decide name by dedup logic
    pretty = _strip_seq_suffix(desired)
    pretty_safe = _sanitize_filename(pretty)

    # 3) 書き込み用データを準備（まだ書かない）
    payload_bytes = None
    src_path = None
    b64val = att_meta.get("bytes_b64")
    if isinstance(b64val, str) and b64val:
        import base64
        try:
            payload_bytes = base64.b64decode(b64val)
        except Exception as e:
            logger.warning(
                f"base64 decode failed while preparing dedup for {pretty_safe}: {e}")
    elif isinstance(att_meta.get("raw_bytes"), (bytes, bytearray)) and len(att_meta.get("raw_bytes")) > 0:
        payload_bytes = bytes(att_meta.get("raw_bytes"))
    else:
        s = att_meta.get("source_path")
        if isinstance(s, str) and s:
            _p = Path(s)
            if _p.exists() and _p.stat().st_size > 0:
                src_path = _p
            else:
                logger.warning(f"source_path missing or 0KB for {s}")

    # 0KB は保存しない（従来仕様踏襲）
    if payload_bytes is not None and len(payload_bytes) == 0:
        payload_bytes = None

    # 4) ★ 同一性判定 → 既存再利用 or 新規（_<hash8>）
    save_path, _reused, _dig = _decide_and_maybe_write(
        pretty_name=pretty_safe,
        out_dir=attachment_output_dir,
        data_bytes=payload_bytes,
        source_path=src_path
    )

    saved_name = save_path.name

    # 既に処理済み
    if meta_name and meta_name in processed_attachment_map and save_path.exists():
        logger.debug(f"Attachment {saved_name} already processed and exists.")
        # 既存のメタ情報を更新
        att_meta.update(processed_attachment_map[meta_name])
        return

    if not save_path.exists():
        wrote = False
        # base64
        b64val = att_meta.get("bytes_b64")
        if isinstance(b64val, str) and b64val:
            import base64
            try:
                data = base64.b64decode(b64val)
                if len(data) > 0:  # 0KBファイルを禁止 (v1.4)
                    save_path.write_bytes(data)
                    wrote = True
                    logger.debug(
                        f"Wrote bytes from base64: {save_path} ({len(data)} bytes)")
                else:
                    logger.warning(
                        f"Skipped writing 0KB file from base64: {save_path.name}")
            except Exception as e:
                logger.warning(f"base64 decode failed for {save_path}: {e}")
        # raw_bytes
        if not wrote:
            raw = att_meta.get("raw_bytes")
            if isinstance(raw, (bytes, bytearray)) and len(raw) > 0:  # 0KBファイルを禁止 (v1.4)
                try:
                    save_path.write_bytes(bytes(raw))
                    wrote = True
                    logger.debug(
                        f"Wrote raw bytes: {save_path} ({len(raw)} bytes)")
                except Exception as e:
                    logger.warning(
                        f"raw_bytes write failed for {save_path}: {e}")
        # source_path
        if not wrote:
            src = att_meta.get("source_path")
            if isinstance(src, str) and src and Path(src).exists():  # 存在確認
                try:
                    if Path(src).stat().st_size > 0:  # 0KBファイルを禁止 (v1.4)
                        from shutil import copyfile
                        copyfile(src, save_path)
                        wrote = True
                        logger.debug(
                            f"Copied from source_path: {src} -> {save_path}")
                    else:
                        logger.warning(
                            f"Skipped copying 0KB source_path: {src}")
                except Exception as e:
                    logger.warning(
                        f"copy from source_path failed: {src} -> {save_path}: {e}")

        if not wrote:
            # 0KB ファイル書き込みを防止 (v1.4)
            logger.warning(
                f"No bytes found for attachment: {save_path.name}. File not created.")
            att_meta["extraction_error"] = "No non-empty data found (0KB file)."
            # content_path は設定しない
            return
            # raise RuntimeError(f"No bytes for attachment: {save_path.name}") # 厳密化

    # 4) Update metadata with relative path
    # 常に "attachments/" から始まる相対パスを生成
    att_meta["content_path"] = str(
        Path("attachments") / save_path.name).replace("\\\\", "/").replace("\\", "/")
    att_meta["saved_name"] = save_path.name

    # タイムスタンプを設定
    _set_file_timestamp(save_path, att_meta)

    # 処理済みマップに追加（循環参照防止）
    if meta_name:
        processed_attachment_map[meta_name] = copy.deepcopy(att_meta)

# --- (重複していた古い extract_and_save 定義は削除済み) ---


# --- DXL Data Extraction ---


def _find_base64_data_node(root: ET.Element, ns: dict, att_meta: dict) -> Tuple[Optional[ET.Element], Optional[str]]:
    """
    Finds the DXL node containing Base64 data for a given attachment metadata.
    Returns the data node and the determined file extension (if applicable).
    """
    ref_info = att_meta.get("ref", {})
    element_type = ref_info.get("element")
    data_node: Optional[ET.Element] = None
    # Extension determined from DXL (e.g., gif, jpeg)
    file_ext: Optional[str] = None

    try:
        if element_type == "file":
            file_name = att_meta.get("name")
            if not file_name:
                logger.warning(
                    "Attachment metadata (type 'file') is missing 'name'.")
                return None, None
            file_el = None
            file_items = root.findall(
                f".//{_q('item', ns)}[@name='$FILE']", ns)
            logger.debug(
                f"Found {len(file_items)} <item name='$FILE'> elements for searching '{file_name}'.")
            found = False
            for item_el in file_items:
                xpath_name_attr = f'@name="{file_name}"'
                if '"' in file_name and "'" not in file_name:
                    xpath_name_attr = f"@name='{file_name}'"
                elif "'" in file_name and '"' not in file_name:
                    xpath_name_attr = f'@name="{file_name}"'
                elif '"' in file_name and "'" in file_name:
                    logger.warning(
                        f"Filename '{file_name}' contains both quotes. Falling back to iteration.")
                    obj_el = _safe_find(item_el, f"./{_q('object', ns)}", ns)
                    target_container = obj_el if obj_el is not None else item_el
                    for child_file in target_container.findall(f"./{_q('file', ns)}", ns):
                        if child_file.get("name") == file_name:
                            file_el = child_file
                            found = True
                            break
                    if found:
                        break
                    continue
                if not found:
                    xpath1 = f"./{_q('object', ns)}/{_q('file', ns)}[{xpath_name_attr}]"
                    file_el = item_el.find(xpath1, ns)
                    if file_el is not None:
                        logger.debug(
                            f"Found file element for '{file_name}' using pattern 1.")
                        found = True
                    if not found:
                        xpath2 = f"./{_q('file', ns)}[{xpath_name_attr}]"
                        file_el = item_el.find(xpath2, ns)
                        if file_el is not None:
                            logger.debug(
                                f"Found file element for '{file_name}' using pattern 2.")
                            found = True
                if found:
                    break
            if file_el is None:
                logger.warning(
                    f"Could not find specific <file> element for $FILE attachment: '{file_name}'.")
                return None, None
            data_node = _safe_find(file_el, f"./{_q('filedata', ns)}", ns)
            if data_node is None:
                logger.warning(
                    f"Found <file> for '{file_name}', but missing <filedata>.")
            # $FILE の場合は file_ext は設定しない (name に拡張子が含まれる前提)

        elif element_type == "attachmentref":
            # (D) v1.5-dev: parser.py v1.5 は attachmentref icon をメタデータに含めない
            #                しかし、アイコンの Base64 抽出のために DXL を探索する必要がある (v1.4.6 互換)
            ref_name = ref_info.get("name")
            if not ref_name:
                logger.warning(
                    "Metadata (type 'attachmentref') missing 'ref.name'.")
                return None, None
            xpath_name_attr = f'@name="{ref_name}"'
            if '"' in ref_name and "'" not in ref_name:
                xpath_name_attr = f"@name='{ref_name}'"
            elif "'" in ref_name and '"' not in ref_name:
                xpath_name_attr = f'@name="{ref_name}"'
            elif '"' in ref_name and "'" in ref_name:
                logger.warning(f"Ref name '{ref_name}' contains both quotes.")
                xpath_name_attr = ""  # 属性セレクタを無効化
            else:
                xpath_name_attr = f'[{xpath_name_attr}]'

            # (D) v1.5-dev: ref_name を持つ <attachmentref> をすべて探す
            attref_candidates = root.findall(
                f".//{_q('attachmentref', ns)}{xpath_name_attr}", ns)

            if not xpath_name_attr:  # 属性セレクタが無効な場合、手動でフィルタリング
                attref_candidates = [
                    el for el in root.iterfind(f".//{_q('attachmentref', ns)}", ns) if el.get("name") == ref_name]

            if not attref_candidates:
                logger.warning(
                    f"Could not find <attachmentref> for ref name: '{ref_name}'")
                return None, None

            attref_el = attref_candidates[0]  # 最初に見つかったものを使用
            picture_el = _safe_find(attref_el, f"./{_q('picture', ns)}", ns)
            if picture_el is None:
                logger.info(
                    f"<attachmentref> for '{ref_name}' missing inner <picture>. This is OK for file stubs.")
                # アイコン実体が無い (ファイル本体のみ参照されている) 場合はデータノード無し
                return None, None

            img_tag_names = ['gif', 'jpeg', 'png', 'bmp', 'notesbitmap']
            for tag_name in img_tag_names:
                data_node = _safe_find(picture_el, f"./{_q(tag_name, ns)}", ns)
                if data_node is not None:
                    file_ext = tag_name if tag_name != 'notesbitmap' else 'bin'
                    logger.debug(
                        f"Found icon data for '{ref_name}' type: {tag_name}")
                    break
            if data_node is None:
                logger.warning(
                    f"Found <picture> icon for '{ref_name}', but no image data tag inside.")

        elif element_type == "picture":
            ref_index = ref_info.get("index")
            if ref_index is None:
                logger.warning("Inline picture metadata missing 'ref.index'.")
                return None, None

            # アイコンとして使われる <picture> を除外
            icon_pictures = set()
            for attref in root.iterfind(f".//{_q('attachmentref', ns)}", ns):
                icon_pic = _safe_find(attref, f"./{_q('picture', ns)}", ns)
                if icon_pic is not None:
                    icon_pictures.add(icon_pic)

            # 本文中のインライン <picture> を DXL 文書順に収集
            inline_pics = []
            for pic in root.iterfind(f".//{_q('picture', ns)}", ns):
                if pic not in icon_pictures:
                    inline_pics.append(pic)

            if ref_index >= len(inline_pics):
                logger.warning(
                    f"Metadata expects inline picture index {ref_index}, found {len(inline_pics)}.")
                return None, None

            picture_el = inline_pics[ref_index]
            img_tag_names = ['gif', 'jpeg', 'png', 'bmp', 'notesbitmap']
            for tag_name in img_tag_names:
                data_node = _safe_find(picture_el, f"./{_q(tag_name, ns)}", ns)
                if data_node is not None:
                    file_ext = tag_name if tag_name != 'notesbitmap' else 'bin'
                    logger.debug(
                        f"Found inline picture data (index {ref_index}) type: {tag_name}")
                    break
            if data_node is None:
                logger.warning(
                    f"Found inline <picture> index {ref_index}, but no image data tag inside.")

        else:
            logger.warning(
                f"Unsupported attachment ref.element type: '{element_type}'")

    except Exception as e:
        logger.error(
            f"Error finding data node for attachment '{att_meta.get('name')}': {e}", exc_info=True)
        return None, None

    return data_node, file_ext


# --- JSON Path Update ---

def _update_runs_paths(runs: Optional[List[Dict[str, Any]]], attachment_map: Dict[str, Dict[str, Any]]):
    """
    Recursively updates 'src' for 'img' tokens and 'content_path' for 'attachmentref' tokens.
    - attachmentref → 本体 (type:"file") の content_path を最優先で解決 (v1.4)
    - img           → inline picture (type:"image", ref.element:"picture") の content_path を src に設定
    """
    if not runs:
        return

    # 1) マップ作成
    #    - attref は "name"（元ファイル名）で解決（本体のみ）
    #    - img   は "name"/"displayname" (parser.py が生成する alt テキスト) で解決

    # ---- Pass 1: 本体ファイル (type:"file") を attref 解決用に登録 ----
    attref_path_map: Dict[str, Optional[str]] = {}
    for meta in attachment_map.values():
        # v1.4: 厳密に type:"file" のみを attref の解決対象とする
        if meta.get("type") == "file":
            meta_name = meta.get("name")  # $FILE@name
            content_path = meta.get("content_path")
            if meta_name and content_path:
                attref_path_map[meta_name] = content_path
            # v1.4: displayname (runs から収集した名前) もキーとして登録
            displayname = meta.get("displayname")
            if displayname and content_path:
                attref_path_map.setdefault(displayname, content_path)

    # ---- Pass 2: inline picture (type:"image", ref.element:"picture") を img 解決用に登録 ----
    img_path_map: Dict[str, Optional[str]] = {}
    for meta in attachment_map.values():
        ref = (meta.get("ref") or {})
        # v1.4: 厳密に ref.element:"picture" のみ (attachmentref アイコンは除外)
        if meta.get("type") == "image" and ref.get("element") == "picture":
            content_path = meta.get("content_path")
            if not content_path:
                continue
            # parser.py は "inline_image_0" のような "name" を生成する
            meta_name = meta.get("name")
            if meta_name:
                img_path_map.setdefault(meta_name, content_path)
                # ★ 修正点(v1.4.5): 拡張子を除いた名前 (parser.py の alt と一致する) もキーとして登録
                stem_name = Path(meta_name).stem
                if stem_name != meta_name:
                    img_path_map.setdefault(stem_name, content_path)

            # displayname も (もしあれば) 登録
            displayname = meta.get("displayname")
            if displayname:
                img_path_map.setdefault(displayname, content_path)

    # ---- Recursive walker (also walks nested tables) ----
    def _walk(rlist: List[Dict[str, Any]]):
        if not isinstance(rlist, list):
            return

        for r in rlist:
            if not isinstance(r, dict):
                continue

            t = r.get("t")
            if t == "img":
                # img トークンの alt は parser.py が "name" (拡張子なし) から生成する
                key_candidates = [r.get("alt"), r.get(
                    "name"), r.get("displayname")]
                key_candidates = [
                    k for k in key_candidates if isinstance(k, str) and k]
                found = None
                for k in key_candidates:
                    # ★ 修正点(v1.4.5): img_path_map は "inline_image_0" と "inline_image_0.gif" の両方を持つ
                    if k in img_path_map and img_path_map[k]:
                        found = img_path_map[k]
                        break
                if found:
                    r["src"] = str(found).replace(
                        "\\\\", "/").replace("\\", "/")
                    logger.debug(
                        f"Updated img src: key={key_candidates[0]!r} -> {r['src']}")
                else:
                    logger.warning(
                        f"Could not resolve img src for keys={key_candidates!r}")

            elif t == "attachmentref":
                # attachmentref トークンは name ($FILE@name) と displayname (DXL@displayname) を持つ
                name = r.get("name")  # $FILE@name
                displayname = r.get("displayname")  # DXL@displayname

                # v1.4: 解決キーとして displayname を name より優先 (displayname があれば)
                key = displayname if displayname else name

                if isinstance(key, str) and key:
                    path = attref_path_map.get(key)
                    if not path and key != name and name:  # displayname で見つからなければ name で再試行
                        path = attref_path_map.get(name)

                    if path:
                        # v1.4: 解決したパスは content_path に設定
                        r["content_path"] = str(path).replace(
                            "\\\\", "/").replace("\\", "/")
                        logger.debug(
                            f"Updated attachmentref content_path: {key!r} -> {r['content_path']}")
                    else:
                        logger.error(
                            f"Could not resolve attachmentref content_path for key={key!r} (name={name!r})")
                else:
                    logger.error(
                        "attachmentref without a valid 'name' or 'displayname' field.")

            elif t == "table":
                for row in (r.get("rows") or []):
                    for cell in (row.get("cells", [])):
                        _walk(cell.get("runs") or [])

            # (B) v1.5-dev: section を再帰的に探索
            elif t == "section":
                _walk(r.get("title_runs") or [])
                _walk(r.get("body_runs") or [])

    _walk(runs)


# --- Main Function ---
def extract_and_save_json_paths(dxl_path: Path,
                                initial_json_path: Path,
                                attachment_output_dir: Path) -> Optional[Dict[str, Any]]:
    """
    1) Load initial normalized JSON (from parser.py).
    2) Load DXL for Base64 data.
    3) Extract & save attachments into 'attachment_output_dir' (v1.4: shared icons).
    4) Populate attachments[*].content_path and (optionally) icon_path for files.
    5) Update runs[].attachmentref.content_path and runs[].img.src across all richtext fields.
    """
    try:
        # ★ 修正点: initial_json_path は Path オブジェクトとして open する
        with initial_json_path.open("r", encoding="utf-8") as f:
            initial_json = json.load(f)
    except FileNotFoundError:
        logger.error(f"Initial JSON file not found at: {initial_json_path}")
        return None
    except json.JSONDecodeError:
        logger.error(f"Failed to decode JSON from: {initial_json_path}")
        return None

    try:
        try:
            dxl_text = dxl_path.read_text(encoding='utf-8')
        except UnicodeDecodeError:
            dxl_text = dxl_path.read_text(
                encoding='utf-8-sig', errors='ignore')
        dxl_text = _sanitize_xml_text(dxl_text)
        try:
            root = ET.fromstring(dxl_text)
        except MemoryError as e_mem:
            logger.error(f"DXL parsing MemoryError for {dxl_path}: {e_mem}")
            return None
        except ET.ParseError as e_inner:
            logger.error(f"Failed to parse DXL XML from {dxl_path}: {e_inner}")
            return None
        ns = _ns_from_root(root)
    except FileNotFoundError:
        logger.error(f"DXL file not found at: {dxl_path}")
        return updated_json

    # Deep copy to avoid modifying original
    updated_json = copy.deepcopy(initial_json)

    # v1.4: Displayname source map from runs (key = stripped displayname)
    displayname_map = _collect_attachmentref_displaynames(updated_json)

    # Make displayname map available to 3-arg callers as well
    global _CTX_USED_SAVED_NAMES, _CTX_DISPLAYNAME_MAP
    _CTX_USED_SAVED_NAMES = set()
    _CTX_DISPLAYNAME_MAP = displayname_map

    # Map: meta_name -> updated meta dict (after extraction)
    processed_attachment_map: Dict[str, Dict[str, Any]] = {}
    # Set: a.txt, a_2.txt (for collision avoidance)
    used_saved_names: Set[str] = set()

    # (A) v1.5-dev: 共有アイコン環境変数の読み取り
    env_shared_icons_dir = os.getenv("DXL_SHARED_ICONS_DIR")
    env_icon_path_mode = os.getenv("DXL_ICON_PATH_MODE", "local").lower()

    if env_icon_path_mode not in ("shared", "local"):
        logger.warning(
            f"Invalid DXL_ICON_PATH_MODE: '{env_icon_path_mode}'. Defaulting to 'local'.")
        env_icon_path_mode = "local"

    # (A) v1.5-dev: アイコンの物理書き込み先ディレクトリを決定
    icon_root_dir: Path
    if env_shared_icons_dir:
        icon_root_dir = Path(env_shared_icons_dir)
        logger.info(
            f"Using shared icons directory (DXL_SHARED_ICONS_DIR): {icon_root_dir}")
    else:
        icon_root_dir = attachment_output_dir / ICON_DIR_NAME
        logger.info(f"Using local icons directory (default): {icon_root_dir}")

    # (A) v1.5-dev: アイコンパスモード
    icon_path_mode: Literal["shared",
                            "local"] = env_icon_path_mode  # type: ignore
    logger.info(f"Icon path mode (DXL_ICON_PATH_MODE): '{icon_path_mode}'")

    # ★ 修正 (v1.4.6): DXLから拡張子別アイコンのBase64マップを事前作成
    ext_icon_b64_map: Dict[str, str] = {}
    logger.debug("Scanning for DXL attachmentref icons...")
    # (D) v1.5-dev: parser.py v1.5 はアイコンを attachments に含めない。
    #                DXL を直接スキャンしてアイコン情報を集める必要がある。
    for attref_el in root.iterfind(f".//{_q('attachmentref', ns)}", ns):
        related_file_name = attref_el.get("name")
        if not related_file_name:
            continue

        ext = Path(related_file_name).suffix.lstrip(".").lower()
        if not ext or ext in ext_icon_b64_map:
            continue  # 拡張子がないか、既に見つけている

        # この attachmentref がアイコン (picture) を持っているか探す
        picture_el = _safe_find(attref_el, f"./{_q('picture', ns)}", ns)
        if picture_el is None:
            continue

        # アイコンの Base64 データを探す
        img_tag_names = ['gif', 'jpeg', 'png', 'bmp', 'notesbitmap']
        data_node = None
        for tag_name in img_tag_names:
            data_node = _safe_find(picture_el, f"./{_q(tag_name, ns)}", ns)
            if data_node is not None:
                break

        if data_node is not None and data_node.text:
            b64_data = data_node.text.strip()
            if b64_data:
                logger.info(
                    f"Found and mapped DXL icon for .{ext} from '{related_file_name}'.")
                ext_icon_b64_map[ext] = b64_data

    logger.debug(
        f"Found DXL icon data for extensions: {list(ext_icon_b64_map.keys())}")

    attachments_out: List[Dict[str, Any]] = []

    for att in updated_json.get("attachments", []) or []:
        meta = att  # Use the dict directly
        t = meta.get("type")
        ref = meta.get("ref", {}) or {}
        meta_name = meta.get("name")

        try:
            # 1. Find Base64 data node in DXL (inline picture / $FILE)
            # (DXL アイコンは ext_icon_b64_map で処理済みなのでここでは不要)
            if t == "file" or (t == "image" and ref.get("element") == "picture"):
                data_node, file_ext = _find_base64_data_node(root, ns, meta)
                if data_node is not None and data_node.text:
                    meta["bytes_b64"] = data_node.text.strip()
                # ★ 修正 (v1.4.6): parser.py が name に拡張子を付けない場合があるため、ここで補完
                if t == "image" and file_ext and "." not in (meta.get("name") or ""):
                    meta["name"] = f"{meta_name or 'inline_pic'}.{file_ext}"
                    meta["file_ext"] = file_ext  # extract_and_save で使うため

            # 2. Persist & set content_path/saved_name (v1.4 logic)
            if t == "file":
                extract_and_save(meta, attachment_output_dir,
                                 processed_attachment_map, used_saved_names, displayname_map)

                if meta.get("content_path"):  # Only add icon if file was saved
                    try:
                        ext = Path(meta.get("saved_name") or meta.get(
                            "name") or "").suffix.lstrip(".")
                        # (A) v1.5-dev: 修正した関数呼び出し
                        icon_rel_path = _ensure_extension_icon(
                            ext, icon_root_dir, icon_path_mode, ext_icon_b64_map)
                        # (A) v1.5-dev: 返された相対パスを設定
                        meta["icon_path"] = icon_rel_path
                    except Exception as e:
                        logger.debug(f"icon_path assignment skipped: {e}")

                attachments_out.append(meta)

            elif t == "image" and ref.get("element") == "picture":
                extract_and_save(meta, attachment_output_dir,
                                 processed_attachment_map, used_saved_names, displayname_map)
                attachments_out.append(meta)  # Keep inline pictures

            elif t == "image" and ref.get("element") == "attachmentref":
                # v1.4/v1.5: per-attachment icons are dropped
                logger.debug(
                    f"Skipping per-attachment icon entry: {meta.get('name')}")
                continue

            else:
                # OLE or future types?
                logger.warning(
                    f"Processing unknown attachment type '{t}' (ref: {ref.get('element')})")
                extract_and_save(meta, attachment_output_dir,
                                 processed_attachment_map, used_saved_names, displayname_map)
                attachments_out.append(meta)

        except Exception as e:
            logger.error(
                f"Failed to process attachment '{meta_name}': {e}", exc_info=True)
            meta["extraction_error"] = str(e)
            # Append even on error to show failure
            attachments_out.append(meta)

    # (v1.5-dev) Slim the JSON by stripping payloads and adding file hashes
    for _m in attachments_out:
        _finalize_attachment_meta(_m, attachment_output_dir)

    updated_json["attachments"] = attachments_out

    # Update all richtext runs (v1.4 logic)
    for field_data in updated_json.get("fields", {}).values():
        if isinstance(field_data, dict) and field_data.get("type") == "richtext":
            runs = field_data.get("runs")
            if runs:
                # processed_attachment_map には extract_and_save で更新された
                # { "name": ..., "content_path": ..., "displayname": ... } が入っている
                _update_runs_paths(runs, processed_attachment_map)

    logger.info(
        f"Attachment extraction and JSON path update finished for {initial_json_path.name}.")
    return updated_json


# --- Unit Test Entrypoint (optional) ---
if __name__ == "__main__":
    import tempfile
    import shutil
    import unittest

    logger.setLevel(logging.INFO)
    if logger.handlers:
        logger.handlers[0].setLevel(logging.INFO)

    class AttachmentsTest(unittest.TestCase):
        temp_output_dir: Path = None  # type: ignore
        # (A) v1.5-dev: 共有アイコン用ディレクトリ
        temp_shared_icon_dir: Path = None  # type: ignore

        @classmethod
        def setUpClass(cls):
            cls.temp_output_dir = Path(tempfile.mkdtemp(prefix="att_test_"))
            # (A) v1.5-dev: 共有アイコン用
            cls.temp_shared_icon_dir = Path(
                tempfile.mkdtemp(prefix="att_shared_icons_"))
            logger.info(
                f"Created temporary output directory: {cls.temp_output_dir}")
            logger.info(
                f"Created temporary shared icon directory: {cls.temp_shared_icon_dir}")

        @classmethod
        def tearDownClass(cls):
            if cls.temp_output_dir and cls.temp_output_dir.exists():
                try:
                    shutil.rmtree(cls.temp_output_dir)
                    logger.info(
                        f"Removed temporary output directory: {cls.temp_output_dir}")
                except Exception as e:
                    logger.error(
                        f"Error removing temp directory {cls.temp_output_dir}: {e}")
            # (A) v1.5-dev: 共有アイコン用
            if cls.temp_shared_icon_dir and cls.temp_shared_icon_dir.exists():
                try:
                    shutil.rmtree(cls.temp_shared_icon_dir)
                    logger.info(
                        f"Removed temporary shared icon directory: {cls.temp_shared_icon_dir}")
                except Exception as e:
                    logger.error(
                        f"Error removing temp shared icon directory {cls.temp_shared_icon_dir}: {e}")

        def setUp(self):
            """テストケースごとに環境変数をリセット"""
            if "DXL_SHARED_ICONS_DIR" in os.environ:
                del os.environ["DXL_SHARED_ICONS_DIR"]
            if "DXL_ICON_PATH_MODE" in os.environ:
                del os.environ["DXL_ICON_PATH_MODE"]

        def _run_test(self, mode: Literal["local", "shared_dir"]):
            """v1.5 の環境変数サポートをテストする共通ロジック"""

            case_name = f"test_v1_5_env_{mode}"
            logger.info(f"--- Running {case_name} ---")

            # Minimal synthetic initial JSON (from parser.py v1.5)
            # (D) v1.5-dev: DXL アイコン (type:"image", ref:"attachmentref") は parser.py v1.5 から削除される
            initial_json = {
                "attachments": [
                    # 1. 本体ファイル
                    {"name": "README.md", "type": "file", "size": 100,
                        "ref": {"element": "file"}, "content_path": None, "saved_name": None},
                    # 2. インライン画像
                    {"name": "inline_image_0", "type": "image", "size": 200,
                        "ref": {"element": "picture", "index": 0}, "content_path": None, "saved_name": None},
                    # 3. .002 サフィックス付きの displayname を持つファイル
                    {"name": "data.txt", "type": "file", "size": 300,
                        "ref": {"element": "file"}, "content_path": None, "saved_name": None},
                ],
                "fields": {
                    "Body": {
                        "type": "richtext",
                        "runs": [
                            {"t": "attachmentref", "name": "README.md",
                                "displayname": "README.md"},
                            {"t": "img", "alt": "inline_image_0"},
                            # (B) v1.5-dev: section をテストケースに追加
                            {
                                "t": "section",
                                "title_runs": [
                                    {"t": "attachmentref", "name": "data.txt",
                                        "displayname": "文書_v1.002.txt"}
                                ],
                                "body_runs": [
                                    {
                                        "t": "table",
                                        "rows": [
                                            {"cells": [
                                                {"runs": [
                                                    {"t": "attachmentref", "name": "README.md", "displayname": "README.md"}]}
                                            ]}
                                        ],
                                    },
                                ]
                            }
                        ],
                    }
                },
            }

            b64_md_icon = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"  # 1x1 gif
            b64_txt_icon = "R0lGODlhAQABAPAAAAAAAAAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=="
            b64_inline_gif = "R0lGODlhAQABAIAAAP8AAAAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=="

            dummy_dxl_xml = f"""<?xml version='1.0' encoding='utf-8'?>
<note xmlns='http://www.lotus.com/dxl' version='9.0' maintenanceversion='1.0'>
    <item name='$FILE'>
        <object><file name='README.md' size='100'>
            <filedata>UkVBRE1FLm1kIENvbnRlbnQ=</filedata>
        </file></object>
    </item>
    <item name='$FILE'>
        <object><file name='data.txt' size='300'>
            <filedata>ZGF0YS50eHQgQ29udGVudA==</filedata>
        </file></object>
    </item>
    <item name='Body'><richtext>
        <attachmentref name='README.md' displayname='README.md'>
            <picture><gif>{b64_md_icon}</gif></picture>
        </attachmentref>
        <picture><gif>{b64_inline_gif}</gif></picture>
        <section>
            <sectiontitle>
                <attachmentref name='data.txt' displayname='文書_v1.002.txt'>
                    <picture><gif>{b64_txt_icon}</gif></picture>
                </attachmentref>
            </sectiontitle>
        </section>
    </richtext></item>
</note>
"""
            tmp_root = self.temp_output_dir
            att_dir = tmp_root  # 本体添付ファイルの出力先
            initial_json_path = tmp_root / "initial.json"
            dxl_dummy_path = tmp_root / "dummy.xml"

            initial_json_path.write_text(json.dumps(
                initial_json, ensure_ascii=False, indent=2), encoding="utf-8")
            dxl_dummy_path.write_text(dummy_dxl_xml, encoding="utf-8")

            # --- 環境変数の設定 (A) v1.5-dev ---
            expected_icon_path_prefix: str
            expected_icon_physical_dir: Path

            if mode == "shared_dir":
                # 共有ディレクトリを指定 (パスモードは 'shared' を想定)
                os.environ["DXL_SHARED_ICONS_DIR"] = str(
                    self.temp_shared_icon_dir)
                os.environ["DXL_ICON_PATH_MODE"] = "shared"
                expected_icon_path_prefix = "icons/"  # "icons/md.gif"
                expected_icon_physical_dir = self.temp_shared_icon_dir
            else:  # "local"
                # 環境変数は設定しない (デフォルト)
                expected_icon_path_prefix = "attachments/icons/"  # "attachments/icons/md.gif"
                expected_icon_physical_dir = att_dir / ICON_DIR_NAME

            # --- 実行 ---
            updated = extract_and_save_json_paths(
                dxl_dummy_path, initial_json_path, att_dir)

            self.assertIsNotNone(
                updated, f"[{case_name}] extract returned None")

            # --- 検証 ---
            # 1. attachments 配列: file x2 + picture x1
            self.assertEqual(len(
                updated["attachments"]), 3, f"[{case_name}] Expected 3 attachments (2 file, 1 picture)")

            readme_file = next(
                a for a in updated["attachments"] if a["name"] == "README.md")
            data_file = next(
                a for a in updated["attachments"] if a["name"] == "data.txt")
            inline_img = next(
                a for a in updated["attachments"] if a["name"] == "inline_image_0.gif")

            # 2. content_path (本体) と icon_path (共有) の確認
            self.assertEqual(
                readme_file["content_path"], "attachments/README.md")
            self.assertEqual(readme_file["icon_path"],
                             f"{expected_icon_path_prefix}md.gif")

            self.assertEqual(data_file["content_path"],
                             "attachments/文書_v1.txt")  # サニタイズ＋.002除去
            self.assertEqual(data_file["icon_path"],
                             f"{expected_icon_path_prefix}txt.gif")

            self.assertEqual(inline_img["content_path"],
                             "attachments/inline_image_0.gif")
            self.assertNotIn(
                "icon_path", inline_img, f"[{case_name}] Inline image should not have icon_path")

            # 3. runs 内のパスが正しく解決されているか (v1.5 section 対応)
            runs = updated["fields"]["Body"]["runs"]
            all_runs = []

            def find_runs_recursive(rlist):
                if not isinstance(rlist, list):
                    return
                for r in rlist:
                    if not isinstance(r, dict):
                        continue
                    all_runs.append(r)
                    if r.get("t") == "table":
                        for row in r.get("rows", []):
                            for cell in row.get("cells", []):
                                find_runs_recursive(cell.get("runs", []))
                    # (B) v1.5-dev: section を再帰
                    elif r.get("t") == "section":
                        find_runs_recursive(r.get("title_runs", []))
                        find_runs_recursive(r.get("body_runs", []))

            find_runs_recursive(runs)

            attrefs = [r for r in all_runs if r.get("t") == "attachmentref"]
            imgs = [r for r in all_runs if r.get("t") == "img"]

            self.assertEqual(
                len(attrefs), 3, f"[{case_name}] Expected 3 attachmentrefs recursively")
            self.assertEqual(
                len(imgs), 1, f"[{case_name}] Expected 1 img recursively")

            self.assertEqual(imgs[0]["src"], "attachments/inline_image_0.gif")

            # attref[0] (Body直下 README.md)
            self.assertEqual(attrefs[0]["content_path"],
                             "attachments/README.md")
            # attref[1] (Section Title内 data.txt)
            self.assertEqual(attrefs[1]["content_path"],
                             "attachments/文書_v1.txt")
            # attref[2] (Section Body内 Table内 README.md)
            self.assertEqual(attrefs[2]["content_path"],
                             "attachments/README.md")

            # 4. 共有アイコンが DXL から抽出され、正しい物理パスに書き込まれているか
            icon_md_path = expected_icon_physical_dir / "md.gif"
            icon_txt_path = expected_icon_physical_dir / "txt.gif"
            icon_gif_path = expected_icon_physical_dir / "gif.gif"  # inline picture 用

            self.assertTrue(icon_md_path.exists(),
                            f"[{case_name}] md.gif not found at {icon_md_path}")
            self.assertTrue(icon_txt_path.exists(
            ), f"[{case_name}] txt.gif not found at {icon_txt_path}")
            self.assertTrue(icon_gif_path.exists(
            ), f"[{case_name}] gif.gif not found at {icon_gif_path}")

            self.assertEqual(icon_md_path.read_bytes(),
                             base64.b64decode(b64_md_icon))
            self.assertEqual(icon_txt_path.read_bytes(),
                             base64.b64decode(b64_txt_icon))
            self.assertEqual(icon_gif_path.read_bytes(), _GIF_PLACEHOLDER)

            # 5. 本体ファイルが DXL から抽出されているか (att_dir)
            readme_path = att_dir / "README.md"
            data_path = att_dir / "文書_v1.txt"
            inline_path = att_dir / "inline_image_0.gif"

            self.assertTrue(readme_path.exists())
            self.assertTrue(data_path.exists())
            self.assertTrue(inline_path.exists())

            self.assertEqual(readme_path.read_bytes(),
                             base64.b64decode("UkVBRE1FLm1kIENvbnRlbnQ="))
            self.assertEqual(data_path.read_bytes(),
                             base64.b64decode("ZGF0YS50eHQgQ29udGVudA=="))
            self.assertEqual(inline_path.read_bytes(),
                             base64.b64decode(b64_inline_gif))

        def test_mode_local(self):
            """環境変数なし (デフォルト = local) で実行"""
            self._run_test("local")

        def test_mode_shared(self):
            """環境変数あり (shared) で実行"""
            self._run_test("shared_dir")

    # TestLoader を使用して実行
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(AttachmentsTest))
    runner = unittest.TextTestRunner()
    runner.run(suite)
