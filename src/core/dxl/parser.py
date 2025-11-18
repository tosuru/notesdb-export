# coding: utf-8
"""
DXL → normalized.json (v1.5-dev) - v1.5.1 Schema Support (Fixes)

- v1.5.1 (Fix):
  - `_style_from_generic_tag` (L361): `<run>` タグ自体に付与されたスタイル属性
    (例: `highlight='yellow'`) を解釈できるように修正。(問題点2 対応)
  - `_walk` (L780): `<run>` タグの処理を、スタイルコンテナとして PUSH/POP を行う
    ように修正。(問題点2 対応)
  - `_walk` (L793, `is_style_tag` ブロック): DXL特有の `<font .../>TEXT` パターン
    (テキストが `tail` に来る) を検出し、`tail` にスタイルが適用されるよう
    救済ロジックを追加。(問題点1 対応)

- v1.5-dev (Feature) からの継承:
  - Schema version "1.5-dev".
  - `parse_dxl_document_from_string`: `meta.schema_version` を追加。
  - `_style_from_font`: `font[name]` 属性を `a["font_family"]` として抽出。
  - `RichTextParser`:
    - `__init__`: `ns` (namespace) を受け取り保持。
    - `_parse_table`: `refwidth`, `columns`, `colspan`/`rowspan`, `style`,
      `attributes.tablabel` を v1.5 仕様で処理。
    - `_walk`: `<horizrule>` ("hr"), `<section>` ("section") を処理。
"""
from __future__ import annotations

import json
import os
import logging
import base64  # attachments メタデータ抽出に必要
from typing import Any, Dict, List, Optional, Tuple, Literal
from xml.etree import ElementTree as ET
from datetime import datetime, timezone, timedelta
from pathlib import Path  # Pathlib をインポート
import traceback  # エラー詳細表示用


# --- Type Definitions ---
Style = Dict[str, Any]  # {"s":[...], "a":{...}}
AttachmentRefElement = Literal["file", "attachmentref", "picture"]
AttachmentType = Literal["file", "image", "ole"]

# --- Constants ---
# (A) v1.5-dev: スキーマバージョン更新
SCHEMA_VERSION = "1.5-dev"  # Updated schema version

# --- Logger Setup ---
logger = logging.getLogger(__name__)

# ---------------------------
# XML / DXL helpers
# ---------------------------


def _ns_from_root(root: ET.Element) -> Dict[str, str]:
    """Extracts the DXL namespace URI from the root element."""
    if '}' in root.tag:
        uri = root.tag.split('}')[0].strip('{')
        return {'dxl': uri}
    logger.debug("No namespace found in root tag.")
    return {}


def _q(tag: str, ns: Dict[str, str]) -> str:
    """Qualifies a tag name with the namespace prefix if available."""
    return f"{{{ns['dxl']}}}{tag}" if ns else tag


def _localname(tag: str) -> str:
    """'{uri}name' or 'name' からローカル名だけを返します。"""
    return tag.split('}')[-1]


def _safe_find(node: Optional[ET.Element], xpath: str, ns: Dict[str, str] | None = None) -> Optional[ET.Element]:
    """Safely finds an element, returning None on error or if node is None."""
    if node is None:
        return None
    try:
        return node.find(xpath, ns or {})
    except SyntaxError:  # Handle potential invalid XPath
        logger.warning(f"Invalid XPath syntax: {xpath}")
        return None
    except Exception as e:
        logger.error(
            f"Error finding element with XPath '{xpath}': {e}", exc_info=False)
        return None


def _safe_findall(node: Optional[ET.Element], xpath: str, ns: Dict[str, str] | None = None) -> List[ET.Element]:
    """Safely finds all elements, returning an empty list on error or if node is None."""
    if node is None:
        return []
    try:
        return node.findall(xpath, ns or {})
    except SyntaxError:  # Handle potential invalid XPath
        logger.warning(f"Invalid XPath syntax: {xpath}")
        return []
    except Exception as e:
        logger.error(
            f"Error finding elements with XPath '{xpath}': {e}", exc_info=False)
        return []


def _parse_dxl_datetime(dt_el: Optional[ET.Element]) -> Optional[str]:
    """Domino datetime: YYYYMMDDTHHMMSS,ff[+|-]TZ -> ISO8601"""
    if dt_el is None or not dt_el.text:
        return None
    raw = dt_el.text.strip()
    if not raw:
        return None

    try:
        # Simplified timezone handling - assumes [+|-]HHMM or [+|-]HH
        base = raw
        sign = '+'
        tz_offset_str = "0000"  # Default UTC+0

        # Look for the last + or - sign, potential timezone separator
        last_plus = raw.rfind('+')
        last_minus = raw.rfind('-')
        tz_sep_index = max(last_plus, last_minus)

        if tz_sep_index > 0:  # Found a potential sign
            potential_tz = raw[tz_sep_index+1:]
            # Basic check: offset should be digits, length 2 (HH) or 4 (HHMM)
            if potential_tz.isdigit() and len(potential_tz) in [2, 4]:
                base = raw[:tz_sep_index]
                sign = raw[tz_sep_index]
                tz_offset_str = potential_tz
            # Else: The sign was likely part of the time/date itself, not a TZ offset

        # Split date and time
        if 'T' not in base:
            # Accept bare date like YYYYMMDD
            if len(base) == 8 and base.isdigit():
                # Convert to ISO date
                iso_date = f"{base[0:4]}-{base[4:6]}-{base[6:8]}"
                return iso_date
            logger.warning(
                f"Invalid DXL datetime format (missing T): '{raw}'.")
            return raw
        date_part, time_part = base.split('T', 1)

        # Handle fractional seconds
        if ',' in time_part:
            sec_part, frac = time_part.split(',', 1)
            micro = (frac + "000000")[:6]  # Normalize to 6 digits
        else:
            sec_part = time_part
            micro = "000000"

        # Check main parts length
        if len(date_part) != 8 or len(sec_part) != 6:
            logger.warning(
                f"Invalid DXL date/time format (length mismatch): '{raw}'.")
            return raw

        # Parse main datetime
        dt = datetime.strptime(
            date_part + sec_part, "%Y%m%d%H%M%S").replace(microsecond=int(micro))

        # Normalize timezone offset string HH or HHMM -> HHMM
        if len(tz_offset_str) == 2:
            tz_offset_str += "00"
        if len(tz_offset_str) != 4:  # Should be HHMM now
            logger.warning(
                f"Invalid DXL timezone offset format: '{tz_offset_str}' derived from '{raw}'. Defaulting to +0000.")
            tz_offset_str = "0000"
            sign = '+'

        hours = int(tz_offset_str[:2])
        mins = int(tz_offset_str[2:])
        offset = timedelta(hours=hours, minutes=mins)
        if sign == '-':
            offset = -offset

        # Check if offset is within valid range
        if not (-timedelta(hours=23, minutes=59) <= offset <= timedelta(hours=23, minutes=59)):
            logger.warning(
                f"DXL timezone offset out of range: '{sign}{tz_offset_str}' derived from '{raw}'. Defaulting to +0000.")
            offset = timedelta(0)

        return dt.replace(tzinfo=timezone(offset)).isoformat()

    except ValueError as e:
        logger.warning(
            f"Failed to parse DXL datetime '{raw}' due to invalid format/value. Error: {e}")
        return raw  # Return raw on parsing errors
    except Exception as e:
        logger.warning(
            f"Unexpected error parsing DXL datetime '{raw}'. Error: {e}", exc_info=False)
        return raw  # Return raw on other errors

# ---------------------------
# Style helpers
# ---------------------------


def _merge_styles(a: Style | None, b: Style | None) -> Style:
    """Merges style dictionaries a and b. b takes precedence for attributes."""
    out: Style = {}
    styles_set = set()  # Use set for efficient uniqueness check for 's'

    for src in [a, b]:
        if not isinstance(src, dict):
            continue

        # merge s[] (style flags)
        s_list = src.get('s')
        if isinstance(s_list, list):
            styles_set.update(s_list)

        # merge a{} (attributes)
        a_dict = src.get('a')
        if isinstance(a_dict, dict):
            out.setdefault('a', {})
            for k, v in a_dict.items():
                if isinstance(v, list):  # Special handling for list attributes like 'fx'
                    out['a'].setdefault(k, [])
                    # Add unique items from v to existing list
                    out['a'][k].extend(
                        item for item in v if item not in out['a'][k])
                elif v is not None:
                    out['a'][k] = v  # Overwrite/set attribute

    # Convert set back to list for 's'
    if styles_set:
        out['s'] = sorted(list(styles_set))  # Sort for consistent output

    return out


def _style_from_font(font_el: ET.Element) -> Style:
    """Extract style from <font> element."""
    s: List[str] = []
    a: Dict[str, Any] = {}
    fx: List[str] = []

    # Basic attributes
    if val := font_el.get("color"):
        a["color"] = val
    if val := font_el.get("size"):
        a["size"] = val
    # Handle bgcolor (prefer bgcolor, fallback to background/highlight)
    bgcolor = font_el.get("bgcolor") or font_el.get(
        "background") or font_el.get("highlight")
    if bgcolor:
        a["bgcolor"] = bgcolor

    # (B) v1.5-dev: font_family を追加
    if val := font_el.get("name"):
        a["font_family"] = val

    # Style flags
    styles = set((font_el.get("style", "").lower().replace(",", " ").split()))
    if "bold" in styles:
        s.append("b")
    if "italic" in styles:
        s.append("i")
    if "underline" in styles:
        s.append("u")
    if "strikethrough" in styles or "strikeout" in styles:
        s.append("s")

    # Effects
    for k in ("shadow", "emboss", "extrude"):
        if k in styles:
            fx.append(k)

    # Script (superscript/subscript)
    baseline = font_el.get("baseline") or font_el.get("position")
    is_super = "superscript" in styles or baseline in ("super", "superscript")
    is_sub = "subscript" in styles or baseline in ("sub", "subscript")
    if is_super and not is_sub:
        a["script"] = "super"
        fx.append("super")
    elif is_sub and not is_super:
        a["script"] = "sub"
        fx.append("sub")

    if fx:
        a["fx"] = list(dict.fromkeys(fx))  # Remove duplicates

    st: Style = {}
    if s:
        st["s"] = s
    if a:
        st["a"] = a
    return st


def _style_from_generic_tag(el: ET.Element) -> Style:
    """Extract style from simple tags like <b>, <i>, <sup>, <run> etc."""
    tag = _localname(el.tag)
    if tag == "font":
        return _style_from_font(el)  # Delegate to specific handler

    s: List[str] = []
    a: Dict[str, Any] = {}
    fx: List[str] = []

    # (v1.5.1) <run> タグの属性 (highlight など) をスタイルとして解釈
    if tag == "run":
        # <font> と同様に highlight/background/bgcolor をチェック
        bgcolor = el.get("bgcolor") or el.get("background") or el.get("highlight")
        if bgcolor:
            a["bgcolor"] = bgcolor
        # 他に run に付く可能性のある属性があればここに追加

    elif tag == "b":
        s.append("b")
    elif tag == "i":
        s.append("i")
    elif tag == "u":
        s.append("u")
    elif tag == "strike":
        s.append("s")
    elif tag == "sup":
        a["script"] = "super"
        fx.append("super")
    elif tag == "sub":
        a["script"] = "sub"
        fx.append("sub")
    # Add other simple tags here if needed (e.g., tt -> mono)

    if fx:
        a["fx"] = fx

    st: Style = {}
    if s:
        st["s"] = s
    if a:
        st["a"] = a
    return st


def _collect_inline_style(el: ET.Element) -> Style:
    """Collect styles from el and all relevant descendants."""
    agg: Style = _style_from_generic_tag(el)  # Style from element itself
    for n in el.iter():
        if n == el:
            continue
        # Only consider style tags for descendant styles
        if _localname(n.tag) in ("font", "b", "i", "u", "strike", "sup", "sub"):
            agg = _merge_styles(agg, _style_from_generic_tag(n))
    return agg

# ---------------------------
# Paragraph handling
# ---------------------------


def _collect_pardefs(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    """Collects all <pardef> attributes, keyed by id."""
    ns = _ns_from_root(root)
    m: Dict[str, Dict[str, Any]] = {}

    for pd in _safe_findall(root, f".//{_q('pardef', ns)}", ns):
        pid = pd.get("id")
        if not pid:
            continue

        a: Dict[str, Any] = {}
        # align
        align = pd.get("align")
        if align == "full":
            a["align"] = "justify"
        elif align in ("center", "right"):
            a["align"] = align
        # margins/spacing
        if val := pd.get("leftmargin"):
            a["leftmargin"] = val
        if val := pd.get("spaceafter"):
            a["spaceafter"] = val
        # parstyle name
        ps_el = _safe_find(pd, f"./{_q('parstyle', ns)}", ns)
        if ps_el is not None and (name := ps_el.get("name")):
            a["parstyle"] = name
        elif name := pd.get("name"):
            a["parstyle"] = name  # Fallback
        # list
        if list_type := pd.get("list"):
            # Map common Notes list types to more standard names
            list_map = {
                "bullet": "bullet",
                "number": "number",
                "uncheck": "uncheck",
                "square": "square",
                "alphaupper": "alphaupper",
                "alphalower": "alphalower",
                "romanupper": "romanupper",
                "romanlower": "romanlower",
                # Add other mappings if needed
            }
            # Use mapped type if available, otherwise raw
            mapped_type = list_map.get(list_type.lower(), list_type)
            a["list"] = {"type": mapped_type, "raw": list_type}

        m[pid] = a
    return m


def _par_attrs_equal(a: Dict[str, Any] | None, b: Dict[str, Any] | None) -> bool:
    """Compares two paragraph attribute dictionaries for equality."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        a_filtered = {k: v for k,
                      v in a.items() if v not in (None, "", [], {})}
        b_filtered = {k: v for k,
                      v in b.items() if v not in (None, "", [], {})}
        return json.dumps(a_filtered, sort_keys=True) == json.dumps(b_filtered, sort_keys=True)
    except Exception:
        return False

# ---------------------------
# Attachment Metadata Extraction
# ---------------------------


def _extract_attachments_metadata(root: ET.Element, ns: Dict[str, str]) -> List[Dict[str, Any]]:
    """Extracts metadata for files ($FILE), icons (in attachmentref), and inline images."""
    attachments: List[Dict[str, Any]] = []
    processed_file_names = set()
    inline_image_index = 0
    # name -> {attref_el: ..., picture_el: ...}
    attachment_refs_info: Dict[str, Dict[str, Any | ET.Element | None]] = {}
    icon_picture_elements = set()  # アイコンとして処理するpicture要素を記憶

    # 1. Extract from <item name='$FILE'>/<object>/<file>
    for item_el in _safe_findall(root, f".//{_q('item', ns)}[@name='$FILE']", ns):
        for file_el in _safe_findall(item_el, f".//{_q('file', ns)}", ns):
            name = file_el.get("name")
            if not name or name in processed_file_names:
                continue

            meta: Dict[str, Any] = {
                "name": name,
                "type": "file",
                "ref": {"element": "file"},
                "size": int(file_el.get("size", 0)),
                "created": _parse_dxl_datetime(_safe_find(file_el, f"./{_q('created', ns)}/{_q('datetime', ns)}", ns)),
                "modified": _parse_dxl_datetime(_safe_find(file_el, f"./{_q('modified', ns)}/{_q('datetime', ns)}", ns)),
                "content_path": None,  # Initialized to None (v1.4.4)
                "saved_name": None,  # Initialized to None (v1.4.4)
            }
            attachments.append(meta)
            processed_file_names.add(name)
            logger.debug(f"Found $FILE metadata: {name}")

    # 2. Collect attachmentref elements and their picture icons first
    for attref_el in root.iterfind(f".//{_q('attachmentref', ns)}", ns):
        att_name = attref_el.get("name")
        if not att_name:
            logger.warning("Found <attachmentref> without 'name' attribute.")
            continue
        # attachmentref に紐づく picture (アイコン) を探す
        picture_el = _safe_find(attref_el, f"./{_q('picture', ns)}", ns)
        attachment_refs_info[att_name] = {
            "attref_el": attref_el, "picture_el": picture_el}
        if picture_el is not None:
            icon_picture_elements.add(picture_el)  # このpictureはアイコン用としてマーク

    # 3. Process all <picture> elements, distinguishing icons from inline images
    all_pictures = list(root.iterfind(f".//{_q('picture', ns)}", ns))

    for pic_el in all_pictures:
        if pic_el in icon_picture_elements:
            # This is an icon picture, already marked. Generate metadata for the attachmentref icon.
            # Find the corresponding attachmentref info
            found_att_name = None
            attref_el = None
            for name, info in attachment_refs_info.items():
                if info["picture_el"] == pic_el:
                    found_att_name = name
                    attref_el = info["attref_el"]
                    break

            if found_att_name and attref_el is not None:
                display_name = attref_el.get("displayname") or found_att_name
                width_str = pic_el.get("width", "0px").replace("px", "")
                height_str = pic_el.get("height", "0px").replace("px", "")
                try:
                    # Estimate size from base64 data if available, fallback to dimensions
                    gif_el = _safe_find(pic_el, f".//{_q('gif', ns)}", ns)
                    if gif_el is not None and gif_el.text:
                        # Decode to estimate byte size
                        est_size = len(base64.b64decode(gif_el.text.strip()))
                    else:
                        est_size = int(width_str) * int(height_str)
                except Exception:
                    est_size = 0

                meta_icon: Dict[str, Any] = {
                    "name": display_name,
                    "type": "image",
                    "ref": {"element": "attachmentref", "name": found_att_name},
                    "size": est_size,
                    "created": None,
                    "modified": None,
                    "content_path": None,  # Initialized to None (v1.4.4)
                    "saved_name": None,  # Initialized to None (v1.4.4)
                }
                # (D) v1.5-dev (v1.4.4) 仕様: アイコン項目は廃止
                # attachments.append(meta_icon)
                logger.debug(
                    f"Skipping attachmentref icon metadata (v1.5 spec): {display_name} (refs file: {found_att_name})")
            else:
                logger.warning(
                    "Icon picture element found but couldn't link back to attachmentref info.")

        else:
            # This is an inline <picture>
            img_name = f"inline_image_{inline_image_index}"
            width_str = pic_el.get("width", "0px").replace("px", "")
            height_str = pic_el.get("height", "0px").replace("px", "")
            try:
                # Estimate size from base64 data if available
                gif_el = _safe_find(pic_el, f".//{_q('gif', ns)}", ns)
                if gif_el is not None and gif_el.text:
                    est_size = len(base64.b64decode(gif_el.text.strip()))
                else:
                    est_size = int(width_str) * int(height_str)
            except Exception:
                est_size = 0

            meta_inline: Dict[str, Any] = {
                "name": img_name,
                "type": "image",
                "ref": {"element": "picture", "index": inline_image_index},
                "size": est_size,
                "created": None,
                "modified": None,
                "content_path": None,  # Initialized to None (v1.4.4)
                "saved_name": None,  # Initialized to None (v1.4.4)
            }
            attachments.append(meta_inline)
            logger.debug(
                f"Found inline image metadata: {img_name} (index: {inline_image_index})")
            inline_image_index += 1

    # 4. (Optional) Add metadata for attachmentrefs that had NO icon and were not in $FILE
    for att_name, info in attachment_refs_info.items():
        attref_el = info["attref_el"]
        # Check if this file was already added (either via $FILE or via its icon)
        already_added = any(a.get("ref", {}).get("name") == att_name or a.get(
            "name") == att_name for a in attachments)
        if not already_added:
            display_name = attref_el.get("displayname") or att_name
            logger.info(
                f"Adding stub metadata for attachmentref without icon/FILE entry: {display_name}")
            meta_stub: Dict[str, Any] = {
                "name": display_name,
                "type": "file",  # Assume file, could be OLE?
                "ref": {"element": "attachmentref", "name": att_name},
                "size": 0,  # Size unknown
                "created": None,
                "modified": None,
                "content_path": None,  # Initialized to None (v1.4.4)
                "saved_name": None,  # Initialized to None (v1.4.4)
            }
            attachments.append(meta_stub)

    logger.info(
        f"Extracted metadata for {len(attachments)} attachments/images.")
    attachments.sort(key=lambda x: (x.get("name", ""), x.get("type", "")))
    return attachments

# ---------------------------
# Richtext Parser
# ---------------------------


class RichTextParser:
    def __init__(self, pardefs: Dict[str, Dict[str, Any]], attachment_metas: List[Dict[str, Any]], ns: Dict[str, str]):
        self._pardefs = pardefs
        self._attachment_metas = attachment_metas
        # (B) v1.5-dev: ns を保持
        self._ns = ns
        self._inline_images_meta = [m for m in attachment_metas if m.get(
            "ref", {}).get("element") == "picture"]
        self._runs: List[Dict[str, Any]] = []
        self._plain: List[str] = []
        self._style_stack: List[Style] = []
        self._last_par_attrs: Optional[Dict[str, Any]] = None
        # [修正] None ではなく空辞書で初期化
        self._pending_par_attrs: Optional[Dict[str, Any]] = {}
        self._emitted_anything = False
        self._inline_image_index = 0
        self._emitted_content_since_last_par = False  # [追加] par 発行後にコンテントがあったか
        logger.debug("RichTextParser initialized.")

    def _cur(self) -> Style:
        """Aggregates the current style stack."""
        cur: Style = {}
        for st in self._style_stack:
            cur = _merge_styles(cur, st)
        return cur

    def _ensure_par_before_content(self):
        """
        [v1.4.5 改行対応]
        コンテンツ(text, img等)の前に par が一度も発行されていない場合、
        デフォルトの par を発行するフォールバック。
        """
        # _pending_par_attrs が None でない場合 (par が保留中か、一度も発行されていない場合)
        if self._pending_par_attrs is not None:
            attrs = self._pending_par_attrs
            logger.debug(f"Ensure par: Emitting pending/default par: {attrs}")
            # _emit_par を呼び出す（重複チェックもここで実施される）
            self._emit_par(attrs)
            # _emit_par が _pending_par_attrs を None に設定する

        # コンテントが発行されることを記録
        self._emitted_content_since_last_par = True

    def _emit_par(self, attrs: Dict[str, Any]):
        """
        [v1.4.5 改行対応 新設]
        par トークンを発行します。
        連続する同じ属性の par も（改行として）発行します。
        ただし、直前の run が既に同じ属性の par で、かつ内容が空だった場合は重複発行を防ぎます。
        """

        # 属性をフィルタリング
        attrs_filtered = {
            k: v for k, v in attrs.items() if v not in (None, "", [], {})}

        # --- 重複空 par の防止 ---
        if self._runs:
            last_run = self._runs[-1]
            if last_run.get("t") == "par":
                # 直前の par と属性を比較
                last_attrs = last_run.get("a", {})
                if last_attrs == attrs_filtered:
                    # 直前が par で属性も同じ場合、
                    # 間に content (text, img等) が無ければ発行しない
                    # _emitted_content_since_last_par をチェック
                    if not self._emitted_content_since_last_par:
                        logger.debug(
                            f"Suppressing duplicate empty par token: {attrs_filtered}")
                        self._pending_par_attrs = attrs  # 保留状態は維持
                        return
        # ------------------------

        if self._emitted_anything and not (self._plain and self._plain[-1].endswith("\n")):
            self._plain.append("\n")

        token = {"t": "par"}
        if attrs_filtered:
            token["a"] = attrs_filtered

        logger.debug(f"Emitting par token: {attrs_filtered}")
        self._runs.append(token)
        self._last_par_attrs = attrs  # (ensure_par..のフォールバック用に残す)
        self._pending_par_attrs = None  # 保留をクリア (発行したので)
        self._emitted_anything = True
        self._emitted_content_since_last_par = False  # par を発行したのでリセット

    def _emit_text(self, text: str):
        # 1) 前後空白の規格化（空文字は無視）
        text = (text or "").strip()
        if not text:
            return

        # 2) 段落ヘッダの確実化
        self._ensure_par_before_content()

        # 3) 現在のスタイルを正規化して比較可能にする
        def _norm_style(style: dict) -> dict:
            if not style:
                return {}
            out = {}
            s = style.get("s") or []
            a = style.get("a") or {}
            # s は順序非依存：集合→昇順リスト
            out["s"] = sorted(set(s))
            # a はキー順序で固定化
            out["a"] = {k: a[k] for k in sorted(a.keys())}
            return out

        cur_st = _norm_style(self._cur())

        # 4) 直前 run が text なら、スタイルで判定してマージ／抑止
        if self._runs:
            last = self._runs[-1]
            if isinstance(last, dict) and last.get("t") == "text":
                last_st = _norm_style({"s": last.get("s"), "a": last.get("a")})
                # 4-1) スタイル完全一致 → 文字列を連結して run を1つに保つ
                if last_st == cur_st:
                    last["text"] = f"{last.get('text', '')}{text}"
                    logger.debug(
                        "Merged adjacent text run with identical style.")
                    # plain も重複追加しない（最後の要素に連結）
                    if self._plain:
                        self._plain[-1] = f"{self._plain[-1]}{text}"
                    else:
                        self._plain.append(text)
                    self._emitted_anything = True
                    self._emitted_content_since_last_par = True  # テキスト発行を記録
                    # [修正] v1.4.6 推奨修正: return のコメントアウトを解除
                    return

                # --- [修正] ---
                # 4-2) スタイルが異なる場合は、テキスト内容が同じでもマージしない方針に変更。
                #      以前の複雑なマージロジック (同一テキストでスタイル違いの場合) を削除。
                #      これにより、par境界とテキストの関係性がより明確になる。
                # ----------------
                # (削除された elif ブロックはここにありました)

        # 5) ここまで来たら新規に text run を発行
        tok = {"t": "text", "text": text}
        if cur_st.get("s"):
            tok["s"] = cur_st["s"]
        if cur_st.get("a"):
            tok["a"] = cur_st["a"]

        logger.debug(f"Emitting text: '{text}' with style {cur_st}")
        self._runs.append(tok)
        self._plain.append(text)
        self._emitted_anything = True
        self._emitted_content_since_last_par = True  # テキスト発行を記録

    def _emit_link(self, href: str, label: str, st: Style, notes: Optional[Dict[str, Any]] = None):
        self._ensure_par_before_content()
        label = label.strip()
        href = href.strip()
        if not label:
            label = href or "Notes Link"  # Provide fallback label

        tok = {"t": "link", "href": href, "label": label}
        current_st = self._cur()
        merged_st = _merge_styles(current_st, st)
        if s_list := merged_st.get("s"):
            tok["s"] = s_list
        if a_dict := merged_st.get("a"):
            tok["a"] = a_dict

        tok.setdefault("s", [])
        if "u" not in tok["s"]:
            tok["s"].append("u")

        if notes:
            tok["notes"] = notes
        logger.debug(f"Emitting link: '{label}' -> '{href}'")
        self._runs.append(tok)
        self._plain.append(label)
        self._emitted_anything = True
        self._emitted_content_since_last_par = True  # [修正] リンクもコンテントとして扱う

    def _emit_attachmentref(self, name: str, displayname: Optional[str]):
        self._ensure_par_before_content()
        effective_displayname = (displayname or name or "").strip()
        tok: Dict[str, Any] = {
            "t": "attachmentref",
            "name": name,
            "displayname": effective_displayname
        }
        # No content_path here (v1.4.4 change)
        st = self._cur()
        if s_list := st.get("s"):
            tok["s"] = s_list
        if a_dict := st.get("a"):
            tok["a"] = a_dict

        logger.debug(
            f"Emitting attachmentref token: name='{name}', display='{effective_displayname}'")
        self._runs.append(tok)
        self._plain.append(f"[{effective_displayname}]")
        self._emitted_anything = True
        self._emitted_content_since_last_par = True  # [修正] attref もコンテントとして扱う

    def _emit_inline_image(self):
        self._ensure_par_before_content()
        if self._inline_image_index < len(self._inline_images_meta):
            meta = self._inline_images_meta[self._inline_image_index]
            alt_text = meta.get(
                "name", f"inline_image_{self._inline_image_index}")
            tok: Dict[str, Any] = {"t": "img", "alt": alt_text}
            # No src here (v1.4.4 change)

            logger.debug(
                f"Emitting img token (inline): alt='{alt_text}' (index {self._inline_image_index})")
            self._runs.append(tok)
            self._plain.append(f"[Image: {alt_text}]")
            self._emitted_anything = True
            self._inline_image_index += 1
            self._emitted_content_since_last_par = True  # [修正] img もコンテントとして扱う
        else:
            logger.error(
                f"Attempted to emit inline image token index {self._inline_image_index}, but only {len(self._inline_images_meta)} inline metadata found.")

    def _parse_table(self, tbl: ET.Element) -> Dict[str, Any]:
        """Parses a <table> element into a JSON token."""
        tag = _localname(tbl.tag)
        assert tag == "table"

        table_attrs: Dict[str, Any] = {
            k: v for k, v in tbl.attrib.items() if v is not None}
        rows_out: List[Dict[str, Any]] = []
        # (B) v1.5-dev: 保持している ns を使用
        ns = self._ns

        # (C-1) v1.5-dev: <tablecolumn> の情報を収集
        columns_out: List[Dict[str, Any]] = []
        for col_el in _safe_findall(tbl, f"./{_q('tablecolumn', ns)}", ns):
            col_data: Dict[str, Any] = {}
            if width := col_el.get("width"):
                col_data["width"] = width
            # 他に必要な列属性があればここに追加
            columns_out.append(col_data)

        for row_el in _safe_findall(tbl, f"./{_q('tablerow', ns)}", ns):
            row_obj: Dict[str, Any] = {}
            row_attrs: Dict[str, Any] = {
                k: v for k, v in row_el.attrib.items() if v is not None}
            
            # (C-2) v1.5-dev: tablabel を attributes オブジェクト内に格納 (直下キー 'label' を廃止)
            if row_attrs:
                row_obj["attributes"] = row_attrs  # tablabel があればここに含まれる

            cells_out: List[Dict[str, Any]] = []
            for cell_el in _safe_findall(row_el, f"./{_q('tablecell', ns)}", ns):
                cell_data: Dict[str, Any] = {}
                cell_style_attrs: Dict[str, Any] = {}  # bgcolor などの style 属性用

                # (C-4) v1.5-dev: colspan/rowspan をセル直下のキーに設定
                if val := cell_el.get("colspan"):
                    try:
                        cell_data["colspan"] = int(val)
                    except ValueError:
                        logger.warning(f"Invalid colspan value: {val}")
                if val := cell_el.get("rowspan"):
                    try:
                        cell_data["rowspan"] = int(val)
                    except ValueError:
                        logger.warning(f"Invalid rowspan value: {val}")

                # (C-4) v1.5-dev: style 属性 (bgcolor など) を収集
                for k, v in cell_el.attrib.items():
                    if k not in ("colspan", "rowspan") and v is not None:
                        cell_style_attrs[k] = v

                # (B) v1.5-dev: ns を渡す
                sub = RichTextParser(self._pardefs, self._attachment_metas, self._ns)
                sub._inline_image_index = self._inline_image_index
                logger.debug("--- Parsing TableCell ---")
                if cell_el.text:
                    # Process text before children
                    sub._emit_text(cell_el.text)
                for ch in list(cell_el):
                    sub._walk(ch)
                    if ch.tail:
                        sub._emit_text(ch.tail)  # Process tail after child
                logger.debug(f"--- End TableCell (Runs: {len(sub._runs)}) ---")

                self._inline_image_index = sub._inline_image_index  # Update index
                if cell_plain := "".join(sub._plain).rstrip():
                    self._plain.append(cell_plain + " ")

                cell_data["runs"] = sub._runs
                
                # (C-4) v1.5-dev: style 属性があれば "style" キーに格納
                if cell_style_attrs:
                    cell_data["style"] = cell_style_attrs
                    
                cells_out.append(cell_data)

            row_obj["cells"] = cells_out
            rows_out.append(row_obj)

        tbl_tok = {"t": "table", "rows": rows_out}
        if table_attrs:
            # (C-1) v1.5-dev: attributes はそのまま (refwidth も含まれる)
            tbl_tok["attributes"] = table_attrs
            
        # (C-1) v1.5-dev: columns を追加
        if columns_out:
            tbl_tok["columns"] = columns_out
            
        return tbl_tok

    def parse(self, richtext_el: ET.Element) -> Dict[str, Any]:
        """Parses the content of a <richtext> element."""
        logger.debug("--- Starting Richtext Parse ---")
        if richtext_el.text:
            self._emit_text(richtext_el.text)  # Process text before children
        for ch in list(richtext_el):
            self._walk(ch)
            if ch.tail:
                self._emit_text(ch.tail)  # Process tail after child
        logger.debug(
            f"--- Finished Richtext Parse (Runs: {len(self._runs)}) ---")
        return {"text": "".join(self._plain).rstrip(), "runs": self._runs}

    def _walk(self, el: ET.Element):
        """Main recursive walker for DXL elements within richtext."""
        tag = _localname(el.tag)
        logger.debug(f"Walking <{tag}>...")

        # --- Structure Handlers ---
        if tag == "par":
            attrs = self._pardefs.get(el.get("def", ""), {})

            # [v1.4.5 改行対応]
            # _pending_par_attrs にセットするのではなく、即座に発行を試みる
            self._emit_par(attrs)

            # par タグ内のテキストや子要素の処理
            if el.text:
                self._emit_text(el.text)
            for ch in list(el):
                self._walk(ch)
                if ch.tail:
                    self._emit_text(ch.tail)

            # par タグの処理が終わったので、次のコンテント発行は新しい par の後であることを記録
            # (もし子が何もなければ _emitted_content_since_last_par は False のまま)
            return

        if tag == "table":
            self._ensure_par_before_content()
            self._runs.append(self._parse_table(el))
            self._emitted_anything = True
            self._emitted_content_since_last_par = True  # [修正] テーブルもコンテントとして扱う
            return

        # (B) v1.5-dev: <horizrule> (水平線) のサポート
        if tag == "horizrule":
            self._ensure_par_before_content()
            tok: Dict[str, Any] = {"t": "hr"}
            attrs: Dict[str, Any] = {
                k: v for k, v in el.attrib.items() if v is not None}
            if attrs:
                tok["a"] = attrs
            logger.debug(f"Emitting hr token: {attrs}")
            self._runs.append(tok)
            self._plain.append("\n---\n")
            self._emitted_anything = True
            self._emitted_content_since_last_par = True # hr もコンテントとして扱う
            return # hr には子要素も tail もない

        # (B) v1.5-dev: <section> (セクション) のサポート
        if tag == "section":
            self._ensure_par_before_content()
            ns = self._ns # (B) v1.5-dev: 保持している ns を使用
            
            section_tok: Dict[str, Any] = {"t": "section"}
            
            # 1. <sectiontitle> のパース
            title_el = _safe_find(el, f"./{_q('sectiontitle', ns)}", ns)
            title_runs: List[Dict[str, Any]] = []
            
            # (B) v1.5-dev: ns を渡す
            sub_title = RichTextParser(self._pardefs, self._attachment_metas, self._ns)
            sub_title._inline_image_index = self._inline_image_index
            logger.debug("--- Parsing SectionTitle ---")
            
            if title_el is not None:
                if title_el.text:
                    sub_title._emit_text(title_el.text)
                for ch in list(title_el):
                    sub_title._walk(ch)
                    if ch.tail:
                        sub_title._emit_text(ch.tail)
            else:
                logger.warning("Found <section> without <sectiontitle>. Emitting empty title_runs.")
            
            logger.debug(f"--- End SectionTitle (Runs: {len(sub_title._runs)}) ---")
            title_runs = sub_title._runs
            self._inline_image_index = sub_title._inline_image_index # index を引き継ぐ
            
            section_tok["title_runs"] = title_runs

            # 2. セクション本文 (sectiontitle 以外の要素) のパース
            body_runs: List[Dict[str, Any]] = []
            # (B) v1.5-dev: ns を渡す
            sub_body = RichTextParser(self._pardefs, self._attachment_metas, self._ns)
            sub_body._inline_image_index = self._inline_image_index
            logger.debug("--- Parsing SectionBody ---")
            
            # el (section) の子要素をイテレート
            # section 直下のテキストノード (もしあれば)
            if el.text:
                 sub_body._emit_text(el.text)
                 
            for ch in list(el):
                tag_name = _localname(ch.tag)
                # pardef や sectiontitle はスキップ
                if tag_name in ("pardef", "sectiontitle"): 
                    if ch.tail:
                        # sectiontitle の tail は body に属する
                        sub_body._emit_text(ch.tail)
                    continue
                
                # par, table などの body 要素
                sub_body._walk(ch)
                
                if ch.tail:
                    sub_body._emit_text(ch.tail)

            logger.debug(f"--- End SectionBody (Runs: {len(sub_body._runs)}) ---")
            body_runs = sub_body._runs
            self._inline_image_index = sub_body._inline_image_index # index を引き継ぐ
            
            section_tok["body_runs"] = body_runs
            
            # 3. セクション自体の属性 (展開状態など)
            attrs: Dict[str, Any] = {
                k: v for k, v in el.attrib.items() if v is not None}
            if attrs:
                section_tok["attributes"] = attrs
            
            self._runs.append(section_tok)
            self._emitted_anything = True
            self._emitted_content_since_last_par = True
            
            return

        # --- Container/Style Handler: <run> (v1.5.1 修正) ---
        if tag == "run":
            style_applied = _style_from_generic_tag(el) # <run highlight="..."> を解釈
            if style_applied:
                self._style_stack.append(style_applied)
                logger.debug(f"Pushed style (for <run>): {style_applied}")

            if el.text:
                self._emit_text(el.text)
            for ch in list(el):
                self._walk(ch)
                if ch.tail:
                    self._emit_text(ch.tail)
            
            if style_applied:
                self._style_stack.pop()
                logger.debug(f"Popped style (for </run>)")
            
            return

        # --- Style Handlers ---
        style_applied = None
        # etc.
        is_style_tag = tag in ("font", "b", "i", "u", "strike", "sup", "sub")

        if is_style_tag:
            style_applied = _style_from_generic_tag(el)

            # (v1.5.1) DXL <font.../>TEXT (tail) パターン救済
            is_font_slash_text_pattern = (
                tag == "font" and
                (el.text is None or not el.text.strip()) and
                (el.tail and el.tail.strip()) and
                not list(el)  # 子要素がない
            )

            if is_font_slash_text_pattern:
                # 特別パターン: スタイルを適用して tail を emit し、即 return
                if style_applied:
                    self._style_stack.append(style_applied)
                    logger.debug(
                        f"Pushed style (font.../>TEXT): {style_applied}")

                logger.debug(
                    f"Emitting text from font tail (font.../>TEXT): {el.tail.strip()}")
                self._emit_text(el.tail)  # スタイル適用中に tail を emit

                if style_applied:
                    self._style_stack.pop()
                    logger.debug(f"Popped style (font.../>TEXT) for </{tag}>")

                # 親ループ (L806など) で二重処理されないよう、tail をクリアする
                el.tail = None
                return  # Handled
            # --- (修正ここまで) ---

            if style_applied:
                self._style_stack.append(style_applied)
                logger.debug(f"Pushed style: {style_applied}")

            # Process content/children within style context
            if el.text:
                self._emit_text(el.text)
            for ch in list(el):
                self._walk(ch)
                if ch.tail:
                    self._emit_text(ch.tail)  # Tail is *within* style

            # [v1.4.6 バグ修正]
            # el.tail (自分自身の tail) の処理を削除します。
            # tail の処理は、常に親要素 (par や run) のループ (L785 や L769) で
            # ch.tail として処理されるべきです。

            if style_applied:
                self._style_stack.pop()
                logger.debug(f"Popped style for </{tag}>")
            return  # Handled

        # --- Specific Element Handlers ---
        if tag == "urllink":
            href = el.get("href", "")
            label = "".join(el.itertext()).strip() or el.get("title") or href
            st = _collect_inline_style(el)
            self._emit_link(href, label, st)
            return

        if tag == "doclink":
            notes = {k: v for k, v in {
                "server": el.get("server"),
                "replica": el.get("database") or el.get("db"),
                "unid": el.get("document") or el.get("unid"),
                "view": el.get("view")
            }.items() if v}
            label = "".join(el.itertext()).strip() or el.get(
                "description") or "DocLink"
            st = _collect_inline_style(el)
            self._emit_link("", label, st, notes=notes)
            return

        if tag == "attachmentref":
            name = el.get("name")
            displayname = el.get("displayname")
            if name:
                self._emit_attachmentref(name, displayname)
            else:
                logger.warning(
                    "<attachmentref> found without 'name' attribute.")
            # Ignore children (icon picture) for runs
            return

        if tag == "picture":
            parent_el = None
            # ElementTree 1.2 (Python 3.7) doesn't have parent maps or find("..")
            # We need to rely on the limited info we have.
            # Assume _walk is called on a child of attachmentref IF it's an icon.
            # But _emit_inline_image decides based on metadata index, not parent tag.
            # This logic might be flawed if DXL structure is complex.
            # For now, rely on _inline_images_meta index.
            
            # (v1.5.1 追記)
            # `_extract_attachments_metadata` (L519) は、`attachmentref` の中の `picture` を
            # `icon_picture_elements` セットに追加します。
            # `_extract_attachments_metadata` (L527) は、`picture` をイテレートする際、
            # `icon_picture_elements` に *含まれない* `picture` のみ
            # `inline_image` (type: "image", ref: {"element": "picture", "index": ...}) として
            # `attachments` リストに追加します。
            # `RichTextParser.__init__` (L431) は、このリストを `_inline_images_meta` に保存します。
            # したがって、`_emit_inline_image` (L703) が処理する `picture` は、
            # `attachmentref` の子ではない、インライン画像のみのはずです。
            is_inside_attref = False # このコンテキストでは正確な判定は難しいが、
                                      # _inline_images_meta の構築ロジックにより不要なはず。
            
            if not is_inside_attref:
                self._emit_inline_image()
            else:
                logger.debug(
                    "Skipping <picture> inside <attachmentref> for runs.")
            # Ignore children (image data) for runs
            return

        if tag == "br" or tag == "break":
            # (B) v1.5-dev: <br> を明示的なトークンとしてサポート (仕様書 3.2.3 💠 3)
            self._ensure_par_before_content()
            logger.debug("Emitting br token")
            self._runs.append({"t": "br"})
            self._plain.append("\n") # プレーンテキストにも改行を反映
            self._emitted_anything = True
            # br は par 内の改行であり、content_since_last_par を True にする
            self._emitted_content_since_last_par = True
            return

        # --- Ignored / Unhandled Tags ---
        ignored_tags = {
            "pardef", "parstyle", "fonttable", "colortable", "object", "file", "filedata",
            "gif", "jpeg", "png", "bmp", "notesbitmap", "caption", "region"
            # (B) v1.5-dev: section, sectiontitle, horizrule を削除
            # Add others as needed
        }
        if tag in ignored_tags:
            logger.debug(
                f"Ignoring known DXL tag <{tag}> in RichTextParser walk.")
            return  # Assume these contain no direct run content

        # --- Fallback for unknown tags ---
        logger.warning(
            f"Unhandled DXL tag <{tag}>. Processing content/children as plain text.")
        if el.text:
            self._emit_text(el.text)
        for ch in list(el):
            self._walk(ch)
            if ch.tail:
                self._emit_text(ch.tail)


# ---------------------------
# Top-level document parse
# ---------------------------

def parse_dxl_document_from_string(dxl_text: str, db_title: str, richtext_item_name: str = "Body") -> Dict[str, Any]:
    """
    Parses the entire DXL string and returns the initial normalized JSON (dict).
    Attachment/image path info (`content_path`, `src`) will be None.
    """
    logger.info(f"Starting DXL parse ({SCHEMA_VERSION}) for db: {db_title}")

    try:
        # Use XMLPullParser for potentially large files, though fromstring is simpler for now
        root = ET.fromstring(dxl_text)
        ns = _ns_from_root(root)
    except ET.ParseError as e:
        # Log with traceback
        logger.exception(f"Fatal DXL XML Parse Error: {e}")
        return {
            "schema_version": SCHEMA_VERSION,
            "meta": {"db_title": db_title, "unid": None, "form": "Document", "error": f"DXL XML Parse Error: {e}", "schema_version": SCHEMA_VERSION},
            "fields": {}, "attachments": [], "links": {"notes": [], "http": []}, "layout": {}
        }
    except Exception as e:
        logger.exception(f"Unexpected error during XML parsing setup: {e}")
        return {
            "schema_version": SCHEMA_VERSION,
            "meta": {"db_title": db_title, "unid": None, "form": "Document", "error": f"Unexpected XML Setup Error: {e}", "schema_version": SCHEMA_VERSION},
            "fields": {}, "attachments": [], "links": {"notes": [], "http": []}, "layout": {}
        }

    out: Dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "meta": {
            "db_title": db_title,
            "unid": None,
            "form": root.get("form") or "Document",
            # (A) v1.5-dev: meta にもバージョンを追加
            "schema_version": SCHEMA_VERSION
        },
        "fields": {},
        "attachments": [],
        "links": {"notes": [], "http": []},
        "layout": {
            "primary_fields_allowlist": ["Subject", "From", "To", "CC", "Body", "Categories", "Created", "Modified"],
            "used_in_body": ["Subject", richtext_item_name],
        },
    }

    # Extract Metadata
    ni = _safe_find(root, f".//{_q('noteinfo', ns)}", ns)
    if ni is not None:
        out["meta"]["unid"] = ni.get("unid")
        for key, path in [("created", "created"), ("modified", "modified"), ("revised", "revised")]:
            dt_el = _safe_find(
                ni, f"./{_q(path, ns)}/{_q('datetime', ns)}", ns)
            out["meta"][key] = _parse_dxl_datetime(dt_el)
    else:
        out["meta"]["unid"] = root.get("unid")  # Fallback to root unid
        logger.warning(
            f"Could not find <noteinfo>. Using root UNID: {out['meta']['unid']}")

    # Collect Pardefs
    pardefs = _collect_pardefs(root)
    logger.info(f"Collected {len(pardefs)} pardefs.")

    # Extract Attachments Metadata (Initial - paths are None)
    out["attachments"] = _extract_attachments_metadata(root, ns)

    # Extract Fields (Items)
    for item in _safe_findall(root, f".//{_q('item', ns)}", ns):
        name = item.get("name")
        if not name or name.startswith("$") or name == "Form":
            continue

        # --- Handle Richtext Field ---
        if name == richtext_item_name:
            rt_el = _safe_find(item, f"./{_q('richtext', ns)}", ns)
            if rt_el is not None:
                # (B) v1.5-dev: ns をパーサーに渡す
                parser = RichTextParser(pardefs, out["attachments"], ns)
                body_parsed = parser.parse(rt_el)
                out["fields"][name] = {"type": "richtext",
                                       "text": body_parsed.get("text", ""),
                                       "runs": body_parsed.get("runs", [])}
            else:  # Richtext item without <richtext> tag
                text_val = "".join(t.strip()
                                   for t in item.itertext() if t and t.strip())
                out["fields"][name] = {"type": "richtext", "text": text_val,
                                       "runs": [{"t": "text", "text": text_val}] if text_val else []}
        # --- Handle Other Field Types ---
        else:
            field_data = None
            # Prioritize list types first
            if els := _safe_findall(item, f".//{_q('textlist', ns)}/{_q('text', ns)}", ns):
                field_data = {"type": "textlist", "value": [
                    el.text.strip() for el in els if el.text]}
            elif els := _safe_findall(item, f".//{_q('datetimelist', ns)}/{_q('datetime', ns)}", ns):
                field_data = {"type": "datetimelist", "value": [
                    _parse_dxl_datetime(el) for el in els]}
            elif els := _safe_findall(item, f".//{_q('numberlist', ns)}/{_q('number', ns)}", ns):
                field_data = {"type": "numberlist", "value": [
                    el.text.strip() for el in els if el.text]}  # TODO: Parse number?
            # Then single value types
            elif el := _safe_find(item, f"./{_q('text', ns)}", ns):
                if el.text is not None:
                    field_data = {"type": "text", "value": el.text.strip()}
            elif el := _safe_find(item, f"./{_q('datetime', ns)}", ns):
                field_data = {"type": "datetime",
                              "value": _parse_dxl_datetime(el)}
            elif el := _safe_find(item, f"./{_q('number', ns)}", ns):
                if el.text is not None:
                    # TODO: Parse number?
                    field_data = {"type": "number", "value": el.text.strip()}
            else:  # Fallback text extraction
                fallback_text = "".join(t.strip()
                                        for t in item.itertext() if t and t.strip())
                if fallback_text:
                    field_data = {
                        "type": "text", "value": fallback_text, "notes": "Fallback text"}

            if field_data:
                # Clean up None values in lists and only add if value exists
                if 'value' in field_data:
                    if isinstance(field_data['value'], list):
                        field_data['value'] = [
                            v for v in field_data['value'] if v is not None]
                        if field_data['value']:  # Only add if list is not empty after cleanup
                            out["fields"][name] = field_data
                    # Add if single value is not None
                    elif field_data['value'] is not None:
                        out["fields"][name] = field_data
                # else: Don't add if field_data ended up empty

    # Collect Links from Body Runs
    links = {"notes": [], "http": []}
    body_runs = out.get("fields", {}).get(
        richtext_item_name, {}).get("runs", [])

    def walk_for_links(run_list: List[Dict[str, Any]]):
        if not run_list:
            return
        for r in run_list:
            if r.get("t") == "link":
                href = (r.get("href") or "").strip()
                label = r.get("label")
                notes = r.get("notes") or {}
                if notes.get("unid"):
                    raw_link = href or f"notes:///{notes.get('server', '')}/{notes.get('replica', '')}/0/{notes['unid']}?OpenDocument"
                    link_info = {k: v for k, v in {
                        "raw": raw_link, "server": notes.get("server"), "replica": notes.get("replica"),
                        "unid": notes["unid"], "query": "OpenDocument", "resolved_url": None,
                        "meta": {"label": label}
                    }.items() if v is not None}  # Store non-None values
                    links["notes"].append(link_info)
                elif href.startswith("http://") or href.startswith("https://"):
                    links["http"].append(
                        {"raw": href, "label": label, "resolved": None})
            elif r.get("t") == "table":
                for row in r.get("rows", []):
                    for cell in row.get("cells", []):
                        walk_for_links(cell.get("runs", []))
            # (B) v1.5-dev: section 内も再帰的に探索
            elif r.get("t") == "section":
                walk_for_links(r.get("title_runs", []))
                walk_for_links(r.get("body_runs", []))

    walk_for_links(body_runs)
    out["links"] = links

    logger.info(f"DXL parse complete for UNID {out['meta']['unid']}.")
    return out


# ---------------------------
# Unit test entrypoint helpers (NEW)
# ---------------------------


def _find_run_recursive(run_list: List[Dict[str, Any]], predicate: callable) -> Optional[Dict[str, Any]]:
    """runs のリストを再帰的に探索し、述語(predicate)に一致する最初の run を返す"""
    if not run_list:
        return None
    for r in run_list:
        if predicate(r):
            return r
        if r.get("t") == "table":  # Recurse into table cells
            for row in r.get("rows", []):
                for cell in row.get("cells", []):
                    found = _find_run_recursive(
                        cell.get("runs", []), predicate)
                    if found:
                        return found
        # (B) v1.5-dev: section 内も再帰的に探索
        if r.get("t") == "section":
            found_title = _find_run_recursive(r.get("title_runs", []), predicate)
            if found_title:
                return found_title
            found_body = _find_run_recursive(r.get("body_runs", []), predicate)
            if found_body:
                return found_body
    return None


def count_tokens_recursive(run_list: List[Dict[str, Any]], token_type: str) -> int:
    """Recursively counts tokens of a specific type within a runs list, including tables and sections."""
    count = 0
    if not run_list:
        return 0
    for r in run_list:
        if r.get("t") == token_type:
            count += 1
        elif r.get("t") == "table":
            for row in r.get("rows", []):
                for cell in row.get("cells", []):
                    count += count_tokens_recursive(
                        cell.get("runs", []), token_type)
        # (B) v1.5-dev: section 内も再帰的に探索
        elif r.get("t") == "section":
            count += count_tokens_recursive(r.get("title_runs", []), token_type)
            count += count_tokens_recursive(r.get("body_runs", []), token_type)
    return count


def _find_all_runs_recursive(run_list: List[Dict[str, Any]], token_type: str) -> List[Dict[str, Any]]:
    """Recursively finds all tokens of a specific type within a runs list, including tables and sections."""
    found_runs = []
    if not run_list:
        return found_runs
    for r in run_list:
        if r.get("t") == token_type:
            found_runs.append(r)
        elif r.get("t") == "table":
            for row in r.get("rows", []):
                for cell in row.get("cells", []):
                    found_runs.extend(_find_all_runs_recursive(
                        cell.get("runs", []), token_type))
        # (B) v1.5-dev: section 内も再帰的に探索
        elif r.get("t") == "section":
            found_runs.extend(_find_all_runs_recursive(
                r.get("title_runs", []), token_type))
            found_runs.extend(_find_all_runs_recursive(
                r.get("body_runs", []), token_type))
    return found_runs


# ---------------------------
# Unit test entrypoint
# ---------------------------
if __name__ == "__main__":
    logger.setLevel(logging.INFO)  # Set log level for testing
    if logger.handlers:
        logger.handlers[0].setLevel(logging.INFO)

    # --- Configuration for Test ---
    # !!! Adjust DXL path if necessary !!!
    # (v1.5) テストパスをデバッグ用にローカルの絶対パスに変更
    # (実行環境に合わせて変更してください)
    # xml_path_str = r"C:\path\to\your\test\670309E2B88C5E9649258D2C000ADBDE.xml"
    # DXL ファイルをカレントディレクトリから探すように変更
    script_dir = Path(__file__).parent.resolve()
    xml_path = script_dir / "670309E2B88C5E9649258D2C000ADBDE.xml"
    
    if not xml_path.exists():
        # ローカルでのデバッグパスフォールバック (必要に応じて変更)
        logger.warning(f"DXL not found at {xml_path}, trying debug path...")
        xml_path_str_debug = r"C:\Users\A512292\py2024\notes_to_box_v4_2_0\dist\output\manifest_run\012\_dxl\TEMP_DB\670309E2B88C5E9649258D2C000ADBDE.xml"
        xml_path = Path(xml_path_str_debug)

    # Output JSON path (in the same directory as DXL)
    json_out_dir = xml_path.parent
    # Indicate v1.5 state
    json_filename = f"{xml_path.stem}.normalized.{SCHEMA_VERSION}.json"
    out_path = json_out_dir / json_filename

    db_title_str = "TEMP_DB"  # Example DB title
    # --- End Configuration ---

    if xml_path.exists():
        print(
            f"--- Running {SCHEMA_VERSION} Parser Unit Test (DXL -> Initial JSON) ---")
        print(f"Input DXL: {xml_path}")
        print(f"Output JSON: {out_path}")

        try:
            # Read DXL
            try:
                dxl = xml_path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                logger.warning("UTF-8 decode failed, trying 'utf-8-sig'")
                dxl = xml_path.read_text(encoding="utf-8-sig", errors="ignore")

            # Parse DXL to Initial JSON (dict)
            result = parse_dxl_document_from_string(dxl, db_title=db_title_str)

            # Ensure output directory exists
            json_out_dir.mkdir(parents=True, exist_ok=True)

            # Write Initial JSON to file
            with out_path.open("w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(
                f"\n[OK] Generated and wrote Initial JSON ({SCHEMA_VERSION}) to: {out_path}")

            # --- Verification (v1.5) ---
            print(f"\n--- Verification of Initial JSON ({SCHEMA_VERSION}) ---")
            meta = result.get("meta", {})
            fields = result.get("fields", {})
            attachments = result.get("attachments", [])
            links = result.get("links", {})
            body = fields.get("Body", {})
            runs = body.get("runs", [])

            # Basic structure checks
            assert result.get(
                "schema_version") == SCHEMA_VERSION, f"Schema version mismatch (root)"
            assert meta.get(
                "schema_version") == SCHEMA_VERSION, f"Schema version mismatch (meta)"
            assert meta.get(
                "unid") == "670309E2B88C5E9649258D2C000ADBDE", f"UNID mismatch"
            assert len(fields) > 0, "No fields parsed"
            assert "Body" in fields, "Body field missing"
            assert len(runs) > 0, "Body runs are empty"

            print(f"  Schema Version (root): OK ({SCHEMA_VERSION})")
            print(f"  Schema Version (meta): OK ({meta.get('schema_version')})")
            print(f"  UNID: OK ({meta.get('unid')})")
            
            # 1. Verify Attachments Metadata (v1.5 - アイコンは除外)
            print(f"  Total Attachments (Metadata): {len(attachments)}")
            # $FILE (3) + inline image (1) = 4
            expected_att_count = 4
            assert len(
                attachments) == expected_att_count, f"Expected {expected_att_count} attachment metadata entries (files+inline), found {len(attachments)}"
            all_paths_none = all(a.get('content_path') is None and a.get(
                'saved_name') is None for a in attachments)
            assert all_paths_none, "Found non-None 'content_path' or 'saved_name' in initial attachments metadata!"
            print(f"  [Check] Attachments metadata paths are None: OK")
            
            # 2. Verify Runs Tokens (Recursive)
            img_token_count = count_tokens_recursive(runs, "img")
            attref_token_count = count_tokens_recursive(runs, "attachmentref")
            assert img_token_count == 1, f"Expected 1 'img' token recursively, found {img_token_count}"
            assert attref_token_count == 3, f"Expected 3 'attachmentref' tokens recursively, found {attref_token_count}"
            print(f"  [Check] Found {img_token_count} 'img' tokens: OK")
            print(
                f"  [Check] Found {attref_token_count} 'attachmentref' tokens: OK")

            # 3. Verify v1.5 Elements (hr, section, br)
            hr_token_count = count_tokens_recursive(runs, "hr")
            section_token_count = count_tokens_recursive(runs, "section")
            br_token_count = count_tokens_recursive(runs, "br")
            
            assert hr_token_count == 1, f"Expected 1 'hr' token, found {hr_token_count}"
            assert section_token_count == 1, f"Expected 1 'section' token, found {section_token_count}"
            # このサンプルDXLには <br> はない
            assert br_token_count == 0, f"Expected 0 'br' tokens, found {br_token_count}" 
            print(f"  [Check] Found {hr_token_count} 'hr' token: OK")
            print(f"  [Check] Found {section_token_count} 'section' token: OK")
            
            section_run = _find_run_recursive(runs, lambda r: r.get("t") == "section")
            assert section_run is not None, "Section run not found"
            assert len(section_run.get("title_runs", [])) > 0, "Section title_runs is empty"
            assert len(section_run.get("body_runs", [])) == 0, f"Section body_runs should be empty (or just par), found {len(section_run.get('body_runs', []))}"
            print(f"  [Check] Section content (title/body runs): OK")

            # 4. Verify v1.5 Table structure (columns, refwidth, attributes.tablabel, cell.colspan)
            table_run = _find_run_recursive(runs, lambda r: r.get("t") == "table" and r.get("attributes", {}).get("rowdisplay") == "tabs")
            assert table_run is not None, "Tabbed table not found"
            
            assert "refwidth" in table_run.get("attributes", {}), "Tabbed table missing attributes.refwidth"
            assert "columns" in table_run, "Tabbed table missing columns"
            assert len(table_run.get("columns", [])) == 2, "Tabbed table columns count mismatch"
            
            first_row = (table_run.get("rows", []) or [{}])[0]
            assert "tablabel" in first_row.get("attributes", {}), "Tabbed table row missing attributes.tablabel"
            
            first_cell = (first_row.get("cells", []) or [{}])[0]
            # このサンプルでは colspan/rowspan はない
            assert "colspan" not in first_cell, "Tabbed table cell has unexpected colspan"
            assert "style" in first_cell, "Tabbed table cell missing style"
            assert "bgcolor" in first_cell.get("style", {}), "Tabbed table cell style missing bgcolor"
            print(f"  [Check] v1.5 Table structure (refwidth, columns, tablabel, style): OK")
            
            # 5. Verify v1.5 font_family (v1.5.1 修正により検出されるはず)
            font_run = _find_run_recursive(runs, lambda r: r.get("a", {}).get("font_family") == "ＭＳ Ｐゴシック")
            assert font_run is not None, "font_family 'ＭＳ Ｐゴシック' not found (v1.5.1 fix failed?)"
            print(f"  [Check] Found 'font_family' attribute (v1.5.1 fix): OK")

            # 6. Verify v1.5.1 style fixes (highlight, color, style)
            highlight_run = _find_run_recursive(runs, lambda r: r.get("text") == "蛍光ペ" and r.get("a", {}).get("bgcolor") == "yellow")
            assert highlight_run is not None, "highlight='yellow' style missing (v1.5.1 fix failed?)"
            print(f"  [Check] Found 'highlight' (bgcolor) attribute (v1.5.1 fix): OK")
            
            red_font_run = _find_run_recursive(runs, lambda r: r.get("text") == "AAAA" and r.get("a", {}).get("color") == "red")
            assert red_font_run is not None, "red color style missing (v1.5.1 fix failed?)"
            assert "s" in red_font_run and "i" in red_font_run["s"] and "s" in red_font_run["s"], "italic/strikethrough style missing (v1.5.1 fix failed?)"
            print(f"  [Check] Found color/italic/strike styles (v1.5.1 fix): OK")


            print("\n--- Verification Complete: Initial JSON (v1.5.1 fix) seems OK ---")

        # ... (既存のエラーハンドリング) ...
        except ET.ParseError as e:
            print(f"[ERROR] Failed to parse XML file: {xml_path}. Error: {e}")
            logger.exception("XML Parse Error")
            print(traceback.format_exc())
        except FileNotFoundError as e:
            print(f"[ERROR] File not found: {e}")
            logger.exception("File Not Found Error")
            print(traceback.format_exc())
        except AssertionError as e:
            print(f"[FAIL] Verification failed: {e}")
            logger.error(f"Verification failed: {e}", exc_info=True)
            print(traceback.format_exc())
        except Exception as e:
            print(f"[ERROR] An unexpected error occurred: {e}")
            logger.exception("Unexpected Error")
            print(traceback.format_exc())

    else:
        print(f"[ERROR] Input DXL file not found: {xml_path}")
        logger.error(f"Input DXL file not found: {xml_path}")
