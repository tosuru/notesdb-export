"""
HTML レンダリングエンジン (HtmlRenderer v1.6-dev)

[v1.6 改善提案による修正]
- [提案1] document.jinja2.html 側で、通常テーブルの <td> 属性の
           レンダリング方法を修正し、タグ内部の不要な空白を除去。
- [提案2] BeautifulSoup4 を使用した「美化 (pretty-print)」機能を
           render() メソッドのオプションとして追加。
- [提案2] BeautifulSoup4 がインストールされていない場合でも
           フォールバックして動作するように修正。

[v1.6.2 修正]
- colspan/columnspan が style の中/外どちらにあっても処理するよう修正
- render() の pretty=True をデフォルトに変更
"""

from __future__ import annotations
import os
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

try:
    from jinja2 import Environment, FileSystemLoader, select_autoescape
except ImportError:
    raise ImportError(
        "Jinja2 is required for HTML rendering. Install with: pip install Jinja2")

# [提案2] BeautifulSoup4 をインポート
try:
    from bs4 import BeautifulSoup
    _BS4_AVAILABLE = True
except ImportError:
    BeautifulSoup = None
    _BS4_AVAILABLE = False
    # logger はここではまだ使えないので、モジュールレベルで警告
    logging.warning("beautifulsoup4 not installed. 'pretty=True' formatting will be unavailable.")


# --- common.py から BaseRenderer をインポート ---
try:
    # パッケージとして実行される場合
    from ..common import BaseRenderer, Style
except ImportError:
    # スクリプトとして直接実行される場合など
    import sys
    # common.py が存在するディレクトリパスを取得
    common_dir = Path(__file__).resolve().parent.parent
    if str(common_dir) not in sys.path:
        sys.path.insert(0, str(common_dir))  # common.py の親ディレクトリをパスに追加
    from common import BaseRenderer, Style  # type: ignore


# --- 依存ユーティリティ (スタブ - 実際には別モジュールからインポート) ---
# utils/font.py 等に実装されている想定
def resolve_font_path() -> Optional[Path]:
    """フォントパスを解決する (スタブ)"""
    # logger.warning("Using stub for resolve_font_path()") # デバッグレベルで十分
    return None


def font_file_url(font_path: Optional[Path]) -> Optional[str]:
    """フォントパスを data: URL に変換する (スタブ)"""
    import base64
    if font_path and font_path.exists():
        logger.debug(f"Generating data URL for font: {font_path}")
        try:
            mime_types = {'.ttf': 'font/truetype', '.otf': 'font/opentype',
                          '.woff': 'font/woff', '.woff2': 'font/woff2'}
            mime = mime_types.get(font_path.suffix.lower())
            if not mime:
                logger.warning(
                    f"Unknown font extension: {font_path.suffix}, cannot generate correct data URL.")
                return None
            with open(font_path, 'rb') as f:
                encoded = base64.b64encode(f.read()).decode('utf-8')
            return f"data:{mime};base64,{encoded}"
        except Exception as e:
            logger.error(
                f"Failed to generate data URL for font {font_path}: {e}")
            return None
    return None
# --- 依存ユーティリティここまで ---


def _reset_list_state_before_block(self):
    """ブロック要素開始前に list_state を確実にリセット"""
    try:
        ls = getattr(self.context, 'list_state', {})
        ls['level'] = 0
        ls['type'] = None
        self.context.list_state = ls
    except Exception as e:
        try:
            self.logger.warning(f"Failed to reset list state: {e}")
        except Exception:
            pass


logger = logging.getLogger(__name__)

# --- Jinja2 テンプレート関連の定数 ---
_CURRENT_DIR = Path(__file__).resolve().parent
# パスの階層を修正 (src/core/render/engines -> root/render/templates/html)
_DEFAULT_TEMPLATE_DIR = _CURRENT_DIR.parent.parent / "render" / "templates" / "html"
_DEFAULT_TEMPLATE_NAME = "document.jinja2.html"


class HtmlRenderer(BaseRenderer):

    def _default_link_attrs(self) -> Dict[str, str]:
        # すべての a 要素を別タブ＋安全に開く
        return {"target": "_blank", "rel": "noopener noreferrer"}

    def _reset_list_state_before_block(self):
        """Ensure list_state is reset before emitting a block element."""
        try:
            ls = getattr(self.context, 'list_state', {}) or {}
            ls['level'] = 0
            ls['type'] = None
            self.context.list_state = ls
        except Exception as e:
            # HtmlRenderer は self.logger を持たないので module logger を使う
            logger.warning(f"Failed to reset list state: {e}")

    def _update_list_state_from_par(self, run):
        """Update list_state from paragraph run.attrs.list (simple level mapping)."""
        try:
            a = (run.get('a') or {})
            lst = (a.get('list') or {})
            t = lst.get('type')
            ls = getattr(self.context, 'list_state', {})
            if t:
                ls['level'] = 1
                ls['type'] = t
            else:
                ls['level'] = 0
                ls['type'] = None
            self.context.list_state = ls
        except Exception as e:
            try:
                self.logger.warning(f"Failed to update list state: {e}")
            except Exception:
                pass

    def _indent_class_from_par(self, par_attrs):
        try:
            lm = par_attrs.get('leftmargin')
            if not lm:
                lst = par_attrs.get('list') or {}
                return ""

            def to_in(v):
                if isinstance(v, (int, float)):
                    return float(v)
                v = str(v)
                if v.endswith('in'):
                    return float(v[:-2])
                if v.endswith('pt'):
                    return float(v[:-2]) / 72.0
                return 0.0
            inches = to_in(lm)
            if 0.75 <= inches < 1.25:
                return " indent-1"
            elif 1.25 <= inches < 1.75:
                return " indent-2"
            elif inches >= 1.75:
                return " indent-3"
            else:
                return ""
        except Exception:
            return ""
    """
    normalized.json (runs) を HTML 文字列に変換するレンダラー。
    Jinja2 テンプレートを使用して最終的なHTMLを構築します。
    """

    def __init__(self, doc_path: Path, doc: Optional[Dict[str, Any]] = None,
                 shared_jinja_env: Optional[Environment] = None):
        """
        Jinja2 環境を初期化し、BaseRenderer を呼び出します。
        [v1.4] 添付アイコンマップを事前構築します。
        """
        self.template_dir = _DEFAULT_TEMPLATE_DIR

        if shared_jinja_env:
            self.jinja_env = shared_jinja_env
        elif not hasattr(self, 'jinja_env'):
            try:
                self.jinja_env = Environment(
                    loader=FileSystemLoader(
                        str(self.template_dir), followlinks=True),
                    autoescape=select_autoescape(["html", "xml"]),
                    trim_blocks=True,  # [v1.6] テンプレート側の空白制御を有効化
                    lstrip_blocks=True,  # [v1.6] テンプレート側の空白制御を有効化
                )
                logger.debug(
                    f"Jinja2 environment loaded from: {self.template_dir}")
            except Exception as e:
                logger.error(
                    f"Failed to load Jinja2 environment from {self.template_dir}: {e}", exc_info=True)
                raise RuntimeError(
                    f"Could not initialize Jinja2 environment from {self.template_dir}") from e

        # BaseRenderer の __init__ を呼び出す
        # doc が提供されていれば、ファイルロードはスキップされる
        super().__init__(doc_path, doc=doc)

        # --- HTMLRenderer 固有の状態 ---
        self._current_block: Optional[Dict[str, Any]] = None
        self._current_list_block: Optional[Dict[str, Any]] = None

        # [改善案 7] タブ付き表のグループID用カウンター
        self._tab_group_counter = 0
        # NEW v1.6: track structural blank emission to collapse multiple empty paragraphs
        self._last_emitted_was_blank = False
        # [修正 14] 最後に有効だったインデントクラスを保持（Notesが添付ファイルの直前に空のparを挿入してくる対策）
        self._last_meaningful_indent_cls = ""

        # --- [修正 4.A] v1.4対応: 添付ファイル辞書 (content_path -> icon_path) を構築 ---
        self._attachment_icon_map = self._build_icon_map()
        logger.debug(
            f"Attachment icon map built with {len(self._attachment_icon_map)} entries.")

    @classmethod
    def _create_sub_renderer(cls, parent_renderer: HtmlRenderer) -> HtmlRenderer:
        """
        親レンダラーの doc と jinja_env を共有するサブレンダラーを作成する。
        (Jinja2 env のロードをスキップする)
        """
        sub = cls(
            parent_renderer.doc_path,
            doc=parent_renderer.doc,
            shared_jinja_env=parent_renderer.jinja_env
        )
        sub.template_dir = parent_renderer.template_dir
        # [v1.4] 親で構築したアイコンマップを共有
        sub._attachment_icon_map = parent_renderer._attachment_icon_map
        # [改善案 7] タブグループカウンターを共有 (ネストした場合にインクリメントするため)
        # ただし、サブレンダラーはセル内などで使われるため、親のカウンターを共有するのが正しい
        sub._tab_group_counter = parent_renderer._tab_group_counter
        sub._last_meaningful_indent_cls = parent_renderer._last_meaningful_indent_cls
        return sub

    # --- [提案2] BaseRenderer.render をオーバーライドして pretty 引数を追加 ---
    # [v1.6.2] pretty=True をデフォルトに変更
    def render(self, pretty: bool = True) -> str:
        """
        ドキュメントをレンダリングします。
        :param pretty: True (デフォルト) の場合、BeautifulSoup を使用して
                       インデント付きの HTML を返します (デバッグ用)。
                       False の場合、最小化された HTML を返します。
        :return: レンダリングされた HTML 文字列。
        """
        try:
            self.output = self._init_output()
            self._render_header()

            # --- 本文 (Body) のレンダリング ---
            body_runs = self.doc.get("fields", {}).get(
                "Body", {}).get("runs", [])
            logger.debug(
                f"Starting to process {len(body_runs)} runs for Body...")
            # self.context は BaseRenderer の __init__ で初期化される
            self.context.paragraph_started = False
            self.context.list_state = {'level': 0, 'type': None}

            # _process_runs (v1.5) を呼び出す
            self._process_runs(body_runs)

            # 最後の段落を確定 (v1.5 では _process_runs の最後でも呼ばれるが念のため)
            self._finalize_paragraph()
            logger.debug("Finished processing Body runs.")

            self._render_appendix()
            self._render_footer()

            # get_output に pretty フラグを渡す
            return self.get_output(pretty=pretty)

        except Exception as e:
            logger.error(f"Render failed: {e}", exc_info=True)
            try:
                # エラー発生時も get_output_on_error を呼ぶ
                return self.get_output_on_error(e)
            except Exception as e_handler:
                # エラーハンドラ自体が失敗した場合の最終フォールバック
                return f"<h1>Fatal Render Error</h1><pre>{e}</pre><pre>{e_handler}</pre>"

    # --- [修正 12] v1.5 BaseRenderer._process_runs をオーバーライド ---

    def _process_runs(self, runs: List[Dict[str, Any]]):
        """
        v1.5仕様: runs配列を処理します。
        'par' トークンがインラインコンテンツ (text, link, img, br) のコンテナを開始します。
        'table', 'hr', 'section', 'attachmentref' はブロックレベル要素として扱われます。
        """

        # v1.5 ディスパッチャ (ブロックレベル要素)
        block_handlers = {
            "table": self._handle_table,
            "hr": self._handle_hr,
            "section": self._handle_section,
            "attachmentref": self._handle_attachmentref,
        }

        # v1.5 ディスパッチャ (インライン要素 - par コンテキスト内で処理)
        inline_handlers = {
            "text": self._handle_text,
            "link": self._handle_link,
            "img": self._handle_img,
            "br": self._handle_br,
        }

        if not runs:
            logger.debug("Processing empty runs list.")
            return

        for run in runs:
            run_type = run.get('t')

            if run_type == 'par':
                # 前の段落を確定
                self._finalize_paragraph()

                # [改善案 4] BaseRenderer のリスト状態更新ロジックを呼び出す
                # (BaseRenderer._process_run が list_state を更新すると仮定)
                # この呼び出しにより self.context.list_state が更新される
                try:
                    self._update_list_state_from_par(run)
                except Exception as e:
                    logger.warning(
                        f"Failed to update list state via self._update_list_state_from_par(run): {e}", exc_info=True)

                # 新しい段落を開始
                par_style = run.get('a', {})
                self._start_paragraph(par_style)

            elif run_type in block_handlers:
                # ブロックレベル要素 (table, hr, section, attachmentref)
                # これらは段落の外にあるので、前の段落を確定
                self._finalize_paragraph()

                # [改善案 4] ブロック要素の前にリストが終了したことを
                # BaseRenderer に伝える (par トークンがないため)
                try:
                    self._reset_list_state_before_block()
                    self._finalize_paragraph()
                except Exception as e:
                    logger.warning(
                        f"Failed to reset list state before block element: {e}")

                block_handlers[run_type](run)
                # これらの要素の後は段落が開始されていない状態
                self.context.paragraph_started = False

            elif run_type in inline_handlers:
                # インライン要素 (text, link, img, br)
                # 段落が開始されていなければ、デフォルト段落を開始
                self._ensure_paragraph_started()
                inline_handlers[run_type](run)

            else:
                # 不明なトークン
                logger.warning(
                    f"Unknown run type '{run_type}'. Handling as unknown.")
                self._ensure_paragraph_started()
                self._handle_unknown(run)

        # 最後の段落を確定
        self._finalize_paragraph()

    # [改善案 4] v1.5 では _process_runs がディスパッチするため、
    # _process_run は BaseRenderer の実装を呼び出し、リスト状態の更新のみを担う
    def _process_run(self, run: Dict[str, Any]):
        """
        v1.5: _process_runs から 'par' トークンの場合に呼び出され、
        BaseRenderer のリスト状態更新ロジックを実行する。
        """
        try:
            # BaseRenderer (親) の _process_run を呼び出す
            # これが self.context.list_state を更新すると仮定
            self._update_list_state_from_par(run)
        except AttributeError:
            # BaseRenderer に _process_run がない (または呼び出せない) 場合
            logger.warning(
                "self._update_list_state_from_par(run) failed. List state might be incorrect.", exc_info=True)
        except Exception as e:
            logger.error(
                f"Error during self._update_list_state_from_par(run): {e}", exc_info=True)
    # ----------------------------------------------------

    # --- BaseRenderer 抽象メソッドの実装 ---

    def _init_output(self) -> Dict[str, Any]:
        """
        Jinja2 に渡すコンテキスト辞書を初期化します。
        'body_elements' が本文のブロック要素 (p, list, table) のリストになります。
        """
        return {
            "subject": "No Subject",  # デフォルト値
            "meta": {},
            "body_elements": [],  # 本文ブロック (p, list, table, hr, section) のリスト
            "appendix_rows": [],
            "font_url": None,     # @font-face 用の data URL
            "styles": {},         # 将来的な CSS クラス管理用 (今は未使用)
            "error_message": None  # レンダリングエラーメッセージ用
        }

    # [提案2] get_output に pretty 引数を追加
    def get_output(self, pretty: bool = True) -> str:
        """
        Jinja2 テンプレートをレンダリングして、最終的な HTML 文字列を返します。
        [v1.6.2] pretty=True がデフォルト
        """
        try:
            font_path = resolve_font_path()
            self.output["font_url"] = font_file_url(
                font_path) if font_path else None
        except Exception as e:
            logger.warning(
                f"Could not resolve/process font path: {e}", exc_info=True)
            self.output["font_url"] = None

        try:
            template_path = self.template_dir / _DEFAULT_TEMPLATE_NAME
            if not template_path.is_file():
                raise FileNotFoundError(
                    f"Jinja2 template not found: {template_path}")

            template = self.jinja_env.get_template(_DEFAULT_TEMPLATE_NAME)
            logger.info(
                f"Rendering HTML using template: {_DEFAULT_TEMPLATE_NAME} (Pretty={pretty})")

            # [改善案 7] サブレンダラーのカウンターがインクリメントされている可能性があるため、
            # ルートレンダラーのカウンターをグローバルコンテキストとして渡す
            render_context = self.output.copy()
            render_context["_tab_group_counter"] = self._tab_group_counter

            html = template.render(**render_context)

            # 1. Python による最小化 (タグ間の空白削除)
            html_minified = self._postprocess_html(html)

            # 2. [提案2] pretty フラグに基づき、美化 (prettify) を行う
            if pretty:
                if not _BS4_AVAILABLE:
                    logger.warning(
                        "pretty=True was requested, but beautifulsoup4 is not installed. Returning minified HTML.")
                    return html_minified

                try:
                    logger.debug("Applying BeautifulSoup prettify...")
                    # 最小化されたHTMLをパースし、美化して返す
                    soup = BeautifulSoup(html_minified, "html.parser")
                    # [v1.6.2] soup.prettify() はHTML5のvoidタグ(br, hr, img)を
                    # 閉じてしまう(<img> -> <img/>)ことがある。
                    # formatter="html5" を使うと、それを防ぎ、
                    # 現代的なHTML5の書式で出力する。
                    return soup.prettify(formatter="html5")
                except Exception as e:
                    logger.error(
                        f"BeautifulSoup prettify failed: {e}. Returning minified HTML.", exc_info=True)
                    return html_minified

            # デフォルト (pretty=False) の場合は最小化されたHTMLを返す
            return html_minified

        except Exception as e:
            logger.error(
                f"Failed to render Jinja2 template {_DEFAULT_TEMPLATE_NAME}: {e}", exc_info=True)
            raise RuntimeError(f"Jinja2 rendering failed: {e}") from e

    def _postprocess_html(self, html: str) -> str:
        # """
        # 安全な後処理でホワイトスペースを圧縮。
        #     - <pre>/<code> ブロックは除外
        #     - タグ間スペース (>\s+<) を圧縮
        #     - 余計な空白・改行の縮約
        #     - 空の <span>/<p> の削除
        # """

        import re

        def squash_bt(s: str) -> str:
            # タグとタグの間のすべての空白（改行、インデント、スペース）を削除
            return re.sub(r'>\s+<', '><', s)

        parts = []
        idx = 0
        # <pre> または <code> ブロックを検索 (大文字小文字無視、複数行対応)
        for m in re.finditer(r'(?is)<(pre|code)\b.*?</\1\s*>', html):
            # <pre> までのチャンクを処理
            chunk = html[idx:m.start()]
            chunk = squash_bt(chunk)  # タグ間空白の削除
            chunk = re.sub(r'[ \t\f\v\r]+\n', '\n', chunk)  # 行末の余分なスペースを削除
            chunk = re.sub(r'\n{2,}', '\n', chunk)  # 複数の改行を1つに
            parts.append(chunk)

            # <pre> ブロック自体は変更せずにそのまま追加
            parts.append(m.group(0))  # untouched
            idx = m.end()

        # 最後のチャンク (または <pre> がない場合は全体) を処理
        tail = squash_bt(html[idx:])
        tail = re.sub(r'[ \t\f\v\r]+\n', '\n', tail)
        tail = re.sub(r'\n{2,}', '\n', tail)
        parts.append(tail)

        compact = ''.join(parts)

        # 空のタグを削除
        compact = re.sub(
            r'<span(?:\s+[^>]*)?>\s*</span>', '', compact, flags=re.IGNORECASE)
        compact = re.sub(r'<p(?:\s+[^>]*)?>\s*</p>',
                         '', compact, flags=re.IGNORECASE)

        return compact.strip()  # 最後に全体を strip

    def _render_header(self):
        """文書の Subject と Meta 情報を self.output (コンテキスト) に格納します。"""
        self.output["subject"] = self.doc.get("fields", {}).get(
            "Subject", {}).get("value", "No Subject")
        self.output["meta"] = self.doc.get("meta", {})
        logger.debug(
            f"Header set: Subject='{self.output['subject'][:50]}...', Meta keys={list(self.output['meta'].keys())}")

    def _render_footer(self):
        """フッター (HTMLでは通常不要だが、テンプレート内で定義可能)"""
        pass

    def _render_appendix(self):
        """付録データを self.output (コンテキスト) に格納します。"""
        self.output["appendix_rows"] = self._build_appendix_rows_helper()
        logger.debug(
            f"Appendix rows generated: {len(self.output['appendix_rows'])} rows.")

    def _start_paragraph(self, par_style: Dict[str, Any]):
        """
        'par' トークンに基づき、新しい段落 (<p> または <li>) ブロックを開始します。
        作成したブロックを適切なリスト (body_elements or current_list_block['items']) に追加します。
        [改善案 4] このメソッドは self.context.list_state が更新済みであることを前提とする。
        """
        # [v1.5] _process_runs で par の前に _finalize_paragraph が呼ばれる
        self.context.paragraph_started = True

        # [改善案 4] _process_runs で更新された list_state を参照
        list_state = self.context.list_state
        css_style = self._style_from_par(par_style)  # CSS文字列を生成

        if list_state['level'] > 0:
            # --- リストアイテムとして開始 ---
            list_tag = 'ol' if list_state.get('type') in (
                'number', 'alphaupper', 'alphalower', 'romanupper', 'romanlower') else 'ul'
            list_type_attr = self._get_ol_type_attr(
                list_state.get('type'))  # <ol type="...">

            if not self._current_list_block or self._current_list_block['tag'] != list_tag:
                # [改善案 4] リストタイプが変更された場合も新しいリストブロックを開始
                if self._current_list_block:
                    logger.debug(
                        f"Changing list block type from {self._current_list_block['tag']} to {list_tag}")

                self._current_list_block = {
                    "type": "list",
                    "tag": list_tag,
                    "attributes": {"type": list_type_attr} if list_type_attr else {},
                    "items": [],
                    "style": ""  # リストブロック全体のスタイル (必要であれば)
                }
                self.output['body_elements'].append(self._current_list_block)
                logger.debug(
                    f"Started new list block: <{list_tag}> (Level={list_state['level']})")

            # 新しい <li> ブロックを作成
            self._current_block = {
                "type": "li",
                "content": [],  # このリストに span, a, img などが入る
                "style": css_style,  # <li style="...">
                # indent class if any
                "attributes": {"class": (self._indent_class_from_par(par_style).strip() or None)}
            }
            # <li> を現在のリストブロックに追加
            self._current_list_block['items'].append(self._current_block)
            logger.debug(f"Started new list item (li) inside <{list_tag}>")

        else:
            # --- 通常の段落として開始 ---
            self._current_list_block = None  # 現在のリストブロックへの参照をクリア

            indent_cls = self._indent_class_from_par(par_style).strip()
            if indent_cls:
                self._last_meaningful_indent_cls = indent_cls

            # 新しい <p> ブロックを作成
            self._current_block = {
                "type": "p",
                "content": [],  # このリストに span, a, img などが入る
                "style": css_style,  # <p style="...">
                # indent class if any
                "attributes": {"class": (indent_cls or None)}
            }
            # <p> を body_elements に追加
            self.output['body_elements'].append(self._current_block)
            logger.debug(f"Started new paragraph block (p)")

    def _finalize_paragraph(self):
        """
        現在の段落ブロック (<p> または <li>) を確定します。
        v1.6 空行方針:
        - 空の <p> を &nbsp; で埋めない
        - 直前が段落/リストならそこに <br> を1つだけ差し込み、連続空行は抑止
        """
        if self._current_block:
            block_type = self._current_block.get('type')
            content = self._current_block.get('content', [])

            if block_type == 'p' and not content:
                # 直前の要素に <br> を付与して空行表現とする（連続は抑止）
                if self.output.get("body_elements"):
                    # body_elements から空の <p> を取り除く
                    try:
                        last_p = self.output["body_elements"].pop()
                        if last_p is not self._current_block:
                            # 予期しない状態: 戻して何もしない
                            self.output["body_elements"].append(last_p)
                            logger.warning(
                                "Mismatch in _finalize_paragraph: last element was not current_block.")
                        else:
                            # 空の <p> を削除したので、その前の要素を探す
                            if self.output.get("body_elements"):
                                prev = self.output["body_elements"][-1]
                                if not self._last_emitted_was_blank and prev.get("type") in ("p", "list"):
                                    # prev が <p> の場合
                                    if prev.get("type") == "p":
                                        prev.setdefault("content", []).append(
                                            {"type": "br"})
                                    # prev が <list> の場合、その最後の <li> に <br> を追加
                                    elif prev.get("type") == "list" and prev.get("items"):
                                        prev["items"][-1].setdefault(
                                            "content", []).append({"type": "br"})

                                    self._last_emitted_was_blank = True
                    except IndexError:
                        pass  # body_elements が空だった

                # 空<p> 自体は捨てる
                self._current_block = None
                self.context.paragraph_started = False
                return

            # 非空段落の場合はフラグをリセット
            self._last_emitted_was_blank = False

            # ブロックを確定
            self._current_block = None
            self.context.paragraph_started = False
        else:
            self._last_emitted_was_blank = False

    def _ensure_paragraph_started(self):
        """まだ段落が開始されていない場合、デフォルトの段落 (<p>) を開始します。"""
        if not self.context.paragraph_started:
            logger.debug("No paragraph started, ensuring default paragraph.")

            # [改善案 4] デフォルト段落開始時も list_state を更新(リセット)
            try:
                # 空の par と同等の扱いで list_state をリセット
                self._update_list_state_from_par({"a": {}})
            except Exception as e:
                logger.warning(
                    f"Failed to reset list state for default paragraph: {e}")

            # デフォルトスタイル({})で段落を開始
            self._start_paragraph({})

    def _append_content(self, content_item: Dict[str, Any]):
        """現在のブロック (self._current_block['content']) にインライン要素を追加します。"""
        if not self._current_block:
            logger.warning(
                "Attempted to append content, but no current block exists. Ensuring default paragraph.")
            self._ensure_paragraph_started()
            if not self._current_block:
                logger.error(
                    "Failed to ensure paragraph start. Cannot append content.")
                return

        self._current_block['content'].append(content_item)

    # --- ハンドラメソッド ---

    def _handle_text(self, run: Dict[str, Any]):
        """
        'text' トークンをスタイル付きの 'span' オブジェクトに変換します。
        [改善案 5] text 内の \n はテンプレート側で <br> に変換 *しない*。
        パーサーが 'br' トークンを生成することを期待する。
        """
        # 段落先頭のみ、先頭空白（半角/ノーブレーク/全角）を除去
        _text = run.get('text', '')
        try:
            if self._current_block is not None and not self._current_block.get('content'):
                import re as _re
                _text = _re.sub(r'^[ \u00A0\u3000]+', '', _text)
        except Exception:
            pass

        # テキストが空（または空白除去で空）になった場合は何も追加しない
        if not _text:
            return

        style_css, tags = self._style_from_run(run)
        self._append_content({
            "type": "span",
            "text": _text,
            "style": style_css,  # インライン CSS
            "tags": tags       # ラップするタグ (例: ['b', 'i'])
        })

    # def _handle_link(self, run: Dict[str, Any]):
    #     """'link' トークンを 'a' オブジェクトに変換します。"""
    #     style_css, tags = self._style_from_run(run)
    #     label = run.get('label', run.get('href', ''))
    #     self._append_content({
    #         "type": "a",
    #         "href": run.get('href', '#'),
    #         "label": label,
    #         "style": style_css,
    #         "tags": tags
    #     })

        # --- 追加: リダイレクト基底URLを取得 ---
    def _get_redirect_base(self) -> Optional[str]:
        """
        リダイレクト基底URLを返す。
        優先度: 環境変数 -> doc.meta -> 無し
        """

        env = os.getenv("NOTES_REDIRECT_BASE")
        if env:
            return env.rstrip('&')
        meta = (self.doc or {}).get("meta", {})
        base = meta.get("notes_redirect_base")  # 任意: メタに入っていれば使用
        return base.rstrip('&') if isinstance(base, str) else None

    # --- 追加: Notes:// の生URLを構築 ---
    def _build_notesurl(self, notes: Dict[str, Any]) -> Optional[str]:
        try:
            server = notes.get("server")
            replica = notes.get("replica")
            unid = notes.get("unid")
            view = notes.get("view")  # 任意
            if not (server and replica and unid):
                return None

            import re
            m = re.match(r"(?:CN=)?([^/]+)", server, re.IGNORECASE)
            if m:
                server = m.group(1)
            
            domain_suffix = os.getenv("NOTES_DOMAIN_SUFFIX", "")
            if domain_suffix and "." not in server:
                if not domain_suffix.startswith("."):
                    domain_suffix = "." + domain_suffix
                server = f"{server}{domain_suffix}"

            path = f"{server}/{replica}"
            if view:
                path += f"/{view}"
            path += f"/{unid}"
            return f"Notes://{path}"
        except Exception:
            return None

    # --- 追加: リダイレクト付の最終 href を構築 ---
    def _build_redirect_url(self, notes: Dict[str, Any]) -> Optional[str]:
        base = self._get_redirect_base()
        notesurl = self._build_notesurl(notes)
        if not notesurl:
            return None
        if base:
            # 例: http://.../Index?OpenPage&NotesURL=<URLENCODED>
            joiner = '&' if '?' in base else '?'
            if base.endswith('NotesURL='):
                return f"{base}{quote(notesurl, safe='')}"
            return f"{base}{joiner}NotesURL={quote(notesurl, safe='')}"
        # 基底が無い場合は Notes:// をそのまま返す（社内ハンドラに委ねる）
        return notesurl

    # --- 追加: Box検索URLビルダ ---
    def _get_box_search_base(self) -> str:
        # 末尾に query= を含んでいる形を想定。含まれていなくても ?query= を付与して使う。
        return os.getenv("BOX_SEARCH_BASE", "https://app.box.com/folder/0/search?query=")

    def _build_box_search_url(self, query_text: str) -> str:
        base = self._get_box_search_base()
        if "query=" not in base:
            sep = '&' if '?' in base else '?'
            base = f"{base}{sep}query="
        return f"{base}{quote(query_text or '', safe='')}"

    # --- 追加: Box検索用のクエリ整形関数 ---
    def _extract_query_from_label(self, label: str) -> str:
        """
        Notesリンクのlabelから検索用文字列を抽出。
        例:
          "Database '生産部門 掲示板', View '生産部門会議', Document 'いすゞディーゼルエンジンの歴史'"
          → "いすゞディーゼルエンジンの歴史 生産部門 掲示板 生産部門会議"
        """
        if not label:
            return ""
        # Database, View, Document の順に正規表現で抽出
        # db = re.search(r"Database\s+'([^']+)'", label)
        # view = re.search(r"View\s+'([^']+)'", label)
        doc = re.search(r"Document\s+'([^']+)'", label)

        # ドキュメント名 → データベース名 → ビュー名 の順で結合
        parts = []
        if doc:
            parts.append(doc.group(1))
        # if db:
        #     parts.append(db.group(1))
        # if view:
        #     parts.append(view.group(1))

        return "".join(parts) if parts else label

    def _handle_link(self, run: Dict[str, Any]):
        """'link' トークン -> a要素。Notesリンクのみ Box検索を追加。すべて別タブで開く。"""
        # --- 修正版: _handle_link ---    既存: 'link' の拡張（2本目の「Box検索」を追加）
        style_css, tags = self._style_from_run(run)
        # 元の href を保持して「Notes判定」に使う
        original_href = run.get('href', '') or ''
        notes_meta = run.get('notes') or {}

        # --- Notes判定：notes メタがある or href が Notes://
        is_notes_link = bool(notes_meta) or (isinstance(
            original_href, str) and original_href.startswith("Notes://"))

        label = run.get('label', original_href)
        # カンマ区切りのままだとHTML上で下線が途切れるため、Notes関連リンクではスペースに置換する
        # また、ユーザーの要望によりBox検索と同様のアイコンを先頭に付与
        if is_notes_link and isinstance(label, str):
            label = "📝" + label.replace(", ", " ")

        # --- NotesリダイレクトURLの構築（既存ロジック）
        href = original_href
        if notes_meta:
            _redirect = self._build_redirect_url(notes_meta)
            if _redirect:
                href = _redirect
        elif isinstance(href, str) and href.startswith("Notes://"):
            base = self._get_redirect_base()
            if base:
                if base.endswith('NotesURL='):
                    href = f"{base}{quote(href, safe='')}"
                else:
                    joiner = '&' if '?' in base else '?'
                    href = f"{base}{joiner}NotesURL={quote(href, safe='')}"

        # --- 1本目：通常リンク（別タブ属性を付与）
        self._append_content({
            "type": "a",
            "href": href if href else '#',
            "label": label,
            "style": style_css,
            "tags": tags,
            "attributes": self._default_link_attrs()
        })

        # --- 2本目：Box検索リンク（Notesリンクのときだけ）
        if is_notes_link:
            try:
                subject = self.output.get("subject")
            except Exception:
                subject = None

            query_text = ""
            if run.get("label"):
                query_text = self._extract_query_from_label(run["label"])
            elif subject:
                query_text = subject

            if query_text:
                self._append_content(
                    {"type": "span", "text": " ", "style": "", "tags": []})
                self._append_content({
                    "type": "a",
                    "href": self._build_box_search_url(query_text),
                    "label": "🔎Box検索",
                    "style": "",
                    "tags": [],
                    "attributes": {**self._default_link_attrs(), "style": "margin-left: 0.75rem; text-decoration: none;"}
                })

    # def _handle_link(self, run: Dict[str, Any]):
    #     """'link' トークンを 'a' オブジェクトに変換 (Notesリンク対応)"""
    #     style_css, tags = self._style_from_run(run)
    #     label = run.get('label', run.get('href', ''))

    #     href = run.get('href', '')
    #     notes_meta = run.get('notes') or {}
    #     if notes_meta:
    #         # NJSONに Notes メタがある場合：リダイレクトURLを生成
    #         redirect = self._build_redirect_url(notes_meta)
    #         if redirect:
    #             href = redirect
    #     elif isinstance(href, str) and href.startswith("Notes://"):
    #         # href が Notes:// の場合もリダイレクトに載せ替え
    #         # （view/unid を取り出せないケースはそのまま base にぶら下げ）
    #         base = self._get_redirect_base()
    #         if base:
    #             href = f"{base}{quote(href, safe='')}" if base.endswith('NotesURL=') \
    #                 else f"{base}{('&' if '?' in base else '?')}NotesURL={quote(href, safe='')}"

    #     self._append_content({
    #         "type": "a",
    #         "href": href if href else '#',
    #         "label": label,
    #         "style": style_css,
    #         "tags": tags
    #     })

    def _handle_img(self, run: Dict[str, Any]):
        """
        'img' トークンを 'img' オブジェクトに変換します。
        v1.4 対応: 'content_path' ではなく 'src' を参照します。
        """
        # [修正 3] content_path -> src に変更
        content_path = run.get('src')
        alt_text = run.get('alt', 'image')

        src = content_path if content_path else '#'
        if src == '#':
            # [修正 3] 警告メッセージも 'src' に合わせる
            logger.warning(f"Image run has no 'src'. Alt: {alt_text}")

        self._append_content({
            "type": "img",
            "src": src,
            "alt": alt_text
        })

    def _handle_table(self, run: Dict[str, Any]):
        """
        'table' トークンを 'table' オブジェクト (再帰) に変換し、
        body_elements に追加します。
        [v1.5] _process_runs から呼ばれる (finalize不要)
        """
        table_attributes = run.get('attributes', {})
        table_style_css = self._style_from_table_attrs(table_attributes)

        table_block = {
            "type": "table",
            "style": table_style_css,
            "attributes": table_attributes,  # [修正 8] タブ表判定のため属性全体を渡す
            "columns": run.get('columns', []),  # [修正 8] 列幅定義を渡す
            "rows": [],
            "tab_group_id": None  # [改善案 7] タブ表用ID
        }

        # [改善案 7] タブ付き表の場合、一意なIDを生成
        if table_attributes.get('rowdisplay') == "tabs":
            self._tab_group_counter += 1
            table_block['tab_group_id'] = f"tg-{self._tab_group_counter}"
            logger.debug(
                f"Tabbed table detected. Assigning group ID: {table_block['tab_group_id']}")

        logger.debug("Starting table block.")

        for i, row_data in enumerate(run.get('rows', [])):
            row_attributes = row_data.get('attributes', {})
            row_block = {
                "type": "tr",
                # [v1.5] tablabel は row の attributes に入っている
                "label": row_attributes.get('tablabel'),
                "attributes": {},
                "cells": []
            }

            for j, cell_data in enumerate(row_data.get('cells', [])):
                # [v1.5] v1.5仕様では cell_data は style と runs を持つ
                cell_runs = cell_data.get('runs', [])
                # colspan, rowspan, bgcolor など
                cell_style_attrs = (cell_data.get('style', {}) or {}).copy()

                # [v1.6.2] 堅牢性向上: top-level の span 情報を style_attrs にマージ
                if 'rowspan' in cell_data:
                    cell_style_attrs['rowspan'] = cell_data['rowspan']
                if 'colspan' in cell_data:
                    cell_style_attrs['colspan'] = cell_data['colspan']
                # [v1.6.2] パーサーのタイポ(columnspan)にも top-level で対応
                if 'columnspan' in cell_data:
                    cell_style_attrs['columnspan'] = cell_data['columnspan']

                cell_render_attributes = self._style_from_cell_style(
                    cell_style_attrs)

                # サブレンダラーを作成 (親の doc, jinja_env, icon_map, tab_counter を共有)
                cell_renderer = HtmlRenderer._create_sub_renderer(self)

                logger.debug(f"Processing runs for table cell ({i},{j})...")
                cell_renderer._process_runs(cell_runs)

                # [改善案 7] サブレンダラーでカウンターがインクリメントされた場合、
                # 親レンダラーに反映
                self._tab_group_counter = cell_renderer._tab_group_counter

                cell_content_blocks = cell_renderer.output['body_elements']
                logger.debug(
                    f"Cell ({i},{j}) contains {len(cell_content_blocks)} blocks.")

                cell_block = {
                    "type": "td",
                    "content_blocks": cell_content_blocks,
                    "attributes": cell_render_attributes
                }
                row_block['cells'].append(cell_block)
            table_block['rows'].append(row_block)

        self.output['body_elements'].append(table_block)
        logger.debug("Finished table block.")

    def _handle_attachmentref(self, run: Dict[str, Any]):
        """
        'attachmentref' を 'attachment_block' (pタグでラップ) オブジェクトに変換します。
        [v1.5] _process_runs から呼ばれる (finalize不要)
        """
        content_path = run.get('content_path')
        name = run.get('name', 'file')
        display_name = run.get('displayname', name)
        href = content_path if content_path else '#'

        icon_path = self._find_icon_path(content_path, name)

        block_style = ""  # <p> タグのスタイル (必要なら)

        if href == '#':
            logger.warning(
                f"AttachmentRef run has no 'content_path'. Name: {name}")
            content_item = {
                "type": "span",
                "text": f"[添付: {display_name}]",
                "style": "font-style: italic; color: #555;",
                "tags": []
            }
        else:
            # [修正 11] attachmentref はリンク(a)とアイコン(img)の組み合わせとして content に渡す
            content_item = {
                "type": "attachment_link",  # Jinja2で特別扱いする新しいタイプ
                "href": href,
                "label": display_name,
                "icon_src": icon_path,
                "style": "",  # リンク自体のスタイルはCSSクラス (attachment-link) で制御
                "tags": [],
                # "attributes": {"download": name}
                # download 既存に target/rel を追加（download優先だが target 付与で統一）
                "attributes": {"download": name, **self._default_link_attrs()}
            }

        # [修正 13] 直前の空ブロックの影響を受けないよう、コンテキストから最新のスタイルを直接引継ぎ、独立した段落を作成する
        inherited_class = "attachment-paragraph"
        indent_cls = self._indent_class_from_par(self.context.current_par_style).strip()
        # [修正 14] 直前のパラグラフがインデントを持たない（Notesが直前に空のparを出す）場合、最後に有効だったインデントにフォールバック
        if not indent_cls and self._last_meaningful_indent_cls:
            indent_cls = self._last_meaningful_indent_cls

        if indent_cls:
            inherited_class += f" {indent_cls}"

        attachment_para_block = {
            "type": "p",
            "content": [content_item],
            "style": block_style,
            "attributes": {"class": inherited_class}
        }
        self.output['body_elements'].append(attachment_para_block)
        logger.debug(f"Added attachment paragraph block with class '{inherited_class}': {display_name}")

    def _handle_br(self, run: Dict[str, Any]):
        """'br' トークンを 'br' オブジェクトに変換します。(v1.5定義明確化)"""
        self._append_content({"type": "br"})

    # --- [修正 6] v1.5新設 ---
    def _handle_hr(self, run: Dict[str, Any]):
        """'hr' トークン (水平線) を 'hr' オブジェクトに変換します。(v1.5新設)"""
        # _process_runs 側で _finalize_paragraph() が呼ばれる
        attributes = run.get('a', {})
        style_css = []
        if attributes.get('height'):
            style_css.append(f"height: {attributes['height']};")
        if attributes.get('width'):
            # width: 100% のような指定に対応
            css_width = attributes['width']
            if css_width.replace('.', '', 1).isdigit():  # 単なる数字(twips?)は無視
                logger.debug(f"Ignoring numeric-only hr width: {css_width}")
            else:
                style_css.append(f"width: {css_width};")

        if attributes.get('use3dshading') == 'true':
            style_css.append("border-style: inset; border-width: 1px;")
        else:
            # 3Dシェーディングなし、または指定なしの場合
            style_css.append(
                "border-style: solid; border-width: 1px; color: #ccc;")

        hr_block = {
            "type": "hr",
            "style": " ".join(style_css)
        }
        self.output['body_elements'].append(hr_block)
        logger.debug("Added horizontal rule (hr) block.")

    # --- [修正 7] v1.5新設 ---
    def _handle_section(self, run: Dict[str, Any]):
        """
        'section' トークン (折りたたみセクション) を 'section' オブジェクト (再帰) に変換し、
        body_elements に追加します。(v1.5新設)
        """
        # _process_runs 側で _finalize_paragraph() が呼ばれる
        title_runs = run.get('title_runs', [])
        body_runs = run.get('body_runs', [])

        # --- タイトルを処理 ---
        title_renderer = HtmlRenderer._create_sub_renderer(self)
        logger.debug("Processing runs for section title...")
        title_renderer._process_runs(title_runs)
        # [改善案 7] サブレンダラーのカウンターを反映
        self._tab_group_counter = title_renderer._tab_group_counter
        title_content_blocks = title_renderer.output['body_elements']
        logger.debug(
            f"Section title contains {len(title_content_blocks)} blocks.")

        # --- 本文を処理 ---
        body_renderer = HtmlRenderer._create_sub_renderer(self)
        logger.debug("Processing runs for section body...")
        body_renderer._process_runs(body_runs)
        # [改善案 7] サブレンダラーのカウンターを反映
        self._tab_group_counter = body_renderer._tab_group_counter
        body_content_blocks = body_renderer.output['body_elements']
        logger.debug(
            f"Section body contains {len(body_content_blocks)} blocks.")

        section_block = {
            "type": "section",
            "title_blocks": title_content_blocks,
            "body_blocks": body_content_blocks,
            "attributes": run.get('attributes', {}),
            "style": ""  # <details> タグのスタイル (必要なら)
        }
        self.output['body_elements'].append(section_block)
        logger.debug("Finished section block.")

    def _handle_unknown(self, run: Dict[str, Any]):
        """不明なトークンをデバッグ用 'span' オブジェクトに変換します。"""
        self._append_content({
            "type": "span",
            "text": f"[Unknown Run: {run.get('t')}]",
            "style": "color: red; background-color: yellow; font-weight: bold; ",
            "tags": []
        })

    # --- [修正 4.A] アイコンSVGマッピング (インライン埋め込み) ---
    @staticmethod
    def _svg_to_data_uri(svg_string: str) -> str:
        import base64
        encoded = base64.b64encode(svg_string.strip().encode('utf-8')).decode('ascii')
        return f"data:image/svg+xml;base64,{encoded}"

    # Fluent Design風のリッチで立体的なSVG定義
    _SVG_ICONS = {
        "excel": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgExcel" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#21A366"/>
      <stop offset="100%" stop-color="#107C41"/>
    </linearGradient>
  </defs>
  <rect x="4" y="2" width="24" height="28" rx="3" fill="url(#bgExcel)"/>
  <path d="M14 8h10v16H14z" fill="#fff" opacity="0.9"/>
  <path d="M16 11h6v2h-6zm0 4h6v2h-6zm0 4h6v2h-6z" fill="#107C41"/>
  <rect x="2" y="8" width="14" height="14" rx="2" fill="#185C37"/>
  <text x="9" y="19" font-family="'Segoe UI', Arial, sans-serif" font-weight="700" font-size="11" fill="#fff" text-anchor="middle">X</text>
</svg>"""),
        "word": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgWord" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#2B579A"/>
      <stop offset="100%" stop-color="#143A75"/>
    </linearGradient>
  </defs>
  <rect x="4" y="2" width="24" height="28" rx="3" fill="url(#bgWord)"/>
  <path d="M14 8h10v16H14z" fill="#fff" opacity="0.9"/>
  <path d="M16 11h6v2h-6zm0 4h6v2h-6zm0 4h4v2h-4z" fill="#143A75"/>
  <rect x="2" y="8" width="14" height="14" rx="2" fill="#0D254F"/>
  <text x="9" y="19" font-family="'Segoe UI', Arial, sans-serif" font-weight="700" font-size="11" fill="#fff" text-anchor="middle">W</text>
</svg>"""),
        "ppt": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgPpt" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#D24726"/>
      <stop offset="100%" stop-color="#A5351A"/>
    </linearGradient>
  </defs>
  <rect x="4" y="2" width="24" height="28" rx="3" fill="url(#bgPpt)"/>
  <path d="M14 8h10v16H14z" fill="#fff" opacity="0.9"/>
  <path d="M16 10h6v4h-6z" fill="#A5351A"/>
  <path d="M16 15h4v1h-4zm0 2h4v1h-4z" fill="#A5351A"/>
  <rect x="2" y="8" width="14" height="14" rx="2" fill="#75220F"/>
  <text x="9" y="19" font-family="'Segoe UI', Arial, sans-serif" font-weight="700" font-size="11" fill="#fff" text-anchor="middle">P</text>
</svg>"""),
        "pdf": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgPdf" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#F40F02"/>
      <stop offset="100%" stop-color="#B20B00"/>
    </linearGradient>
  </defs>
  <rect x="4" y="2" width="24" height="28" rx="3" fill="url(#bgPdf)"/>
  <path d="M6 10h20v14H6z" fill="#fff" opacity="0.95"/>
  <text x="16" y="20.5" font-family="'Segoe UI', Arial, sans-serif" font-weight="800" font-size="9" fill="#B20B00" text-anchor="middle" letter-spacing="0.5">PDF</text>
</svg>"""),
        "zip": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgZip" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#FBC02D"/>
      <stop offset="100%" stop-color="#F57F17"/>
    </linearGradient>
    <linearGradient id="bgZipTab" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#FFF176"/>
      <stop offset="100%" stop-color="#FBC02D"/>
    </linearGradient>
  </defs>
  <path d="M4 10v18h24V10z" fill="url(#bgZip)"/>
  <path d="M4 10V6c0-1.1.9-2 2-2h8l2 2h8c1.1 0 2 .9 2 2v4z" fill="url(#bgZipTab)"/>
  <rect x="14" y="10" width="4" height="12" fill="#E65100" opacity="0.3"/>
  <rect x="14" y="10" width="4" height="1.5" fill="#4B4B4B"/>
  <rect x="14" y="13" width="4" height="1.5" fill="#4B4B4B"/>
  <rect x="14" y="16" width="4" height="1.5" fill="#4B4B4B"/>
  <rect x="13.5" y="18" width="5" height="4" rx="1" fill="#B0BEC5"/>
  <path d="M14 20h4v1h-4z" fill="#78909C"/>
</svg>"""),
        "image": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgImg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#4FC3F7"/>
      <stop offset="100%" stop-color="#0288D1"/>
    </linearGradient>
  </defs>
  <rect x="2" y="4" width="28" height="24" rx="3" fill="url(#bgImg)"/>
  <circle cx="10" cy="11" r="3" fill="#fff" opacity="0.9"/>
  <path d="M2 24l8-10 6 8 4-5 10 11v-4H2z" fill="#fff" opacity="0.6"/>
  <path d="M2 28v-4l8-10 6 8 4-5 10 11v3a3 3 0 01-3 3H5a3 3 0 01-3-3z" fill="#01579B" opacity="0.4"/>
</svg>"""),
        "default": _svg_to_data_uri("""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <defs>
    <linearGradient id="bgFile" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#E0E0E0"/>
      <stop offset="100%" stop-color="#9E9E9E"/>
    </linearGradient>
  </defs>
  <path d="M6 3h12l8 8v17a1 1 0 01-1 1H6a1 1 0 01-1-1V4a1 1 0 011-1z" fill="url(#bgFile)"/>
  <path d="M18 3v8h8z" fill="#F5F5F5" opacity="0.8"/>
  <path d="M18 11h8l-8-8v8z" fill="#757575" opacity="0.2"/>
  <rect x="10" y="16" width="12" height="2" rx="1" fill="#757575" opacity="0.6"/>
  <rect x="10" y="20" width="12" height="2" rx="1" fill="#757575" opacity="0.6"/>
  <rect x="10" y="24" width="8" height="2" rx="1" fill="#757575" opacity="0.6"/>
</svg>""")
    }

    @classmethod
    def _map_extension_to_svg(cls, filename: str) -> str:
        """ファイル名拡張子から直接SVG Data URIを返す"""
        if not filename:
            return cls._SVG_ICONS["default"]
        ext_dot = filename.rfind('.')
        if ext_dot == -1:
            return cls._SVG_ICONS["default"]
        
        ext = filename[ext_dot+1:].lower()
        if ext in ['xls', 'xlsx', 'xlsm', 'csv']:
            return cls._SVG_ICONS["excel"]
        elif ext in ['ppt', 'pptx', 'pptm']:
            return cls._SVG_ICONS["ppt"]
        elif ext in ['doc', 'docx', 'docm', 'rtf']:
            return cls._SVG_ICONS["word"]
        elif ext == 'pdf':
            return cls._SVG_ICONS["pdf"]
        elif ext in ['zip', 'rar', '7z', 'tar', 'gz']:
            return cls._SVG_ICONS["zip"]
        elif ext in ['jpg', 'jpeg', 'png', 'gif', 'bmp']:
            return cls._SVG_ICONS["image"]
        
        return cls._SVG_ICONS["default"]

    def _build_icon_map(self) -> Dict[str, str]:
        """attachments 配列から content_path -> icon_path の辞書を作成"""
        icon_map = {}
        try:
            for att in self.doc.get("attachments", []):
                if att.get("type") == "file" and att.get("content_path"):
                    content_path = att["content_path"]
                    name = att.get("name", "")
                    
                    # 強制的にファイル名拡張子からSVG Data URIを生成して割り当てる
                    # 外部GIFへの依存や間違った名前の焼き込み問題を完全に回避
                    icon_map[content_path] = self._map_extension_to_svg(name)
        except Exception as e:
            logger.warning(
                f"Failed to build attachment icon map: {e}", exc_info=True)
        return icon_map

    # --- [修正 4.B] v1.4対応: アイコンパス検索ヘルパー ---
    def _find_icon_path(self, content_path: Optional[str], name: str) -> str:
        """content_path (優先) または name から SVG Data URI を特定する"""
        
        if content_path and content_path in self._attachment_icon_map:
            return self._attachment_icon_map[content_path]

        # マップにない場合もファイル名からSVG Data URIを生成
        return self._map_extension_to_svg(name)

    # --- ヘルパーメソッド ---

    def _style_from_par(self, par_style: Dict[str, Any]) -> str:
        """par トークンの 'a' 属性から CSS インラインスタイル文字列を生成"""
        css = []
        align = par_style.get('align')
        align_map = {'left': 'left', 'center': 'center',
                     'right': 'right', 'full': 'justify', 'justify': 'justify'}
        if align and align in align_map:
            css.append(f"text-align: {align_map[align]};")

        left_margin = par_style.get('leftmargin')
        # [patched] leftmargin is mapped to indent classes; no inline margin-left.
        # (indent class assignment happens in _start_paragraph)

        space_after = par_style.get('spaceafter')
        if space_after:
            try:
                # v1.5 DOCX ログで '1.5' が警告されているため、float に対応
                margin_bottom = float(space_after)
                # Notes の 1.5 (行間) は 1.0 を引いて調整 (経験則)
                css.append(
                    f"margin-bottom: {max(0, margin_bottom - 1.0) * 0.8}em;")
            except (ValueError, TypeError):
                logger.warning(
                    f"Could not parse spaceafter value: {space_after}")

        return " ".join([c for c in css if isinstance(c, str) and c.strip()]).strip()

    def _style_from_run(self, run: Dict[str, Any]) -> Tuple[str, List[str]]:
        """
        run トークンの 's' と 'a' から CSS 文字列とラッパータグリストを生成。
        Jinja2 テンプレート側でタグをラップすることを想定。
        """
        s = run.get('s', [])
        a = run.get('a', {})
        css = []
        tags = []

        if 'b' in s:
            tags.append('strong')
        if 'i' in s:
            tags.append('em')
        if 'u' in s:
            tags.append('u')
        if 's' in s or 'strike' in s:
            tags.append('s')

        if a.get('color'):
            css.append(f"color: {a['color']};")
        if a.get('bgcolor'):
            css.append(f"background-color: {a['bgcolor']};")
        if a.get('size'):
            css.append(f"font-size: {a['size']};")

        script = a.get('script') or next(
            (fx for fx in a.get('fx', []) if fx in ['super', 'sub']), None)
        if script == 'super':
            tags.append('sup')
        elif script == 'sub':
            tags.append('sub')

        fx = a.get('fx', [])
        if 'shadow' in fx:
            css.append("text-shadow: 1px 1px 2px rgba(0,0,0,0.5);")
        if 'emboss' in fx:
            css.append(
                "text-shadow: -1px -1px 1px rgba(255,255,255,0.8), 1px 1px 1px rgba(0,0,0,0.2);")

        # --- [修正 1] v1.4 fx:extrude 対応 ---
        if 'extrude' in fx:
            # 簡易的な押し出し効果
            css.append(
                "text-shadow: 1px 1px 0px rgba(0,0,0,0.4), 2px 2px 0px rgba(0,0,0,0.3);")
        # --------------------------------------

        if 'mono' in s:
            css.append("font-family: Consolas, 'Courier New', monospace;")

        return " ".join([c for c in css if isinstance(c, str) and c.strip()]).strip(), tags

    def _get_ol_type_attr(self, list_type: Optional[str]) -> Optional[str]:
        """Notes のリストタイプを HTML の <ol type="..."> 属性にマッピング"""
        if list_type == 'number':
            return '1'
        if list_type == 'alphaupper':
            return 'A'
        if list_type == 'alphalower':
            return 'a'
        if list_type == 'romanupper':
            return 'I'
        if list_type == 'romanlower':
            return 'i'
        return None

    def _style_from_table_attrs(self, table_attributes: Dict[str, Any]) -> str:
        """<div class="table-wrap">\1 タグの attributes から CSS 文字列を生成"""
        css = ["border-collapse: collapse;",
               "margin-bottom: 1em;"]
        width_type = table_attributes.get('widthtype')
        ref_width = table_attributes.get('refwidth')  # [修正 9]

        if width_type == 'fitmargins':
            css.append("width: 100%;")
        elif width_type == 'fixedleft' or width_type == 'fixed':
            # [修正 9] refwidth があればそれを優先
            if ref_width:
                css.append(f"width: {ref_width};")
            elif width_type == 'fixedleft':
                css.append("width: auto;")
            else:  # fixed (refwidth なし)
                css.append("width: 100%;")

            # [修正 9] 列幅指定がある(columnsが使われる)ことを想定し、table-layout: fixed を追加
            css.append("table-layout: fixed;")

        elif width_type == 'fixedleft':
            # (v1.4互換) width_type が 'fixedleft' (refwidthなし) の場合のフォールバック
            css.append("width: auto;")
        return " ".join(css)

    def _style_from_cell_style(self, cell_style_attrs: Dict[str, Any]) -> Dict[str, Any]:
        """<td>, <th> の style属性 (colspan/rowspan含む) から属性辞書を生成"""
        attributes = {}
        css = []  # start empty; avoid stray spaces

        # [v1.6.1 修正]
        # JSONパーサーが "colspan" ではなく "columnspan" を
        # 出力する場合があるため、両方に対応する。
        colspan = cell_style_attrs.get(
            'colspan') or cell_style_attrs.get('columnspan')
        rowspan = cell_style_attrs.get('rowspan')

        if colspan:
            try:
                if int(colspan) > 1:
                    attributes['colspan'] = str(colspan)
            except (ValueError, TypeError):
                logger.warning(f"Invalid non-numeric colspan value: {colspan}")

        if rowspan:
            try:
                if int(rowspan) > 1:
                    attributes['rowspan'] = str(rowspan)
            except (ValueError, TypeError):
                logger.warning(f"Invalid non-numeric rowspan value: {rowspan}")

        # [v1.5 互換性修正]
        # v1.5 JSON: "style": { "bgcolor": "lime", "colspan": 1, ... }
        # v1.4 JSON: "style": { "colspan": 1, "style": { "bgcolor": "teal" } }

        # v1.5 仕様 (style直下) を優先
        if cell_style_attrs.get('bgcolor'):
            css.append(f"background-color: {cell_style_attrs['bgcolor']};")
        else:
            # v1.4 互換 (style.style.bgcolor)
            style_inner = cell_style_attrs.get('style', {})
            if style_inner.get('bgcolor'):
                css.append(f"background-color: {style_inner['bgcolor']};")

        css_str = " ".join(css).strip()
        if css_str:
            attributes['style'] = css_str
        return attributes

    def _build_appendix_rows_helper(self) -> List[Dict[str, Any]]:
        """
        付録テーブル用の行データを構築します。
        BaseRenderer の self.doc を参照します。
        """
        rows = []
        fields = self.doc.get("fields", {}) or {}
        layout = self.doc.get("layout", {}) or {}
        allow = set(layout.get("primary_fields_allowlist", []) or [])
        used = set(layout.get("used_in_body", []) or [])
        logger.debug(f"Appendix: allowlist={allow}, used_in_body={used}")

        # 無意味なバイナリや文字化けの原因となる、システム・セキュリティ系のフィールドを除外
        ignore_fields = {
            "LastScannedVersionCheck",
            "$UpdatedBy", "$Revisions", "$MessageID",
            "$KeepPrivate", "sign", "$Writers", "Categories"
        }

        processed_names = set()

        for name, meta in fields.items():
            if name in allow or name in used or name in processed_names or name in ignore_fields or name.startswith("$"):
                continue

            ftype = meta.get("type", "unknown")
            val = meta.get("value")
            preview = ""
            try:
                if isinstance(val, list):
                    if val and all(isinstance(x, (str, int, float, bool)) for x in val):
                        preview = ", ".join(map(str, val[:5]))
                        if len(val) > 5:
                            preview += " ..."
                    else:
                        preview = f"[List of {len(val)} items]"
                elif isinstance(val, dict):
                    preview = "{...}"
                elif val is None:
                    preview = "[None]"
                else:
                    preview = str(val)

                if len(preview) > 100:
                    preview = preview[:100] + " ..."
                preview = preview.replace('&', '&amp;').replace(
                    '<', '&lt;').replace('>', '&gt;')
            except Exception as e:
                preview = f"[Error previewing: {e}]"
                logger.warning(
                    f"Error generating preview for field '{name}': {e}")

            rows.append({"name": name, "type": ftype, "preview": preview})
            processed_names.add(name)

        rows.sort(key=lambda r: r["name"].lower())
        logger.debug(f"Generated {len(rows)} appendix rows.")
        return rows

    def get_output_on_error(self, e: Exception) -> str:
        """
        レンダリングエラー発生時にエラー情報を HTML で返す
        (get_output() は呼び出さず、安全にHTMLを構築する)
        """
        logger.error(f"Render error occurred: {e}", exc_info=True)

        try:
            error_subject = f"Render Error: {e}"
            meta_preview = self.output.get(
                "meta", {"error": "Meta unavailable"})
            context_preview = {k: v for k, v in self.output.items(
            ) if k not in ['font_url', 'body_elements', 'appendix_rows', 'meta']}
        except Exception:
            # self.output が壊れている場合
            meta_preview = {"error": "Context unavailable"}
            context_preview = {"error": "Context unavailable"}

        # [提案2] エラーHTMLも BeautifulSoup で美化する (読みやすさのため)
        error_html = f"""<!DOCTYPE html>
            <html lang="en">
            <head><title>{error_subject}</title>
            <style>body {{ font-family: sans-serif; margin: 2em; }} pre {{ background: #f0f0f0; padding: 1em; border-radius: 5px; }}</style>
            </head>
            <body>
                <h1>Render Error</h1>
                <p>An error occurred while rendering the document:</p>
                <pre>{type(e).__name__}: {e}</pre>
                <hr>
                <h2>Document Meta (partial)</h2>
                <pre>{meta_preview}</pre>
                <h2>Context Data (partial)</h2>
                <pre>{context_preview}</pre>
            </body>
            </html>"""

        if _BS4_AVAILABLE and BeautifulSoup:
            try:
                soup = BeautifulSoup(error_html, "html.parser")
                return soup.prettify(formatter="html5")
            except Exception:
                pass  # フォールバック
        return error_html


# --- [提案2] 簡易テストコード (pretty=True と False の両方をテスト) ---
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

    # v1.5 のテスト JSON を使用 (パスを v1.4.4 -> v1.5-dev に想定変更)
    test_json_path = Path(
        r"C:\Users\A512292\py2024\notes_to_box_v4_2_0\dist\output\manifest_run\020\TEMP_DB\Document\その他\Doc_20251023_TEST4\Doc_20251023_TEST4.normalized.json")

    # もし v1.5-dev がなければ、アップロードされたファイル名 (Doc_20251023_TEST4.normalized.json) を
    # ローカルのテストパスとして想定 (このスクリプトと同じディレクトリにあると仮定)
    if not test_json_path.exists():
        test_json_path_local = Path(
            __file__).parent / "Doc_20251023_TEST4.normalized.json"
        if test_json_path_local.exists():
            test_json_path = test_json_path_local
        else:
            # v1.4.4 のパスに戻す (フォールバック)
            test_json_path = Path(
                r"C:\Users\A512292\py2024\notes_to_box_v4_2_0\dist\output\manifest_run\012\_dxl\TEMP_DB\670309E2B88C5E9649258D2C000ADBDE.normalized.1.4.4-dev_initial.json")

    if test_json_path.exists():
        print(f"--- Running HTML Render Test (v1.6) for: {test_json_path} ---")
        try:
            renderer = HtmlRenderer(test_json_path)
            logger.info(
                f"Using default template dir for test: {renderer.template_dir}")

            # --- 1. 最小化版 (pretty=False) ---
            output_html_minified = renderer.render(pretty=False)
            output_html_path_min = test_json_path.parent / \
                (test_json_path.stem + "_render_test_v1.6_MINIFIED.html")
            with open(output_html_path_min, "w", encoding='utf-8') as f:
                f.write(output_html_minified)
            print(
                f"+++ Minified HTML render successful. Output: {output_html_path_min} +++")

            # --- 2. 美化版 (デフォルト) ---
            if _BS4_AVAILABLE:
                # [v1.6.2] render() のデフォルトが pretty=True になった
                output_html_pretty = renderer.render()
                output_html_path_pretty = test_json_path.parent / \
                    (test_json_path.stem + "_render_test_v1.6_PRETTY.html")
                with open(output_html_path_pretty, "w", encoding='utf-8') as f:
                    f.write(output_html_pretty)
                print(
                    f"+++ Pretty HTML render successful (default). Output: {output_html_path_pretty} +++")
            else:
                print(
                    "--- Skipping Pretty HTML render (beautifulsoup4 not installed) ---")

            if renderer.output.get("appendix_rows"):
                print("\n--- Appendix Preview (first 5 rows) ---")
                for row in renderer.output["appendix_rows"][:5]:
                    print(f"- {row['name']} ({row['type']}): {row['preview']}")
                if len(renderer.output["appendix_rows"]) > 5:
                    print("...")

        except Exception as main_e:
            print(f"--- HTML Render Test FAILED (v1.6) ---")
            logger.exception("Error during HTML render test")

    else:
        print(f"!!! Test JSON file not found: {test_json_path} !!!")
        print("Please create a test JSON file or adjust the path in html.py.")
