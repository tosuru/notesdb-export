"""
レンダリングエンジン共通基盤 (v1.1)
BaseRenderer と RenderContext を定義します。

v1.1:
- BaseRenderer のコンストラクタ引数を doc_path (Path) に変更。
- JSON ファイルのロードと attachment_dir の解決を BaseRenderer 内で行う。
- resolve_attachment_path ヘルパーメソッドを追加。
"""

from __future__ import annotations
import json
import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, List, Optional

# ロガー設定
logger = logging.getLogger(__name__)

# --- スタイルユーティリティ ---
# parser.py から持ってきたスタイルマージロジック
# RenderContext で文字スタイルスタックを集約するために使用
Style = Dict[str, Any]  # {"s":[...], "a":{...}}


def _merge_styles(a: Style | None, b: Style | None) -> Style:
    """
    スタイル辞書 a と b をマージする。
    s[] はリストを結合 (重複削除)。
    a{} は辞書をマージ (b が優先)。
    """
    out: Style = {}
    for src in (a or {}), (b or {}):
        if not isinstance(src, dict):
            continue
        # merge s[] (スタイル配列)
        if 's' in src and isinstance(src['s'], list):
            out.setdefault('s', [])
            for v in src['s']:
                if v not in out['s']:
                    out['s'].append(v)
        # merge a{} (属性辞書)
        if 'a' in src and isinstance(src['a'], dict):
            out.setdefault('a', {})
            for k, v in src['a'].items():
                if isinstance(v, list):  # a.fx など
                    out['a'].setdefault(k, [])
                    ex = out['a'][k]
                    for it in v:
                        if it not in ex:
                            ex.append(it)
                elif v is not None:
                    out['a'][k] = v
    return out

# --- レンダリングコンテキスト ---


class RenderContext:
    """
    レンダリング中の状態 (段落スタイル、文字スタイル等) を保持します。
    """

    def __init__(self):
        # 現在の段落スタイル (par トークンの 'a' 属性)
        self.current_par_style: Dict[str, Any] = {}
        # 現在の文字スタイルスタック (ネストされたスタイルに対応)
        self.current_char_style_stack: List[Dict[str, Any]] = [{}]  # ボトムスタイル
        # 現在のリスト状態
        self.list_state: Dict[str, Any] = {"level": 0, "type": None}
        # 段落が開始されているかどうかのフラグ
        self.paragraph_started: bool = False
        logger.debug("RenderContext initialized.")

    def update_par_style(self, par_style: Dict[str, Any]):
        """par トークンに基づき段落スタイルとリスト状態を更新"""
        self.current_par_style = par_style
        list_info = par_style.get("list")
        if list_info:
            # TODO: 将来的にネストレベル (leftmargin) に応じた level 計算が必要
            self.list_state = {"level": 1, "type": list_info.get(
                "type"), "raw": list_info.get("raw")}
        else:
            self.list_state = {"level": 0, "type": None}
        logger.debug(
            f"Context par style updated. List state: {self.list_state}")

    def push_char_style(self, style: Style):
        """文字スタイルをスタックに積む"""
        self.current_char_style_stack.append(style)

    def pop_char_style(self):
        """文字スタイルをスタックから降ろす"""
        if len(self.current_char_style_stack) > 1:
            self.current_char_style_stack.pop()
        else:
            logger.warning(
                "Tried to pop the bottom character style from the stack.")

    def current_char_style(self) -> Style:
        """現在の文字スタイルスタックを集約して返す"""
        merged: Style = {}
        for style in self.current_char_style_stack:
            merged = _merge_styles(merged, style)
        return merged

# --- 基底レンダラークラス ---


class BaseRenderer(ABC):
    """
    レンダリングエンジンの抽象基底クラス (Visitor パターン)。
    normalized.json の 'runs' 配列を走査し、トークンごとに
    サブクラスで実装された _handle_* メソッドを呼び出します。
    """

    def __init__(self, doc_path: Path, doc: Optional[Dict[str, Any]] = None):
        """
        v1.2: オプショナルな doc 辞書をコンストラクタに追加。
        もし doc が提供されれば、それを使い、ファイルロードをスキップする。

        @param doc_path: normalized.json ファイルの Path オブジェクト
        @param doc: (Optional) 事前にロードされた JSON 辞書
        @raises FileNotFoundError: doc_path が存在しないファイルの場合 (doc is None)
        """
        if not doc_path or not doc_path.is_file():
            raise FileNotFoundError(
                f"Normalized JSON file not found: {doc_path}")

        self.doc_path = doc_path
        if doc:
            self.doc = doc
            logger.debug("Using pre-loaded doc dictionary.")
        else:
            self.doc: Dict[str, Any] = self._load_doc()  # JSONデータをロード

        self.attachment_dir: Path = doc_path.parent

        self.context = RenderContext()
        self.output: Any = self._init_output()
        logger.debug(f"{self.__class__.__name__} initialized for {doc_path}.")

    def _load_doc(self) -> Dict[str, Any]:
        """JSON ファイルをロードする"""
        try:
            with open(self.doc_path, 'r', encoding='utf-8') as f:
                logger.info(f"Loading JSON from: {self.doc_path}")
                return json.load(f)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to decode JSON from {self.doc_path}: {e}")
            raise ValueError(f"Invalid JSON file: {self.doc_path}") from e
        except Exception as e:
            logger.error(f"Failed to load JSON from {self.doc_path}: {e}")
            raise IOError(f"Could not read file: {self.doc_path}") from e

    @abstractmethod
    def _init_output(self) -> Any:
        """
        出力オブジェクト (例: [], "", Document, flowables) を初期化します。
        @return: サブクラス固有の出力オブジェクト
        """
        pass

    @abstractmethod
    def get_output(self) -> Any:
        """
        最終的な出力結果 (str, bytes など) を取得します。
        @return: レンダリング結果
        """
        pass

    def render(self) -> Any:
        """
        レンダリング処理のエントリポイント。
        ヘッダー、本文 (runs)、フッター、付録の順に処理します。
        """
        logger.info(f"Starting render for {self.__class__.__name__}...")
        try:
            self._render_header()

            body_runs = self.doc.get("fields", {}).get(
                "Body", {}).get("runs", [])
            logger.debug(f"Processing {len(body_runs)} runs in Body...")
            self._process_runs(body_runs)

            self._render_footer()
            self._render_appendix()

            logger.info(f"Render finished for {self.__class__.__name__}.")
            return self.get_output()
        except Exception as e:
            logger.error(
                f"Error during render in {self.__class__.__name__}: {e}", exc_info=True)
            # エラー発生時でも、途中までの結果を返すか、エラー情報を返す
            return self.get_output_on_error(e)

    def _process_runs(self, runs: List[Dict[str, Any]]):
        """
        runs 配列を走査し、トークン種別に応じてハンドラを呼び出す (Visitor)
        """
        if not runs:
            logger.debug("Runs list is empty.")
            # 空の本文でもデフォルトの段落を開始する必要があるか確認
            self._ensure_paragraph_started()
            # finalizeを呼んでおく (特にリストなどのコンテキストを閉じるため)
            self._finalize_paragraph()
            return

        for i, run in enumerate(runs):
            run_type = run.get("t")
            logger.debug(
                # Log run data briefly
                f"Processing run {i+1}/{len(runs)}: type='{run_type}' data='{str(run)[:80]}...'")

            # --- 段落 (Paragraph) ハンドリング ---
            if run_type == "par":
                # finalize は現在のブロックを閉じる
                # (例: Markdownなら現在の行をoutputに追加、HTMLならリスト要素を閉じるなど)
                self._finalize_paragraph()
                par_style = run.get("a", {})
                self.context.update_par_style(par_style)
                # start は新しいブロックを開始する
                # (例: Markdownならprefix設定、HTMLなら新しい<p>や<li>要素作成)
                self._start_paragraph(par_style)
                self.context.paragraph_started = True
                continue  # par トークン自体は直接描画しない

            # --- コンテンツ (Content) ハンドリング ---
            # par 以外のトークンが来たら、必要に応じてデフォルト段落を開始
            self._ensure_paragraph_started()

            # 各ハンドラは self.output または現在構築中の要素 (例: HTMLのcurrent_block) を変更する
            if run_type == "text":
                self._handle_text(run)
            elif run_type == "link":
                self._handle_link(run)
            elif run_type == "img":
                self._handle_img(run)
            elif run_type == "table":
                # テーブルはブロック要素なので、現在の段落を確定してからテーブル処理
                self._finalize_paragraph()
                self._handle_table(run)
                # テーブルの直後に新しい段落コンテキストを開始しておく
                # (テーブルの後にテキストが続く場合に備える)
                self.context.update_par_style({})  # デフォルトスタイルに戻す
                # 次のrun処理(_ensure_paragraph_started)で実際の段落要素が作られる
                self.context.paragraph_started = False  # startではなくensureに任せる
            elif run_type == "attachmentref":
                self._handle_attachmentref(run)
            elif run_type == "br":
                # v1.4 では非推奨だが、互換性のために処理
                self._handle_br(run)
            else:
                logger.warning(f"Unknown run type encountered: '{run_type}'")
                self._handle_unknown(run)

        # 最後のrun処理が終わった後、最後の段落を確定させる
        self._finalize_paragraph()

    def resolve_attachment_path(self, content_path: Optional[str]) -> Optional[Path]:
        """
        JSON 内の content_path (相対パス) と self.attachment_dir から
        添付ファイルの絶対パスを解決するヘルパーメソッド。

        @param content_path: JSON 内の attachments[n].content_path または runs[n].content_path
        @return: 解決された Path オブジェクト、または解決できない場合は None
        """
        if not content_path:
            logger.warning(
                "resolve_attachment_path called with empty content_path.")
            return None
        if not self.attachment_dir:
            # attachment_dir が何らかの理由で設定されていない場合
            # (通常は __init__ で設定されるはず)
            logger.error("BaseRenderer.attachment_dir is not set.")
            return None

        try:
            # content_path が既に絶対パスかどうかをチェック (基本は相対パス想定)
            if Path(content_path).is_absolute():
                logger.warning(
                    f"content_path '{content_path}' seems to be absolute, using it directly.")
                resolved_path = Path(content_path)
            else:
                # self.attachment_dir (doc_path.parent) を基準に解決
                resolved_path = (self.attachment_dir / content_path).resolve()

            # セキュリティチェック: 解決後のパスが attachment_dir の管理下にあるか
            # (例: ../../のようなパス操作で意図しない場所にアクセスさせないため)
            # resolve() はシンボリックリンクも解決するので、より安全
            if not str(resolved_path).startswith(str(self.attachment_dir.resolve())):
                logger.warning(
                    f"Resolved path '{resolved_path}' is outside the expected attachment directory '{self.attachment_dir.resolve()}'. Denying access.")
                return None

            logger.debug(
                f"Resolved attachment path: '{resolved_path}' from content_path: '{content_path}'")
            return resolved_path
        except Exception as e:
            # Path 操作中の予期せぬエラー (例: 不正なパス文字列)
            logger.error(
                f"Error resolving attachment path for '{content_path}' relative to '{self.attachment_dir}': {e}", exc_info=True)
            return None

    # --- サブクラスでの実装が必須な抽象メソッド ---

    @abstractmethod
    def _render_header(self):
        """文書のヘッダー (タイトル、メタ情報など) を描画します。"""
        pass

    @abstractmethod
    def _render_footer(self):
        """文書のフッターを描画します。"""
        pass

    @abstractmethod
    def _render_appendix(self):
        """文書の付録 (その他フィールド一覧など) を描画します。"""
        pass

    @abstractmethod
    def _start_paragraph(self, par_style: Dict[str, Any]):
        """
        'par' トークンに基づき、新しい段落コンテキストを開始します。
        (例: <p style="...">, doc.add_paragraph(), ParagraphStyle 設定)
        """
        pass

    @abstractmethod
    def _finalize_paragraph(self):
        """
        現在の段落を確定/終了します。
        (例: </p>, flowables.append(Paragraph(...)))
        """
        pass

    @abstractmethod
    def _ensure_paragraph_started(self):
        """
        まだ段落が開始されていない場合 (例: par トークンがない)、
        デフォルトの段落を開始します。
        """
        pass

    @abstractmethod
    def _handle_text(self, run: Dict[str, Any]):
        """
        "text" トークンを処理します。
        run['s'] (スタイル) と run['a'] (属性) に基づき装飾します。
        """
        pass

    @abstractmethod
    def _handle_link(self, run: Dict[str, Any]):
        """"link" トークン (HTTP または Notes) を処理します。"""
        pass

    @abstractmethod
    def _handle_img(self, run: Dict[str, Any]):
        """
        "img" トークンを処理します。
        HTML/MD: 'content_path' からパスを生成します。
        DOCX/PDF: 'attachment_dir' (自動解決済) と 'content_path' から画像を埋め込みます。
        """
        pass

    @abstractmethod
    def _handle_table(self, run: Dict[str, Any]):
        """
        "table" トークンを処理します。
        セル内の 'runs' を処理するために _process_runs を再帰的に呼び出す必要があります。
        """
        pass

    @abstractmethod
    def _handle_attachmentref(self, run: Dict[str, Any]):
        """"attachmentref" トークン (添付ファイルアイコン) を処理します。"""
        pass

    @abstractmethod
    def _handle_br(self, run: Dict[str, Any]):
        """
        "br" トークン (改行) を処理します。 (v1.4 では非推奨)
        通常は <br/> や \\n の挿入。
        """
        pass

    @abstractmethod
    def _handle_unknown(self, run: Dict[str, Any]):
        """不明なトークンタイプを処理します (フォールバック)。"""
        pass

    def get_output_on_error(self, e: Exception) -> Any:
        """レンダリング中にエラーが発生した場合のフォールバック"""
        logger.error(f"Returning partial or error output due to: {e}")
        # デフォルトでは、(おそらく不完全な) 出力をそのまま返す
        # サブクラスでオーバーライドして、エラーメッセージを埋め込むことも可能
        # 例: HTMLならエラーメッセージを含むHTMLを返す
        # if isinstance(self.output, list) and isinstance(self.output[0], str): # Markdown
        #    return "\n".join(self.output) + f"\n\n**Render Error:**\n```\n{e}\n```"
        return self.get_output()
