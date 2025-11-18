"""
Markdown レンダリングエンジン (MarkdownRenderer v1.1)

- BaseRenderer を継承
- normalized.json の runs 配列を Markdown (GFM) 文字列に変換します。
- スタイル (色、サイズ、fx) は Markdown の制約上、ほぼ無視されます。
- テーブルは GFM 形式で出力します (セル結合は非対応)。
- v1.1: BaseRenderer の変更に対応 (doc_path 入力)
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
import re  # for code block escaping

# --- common.py から BaseRenderer をインポート ---
try:
    # パッケージとして実行される場合
    from ..common import BaseRenderer, Style
except ImportError:
    # スクリプトとして直接実行される場合など
    import sys
    common_dir = Path(__file__).resolve().parent.parent
    if str(common_dir) not in sys.path:
        sys.path.insert(0, str(common_dir))
    from common import BaseRenderer, Style  # type: ignore


logger = logging.getLogger(__name__)


class MarkdownRenderer(BaseRenderer):
    """
    normalized.json (runs) を Markdown 文字列に変換するレンダラー。
    """

    def __init__(self, doc_path: Path):
        """
        v1.1: コンストラクタ引数を doc_path に変更。
        """
        # BaseRenderer の __init__ を呼び出す
        # これにより self.doc, self.attachment_dir, self.context, self.output が初期化される
        super().__init__(doc_path)

        # MarkdownRenderer 固有の状態
        # 現在の段落 (リストアイテム含む) のインライン要素 (文字列) を蓄積
        self._current_line_fragments: List[str] = []
        # 段落プレフィックス (リストマーカーやインデント)
        self._current_prefix: str = ""
        # 連続したリストアイテム間の改行を制御 (空行を入れるか)
        self._is_in_list: bool = False
        # 番号付きリストの現在の番号 (ネスト未対応)
        self._current_list_number: int = 1

    # --- BaseRenderer 抽象メソッドの実装 ---

    def _init_output(self) -> List[str]:
        """
        最終的な Markdown の「ブロック」要素 (段落、テーブル、ヘッダーなど) を
        格納するリストを初期化します。
        """
        return []

    def get_output(self) -> str:
        """
        蓄積されたブロック要素を2つの改行 (空行) で結合し、
        最終的な Markdown 文字列を返します。
        """
        # 最後に残ったフラグメントをフラッシュ (render() の最後でも呼ばれるが念のため)
        self._flush_current_line()

        # ブロック要素間に空行を入れる
        # filter(None, ...) で空文字列の要素を除去
        # strip() で前後の空白を除去してから結合
        rendered_blocks = filter(None, (block.strip()
                                 for block in self.output))
        return "\n\n".join(rendered_blocks)

    def _flush_current_line(self):
        """蓄積されたインライン要素 (_current_line_fragments) を
        プレフィックスと結合し、一つの段落 (文字列) として self.output に追加"""
        if not self._current_line_fragments:
            # フラッシュする内容がない場合
            # リストアイテム中で、かつプレフィックスがある場合のみ空リストアイテムとして出力する
            # (Markdownパーサーによっては空行だけだとリストが途切れるため)
            if self._is_in_list and self._current_prefix.strip():
                # 空のリストアイテムを追加 (例: "* ")
                self.output.append(self._current_prefix.rstrip())
                logger.debug(
                    f"Flushed empty list item with prefix: '{self._current_prefix.rstrip()}'")
            else:
                logger.debug("Flushing empty line (ignored).")
            return

        # インライン要素を結合し、前後の空白を除去
        line_content = "".join(self._current_line_fragments).strip()

        # プレフィックス (リストマーカー等) を付けて出力リストに追加
        full_line = f"{self._current_prefix}{line_content}"
        self.output.append(full_line.strip())  # 末尾の空白を削除して追加
        logger.debug(f"Flushed line: '{full_line[:80]}...'")

        # 状態をリセット
        self._current_line_fragments = []
        # プレフィックス (_current_prefix) は _start_paragraph でリセットされるため、ここではクリアしない

    def _render_header(self):
        """文書の Subject (H1) と Meta 情報を self.output に追加します。"""
        subject = self.doc.get("fields", {}).get(
            "Subject", {}).get("value", "No Subject")
        meta = self.doc.get("meta", {})

        self.output.append(f"# {subject}")

        # メタ情報は箇条書きで表現
        meta_items = []
        if meta.get('unid'):
            meta_items.append(f"- **UNID:** `{meta['unid']}`")
        if meta.get('form'):
            meta_items.append(f"- **Form:** {meta['form']}")
        if meta.get('created'):
            meta_items.append(f"- **Created:** {meta['created']}")
        if meta.get('modified'):
            meta_items.append(f"- **Modified:** {meta['modified']}")
        # 他のメタ情報も必要に応じて追加
        if meta_items:
            self.output.append("\n".join(meta_items))

        self.output.append("---")  # 水平線
        self.output.append("## 本文")  # 本文セクション見出し

    def _render_footer(self):
        """フッター (Markdownでは特になし)"""
        pass

    def _render_appendix(self):
        """付録データを Markdown テーブルとして self.output に追加します。"""
        rows = self._build_appendix_rows_helper()
        if not rows:
            logger.debug("No appendix data to render.")
            return

        # 最後の本文ブロックとの間に区切りを入れる
        # outputが空、または最後の要素が水平線でない場合に水平線を追加
        if not self.output or self.output[-1].strip() != "---":
            self.output.append("---")

        self.output.append("## 付録：その他のフィールド")

        headers = ["フィールド名", "型", "プレビュー"]
        # ヘッダー行と区切り行
        md_table_lines = [
            f"| {' | '.join(headers)} |",
            f"| {' | '.join(['---'] * len(headers))} |"
        ]

        # データ行
        for row in rows:
            # Markdownテーブル内の特殊文字 (| と改行) をエスケープ
            name = str(row.get('name', '')).replace(
                '|', '\\|').replace('\n', '<br>')
            ftype = str(row.get('type', '')).replace(
                '|', '\\|').replace('\n', '<br>')
            preview = str(row.get('preview', '')).replace(
                '|', '\\|').replace('\n', '<br>')  # 改行は <br> へ
            md_table_lines.append(f"| {name} | {ftype} | {preview} |")

        self.output.append("\n".join(md_table_lines))
        logger.debug(f"Appendix table rendered with {len(rows)} rows.")

    def _start_paragraph(self, par_style: Dict[str, Any]):
        """
        'par' トークンに基づき、プレフィックス (リストマーカー、インデント) を設定します。
        前の行 (_current_line_fragments) があればフラッシュします。
        """
        par_style = par_style or {}
        self._flush_current_line()  # 前の行の内容を出力リストへ

        self.context.paragraph_started = True  # BaseRenderer の状態更新
        list_state = self.context.list_state   # BaseRenderer が更新したリスト状態を取得

        prefix = ""
        is_new_list_start = False

        # 1. インデント
        margin_str = par_style.get(
            'leftmargin', '0in').lower().replace('in', '')
        try:
            margin_in = float(margin_str)
            indent_level = int(margin_in / 0.5)  # 0.5インチ単位
            # Markdownではインデントレベルをスペースで表現 (リスト内でも適用)
            # リストマーカーの前にスペースを追加
            prefix += "  " * indent_level  # スペース2つ * レベル
        except ValueError:
            logger.warning(
                f"Could not parse leftmargin: {par_style.get('leftmargin')}")
            pass

        # 2. リストマーカー
        if list_state['level'] > 0:
            if not self._is_in_list:  # リスト開始
                is_new_list_start = True
                self._current_list_number = 1  # 開始時にリセット

            list_type = list_state.get('type')
            # 番号付きリスト
            if list_type in ('number', 'alphaupper', 'alphalower', 'romanupper', 'romanlower'):
                prefix += f"{self._current_list_number}. "
                self._current_list_number += 1
            # 箇条書きリスト
            else:  # bullet, uncheck, square など
                prefix += "* "
            self._is_in_list = True
        else:
            # リスト外に出た
            if self._is_in_list:  # 直前までリストだった場合
                # リストと通常段落の間に空行を入れるかどうかの判断
                # ここでは特に追加しない (_flush_current_line と get_output の \n\n で対応)
                pass
            self._is_in_list = False
            self._current_list_number = 1  # リスト番号リセット

        # 新しいリストが始まる場合、前のブロックとの間に空行を入れる
        # (ただし、ヘッダー直後の場合は除く)
        # if is_new_list_start and self.output and not self.output[-1].startswith('#'):
        #     # outputの最後が空でなければ空行を追加
        #     if self.output[-1].strip():
        #          self.output.append("") # 空行を追加してリストとの間隔を確保

        self._current_prefix = prefix
        logger.debug(
            f"Started paragraph with prefix: '{prefix}', list_state: {list_state}")

    def _finalize_paragraph(self):
        """
        v1.1: このメソッドは _start_paragraph や _handle_table の前に呼ばれ、
              溜まっているフラグメントを書き出す役割。
              _start_paragraph の先頭で _flush_current_line を呼ぶため、処理不要。
        """
        pass

    def _ensure_paragraph_started(self):
        """まだ段落が開始されていない場合、デフォルトの段落を開始します。"""
        if not self.context.paragraph_started:
            logger.debug("No paragraph started, ensuring default paragraph.")
            # デフォルトスタイル({})で段落を開始 (これにより _flush_current_line も呼ばれる)
            self._start_paragraph({})

    def _append_fragment(self, text: str):
        """現在の行のフラグメントリストにテキストを追加"""
        if text:  # 空文字列は追加しない
            self._current_line_fragments.append(text)

    # --- ハンドラメソッド ---

    def _handle_text(self, run: Dict[str, Any]):
        """'text' トークンを Markdown 装飾付きテキストに変換し、フラグメントに追加します。"""
        text = run.get('text', '')
        if not text:
            return  # 空テキストは無視

        s = run.get('s', [])
        a = run.get('a', {})  # 'a' は基本的に無視

        # Markdown スタイル適用 (ネスト順に注意 - 外側から適用)
        # スクリプト (上/下付き) は HTML タグで表現 (GFM など一部パーサーが対応)
        script = a.get('script') or next(
            (fx for fx in a.get('fx', []) if fx in ['super', 'sub']), None)
        if script == 'super':
            text = f"<sup>{text}</sup>"
        if script == 'sub':
            text = f"<sub>{text}</sub>"

        if 'mono' in s:  # 等幅フォント -> バッククォート
            if '`' in text:  # コード内にバッククォートがある場合
                import re
                matches = re.findall(r'`+', text)
                max_backticks = max(len(m) for m in matches) if matches else 0
                delimiter = '`' * (max_backticks + 1)
                space = " " if text.startswith(
                    "`") or text.endswith("`") else ""
                text = f"{delimiter}{space}{text}{space}{delimiter}"
            else:
                text = f"`{text}`"

        if 's' in s or 'strike' in s:
            text = f"~~{text}~~"
        if 'i' in s:
            text = f"*{text}*"
        if 'b' in s:
            text = f"**{text}**"

        # 無視するスタイル
        if 'u' in s:
            logger.debug("Ignoring underline style ('u') for Markdown.")
        if a.get('color') or a.get('bgcolor') or a.get('size') or a.get('fx'):
            ignored_attrs = {k: v for k, v in a.items(
            ) if k in ['color', 'bgcolor', 'size', 'fx'] and v}
            if ignored_attrs:
                logger.debug(
                    f"Ignoring attributes {ignored_attrs} for Markdown.")

        self._append_fragment(text)

    def _handle_link(self, run: Dict[str, Any]):
        """'link' トークンを [label](href) 形式に変換し、フラグメントに追加します。"""
        label = run.get('label', run.get('href', ''))
        href = run.get('href', '#')

        # ラベル内の [] と href 内の () をエスケープ
        label = label.replace('[', '\\[').replace(']', '\\]')
        href = href.replace('(', '\\(').replace(')', '\\)').replace(
            ' ', '%20')  # URL内のスペースもエンコード

        md_link = f"[{label}]({href})"
        self._append_fragment(md_link)

    def _handle_img(self, run: Dict[str, Any]):
        """'img' トークンを ![alt](content_path) 形式に変換し、フラグメントに追加します。"""
        alt = run.get('alt', 'image').replace('[', '\\[').replace(']', '\\]')
        # content_path を使用
        src = run.get('content_path', '#').replace(
            '(', '\\(').replace(')', '\\)').replace(' ', '%20')

        if src == '#':
            logger.warning(f"Image run has no 'content_path'. Alt: {alt}")
            self._append_fragment(f"*[Image: {alt}]*")  # イタリックで代替テキスト表示
        else:
            md_img = f"![{alt}]({src})"
            self._append_fragment(md_img)

    def _handle_table(self, run: Dict[str, Any]):
        """'table' トークンを GFM テーブル文字列に変換し、output に直接追加します。"""
        self._flush_current_line()
        self.context.paragraph_started = False
        self._is_in_list = False
        self._current_list_number = 1
        self._current_prefix = ""

        logger.debug("Starting GFM table rendering.")
        rows_data = run.get('rows', [])
        if not rows_data:
            logger.warning("Table run found with no rows.")
            return

        md_table_lines = []
        num_cols = 0
        header_rendered = False

        # タブ付き表かどうか
        is_tabbed = run.get('attributes', {}).get('rowdisplay') == 'tabs'
        if is_tabbed:
            logger.warning(
                "Tabbed table cannot be fully represented in Markdown. Rendering all rows with labels.")

        for i, row in enumerate(rows_data):
            tab_label = row.get('attributes', {}).get(
                'tablabel') if is_tabbed else None
            cells = row.get('cells', [])

            # 最初の行から列数を決定 (ヘッダー行がない場合も考慮)
            if i == 0:
                num_cols = len(cells)
                if num_cols == 0:
                    logger.warning(
                        "Table has rows but the first row has no cells.")
                    return  # 空のテーブルは出力しない

                # ヘッダー行を生成 (常に最初の行をヘッダーとして扱う)
                header_texts = []
                for cell_data in cells:
                    header_text = self._extract_text_from_runs(
                        cell_data.get('runs', []))
                    header_texts.append(header_text)
                md_table_lines.append(f"| {' | '.join(header_texts)} |")
                md_table_lines.append(f"| {' | '.join(['---'] * num_cols)} |")
                header_rendered = True
                continue  # ヘッダー行はデータ行としては処理しない

            # データ行の処理
            cell_texts = []
            if len(cells) != num_cols:
                logger.warning(
                    f"Table row {i+1} has {len(cells)} cells, expected {num_cols}. Padding with empty cells.")
                cells.extend([{'runs': []}] *
                             max(0, num_cols - len(cells)))  # 足りない分を追加

            for j, cell_data in enumerate(cells[:num_cols]):  # 確定した列数まで
                if cell_data.get('colspan', 1) != 1 or cell_data.get('rowspan', 1) != 1:
                    logger.debug(
                        f"Ignoring colspan/rowspan for GFM table cell ({i+1},{j+1}).")
                cell_text = self._extract_text_from_runs(
                    cell_data.get('runs', []))
                # タブ付き表の場合、各行の先頭セルにラベルを追加 (オプション)
                # if is_tabbed and j == 0 and tab_label:
                #    cell_text = f"`{tab_label}`<br>{cell_text}"
                cell_texts.append(cell_text)

            md_table_lines.append(f"| {' | '.join(cell_texts)} |")

        # テーブル全体を1つのブロックとして output に追加
        if md_table_lines:  # 何かレンダリングされた場合のみ追加
            self.output.append("\n".join(md_table_lines))
            logger.debug(
                f"Finished GFM table rendering with {len(rows_data)} rows, {num_cols} columns.")
        else:
            logger.warning("No lines generated for table.")

    def _handle_attachmentref(self, run: Dict[str, Any]):
        """'attachmentref' を [添付: name](content_path) 形式のリンクまたはテキストに変換します。"""
        name = run.get('name', 'file')
        display_name = run.get('displayname', name)
        # content_path を使用
        content_path = run.get('content_path')

        if content_path:
            # content_path があればリンクにする
            href = content_path.replace(
                '(', '\\(').replace(')', '\\)').replace(' ', '%20')
            label = f"添付: {display_name}".replace(
                '[', '\\[').replace(']', '\\]')
            # イタリック + リンク
            self._append_fragment(f"[*[{label}]({href})*]")
        else:
            logger.warning(
                f"AttachmentRef run has no 'content_path'. Name: {name}")
            # content_path がなければイタリックテキストのみ
            self._append_fragment(f"*[添付: {display_name}]*")

    def _handle_br(self, run: Dict[str, Any]):
        """'br' トークンを強制改行 (スペース2つ + \\n) に変換し、フラグメントに追加します。"""
        # 現在のフラグメントの末尾が既に改行マーカーでないことを確認
        if not self._current_line_fragments or not self._current_line_fragments[-1].endswith("  \n"):
            self._append_fragment("  \n")
            logger.debug("Added line break ('br').")
        else:
            logger.debug("Skipping redundant 'br' token.")

    def _handle_unknown(self, run: Dict[str, Any]):
        """不明なトークンをデバッグ用テキストに変換し、フラグメントに追加します。"""
        run_type = run.get('t', 'unknown')
        logger.warning(
            f"Encountered unknown run type: '{run_type}'. Rendering as placeholder.")
        self._append_fragment(f"`[Unknown Run: {run_type}]`")

    # --- ヘルパーメソッド ---

    def _extract_text_from_runs(self, runs: List[Dict[str, Any]]) -> str:
        """
        テーブルセル用: runs 配列からスタイルを無視してテキストのみを抽出・結合します。
        Markdown テーブル内の特殊文字 (| と改行) をエスケープします。
        """
        parts = []
        if not runs:
            return ""  # 空の runs なら空文字列

        for run in runs:
            run_type = run.get('t')
            if run_type == 'text':
                parts.append(run.get('text', ''))
            elif run_type == 'link':
                parts.append(run.get('label', run.get('href', '')))
            elif run_type == 'img':
                parts.append(f"[Image: {run.get('alt', 'image')}]")
            elif run_type == 'attachmentref':
                parts.append(
                    f"[Attachment: {run.get('displayname', run.get('name', 'file'))}]")
            elif run_type == 'br':
                # GFM テーブルセル内では <br> が有効
                parts.append('<br>')
            elif run_type == 'table':
                parts.append('[Nested Table]')
                logger.warning(
                    "Nested table found inside a table cell. Cannot render text content in GFM.")
            # 他のタイプ (par など) はセル内では無視されるべき

        # | と ` をエスケープ, 改行は <br>
        text = "".join(parts)
        text = text.replace('|', '\\|').replace(
            '`', '\\`').replace('\n', '<br>')
        return text

    def _build_appendix_rows_helper(self) -> List[Dict[str, Any]]:
        """
        付録テーブル用の行データを構築します。
        BaseRenderer の self.doc を参照します。
        """
        # (html.py からの移植 - 変更なし)
        rows = []
        fields = self.doc.get("fields", {}) or {}
        layout = self.doc.get("layout", {}) or {}
        allow = set(layout.get("primary_fields_allowlist", []) or [])
        used = set(layout.get("used_in_body", []) or [])
        logger.debug(f"Appendix build: allowlist={allow}, used_in_body={used}")
        processed_names = set()

        for name, meta in fields.items():
            # $で始まる内部フィールド、$OLEObj* フィールドなどをスキップ (より厳格に)
            if name.startswith('$') or name in allow or name in used or name in processed_names:
                continue

            ftype = meta.get("type", "unknown")
            val = meta.get("value")
            preview = ""
            try:
                if ftype == 'richtext':
                    # リッチテキストはプレーンテキスト版を使用
                    # .get('text') を使う
                    preview = meta.get('text', '[RichText]')
                elif isinstance(val, list):
                    if val and all(isinstance(x, (str, int, float, bool)) for x in val):
                        preview = ", ".join(map(str, val[:5]))  # 先頭5件表示
                        if len(val) > 5:
                            preview += " ..."
                    else:
                        preview = f"[List of {len(val)} items]"
                elif isinstance(val, dict):
                    preview = "{...}"  # 辞書の中身は表示しない
                elif val is None:
                    preview = "[None]"
                else:
                    preview = str(val)

                # プレビューの最大長
                max_len = 100
                if len(preview) > max_len:
                    preview = preview[:max_len] + " ..."

            except Exception as e:
                preview = f"[Error previewing: {e}]"
                logger.warning(
                    # 簡略ログ
                    f"Error generating preview for field '{name}': {e}", exc_info=False)

            rows.append({"name": name, "type": ftype, "preview": preview})
            processed_names.add(name)

        rows.sort(key=lambda r: r["name"].lower())
        logger.debug(f"Generated {len(rows)} appendix rows.")
        return rows

    def get_output_on_error(self, e: Exception) -> str:
        """レンダリングエラー発生時にエラー情報を Markdown で返す"""
        logger.error(f"Render error occurred: {e}", exc_info=True)
        # 既存の出力に追加
        # エラー発生前の最後のブロックとの間に区切りを入れる
        if self.output and self.output[-1].strip() != "---":
            self.output.append("---")
        self.output.append("\n\n**Render Error:**")
        # エラーメッセージとスタックトレースの一部をコードブロックで表示
        import traceback
        tb_str = "".join(traceback.format_exception(
            type(e), e, e.__traceback__, limit=5))
        self.output.append(f"```\n{tb_str}\n```")

        # 部分的な結果とエラーメッセージを返す
        return self.get_output()


# --- 簡易テストコード ---
if __name__ == "__main__":
    # このスクリプトを直接実行した場合のテスト
    logging.basicConfig(
        level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

    # テスト用の normalized.json ファイルのパスを指定
    # プロジェクトルートからの相対パスなどを想定
    script_dir = Path(__file__).resolve().parent
    try:
        # このファイルがあるディレクトリの2つ上がプロジェクトルートと仮定
        project_root = script_dir.parent.parent
        # test_data ディレクトリが存在するか確認
        test_data_dir = project_root / "test_data"
        if not test_data_dir.is_dir():
            # test_data がなければ、スクリプトと同じディレクトリを試す
            test_data_dir = script_dir
            logger.warning(
                f"Directory '{project_root / 'test_data'}' not found. Looking for test file in '{script_dir}'.")

        test_json_path = test_data_dir / \
            "670309E2B88C5E9649258D2C000ADBDE.normalized.1.4.4-dev_initial.json"

    except Exception as e:
        print(f"Error constructing test JSON path: {e}")
        sys.exit(1)

    if test_json_path.exists():
        print(f"--- Running Markdown Render Test for: {test_json_path} ---")
        try:
            renderer = MarkdownRenderer(test_json_path)
            output_md = renderer.render()

            # 出力先Markdownファイルのパス (テストデータと同じディレクトリに保存)
            output_md_path = test_json_path.parent / \
                (test_json_path.stem + "_render_test.md")

            with open(output_md_path, "w", encoding='utf-8') as f:
                f.write(output_md)
            print(
                f"+++ Markdown render test successful. Output saved to: {output_md_path.resolve()} +++")

            # 付録の内容を少し表示 (デバッグ用)
            if hasattr(renderer, 'output') and isinstance(renderer.output, list):
                appendix_section_found = False
                appendix_content = []
                for block in renderer.output:
                    # 見出しを見つけてフラグを立てる
                    if block.strip().startswith("## 付録"):
                        appendix_section_found = True
                    # フラグが立っていれば内容を追加
                    if appendix_section_found:
                        appendix_content.append(block)

                if appendix_content:
                    print("\n--- Appendix Preview (from rendered output) ---")
                    # テーブル部分のみ表示 (簡易)
                    table_lines = [line for line in "\n\n".join(
                        appendix_content).splitlines() if line.strip().startswith('|')]
                    print("\n".join(table_lines[:8]))  # ヘッダ+区切り+5行
                    if len(table_lines) > 8:
                        print("...")

        except Exception as main_e:
            print(f"--- Markdown Render Test FAILED ---")
            logger.exception(
                "Error during Markdown render test")  # スタックトレースを出力

    else:
        print(f"!!! Test JSON file not found: {test_json_path.resolve()} !!!")
        print("Please create a test JSON file (e.g., in a 'test_data' directory next to 'src') or adjust the path in md.py.")
        # テスト用JSONがない場合、空のJSONでテストを試みる (オプション)
        try:
            print("\n--- Running with dummy data as fallback ---")
            # スクリプトディレクトリに一時ファイル作成
            dummy_path = script_dir / "dummy_render_test.json"
            dummy_doc = {"schema_version": "1.4-dev", "meta": {"unid": "DUMMY"}, "fields": {"Subject": {"value": "Dummy Doc"}, "Body": {"runs": [{"t": "par"}, {
                "t": "text", "text": "Hello, **Markdown**!"}, {"t": "br"}, {"t": "text", "text": "*List item 1*", "a": {"leftmargin": "0.5in"}, "list": {"type": "bullet"}}]}}}
            with open(dummy_path, "w", encoding='utf-8') as f:
                json.dump(dummy_doc, f, ensure_ascii=False, indent=2)

            renderer = MarkdownRenderer(dummy_path)
            output_md = renderer.render()
            print("\n--- Dummy Render Output ---")
            print(output_md)
            print("--- End Dummy Render Output ---")
            dummy_path.unlink()  # delete dummy file
            print(f"(Dummy test file '{dummy_path.name}' created and deleted)")
        except Exception as dummy_e:
            print(f"Dummy data test failed: {dummy_e}")
            logger.exception("Error during dummy data test")
