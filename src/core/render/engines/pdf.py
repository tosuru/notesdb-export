"""
PDF レンダリングエンジン (PdfRenderer v1.1)

- BaseRenderer を継承
- normalized.json の runs 配列を ReportLab を使用して PDF 文書に変換します。
- ReportLab の Paragraph と Flowables を中心に構築します。
- スタイル (太字、斜体、色、サイズ、リンク) を ReportLab の XML タグで表現します。
- 配置、インデント、リスト、テーブル (セル結合、背景色含む) を適用します。
- 画像 (img) を埋め込みます。
- 日本語フォント (Noto Sans JP 推奨) を登録して使用します。
- v1.1: BaseRenderer の変更に対応 (doc_path 入力)
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from io import BytesIO
import colorsys  # 色変換用

# --- ReportLab のインポート ---
try:
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, BaseDocTemplate, PageTemplate, Frame
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.units import inch, pt
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader  # 画像サイズ取得用
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    HAVE_RL = True
except ImportError as e:
    HAVE_RL = False
    _import_error_message = (
        "ReportLab is required for PDF rendering. "
        f"Install with: pip install reportlab. Original error: {e}"
    )

# --- common.py から BaseRenderer をインポート ---
try:
    from ..common import BaseRenderer, Style, _merge_styles
except ImportError:
    import sys
    common_dir = Path(__file__).resolve().parent.parent
    if str(common_dir) not in sys.path:
        sys.path.insert(0, str(common_dir))
    from common import BaseRenderer, Style, _merge_styles  # type: ignore

# --- Utilities ---
# (docx.py から移植・調整)


def parse_color_rl(color_str: Optional[str]) -> Optional[colors.Color]:
    """ #RRGGBB 形式を ReportLab の Color に変換 """
    if not color_str or not color_str.startswith('#') or len(color_str) != 7:
        return None
    try:
        r = int(color_str[1:3], 16) / 255.0
        g = int(color_str[3:5], 16) / 255.0
        b = int(color_str[5:7], 16) / 255.0
        return colors.Color(r, g, b)
    except ValueError:
        return None


def parse_pt_size_rl(size_str: Optional[str]) -> Optional[float]:
    """ '12pt' 形式を float (ポイント数) に変換 """
    if not size_str or not size_str.lower().endswith('pt'):
        return None
    try:
        return float(size_str[:-2])
    except ValueError:
        return None


def parse_length_rl(length_str: Optional[str], default_unit=inch) -> Optional[float]:
    """ '1.5in', '12pt' などを ReportLab 単位 (ポイント) に変換 """
    if not length_str:
        return None
    length_str = length_str.lower()
    try:
        if length_str.endswith('in'):
            return float(length_str[:-2]) * inch
        elif length_str.endswith('pt'):
            return float(length_str[:-2]) * pt
        # 他の単位 (cm, mm) のサポートを追加する場合はここに記述
        # elif length_str.endswith('mm'):
        #     return float(length_str[:-2]) * mm
        else:  # 単位なしの場合はデフォルト単位とみなす (例: spaceafter の '1.5')
            # ParagraphStyleのleadingやspaceAfterは数値でポイント数を表すことが多い
            return float(length_str) * pt  # ポイントとして解釈
    except ValueError:
        pass
    logger.warning(f"Could not parse ReportLab length: '{length_str}'")
    return None


# --- Font Registration ---
DEFAULT_FONT_NAME = "NotoSansJP"  # ReportLab内で使用するフォント名
FALLBACK_FONT_NAME = "Helvetica"  # フォントが見つからない場合の代替


def find_font_path(preferred_font_name="Noto Sans CJK JP") -> Optional[Path]:
    """ システムから指定されたフォントファイルを探す (簡易版) """
    font_dirs = []
    if sys.platform == "win32":
        font_dirs.append(
            Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts")
    elif sys.platform == "darwin":
        font_dirs.extend(
            [Path("/Library/Fonts"), Path.home() / "Library/Fonts"])
    else:  # Linuxなど
        font_dirs.extend(
            ["/usr/share/fonts", "/usr/local/share/fonts", Path.home() / ".fonts"])

    # よくあるフォントファイル名 (優先順位順)
    font_patterns = [
        f"{preferred_font_name}-Regular.otf",  # Noto Sans CJK JP の例
        f"{preferred_font_name}*.ttf",
        "*.ttf",  # その他のTTF
        "*.otf",  # その他のOTF
    ]
    # Noto Sans JP Variable Font の場合
    if "Noto Sans JP" in preferred_font_name:
        # Variable Font を優先
        font_patterns.insert(0, "NotoSansJPVariable-VF.ttf")

    for font_dir in font_dirs:
        if font_dir.is_dir():
            for pattern in font_patterns:
                try:
                    found = list(font_dir.glob(f"**/{pattern}"))  # サブディレクトリも検索
                    if found:
                        logger.info(f"Font file found: {found[0]}")
                        return found[0]  # 最初に見つかったものを返す
                except Exception as e:
                    logger.warning(
                        f"Error searching for font '{pattern}' in {font_dir}: {e}")
    logger.warning(
        f"Could not find font '{preferred_font_name}' or fallback TTF/OTF.")
    return None


def register_rl_font(font_path: Optional[Path]) -> str:
    """ ReportLab に TTFont を登録し、登録名を返す """
    if not HAVE_RL or not font_path or not font_path.exists():
        logger.warning(
            "ReportLab or font file not available. Using fallback font.")
        return FALLBACK_FONT_NAME
    try:
        pdfmetrics.registerFont(TTFont(DEFAULT_FONT_NAME, str(font_path)))
        logger.info(f"Registered font '{DEFAULT_FONT_NAME}' from {font_path}")
        return DEFAULT_FONT_NAME
    except Exception as e:
        logger.error(
            f"Failed to register font {font_path}: {e}", exc_info=True)
        return FALLBACK_FONT_NAME


# --- Logger setup ---
logger = logging.getLogger(__name__)
if not HAVE_RL:
    logger.error(_import_error_message)

# --- Main Renderer Class ---


class PdfRenderer(BaseRenderer):
    """
    normalized.json (runs) を ReportLab を使用して PDF 文書に変換するレンダラー。
    """

    def __init__(self, doc_path: Path):
        if not HAVE_RL:
            raise ImportError(_import_error_message)

        super().__init__(doc_path)

        # ReportLab 固有の状態
        self.flowables: List[Any] = []
        self.styles: Dict[str, ParagraphStyle] = getSampleStyleSheet()
        self._registered_font_name: str = FALLBACK_FONT_NAME

        # フォント登録とスタイルの更新
        font_path = find_font_path()  # 環境に合わせてフォント名調整
        self._registered_font_name = register_rl_font(font_path)
        self._update_styles_with_font()

        # 現在の段落を構築するための状態
        self._current_paragraph_fragments: List[str] = []
        self._current_paragraph_style: ParagraphStyle = self.styles['Normal']

    def _update_styles_with_font(self):
        """ getSampleStyleSheet() の主要スタイルに登録したフォントを適用 """
        for style_name in ['Normal', 'BodyText', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'Title', 'Italic']:
            if style_name in self.styles:
                try:
                    self.styles[style_name].fontName = self._registered_font_name
                    # 必要に応じてデフォルトサイズなども調整
                    if style_name in ['Normal', 'BodyText']:
                        self.styles[style_name].fontSize = 10.5  # 例: 10.5pt
                    elif style_name == 'h1':
                        self.styles[style_name].fontSize = 18
                    elif style_name == 'h2':
                        self.styles[style_name].fontSize = 14

                except Exception as e:
                    logger.warning(
                        f"Could not update font for style '{style_name}': {e}")
        # 新しい基本スタイルを作成 (デフォルトフォント適用済み)
        self.styles.add(ParagraphStyle(
            name='BaseStyle', fontName=self._registered_font_name, fontSize=10.5, leading=14))

    # --- BaseRenderer 抽象メソッドの実装 ---

    def _init_output(self) -> BytesIO:
        """ PDF 内容を書き込む BytesIO バッファを初期化 """
        return BytesIO()

    def get_output(self) -> bytes:
        """ SimpleDocTemplate で PDF をビルドし、バイト列を返す """
        if not isinstance(self.output, BytesIO):
            logger.error("Output buffer is not initialized.")
            return b"Error: Output buffer invalid."
        if not self.flowables:
            logger.warning("No flowables generated for PDF.")
            # 空のPDFを生成するか、エラーを返すか
            # ここでは空でも生成を試みる

        try:
            # ページサイズとマージン設定
            doc = SimpleDocTemplate(self.output,
                                    pagesize='A4',
                                    leftMargin=1*inch,
                                    rightMargin=1*inch,
                                    topMargin=1*inch,
                                    bottomMargin=1*inch)
            doc.build(self.flowables)
            return self.output.getvalue()
        except Exception as e:
            logger.error(f"Failed to build PDF document: {e}", exc_info=True)
            return f"Error building PDF: {e}".encode('utf-8')

    def _render_header(self):
        """ 文書タイトル (Subject) と UNID を Flowables に追加 """
        subject = self.doc.get("fields", {}).get(
            "Subject", {}).get("value", self.doc_path.stem)
        meta = self.doc.get("meta", {})
        unid = meta.get("unid")

        self.flowables.append(Paragraph(subject, self.styles['h1']))
        if unid:
            # UNID 用に少し小さいスタイルを使う (例)
            unid_style = ParagraphStyle(
                name='UnidStyle', parent=self.styles['Normal'], fontSize=9, textColor=colors.grey)
            self.flowables.append(Paragraph(f"UNID: {unid}", unid_style))
        self.flowables.append(Spacer(1, 0.2*inch))  # タイトル後のスペース

    def _render_footer(self):
        """ フッター (ここでは何もしない) """
        pass

    def _render_appendix(self):
        """ 付録データをテーブルとして Flowables に追加 """
        rows_data = self._build_appendix_rows_helper()
        if not rows_data:
            return

        self.flowables.append(Spacer(1, 0.2*inch))  # 本文とのスペース
        self.flowables.append(Paragraph("付録：その他のフィールド", self.styles['h2']))

        table_data = [["フィールド名", "型", "プレビュー"]]  # ヘッダー
        col_widths = [2*inch, 1*inch, 3.5*inch]  # 列幅 (目安)

        for r in rows_data:
            # ReportLabテーブルのセルにはFlowable (Paragraphなど) を入れる
            name_p = Paragraph(str(r.get("name", "")), self.styles['Normal'])
            type_p = Paragraph(str(r.get("type", "")), self.styles['Normal'])
            # プレビューは改行を <br/> にして Paragraph に
            preview_text = str(r.get("preview", "")).replace('\n', '<br/>')
            preview_p = Paragraph(preview_text, self.styles['Normal'])
            table_data.append([name_p, type_p, preview_p])

        try:
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # ヘッダー背景
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 上揃え
                # ヘッダー太字 (フォント名に依存)
                ('FONTNAME', (0, 0), (-1, 0), self._registered_font_name + '-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),  # 罫線
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))
            self.flowables.append(table)
            logger.debug(
                f"Appendix table rendered with {len(rows_data)} rows.")
        except Exception as e:
            logger.error(
                f"Failed to render appendix table: {e}", exc_info=True)
            self.flowables.append(
                Paragraph(f"[付録テーブルの描画エラー: {e}]", self.styles['Normal']))

    def _flush_paragraph(self):
        """ 蓄積されたフラグメントから Paragraph を生成し、flowables に追加 """
        if self._current_paragraph_fragments:
            full_html = "".join(self._current_paragraph_fragments)
            # 空白のみの段落は Spacer に置き換えるか、無視する (ここでは無視)
            if full_html.strip():
                try:
                    para = Paragraph(full_html, self._current_paragraph_style)
                    self.flowables.append(para)
                    logger.debug(
                        f"Flushed paragraph: '{full_html[:50]}...' with style '{self._current_paragraph_style.name}'")
                except Exception as e:
                    logger.error(
                        f"Error creating Paragraph from HTML: '{full_html[:100]}...'. Error: {e}")
                    # エラー時もテキストとして追加を試みる (スタイルなし)
                    try:
                        from xml.sax.saxutils import escape
                        safe_text = escape(full_html)  # XMLタグをエスケープ
                        self.flowables.append(
                            Paragraph(f"[Paragraph Error] {safe_text}", self.styles['Normal']))
                    except:
                        pass  # それも失敗したら諦める
            else:
                logger.debug("Flushed empty paragraph (ignored).")

        # 状態リセット
        self._current_paragraph_fragments = []
        # スタイルは _start_paragraph で設定される

    def _start_paragraph(self, par_style: Dict[str, Any]):
        """ 新しい段落を開始。前の段落があればフラッシュし、新しい ParagraphStyle を準備 """
        self._flush_paragraph()  # 前の段落を確定

        par_style = par_style or {}
        self.context.paragraph_started = True

        # --- 新しい ParagraphStyle を生成 ---
        # 基本スタイルをコピーして変更を加える
        new_style = ParagraphStyle(
            # BaseStyleを継承
            name=f"par_{len(self.flowables)}", parent=self.styles['BaseStyle'])

        # 1. 配置 (Alignment)
        align_map = {
            "left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT,
            "justify": TA_JUSTIFY, "full": TA_JUSTIFY,
        }
        alignment = align_map.get(str(par_style.get('align', '')).lower())
        if alignment is not None:
            new_style.alignment = alignment

        # 2. インデント (Left Indent)
        left_indent = parse_length_rl(par_style.get(
            'leftmargin'), default_unit=pt)  # ポイントで取得
        if left_indent is not None:
            new_style.leftIndent = left_indent

        # 3. リスト (Bullet Indent/Text)
        list_state = self.context.list_state
        if list_state['level'] > 0:
            # ReportLabのリストは bulletIndent と bulletText で制御
            # 簡易的にインデントと '*' マーカーを設定
            indent = (list_state['level'] * 0.25 * inch)  # 0.25インチ * レベル
            new_style.leftIndent = indent
            # new_style.bulletIndent = indent - (0.15 * inch) # マーカー位置調整
            new_style.firstLineIndent = - (0.15 * inch)  # マーカー分のインデントを戻す

            list_type = list_state.get('type')
            # TODO: 番号付きリスト対応 (<bullet>タグ内で連番を使うなど)
            # ここでは簡易的に '*'
            bullet_text = '* '
            if list_type in ('number', 'alphaupper', 'alphalower', 'romanupper', 'romanlower'):
                # Paragraph で <seq id="list_id"/> タグを使う必要がある
                # ここではスタイル設定のみ
                # 自動連番 (要 BaseDocTemplate 設定)
                bullet_text = '<seq id="auto_list" />. '
                pass
            elif list_type == 'uncheck':
                bullet_text = '- '  # チェックボックス代替
            elif list_type == 'square':
                bullet_text = '- '  # 四角代替

            new_style.bulletText = bullet_text

        # 4. 段落後スペース (Space After)
        space_after = parse_length_rl(
            par_style.get('spaceafter'), default_unit=pt)
        if space_after is not None:
            new_style.spaceAfter = space_after

        # 5. 行間 (Leading - spaceafter が数値の場合に近似?)
        # ParagraphStyle の leading プロパティで設定 (fontSize の 1.2 倍などが一般的)
        # new_style.leading = new_style.fontSize * 1.4

        self._current_paragraph_style = new_style
        logger.debug(
            f"Started paragraph with RL style: name={new_style.name}, align={new_style.alignment}, leftIndent={new_style.leftIndent}, bullet={new_style.bulletText}")

    def _finalize_paragraph(self):
        """ 現在の段落をフラッシュ (確定) させる """
        self._flush_paragraph()
        # テーブルなどの後、paragraph_started を False にして新しい段落を要求
        if not self.context.paragraph_started:
            self._current_paragraph_style = self.styles['Normal']  # デフォルトに戻す

    def _ensure_paragraph_started(self):
        """ 段落フラグメントが空の場合、新しい段落を開始 """
        # ReportLab ではフラグメントを蓄積するので、 paragraph_started フラグではなく
        # フラグメントリストが空かで判断
        if not self.context.paragraph_started:  # BaseRendererのフラグも考慮
            logger.debug("No paragraph started context, ensuring paragraph.")
            self._start_paragraph(self.context.current_par_style)

    def _append_fragment(self, html_fragment: str):
        """ 現在の段落のHTMLフラグメントリストに追加 """
        if html_fragment:
            self._current_paragraph_fragments.append(html_fragment)

    # --- ハンドラメソッド ---

    def _handle_text(self, run_data: Dict[str, Any]):
        """ text run を ReportLab XML タグ付き文字列に変換してフラグメントに追加 """
        text = run_data.get('text', '')
        if not text:
            return
        text = text.replace('&', '&amp;').replace(
            '<', '&lt;').replace('>', '&gt;')  # 基本的なHTMLエスケープ
        text = text.replace('\n', '<br/>')  # 改行は <br/>

        s = run_data.get('s', [])
        a = run_data.get('a', {})
        tags_open = []
        tags_close = []

        # スタイルタグ
        if 'b' in s:
            tags_open.append('<b>')
            tags_close.insert(0, '</b>')
        if 'i' in s:
            tags_open.append('<i>')
            tags_close.insert(0, '</i>')
        if 'u' in s:
            tags_open.append('<u>')
            tags_close.insert(0, '</u>')
        if 's' in s or 'strike' in s:
            tags_open.append('<strike>')
            tags_close.insert(0, '</strike>')
        # if 'mono' in s: tags_open.append('<font name="Courier">'); tags_close.insert(0, '</font>') # 等幅 (要フォント登録)

        script = a.get('script') or next(
            (fx for fx in a.get('fx', []) if fx in ['super', 'sub']), None)
        if script == 'super':
            tags_open.append('<super>')
            tags_close.insert(0, '</super>')
        if script == 'sub':
            tags_open.append('<sub>')
            tags_close.insert(0, '</sub>')

        # <font> タグ (色、サイズ) - bgcolor は Paragraph では難しい
        font_attrs = []
        color = parse_color_rl(a.get('color'))
        if color:
            font_attrs.append(f'color="{color.hex()}"')  # '#RRGGBB' 形式
        size = parse_pt_size_rl(a.get('size'))
        if size:
            font_attrs.append(f'size="{size}"')  # ポイント数

        if font_attrs:
            tags_open.append(f'<font {" ".join(font_attrs)}>')
            tags_close.insert(0, '</font>')

        # fx は無視 (shadow, emboss など)
        fx = a.get('fx', [])
        ignored_fx = [f for f in fx if f not in ['super', 'sub']]
        if ignored_fx:
            logger.debug(f"Ignoring fx effects {ignored_fx} for PDF.")

        self._append_fragment("".join(tags_open) + text + "".join(tags_close))

    def _handle_link(self, run_data: Dict[str, Any]):
        """ link run を <a href="..."> タグに変換 """
        label = run_data.get('label', run_data.get('href', ''))
        href = run_data.get('href', '#')
        label_escaped = label.replace('&', '&amp;').replace(
            '<', '&lt;').replace('>', '&gt;')
        href_escaped = href.replace('&', '&amp;').replace(
            '"', '&quot;')  # URL内の & と " をエスケープ

        # リンク色を適用 (例: 青)
        link_html = f'<a href="{href_escaped}"><font color="blue"><u>{label_escaped}</u></font></a>'
        self._append_fragment(link_html)

    def _handle_img(self, run_data: Dict[str, Any]):
        """ img run を Image flowable として追加 """
        self._flush_paragraph()  # 画像の前に段落を確定

        alt_text = run_data.get('alt', 'image')
        content_path_rel = run_data.get('content_path')
        if not content_path_rel:
            logger.warning(f"Image run has no 'content_path'. Alt: {alt_text}")
            self.flowables.append(
                Paragraph(f"[画像(パスなし): {alt_text}]", self.styles['Normal']))
            return

        img_path = self.resolve_attachment_path(content_path_rel)
        if img_path and img_path.is_file():
            try:
                # 画像サイズを取得してページ幅に合わせる (最大幅制限)
                max_width = self.styles['Normal'].width  # 段落の利用可能幅 (近似)
                img_reader = ImageReader(str(img_path))
                img_w, img_h = img_reader.getSize()
                aspect = img_h / float(img_w)
                display_width = min(img_w, max_width)
                display_height = display_width * aspect

                img = Image(str(img_path), width=display_width,
                            height=display_height)
                self.flowables.append(img)
                self.flowables.append(Spacer(1, 0.1*inch))  # 画像後のスペース
                logger.debug(
                    f"Added image: {img_path} ({display_width:.1f}pt x {display_height:.1f}pt)")
            except Exception as e:
                logger.error(
                    f"Failed to process or add image '{img_path}': {e}", exc_info=True)
                self.flowables.append(
                    Paragraph(f"[画像(処理エラー): {alt_text}]", self.styles['Normal']))
        else:
            logger.warning(
                f"Image file not found: {img_path} (from {content_path_rel})")
            self.flowables.append(
                Paragraph(f"[画像(ファイル無し): {alt_text}]", self.styles['Normal']))

    def _handle_table(self, run_data: Dict[str, Any]):
        """ table run を Table flowable として追加 """
        self._flush_paragraph()  # テーブルの前に段落を確定
        self.context.paragraph_started = False  # テーブルはブロック

        rows_data = run_data.get('rows', [])
        if not rows_data:
            return

        logger.debug(f"Starting PDF table rendering: {len(rows_data)} rows.")
        table_data = []  # ReportLab Table 用のデータ (セルの内容は Flowable)
        table_styles_cmds = [  # TableStyle コマンド
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),  # 基本的な罫線
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 上揃え
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        col_widths = None  # 列幅 (後で計算する場合あり)

        # SPAN コマンド [( 'SPAN', (col, row), (col, row) ), ...]
        span_commands = []
        # BACKGROUND コマンド [( 'BACKGROUND', (col, row), (col, row), color ), ...]
        background_commands = []

        # 結合セルを追跡するためのグリッド (r, c) -> True (結合済み)
        merged_cells_tracker = {}

        # タブ付き表の警告
        if run_data.get('attributes', {}).get('rowdisplay') == 'tabs':
            logger.warning(
                "Tabbed table cannot be fully represented in PDF. Rendering all rows.")

        max_cols = 0  # テーブルの最大列数を計算

        for r_idx, row in enumerate(rows_data):
            cells_data = row.get('cells', [])
            row_content = []  # この行のセル内容 (Flowables)
            c_idx_tracker = 0  # 結合を考慮した実際の列インデックス

            for c_idx_data, cell_data in enumerate(cells_data):
                # 結合によってスキップされるセルかチェック
                while (r_idx, c_idx_tracker) in merged_cells_tracker:
                    # スキップする場合はプレースホルダ (None) を追加して列数を維持
                    row_content.append(None)
                    c_idx_tracker += 1

                if c_idx_tracker >= max_cols and r_idx > 0:  # 最初の行以外で最大列数を超えたら警告
                    logger.warning(
                        f"Table row {r_idx} seems to have more cells than previous rows due to spans. Max cols: {max_cols}")
                    # continue # このセルは無視する

                # --- セル内容のレンダリング (Paragraph) ---
                # セル内の runs を処理して Paragraph を生成
                cell_runs = cell_data.get('runs', [])
                # 簡易的な ParagraphStyle (セルのパディングは TableStyle で制御)
                cell_style = self.styles['Normal']  # 基本スタイル
                # runs から HTML 文字列を生成 (再帰ではなく直接処理)
                cell_html_fragments = []
                temp_renderer = PdfRenderer(self.doc_path)  # スタイル取得用の一時インスタンス
                temp_renderer._current_paragraph_fragments = cell_html_fragments  # フラグメントリストを差し替え
                temp_renderer._process_runs(cell_runs)  # runs を処理してフラグメントを生成
                cell_html = "".join(cell_html_fragments)

                if cell_html.strip():
                    cell_paragraph = Paragraph(cell_html, cell_style)
                else:
                    cell_paragraph = Spacer(0, 0)  # 空のセル

                row_content.append(cell_paragraph)

                # --- セル結合処理 ---
                colspan = cell_data.get('colspan', 1)
                rowspan = cell_data.get('rowspan', 1)
                if colspan > 1 or rowspan > 1:
                    start_col, start_row = c_idx_tracker, r_idx
                    end_col = start_col + colspan - 1
                    end_row = start_row + rowspan - 1
                    span_commands.append(
                        ('SPAN', (start_col, start_row), (end_col, end_row)))
                    # 結合されるセルを tracker に記録
                    for rr in range(start_row, end_row + 1):
                        for cc in range(start_col, end_col + 1):
                            if (rr, cc) != (start_row, start_col):
                                merged_cells_tracker[(rr, cc)] = True

                # --- セル背景色 ---
                cell_style_attr = cell_data.get('style', {})
                if isinstance(cell_style_attr, dict):
                    inner_style = cell_style_attr.get('style', {})
                    bgcolor_str = inner_style.get('bgcolor')
                    bgcolor = parse_color_rl(bgcolor_str)
                    if bgcolor:
                        background_commands.append(('BACKGROUND', (c_idx_tracker, r_idx), (
                            c_idx_tracker + colspan - 1, r_idx + rowspan - 1), bgcolor))

                # 次の実際の列インデックスへ
                c_idx_tracker += colspan

            # 行の実際のセル数（プレースホルダ含む）を最大列数として更新
            max_cols = max(max_cols, len(row_content))

            # 不足しているセルがあれば None で埋める
            row_content.extend([None] * (max_cols - len(row_content)))
            table_data.append(row_content)

        # table_data の None を空 Spacer で置き換える
        for r in range(len(table_data)):
            for c in range(len(table_data[r])):
                if table_data[r][c] is None:
                    table_data[r][c] = Spacer(0, 0)

        # スタイルコマンドを追加
        table_styles_cmds.extend(span_commands)
        table_styles_cmds.extend(background_commands)

        try:
            # Table オブジェクト作成
            table = Table(table_data, colWidths=col_widths,
                          repeatRows=1)  # repeatRows=1 でヘッダー繰り返し
            table.setStyle(TableStyle(table_styles_cmds))
            self.flowables.append(table)
            self.flowables.append(Spacer(1, 0.1*inch))  # テーブル後のスペース
            logger.debug(f"Finished PDF table rendering.")
        except Exception as e:
            logger.error(f"Error creating ReportLab Table: {e}", exc_info=True)
            self.flowables.append(
                Paragraph(f"[テーブル生成エラー: {e}]", self.styles['Normal']))

        # テーブルの後、新しい段落を開始するための準備
        self.context.paragraph_started = False

    def _handle_attachmentref(self, run_data: Dict[str, Any]):
        """ attachmentref run をテキストとしてフラグメントに追加 """
        name = run_data.get('displayname', run_data.get('name', 'file'))
        # PDFではリンクを埋め込むのが難しいため、テキスト表示のみ
        # イタリック体で表現
        self._append_fragment(f"<i>[添付: {name}]</i>")

    def _handle_br(self, run_data: Dict[str, Any]):
        """ br run を <br/> タグとしてフラグメントに追加 """
        # Paragraph内で <br/> タグとして処理される
        self._append_fragment("<br/>")

    def _handle_unknown(self, run_data: Dict[str, Any]):
        """ 不明な run をプレースホルダテキストとして追加 """
        run_type = run_data.get('t', 'unknown')
        logger.warning(f"Encountered unknown run type: '{run_type}'.")
        # 赤字で表示
        self._append_fragment(
            f'<font color="red">[Unknown Run: {run_type}]</font>')

    # --- Appendix Helper ---
    def _build_appendix_rows_helper(self) -> List[Dict[str, Any]]:
        """ 付録テーブル用の行データを構築 (md.py と同じロジック) """
        # (md.py からコピー＆ペースト)
        rows = []
        fields = self.doc.get("fields", {}) or {}
        layout = self.doc.get("layout", {}) or {}
        allow = set(layout.get("primary_fields_allowlist", []) or [])
        used = set(layout.get("used_in_body", []) or [])
        logger.debug(f"Appendix build: allowlist={allow}, used_in_body={used}")
        processed_names = set()

        for name, meta in fields.items():
            if name.startswith('$') or name in allow or name in used or name in processed_names:
                continue

            ftype = meta.get("type", "unknown")
            val = meta.get("value")
            preview = ""
            try:
                if ftype == 'richtext':
                    preview = meta.get('text', '[RichText Body]')
                elif isinstance(val, list):
                    if val and all(isinstance(x, (str, int, float, bool, type(None))) for x in val):
                        preview = ", ".join(map(str, val[:5]))
                        if len(val) > 5:
                            preview += " ..."
                    else:
                        preview = f"[List of {len(val)} complex items]"
                elif isinstance(val, dict):
                    preview = "{...}"
                elif val is None:
                    preview = "[None]"
                else:
                    preview = str(val)

                max_len = 100
                if len(preview) > max_len:
                    preview = preview[:max_len] + " ..."

            except Exception as e:
                preview = f"[Error previewing: {e}]"
                logger.warning(
                    f"Error generating preview for field '{name}': {e}", exc_info=False)

            rows.append({"name": name, "type": ftype, "preview": preview})
            processed_names.add(name)

        rows.sort(key=lambda r: r["name"].lower())
        logger.debug(f"Generated {len(rows)} appendix rows.")
        return rows

    def get_output_on_error(self, e: Exception) -> bytes:
        """ エラー発生時にエラー情報を埋め込んだPDFを生成試行 """
        logger.error(
            f"Render error occurred during PDF generation: {e}", exc_info=True)
        try:
            # 既存の flowables にエラー情報を追加
            self.flowables.append(Spacer(1, 0.2*inch))
            self.flowables.append(
                Paragraph("Render Error:", self.styles['h2']))
            error_style = ParagraphStyle(
                name='ErrorStyle', parent=self.styles['Normal'], textColor=colors.red)
            error_msg = f"{type(e).__name__}: {e}".replace(
                '&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            self.flowables.append(Paragraph(error_msg, error_style))

            # スタックトレース (簡易)
            import traceback
            tb_lines = traceback.format_exception(
                type(e), e, e.__traceback__, limit=3)
            tb_text = "".join(tb_lines).replace('&', '&amp;').replace(
                '<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')
            tb_style = ParagraphStyle(
                name='TracebackStyle', parent=self.styles['Normal'], fontSize=8, textColor=colors.darkgrey)
            self.flowables.append(Paragraph(tb_text, tb_style))

            # エラー情報を含むPDFを生成
            error_output = BytesIO()
            doc = SimpleDocTemplate(error_output, pagesize='A4', leftMargin=1 *
                                    inch, rightMargin=1*inch, topMargin=1*inch, bottomMargin=1*inch)
            doc.build(self.flowables)
            return error_output.getvalue()

        except Exception as build_err:
            logger.error(
                f"Failed to build PDF even with error info: {build_err}")
            return f"Render Error: {e}\nFailed to build PDF with error details.".encode('utf-8')


# --- 簡易テストコード ---
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
    import os  # find_font_path 用

    script_dir = Path(__file__).resolve().parent
    try:
        project_root = script_dir.parent.parent
        test_data_dir = project_root / "test_data"
        if not test_data_dir.is_dir():
            test_data_dir = script_dir
            logger.warning(
                f"Directory '{project_root / 'test_data'}' not found. Using '{script_dir}'.")
        test_json_path = test_data_dir / \
            "670309E2B88C5E9649258D2C000ADBDE.normalized.1.4.4-dev_initial.json"

    except Exception as e:
        print(f"Error constructing test JSON path: {e}")
        sys.exit(1)

    if test_json_path.exists():
        print(f"--- Running PDF Render Test for: {test_json_path} ---")
        try:
            # attachments ディレクトリが JSON と同じ場所にあると仮定
            # テスト用に content_path をダミーで設定
            import json
            with open(test_json_path, 'r', encoding='utf-8') as f:
                doc_data = json.load(f)

            if 'attachments' in doc_data and isinstance(doc_data['attachments'], list):
                for i, att in enumerate(doc_data['attachments']):
                    if 'name' in att and att.get('content_path') is None:
                        fname = Path(att['name'])
                        # 存在しないかもしれないダミーパスを設定
                        dummy_rel_path = f"attachments/{fname.stem}_{i}{fname.suffix}"
                        att['content_path'] = dummy_rel_path
                        # ダミー画像ファイルを作成 (テスト時のみ)
                        # img_abs_path = test_json_path.parent / dummy_rel_path
                        # img_abs_path.parent.mkdir(exist_ok=True)
                        # if not img_abs_path.exists() and fname.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif']:
                        #    try:
                        #       from PIL import Image as PILImage
                        #       dummy_img = PILImage.new('RGB', (60, 30), color = 'red')
                        #       dummy_img.save(img_abs_path)
                        #       logger.info(f"Created dummy image for testing: {img_abs_path}")
                        #    except ImportError:
                        #        logger.warning("Pillow not installed, cannot create dummy images for testing.")
                        #    except Exception as img_e:
                        #        logger.warning(f"Could not create dummy image {img_abs_path}: {img_e}")

            temp_json_path = test_json_path.with_name(
                test_json_path.stem + "_temp_test.json")
            with open(temp_json_path, 'w', encoding='utf-8') as f:
                json.dump(doc_data, f, ensure_ascii=False, indent=2)

            renderer = PdfRenderer(temp_json_path)
            output_bytes = renderer.render()

            output_pdf_path = test_json_path.parent / \
                (test_json_path.stem + "_render_test.pdf")
            with open(output_pdf_path, "wb") as f:
                f.write(output_bytes)
            print(
                f"+++ PDF render test successful. Output saved to: {output_pdf_path.resolve()} +++")

            temp_json_path.unlink()  # 一時JSON削除

        except ImportError as import_err:
            print(
                f"!!! PDF Render Test SKIPPED due to missing library: {import_err} !!!")
        except FileNotFoundError as fnf_err:
            print(
                f"!!! PDF Render Test FAILED: File not found - {fnf_err} !!!")
            logger.error(
                f"Make sure the input JSON and potentially referenced attachment files exist.")
        except Exception as main_e:
            print(f"--- PDF Render Test FAILED ---")
            logger.exception("Error during PDF render test")
            # エラー時のPDFも保存
            try:
                if 'output_bytes' in locals() and isinstance(output_bytes, bytes):
                    error_pdf_path = test_json_path.parent / \
                        (test_json_path.stem + "_render_error.pdf")
                    with open(error_pdf_path, "wb") as f_err:
                        f_err.write(output_bytes)
                    print(
                        f"--- Error output saved to: {error_pdf_path.resolve()} ---")
            except Exception as save_err:
                print(f"Could not save error PDF: {save_err}")

    else:
        print(f"!!! Test JSON file not found: {test_json_path.resolve()} !!!")
        print("Please create a test JSON file or adjust the path.")
