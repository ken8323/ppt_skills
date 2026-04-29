from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class Grid:
    """12列グリッド座標計算。margin_left/right, content_area_top, margin_bottom で枠を決め、
    12 分割のカラム番号と行番号から left/top/width/height を返す。"""

    TOTAL_COLS = 12
    GUTTER = Inches(0.2)

    def __init__(self, theme):
        self._theme = theme

    @property
    def inner_left(self):
        return self._theme.margin_left

    @property
    def inner_top(self):
        return self._theme.content_area_top

    @property
    def inner_width(self):
        return self._theme.content_width

    @property
    def inner_height(self):
        return self._theme.content_height

    @property
    def col_width(self):
        gutters = self.GUTTER * (self.TOTAL_COLS - 1)
        return (self.inner_width - gutters) // self.TOTAL_COLS

    def span_width(self, span_cols: int):
        """span_cols マス分の横幅 (隙間込み)"""
        return self.col_width * span_cols + self.GUTTER * (span_cols - 1)

    def col_x(self, col: int):
        """col 番目 (0 始まり) の左端 x"""
        return self.inner_left + (self.col_width + self.GUTTER) * col

    def cell(self, col: int, top, span_cols: int = 1, height=None):
        """左上 (col, top) に span_cols マス占有するセルを返す。
        戻り値: (left, top, width, height)"""
        left = self.col_x(col)
        width = self.span_width(span_cols)
        return left, top, width, height


class Theme:
    primary: RGBColor
    secondary: RGBColor
    background: RGBColor
    text_primary: RGBColor
    text_secondary: RGBColor
    border: RGBColor
    chart_colors: list

    # Semantic palette (各テーマで上書き可)
    success: RGBColor = RGBColor(0x2E, 0x8B, 0x57)
    warning: RGBColor = RGBColor(0xE6, 0x9A, 0x1F)
    danger: RGBColor = RGBColor(0xC8, 0x10, 0x2E)
    info: RGBColor = RGBColor(0x4A, 0x7F, 0xB5)
    neutral: RGBColor = RGBColor(0x8B, 0x9D, 0xAF)

    font_title: str = "Yu Gothic"
    font_body: str = "Yu Gothic"

    # 6 段階のフォント階層
    font_size_h1: int = Pt(24)
    font_size_h2: int = Pt(18)
    font_size_h3: int = Pt(14)
    font_size_body: int = Pt(14)
    font_size_caption: int = Pt(10)
    font_size_footnote: int = Pt(9)

    # 既存コード互換エイリアス (古いコードがアクセスしても動くよう保つ)
    @property
    def font_size_title(self):
        return self.font_size_h1

    @property
    def font_size_subtitle(self):
        return self.font_size_h2

    margin_top: int = Inches(0.4)
    margin_bottom: int = Inches(0.55)
    margin_left: int = Inches(0.6)
    margin_right: int = Inches(0.6)
    content_area_top: int = Inches(1.4)
    line_spacing: float = 1.2

    # ブランド情報 (任意)
    brand_name: str = ""
    brand_logo: str = ""  # ファイルパス (設定時のみ使用)

    @property
    def slide_width(self) -> int:
        return Inches(13.333)

    @property
    def slide_height(self) -> int:
        return Inches(7.5)

    @property
    def content_width(self) -> int:
        return self.slide_width - self.margin_left - self.margin_right

    @property
    def content_height(self) -> int:
        return self.slide_height - self.content_area_top - self.margin_bottom

    @property
    def grid(self) -> Grid:
        return Grid(self)
