from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class Theme:
    primary: RGBColor
    secondary: RGBColor
    background: RGBColor
    text_primary: RGBColor
    text_secondary: RGBColor
    border: RGBColor
    chart_colors: list

    font_title: str = "Yu Gothic"
    font_body: str = "Yu Gothic"
    font_size_title: int = Pt(24)
    font_size_subtitle: int = Pt(16)
    font_size_body: int = Pt(14)
    font_size_caption: int = Pt(10)

    margin_top: int = Inches(0.4)
    margin_bottom: int = Inches(0.4)
    margin_left: int = Inches(0.6)
    margin_right: int = Inches(0.6)
    content_area_top: int = Inches(1.4)
    line_spacing: float = 1.2

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
