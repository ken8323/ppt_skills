import pytest
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

from src.themes.base import Theme
from src.themes.monotone import MonotoneTheme
from src.themes.dark import DarkTheme
from src.themes.colorful import ColorfulTheme
from src.themes import get_theme


class TestThemeBase:
    def test_theme_has_required_color_attrs(self, monotone_theme):
        for attr in ["primary", "secondary", "background", "text_primary",
                      "text_secondary", "border", "chart_colors"]:
            assert hasattr(monotone_theme, attr)

    def test_theme_has_required_font_attrs(self, monotone_theme):
        for attr in ["font_title", "font_body", "font_size_title",
                      "font_size_subtitle", "font_size_body", "font_size_caption"]:
            assert hasattr(monotone_theme, attr)

    def test_theme_has_required_layout_attrs(self, monotone_theme):
        for attr in ["margin_top", "margin_bottom", "margin_left", "margin_right",
                      "content_area_top", "line_spacing"]:
            assert hasattr(monotone_theme, attr)

    def test_chart_colors_has_at_least_5(self, monotone_theme):
        assert len(monotone_theme.chart_colors) >= 5

    def test_color_returns_rgbcolor(self, monotone_theme):
        assert isinstance(monotone_theme.primary, RGBColor)

    def test_font_size_returns_pt(self, monotone_theme):
        assert isinstance(monotone_theme.font_size_title, int)


class TestThemeVariants:
    def test_monotone_white_background(self, monotone_theme):
        assert monotone_theme.background == RGBColor(0xFF, 0xFF, 0xFF)

    def test_dark_dark_background(self, dark_theme):
        assert dark_theme.background == RGBColor(0x1B, 0x2A, 0x4A)

    def test_colorful_white_background(self, colorful_theme):
        assert colorful_theme.background == RGBColor(0xFF, 0xFF, 0xFF)


class TestGetTheme:
    def test_get_monotone(self):
        theme = get_theme("monotone")
        assert isinstance(theme, MonotoneTheme)

    def test_get_dark(self):
        theme = get_theme("dark")
        assert isinstance(theme, DarkTheme)

    def test_get_colorful(self):
        theme = get_theme("colorful")
        assert isinstance(theme, ColorfulTheme)

    def test_unknown_theme_raises(self):
        with pytest.raises(ValueError, match="Unknown theme"):
            get_theme("neon")
