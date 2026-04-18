from src.themes.base import Theme, hex_to_rgb


class DarkTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#FFFFFF")
        self.secondary = hex_to_rgb("#FF6B35")
        self.background = hex_to_rgb("#1B2A4A")
        self.text_primary = hex_to_rgb("#FFFFFF")
        self.text_secondary = hex_to_rgb("#A0B0C0")
        self.border = hex_to_rgb("#3A4F6F")
        self.chart_colors = [
            hex_to_rgb("#FFFFFF"), hex_to_rgb("#FF6B35"),
            hex_to_rgb("#5BA4E6"), hex_to_rgb("#A0B0C0"),
            hex_to_rgb("#FFD166"), hex_to_rgb("#06D6A0"),
        ]
        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
