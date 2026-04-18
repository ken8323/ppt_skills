from src.themes.base import Theme, hex_to_rgb


class ColorfulTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#2D5BFF")
        self.secondary = hex_to_rgb("#00C49A")
        self.background = hex_to_rgb("#FFFFFF")
        self.text_primary = hex_to_rgb("#2C3E50")
        self.text_secondary = hex_to_rgb("#7F8C8D")
        self.border = hex_to_rgb("#E0E0E0")
        self.chart_colors = [
            hex_to_rgb("#2D5BFF"), hex_to_rgb("#00C49A"),
            hex_to_rgb("#FF6B35"), hex_to_rgb("#FFD166"),
            hex_to_rgb("#EF476F"), hex_to_rgb("#7B68EE"),
        ]
        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
