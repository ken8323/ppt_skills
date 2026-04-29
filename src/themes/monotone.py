from src.themes.base import Theme, hex_to_rgb


class MonotoneTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#1B2A4A")
        self.secondary = hex_to_rgb("#C8102E")
        self.background = hex_to_rgb("#FFFFFF")
        self.text_primary = hex_to_rgb("#1B2A4A")
        self.text_secondary = hex_to_rgb("#6B7B8D")
        self.border = hex_to_rgb("#D0D5DD")
        self.chart_colors = [
            hex_to_rgb("#1B2A4A"), hex_to_rgb("#C8102E"),
            hex_to_rgb("#4A7FB5"), hex_to_rgb("#8B9DAF"),
            hex_to_rgb("#D4A574"), hex_to_rgb("#6B8E6B"),
            hex_to_rgb("#9A5B8A"), hex_to_rgb("#3E6B5E"),
            hex_to_rgb("#B89454"), hex_to_rgb("#555E72"),
        ]
        self.success = hex_to_rgb("#2E8B57")
        self.warning = hex_to_rgb("#E69A1F")
        self.danger = hex_to_rgb("#C8102E")
        self.info = hex_to_rgb("#4A7FB5")
        self.neutral = hex_to_rgb("#8B9DAF")
        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
