from pptx.util import Inches

from src.components.text import add_title, add_bullets
from src.components.chart import add_bar_chart, add_line_chart, add_pie_chart, add_waterfall


CHART_FUNCTIONS = {
    "bar": add_bar_chart,
    "line": add_line_chart,
    "pie": add_pie_chart,
    "waterfall": add_waterfall,
}


class ChartPageLayout:
    def render(self, slide, theme, data):
        """チャート主体ページ: チャート + オプションでキーポイント"""
        title = data.get("title", "")
        chart_data = data.get("chart", {})
        key_points = data.get("key_points", [])

        if title:
            add_title(slide, theme, title, theme.margin_left, theme.margin_top)

        content_top = theme.content_area_top
        chart_type = chart_data.get("type", "bar")
        chart_func = CHART_FUNCTIONS.get(chart_type, add_bar_chart)
        unit = chart_data.get("unit", None)

        extra_kwargs = {"unit": unit} if unit and chart_type in ("bar", "line") else {}

        if key_points:
            chart_width = int(theme.content_width * 0.65)
            chart_func(
                slide, theme, chart_data["data"],
                theme.margin_left, content_top,
                width=chart_width, height=theme.content_height,
                **extra_kwargs,
            )

            kp_left = theme.margin_left + chart_width + Inches(0.4)
            kp_width = int(theme.content_width * 0.30)
            add_bullets(
                slide, theme, key_points,
                kp_left, content_top + Inches(0.5),
                width=kp_width,
            )
        else:
            chart_func(
                slide, theme, chart_data["data"],
                theme.margin_left, content_top,
                width=theme.content_width, height=theme.content_height,
                **extra_kwargs,
            )
