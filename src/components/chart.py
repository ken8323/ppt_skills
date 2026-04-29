from pptx.util import Inches, Pt, Emu
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from src.components._style import set_paragraph_text


# python-pptx チャートのプロット領域推定比率。左 15% は y 軸ラベル+目盛。
PLOT_AREA_LEFT_RATIO = 0.12
PLOT_AREA_RIGHT_RATIO = 0.98
PLOT_AREA_TOP_OFFSET = Inches(0.15)


def add_bar_chart(slide, theme, data, left, top, width=None, height=None,
                   horizontal=False, unit=None, annotations=None):
    """棒グラフ: 縦/横対応、データラベル付き、任意で注記オーバーレイ"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    for series in data["series"]:
        chart_data.add_series(series["name"], series["values"])

    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    chart_frame = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)
    chart = chart_frame.chart

    _style_chart(chart, theme, data, unit)

    if annotations and not horizontal:
        _render_annotations(slide, theme, annotations, data["labels"], left, top, width, height)


def add_line_chart(slide, theme, data, left, top, width=None, height=None, unit=None, annotations=None):
    """折れ線グラフ: マーカー付き、任意で注記オーバーレイ"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    for series in data["series"]:
        chart_data.add_series(series["name"], series["values"])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    _style_chart(chart, theme, data, unit)

    for series in chart.series:
        series.smooth = False
        series.format.line.width = Pt(2.5)

    if annotations:
        _render_annotations(slide, theme, annotations, data["labels"], left, top, width, height)


def add_pie_chart(slide, theme, data, left, top, width=None, height=None):
    """円グラフ: ラベル+割合表示"""
    if width is None:
        width = Inches(6.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    chart_data.add_series("", data["values"])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = theme.font_size_caption
    chart.legend.font.name = theme.font_body

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_percentage = True
    data_labels.show_category_name = False
    data_labels.show_value = False
    data_labels.font.size = theme.font_size_body
    data_labels.font.name = theme.font_body

    series = chart.series[0]
    for i, point in enumerate(series.points):
        color_idx = i % len(theme.chart_colors)
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = theme.chart_colors[color_idx]


def add_waterfall(slide, theme, data, left, top, width=None, height=None):
    """ウォーターフォールチャート: 積み上げ棒で増減を表現"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    labels = data["labels"]
    values = data["values"]

    invisible = []
    visible = []
    running = 0
    for i, val in enumerate(values):
        if i == 0:
            invisible.append(0)
            visible.append(val)
            running = val
        elif i == len(values) - 1:
            invisible.append(0)
            visible.append(val)
        else:
            if val >= 0:
                invisible.append(running)
                visible.append(val)
                running += val
            else:
                running += val
                invisible.append(running)
                visible.append(abs(val))

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("base", invisible)
    chart_data.add_series("value", visible)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    base_series = chart.series[0]
    base_series.format.fill.background()
    base_series.format.line.fill.background()

    value_series = chart.series[1]
    for i, val in enumerate(values):
        point = value_series.points[i]
        point.format.fill.solid()
        if i == 0 or i == len(values) - 1:
            point.format.fill.fore_color.rgb = theme.primary
        elif val >= 0:
            point.format.fill.fore_color.rgb = theme.success
        else:
            point.format.fill.fore_color.rgb = theme.danger

    chart.has_legend = False
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.border
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = theme.font_size_caption
    value_axis.tick_labels.font.name = theme.font_body

    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = theme.font_size_caption
    category_axis.tick_labels.font.name = theme.font_body
    category_axis.format.line.color.rgb = theme.border


def _style_chart(chart, theme, data, unit=None):
    """チャート共通スタイリング"""
    if len(data["series"]) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = theme.font_size_caption
        chart.legend.font.name = theme.font_body
    else:
        chart.has_legend = False

    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.border
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = theme.font_size_caption
    value_axis.tick_labels.font.name = theme.font_body
    if unit:
        value_axis.tick_labels.number_format = f'#,##0"{unit}"'
        value_axis.tick_labels.number_format_is_linked = False

    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = theme.font_size_caption
    category_axis.tick_labels.font.name = theme.font_body
    category_axis.format.line.color.rgb = theme.border
    category_axis.has_major_gridlines = False

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_value = True
    data_labels.show_category_name = False
    data_labels.font.size = theme.font_size_caption
    data_labels.font.name = theme.font_body
    data_labels.number_format = 'General'
    data_labels.number_format_is_linked = False

    for i, series in enumerate(chart.series):
        color_idx = i % len(theme.chart_colors)
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = theme.chart_colors[color_idx]


def _render_annotations(slide, theme, annotations, labels, chart_left, chart_top, chart_width, chart_height):
    """チャート上にピル型の注記を配置。
    annotations: [{"category": str|int, "text": str, "position": "top"|"bottom"}]
      - category は labels の要素か 0-based インデックス
      - position: "top"(デフォ) はチャート上端 / "bottom" は下端付近
    """
    plot_left = chart_left + int(chart_width * PLOT_AREA_LEFT_RATIO)
    plot_right = chart_left + int(chart_width * PLOT_AREA_RIGHT_RATIO)
    plot_span = plot_right - plot_left
    n = len(labels)

    pill_width = Inches(1.3)
    pill_height = Inches(0.32)

    for ann in annotations:
        cat = ann.get("category")
        text = ann.get("text", "")
        position = ann.get("position", "top")

        if isinstance(cat, int):
            idx = cat
        else:
            try:
                idx = labels.index(cat)
            except ValueError:
                continue
        if idx < 0 or idx >= n:
            continue

        cat_center_x = plot_left + plot_span * (idx + 0.5) // n
        pill_left = cat_center_x - pill_width // 2
        if position == "bottom":
            pill_top = chart_top + chart_height - Inches(0.9)
        else:
            pill_top = chart_top + PLOT_AREA_TOP_OFFSET

        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            pill_left, pill_top, pill_width, pill_height,
        )
        pill.fill.solid()
        pill.fill.fore_color.rgb = theme.secondary
        pill.line.fill.background()

        tf = pill.text_frame
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_paragraph_text(
            p, text,
            size=theme.font_size_caption, color=RGBColor(0xFF, 0xFF, 0xFF),
            name=theme.font_body, bold=True,
        )
