from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor


def add_bar_chart(slide, theme, data, left, top, width=None, height=None,
                   horizontal=False, unit=None):
    """棒グラフ: 縦/横対応、データラベル付き"""
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


def add_line_chart(slide, theme, data, left, top, width=None, height=None, unit=None):
    """折れ線グラフ: マーカー付き"""
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
            point.format.fill.fore_color.rgb = theme.primary
        else:
            point.format.fill.fore_color.rgb = theme.secondary

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
