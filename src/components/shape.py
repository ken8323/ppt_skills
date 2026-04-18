import math
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN


def add_matrix_2x2(slide, theme, x_axis, y_axis, quadrants, left, top, size=None):
    """2x2マトリクス: 軸ラベル付き、各象限にテキスト配置
    quadrants: [top-left, top-right, bottom-left, bottom-right]
    """
    if size is None:
        size = Inches(5.0)

    cell_size = size // 2
    gap = Inches(0.05)

    colors = [
        theme.chart_colors[0],
        theme.chart_colors[2],
        theme.chart_colors[3],
        theme.chart_colors[1],
    ]

    positions = [
        (left, top),
        (left + cell_size + gap, top),
        (left, top + cell_size + gap),
        (left + cell_size + gap, top + cell_size + gap),
    ]

    for i, (qtext, (ql, qt)) in enumerate(zip(quadrants, positions)):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, ql, qt, cell_size - gap, cell_size - gap
        )
        shape.fill.solid()
        c = colors[i]
        light_r = c[0] + (255 - c[0]) * 8 // 10
        light_g = c[1] + (255 - c[1]) * 8 // 10
        light_b = c[2] + (255 - c[2]) * 8 // 10
        shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
        shape.line.color.rgb = theme.border
        shape.line.width = Pt(0.5)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = qtext
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.name = theme.font_body
            run.font.color.rgb = theme.text_primary

    x_label = slide.shapes.add_textbox(
        left, top + size + Inches(0.1), size, Inches(0.3)
    )
    x_label.text_frame.paragraphs[0].text = x_axis
    x_label.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for run in x_label.text_frame.paragraphs[0].runs:
        run.font.size = theme.font_size_caption
        run.font.color.rgb = theme.text_secondary
        run.font.name = theme.font_body

    y_label = slide.shapes.add_textbox(
        left - Inches(0.8), top + size // 2 - Inches(0.15), Inches(0.7), Inches(0.3)
    )
    y_label.text_frame.paragraphs[0].text = y_axis
    y_label.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    for run in y_label.text_frame.paragraphs[0].runs:
        run.font.size = theme.font_size_caption
        run.font.color.rgb = theme.text_secondary
        run.font.name = theme.font_body


def add_pyramid(slide, theme, levels, left, top, width=None, height=None):
    """ピラミッド: 台形の積み重ね"""
    if width is None:
        width = Inches(6.0)
    if height is None:
        height = Inches(4.5)

    n = len(levels)
    level_height = height // n
    gap = Inches(0.03)

    for i, text in enumerate(levels):
        level_width = int(width * (0.3 + 0.7 * (i + 1) / n))
        level_left = left + (width - level_width) // 2

        shape = slide.shapes.add_shape(
            MSO_SHAPE.TRAPEZOID,
            level_left, top + level_height * i + gap * i,
            level_width, level_height - gap,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shape.line.width = Pt(2)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = text
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True


def add_process_flow(slide, theme, steps, left, top, width=None, height=None):
    """プロセスフロー: 矢印で繋がった角丸矩形、横並び"""
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(1.5)

    n = len(steps)
    arrow_width = Inches(0.4)
    total_arrow_width = arrow_width * (n - 1)
    box_width = (width - total_arrow_width) // n
    box_height = height

    for i, step_text in enumerate(steps):
        box_left = left + i * (box_width + arrow_width)

        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            box_left, top, box_width, box_height,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.fill.background()

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = step_text
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True

        if i < n - 1:
            arrow_left = box_left + box_width
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                arrow_left, top + box_height // 2 - Inches(0.2),
                arrow_width, Inches(0.4),
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = theme.text_secondary
            arrow.line.fill.background()


def add_cycle(slide, theme, items, left, top, size=None):
    """サイクル図: 円形配置の要素"""
    if size is None:
        size = Inches(5.0)

    n = len(items)
    center_x = left + size // 2
    center_y = top + size // 2
    radius = size // 2 - Inches(0.6)
    node_size = Inches(1.4)

    for i, text in enumerate(items):
        angle = -math.pi / 2 + (2 * math.pi * i / n)
        node_x = int(center_x + radius * math.cos(angle) - node_size // 2)
        node_y = int(center_y + radius * math.sin(angle) - node_size // 2)

        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, node_x, node_y, node_size, node_size,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.fill.background()

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = text
        for run in tf.paragraphs[0].runs:
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True

        mid_angle = angle + math.pi / n
        arrow_x = int(center_x + (radius + Inches(0.3)) * math.cos(mid_angle) - Inches(0.15))
        arrow_y = int(center_y + (radius + Inches(0.3)) * math.sin(mid_angle) - Inches(0.15))

        arrow = slide.shapes.add_shape(
            MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, Inches(0.3), Inches(0.3),
        )
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = theme.text_secondary
        arrow.line.fill.background()
        arrow.rotation = math.degrees(mid_angle + math.pi / 2)


def add_org_chart(slide, theme, data, left, top, width=None, height=None):
    """組織図: ツリー構造、線で接続
    data: {"name": "CEO", "children": [{"name": "CTO"}, ...]}
    """
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(4.5)

    levels = []
    _collect_levels(data, 0, levels)

    level_height = height // max(len(levels), 1)
    node_height = Inches(0.6)

    _render_org_node(slide, theme, data, left, top, width, level_height, node_height)


def _collect_levels(node, depth, levels):
    while len(levels) <= depth:
        levels.append(0)
    levels[depth] += 1
    for child in node.get("children", []):
        _collect_levels(child, depth + 1, levels)


def _render_org_node(slide, theme, node, area_left, area_top, area_width,
                     level_height, node_height, depth=0):
    node_width = Inches(2.0)
    node_left = area_left + (area_width - node_width) // 2
    node_top = area_top

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, node_left, node_top, node_width, node_height,
    )
    if depth == 0:
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.primary
        text_color = RGBColor(0xFF, 0xFF, 0xFF)
    else:
        shape.fill.solid()
        color_idx = depth % len(theme.chart_colors)
        c = theme.chart_colors[color_idx]
        light_r = c[0] + (255 - c[0]) * 7 // 10
        light_g = c[1] + (255 - c[1]) * 7 // 10
        light_b = c[2] + (255 - c[2]) * 7 // 10
        shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
        text_color = theme.text_primary
    shape.line.color.rgb = theme.border
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].text = node["name"]
    for run in tf.paragraphs[0].runs:
        run.font.size = Pt(12)
        run.font.color.rgb = text_color
        run.font.name = theme.font_body
        if depth == 0:
            run.font.bold = True

    children = node.get("children", [])
    if children:
        n = len(children)
        child_area_width = area_width // n

        for i, child in enumerate(children):
            child_left = area_left + child_area_width * i
            child_top = area_top + level_height

            parent_center_x = node_left + node_width // 2
            child_center_x = child_left + child_area_width // 2

            mid_y = area_top + node_height + (level_height - node_height) // 2
            line1 = slide.shapes.add_connector(
                1, parent_center_x, node_top + node_height,
                parent_center_x, mid_y,
            )
            line1.line.color.rgb = theme.border
            line1.line.width = Pt(1.5)

            line2 = slide.shapes.add_connector(
                1, parent_center_x, mid_y, child_center_x, mid_y,
            )
            line2.line.color.rgb = theme.border
            line2.line.width = Pt(1.5)

            line3 = slide.shapes.add_connector(
                1, child_center_x, mid_y, child_center_x, child_top,
            )
            line3.line.color.rgb = theme.border
            line3.line.width = Pt(1.5)

            _render_org_node(
                slide, theme, child, child_left, child_top,
                child_area_width, level_height, node_height, depth + 1,
            )
