from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


def add_title(slide, theme, text, left, top, width=None):
    """スライドタイトル: 左上配置、太字、下に区切り線"""
    if width is None:
        width = theme.content_width

    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.6))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_title
    run.font.bold = True
    run.font.color.rgb = theme.text_primary
    run.font.name = theme.font_title

    line_top = top + Inches(0.6)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, line_top, width, Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = theme.secondary
    line.line.fill.background()


def add_subtitle(slide, theme, text, left, top, width=None):
    """サブタイトル: タイトル直下、やや小さく"""
    if width is None:
        width = theme.content_width

    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.4))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_subtitle
    run.font.color.rgb = theme.text_secondary
    run.font.name = theme.font_body


def add_bullets(slide, theme, items, left, top, width=None, height=None):
    """箇条書き: インデント2階層対応、行頭は控えめ記号"""
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(4.0)

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        if isinstance(item, dict):
            text = item["text"]
            level = item.get("level", 0)
        else:
            text = item
            level = 0

        prefix = "  " * level + "― "
        p.text = prefix + text
        for run in p.runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = theme.text_primary
            run.font.name = theme.font_body
        p.space_after = Pt(6)

        if level > 0:
            for run in p.runs:
                run.font.color.rgb = theme.text_secondary
            p.level = level


def add_callout(slide, theme, text, left, top, width=None, height=None):
    """強調ボックス: 背景色付き矩形内にテキスト"""
    if width is None:
        width = theme.content_width
    if height is None:
        line_count = text.count("\n") + 1
        height = Inches(0.3 + 0.3 * line_count)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    r = theme.primary[0]
    g = theme.primary[1]
    b = theme.primary[2]
    light_r = r + (255 - r) * 9 // 10
    light_g = g + (255 - g) * 9 // 10
    light_b = b + (255 - b) * 9 // 10
    shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
    shape.line.color.rgb = theme.primary
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.08)
    tf.margin_bottom = Inches(0.08)
    tf.word_wrap = True
    tf.paragraphs[0].text = text
    for run in tf.paragraphs[0].runs:
        run.font.size = theme.font_size_body
        run.font.color.rgb = theme.text_primary
        run.font.name = theme.font_body
        run.font.bold = True


def add_footnote(slide, theme, text, left, bottom_margin=None):
    """脚注: スライド下部、小フォント"""
    if bottom_margin is None:
        bottom_margin = theme.margin_bottom

    top = theme.slide_height - bottom_margin - Inches(0.3)
    width = theme.content_width

    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.3))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_caption
    run.font.color.rgb = theme.text_secondary
    run.font.name = theme.font_body
