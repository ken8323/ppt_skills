from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN


class CoverLayout:
    def render(self, slide, theme, data):
        """表紙: 中央にタイトル、サブタイトル、左下にクライアント名、右下に日付"""
        sw = theme.slide_width
        sh = theme.slide_height

        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, sh)
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.background
        bg.line.fill.background()

        bar_height = Inches(0.08)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, sh - Inches(1.2), sw, bar_height,
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = theme.secondary
        bar.line.fill.background()

        title_width = Inches(10.0)
        title_left = (sw - title_width) // 2
        title_top = sh // 2 - Inches(1.2)

        txBox = slide.shapes.add_textbox(title_left, title_top, title_width, Inches(1.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = data.get("title", "")
        run = p.runs[0]
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = theme.text_primary
        run.font.name = theme.font_title

        subtitle = data.get("subtitle", "")
        if subtitle:
            sub_box = slide.shapes.add_textbox(
                title_left, title_top + Inches(1.2), title_width, Inches(0.6),
            )
            tf2 = sub_box.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            p2.text = subtitle
            run2 = p2.runs[0]
            run2.font.size = Pt(20)
            run2.font.color.rgb = theme.text_secondary
            run2.font.name = theme.font_body

        client = data.get("client", "")
        if client:
            cl_box = slide.shapes.add_textbox(
                theme.margin_left, sh - Inches(1.0), Inches(4.0), Inches(0.4),
            )
            cl_box.text_frame.paragraphs[0].text = client
            for run in cl_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(14)
                run.font.color.rgb = theme.text_secondary
                run.font.name = theme.font_body

        date = data.get("date", "")
        if date:
            dt_box = slide.shapes.add_textbox(
                sw - theme.margin_right - Inches(3.0), sh - Inches(1.0),
                Inches(3.0), Inches(0.4),
            )
            dt_box.text_frame.paragraphs[0].text = date
            dt_box.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            for run in dt_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(14)
                run.font.color.rgb = theme.text_secondary
                run.font.name = theme.font_body
