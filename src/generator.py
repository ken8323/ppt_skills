from pptx import Presentation
from pptx.util import Inches

from src.themes import get_theme
from src.layouts import get_layout


BLANK_LAYOUT_INDEX = 6


def generate_pptx(config: dict, output_path: str) -> Presentation:
    """設定辞書からpptxを生成してファイル保存。

    config 構造:
        {
            "theme": "monotone" | "dark" | "colorful",
            "slides": [
                {"layout": "cover", "data": {...}},
                ...
            ],
        }
    """
    theme_name = config.get("theme", "monotone")
    theme = get_theme(theme_name)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[BLANK_LAYOUT_INDEX]

    for slide_cfg in config.get("slides", []):
        layout_name = slide_cfg.get("layout")
        data = slide_cfg.get("data", {})

        slide = prs.slides.add_slide(blank_layout)
        layout = get_layout(layout_name)
        layout.render(slide, theme, data)

    prs.save(output_path)
    return prs
