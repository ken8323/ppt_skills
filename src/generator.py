from pptx import Presentation
from pptx.util import Inches

from src.themes import get_theme
from src.layouts import get_layout
from src.components.footer import add_page_footer
from src.validator import validate_config


BLANK_LAYOUT_INDEX = 6

# フッターを付与しないレイアウト (背景塗りつぶし系)
FOOTER_SKIP_LAYOUTS = {"cover", "section_divider"}


def _should_skip_footer(layout_name: str, data: dict) -> bool:
    if layout_name in FOOTER_SKIP_LAYOUTS:
        return True
    if layout_name == "closing" and data.get("type") == "thank_you":
        return True
    return False


def generate_pptx(config: dict, output_path: str, strict: bool = False) -> Presentation:
    """設定辞書からpptxを生成してファイル保存。

    config 構造:
        {
            "theme": "monotone" | "dark" | "colorful",
            "footer": "株式会社ABC | 社外秘",  # 任意。各ページ左下に表示
            "brand_name": "...",  # 任意。theme.brand_name を上書き
            "slides": [
                {"layout": "cover", "data": {...}},
                ...
            ],
        }
    """
    if strict:
        validate_config(config)

    theme_name = config.get("theme", "monotone")
    theme = get_theme(theme_name)

    if "brand_name" in config:
        theme.brand_name = config["brand_name"]

    footer_text = config.get("footer", "")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[BLANK_LAYOUT_INDEX]

    slide_configs = config.get("slides", [])
    total = len(slide_configs)

    for idx, slide_cfg in enumerate(slide_configs, start=1):
        layout_name = slide_cfg.get("layout")
        data = slide_cfg.get("data", {})

        slide = prs.slides.add_slide(blank_layout)
        layout = get_layout(layout_name)
        layout.render(slide, theme, data)

        if not _should_skip_footer(layout_name, data):
            add_page_footer(slide, theme, idx, total, footer_text=footer_text)

    prs.save(output_path)
    return prs
