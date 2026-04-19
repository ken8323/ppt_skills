"""レイアウト"""
from src.layouts.cover import CoverLayout
from src.layouts.section_divider import SectionDividerLayout
from src.layouts.agenda import AgendaLayout
from src.layouts.content import ContentLayout
from src.layouts.chart_page import ChartPageLayout
from src.layouts.comparison import ComparisonLayout
from src.layouts.closing import ClosingLayout


LAYOUT_MAP = {
    "cover": CoverLayout,
    "section_divider": SectionDividerLayout,
    "agenda": AgendaLayout,
    "content": ContentLayout,
    "chart_page": ChartPageLayout,
    "comparison": ComparisonLayout,
    "closing": ClosingLayout,
}


def get_layout(name: str):
    """レイアウト名からクラスを取得"""
    layout_cls = LAYOUT_MAP.get(name)
    if layout_cls is None:
        raise ValueError(f"Unknown layout: {name}")
    return layout_cls()
