from pptx.util import Inches

from src.components.text import add_title, add_bullets, add_callout
from src.components.table import add_table
from src.components.shape import (
    add_matrix_2x2, add_pyramid, add_process_flow, add_cycle, add_org_chart,
)
from src.components.timeline import add_timeline, add_gantt
from src.components.icon import add_icon_row, add_kpi_card


class ContentLayout:
    def render(self, slide, theme, data):
        """汎用コンテンツ: 1/2/3カラム切り替え"""
        title = data.get("title", "")
        columns = data.get("columns", 1)
        components = data.get("components", [])

        if title:
            add_title(slide, theme, title, theme.margin_left, theme.margin_top)

        content_top = theme.content_area_top
        content_width = theme.content_width

        if columns == 1:
            self._render_components(
                slide, theme, components,
                theme.margin_left, content_top, content_width,
            )
        elif columns == 2:
            col_width = (content_width - Inches(0.4)) // 2
            left_comps = components[:len(components) // 2] if len(components) > 1 else components
            right_comps = components[len(components) // 2:] if len(components) > 1 else []

            self._render_components(
                slide, theme, left_comps,
                theme.margin_left, content_top, col_width,
            )
            self._render_components(
                slide, theme, right_comps,
                theme.margin_left + col_width + Inches(0.4), content_top, col_width,
            )
        elif columns == 3:
            col_width = (content_width - Inches(0.8)) // 3
            third = max(1, len(components) // 3)
            for col_i in range(3):
                col_comps = components[col_i * third:(col_i + 1) * third]
                col_left = theme.margin_left + col_i * (col_width + Inches(0.4))
                self._render_components(
                    slide, theme, col_comps, col_left, content_top, col_width,
                )

    def _render_components(self, slide, theme, components, left, top, width):
        """コンポーネントリストを順番に描画"""
        current_top = top
        for comp in components:
            comp_type = comp.get("type", "")
            if comp_type == "bullets":
                items = comp.get("items", [])
                height = Inches(0.4 * len(items))
                add_bullets(slide, theme, items, left, current_top, width=width, height=height)
                current_top += height + Inches(0.2)

            elif comp_type == "callout":
                text = comp.get("text", "")
                add_callout(slide, theme, text, left, current_top, width=width)
                current_top += Inches(1.0)

            elif comp_type == "table":
                headers = comp.get("headers", [])
                rows = comp.get("rows", [])
                add_table(slide, theme, headers, rows, left, current_top, width=width)
                current_top += Inches(0.4 * (len(rows) + 1)) + Inches(0.2)

            elif comp_type == "matrix_2x2":
                matrix_size = Inches(5.0)
                matrix_left = left + (width - matrix_size) // 2
                add_matrix_2x2(
                    slide, theme,
                    x_axis=comp.get("x_axis", ""),
                    y_axis=comp.get("y_axis", ""),
                    quadrants=comp.get("quadrants", ["", "", "", ""]),
                    left=matrix_left, top=current_top,
                )
                current_top += Inches(5.5)

            elif comp_type == "pyramid":
                add_pyramid(slide, theme, comp.get("levels", []), left, current_top, width=width)
                current_top += Inches(5.0)

            elif comp_type == "process_flow":
                add_process_flow(slide, theme, comp.get("steps", []), left, current_top, width=width)
                current_top += Inches(2.0)

            elif comp_type == "cycle":
                cycle_size = Inches(5.0)
                cycle_left = left + (width - cycle_size) // 2
                add_cycle(slide, theme, comp.get("items", []), cycle_left, current_top)
                current_top += Inches(5.5)

            elif comp_type == "org_chart":
                add_org_chart(slide, theme, comp.get("data", {}), left, current_top, width=width)
                current_top += Inches(5.0)

            elif comp_type == "timeline":
                add_timeline(slide, theme, comp.get("milestones", []), left, current_top, width=width)
                current_top += Inches(3.5)

            elif comp_type == "gantt":
                add_gantt(
                    slide, theme,
                    comp.get("tasks", []), comp.get("phases", []),
                    left, current_top, width=width,
                )
                current_top += Inches(4.5)

            elif comp_type == "icon_row":
                add_icon_row(slide, theme, comp.get("items", []), left, current_top, width=width)
                current_top += Inches(2.8)

            elif comp_type == "kpi_cards":
                cards = comp.get("cards", [])
                card_width = (width - Inches(0.3) * (len(cards) - 1)) // len(cards)
                for ci, card in enumerate(cards):
                    card_left = left + ci * (card_width + Inches(0.3))
                    add_kpi_card(
                        slide, theme,
                        card.get("value", ""), card.get("unit", ""), card.get("label", ""),
                        card_left, current_top,
                        width=card_width, color_idx=ci,
                    )
                current_top += Inches(2.5)
