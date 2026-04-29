from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR


def add_table(slide, theme, headers, rows, left, top, width=None, col_widths=None,
              highlight_rows=None, highlight_cells=None):
    """データ表: ヘッダ行primary色背景+白文字、ストライプ行
    highlight_rows: list[int] — 0始まりのデータ行インデックス (ヘッダ除く)
    highlight_cells: list[{"row": int, "col": int}] — 0始まり
    """
    if width is None:
        width = theme.content_width

    hl_rows = set(highlight_rows or [])
    hl_cells = {(hc["row"], hc["col"]) for hc in (highlight_cells or [])}

    num_rows = len(rows) + 1
    num_cols = len(headers)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols, left, top, width, Inches(0.4 * num_rows)
    )
    table = table_shape.table

    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    else:
        col_width = width // num_cols
        for i in range(num_cols):
            table.columns[i].width = col_width

    for j, header_text in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header_text
        _style_cell(cell, theme, is_header=True)

    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i + 1, j)
            cell.text = str(cell_text)
            is_stripe = (i % 2 == 1)
            _style_cell(
                cell, theme,
                is_header=False,
                is_stripe=is_stripe,
                is_highlight_row=(i in hl_rows),
                is_highlight_cell=((i, j) in hl_cells),
            )


def _style_cell(cell, theme, is_header=False, is_stripe=False,
                is_highlight_row=False, is_highlight_cell=False):
    if is_header:
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.primary
    elif is_highlight_cell:
        c = theme.primary
        tint_r = c[0] + (255 - c[0]) * 70 // 100
        tint_g = c[1] + (255 - c[1]) * 70 // 100
        tint_b = c[2] + (255 - c[2]) * 70 // 100
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(tint_r, tint_g, tint_b)
    elif is_highlight_row:
        c = theme.primary
        tint_r = c[0] + (255 - c[0]) * 85 // 100
        tint_g = c[1] + (255 - c[1]) * 85 // 100
        tint_b = c[2] + (255 - c[2]) * 85 // 100
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(tint_r, tint_g, tint_b)
    elif is_stripe:
        cell.fill.solid()
        r = theme.primary[0]
        g = theme.primary[1]
        b = theme.primary[2]
        light_r = r + (255 - r) * 95 // 100
        light_g = g + (255 - g) * 95 // 100
        light_b = b + (255 - b) * 95 // 100
        cell.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.1)
    cell.margin_right = Inches(0.1)
    cell.margin_top = Inches(0.05)
    cell.margin_bottom = Inches(0.05)

    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = theme.font_size_body
            run.font.name = theme.font_body
            if is_header:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
            elif is_highlight_cell:
                run.font.color.rgb = theme.primary
                run.font.bold = True
            elif is_highlight_row:
                run.font.color.rgb = theme.text_primary
                run.font.bold = True
            else:
                run.font.color.rgb = theme.text_primary
