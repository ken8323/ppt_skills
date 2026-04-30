from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN


_ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "right": PP_ALIGN.RIGHT,
    "center": PP_ALIGN.CENTER,
}


def add_table(slide, theme, headers, rows, left, top, width=None,
              col_widths=None, col_widths_ratio=None,
              align=None, banded=True,
              totals_row=None,
              highlight_rows=None, highlight_cells=None):
    """データ表。

    Args:
        align: str (全列共通) または list[str] (列ごと)。"left"/"right"/"center"。
               未指定時は数値列を right、それ以外を left に自動推定。
        banded: True で偶数行ストライプ (既定)。False で全行白。
        totals_row: True で数値列の合計を末尾行に自動追加 (非数値列は空文字)。
                    list[str] を渡せば明示的な合計行として使用。
                    末尾行は太字＋上線で強調。
        col_widths_ratio: list[float] — 列幅の相対比 (例: [2, 1, 1])。col_widths より優先。
        highlight_rows/highlight_cells: 0始まりインデックス。
    """
    if width is None:
        width = theme.content_width

    hl_rows = set(highlight_rows or [])
    hl_cells = {(hc["row"], hc["col"]) for hc in (highlight_cells or [])}

    num_cols = len(headers)
    aligns = _resolve_alignments(align, headers, rows, num_cols)
    total_values = _resolve_totals(totals_row, rows, num_cols)
    has_totals = total_values is not None

    num_rows = len(rows) + 1 + (1 if has_totals else 0)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols, left, top, width, Inches(0.4 * num_rows)
    )
    table = table_shape.table

    _apply_column_widths(table, num_cols, width, col_widths, col_widths_ratio)

    for j, header_text in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header_text
        _style_cell(cell, theme, is_header=True, align=aligns[j])

    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i + 1, j)
            cell.text = str(cell_text)
            is_stripe = banded and (i % 2 == 1)
            _style_cell(
                cell, theme,
                is_header=False,
                is_stripe=is_stripe,
                is_highlight_row=(i in hl_rows),
                is_highlight_cell=((i, j) in hl_cells),
                align=aligns[j],
            )

    if has_totals:
        totals_row_idx = num_rows - 1
        for j, val in enumerate(total_values):
            cell = table.cell(totals_row_idx, j)
            cell.text = str(val) if val is not None else ""
            _style_cell(cell, theme, is_header=False, is_totals=True, align=aligns[j])


def _apply_column_widths(table, num_cols, width, col_widths, col_widths_ratio):
    if col_widths_ratio:
        if len(col_widths_ratio) != num_cols:
            raise ValueError(
                f"col_widths_ratio の長さ {len(col_widths_ratio)} が列数 {num_cols} と一致しません。"
            )
        total = float(sum(col_widths_ratio))
        for i, r in enumerate(col_widths_ratio):
            table.columns[i].width = int(width * (r / total))
        return
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
        return
    col_width = width // num_cols
    for i in range(num_cols):
        table.columns[i].width = col_width


def _resolve_alignments(align, headers, rows, num_cols):
    if isinstance(align, list):
        if len(align) != num_cols:
            raise ValueError(
                f"align の長さ {len(align)} が列数 {num_cols} と一致しません。"
            )
        return align
    if isinstance(align, str):
        return [align] * num_cols
    # 自動推定: 数値列は right、それ以外は left
    result = []
    for j in range(num_cols):
        is_numeric = _column_is_numeric(rows, j)
        result.append("right" if is_numeric else "left")
    return result


def _column_is_numeric(rows, col_idx):
    if not rows:
        return False
    saw_value = False
    for row in rows:
        if col_idx >= len(row):
            continue
        cell = row[col_idx]
        if cell is None or cell == "":
            continue
        saw_value = True
        if isinstance(cell, (int, float)):
            continue
        if isinstance(cell, str) and _is_numeric_string(cell):
            continue
        return False
    return saw_value


def _is_numeric_string(s: str) -> bool:
    """カンマ・通貨記号・%を含む数値文字列も許容。"""
    cleaned = s.replace(",", "").replace("%", "").replace("¥", "").replace("$", "").strip()
    if not cleaned:
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def _resolve_totals(totals_row, rows, num_cols):
    if totals_row is None or totals_row is False:
        return None
    if isinstance(totals_row, list):
        if len(totals_row) != num_cols:
            raise ValueError(
                f"totals_row の長さ {len(totals_row)} が列数 {num_cols} と一致しません。"
            )
        return totals_row
    if totals_row is True:
        return _auto_totals(rows, num_cols)
    return None


def _auto_totals(rows, num_cols):
    totals = [None] * num_cols
    has_label = False
    for j in range(num_cols):
        if not _column_is_numeric(rows, j):
            continue
        s = 0.0
        is_int = True
        for row in rows:
            if j >= len(row):
                continue
            cell = row[j]
            if cell is None or cell == "":
                continue
            if isinstance(cell, (int, float)):
                s += cell
                if isinstance(cell, float):
                    is_int = False
            elif isinstance(cell, str) and _is_numeric_string(cell):
                cleaned = cell.replace(",", "").replace("¥", "").replace("$", "").strip()
                has_pct = cleaned.endswith("%")
                cleaned = cleaned.rstrip("%")
                v = float(cleaned)
                s += v
                if "." in cleaned or has_pct:
                    is_int = False
        if is_int:
            totals[j] = f"{int(s):,}"
        else:
            totals[j] = f"{s:,.2f}"
    # 合計行の最初の非数値列に "合計" を入れる
    for j in range(num_cols):
        if totals[j] is None:
            totals[j] = "合計" if not has_label else ""
            has_label = True
            break
    return totals


def _style_cell(cell, theme, is_header=False, is_stripe=False,
                is_highlight_row=False, is_highlight_cell=False,
                is_totals=False, align="left"):
    if is_header:
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.primary
    elif is_totals:
        c = theme.primary
        tint_r = c[0] + (255 - c[0]) * 88 // 100
        tint_g = c[1] + (255 - c[1]) * 88 // 100
        tint_b = c[2] + (255 - c[2]) * 88 // 100
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(tint_r, tint_g, tint_b)
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

    align_value = _ALIGN_MAP.get(align, PP_ALIGN.LEFT)

    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = align_value
        for run in paragraph.runs:
            run.font.size = theme.font_size_body
            run.font.name = theme.font_body
            if is_header:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
            elif is_totals:
                run.font.color.rgb = theme.primary
                run.font.bold = True
            elif is_highlight_cell:
                run.font.color.rgb = theme.primary
                run.font.bold = True
            elif is_highlight_row:
                run.font.color.rgb = theme.text_primary
                run.font.bold = True
            else:
                run.font.color.rgb = theme.text_primary
