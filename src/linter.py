"""構成 JSON の静的チェック (オーバーフロー / 推奨範囲逸脱)。

validator が「正しい/間違っている」の判定なのに対し、linter は
「動くが見た目が崩れたり読みにくくなる可能性がある」ケースを警告する。
generate_pptx() から既定で呼ばれ、stderr に出力される。
"""
from __future__ import annotations

# 推奨レンジ (skill.md の設計原則と整合)
TITLE_MAX = 40
SUBTITLE_MAX = 60
BULLET_ITEMS_MAX = 5
BULLET_LINE_MAX = 80
PILLARS_RANGE = (3, 5)
ICON_ROW_RANGE = (3, 5)
KPI_CARDS_RANGE = (2, 4)
TABLE_ROWS_MAX = 8
TABLE_COLS_MAX = 6
PROCESS_FLOW_RANGE = (3, 6)
CYCLE_RANGE = (3, 6)
PYRAMID_RANGE = (3, 5)


def lint_config(config: dict) -> list[str]:
    """構成全体を走査し、警告メッセージのリストを返す。"""
    warnings: list[str] = []
    for i, slide in enumerate(config.get("slides", [])):
        layout = slide.get("layout", "")
        data = slide.get("data", {})
        path = f"slides[{i}]({layout})"
        _lint_slide(layout, data, path, warnings)
    return warnings


def _lint_slide(layout: str, data: dict, path: str, warnings: list[str]) -> None:
    title = data.get("title", "")
    if isinstance(title, str) and len(title) > TITLE_MAX:
        warnings.append(
            f"{path}.title 文字数 {len(title)} 文字 (推奨 {TITLE_MAX} 文字以内)。"
            "アクションタイトルは短く端的に。"
        )

    subtitle = data.get("subtitle", "")
    if isinstance(subtitle, str) and len(subtitle) > SUBTITLE_MAX:
        warnings.append(
            f"{path}.subtitle 文字数 {len(subtitle)} 文字 (推奨 {SUBTITLE_MAX} 文字以内)。"
        )

    if layout == "chart_page" and not data.get("source"):
        warnings.append(
            f"{path} に source が未指定。数値スライドには出典の明記を推奨 "
            "(社内データ/調査機関名/年度。推計値は『※サンプル値』)。"
        )

    if layout == "agenda":
        items = data.get("items", [])
        if len(items) > 7:
            warnings.append(
                f"{path}.items が {len(items)} 件 (推奨 7 件以内)。"
            )

    for j, comp in enumerate(data.get("components", [])):
        _lint_component(comp, f"{path}.components[{j}]", warnings)
    for key in ("left_components", "right_components"):
        for j, comp in enumerate(data.get(key, [])):
            _lint_component(comp, f"{path}.{key}[{j}]", warnings)


def _lint_component(comp: dict, path: str, warnings: list[str]) -> None:
    t = comp.get("type", "")

    if t == "bullets":
        items = comp.get("items", [])
        if len(items) > BULLET_ITEMS_MAX:
            warnings.append(
                f"{path} bullets が {len(items)} 項目 "
                f"(推奨 {BULLET_ITEMS_MAX} 項目以内)。3項目以上の概念は図解優先。"
            )
        for k, item in enumerate(items):
            if isinstance(item, str) and len(item) > BULLET_LINE_MAX:
                warnings.append(
                    f"{path}.items[{k}] が {len(item)} 文字 "
                    f"(推奨 {BULLET_LINE_MAX} 文字以内)。1-2行に収める。"
                )

    elif t == "pillars":
        items = comp.get("items", [])
        if not (PILLARS_RANGE[0] <= len(items) <= PILLARS_RANGE[1]):
            warnings.append(
                f"{path} pillars が {len(items)} 本 "
                f"(推奨 {PILLARS_RANGE[0]}-{PILLARS_RANGE[1]} 本)。"
            )

    elif t == "icon_row":
        items = comp.get("items", [])
        if not (ICON_ROW_RANGE[0] <= len(items) <= ICON_ROW_RANGE[1]):
            warnings.append(
                f"{path} icon_row が {len(items)} 個 "
                f"(推奨 {ICON_ROW_RANGE[0]}-{ICON_ROW_RANGE[1]} 個)。"
            )

    elif t == "kpi_cards":
        cards = comp.get("cards", [])
        if not (KPI_CARDS_RANGE[0] <= len(cards) <= KPI_CARDS_RANGE[1]):
            warnings.append(
                f"{path} kpi_cards が {len(cards)} 枚 "
                f"(推奨 {KPI_CARDS_RANGE[0]}-{KPI_CARDS_RANGE[1]} 枚)。"
            )
        for k, card in enumerate(cards):
            if "delta" in card and "delta_direction" not in card:
                warnings.append(
                    f"{path}.cards[{k}] delta があるが delta_direction が未指定。"
                )

    elif t == "table":
        rows = comp.get("rows", [])
        headers = comp.get("headers", [])
        if len(rows) > TABLE_ROWS_MAX:
            warnings.append(
                f"{path} table の行数 {len(rows)} (推奨 {TABLE_ROWS_MAX} 行以内)。"
                "縦に長すぎる場合は分割するか、要点を抽出した別図解を検討。"
            )
        if len(headers) > TABLE_COLS_MAX:
            warnings.append(
                f"{path} table の列数 {len(headers)} (推奨 {TABLE_COLS_MAX} 列以内)。"
            )

    elif t == "process_flow":
        steps = comp.get("steps", [])
        if not (PROCESS_FLOW_RANGE[0] <= len(steps) <= PROCESS_FLOW_RANGE[1]):
            warnings.append(
                f"{path} process_flow のステップ数 {len(steps)} "
                f"(推奨 {PROCESS_FLOW_RANGE[0]}-{PROCESS_FLOW_RANGE[1]})。"
            )

    elif t == "cycle":
        items = comp.get("items", [])
        if not (CYCLE_RANGE[0] <= len(items) <= CYCLE_RANGE[1]):
            warnings.append(
                f"{path} cycle の要素数 {len(items)} "
                f"(推奨 {CYCLE_RANGE[0]}-{CYCLE_RANGE[1]})。"
            )

    elif t == "pyramid":
        levels = comp.get("levels", [])
        if not (PYRAMID_RANGE[0] <= len(levels) <= PYRAMID_RANGE[1]):
            warnings.append(
                f"{path} pyramid の段数 {len(levels)} "
                f"(推奨 {PYRAMID_RANGE[0]}-{PYRAMID_RANGE[1]})。"
            )

    elif t == "callout":
        text = comp.get("text", "")
        if isinstance(text, str) and len(text) > 60:
            warnings.append(
                f"{path} callout の text が {len(text)} 文字 "
                "(推奨 60 文字以内)。callout は強調したい一文に絞る。"
            )
