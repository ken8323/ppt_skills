"""設定JSONのバリデーション。generate_pptx の strict=True 時に呼ばれる。"""
import json
from pathlib import Path

import jsonschema

_SCHEMA_PATH = Path(__file__).parent / "schema.json"
_schema = None


def _get_schema():
    global _schema
    if _schema is None:
        _schema = json.loads(_SCHEMA_PATH.read_text(encoding="utf-8"))
    return _schema


class ConfigValidationError(ValueError):
    """スキーマ検証失敗時に送出する例外。メッセージはユーザー向けに整形済み。"""


def validate_config(config: dict) -> None:
    """config を JSON Schema で検証。不整合があれば ConfigValidationError を送出。"""
    schema = _get_schema()
    validator = jsonschema.Draft7Validator(schema)
    errors = sorted(validator.iter_errors(config), key=lambda e: list(e.path))
    if not errors:
        _validate_business_rules(config)
        return

    messages = []
    for err in errors:
        path = " > ".join(str(p) for p in err.absolute_path) or "(root)"
        messages.append(f"  [{path}] {err.message}")
    raise ConfigValidationError("設定JSONにエラーがあります:\n" + "\n".join(messages))


def _validate_business_rules(config: dict) -> None:
    """JSON Schema で表現しにくいビジネスルールを追加検証。"""
    for i, slide in enumerate(config.get("slides", [])):
        data = slide.get("data", {})
        layout = slide.get("layout", "")

        if layout == "chart_page":
            chart = data.get("chart", {})
            _validate_chart(chart, f"slides[{i}].data.chart")

        for j, comp in enumerate(data.get("components", [])):
            _validate_component(comp, f"slides[{i}].data.components[{j}]")

        for key in ("left_components", "right_components"):
            for j, comp in enumerate(data.get(key, [])):
                _validate_component(comp, f"slides[{i}].data.{key}[{j}]")


def _validate_chart(chart: dict, path: str) -> None:
    chart_type = chart.get("type", "")
    data = chart.get("data", {})

    if chart_type in ("pie", "waterfall"):
        if "series" in data:
            raise ConfigValidationError(
                f"[{path}] type='{chart_type}' では data.series は使えません。"
                f" data.labels と data.values を使ってください。"
            )
        if "values" not in data or "labels" not in data:
            raise ConfigValidationError(
                f"[{path}] type='{chart_type}' には data.labels と data.values が必要です。"
            )
        labels = data.get("labels", [])
        values = data.get("values", [])
        if len(labels) != len(values):
            raise ConfigValidationError(
                f"[{path}] labels の長さ ({len(labels)}) と values の長さ ({len(values)}) が一致しません。"
            )

    if chart_type in ("bar", "line", "stacked_bar", "area"):
        if "series" not in data:
            raise ConfigValidationError(
                f"[{path}] type='{chart_type}' には data.series が必要です。"
            )
        labels = data.get("labels", [])
        for k, series in enumerate(data.get("series", [])):
            vals = series.get("values", [])
            if len(vals) != len(labels):
                raise ConfigValidationError(
                    f"[{path}].data.series[{k}] の values の長さ ({len(vals)}) が"
                    f" labels の長さ ({len(labels)}) と一致しません。"
                )

    if chart_type == "scatter":
        series = data.get("series", [])
        if not series:
            raise ConfigValidationError(
                f"[{path}] type='scatter' には data.series が必要です。"
            )
        for k, s in enumerate(series):
            pts = s.get("points")
            if pts is None:
                raise ConfigValidationError(
                    f"[{path}].data.series[{k}] には points (XY ペア配列) が必要です。"
                )
            for p, pt in enumerate(pts):
                if not (isinstance(pt, (list, tuple)) and len(pt) == 2):
                    raise ConfigValidationError(
                        f"[{path}].data.series[{k}].points[{p}] は [x, y] の2要素配列で指定してください。"
                    )

    if chart_type == "combo":
        bars = data.get("bars", [])
        lines = data.get("lines", [])
        if not bars and not lines:
            raise ConfigValidationError(
                f"[{path}] type='combo' には data.bars または data.lines のいずれかが必要です。"
            )
        labels = data.get("labels", [])
        if not labels:
            raise ConfigValidationError(f"[{path}] type='combo' には data.labels が必要です。")
        for group_key in ("bars", "lines"):
            for k, s in enumerate(data.get(group_key, [])):
                vals = s.get("values", [])
                if len(vals) != len(labels):
                    raise ConfigValidationError(
                        f"[{path}].data.{group_key}[{k}] の values の長さ ({len(vals)}) が"
                        f" labels の長さ ({len(labels)}) と一致しません。"
                    )

    if chart_type in ("pie", "waterfall") and chart.get("annotations"):
        raise ConfigValidationError(
            f"[{path}] annotations は bar / line / stacked_bar / area / combo でのみ使用できます。"
        )

    if chart_type == "scatter" and chart.get("annotations"):
        raise ConfigValidationError(
            f"[{path}] annotations は scatter では使用できません (XY軸でカテゴリ位置が定まらないため)。"
        )


def _validate_component(comp: dict, path: str) -> None:
    comp_type = comp.get("type", "")

    if comp_type == "kpi_cards":
        for k, card in enumerate(comp.get("cards", [])):
            has_delta = "delta" in card
            has_dir = "delta_direction" in card
            if has_delta != has_dir:
                raise ConfigValidationError(
                    f"[{path}.cards[{k}]] delta と delta_direction は両方セットで指定してください。"
                )

    if comp_type == "heatmap":
        col_h = comp.get("col_headers", [])
        row_h = comp.get("row_headers", [])
        values = comp.get("values", [])
        if len(values) != len(row_h):
            raise ConfigValidationError(
                f"[{path}] values の行数 ({len(values)}) が row_headers の長さ ({len(row_h)}) と一致しません。"
            )
        for r, row in enumerate(values):
            if len(row) != len(col_h):
                raise ConfigValidationError(
                    f"[{path}].values[{r}] の列数 ({len(row)}) が col_headers の長さ ({len(col_h)}) と一致しません。"
                )

    if comp_type == "matrix_2x2":
        quads = comp.get("quadrants", [])
        if quads and len(quads) != 4:
            raise ConfigValidationError(
                f"[{path}] quadrants は4要素 (左上/右上/左下/右下) でなければなりません。現在: {len(quads)}要素。"
            )

    if comp_type == "swot":
        cells = comp.get("cells", [])
        if cells and len(cells) != 4:
            raise ConfigValidationError(
                f"[{path}] cells は4要素 (左上/右上/左下/右下) でなければなりません。現在: {len(cells)}要素。"
            )
