"""table コンポーネントの拡張機能 (align / totals_row / banded / col_widths_ratio) のテスト。"""
import pytest

from src.generator import generate_pptx
from src.validator import validate_config, ConfigValidationError


def _wrap_table(table_comp, headers=None, rows=None):
    base = {"type": "table"}
    if headers is not None:
        base["headers"] = headers
    if rows is not None:
        base["rows"] = rows
    base.update(table_comp)
    return {
        "theme": "monotone",
        "slides": [{"layout": "content", "data": {
            "title": "T", "columns": 1,
            "components": [base],
        }}],
    }


class TestAlign:
    def test_string_align_valid(self, tmp_path):
        config = _wrap_table(
            {"align": "right"},
            headers=["A", "B"], rows=[["1", "2"]],
        )
        validate_config(config)
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_list_align_valid(self, tmp_path):
        config = _wrap_table(
            {"align": ["left", "right", "center"]},
            headers=["A", "B", "C"], rows=[["x", "1", "y"]],
        )
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_invalid_align_string_raises(self):
        config = _wrap_table(
            {"align": "middle"},
            headers=["A"], rows=[["x"]],
        )
        with pytest.raises(ConfigValidationError, match="align"):
            validate_config(config)

    def test_align_length_mismatch_raises(self):
        config = _wrap_table(
            {"align": ["left", "right"]},
            headers=["A", "B", "C"], rows=[["x", "y", "z"]],
        )
        with pytest.raises(ConfigValidationError, match="一致"):
            validate_config(config)


class TestTotalsRow:
    def test_auto_totals_numeric(self, tmp_path):
        config = _wrap_table(
            {"totals_row": True},
            headers=["項目", "金額"], rows=[["A", 100], ["B", 250]],
        )
        validate_config(config)
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_auto_totals_with_comma_strings(self, tmp_path):
        config = _wrap_table(
            {"totals_row": True},
            headers=["商品", "売上"], rows=[["X", "1,200"], ["Y", "3,400"]],
        )
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_explicit_totals_row(self, tmp_path):
        config = _wrap_table(
            {"totals_row": ["合計", "350", "—"]},
            headers=["項目", "金額", "備考"],
            rows=[["A", 100, "ok"], ["B", 250, "ok"]],
        )
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_totals_length_mismatch_raises(self):
        config = _wrap_table(
            {"totals_row": ["合計", "100"]},
            headers=["A", "B", "C"],
            rows=[["x", "y", "z"]],
        )
        with pytest.raises(ConfigValidationError, match="totals_row"):
            validate_config(config)


class TestBanded:
    def test_banded_disabled(self, tmp_path):
        config = _wrap_table(
            {"banded": False},
            headers=["A"], rows=[["x"], ["y"], ["z"]],
        )
        generate_pptx(config, str(tmp_path / "o.pptx"))


class TestColWidthsRatio:
    def test_valid_ratio(self, tmp_path):
        config = _wrap_table(
            {"col_widths_ratio": [3, 1, 1]},
            headers=["項目", "数値1", "数値2"],
            rows=[["長めのテキスト", "100", "200"]],
        )
        validate_config(config)
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_length_mismatch_raises(self):
        config = _wrap_table(
            {"col_widths_ratio": [1, 2]},
            headers=["A", "B", "C"], rows=[["x", "y", "z"]],
        )
        with pytest.raises(ConfigValidationError, match="col_widths_ratio"):
            validate_config(config)

    def test_negative_value_raises(self):
        config = _wrap_table(
            {"col_widths_ratio": [1, -1]},
            headers=["A", "B"], rows=[["x", "y"]],
        )
        with pytest.raises(ConfigValidationError, match="正の数値"):
            validate_config(config)


class TestAutoNumericAlignment:
    def test_numeric_column_right_aligned_by_default(self, tmp_path):
        # 数値列が自動で right-align されることは描画ベースなので
        # ここでは生成が成功することのみ確認 (回帰防止)
        config = _wrap_table(
            {},
            headers=["項目", "売上", "増減"],
            rows=[["A", "100", "+5"], ["B", "200", "-3"]],
        )
        generate_pptx(config, str(tmp_path / "o.pptx"))
