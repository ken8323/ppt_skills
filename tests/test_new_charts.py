"""新規チャート (stacked_bar / area / scatter / combo) の生成・検証テスト。"""
import pytest

from src.generator import generate_pptx
from src.validator import validate_config, ConfigValidationError


def _wrap_chart(chart):
    return {
        "theme": "monotone",
        "slides": [{
            "layout": "chart_page",
            "data": {"title": "T", "source": "x", "chart": chart},
        }],
    }


class TestStackedBar:
    def test_generates(self, tmp_path):
        config = _wrap_chart({
            "type": "stacked_bar", "unit": "億円",
            "data": {"labels": ["A", "B"],
                     "series": [{"name": "x", "values": [1, 2]},
                                {"name": "y", "values": [3, 4]}]},
        })
        out = tmp_path / "out.pptx"
        generate_pptx(config, str(out))
        assert out.exists()

    def test_horizontal(self, tmp_path):
        config = _wrap_chart({
            "type": "stacked_bar", "horizontal": True,
            "data": {"labels": ["A"], "series": [{"name": "x", "values": [10]}]},
        })
        generate_pptx(config, str(tmp_path / "h.pptx"))

    def test_length_mismatch_raises(self):
        with pytest.raises(ConfigValidationError, match="一致"):
            validate_config(_wrap_chart({
                "type": "stacked_bar",
                "data": {"labels": ["A", "B"],
                         "series": [{"name": "x", "values": [1]}]},
            }))


class TestArea:
    def test_generates(self, tmp_path):
        generate_pptx(_wrap_chart({
            "type": "area",
            "data": {"labels": ["1", "2", "3"],
                     "series": [{"name": "x", "values": [1, 2, 3]}]},
        }), str(tmp_path / "a.pptx"))

    def test_stacked(self, tmp_path):
        generate_pptx(_wrap_chart({
            "type": "area", "stacked": True,
            "data": {"labels": ["A", "B"],
                     "series": [{"name": "x", "values": [1, 2]},
                                {"name": "y", "values": [3, 4]}]},
        }), str(tmp_path / "as.pptx"))


class TestScatter:
    def test_generates(self, tmp_path):
        generate_pptx(_wrap_chart({
            "type": "scatter", "x_label": "X", "y_label": "Y",
            "data": {"series": [{"name": "s", "points": [[1, 2], [3, 4]]}]},
        }), str(tmp_path / "s.pptx"))

    def test_missing_points_raises(self):
        with pytest.raises(ConfigValidationError, match="points"):
            validate_config(_wrap_chart({
                "type": "scatter",
                "data": {"series": [{"name": "s"}]},
            }))

    def test_invalid_point_format_raises(self):
        with pytest.raises(ConfigValidationError, match=r"\[x, y\]"):
            validate_config(_wrap_chart({
                "type": "scatter",
                "data": {"series": [{"name": "s", "points": [[1, 2, 3]]}]},
            }))

    def test_annotations_rejected(self):
        with pytest.raises(ConfigValidationError, match="annotations"):
            validate_config(_wrap_chart({
                "type": "scatter",
                "data": {"series": [{"name": "s", "points": [[1, 2]]}]},
                "annotations": [{"category": "A", "text": "x"}],
            }))


class TestCombo:
    def test_generates_with_secondary_axis(self, tmp_path):
        generate_pptx(_wrap_chart({
            "type": "combo",
            "data": {
                "labels": ["2023", "2024", "2025"],
                "bars": [{"name": "売上", "values": [100, 150, 220], "unit": "億円"}],
                "lines": [{"name": "成長率", "values": [10, 50, 47], "unit": "%", "secondary_axis": True}],
            },
        }), str(tmp_path / "c.pptx"))

    def test_generates_without_secondary(self, tmp_path):
        generate_pptx(_wrap_chart({
            "type": "combo",
            "data": {
                "labels": ["A", "B"],
                "bars": [{"name": "b1", "values": [10, 20]}],
                "lines": [{"name": "l1", "values": [5, 15]}],
            },
        }), str(tmp_path / "c2.pptx"))

    def test_missing_labels_raises(self):
        with pytest.raises(ConfigValidationError, match="labels"):
            validate_config(_wrap_chart({
                "type": "combo",
                "data": {"bars": [{"name": "b", "values": [1]}]},
            }))

    def test_length_mismatch_raises(self):
        with pytest.raises(ConfigValidationError, match="一致"):
            validate_config(_wrap_chart({
                "type": "combo",
                "data": {
                    "labels": ["A", "B"],
                    "bars": [{"name": "b", "values": [1, 2, 3]}],
                },
            }))

    def test_empty_data_raises(self):
        with pytest.raises(ConfigValidationError, match="bars"):
            validate_config(_wrap_chart({
                "type": "combo",
                "data": {"labels": ["A"]},
            }))
