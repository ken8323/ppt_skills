"""source / sources を全レイアウト共通フィールドとして扱うテスト。"""
import pytest

from src.generator import generate_pptx
from src.validator import validate_config, ConfigValidationError


def _wrap(layout, data):
    return {"theme": "monotone", "slides": [{"layout": layout, "data": data}]}


class TestSourceOnAllLayouts:
    @pytest.mark.parametrize("layout,data", [
        ("content", {"title": "T", "columns": 1, "components": [], "source": "社内データ 2026"}),
        ("comparison", {
            "title": "T", "left_title": "L", "right_title": "R",
            "left_components": [], "right_components": [],
            "source": "顧客アンケート 2026年3月",
        }),
        ("chart_page", {
            "title": "売上",
            "chart": {"type": "bar", "data": {
                "labels": ["A", "B"], "series": [{"name": "s", "values": [1, 2]}]}},
            "source": "経産省 2026",
        }),
        ("agenda", {"items": ["A", "B"], "source": "戦略本部資料"}),
        ("closing", {"summary": ["x"], "next_steps": ["y"], "source": "Board paper"}),
    ])
    def test_source_renders_on_layout(self, tmp_path, layout, data):
        config = _wrap(layout, data)
        validate_config(config)
        out = tmp_path / "o.pptx"
        generate_pptx(config, str(out))
        assert out.exists()


class TestSourcesList:
    def test_sources_string_list(self, tmp_path):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "sources": ["社内CRM 2026年3月", "Gartner 2025"],
        })
        validate_config(config)
        generate_pptx(config, str(tmp_path / "o.pptx"))

    def test_sources_label_url_dict(self, tmp_path):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "sources": [
                {"label": "Anthropic Press Release", "url": "https://example.com/a"},
                {"label": "Google Blog"},
            ],
        })
        validate_config(config)
        generate_pptx(config, str(tmp_path / "o.pptx"))


class TestValidationErrors:
    def test_both_source_and_sources_raises(self):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "source": "x", "sources": ["y"],
        })
        with pytest.raises(ConfigValidationError, match="同時指定"):
            validate_config(config)

    def test_source_non_string_raises(self):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "source": 12345,
        })
        with pytest.raises(ConfigValidationError, match="source"):
            validate_config(config)

    def test_sources_non_list_raises(self):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "sources": "should be list",
        })
        with pytest.raises(ConfigValidationError, match="list"):
            validate_config(config)

    def test_sources_dict_without_label_or_url_raises(self):
        config = _wrap("content", {
            "title": "T", "columns": 1, "components": [],
            "sources": [{"foo": "bar"}],
        })
        with pytest.raises(ConfigValidationError, match="label"):
            validate_config(config)


class TestChartPageBackwardCompat:
    """既存の chart_page.source も引き続き動くこと。"""
    def test_chart_page_source_still_works(self, tmp_path):
        config = _wrap("chart_page", {
            "title": "売上推移",
            "source": "社内財務 2026年3月",
            "chart": {"type": "bar", "data": {
                "labels": ["A", "B"], "series": [{"name": "s", "values": [1, 2]}]}},
        })
        out = tmp_path / "o.pptx"
        generate_pptx(config, str(out))
        assert out.exists()
