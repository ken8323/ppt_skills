"""拡張アイコンライブラリのテスト。"""
import pytest

from src.components.icon import ICON_SHAPES, list_icons
from src.generator import generate_pptx
from src.linter import lint_config


def _wrap_icons(items):
    return {
        "theme": "monotone",
        "slides": [{"layout": "content", "data": {
            "title": "T", "columns": 1,
            "components": [{"type": "icon_row", "items": items}],
        }}],
    }


class TestExpandedLibrary:
    def test_more_than_50_icons(self):
        assert len(ICON_SHAPES) >= 50

    def test_business_keywords_present(self):
        for kw in ["code", "user", "chart", "doc", "shield", "database",
                    "globe", "target", "clock", "ai", "rocket", "growth"]:
            assert kw in ICON_SHAPES

    def test_list_icons_returns_sorted(self):
        names = list_icons()
        assert names == sorted(names)


class TestRendering:
    def test_business_icons_generate(self, tmp_path):
        config = _wrap_icons([
            {"icon": "code", "label": "開発"},
            {"icon": "shield", "label": "セキュリティ"},
            {"icon": "database", "label": "DB"},
            {"icon": "rocket", "label": "成長"},
        ])
        generate_pptx(config, str(tmp_path / "biz.pptx"))

    def test_symbol_override(self, tmp_path):
        # 形は circle のままで、シンボルだけ ★ に上書き
        config = _wrap_icons([
            {"icon": "circle", "label": "強調", "symbol": "★"},
            {"icon": "square", "label": "標準"},
        ])
        generate_pptx(config, str(tmp_path / "sym.pptx"))


class TestLinterUnknownIcon:
    def test_unknown_icon_warns(self):
        warnings = lint_config(_wrap_icons([
            {"icon": "totally_made_up", "label": "X"},
            {"icon": "code", "label": "Y"},
            {"icon": "rocket", "label": "Z"},
        ]))
        assert any("totally_made_up" in w for w in warnings)
