"""Linter (オーバーフロー / 推奨範囲逸脱) のテスト。"""
from src.linter import lint_config


def _wrap(layout: str, data: dict) -> dict:
    return {"theme": "monotone", "slides": [{"layout": layout, "data": data}]}


class TestNoWarnings:
    def test_clean_config_no_warnings(self):
        config = _wrap("content", {
            "title": "簡潔なタイトル",
            "columns": 1,
            "components": [
                {"type": "bullets", "items": ["A", "B", "C"]},
                {"type": "pillars", "items": [
                    {"title": "X", "body": "x"},
                    {"title": "Y", "body": "y"},
                    {"title": "Z", "body": "z"},
                ]},
            ],
        })
        assert lint_config(config) == []


class TestTitleOverflow:
    def test_long_title_warns(self):
        long = "あ" * 50
        warnings = lint_config(_wrap("content", {"title": long, "columns": 1, "components": []}))
        assert any("title" in w for w in warnings)

    def test_long_subtitle_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "短い",
            "subtitle": "あ" * 80,
            "columns": 1,
            "components": [],
        }))
        assert any("subtitle" in w for w in warnings)


class TestBullets:
    def test_too_many_bullets_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "bullets", "items": [f"item{i}" for i in range(7)]}],
        }))
        assert any("bullets" in w and "7" in w for w in warnings)

    def test_long_bullet_line_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "bullets", "items": ["あ" * 100]}],
        }))
        assert any("100" in w for w in warnings)


class TestPillars:
    def test_two_pillars_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "pillars", "items": [
                {"title": "A", "body": "a"},
                {"title": "B", "body": "b"},
            ]}],
        }))
        assert any("pillars" in w for w in warnings)

    def test_six_pillars_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "pillars", "items": [
                {"title": f"P{i}", "body": "x"} for i in range(6)
            ]}],
        }))
        assert any("pillars" in w for w in warnings)


class TestKpiCards:
    def test_kpi_delta_without_direction_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "kpi_cards", "cards": [
                {"value": "100", "unit": "%", "label": "X", "delta": "+5%"},
                {"value": "50", "unit": "%", "label": "Y"},
            ]}],
        }))
        assert any("delta_direction" in w for w in warnings)


class TestChartSource:
    def test_chart_page_without_source_warns(self):
        warnings = lint_config(_wrap("chart_page", {
            "title": "売上推移",
            "chart": {
                "type": "bar",
                "data": {"labels": ["1", "2"], "series": [{"name": "s", "values": [1, 2]}]},
            },
        }))
        assert any("source" in w for w in warnings)

    def test_chart_page_with_source_no_warning(self):
        warnings = lint_config(_wrap("chart_page", {
            "title": "売上推移",
            "source": "社内データ 2026年3月",
            "chart": {
                "type": "bar",
                "data": {"labels": ["1", "2"], "series": [{"name": "s", "values": [1, 2]}]},
            },
        }))
        assert not any("source" in w for w in warnings)


class TestTable:
    def test_too_many_rows_warns(self):
        warnings = lint_config(_wrap("content", {
            "title": "T", "columns": 1,
            "components": [{"type": "table",
                            "headers": ["A", "B"],
                            "rows": [["x", "y"]] * 10}],
        }))
        assert any("table" in w and "10" in w for w in warnings)


class TestGeneratorIntegration:
    def test_generator_emits_warnings_to_stderr(self, tmp_path, capsys):
        from src.generator import generate_pptx
        config = _wrap("content", {
            "title": "あ" * 50,
            "columns": 1,
            "components": [],
        })
        out = tmp_path / "out.pptx"
        generate_pptx(config, str(out))
        captured = capsys.readouterr()
        assert "Lint warnings" in captured.err
        assert "title" in captured.err

    def test_generator_lint_disabled(self, tmp_path, capsys):
        from src.generator import generate_pptx
        config = _wrap("content", {
            "title": "あ" * 50,
            "columns": 1,
            "components": [],
        })
        out = tmp_path / "out.pptx"
        generate_pptx(config, str(out), lint=False)
        captured = capsys.readouterr()
        assert "Lint warnings" not in captured.err
