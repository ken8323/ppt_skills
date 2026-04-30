"""scaffold() 各テンプレートが validate を pass し、generate も成功することを担保。"""
import pytest

from src.generator import generate_pptx
from src.scaffold import scaffold, list_templates, template_info
from src.validator import validate_config


class TestListAndInfo:
    def test_list_returns_all_templates(self):
        names = list_templates()
        assert "consulting_proposal" in names
        assert "monthly_report" in names
        assert "project_kickoff" in names
        assert "briefing" in names

    def test_template_info_includes_metadata(self):
        info = template_info("monthly_report")
        assert info["name"] == "monthly_report"
        assert info["slide_count"] > 0
        assert "story_arc" in info
        assert isinstance(info["story_arc"], list)

    def test_unknown_template_info_raises(self):
        with pytest.raises(ValueError, match="unknown template"):
            template_info("nonexistent")


class TestScaffoldOutput:
    @pytest.mark.parametrize("name", list_templates())
    def test_scaffold_validates(self, name):
        config = scaffold(name)
        validate_config(config)

    @pytest.mark.parametrize("name", list_templates())
    def test_scaffold_generates(self, name, tmp_path):
        config = scaffold(name)
        out = tmp_path / f"{name}.pptx"
        generate_pptx(config, str(out), lint=False)
        assert out.exists()


class TestScaffoldOverrides:
    def test_theme_override(self):
        config = scaffold("monthly_report", theme="dark")
        assert config["theme"] == "dark"

    def test_brand_and_footer_override(self):
        config = scaffold("monthly_report",
                          brand_name="営業部", footer="社外秘")
        assert config["brand_name"] == "営業部"
        assert config["footer"] == "社外秘"

    def test_cover_overrides_applied(self):
        config = scaffold("consulting_proposal",
                          title="DX推進提案書", client="株式会社ABC", date="2026年4月")
        cover = config["slides"][0]["data"]
        assert cover["title"] == "DX推進提案書"
        assert cover["client"] == "株式会社ABC"
        assert cover["date"] == "2026年4月"

    def test_unknown_template_raises(self):
        with pytest.raises(ValueError, match="unknown template"):
            scaffold("nonexistent")


class TestEditableScaffold:
    """骨子を受け取って編集→生成の流れが意図通り動くこと。"""
    def test_edit_and_generate_e2e(self, tmp_path):
        config = scaffold("monthly_report", theme="dark")
        # cover の上書き
        config["slides"][0]["data"]["title"] = "営業部 月次報告"
        # KPIスライドのプレースホルダを実値に
        kpi_slide = config["slides"][1]
        kpi_slide["data"]["title"] = "売上は計画比 108%"
        kpi_slide["data"]["subtitle"] = "新規受注の拡大が牽引"
        kpi_slide["data"]["components"][0]["cards"] = [
            {"value": "12.5", "unit": "億円", "label": "月次売上",
             "delta": "+8%", "delta_direction": "up"},
        ]

        validate_config(config)
        out = tmp_path / "edited.pptx"
        generate_pptx(config, str(out), lint=False)
        assert out.exists()
