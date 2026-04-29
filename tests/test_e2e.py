"""E2Eテスト: 実際のコンサル資料相当の全機能を組み合わせて生成。"""
from pathlib import Path

import pytest

from src.generator import generate_pptx


FULL_DECK_CONFIG = {
    "theme": "monotone",
    "footer": "株式会社ABC ホールディングス | 社外秘",
    "brand_name": "ABC コンサルティング",
    "slides": [
        {
            "layout": "cover",
            "data": {
                "title": "DX推進戦略提案書",
                "subtitle": "2026年度全社デジタル変革ロードマップ",
                "client": "株式会社ABC ホールディングス",
                "date": "2026年4月19日",
            },
        },
        {
            "layout": "agenda",
            "data": {
                "items": [
                    "エグゼクティブサマリー",
                    "現状分析と課題認識",
                    "戦略提案",
                    "実行計画とロードマップ",
                    "投資対効果",
                    "次のステップ",
                ],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 1, "section_title": "エグゼクティブサマリー"},
        },
        {
            "layout": "content",
            "data": {
                "title": "3つの戦略的提言",
                "subtitle": "3年間で業務効率30%改善、5.2億円の累計効果を実現する",
                "columns": 1,
                "components": [
                    {
                        "type": "kpi_cards",
                        "cards": [
                            {"value": "30", "unit": "%", "label": "業務効率改善目標",
                             "delta": "+8pt", "delta_direction": "up"},
                            {"value": "5.2", "unit": "億円", "label": "3年間累計効果",
                             "delta": "+1.2億", "delta_direction": "up"},
                            {"value": "18", "unit": "ヶ月", "label": "投資回収期間",
                             "delta": "-6ヶ月", "delta_direction": "down"},
                        ],
                    },
                    {
                        "type": "callout",
                        "text": "全社横断でのデータ基盤統合を優先し、短期成果と中期基盤の両立を図る",
                    },
                ],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 2, "section_title": "現状分析と課題認識"},
        },
        {
            "layout": "chart_page",
            "data": {
                "title": "業績推移と競合比較",
                "subtitle": "自社は年率7%成長だが競合は15-20%で市場シェアが縮小",
                "chart": {
                    "type": "bar",
                    "unit": "億円",
                    "data": {
                        "labels": ["2021", "2022", "2023", "2024", "2025"],
                        "series": [
                            {"name": "自社", "values": [120, 135, 142, 148, 155]},
                            {"name": "競合A", "values": [140, 160, 185, 210, 240]},
                            {"name": "競合B", "values": [110, 125, 150, 175, 205]},
                        ],
                    },
                    "annotations": [
                        {"category": "2025", "text": "競合比▲55%"},
                    ],
                },
                "key_points": [
                    "自社は年率約7%の成長",
                    "競合は15-20%の高成長",
                    "市場シェアは縮小傾向",
                    "差別化戦略の再定義が急務",
                ],
                "source": "各社有価証券報告書 (2021-2025年) を基に作成",
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "現状の4つの構造課題",
                "columns": 1,
                "components": [
                    {
                        "type": "matrix_2x2",
                        "x_axis": "実現の容易さ",
                        "y_axis": "インパクト",
                        "quadrants": [
                            "データ基盤統合\n業務自動化",
                            "AI活用\n新規事業開発",
                            "既存システム改善\nプロセス標準化",
                            "人材育成\nガバナンス強化",
                        ],
                    },
                ],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "課題の影響範囲",
                "columns": 2,
                "components": [
                    {
                        "type": "bullets",
                        "items": [
                            "レガシーシステムの運用負荷",
                            "部門別サイロ化したデータ",
                            "手作業中心のオペレーション",
                            "分析基盤の不足",
                        ],
                    },
                    {
                        "type": "bullets",
                        "items": [
                            "意思決定スピードの低下",
                            "顧客体験の競合優位喪失",
                            "人材採用における魅力度低下",
                            "コスト構造の硬直化",
                        ],
                    },
                ],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 3, "section_title": "戦略提案"},
        },
        {
            "layout": "content",
            "data": {
                "title": "DX推進の3本柱",
                "columns": 1,
                "components": [
                    {
                        "type": "pyramid",
                        "levels": ["価値創造", "業務変革", "基盤整備"],
                    },
                ],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "推進プロセス",
                "columns": 1,
                "components": [
                    {
                        "type": "process_flow",
                        "steps": ["現状把握", "戦略策定", "PoC実行", "本格展開", "継続改善"],
                    },
                ],
            },
        },
        {
            "layout": "comparison",
            "data": {
                "title": "Before / After: 目指す姿",
                "left_title": "Before (現状)",
                "left_components": [
                    {
                        "type": "bullets",
                        "items": [
                            "分断されたシステム群",
                            "手作業による集計",
                            "月次での業績把握",
                            "勘と経験による意思決定",
                        ],
                    },
                ],
                "right_title": "After (3年後)",
                "right_components": [
                    {
                        "type": "bullets",
                        "items": [
                            "統合データプラットフォーム",
                            "自動化されたレポーティング",
                            "リアルタイムダッシュボード",
                            "データドリブン経営",
                        ],
                    },
                ],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 4, "section_title": "実行計画とロードマップ"},
        },
        {
            "layout": "content",
            "data": {
                "title": "3年間のロードマップ",
                "columns": 1,
                "components": [
                    {
                        "type": "gantt",
                        "phases": ["2026", "2027", "2028"],
                        "tasks": [
                            {"name": "データ基盤構築", "start": 0, "duration": 4},
                            {"name": "業務自動化PoC", "start": 2, "duration": 3},
                            {"name": "全社展開", "start": 5, "duration": 4},
                            {"name": "AI活用", "start": 7, "duration": 5},
                        ],
                    },
                ],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "主要マイルストーン",
                "columns": 1,
                "components": [
                    {
                        "type": "timeline",
                        "milestones": [
                            {"date": "2026Q2", "label": "基盤設計完了"},
                            {"date": "2026Q4", "label": "PoC開始"},
                            {"date": "2027Q2", "label": "初期展開"},
                            {"date": "2027Q4", "label": "全社展開"},
                            {"date": "2028Q4", "label": "AI本格活用"},
                        ],
                    },
                ],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "投資配分計画",
                "columns": 1,
                "components": [
                    {
                        "type": "table",
                        "headers": ["領域", "2026", "2027", "2028", "合計"],
                        "rows": [
                            ["データ基盤", "80", "60", "40", "180"],
                            ["業務自動化", "40", "80", "60", "180"],
                            ["AI活用", "20", "40", "80", "140"],
                            ["人材育成", "30", "30", "30", "90"],
                            ["合計", "170", "210", "210", "590"],
                        ],
                    },
                ],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 5, "section_title": "投資対効果"},
        },
        {
            "layout": "chart_page",
            "data": {
                "title": "投資効果の累計推移",
                "subtitle": "初期投資17億を18ヶ月で回収、3年累計52億の効果",
                "chart": {
                    "type": "waterfall",
                    "data": {
                        "labels": ["初期投資", "2026効果", "2027効果", "2028効果", "累計"],
                        "values": [-170, 80, 180, 262, 352],
                    },
                },
                "key_points": [
                    "18ヶ月で投資回収",
                    "3年累計で5.2億円の効果",
                    "ROI 188%を達成",
                ],
                "source": "社内試算モデル (※サンプル値)",
            },
        },
        {
            "layout": "chart_page",
            "data": {
                "title": "効果内訳",
                "chart": {
                    "type": "pie",
                    "data": {
                        "labels": ["業務効率化", "売上拡大", "コスト削減", "その他"],
                        "values": [45, 25, 20, 10],
                    },
                },
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "DX推進の3本柱 (pillars)",
                "columns": 1,
                "components": [{
                    "type": "pillars",
                    "items": [
                        {"title": "基盤整備", "body": "データ統合\nシステム刷新", "kpi": "180M円"},
                        {"title": "業務変革", "body": "RPA導入\nプロセス自動化", "kpi": "180M円"},
                        {"title": "価値創造", "body": "AI活用\n新規事業開発", "kpi": "140M円"},
                    ],
                }],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "SWOT分析",
                "columns": 1,
                "components": [{
                    "type": "swot",
                    "cells": [
                        {"title": "強み", "items": ["ブランド力", "技術蓄積", "顧客基盤"]},
                        {"title": "弱み", "items": ["コスト構造の硬直化", "デジタル人材不足"]},
                        {"title": "機会", "items": ["AI普及", "規制緩和", "市場拡大"]},
                        {"title": "脅威", "items": ["競合参入加速", "景気悪化リスク"]},
                    ],
                }],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "競合ベンチマーク比較",
                "columns": 1,
                "components": [{
                    "type": "benchmark_bar",
                    "unit": "%",
                    "items": [
                        {"label": "自社", "value": 72, "is_self": True},
                        {"label": "競合A", "value": 88},
                        {"label": "競合B", "value": 81},
                        {"label": "業界平均", "value": 76},
                    ],
                }],
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "優先度×インパクト マトリクス",
                "columns": 1,
                "components": [{
                    "type": "heatmap",
                    "col_headers": ["コスト削減", "売上拡大", "顧客体験", "リスク低減"],
                    "row_headers": ["データ基盤", "業務自動化", "AI活用"],
                    "values": [
                        [90, 60, 70, 80],
                        [80, 50, 85, 60],
                        [50, 95, 90, 40],
                    ],
                }],
            },
        },
        {
            "layout": "section_divider",
            "data": {"section_number": 6, "section_title": "次のステップ"},
        },
        {
            "layout": "closing",
            "data": {
                "summary": [
                    "DX推進は全社的な経営アジェンダ",
                    "データ基盤を起点とした段階的変革",
                    "3年で590百万円の投資、5.2億円の効果",
                ],
                "next_steps": [
                    "経営会議での承認取得 (2026年5月)",
                    "プロジェクト体制の組成 (2026年6月)",
                    "キックオフと詳細設計開始 (2026年7月)",
                ],
            },
        },
        {
            "layout": "closing",
            "data": {
                "type": "thank_you",
                "contact": "contact@consulting-firm.co.jp",
            },
        },
    ],
}


@pytest.mark.parametrize("theme", ["monotone", "dark", "colorful"])
def test_full_deck_all_themes(tmp_path, theme):
    config = {**FULL_DECK_CONFIG, "theme": theme}
    output = tmp_path / f"full_{theme}.pptx"
    prs = generate_pptx(config, str(output))
    assert output.exists()
    assert len(prs.slides) == len(FULL_DECK_CONFIG["slides"])
    assert output.stat().st_size > 10_000  # 実ファイルサイズが最低限あること


def test_full_deck_writes_inspectable_artifact():
    """手動確認用の成果物を output/ に保存する (コミット対象外)。"""
    out_dir = Path(__file__).resolve().parent.parent / "output"
    out_dir.mkdir(exist_ok=True)
    for theme in ["monotone", "dark", "colorful"]:
        config = {**FULL_DECK_CONFIG, "theme": theme}
        output = out_dir / f"sample_{theme}.pptx"
        generate_pptx(config, str(output))
        assert output.exists()
