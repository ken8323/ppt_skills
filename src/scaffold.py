"""ストーリーテンプレートからスライド骨子 (validate を pass する config dict) を返す。

Claude のフェーズ2 (構成提案) を半自動化する。骨子を受け取って、
プレースホルダ (角括弧 [XXX]) を実際の内容に置き換えてから generate_pptx に渡す。

使い方:
    from src.scaffold import scaffold, list_templates

    config = scaffold("monthly_report", theme="dark", brand_name="営業部")
    # config をユーザー要件に合わせて編集
    generate_pptx(config, "/tmp/out.pptx")
"""
from __future__ import annotations

from copy import deepcopy


def list_templates() -> list[str]:
    """利用可能なテンプレート名の一覧。"""
    return sorted(_TEMPLATES.keys())


def template_info(name: str) -> dict:
    """テンプレートのメタ情報 (用途・推奨枚数)。"""
    if name not in _TEMPLATES:
        raise ValueError(f"unknown template '{name}'. available: {list_templates()}")
    info = _TEMPLATE_INFO[name]
    return {
        "name": name,
        "description": info["description"],
        "slide_count": len(_TEMPLATES[name]()["slides"]),
        "story_arc": info["story_arc"],
    }


def scaffold(template: str, theme: str = "monotone",
             footer: str = "", brand_name: str = "",
             title: str = "", client: str = "", date: str = "") -> dict:
    """テンプレート名から骨子 config を生成。

    Args:
        template:   テンプレート名 (list_templates() で確認)
        theme:      "monotone" | "dark" | "colorful"
        footer:     全スライド共通のフッター文字列
        brand_name: 表紙左パネル + フッターのブランド表示
        title:      表紙タイトル (未指定時はテンプレートのデフォルト)
        client:     表紙の対象読者
        date:       表紙の日付
    """
    if template not in _TEMPLATES:
        raise ValueError(f"unknown template '{template}'. available: {list_templates()}")

    config = _TEMPLATES[template]()
    config["theme"] = theme
    if footer:
        config["footer"] = footer
    if brand_name:
        config["brand_name"] = brand_name

    # 表紙データの上書き
    if config["slides"] and config["slides"][0]["layout"] == "cover":
        cover_data = config["slides"][0]["data"]
        if title:
            cover_data["title"] = title
        if client:
            cover_data["client"] = client
        if date:
            cover_data["date"] = date

    return config


# ---------- テンプレート定義 ----------

def _consulting_proposal() -> dict:
    return {
        "theme": "monotone",
        "slides": [
            {"layout": "cover", "data": {
                "title": "[提案書タイトル]",
                "subtitle": "[サブタイトル: 対象期間や領域]",
                "client": "[クライアント名]",
                "date": "[YYYY年M月]",
            }},
            {"layout": "agenda", "data": {
                "items": ["現状分析", "課題整理", "解決策の提案", "期待効果", "実行計画"],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 1, "section_title": "現状分析",
            }},
            {"layout": "chart_page", "data": {
                "title": "[現状を端的に表すアクションタイトル]",
                "subtitle": "[根拠数値や背景を1行で]",
                "source": "[出典: 社内データ / 調査機関 / 年度]",
                "chart": {
                    "type": "bar", "unit": "[単位]",
                    "data": {
                        "labels": ["[期間1]", "[期間2]", "[期間3]"],
                        "series": [{"name": "[系列名]", "values": [0, 0, 0]}],
                    },
                    "annotations": [{"category": "[期間3]", "text": "[着目点]"}],
                },
                "key_points": ["[要点1]", "[要点2]", "[要点3]"],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 2, "section_title": "課題整理",
            }},
            {"layout": "content", "data": {
                "title": "[最重要課題のアクションタイトル]",
                "subtitle": "[影響度・緊急度の整理]",
                "columns": 1,
                "components": [
                    {"type": "matrix_2x2",
                     "x_axis": "影響度", "y_axis": "緊急度",
                     "quadrants": ["[左上: 高緊急/低影響]", "[右上: 高緊急/高影響]",
                                   "[左下: 低緊急/低影響]", "[右下: 低緊急/高影響]"],
                     "recommended_quadrant": 1},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 3, "section_title": "解決策の提案",
            }},
            {"layout": "content", "data": {
                "title": "[提案の3本柱を端的に]",
                "subtitle": "[3軸の関係性を1行で]",
                "columns": 1,
                "components": [
                    {"type": "pillars", "items": [
                        {"title": "[柱1]", "body": "[施策の中身]", "kpi": "[時期/規模]"},
                        {"title": "[柱2]", "body": "[施策の中身]", "kpi": "[時期/規模]"},
                        {"title": "[柱3]", "body": "[施策の中身]", "kpi": "[時期/規模]"},
                    ]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 4, "section_title": "期待効果",
            }},
            {"layout": "content", "data": {
                "title": "[効果を数値で表すアクションタイトル]",
                "subtitle": "[投資回収期間 / ROI]",
                "columns": 1,
                "components": [
                    {"type": "kpi_cards", "cards": [
                        {"value": "[数値]", "unit": "[単位]", "label": "[指標1]",
                         "delta": "[差分]", "delta_direction": "up"},
                        {"value": "[数値]", "unit": "[単位]", "label": "[指標2]",
                         "delta": "[差分]", "delta_direction": "up"},
                        {"value": "[数値]", "unit": "[単位]", "label": "[指標3]",
                         "delta": "[差分]", "delta_direction": "up"},
                    ]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 5, "section_title": "実行計画",
            }},
            {"layout": "content", "data": {
                "title": "[実行スケジュールのアクションタイトル]",
                "subtitle": "[マイルストーン]",
                "columns": 1,
                "components": [
                    {"type": "gantt",
                     "phases": ["Q1", "Q2", "Q3", "Q4"],
                     "tasks": [
                         {"name": "[タスク1]", "start": 0, "duration": 2, "progress": 0.0},
                         {"name": "[タスク2]", "start": 1, "duration": 2, "progress": 0.0},
                         {"name": "[タスク3]", "start": 0, "duration": 4, "progress": 0.0},
                     ]},
                ],
            }},
            {"layout": "closing", "data": {
                "summary": ["[要点1]", "[要点2]", "[要点3]"],
                "next_steps": ["[次の一歩1]", "[次の一歩2]", "[次の一歩3]"],
            }},
        ],
    }


def _monthly_report() -> dict:
    return {
        "theme": "monotone",
        "slides": [
            {"layout": "cover", "data": {
                "title": "[部門名] 月次報告",
                "subtitle": "[YYYY年M月度]",
                "client": "[報告先]",
                "date": "[YYYY年M月D日]",
            }},
            {"layout": "content", "data": {
                "title": "[今月の総括: 計画比/前月比のアクションタイトル]",
                "subtitle": "[要因の1行サマリ]",
                "columns": 1,
                "components": [
                    {"type": "kpi_cards", "cards": [
                        {"value": "[数値]", "unit": "[単位]", "label": "[KPI1]",
                         "delta": "[差分]", "delta_direction": "up"},
                        {"value": "[数値]", "unit": "[単位]", "label": "[KPI2]",
                         "delta": "[差分]", "delta_direction": "up"},
                        {"value": "[数値]", "unit": "[単位]", "label": "[KPI3]",
                         "delta": "[差分]", "delta_direction": "up"},
                    ]},
                ],
            }},
            {"layout": "chart_page", "data": {
                "title": "[トレンドのアクションタイトル]",
                "subtitle": "[何が牽引しているか]",
                "source": "[出典: 社内CRM / ERP 等]",
                "chart": {
                    "type": "line", "unit": "[単位]",
                    "data": {
                        "labels": ["[月-2]", "[月-1]", "[当月]"],
                        "series": [{"name": "[系列1]", "values": [0, 0, 0]}],
                    },
                    "annotations": [{"category": "[当月]", "text": "[着目点]"}],
                },
                "key_points": ["[要点1]", "[要点2]"],
            }},
            {"layout": "content", "data": {
                "title": "[進行中案件のアクションタイトル]",
                "subtitle": "[今月確定見込み件数]",
                "columns": 1,
                "components": [
                    {"type": "table",
                     "headers": ["案件名", "金額", "ステータス", "受注/完了時期"],
                     "rows": [
                         ["[案件1]", "[金額]", "[ステータス]", "[時期]"],
                         ["[案件2]", "[金額]", "[ステータス]", "[時期]"],
                         ["[案件3]", "[金額]", "[ステータス]", "[時期]"],
                     ],
                     "highlight_rows": [0]},
                ],
            }},
            {"layout": "comparison", "data": {
                "title": "[光と影のアクションタイトル]",
                "subtitle": "[全体評価の1行]",
                "left_title": "順調な点",
                "left_components": [
                    {"type": "bullets", "items": ["[好調1]", "[好調2]", "[好調3]"]},
                    {"type": "callout", "text": "[ポジティブな結論]", "variant": "success"},
                ],
                "right_title": "課題",
                "right_components": [
                    {"type": "bullets", "items": ["[課題1]", "[課題2]", "[課題3]"]},
                    {"type": "callout", "text": "[アラート]", "variant": "warning"},
                ],
            }},
            {"layout": "content", "data": {
                "title": "[来月の重点アクションのアクションタイトル]",
                "subtitle": "[3つの並行アクション]",
                "columns": 1,
                "components": [
                    {"type": "process_flow",
                     "steps": ["[アクション1]", "[アクション2]", "[アクション3]"]},
                ],
            }},
            {"layout": "closing", "data": {
                "summary": ["[今月の総括1]", "[今月の総括2]", "[今月の総括3]"],
                "next_steps": ["[来月の一歩1]", "[来月の一歩2]", "[来月の一歩3]"],
            }},
        ],
    }


def _project_kickoff() -> dict:
    return {
        "theme": "colorful",
        "slides": [
            {"layout": "cover", "data": {
                "title": "[プロジェクト名] キックオフ",
                "subtitle": "[サブタイトル: スコープ概要]",
                "client": "[ステアリングコミッティ向け 等]",
                "date": "[YYYY年M月]",
            }},
            {"layout": "agenda", "data": {
                "items": ["背景と目的", "スコープ", "体制", "スケジュール", "リスクと対応"],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 1, "section_title": "背景と目的",
            }},
            {"layout": "content", "data": {
                "title": "[なぜ今このプロジェクトかのアクションタイトル]",
                "subtitle": "[トリガーとなった事象]",
                "columns": 1,
                "components": [
                    {"type": "callout", "text": "[本質課題の一文]", "variant": "info"},
                    {"type": "bullets", "items": ["[背景1]", "[背景2]", "[背景3]"]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 2, "section_title": "スコープ",
            }},
            {"layout": "comparison", "data": {
                "title": "[スコープを明確に: In/Out のアクションタイトル]",
                "subtitle": "[線引きの基準]",
                "left_title": "対象 (In Scope)",
                "left_components": [
                    {"type": "bullets", "items": ["[対象1]", "[対象2]", "[対象3]"]},
                ],
                "right_title": "対象外 (Out of Scope)",
                "right_components": [
                    {"type": "bullets", "items": ["[対象外1]", "[対象外2]", "[対象外3]"]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 3, "section_title": "体制",
            }},
            {"layout": "content", "data": {
                "title": "[体制構成のアクションタイトル]",
                "subtitle": "[役割分担の1行サマリ]",
                "columns": 1,
                "components": [
                    {"type": "org_chart", "data": {
                        "name": "[プロジェクトオーナー]",
                        "children": [
                            {"name": "[PM]", "children": [
                                {"name": "[サブチーム1リーダー]"},
                                {"name": "[サブチーム2リーダー]"},
                            ]},
                        ],
                    }},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 4, "section_title": "スケジュール",
            }},
            {"layout": "content", "data": {
                "title": "[全体スケジュールのアクションタイトル]",
                "subtitle": "[マイルストーン]",
                "columns": 1,
                "components": [
                    {"type": "gantt",
                     "phases": ["Q1", "Q2", "Q3", "Q4"],
                     "tasks": [
                         {"name": "[フェーズ1: 計画]", "start": 0, "duration": 1, "progress": 0.0},
                         {"name": "[フェーズ2: 構築]", "start": 1, "duration": 2, "progress": 0.0},
                         {"name": "[フェーズ3: 展開]", "start": 3, "duration": 1, "progress": 0.0},
                     ]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 5, "section_title": "リスクと対応",
            }},
            {"layout": "content", "data": {
                "title": "[主要リスクの俯瞰アクションタイトル]",
                "subtitle": "[影響度×発生確率での評価]",
                "columns": 1,
                "components": [
                    {"type": "table",
                     "headers": ["リスク", "影響度", "発生確率", "対応策"],
                     "rows": [
                         ["[リスク1]", "[高/中/低]", "[高/中/低]", "[対応策1]"],
                         ["[リスク2]", "[高/中/低]", "[高/中/低]", "[対応策2]"],
                         ["[リスク3]", "[高/中/低]", "[高/中/低]", "[対応策3]"],
                     ]},
                ],
            }},
            {"layout": "closing", "data": {
                "summary": ["[キックオフ要点1]", "[キックオフ要点2]", "[キックオフ要点3]"],
                "next_steps": ["[直近の一歩1]", "[直近の一歩2]", "[直近の一歩3]"],
            }},
        ],
    }


def _briefing() -> dict:
    return {
        "theme": "colorful",
        "slides": [
            {"layout": "cover", "data": {
                "title": "[ブリーフィングタイトル]",
                "subtitle": "[サブタイトル: 主要メッセージ]",
                "client": "[対象読者]",
                "date": "[YYYY年M月]",
            }},
            {"layout": "agenda", "data": {
                "items": ["エグゼクティブサマリ", "市場/業界トレンド", "事例", "示唆と提言"],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 1, "section_title": "エグゼクティブサマリ",
            }},
            {"layout": "content", "data": {
                "title": "[3〜4本の主要潮流を端的にまとめるアクションタイトル]",
                "subtitle": "[なぜ今これが重要か]",
                "columns": 1,
                "components": [
                    {"type": "pillars", "items": [
                        {"title": "[潮流1]", "body": "[端的な説明]"},
                        {"title": "[潮流2]", "body": "[端的な説明]"},
                        {"title": "[潮流3]", "body": "[端的な説明]"},
                        {"title": "[潮流4]", "body": "[端的な説明]"},
                    ]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 2, "section_title": "市場/業界トレンド",
            }},
            {"layout": "content", "data": {
                "title": "[トレンドの俯瞰アクションタイトル]",
                "subtitle": "[ベンチマーク観点]",
                "columns": 1,
                "sources": ["[出典1]", "[出典2]"],
                "components": [
                    {"type": "table",
                     "headers": ["[項目]", "[列1]", "[列2]", "[列3]"],
                     "rows": [
                         ["[行1]", "[値]", "[値]", "[値]"],
                         ["[行2]", "[値]", "[値]", "[値]"],
                         ["[行3]", "[値]", "[値]", "[値]"],
                     ]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 3, "section_title": "事例",
            }},
            {"layout": "content", "data": {
                "title": "[事例から見えるパターンのアクションタイトル]",
                "subtitle": "[共通する成功要因]",
                "columns": 1,
                "sources": ["[出典 (例: 各社IR)]"],
                "components": [
                    {"type": "table",
                     "headers": ["企業/組織", "取組内容", "成果"],
                     "rows": [
                         ["[事例1]", "[内容]", "[成果]"],
                         ["[事例2]", "[内容]", "[成果]"],
                         ["[事例3]", "[内容]", "[成果]"],
                     ],
                     "highlight_rows": [0]},
                ],
            }},
            {"layout": "section_divider", "data": {
                "section_number": 4, "section_title": "示唆と提言",
            }},
            {"layout": "content", "data": {
                "title": "[自社が取るべきアクションのアクションタイトル]",
                "subtitle": "[段階的アプローチの方針]",
                "columns": 1,
                "components": [
                    {"type": "process_flow",
                     "steps": ["[ステップ1]", "[ステップ2]", "[ステップ3]", "[ステップ4]"]},
                ],
            }},
            {"layout": "closing", "data": {
                "summary": ["[要点1]", "[要点2]", "[要点3]"],
                "next_steps": ["[次の一歩1]", "[次の一歩2]", "[次の一歩3]"],
            }},
        ],
    }


_TEMPLATES = {
    "consulting_proposal": _consulting_proposal,
    "monthly_report": _monthly_report,
    "project_kickoff": _project_kickoff,
    "briefing": _briefing,
}

_TEMPLATE_INFO = {
    "consulting_proposal": {
        "description": "コンサル提案書 (現状→課題→解決策→効果→計画)",
        "story_arc": ["現状分析", "課題整理", "解決策", "期待効果", "実行計画"],
    },
    "monthly_report": {
        "description": "部門月次報告 (KPI主体・進捗・課題・次月計画)",
        "story_arc": ["総括KPI", "トレンド", "進行案件", "光と影", "次月アクション"],
    },
    "project_kickoff": {
        "description": "プロジェクトキックオフ (背景→スコープ→体制→計画→リスク)",
        "story_arc": ["背景と目的", "スコープ", "体制", "スケジュール", "リスクと対応"],
    },
    "briefing": {
        "description": "業界・社内ブリーフィング (サマリ→トレンド→事例→提言)",
        "story_arc": ["エグゼサマリ", "業界トレンド", "事例", "示唆と提言"],
    },
}
