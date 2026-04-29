# ppt_skills

コンサルティングファームレベルの PowerPoint 資料を、JSON ライクな構成データから直接 `.pptx` として生成する Python スキルです。Claude Code の **skill** として動作することを前提に設計されており、`python-pptx` のみで完結するためテンプレートファイルや外部ツールは不要です。

## 特徴

- **テンプレート不要**: `python-pptx` で全要素を直接描画。フォント・色・レイアウトはコードで一貫管理
- **3 つのテーマ**: `monotone`（端正・フォーマル）/ `dark`（重厚・インパクト）/ `colorful`（モダン・カジュアル）
- **7 種のレイアウト**: 表紙・章区切り・アジェンダ・汎用コンテンツ・チャート・左右比較・クロージング
- **17 種のコンポーネント**: KPI カード、ピラミッド、プロセスフロー、サイクル、組織図、タイムライン、ガント、SWOT、ヒートマップ、ベンチマーク棒など、コンサル資料に頻出する図解を網羅
- **4 種のチャート**: 棒・折れ線・円・ウォーターフォール（ウォーターフォールは正/負値で自動配色、棒/折れ線は注記ピル対応）
- **アクションタイトル前提**: タイトル直下に根拠や数値を補足する `subtitle`（サブヘッド）を全主要レイアウトで対応
- **フッター/ページ番号自動注入**: トップレベル `footer` / `brand_name` を一度指定すれば全スライドに反映
- **スキーマ検証**: `src/schema.json` + `src/validator.py` で構成 JSON を事前バリデート

## ディレクトリ構成

```
ppt_skills/
├── skill.md                # Claude Code skill 定義（プロンプトと使い方）
├── src/
│   ├── generator.py        # エントリーポイント: generate_pptx(config, output_path)
│   ├── schema.json         # 構成 JSON のスキーマ
│   ├── validator.py        # スキーマ検証
│   ├── layouts/            # cover / agenda / content / chart_page / comparison / section_divider / closing
│   ├── components/         # bullets, callout, table, kpi_cards, pyramid, process_flow, ...
│   └── themes/             # monotone / dark / colorful
├── tests/
│   ├── test_e2e.py         # 全機能を網羅した 26 枚サンプルデッキの生成テスト
│   ├── test_new_features.py
│   └── test_validator.py
└── requirements.txt
```

## セットアップ

```bash
git clone https://github.com/ken8323/ppt_skills.git
cd ppt_skills
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

依存: `python-pptx`, `pytest`, `jsonschema`

## 使い方

### 1. Python から直接呼ぶ

```python
import sys
sys.path.insert(0, "/path/to/ppt_skills")
from src.generator import generate_pptx

config = {
    "theme": "monotone",
    "footer": "株式会社ABC | 社外秘",
    "brand_name": "ABC Consulting",
    "slides": [
        {
            "layout": "cover",
            "data": {
                "title": "DX推進戦略提案書",
                "subtitle": "2026年度計画",
                "client": "株式会社ABC",
                "date": "2026年4月",
            },
        },
        {
            "layout": "content",
            "data": {
                "title": "売上は3年で2.2倍に成長",
                "subtitle": "北米市場の拡大が牽引、CAGR 48% を達成",
                "columns": 1,
                "components": [
                    {
                        "type": "kpi_cards",
                        "cards": [
                            {"value": "220", "unit": "億円", "label": "2025年売上",
                             "delta": "+47%", "delta_direction": "up"},
                            {"value": "48", "unit": "%", "label": "3年 CAGR",
                             "delta": "+12pt", "delta_direction": "up"},
                        ],
                    },
                    {"type": "callout", "text": "CAGR 48% の高成長を実現"},
                ],
            },
        },
    ],
}

generate_pptx(config, "/tmp/出力.pptx")
```

### 2. Claude Code の skill として使う

`skill.md` を Claude Code が読み込むスキルとして配置し、

> 「DX 推進の提案書を dark テーマで作って」

のように依頼すると、Claude が

1. テーマ・目的のヒアリング
2. ストーリーラインに沿った構成提案（アクションタイトル + 図解中心）
3. 承認後に `generate_pptx` を実行

という手順で `.pptx` を生成します。詳細なプロンプト設計は `skill.md` を参照してください。

## レイアウトとコンポーネント

| カテゴリ | 一覧 |
|---|---|
| レイアウト | `cover`, `section_divider`, `agenda`, `content`, `chart_page`, `comparison`, `closing` |
| 図解コンポーネント | `process_flow`, `cycle`, `pyramid`, `matrix_2x2`, `org_chart`, `timeline`, `gantt`, `icon_row`, `kpi_cards`, `pillars`, `swot`, `heatmap`, `benchmark_bar` |
| データ系 | `bullets`, `callout`, `table` |
| チャート | `bar`, `line`, `pie`, `waterfall` |

各コンポーネントの必須/任意パラメータは `skill.md` の「構成 JSON リファレンス」に一覧があります。

## 設計原則

- **1 スライド 1 メッセージ**
- **アクションタイトル**: 「売上推移」ではなく「売上は3年で2.2倍に成長」のように結論を述べる
- **図解ファースト**: 3 項目以上の概念・関係・時系列・手順は箇条書きより図解を優先
- **数値には出典を**: `chart_page.source` に出典を必ず記載、推計値は「※サンプル値」と明示
- **絵文字は使わない**

## テスト

```bash
pytest tests/
```

`tests/test_e2e.py` の `FULL_DECK_CONFIG` は全機能を含む 26 枚のサンプルデッキで、出力結果が壊れていないかをエンドツーエンドで検証します。

## ライセンス

未設定（必要に応じて追加してください）。
