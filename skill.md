---
name: ppt-generator
description: Use when the user asks to create a PowerPoint / pptx / スライド / プレゼン / 提案書 / 報告書 / プロジェクト計画書. Generates consulting-firm-quality decks via python-pptx. Japanese only, no emoji.
---

# PowerPoint資料生成スキル

コンサルティングファームレベルのPowerPoint資料を生成する。python-pptxで直接.pptxを生成するため、テンプレートや外部ツールは不要。

## いつ起動するか

ユーザーが以下のいずれかを依頼したとき:
- PowerPoint / pptx / スライド / プレゼン資料の作成
- 提案書 / 報告書 / プロジェクト計画書 / 分析資料の作成

## 処理フロー

### フェーズ1: ヒアリング

**必須質問**: 以下2点は必ずユーザーに確認する (ユーザーの最初の依頼に明記されていない場合):

1. **テーマ** (必ず選択させる。デフォルトで勝手に決めない):
   - `monotone`: 白背景 + 濃紺 + 赤アクセント（端正・フォーマル）
   - `dark`: 濃紺背景 + 白 + 橙アクセント（重厚・インパクト）
   - `colorful`: 白背景 + 青/緑/橙（モダン・カジュアル）
2. **資料の目的と対象読者** (例: 経営層向けDX提案、現場向け報告) — 依頼文から自明ならスキップ可

**任意の追加質問** (不足時のみ):
- 盛り込みたいデータ・論点・キーメッセージ

ユーザーが最初の依頼時点でテーマを明示している場合 (例: 「darkで作って」) は質問をスキップしてフェーズ2へ。

### フェーズ2: 構成提案

スライド構成をMarkdownで提示し、ユーザーの承認を得る。

**構成設計の原則:**

- **ストーリーライン重視**: 「状況 → 課題 → 原因 → 解決策 → 効果 → 実行計画」等の論理構造
- **1スライド1メッセージ**
- **アクションタイトル** (結論を述べる形にする):
  - ○「売上は3年で2.2倍に成長」「レガシー刷新が最優先課題」
  - ×「売上推移」「課題一覧」
- **図解ファースト**: 3項目以上の概念・関係性・時系列・手順は**必ず図解コンポーネントを優先**する。箇条書き (`bullets`) は図解で表現できない補足説明や例示に限定する。全スライドの**過半数に図解・チャート・表のいずれかを含める**ことを目安にする
- **箇条書きは3-5項目、各1-2行以内** (使う場合のみ)
- **絵文字は一切使わない**
- データが不明な場合はサンプル値を入れ「※サンプル値」と注記

**図解選定ガイド** (内容パターン → 使うコンポーネント):

| 内容パターン | 推奨コンポーネント |
|---|---|
| 段階的な手順・プロセス (3-6段) | `process_flow` |
| 循環する活動 (PDCA等) | `cycle` |
| 階層・重要度の順序付け | `pyramid` |
| 2軸による分類・ポジショニング | `matrix_2x2` |
| 組織・親子関係 | `org_chart` |
| 時系列イベント・マイルストーン | `timeline` |
| プロジェクトスケジュール | `gantt` |
| 並列する3-5個の概念・原則 | `icon_row` |
| 重要指標・成果数値 | `kpi_cards` |
| Before/After・対比 | `comparison` レイアウト |
| 数値の推移・比較 | `chart_page` (bar/line/pie/waterfall) |
| 構造化されたデータ | `table` |
| 強調したい一文 | `callout` |

**スライド枚数の目安**: 10-25枚。短い報告なら10枚、戦略提案なら20-25枚。Claudeが内容から判断する。

ユーザーに「この構成で生成してよいか」を確認し、修正があれば反映する。

### フェーズ3: 生成

承認後、Pythonコードを実行して.pptxを生成する:

```python
import sys
sys.path.insert(0, "/Users/kenichi/Desktop/project/ppt_skills")
from src.generator import generate_pptx

config = { ... }  # 構成JSON (下記リファレンス参照)

output_path = "/Users/kenichi/Desktop/生成資料.pptx"  # ユーザーの希望パス
generate_pptx(config, output_path)
print(f"生成完了: {output_path}")
```

生成完了後、ファイルパスを報告する。

## 構成JSONリファレンス

### トップレベル

```json
{
  "theme": "monotone" | "dark" | "colorful",
  "slides": [ { "layout": "...", "data": { ... } }, ... ]
}
```

各スライドは `{"layout": "<レイアウト名>", "data": {<レイアウト固有データ>}}` の形。

### レイアウト一覧

| layout | 用途 | data キー |
|---|---|---|
| `cover` | 表紙 | `title`, `subtitle`, `client`, `date` |
| `section_divider` | 章区切り | `section_number`, `section_title` |
| `agenda` | アジェンダ | `items: list[str]`, `highlight: int?` |
| `content` | 汎用コンテンツ | `title`, `columns: 1|2|3`, `components: list[Component]` |
| `chart_page` | チャート主体 | `title`, `chart: Chart`, `key_points: list[str]?` |
| `comparison` | 左右比較 | `title`, `left_title`, `left_components`, `right_title`, `right_components` |
| `closing` | まとめ/お礼 | `summary: list[str]`, `next_steps: list[str]` または `type: "thank_you"`, `contact` |

### コンポーネント (contentレイアウト `components` 内)

| type | 用途 | 必須パラメータ |
|---|---|---|
| `bullets` | 箇条書き | `items: list[str]` |
| `callout` | 強調ボックス | `text: str` |
| `table` | 表 | `headers: list[str]`, `rows: list[list[str]]` |
| `matrix_2x2` | 2x2マトリクス | `x_axis`, `y_axis`, `quadrants: list[4]` |
| `pyramid` | ピラミッド | `levels: list[str]` (上から順) |
| `process_flow` | プロセスフロー | `steps: list[str]` |
| `cycle` | サイクル図 | `items: list[str]` |
| `org_chart` | 組織図 | `data: {name, children: [{name, children}]}` |
| `timeline` | タイムライン | `milestones: [{date, label}]` |
| `gantt` | ガント | `tasks: [{name, start, duration}]`, `phases: list[str]` |
| `icon_row` | アイコン行 | `items: [{icon, label}]` |
| `kpi_cards` | KPIカード | `cards: [{value, unit, label}]` |

### チャート (chart_page `chart` 内)

| type | dataフォーマット | 追加オプション |
|---|---|---|
| `bar` | `{labels, series: [{name, values}]}` | `unit` |
| `line` | `{labels, series: [{name, values}]}` | `unit` |
| `pie` | `{labels, values}` | — |
| `waterfall` | `{labels, values}` (符号で増減) | — |

## 最小構成例

```json
{
  "theme": "monotone",
  "slides": [
    {
      "layout": "cover",
      "data": {
        "title": "DX推進戦略提案書",
        "subtitle": "2026年度計画",
        "client": "株式会社ABC",
        "date": "2026年4月"
      }
    },
    {
      "layout": "agenda",
      "data": {"items": ["現状分析", "課題整理", "戦略提案", "実行計画"]}
    },
    {
      "layout": "section_divider",
      "data": {"section_number": 1, "section_title": "現状分析"}
    },
    {
      "layout": "content",
      "data": {
        "title": "売上は3年で2.2倍に成長",
        "columns": 1,
        "components": [
          {"type": "bullets", "items": ["2023年: 100億円", "2024年: 150億円", "2025年: 220億円"]},
          {"type": "callout", "text": "CAGR 48%の高成長を実現"}
        ]
      }
    },
    {
      "layout": "chart_page",
      "data": {
        "title": "市場規模は年平均15%で拡大",
        "chart": {
          "type": "bar",
          "unit": "億円",
          "data": {
            "labels": ["2023", "2024", "2025"],
            "series": [{"name": "市場規模", "values": [100, 150, 220]}]
          }
        },
        "key_points": ["CAGR 15%で成長", "2025年に200億円突破"]
      }
    },
    {
      "layout": "closing",
      "data": {
        "summary": ["要点1", "要点2", "要点3"],
        "next_steps": ["ステップ1", "ステップ2", "ステップ3"]
      }
    }
  ]
}
```

## よくある落とし穴

- **pieとwaterfall** の `data` は `{labels, values}` フラット形式。`series` を使わない
- **bar/line** の `data` は `{labels, series: [{name, values}]}` のネスト形式
- **org_chart** の `data` は再帰的な `{name, children: [...]}` 構造
- **gantt** の `start` と `duration` は月数などの整数。`phases` の長さで横軸を決める
- **matrix_2x2** の `quadrants` は左上/右上/左下/右下の順で4要素
- **pyramid** の `levels` は上 (狭い) から下 (広い) の順
- 実行ファイルが存在しない環境では `pip install -r requirements.txt` を事前実行

## 実装参照

- エントリーポイント: `src/generator.py` の `generate_pptx(config, output_path)`
- レイアウト: `src/layouts/`
- コンポーネント: `src/components/`
- テーマ: `src/themes/`
- 動作例 (全機能を網羅した21枚のサンプル): `tests/test_e2e.py` の `FULL_DECK_CONFIG`
