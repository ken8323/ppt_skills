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
- **サブヘッド推奨**: アクションタイトルの 1 行下に、タイトルの根拠や数値補足を `subtitle` で添える (content/chart_page/comparison で対応)。例「CAGR 48% を牽引したのは北米市場」
- **図解ファースト**: 3項目以上の概念・関係性・時系列・手順は**必ず図解コンポーネントを優先**する。箇条書き (`bullets`) は図解で表現できない補足説明や例示に限定する。全スライドの**過半数に図解・チャート・表のいずれかを含める**ことを目安にする
- **数値スライドには出典を明記**: `chart_page` では `source` フィールドに出典 (社内データ/調査機関名/年度) を必ず入れる。推計・サンプル値は「(※サンプル値)」と明示
- **KPI は前年差分を添える**: `kpi_cards` では可能な限り `delta` + `delta_direction` (up/down/flat) で差分バッジを付ける。読み手の「良いのか悪いのか」判断を助ける
- **チャート注記で結論を刺す**: `bar` / `line` の `annotations` に `{"category": "2025", "text": "最高値 +48%"}` 等を入れて、着目ポイントに注釈ピルを重ねる
- **箇条書きは3-5項目、各1-2行以内** (使う場合のみ)
- **絵文字は一切使わない**
- **ページ番号とフッターは自動注入** (cover / section_divider / thank_you を除く)。トップレベル `footer` / `brand_name` で一括設定
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
| 数値の推移・比較 | `chart_page` (bar/line/area) |
| 構成比の時系列推移 | `chart_page` (stacked_bar/area stacked) |
| 売上 + 成長率など 単位の異なる2指標 | `chart_page` (combo + secondary_axis) |
| 2軸でのプロット (例: コスト×効果) | `chart_page` (scatter) |
| 内訳・構成比の単一断面 | `chart_page` (pie) |
| 増減の積み上げ可視化 | `chart_page` (waterfall) |
| 構造化されたデータ | `table` |
| 強調したい一文 | `callout` |
| 戦略テーマ・3本柱 | `pillars` |
| SWOT/3C/4P等フレームワーク | `swot` |
| 優先度×影響度マッピング | `heatmap` |
| 自社vs競合ベンチマーク | `benchmark_bar` |

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
  "footer": "株式会社ABC | 社外秘",       // 任意。各ページ左下に表示
  "brand_name": "ABC Consulting",          // 任意。footer 未指定時のフォールバック
  "slides": [ { "layout": "...", "data": { ... } }, ... ]
}
```

各スライドは `{"layout": "<レイアウト名>", "data": {<レイアウト固有データ>}}` の形。ページ番号は各スライド右下に自動挿入される (cover/section_divider/thank_you を除く)。

### レイアウト一覧

| layout | 用途 | data キー |
|---|---|---|
| `cover` | 表紙 (左 primary パネル + 右タイトル) | `title`, `subtitle`, `client`, `date` ※トップレベル `brand_name` はパネル上部に白文字で自動表示 |
| `section_divider` | 章区切り | `section_number`, `section_title` |
| `agenda` | アジェンダ | `items: list[str]`, `highlight: int?` |
| `content` | 汎用コンテンツ | `title`, `subtitle?`, `columns: 1|2|3`, `components: list[Component]` |
| `chart_page` | チャート主体 | `title`, `subtitle?`, `chart: Chart`, `key_points: list[str]?`, `source: str?` |
| `comparison` | 左右比較 | `title`, `subtitle?`, `left_title`, `left_components`, `right_title`, `right_components` |
| `closing` | まとめ/お礼 | `summary: list[str]`, `next_steps: list[str]` または `type: "thank_you"`, `contact` |

**共通の任意フィールド** (cover/section_divider 以外の全レイアウトで使用可):
- `subtitle`: アクションタイトル直下に配置される 1 行サブヘッド。content/chart_page/comparison で対応
- `source: str`: 「出典: ...」としてフッター上端に italic で小さく描画。例: `"経済産業省 2025年3月"`
- `sources: list`: 複数出典を `/` 区切りで結合して描画。要素は `string` または `{label, url}`。例: `[{"label": "Gartner", "url": "https://..."}, "社内データ"]`
- 数値・事例スライドには可能な限り出典を明記。`source` と `sources` の同時指定は不可

### コンポーネント (contentレイアウト `components` 内)

| type | 用途 | 必須パラメータ | 任意パラメータ |
|---|---|---|---|
| `bullets` | 箇条書き | `items: list[str]` | — |
| `callout` | 左アクセントバー付き強調ボックス | `text: str` | `variant: "info"\|"success"\|"warning"\|"danger"`（バーと背景色に反映。デフォルトは `info` = primary色） |
| `table` | 表 | `headers: list[str]`, `rows: list[list[str]]` | `highlight_rows: list[int]`, `highlight_cells: [{row, col}]`, `align: str\|list[str]`（"left"/"right"/"center"。未指定時は数値列を自動で右寄せ）, `banded: bool`（既定 true。ストライプ無効化は false）, `totals_row: bool\|list[str]`（true で数値列を自動合計し末尾に合計行を追加。明示指定も可）, `col_widths_ratio: list[float]`（列幅の相対比、例 [3,1,1]） |
| `matrix_2x2` | 2x2マトリクス | `x_axis`, `y_axis`, `quadrants: list[4]` | `recommended_quadrant: 0-3`（太枠強調。0=左上, 1=右上, 2=左下, 3=右下） |
| `pyramid` | ピラミッド | `levels: list[str\|{text,note}]` (上から順) | 各 level を `{"text": "...", "note": "右注釈"}` 形式にすると右側に注釈を表示 |
| `process_flow` | プロセスフロー | `steps: list[str]` | `style: "arrow"(デフォ)\|"chevron"` |
| `cycle` | サイクル図 | `items: list[str]` | — |
| `org_chart` | 組織図 | `data: {name, children: [{name, children}]}` | — |
| `timeline` | タイムライン | `milestones: [{date, label}]` | `today: str`（日付文字列。マイルストーンと同じ形式で指定すると「現在」マーカーを表示） |
| `gantt` | ガント | `tasks: [{name, start, duration}]`, `phases: list[str]` | 各 task に `progress: 0.0-1.0`（完了比率バーを濃色で表示） |
| `icon_row` | アイコン行 | `items: [{icon, label}]` | — |
| `kpi_cards` | KPIカード | `cards: [{value, unit, label}]` | 各 card に `delta: str`, `delta_direction: "up"\|"down"\|"flat"` |
| `pillars` | 縦柱（3-5本） | `items: [{title, body}]` | 各 item に `kpi: str`（大きな数値を柱下部に表示） |
| `swot` | SWOTフレームワーク | `cells: [{title, items: list[str]}]`（左上/右上/左下/右下の順） | — |
| `heatmap` | ヒートマップ | `col_headers: list[str]`, `row_headers: list[str]`, `values: list[list[float]]` | — |
| `benchmark_bar` | 横棒ベンチマーク | `items: [{label, value}]` | `unit: str?`、各 item に `is_self: true` で自社を primary 色で強調 |

### チャート (chart_page `chart` 内)

| type | dataフォーマット | 追加オプション |
|---|---|---|
| `bar` | `{labels, series: [{name, values}]}` | `unit: str?`, `annotations: list?` |
| `line` | `{labels, series: [{name, values}]}` | `unit: str?`, `annotations: list?` |
| `stacked_bar` | `{labels, series: [{name, values}]}` (複数系列を積み上げ) | `unit: str?`, `horizontal: bool?`, `annotations: list?` |
| `area` | `{labels, series: [{name, values}]}` | `unit: str?`, `stacked: bool?`, `annotations: list?` |
| `scatter` | `{series: [{name, points: [[x, y], ...]}]}` (XY 散布) | `x_label: str?`, `y_label: str?` ※annotations 非対応 |
| `combo` | `{labels, bars: [{name, values, unit?}], lines: [{name, values, unit?, secondary_axis?}]}` | `annotations: list?` ※`secondary_axis: true` で第2軸に切替 (右側) |
| `pie` | `{labels, values}` | — |
| `waterfall` | `{labels, values}` (符号で増減) | — ※正値は緑(success)、負値は赤(danger)、先頭・末尾のバーは primary で自動着色 |

**annotations フォーマット** (bar/line のみ):

```json
[
  {"category": "2025", "text": "最高値 +48%"},
  {"category": 3, "text": "ROI 達成", "position": "bottom"}
]
```

- `category`: 該当ラベル文字列または 0 始まりのインデックス
- `text`: 注記ピルに表示する文字列 (短く)
- `position`: `"top"` (デフォ) または `"bottom"`

## 最小構成例

```json
{
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
        "subtitle": "北米市場の拡大が牽引、CAGR 48% を達成",
        "columns": 1,
        "components": [
          {
            "type": "kpi_cards",
            "cards": [
              {"value": "220", "unit": "億円", "label": "2025年売上",
               "delta": "+47%", "delta_direction": "up"},
              {"value": "48", "unit": "%", "label": "3年 CAGR",
               "delta": "+12pt", "delta_direction": "up"}
            ]
          },
          {"type": "callout", "text": "CAGR 48%の高成長を実現"}
        ]
      }
    },
    {
      "layout": "chart_page",
      "data": {
        "title": "市場規模は年平均15%で拡大",
        "subtitle": "2025年に 200 億円を突破、競合参入も加速",
        "chart": {
          "type": "bar",
          "unit": "億円",
          "data": {
            "labels": ["2023", "2024", "2025"],
            "series": [{"name": "市場規模", "values": [100, 150, 220]}]
          },
          "annotations": [
            {"category": "2025", "text": "200億円突破"}
          ]
        },
        "key_points": ["CAGR 15%で成長", "2025年に200億円突破"],
        "source": "経済産業省調査 (2025年3月)"
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
- **annotations** は `bar` / `line` でのみ有効。`pie` / `waterfall` には指定しない
- **delta は delta_direction とセット**: 片方のみだとバッジが出ない
- **org_chart** の `data` は再帰的な `{name, children: [...]}` 構造
- **gantt** の `start` と `duration` は月数などの整数。`phases` の長さで横軸を決める
- **matrix_2x2** の `quadrants` は左上/右上/左下/右下の順で4要素
- **pyramid** の `levels` は上 (狭い) から下 (広い) の順
- **waterfall** のバー配色は自動（正値=緑、負値=赤、先頭末尾=primary）。JSONでの指定不要
- 実行ファイルが存在しない環境では `pip install -r requirements.txt` を事前実行

## 実装参照

- エントリーポイント: `src/generator.py` の `generate_pptx(config, output_path)`
  - 既定で schema 検証 + linter (オーバーフロー警告) を実行。`validate=False` / `lint=False` で個別に無効化可能
- レイアウト: `src/layouts/`
- コンポーネント: `src/components/`
- テーマ: `src/themes/`
- スキーマ定義: `src/schema.json`、検証ロジック: `src/validator.py`、linter: `src/linter.py`
- **完成形サンプル**: `examples/` (consulting_proposal / monthly_report / agentic_ai_briefing)。新規依頼時は近い形のものを起点にする
- 動作例 (全機能を網羅した26枚のサンプル): `tests/test_e2e.py` の `FULL_DECK_CONFIG`
