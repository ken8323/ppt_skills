# PowerPoint資料生成スキル 設計書

## 概要

Claude Codeのスキルとして、コンサルティングファームレベルのPowerPoint資料をpython-pptxで生成する。ユーザーの自然言語入力から、構成提案→承認→.pptx生成までを一貫して行う。

## 要件

- 汎用型: 提案資料、プレゼン、プロジェクト計画書、分析レポート等に対応
- 入力: ハイブリッド（自然言語→構成案提示→承認→生成）
- デザイン: テーマ切り替え可能（monotone / dark / colorful）
- ビジュアル: フルセット（テキスト、図形、チャート、表、タイムライン、アイコン）
- 言語: 日本語
- 出力: .pptxファイル
- 枚数: Claude判断
- 絵文字不使用

## アーキテクチャ

コンポーネントライブラリ方式を採用。スライド要素を個別のPython関数としてライブラリ化し、スキルプロンプトがClaude にこれらを組み合わせたコードを生成・実行させる。

### ディレクトリ構成

```
ppt_skills/
├── skill.md                    # スキル定義（Claudeへのプロンプト）
├── src/
│   ├── generator.py            # エントリポイント: 構成JSONを受け取り.pptxを生成
│   ├── themes/
│   │   ├── base.py             # テーマ基底クラス
│   │   ├── monotone.py         # 白背景+濃紺+赤アクセント
│   │   ├── dark.py             # 濃紺背景+白テキスト+橙アクセント
│   │   └── colorful.py         # 白背景+青/緑/橙の3色
│   ├── components/
│   │   ├── text.py             # タイトル、サブタイトル、箇条書き、コールアウト、脚注
│   │   ├── chart.py            # 棒グラフ、折れ線、円グラフ、ウォーターフォール
│   │   ├── table.py            # データ表
│   │   ├── shape.py            # 2x2マトリクス、ピラミッド、プロセスフロー、サイクル、組織図
│   │   ├── timeline.py         # タイムライン、ガントチャート
│   │   └── icon.py             # アイコン+ラベル、KPIカード、アイコン横並び
│   └── layouts/
│       ├── cover.py            # 表紙
│       ├── agenda.py           # アジェンダ/目次
│       ├── section_divider.py  # セクション区切り
│       ├── content.py          # 汎用コンテンツ（1/2/3カラム）
│       ├── chart_page.py       # チャート主体ページ
│       ├── comparison.py       # 比較ページ
│       └── closing.py          # まとめ/Next Steps
└── requirements.txt
```

### 処理フロー

1. スキル起動 → Claudeがユーザーの要件をヒアリング（1-2回の質問）
2. Claudeがスライド構成をMarkdown形式で提示
3. ユーザーが承認（修正があれば反映）
4. Claudeが構成JSONを生成し、generator.pyを呼ぶPythonコードを生成・実行
5. generator.pyがJSONを受け取り、テーマ+レイアウト+コンポーネントで.pptxを出力

## テーマシステム

テーマはデザイントークンの集合体。全コンポーネントがテーマオブジェクトを参照する。

### Theme基底クラス

```python
class Theme:
    # カラー
    primary: str        # メインカラー（タイトル、強調要素）
    secondary: str      # サブカラー（アクセント）
    background: str     # スライド背景
    text_primary: str   # 本文テキスト
    text_secondary: str # 補足テキスト
    border: str         # 線、枠
    chart_colors: list  # チャート用カラーパレット（5-6色）

    # フォント
    font_title: str     # タイトル用（Yu Gothic Bold）
    font_body: str      # 本文用（Yu Gothic）
    font_size_title: Pt
    font_size_subtitle: Pt
    font_size_body: Pt
    font_size_caption: Pt

    # レイアウト定数
    margin_top: Inches
    margin_bottom: Inches
    margin_left: Inches
    margin_right: Inches
    content_area_top: Inches
    line_spacing: float
```

### テーマ定義

| トークン | Monotone | Dark | Colorful |
|---|---|---|---|
| primary | #1B2A4A | #FFFFFF | #2D5BFF |
| secondary | #C8102E | #FF6B35 | #00C49A |
| background | #FFFFFF | #1B2A4A | #FFFFFF |
| text_primary | #1B2A4A | #FFFFFF | #2C3E50 |
| text_secondary | #6B7B8D | #A0B0C0 | #7F8C8D |
| border | #D0D5DD | #3A4F6F | #E0E0E0 |
| font_title | Yu Gothic Bold | Yu Gothic Bold | Yu Gothic Bold |
| font_body | Yu Gothic | Yu Gothic | Yu Gothic |

フォントフォールバック順: Yu Gothic → Meiryo → MS Gothic → sans-serif

## コンポーネント詳細

### text.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_title()` | スライドタイトル | 左上配置、太字、下に区切り線 |
| `add_subtitle()` | サブタイトル | タイトル直下、やや小さく、text_secondary色 |
| `add_bullets()` | 箇条書き | インデント2階層対応、行頭は「―」等の控えめ記号 |
| `add_callout()` | 強調ボックス | 背景色付き矩形内にテキスト |
| `add_footnote()` | 脚注 | スライド下部、小フォント、text_secondary色 |

### chart.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_bar_chart()` | 棒グラフ | 縦/横対応、グリッド線最小限、データラベル付き |
| `add_line_chart()` | 折れ線 | マーカー付き、凡例は右側またはチャート内直接ラベル |
| `add_pie_chart()` | 円グラフ | ドーナツ型対応、ラベル+割合表示 |
| `add_waterfall()` | ウォーターフォール | 増減色分け（増=primary, 減=secondary） |

### table.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_table()` | データ表 | ヘッダ行primary色背景+白文字、ストライプ行、適切なセル内余白 |

### shape.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_matrix_2x2()` | 2x2マトリクス | 軸ラベル付き、各象限にテキスト |
| `add_pyramid()` | ピラミッド | 3-5段、台形の積み重ね |
| `add_process_flow()` | プロセスフロー | 矢印で繋がった角丸矩形、横並び |
| `add_cycle()` | サイクル図 | 円形配置の矢印ループ |
| `add_org_chart()` | 組織図/体制図 | ツリー構造、線で接続 |

### timeline.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_timeline()` | タイムライン | 横軸に時間、マイルストーンを上下にプロット |
| `add_gantt()` | ガントチャート | 横棒で期間表示、フェーズ色分け |

### icon.py

| 関数 | 用途 | 仕様 |
|---|---|---|
| `add_icon_with_label()` | アイコン+ラベル | 丸/角丸内に幾何学記号+下にテキスト |
| `add_kpi_card()` | KPI表示 | 大きな数字+単位+ラベル、カード型 |
| `add_icon_row()` | アイコン横並び | 3-5個のアイコン+ラベルを等間隔配置 |

### コンポーネント共通ルール

- テーマオブジェクトを第1引数で受け取る
- 配置座標はレイアウト側が計算して渡す
- 渡されたslideオブジェクトに直接描画（戻り値なし）

## レイアウト詳細

### cover.py — 表紙
- 中央にタイトル（大きめ）、その下にサブタイトル
- 左下にクライアント名、右下に日付
- テーマに応じて背景色全面塗りまたは下部にアクセントバー

### agenda.py — アジェンダ/目次
- タイトル「Agenda」+ 番号付きリスト
- 各項目: 番号（primary色、大きめ）+ テキスト横並び
- 現在セクションのハイライト表示対応

### section_divider.py — セクション区切り
- 中央にセクション番号+セクション名を大きく表示
- 背景はprimary色ベタ塗り、テキストは白

### content.py — 汎用コンテンツ
- 1/2/3カラムをcolumnsパラメータで切り替え
- 各カラム内にテキスト系コンポーネントを自由に配置
- コンポーネントの種類と数に応じてY座標を自動計算

### chart_page.py — チャート主体
- 左にチャート（幅65%）+ 右にキーポイント箇条書き（幅30%）
- またはチャート全面表示（テキストなしモード）

### comparison.py — 比較
- 左右2分割、中央に区切り線
- Before/After、As-Is/To-Be、Option A/B対応
- 各側にタイトル+任意のコンポーネント

### closing.py — まとめ/Next Steps
- 上部に「まとめ」箇条書き
- 下部にNext Stepsを番号付きで配置
- または「Thank You」+連絡先の最終ページ

### レイアウト共通

- 全レイアウトでタイトル位置を統一（左上、margin_left / margin_top）
- コンテンツ領域: content_area_topからmargin_bottomまでの矩形
- Layoutクラスがrender(slide, theme, content_data)メソッドを持つ

## generator.py 処理フロー

1. テーマ名からThemeインスタンスを生成
2. Presentationオブジェクトを作成（16:9）
3. slides配列をループ:
   a. 空白スライドを追加
   b. layoutに対応するLayoutクラスを取得（LAYOUT_MAP）
   c. Layout.render(slide, theme, content_data)を呼び出し
4. .pptxとして保存

### エラーハンドリング
- 不明なlayout名 → エラーメッセージ出力、該当スライドをスキップ
- 不明なchart type → 同上
- フォント不在 → フォールバック順で代替

## スキルプロンプト（skill.md）

### フェーズ1: ヒアリング
- 資料の目的（提案、報告、計画、分析等）
- 対象読者（経営層、現場、クライアント等）
- 盛り込みたい内容・キーメッセージ
- テーマ選択
- 1-2回の質問で把握。情報が十分なら質問をスキップ

### フェーズ2: 構成提案
- Markdown形式でスライド構成を提示
- コンサルのストーリーライン構築力を発揮
- 「状況→課題→原因→解決策→効果→実行計画」等の論理構造を自動設計
- ユーザー承認を求める

### フェーズ3: 生成・実行
- 構成をJSONに変換
- generator.pyを呼ぶPythonコードを生成・実行
- .pptxを出力

### プロンプト内ルール
- 絵文字を一切使わない
- 1スライド1メッセージ原則
- アクションタイトル（結論を述べる形。例:「売上は3年で2.2倍に成長」）
- 箇条書きは3-5項目、各項目1-2行以内
- データがない場合はプレースホルダー値+その旨を明記

## 構成JSONフォーマット

```json
{
  "theme": "monotone",
  "title": "DX推進戦略提案書",
  "output_path": "output.pptx",
  "slides": [
    {
      "layout": "cover",
      "title": "DX推進戦略提案書",
      "subtitle": "2026年度計画",
      "client": "株式会社ABC",
      "date": "2026年4月"
    },
    {
      "layout": "agenda",
      "items": ["現状分析", "課題整理", "戦略提案", "実行計画"]
    },
    {
      "layout": "section_divider",
      "section_number": 1,
      "section_title": "現状分析"
    },
    {
      "layout": "content",
      "columns": 1,
      "title": "売上は3年で2.2倍に成長",
      "components": [
        {
          "type": "bullets",
          "items": ["要点1", "要点2", "要点3"]
        }
      ]
    },
    {
      "layout": "chart_page",
      "title": "市場規模は年平均15%で拡大",
      "chart": {
        "type": "bar",
        "data": {
          "labels": ["2023", "2024", "2025"],
          "series": [{"name": "市場規模", "values": [100, 150, 220]}]
        },
        "unit": "億円"
      },
      "key_points": ["CAGR 15%で成長", "2025年に200億円突破"]
    },
    {
      "layout": "content",
      "columns": 2,
      "title": "DX成熟度の現状評価",
      "components": [
        {
          "type": "matrix_2x2",
          "x_axis": "実行難易度",
          "y_axis": "事業インパクト",
          "quadrants": ["Quick Win", "戦略投資", "要検討", "後回し"]
        }
      ]
    },
    {
      "layout": "closing",
      "summary": ["要点1", "要点2", "要点3"],
      "next_steps": ["ステップ1", "ステップ2", "ステップ3"]
    }
  ]
}
```
