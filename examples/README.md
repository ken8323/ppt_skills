# Examples Gallery

`generate_pptx(config, output_path)` にそのまま渡せる完成形 JSON を集めたギャラリー。新しい資料を作るときの「似た形」の出発点として使える。

| ファイル | 用途 | テーマ | 枚数 | 主要な見どころ |
|---|---|---|---|---|
| [`consulting_proposal.json`](consulting_proposal.json) | コンサル提案書（状況→課題→解決策→効果→計画） | monotone | 13 | matrix_2x2 / pillars / kpi_cards / gantt / chart_page(bar+annotation) |
| [`monthly_report.json`](monthly_report.json) | 部門月次報告（KPI主体・comparison で光と影） | dark | 7 | kpi_cards / line chart / table(highlight) / comparison / process_flow |
| [`agentic_ai_briefing.json`](agentic_ai_briefing.json) | 全社員向け業界動向ブリーフィング | colorful | 18 | section_divider 多用 / pillars / icon_row / 出典の差し込み方 |

## 使い方

```python
import json
from pathlib import Path
from src.generator import generate_pptx

config = json.loads(Path("examples/consulting_proposal.json").read_text(encoding="utf-8"))
generate_pptx(config, "/tmp/proposal.pptx")
```

## 設計ポイント

各例で以下の skill.md の原則を実装している:

- **アクションタイトル**: 「売上推移」ではなく「売上は3年で2.2倍に成長」
- **サブヘッド**: タイトル直下に根拠を1行で添える (`subtitle`)
- **図解ファースト**: 3項目以上の概念は pillars / matrix_2x2 / process_flow で表現
- **数値スライドには `source`**: chart_page には必ず出典を入れる（`monthly_report.json`, `consulting_proposal.json` 参照）
- **出典の差し込み**: コンポーネント並びの最後に `bullets` で 1 行小さく入れる（`agentic_ai_briefing.json` の事例表/プレイヤー表参照）
- **KPI は delta + delta_direction セット**: 必ずペアで指定（lint 警告対象）

## カスタマイズの目安

- テーマ変更: トップレベル `theme` を `monotone` / `dark` / `colorful` で差し替えるだけ
- フッター/ブランド: `footer`, `brand_name` を編集
- 章を増やす: `section_divider` + `content` のペアを追加
- 図解の置き換え: `components` 内の `type` を別コンポーネントに差し替え（必須キーは skill.md 参照）

## 検証

```bash
pytest tests/test_validator.py tests/test_linter.py
```

すべての example はスキーマ検証と linter を pass している。
