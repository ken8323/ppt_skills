# PowerPoint資料生成スキル 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** コンサルティングファームレベルのPowerPoint資料をpython-pptxで生成するClaude Codeスキルを構築する

**Architecture:** コンポーネントライブラリ方式。テーマシステムがデザイントークンを管理し、コンポーネント（text/chart/table/shape/timeline/icon）がスライド要素を描画し、レイアウト（cover/agenda/content等）がコンポーネントの配置を制御する。generator.pyがJSON構成データを受け取り、テーマ+レイアウト+コンポーネントを組み合わせて.pptxを出力する。

**Tech Stack:** Python 3.10+, python-pptx, pytest

**Spec:** `docs/superpowers/specs/2026-04-18-ppt-skill-design.md`

---

## File Structure

```
ppt_skills/
├── skill.md
├── requirements.txt
├── src/
│   ├── __init__.py
│   ├── generator.py
│   ├── themes/
│   │   ├── __init__.py
│   │   ├── base.py
│   │   ├── monotone.py
│   │   ├── dark.py
│   │   └── colorful.py
│   ├── components/
│   │   ├── __init__.py
│   │   ├── text.py
│   │   ├── chart.py
│   │   ├── table.py
│   │   ├── shape.py
│   │   ├── timeline.py
│   │   └── icon.py
│   └── layouts/
│       ├── __init__.py
│       ├── cover.py
│       ├── agenda.py
│       ├── section_divider.py
│       ├── content.py
│       ├── chart_page.py
│       ├── comparison.py
│       └── closing.py
└── tests/
    ├── __init__.py
    ├── conftest.py
    ├── test_themes.py
    ├── test_text.py
    ├── test_chart.py
    ├── test_table.py
    ├── test_shape.py
    ├── test_timeline.py
    ├── test_icon.py
    ├── test_layouts.py
    ├── test_generator.py
    └── test_e2e.py
```

---

### Task 1: プロジェクトセットアップ

**Files:**
- Create: `requirements.txt`
- Create: `src/__init__.py`
- Create: `src/themes/__init__.py`
- Create: `src/components/__init__.py`
- Create: `src/layouts/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: requirements.txtを作成**

```
python-pptx>=0.6.23
pytest>=7.0
```

- [ ] **Step 2: パッケージ初期化ファイルを作成**

`src/__init__.py`:
```python
"""PowerPoint資料生成ライブラリ"""
```

`src/themes/__init__.py`:
```python
from src.themes.monotone import MonotoneTheme
from src.themes.dark import DarkTheme
from src.themes.colorful import ColorfulTheme

THEME_MAP = {
    "monotone": MonotoneTheme,
    "dark": DarkTheme,
    "colorful": ColorfulTheme,
}


def get_theme(name: str):
    theme_class = THEME_MAP.get(name)
    if theme_class is None:
        raise ValueError(f"Unknown theme: {name}. Available: {list(THEME_MAP.keys())}")
    return theme_class()
```

`src/components/__init__.py`:
```python
"""スライドコンポーネント"""
```

`src/layouts/__init__.py`:
```python
from src.layouts.cover import CoverLayout
from src.layouts.agenda import AgendaLayout
from src.layouts.section_divider import SectionDividerLayout
from src.layouts.content import ContentLayout
from src.layouts.chart_page import ChartPageLayout
from src.layouts.comparison import ComparisonLayout
from src.layouts.closing import ClosingLayout

LAYOUT_MAP = {
    "cover": CoverLayout,
    "agenda": AgendaLayout,
    "section_divider": SectionDividerLayout,
    "content": ContentLayout,
    "chart_page": ChartPageLayout,
    "comparison": ComparisonLayout,
    "closing": ClosingLayout,
}


def get_layout(name: str):
    layout_class = LAYOUT_MAP.get(name)
    if layout_class is None:
        raise ValueError(f"Unknown layout: {name}. Available: {list(LAYOUT_MAP.keys())}")
    return layout_class()
```

`tests/__init__.py`:
```python
```

`tests/conftest.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.themes.monotone import MonotoneTheme
from src.themes.dark import DarkTheme
from src.themes.colorful import ColorfulTheme


@pytest.fixture
def prs():
    """16:9のPresentationオブジェクト"""
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


@pytest.fixture
def blank_slide(prs):
    """空白スライド"""
    layout = prs.slide_layouts[6]  # blank layout
    return prs.slides.add_slide(layout)


@pytest.fixture
def monotone_theme():
    return MonotoneTheme()


@pytest.fixture
def dark_theme():
    return DarkTheme()


@pytest.fixture
def colorful_theme():
    return ColorfulTheme()
```

- [ ] **Step 3: 依存関係をインストール**

Run: `pip install -r requirements.txt`

- [ ] **Step 4: コミット**

```bash
git add requirements.txt src/__init__.py src/themes/__init__.py src/components/__init__.py src/layouts/__init__.py tests/__init__.py tests/conftest.py
git commit -m "feat: project setup with package structure and test fixtures"
```

---

### Task 2: テーマシステム

**Files:**
- Create: `src/themes/base.py`
- Create: `src/themes/monotone.py`
- Create: `src/themes/dark.py`
- Create: `src/themes/colorful.py`
- Create: `tests/test_themes.py`

- [ ] **Step 1: テストを書く**

`tests/test_themes.py`:
```python
import pytest
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

from src.themes.base import Theme
from src.themes.monotone import MonotoneTheme
from src.themes.dark import DarkTheme
from src.themes.colorful import ColorfulTheme
from src.themes import get_theme


class TestThemeBase:
    def test_theme_has_required_color_attrs(self, monotone_theme):
        for attr in ["primary", "secondary", "background", "text_primary",
                      "text_secondary", "border", "chart_colors"]:
            assert hasattr(monotone_theme, attr)

    def test_theme_has_required_font_attrs(self, monotone_theme):
        for attr in ["font_title", "font_body", "font_size_title",
                      "font_size_subtitle", "font_size_body", "font_size_caption"]:
            assert hasattr(monotone_theme, attr)

    def test_theme_has_required_layout_attrs(self, monotone_theme):
        for attr in ["margin_top", "margin_bottom", "margin_left", "margin_right",
                      "content_area_top", "line_spacing"]:
            assert hasattr(monotone_theme, attr)

    def test_chart_colors_has_at_least_5(self, monotone_theme):
        assert len(monotone_theme.chart_colors) >= 5

    def test_color_returns_rgbcolor(self, monotone_theme):
        assert isinstance(monotone_theme.primary, RGBColor)

    def test_font_size_returns_pt(self, monotone_theme):
        assert isinstance(monotone_theme.font_size_title, int)  # Pt returns int (EMU)


class TestThemeVariants:
    def test_monotone_white_background(self, monotone_theme):
        assert monotone_theme.background == RGBColor(0xFF, 0xFF, 0xFF)

    def test_dark_dark_background(self, dark_theme):
        assert dark_theme.background == RGBColor(0x1B, 0x2A, 0x4A)

    def test_colorful_white_background(self, colorful_theme):
        assert colorful_theme.background == RGBColor(0xFF, 0xFF, 0xFF)


class TestGetTheme:
    def test_get_monotone(self):
        theme = get_theme("monotone")
        assert isinstance(theme, MonotoneTheme)

    def test_get_dark(self):
        theme = get_theme("dark")
        assert isinstance(theme, DarkTheme)

    def test_get_colorful(self):
        theme = get_theme("colorful")
        assert isinstance(theme, ColorfulTheme)

    def test_unknown_theme_raises(self):
        with pytest.raises(ValueError, match="Unknown theme"):
            get_theme("neon")
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_themes.py -v`
Expected: FAIL (import errors)

- [ ] **Step 3: base.pyを実装**

`src/themes/base.py`:
```python
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_str: str) -> RGBColor:
    """'#1B2A4A' -> RGBColor(0x1B, 0x2A, 0x4A)"""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class Theme:
    # カラー（サブクラスでオーバーライド）
    primary: RGBColor
    secondary: RGBColor
    background: RGBColor
    text_primary: RGBColor
    text_secondary: RGBColor
    border: RGBColor
    chart_colors: list  # list[RGBColor]

    # フォント
    font_title: str = "Yu Gothic"
    font_body: str = "Yu Gothic"
    font_size_title: int = Pt(24)
    font_size_subtitle: int = Pt(16)
    font_size_body: int = Pt(14)
    font_size_caption: int = Pt(10)

    # レイアウト定数（16:9スライド: 13.333 x 7.5 inches）
    margin_top: int = Inches(0.4)
    margin_bottom: int = Inches(0.4)
    margin_left: int = Inches(0.6)
    margin_right: int = Inches(0.6)
    content_area_top: int = Inches(1.4)
    line_spacing: float = 1.2

    @property
    def slide_width(self) -> int:
        return Inches(13.333)

    @property
    def slide_height(self) -> int:
        return Inches(7.5)

    @property
    def content_width(self) -> int:
        """マージンを除いたコンテンツ領域の幅"""
        return self.slide_width - self.margin_left - self.margin_right

    @property
    def content_height(self) -> int:
        """content_area_topからmargin_bottomまでの高さ"""
        return self.slide_height - self.content_area_top - self.margin_bottom
```

- [ ] **Step 4: monotone.pyを実装**

`src/themes/monotone.py`:
```python
from src.themes.base import Theme, hex_to_rgb


class MonotoneTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#1B2A4A")
        self.secondary = hex_to_rgb("#C8102E")
        self.background = hex_to_rgb("#FFFFFF")
        self.text_primary = hex_to_rgb("#1B2A4A")
        self.text_secondary = hex_to_rgb("#6B7B8D")
        self.border = hex_to_rgb("#D0D5DD")
        self.chart_colors = [
            hex_to_rgb("#1B2A4A"),
            hex_to_rgb("#C8102E"),
            hex_to_rgb("#4A7FB5"),
            hex_to_rgb("#8B9DAF"),
            hex_to_rgb("#D4A574"),
            hex_to_rgb("#6B8E6B"),
        ]

        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
```

- [ ] **Step 5: dark.pyを実装**

`src/themes/dark.py`:
```python
from src.themes.base import Theme, hex_to_rgb


class DarkTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#FFFFFF")
        self.secondary = hex_to_rgb("#FF6B35")
        self.background = hex_to_rgb("#1B2A4A")
        self.text_primary = hex_to_rgb("#FFFFFF")
        self.text_secondary = hex_to_rgb("#A0B0C0")
        self.border = hex_to_rgb("#3A4F6F")
        self.chart_colors = [
            hex_to_rgb("#FFFFFF"),
            hex_to_rgb("#FF6B35"),
            hex_to_rgb("#5BA4E6"),
            hex_to_rgb("#A0B0C0"),
            hex_to_rgb("#FFD166"),
            hex_to_rgb("#06D6A0"),
        ]

        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
```

- [ ] **Step 6: colorful.pyを実装**

`src/themes/colorful.py`:
```python
from src.themes.base import Theme, hex_to_rgb


class ColorfulTheme(Theme):
    def __init__(self):
        self.primary = hex_to_rgb("#2D5BFF")
        self.secondary = hex_to_rgb("#00C49A")
        self.background = hex_to_rgb("#FFFFFF")
        self.text_primary = hex_to_rgb("#2C3E50")
        self.text_secondary = hex_to_rgb("#7F8C8D")
        self.border = hex_to_rgb("#E0E0E0")
        self.chart_colors = [
            hex_to_rgb("#2D5BFF"),
            hex_to_rgb("#00C49A"),
            hex_to_rgb("#FF6B35"),
            hex_to_rgb("#FFD166"),
            hex_to_rgb("#EF476F"),
            hex_to_rgb("#7B68EE"),
        ]

        self.font_title = "Yu Gothic"
        self.font_body = "Yu Gothic"
```

- [ ] **Step 7: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_themes.py -v`
Expected: ALL PASS

- [ ] **Step 8: コミット**

```bash
git add src/themes/ tests/test_themes.py
git commit -m "feat: theme system with monotone, dark, colorful variants"
```

---

### Task 3: テキストコンポーネント

**Files:**
- Create: `src/components/text.py`
- Create: `tests/test_text.py`

- [ ] **Step 1: テストを書く**

`tests/test_text.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from src.components.text import add_title, add_subtitle, add_bullets, add_callout, add_footnote
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestAddTitle:
    def test_adds_textbox(self, slide, theme):
        add_title(slide, theme, "テストタイトル", Inches(0.6), Inches(0.4))
        assert len(slide.shapes) >= 1

    def test_text_content(self, slide, theme):
        add_title(slide, theme, "テストタイトル", Inches(0.6), Inches(0.4))
        text = slide.shapes[0].text_frame.paragraphs[0].text
        assert text == "テストタイトル"

    def test_adds_separator_line(self, slide, theme):
        add_title(slide, theme, "テストタイトル", Inches(0.6), Inches(0.4))
        # タイトルテキストボックス + 区切り線 = 2つのshape
        assert len(slide.shapes) >= 2


class TestAddSubtitle:
    def test_adds_textbox(self, slide, theme):
        add_subtitle(slide, theme, "サブタイトル", Inches(0.6), Inches(1.0))
        assert len(slide.shapes) >= 1

    def test_text_content(self, slide, theme):
        add_subtitle(slide, theme, "サブタイトル", Inches(0.6), Inches(1.0))
        text = slide.shapes[0].text_frame.paragraphs[0].text
        assert text == "サブタイトル"


class TestAddBullets:
    def test_adds_textbox(self, slide, theme):
        add_bullets(slide, theme, ["項目1", "項目2", "項目3"], Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1

    def test_bullet_count(self, slide, theme):
        items = ["項目1", "項目2", "項目3"]
        add_bullets(slide, theme, items, Inches(0.6), Inches(1.5))
        paragraphs = slide.shapes[0].text_frame.paragraphs
        # 各itemが1つのparagraphになる
        texts = [p.text for p in paragraphs if p.text.strip()]
        assert len(texts) == 3

    def test_nested_bullets(self, slide, theme):
        items = ["項目1", {"text": "サブ項目", "level": 1}]
        add_bullets(slide, theme, items, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1


class TestAddCallout:
    def test_adds_shape(self, slide, theme):
        add_callout(slide, theme, "重要なメッセージ", Inches(0.6), Inches(2.0))
        assert len(slide.shapes) >= 1


class TestAddFootnote:
    def test_adds_textbox(self, slide, theme):
        add_footnote(slide, theme, "出典: 調査レポート 2026", Inches(0.6))
        assert len(slide.shapes) >= 1
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_text.py -v`
Expected: FAIL (import error)

- [ ] **Step 3: text.pyを実装**

`src/components/text.py`:
```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def add_title(slide, theme, text, left, top, width=None):
    """スライドタイトル: 左上配置、太字、下に区切り線"""
    if width is None:
        width = theme.content_width

    # タイトルテキストボックス
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.6))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_title
    run.font.bold = True
    run.font.color.rgb = theme.text_primary
    run.font.name = theme.font_title

    # 区切り線
    line_top = top + Inches(0.6)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, line_top, width, Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = theme.secondary
    line.line.fill.background()


def add_subtitle(slide, theme, text, left, top, width=None):
    """サブタイトル: タイトル直下、やや小さく"""
    if width is None:
        width = theme.content_width

    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.4))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_subtitle
    run.font.color.rgb = theme.text_secondary
    run.font.name = theme.font_body


def add_bullets(slide, theme, items, left, top, width=None, height=None):
    """箇条書き: インデント2階層対応、行頭は控えめ記号"""
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(4.0)

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        if isinstance(item, dict):
            text = item["text"]
            level = item.get("level", 0)
        else:
            text = item
            level = 0

        prefix = "  " * level + "― "
        p.text = prefix + text
        run = p.runs[0]
        run.font.size = theme.font_size_body
        run.font.color.rgb = theme.text_primary
        run.font.name = theme.font_body
        p.space_after = Pt(6)

        if level > 0:
            run.font.color.rgb = theme.text_secondary
            p.level = level


def add_callout(slide, theme, text, left, top, width=None, height=None):
    """強調ボックス: 背景色付き矩形内にテキスト"""
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(0.8)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    # primary色を薄くした背景
    r = theme.primary.red
    g = theme.primary.green
    b = theme.primary.blue
    light_r = r + (255 - r) * 9 // 10
    light_g = g + (255 - g) * 9 // 10
    light_b = b + (255 - b) * 9 // 10
    shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
    shape.line.color.rgb = theme.primary
    shape.line.width = Pt(1)

    # 左端にアクセントバー的効果のため、テキスト内にpadding
    tf = shape.text_frame
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.15)
    tf.margin_bottom = Inches(0.15)
    tf.word_wrap = True
    tf.paragraphs[0].text = text
    run = tf.paragraphs[0].runs[0]
    run.font.size = theme.font_size_body
    run.font.color.rgb = theme.text_primary
    run.font.name = theme.font_body
    run.font.bold = True


def add_footnote(slide, theme, text, left, bottom_margin=None):
    """脚注: スライド下部、小フォント"""
    if bottom_margin is None:
        bottom_margin = theme.margin_bottom

    top = theme.slide_height - bottom_margin - Inches(0.3)
    width = theme.content_width

    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.3))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.size = theme.font_size_caption
    run.font.color.rgb = theme.text_secondary
    run.font.name = theme.font_body
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_text.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/text.py tests/test_text.py
git commit -m "feat: text components (title, subtitle, bullets, callout, footnote)"
```

---

### Task 4: テーブルコンポーネント

**Files:**
- Create: `src/components/table.py`
- Create: `tests/test_table.py`

- [ ] **Step 1: テストを書く**

`tests/test_table.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.components.table import add_table
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestAddTable:
    def test_adds_table_shape(self, slide, theme):
        headers = ["項目", "2024", "2025"]
        rows = [["売上", "100", "150"], ["利益", "20", "35"]]
        add_table(slide, theme, headers, rows, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1

    def test_correct_dimensions(self, slide, theme):
        headers = ["項目", "2024", "2025"]
        rows = [["売上", "100", "150"], ["利益", "20", "35"]]
        add_table(slide, theme, headers, rows, Inches(0.6), Inches(1.5))
        table = slide.shapes[0].table
        assert table.rows.__len__() == 3  # 1 header + 2 data
        assert len(table.columns) == 3

    def test_header_text(self, slide, theme):
        headers = ["項目", "2024", "2025"]
        rows = [["売上", "100", "150"]]
        add_table(slide, theme, headers, rows, Inches(0.6), Inches(1.5))
        table = slide.shapes[0].table
        assert table.cell(0, 0).text == "項目"
        assert table.cell(0, 1).text == "2024"

    def test_data_text(self, slide, theme):
        headers = ["項目", "値"]
        rows = [["売上", "100"]]
        add_table(slide, theme, headers, rows, Inches(0.6), Inches(1.5))
        table = slide.shapes[0].table
        assert table.cell(1, 0).text == "売上"
        assert table.cell(1, 1).text == "100"
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_table.py -v`
Expected: FAIL

- [ ] **Step 3: table.pyを実装**

`src/components/table.py`:
```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def add_table(slide, theme, headers, rows, left, top, width=None, col_widths=None):
    """データ表: ヘッダ行primary色背景+白文字、ストライプ行"""
    if width is None:
        width = theme.content_width

    num_rows = len(rows) + 1  # +1 for header
    num_cols = len(headers)

    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, Inches(0.4 * num_rows))
    table = table_shape.table

    # 列幅を設定
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    else:
        col_width = width // num_cols
        for i in range(num_cols):
            table.columns[i].width = col_width

    # ヘッダ行
    for j, header_text in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header_text
        _style_cell(cell, theme, is_header=True)

    # データ行
    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i + 1, j)
            cell.text = str(cell_text)
            is_stripe = (i % 2 == 1)
            _style_cell(cell, theme, is_header=False, is_stripe=is_stripe)


def _style_cell(cell, theme, is_header=False, is_stripe=False):
    """セルのスタイルを設定"""
    # 背景色
    if is_header:
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.primary
    elif is_stripe:
        cell.fill.solid()
        r = theme.primary.red
        g = theme.primary.green
        b = theme.primary.blue
        light_r = r + (255 - r) * 95 // 100
        light_g = g + (255 - g) * 95 // 100
        light_b = b + (255 - b) * 95 // 100
        cell.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # テキストスタイル
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.1)
    cell.margin_right = Inches(0.1)
    cell.margin_top = Inches(0.05)
    cell.margin_bottom = Inches(0.05)

    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = theme.font_size_body
            run.font.name = theme.font_body
            if is_header:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
            else:
                run.font.color.rgb = theme.text_primary
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_table.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/table.py tests/test_table.py
git commit -m "feat: table component with header styling and stripe rows"
```

---

### Task 5: チャートコンポーネント

**Files:**
- Create: `src/components/chart.py`
- Create: `tests/test_chart.py`

- [ ] **Step 1: テストを書く**

`tests/test_chart.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.components.chart import add_bar_chart, add_line_chart, add_pie_chart, add_waterfall
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestBarChart:
    def test_adds_chart(self, slide, theme):
        data = {
            "labels": ["2023", "2024", "2025"],
            "series": [{"name": "売上", "values": [100, 150, 220]}],
        }
        add_bar_chart(slide, theme, data, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1

    def test_multi_series(self, slide, theme):
        data = {
            "labels": ["Q1", "Q2", "Q3"],
            "series": [
                {"name": "売上", "values": [100, 150, 220]},
                {"name": "利益", "values": [20, 35, 50]},
            ],
        }
        add_bar_chart(slide, theme, data, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1

    def test_horizontal_bar(self, slide, theme):
        data = {
            "labels": ["A", "B", "C"],
            "series": [{"name": "値", "values": [10, 20, 30]}],
        }
        add_bar_chart(slide, theme, data, Inches(0.6), Inches(1.5), horizontal=True)
        assert len(slide.shapes) >= 1


class TestLineChart:
    def test_adds_chart(self, slide, theme):
        data = {
            "labels": ["1月", "2月", "3月"],
            "series": [{"name": "推移", "values": [10, 20, 15]}],
        }
        add_line_chart(slide, theme, data, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1


class TestPieChart:
    def test_adds_chart(self, slide, theme):
        data = {
            "labels": ["製品A", "製品B", "製品C"],
            "values": [40, 35, 25],
        }
        add_pie_chart(slide, theme, data, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1


class TestWaterfall:
    def test_adds_chart(self, slide, theme):
        data = {
            "labels": ["開始", "+営業", "+開発", "-コスト", "合計"],
            "values": [100, 50, 30, -20, 160],
        }
        add_waterfall(slide, theme, data, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 1
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_chart.py -v`
Expected: FAIL

- [ ] **Step 3: chart.pyを実装**

`src/components/chart.py`:
```python
from pptx.util import Inches, Pt, Emu
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor


def add_bar_chart(slide, theme, data, left, top, width=None, height=None, horizontal=False, unit=None):
    """棒グラフ: 縦/横対応、データラベル付き"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    for series in data["series"]:
        chart_data.add_series(series["name"], series["values"])

    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    chart_frame = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)
    chart = chart_frame.chart

    _style_chart(chart, theme, data, unit)


def add_line_chart(slide, theme, data, left, top, width=None, height=None, unit=None):
    """折れ線グラフ: マーカー付き"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    for series in data["series"]:
        chart_data.add_series(series["name"], series["values"])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    _style_chart(chart, theme, data, unit)

    # マーカーサイズ設定
    for i, series in enumerate(chart.series):
        series.smooth = False
        series.format.line.width = Pt(2.5)


def add_pie_chart(slide, theme, data, left, top, width=None, height=None):
    """円グラフ: ラベル+割合表示"""
    if width is None:
        width = Inches(6.0)
    if height is None:
        height = Inches(5.0)

    chart_data = CategoryChartData()
    chart_data.categories = data["labels"]
    chart_data.add_series("", data["values"])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = theme.font_size_caption
    chart.legend.font.name = theme.font_body

    # データラベル
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_percentage = True
    data_labels.show_category_name = False
    data_labels.show_value = False
    data_labels.font.size = theme.font_size_body
    data_labels.font.name = theme.font_body

    # 色設定
    series = chart.series[0]
    for i, point in enumerate(series.points):
        color_idx = i % len(theme.chart_colors)
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = theme.chart_colors[color_idx]


def add_waterfall(slide, theme, data, left, top, width=None, height=None):
    """ウォーターフォールチャート: 増減色分け（積み上げ棒グラフで表現）"""
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(5.0)

    labels = data["labels"]
    values = data["values"]

    # ウォーターフォールを積み上げ棒グラフで実現
    # invisible: 各バーの開始位置（透明）, visible: 各バーの高さ（表示）
    invisible = []
    visible = []
    running = 0
    for i, val in enumerate(values):
        if i == 0:
            invisible.append(0)
            visible.append(val)
            running = val
        elif i == len(values) - 1:
            # 合計バー
            invisible.append(0)
            visible.append(running + val if val != running else val)
        else:
            if val >= 0:
                invisible.append(running)
                visible.append(val)
                running += val
            else:
                running += val
                invisible.append(running)
                visible.append(abs(val))

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("base", invisible)
    chart_data.add_series("value", visible)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    # 透明バー
    base_series = chart.series[0]
    base_series.format.fill.background()
    base_series.format.line.fill.background()

    # 値バー: 増減で色分け
    value_series = chart.series[1]
    for i, val in enumerate(values):
        point = value_series.points[i]
        point.format.fill.solid()
        if i == 0 or i == len(values) - 1:
            point.format.fill.fore_color.rgb = theme.primary
        elif val >= 0:
            point.format.fill.fore_color.rgb = theme.primary
        else:
            point.format.fill.fore_color.rgb = theme.secondary

    # スタイリング
    chart.has_legend = False
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.border
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = theme.font_size_caption
    value_axis.tick_labels.font.name = theme.font_body

    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = theme.font_size_caption
    category_axis.tick_labels.font.name = theme.font_body
    category_axis.format.line.color.rgb = theme.border


def _style_chart(chart, theme, data, unit=None):
    """チャート共通スタイリング"""
    # 凡例
    if len(data["series"]) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = theme.font_size_caption
        chart.legend.font.name = theme.font_body
    else:
        chart.has_legend = False

    # 値軸
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.border
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = theme.font_size_caption
    value_axis.tick_labels.font.name = theme.font_body

    # カテゴリ軸
    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = theme.font_size_caption
    category_axis.tick_labels.font.name = theme.font_body
    category_axis.format.line.color.rgb = theme.border
    category_axis.has_major_gridlines = False

    # データラベル
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_value = True
    data_labels.show_category_name = False
    data_labels.font.size = theme.font_size_caption
    data_labels.font.name = theme.font_body
    data_labels.number_format_is_linked = False

    # 系列の色
    for i, series in enumerate(chart.series):
        color_idx = i % len(theme.chart_colors)
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = theme.chart_colors[color_idx]
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_chart.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/chart.py tests/test_chart.py
git commit -m "feat: chart components (bar, line, pie, waterfall)"
```

---

### Task 6: 図形コンポーネント

**Files:**
- Create: `src/components/shape.py`
- Create: `tests/test_shape.py`

- [ ] **Step 1: テストを書く**

`tests/test_shape.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.components.shape import (
    add_matrix_2x2, add_pyramid, add_process_flow, add_cycle, add_org_chart
)
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestMatrix2x2:
    def test_adds_shapes(self, slide, theme):
        add_matrix_2x2(
            slide, theme,
            x_axis="コスト", y_axis="効果",
            quadrants=["高効果/低コスト", "高効果/高コスト", "低効果/低コスト", "低効果/高コスト"],
            left=Inches(1.0), top=Inches(1.5),
        )
        assert len(slide.shapes) >= 4  # 4象限 + 軸ラベル


class TestPyramid:
    def test_adds_shapes(self, slide, theme):
        add_pyramid(
            slide, theme,
            levels=["戦略", "戦術", "実行"],
            left=Inches(3.0), top=Inches(1.5),
        )
        assert len(slide.shapes) >= 3


class TestProcessFlow:
    def test_adds_shapes(self, slide, theme):
        add_process_flow(
            slide, theme,
            steps=["計画", "設計", "実装", "テスト"],
            left=Inches(0.6), top=Inches(2.0),
        )
        assert len(slide.shapes) >= 4  # 4ステップ + 矢印


class TestCycle:
    def test_adds_shapes(self, slide, theme):
        add_cycle(
            slide, theme,
            items=["Plan", "Do", "Check", "Act"],
            left=Inches(3.0), top=Inches(1.5),
        )
        assert len(slide.shapes) >= 4


class TestOrgChart:
    def test_adds_shapes(self, slide, theme):
        add_org_chart(
            slide, theme,
            data={"name": "CEO", "children": [
                {"name": "CTO"},
                {"name": "CFO"},
            ]},
            left=Inches(2.0), top=Inches(1.5),
        )
        assert len(slide.shapes) >= 3  # 3ノード + 接続線
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_shape.py -v`
Expected: FAIL

- [ ] **Step 3: shape.pyを実装**

`src/components/shape.py`:
```python
import math
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def add_matrix_2x2(slide, theme, x_axis, y_axis, quadrants, left, top, size=None):
    """2x2マトリクス: 軸ラベル付き、各象限にテキスト配置
    quadrants: [top-left, top-right, bottom-left, bottom-right]
    """
    if size is None:
        size = Inches(5.0)

    cell_size = size // 2
    gap = Inches(0.05)

    # 背景色を4段階で分ける
    colors = [
        theme.chart_colors[0],  # top-left
        theme.chart_colors[2],  # top-right
        theme.chart_colors[3],  # bottom-left
        theme.chart_colors[1],  # bottom-right
    ]

    positions = [
        (left, top),                          # top-left
        (left + cell_size + gap, top),        # top-right
        (left, top + cell_size + gap),        # bottom-left
        (left + cell_size + gap, top + cell_size + gap),  # bottom-right
    ]

    for i, (qtext, (ql, qt)) in enumerate(zip(quadrants, positions)):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, ql, qt, cell_size - gap, cell_size - gap
        )
        shape.fill.solid()
        # 色を薄くする
        c = colors[i]
        light_r = c.red + (255 - c.red) * 8 // 10
        light_g = c.green + (255 - c.green) * 8 // 10
        light_b = c.blue + (255 - c.blue) * 8 // 10
        shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
        shape.line.color.rgb = theme.border
        shape.line.width = Pt(0.5)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        shape.text_frame.paragraphs[0].text = qtext
        for run in shape.text_frame.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.name = theme.font_body
            run.font.color.rgb = theme.text_primary

    # X軸ラベル
    x_label = slide.shapes.add_textbox(
        left, top + size + Inches(0.1), size, Inches(0.3)
    )
    x_label.text_frame.paragraphs[0].text = x_axis
    x_label.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for run in x_label.text_frame.paragraphs[0].runs:
        run.font.size = theme.font_size_caption
        run.font.color.rgb = theme.text_secondary
        run.font.name = theme.font_body

    # Y軸ラベル
    y_label = slide.shapes.add_textbox(
        left - Inches(0.8), top + size // 2 - Inches(0.15), Inches(0.7), Inches(0.3)
    )
    y_label.text_frame.paragraphs[0].text = y_axis
    y_label.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    for run in y_label.text_frame.paragraphs[0].runs:
        run.font.size = theme.font_size_caption
        run.font.color.rgb = theme.text_secondary
        run.font.name = theme.font_body


def add_pyramid(slide, theme, levels, left, top, width=None, height=None):
    """ピラミッド: 3-5段、上から下へ幅が広がる台形の積み重ね"""
    if width is None:
        width = Inches(6.0)
    if height is None:
        height = Inches(4.5)

    n = len(levels)
    level_height = height // n
    gap = Inches(0.03)

    for i, text in enumerate(levels):
        # 上から下へ幅が広がる
        ratio_top = (i + 0.5) / n
        ratio_bottom = (i + 1.5) / n
        level_width = int(width * (0.3 + 0.7 * (i + 1) / n))
        level_left = left + (width - level_width) // 2

        shape = slide.shapes.add_shape(
            MSO_SHAPE.TRAPEZOID,
            level_left, top + level_height * i + gap * i,
            level_width, level_height - gap,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shape.line.width = Pt(2)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = text
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True


def add_process_flow(slide, theme, steps, left, top, width=None, height=None):
    """プロセスフロー: 矢印で繋がった角丸矩形、横並び"""
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(1.5)

    n = len(steps)
    arrow_width = Inches(0.4)
    total_arrow_width = arrow_width * (n - 1)
    box_width = (width - total_arrow_width) // n
    box_height = height

    for i, step_text in enumerate(steps):
        box_left = left + i * (box_width + arrow_width)

        # ステップボックス
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            box_left, top, box_width, box_height,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.fill.background()

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = step_text
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True

        # 矢印（最後のステップ以外）
        if i < n - 1:
            arrow_left = box_left + box_width
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                arrow_left, top + box_height // 2 - Inches(0.2),
                arrow_width, Inches(0.4),
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = theme.text_secondary
            arrow.line.fill.background()


def add_cycle(slide, theme, items, left, top, size=None):
    """サイクル図: 円形配置の要素"""
    if size is None:
        size = Inches(5.0)

    n = len(items)
    center_x = left + size // 2
    center_y = top + size // 2
    radius = size // 2 - Inches(0.6)
    node_size = Inches(1.4)

    for i, text in enumerate(items):
        angle = -math.pi / 2 + (2 * math.pi * i / n)
        node_x = int(center_x + radius * math.cos(angle) - node_size // 2)
        node_y = int(center_y + radius * math.sin(angle) - node_size // 2)

        # ノード（円）
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, node_x, node_y, node_size, node_size,
        )
        color_idx = i % len(theme.chart_colors)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.chart_colors[color_idx]
        shape.line.fill.background()

        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = text
        for run in tf.paragraphs[0].runs:
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title
            run.font.bold = True

        # 矢印（次ノードへ）
        next_i = (i + 1) % n
        next_angle = -math.pi / 2 + (2 * math.pi * next_i / n)
        # 矢印の開始・終了を円の縁に合わせる
        arrow_start_x = int(center_x + radius * math.cos(angle) + node_size // 2 * math.cos(
            angle + math.pi / (n * 0.8)))
        arrow_start_y = int(center_y + radius * math.sin(angle) + node_size // 2 * math.sin(
            angle + math.pi / (n * 0.8)))

        # 弧状の矢印は難しいので、小さな矢印を配置
        mid_angle = angle + math.pi / n
        arrow_x = int(center_x + (radius + Inches(0.3)) * math.cos(mid_angle) - Inches(0.15))
        arrow_y = int(center_y + (radius + Inches(0.3)) * math.sin(mid_angle) - Inches(0.15))

        arrow = slide.shapes.add_shape(
            MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, Inches(0.3), Inches(0.3),
        )
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = theme.text_secondary
        arrow.line.fill.background()
        # 矢印の回転
        arrow.rotation = math.degrees(mid_angle + math.pi / 2)


def add_org_chart(slide, theme, data, left, top, width=None, height=None):
    """組織図: ツリー構造、線で接続
    data: {"name": "CEO", "children": [{"name": "CTO"}, {"name": "CFO", "children": [...]}]}
    """
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(4.5)

    # ツリーの深さと各レベルのノード数を計算
    levels = []
    _collect_levels(data, 0, levels)

    level_height = height // len(levels)
    node_height = Inches(0.6)

    _render_org_node(slide, theme, data, left, top, width, level_height, node_height, levels)


def _collect_levels(node, depth, levels):
    """各レベルのノード数を集計"""
    while len(levels) <= depth:
        levels.append(0)
    levels[depth] += 1
    for child in node.get("children", []):
        _collect_levels(child, depth + 1, levels)


def _render_org_node(slide, theme, node, area_left, area_top, area_width, level_height, node_height, levels, depth=0):
    """組織図のノードを再帰的に描画"""
    node_width = Inches(2.0)
    node_left = area_left + (area_width - node_width) // 2
    node_top = area_top

    # ノードボックス
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, node_left, node_top, node_width, node_height,
    )
    if depth == 0:
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.primary
        text_color = RGBColor(0xFF, 0xFF, 0xFF)
    else:
        shape.fill.solid()
        color_idx = depth % len(theme.chart_colors)
        c = theme.chart_colors[color_idx]
        light_r = c.red + (255 - c.red) * 7 // 10
        light_g = c.green + (255 - c.green) * 7 // 10
        light_b = c.blue + (255 - c.blue) * 7 // 10
        shape.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
        text_color = theme.text_primary
    shape.line.color.rgb = theme.border
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].text = node["name"]
    for run in tf.paragraphs[0].runs:
        run.font.size = Pt(12)
        run.font.color.rgb = text_color
        run.font.name = theme.font_body
        if depth == 0:
            run.font.bold = True

    # 子ノードの描画
    children = node.get("children", [])
    if children:
        n = len(children)
        child_area_width = area_width // n

        for i, child in enumerate(children):
            child_left = area_left + child_area_width * i
            child_top = area_top + level_height

            # 接続線（親の下端から子の上端へ）
            parent_center_x = node_left + node_width // 2
            child_center_x = child_left + child_area_width // 2

            # 垂直線（親から中間点）
            mid_y = area_top + node_height + (level_height - node_height) // 2
            line1 = slide.shapes.add_connector(
                1,  # straight connector
                parent_center_x, node_top + node_height,
                parent_center_x, mid_y,
            )
            line1.line.color.rgb = theme.border
            line1.line.width = Pt(1.5)

            # 水平線（中間点で横移動）
            line2 = slide.shapes.add_connector(
                1, parent_center_x, mid_y, child_center_x, mid_y,
            )
            line2.line.color.rgb = theme.border
            line2.line.width = Pt(1.5)

            # 垂直線（中間点から子）
            line3 = slide.shapes.add_connector(
                1, child_center_x, mid_y, child_center_x, child_top,
            )
            line3.line.color.rgb = theme.border
            line3.line.width = Pt(1.5)

            _render_org_node(
                slide, theme, child, child_left, child_top,
                child_area_width, level_height, node_height, levels, depth + 1,
            )
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_shape.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/shape.py tests/test_shape.py
git commit -m "feat: shape components (matrix, pyramid, process flow, cycle, org chart)"
```

---

### Task 7: タイムラインコンポーネント

**Files:**
- Create: `src/components/timeline.py`
- Create: `tests/test_timeline.py`

- [ ] **Step 1: テストを書く**

`tests/test_timeline.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.components.timeline import add_timeline, add_gantt
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestTimeline:
    def test_adds_shapes(self, slide, theme):
        milestones = [
            {"date": "2026/4", "label": "キックオフ"},
            {"date": "2026/6", "label": "要件定義完了"},
            {"date": "2026/9", "label": "開発完了"},
            {"date": "2026/12", "label": "リリース"},
        ]
        add_timeline(slide, theme, milestones, Inches(0.6), Inches(2.0))
        # 軸線 + マイルストーン(丸+テキスト) * 4
        assert len(slide.shapes) >= 5


class TestGantt:
    def test_adds_shapes(self, slide, theme):
        tasks = [
            {"name": "要件定義", "start": 0, "duration": 2},
            {"name": "設計", "start": 1, "duration": 3},
            {"name": "実装", "start": 3, "duration": 4},
        ]
        phases = ["4月", "5月", "6月", "7月", "8月", "9月"]
        add_gantt(slide, theme, tasks, phases, Inches(0.6), Inches(1.5))
        assert len(slide.shapes) >= 3  # 最低3つのバー
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_timeline.py -v`
Expected: FAIL

- [ ] **Step 3: timeline.pyを実装**

`src/components/timeline.py`:
```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def add_timeline(slide, theme, milestones, left, top, width=None, height=None):
    """タイムライン: 横軸に時間、マイルストーンを上下にプロット
    milestones: [{"date": "2026/4", "label": "キックオフ"}, ...]
    """
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(3.0)

    n = len(milestones)

    # 中央の横線
    line_y = top + height // 2
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, line_y, width, Pt(3),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = theme.primary
    line.line.fill.background()

    # マイルストーン
    for i, ms in enumerate(milestones):
        x = left + (width * i) // (n - 1) if n > 1 else left + width // 2
        dot_size = Inches(0.25)

        # ドット
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x - dot_size // 2, line_y - dot_size // 2 + Pt(1),
            dot_size, dot_size,
        )
        color_idx = i % len(theme.chart_colors)
        dot.fill.solid()
        dot.fill.fore_color.rgb = theme.chart_colors[color_idx]
        dot.line.fill.background()

        # 日付ラベル（上下交互）
        is_above = (i % 2 == 0)
        label_width = Inches(1.5)
        label_x = x - label_width // 2

        if is_above:
            date_top = line_y - Inches(1.2)
            label_top = line_y - Inches(0.8)
        else:
            date_top = line_y + Inches(0.5)
            label_top = line_y + Inches(0.9)

        # 日付
        date_box = slide.shapes.add_textbox(label_x, date_top, label_width, Inches(0.3))
        date_box.text_frame.paragraphs[0].text = ms["date"]
        date_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        for run in date_box.text_frame.paragraphs[0].runs:
            run.font.size = theme.font_size_caption
            run.font.color.rgb = theme.text_secondary
            run.font.name = theme.font_body
            run.font.bold = True

        # ラベル
        label_box = slide.shapes.add_textbox(label_x, label_top, label_width, Inches(0.4))
        label_box.text_frame.word_wrap = True
        label_box.text_frame.paragraphs[0].text = ms["label"]
        label_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        for run in label_box.text_frame.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = theme.text_primary
            run.font.name = theme.font_body


def add_gantt(slide, theme, tasks, phases, left, top, width=None, height=None):
    """ガントチャート: 横棒で期間表示、フェーズ色分け
    tasks: [{"name": "要件定義", "start": 0, "duration": 2}, ...]
    phases: ["4月", "5月", "6月", ...]  ヘッダラベル
    """
    if width is None:
        width = theme.content_width
    if height is None:
        height = Inches(4.0)

    n_tasks = len(tasks)
    n_phases = len(phases)

    label_width = Inches(2.0)
    chart_width = width - label_width
    phase_width = chart_width // n_phases
    row_height = min(Inches(0.6), height // (n_tasks + 1))

    # フェーズヘッダ
    for j, phase_label in enumerate(phases):
        hdr_left = left + label_width + phase_width * j
        hdr = slide.shapes.add_textbox(hdr_left, top, phase_width, row_height)
        hdr.text_frame.paragraphs[0].text = phase_label
        hdr.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        for run in hdr.text_frame.paragraphs[0].runs:
            run.font.size = theme.font_size_caption
            run.font.color.rgb = theme.text_secondary
            run.font.name = theme.font_body
            run.font.bold = True

    # ヘッダ下線
    hdr_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left + label_width, top + row_height - Pt(1),
        chart_width, Pt(1),
    )
    hdr_line.fill.solid()
    hdr_line.fill.fore_color.rgb = theme.border
    hdr_line.line.fill.background()

    # タスク行
    for i, task in enumerate(tasks):
        row_top = top + row_height * (i + 1) + Inches(0.1)

        # タスク名
        name_box = slide.shapes.add_textbox(left, row_top, label_width, row_height)
        name_box.text_frame.paragraphs[0].text = task["name"]
        for run in name_box.text_frame.paragraphs[0].runs:
            run.font.size = theme.font_size_body
            run.font.color.rgb = theme.text_primary
            run.font.name = theme.font_body

        # ガントバー
        bar_left = left + label_width + phase_width * task["start"]
        bar_width = phase_width * task["duration"]
        bar_height = row_height - Inches(0.15)

        bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            bar_left, row_top + Inches(0.05), bar_width, bar_height,
        )
        color_idx = i % len(theme.chart_colors)
        bar.fill.solid()
        bar.fill.fore_color.rgb = theme.chart_colors[color_idx]
        bar.line.fill.background()

        # バー内テキスト
        tf = bar.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = task["name"]
        for run in tf.paragraphs[0].runs:
            run.font.size = theme.font_size_caption
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_body
            run.font.bold = True
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_timeline.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/timeline.py tests/test_timeline.py
git commit -m "feat: timeline and gantt chart components"
```

---

### Task 8: アイコンコンポーネント

**Files:**
- Create: `src/components/icon.py`
- Create: `tests/test_icon.py`

- [ ] **Step 1: テストを書く**

`tests/test_icon.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.components.icon import add_icon_with_label, add_kpi_card, add_icon_row
from src.themes.monotone import MonotoneTheme


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]
    return prs.slides.add_slide(layout)


class TestIconWithLabel:
    def test_adds_shapes(self, slide, theme):
        add_icon_with_label(slide, theme, "check", "完了", Inches(1.0), Inches(1.0))
        assert len(slide.shapes) >= 2  # 円 + テキスト


class TestKpiCard:
    def test_adds_shapes(self, slide, theme):
        add_kpi_card(slide, theme, "125", "億円", "年間売上", Inches(1.0), Inches(1.0))
        assert len(slide.shapes) >= 1

    def test_text_content(self, slide, theme):
        add_kpi_card(slide, theme, "125", "億円", "年間売上", Inches(1.0), Inches(1.0))
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        texts.append(p.text.strip())
        assert any("125" in t for t in texts)


class TestIconRow:
    def test_adds_shapes(self, slide, theme):
        items = [
            {"icon": "circle", "label": "項目1"},
            {"icon": "square", "label": "項目2"},
            {"icon": "triangle", "label": "項目3"},
        ]
        add_icon_row(slide, theme, items, Inches(0.6), Inches(2.0))
        assert len(slide.shapes) >= 6  # 3アイコン + 3テキスト
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_icon.py -v`
Expected: FAIL

- [ ] **Step 3: icon.pyを実装**

`src/components/icon.py`:
```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# アイコン形状マッピング
ICON_SHAPES = {
    "circle": MSO_SHAPE.OVAL,
    "square": MSO_SHAPE.ROUNDED_RECTANGLE,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
    "star": MSO_SHAPE.STAR_5_POINT,
    "check": MSO_SHAPE.OVAL,
    "arrow_right": MSO_SHAPE.RIGHT_ARROW,
    "arrow_up": MSO_SHAPE.UP_ARROW,
    "hexagon": MSO_SHAPE.HEXAGON,
    "lightning": MSO_SHAPE.LIGHTNING_BOLT,
}

# アイコン内に表示するシンボル文字
ICON_SYMBOLS = {
    "check": "✓",
    "arrow_right": "→",
    "arrow_up": "↑",
    "circle": "",
    "square": "",
    "triangle": "",
    "diamond": "",
    "star": "",
    "hexagon": "",
    "lightning": "",
}


def add_icon_with_label(slide, theme, icon_type, label, left, top, size=None, color_idx=0):
    """アイコン+ラベル: 丸/角丸内に幾何学記号+下にテキスト"""
    if size is None:
        size = Inches(0.8)

    shape_type = ICON_SHAPES.get(icon_type, MSO_SHAPE.OVAL)
    symbol = ICON_SYMBOLS.get(icon_type, "")

    # アイコン図形
    icon = slide.shapes.add_shape(shape_type, left, top, size, size)
    c_idx = color_idx % len(theme.chart_colors)
    icon.fill.solid()
    icon.fill.fore_color.rgb = theme.chart_colors[c_idx]
    icon.line.fill.background()

    if symbol:
        tf = icon.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].text = symbol
        for run in tf.paragraphs[0].runs:
            run.font.size = Pt(int(size / Inches(1) * 18))
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.bold = True

    # ラベル
    label_box = slide.shapes.add_textbox(
        left - Inches(0.3), top + size + Inches(0.1),
        size + Inches(0.6), Inches(0.4),
    )
    label_box.text_frame.word_wrap = True
    label_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    label_box.text_frame.paragraphs[0].text = label
    for run in label_box.text_frame.paragraphs[0].runs:
        run.font.size = theme.font_size_caption
        run.font.color.rgb = theme.text_primary
        run.font.name = theme.font_body


def add_kpi_card(slide, theme, value, unit, label, left, top, width=None, height=None, color_idx=0):
    """KPI表示: 大きな数字+単位+ラベル、カード型"""
    if width is None:
        width = Inches(3.0)
    if height is None:
        height = Inches(2.0)

    # カード背景
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
    )
    card.fill.solid()
    c = theme.chart_colors[color_idx % len(theme.chart_colors)]
    # 薄い背景色
    light_r = c.red + (255 - c.red) * 92 // 100
    light_g = c.green + (255 - c.green) * 92 // 100
    light_b = c.blue + (255 - c.blue) * 92 // 100
    card.fill.fore_color.rgb = RGBColor(light_r, light_g, light_b)
    card.line.color.rgb = theme.border
    card.line.width = Pt(0.5)

    tf = card.text_frame
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.2)
    tf.word_wrap = True

    # 数値 + 単位
    p_value = tf.paragraphs[0]
    p_value.alignment = PP_ALIGN.CENTER
    run_val = p_value.add_run()
    run_val.text = value
    run_val.font.size = Pt(36)
    run_val.font.bold = True
    run_val.font.color.rgb = theme.text_primary
    run_val.font.name = theme.font_title

    run_unit = p_value.add_run()
    run_unit.text = " " + unit
    run_unit.font.size = Pt(16)
    run_unit.font.color.rgb = theme.text_secondary
    run_unit.font.name = theme.font_body

    # ラベル
    p_label = tf.add_paragraph()
    p_label.alignment = PP_ALIGN.CENTER
    p_label.space_before = Pt(8)
    run_label = p_label.add_run()
    run_label.text = label
    run_label.font.size = theme.font_size_body
    run_label.font.color.rgb = theme.text_secondary
    run_label.font.name = theme.font_body


def add_icon_row(slide, theme, items, left, top, width=None, icon_size=None):
    """アイコン横並び: 3-5個のアイコン+ラベルを等間隔配置
    items: [{"icon": "circle", "label": "項目1"}, ...]
    """
    if width is None:
        width = theme.content_width
    if icon_size is None:
        icon_size = Inches(0.8)

    n = len(items)
    spacing = width // n

    for i, item in enumerate(items):
        icon_left = left + spacing * i + (spacing - icon_size) // 2
        add_icon_with_label(
            slide, theme,
            item["icon"], item["label"],
            icon_left, top,
            size=icon_size,
            color_idx=i,
        )
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_icon.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/components/icon.py tests/test_icon.py
git commit -m "feat: icon components (icon with label, KPI card, icon row)"
```

---

### Task 9: レイアウト — 表紙・セクション区切り・アジェンダ

**Files:**
- Create: `src/layouts/cover.py`
- Create: `src/layouts/section_divider.py`
- Create: `src/layouts/agenda.py`
- Create: `tests/test_layouts.py`

- [ ] **Step 1: テストを書く**

`tests/test_layouts.py`:
```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.themes.monotone import MonotoneTheme
from src.themes.dark import DarkTheme
from src.layouts.cover import CoverLayout
from src.layouts.section_divider import SectionDividerLayout
from src.layouts.agenda import AgendaLayout
from src.layouts.content import ContentLayout
from src.layouts.chart_page import ChartPageLayout
from src.layouts.comparison import ComparisonLayout
from src.layouts.closing import ClosingLayout


@pytest.fixture
def theme():
    return MonotoneTheme()


@pytest.fixture
def dark_theme():
    return DarkTheme()


@pytest.fixture
def prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def make_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


class TestCoverLayout:
    def test_render(self, prs, theme):
        slide = make_slide(prs)
        layout = CoverLayout()
        layout.render(slide, theme, {
            "title": "DX推進戦略提案書",
            "subtitle": "2026年度計画",
            "client": "株式会社ABC",
            "date": "2026年4月",
        })
        assert len(slide.shapes) >= 2

    def test_render_dark_theme(self, prs, dark_theme):
        slide = make_slide(prs)
        layout = CoverLayout()
        layout.render(slide, dark_theme, {
            "title": "テスト",
        })
        assert len(slide.shapes) >= 1


class TestSectionDividerLayout:
    def test_render(self, prs, theme):
        slide = make_slide(prs)
        layout = SectionDividerLayout()
        layout.render(slide, theme, {
            "section_number": 1,
            "section_title": "現状分析",
        })
        assert len(slide.shapes) >= 1


class TestAgendaLayout:
    def test_render(self, prs, theme):
        slide = make_slide(prs)
        layout = AgendaLayout()
        layout.render(slide, theme, {
            "items": ["現状分析", "課題整理", "戦略提案", "実行計画"],
        })
        assert len(slide.shapes) >= 1

    def test_render_with_highlight(self, prs, theme):
        slide = make_slide(prs)
        layout = AgendaLayout()
        layout.render(slide, theme, {
            "items": ["現状分析", "課題整理", "戦略提案", "実行計画"],
            "highlight": 1,
        })
        assert len(slide.shapes) >= 1


class TestContentLayout:
    def test_render_single_column(self, prs, theme):
        slide = make_slide(prs)
        layout = ContentLayout()
        layout.render(slide, theme, {
            "title": "テストタイトル",
            "columns": 1,
            "components": [
                {"type": "bullets", "items": ["項目1", "項目2"]},
            ],
        })
        assert len(slide.shapes) >= 2  # タイトル + 区切り線 + 箇条書き

    def test_render_two_columns(self, prs, theme):
        slide = make_slide(prs)
        layout = ContentLayout()
        layout.render(slide, theme, {
            "title": "比較",
            "columns": 2,
            "components": [
                {"type": "bullets", "items": ["左1", "左2"]},
                {"type": "bullets", "items": ["右1", "右2"]},
            ],
        })
        assert len(slide.shapes) >= 3


class TestChartPageLayout:
    def test_render_with_key_points(self, prs, theme):
        slide = make_slide(prs)
        layout = ChartPageLayout()
        layout.render(slide, theme, {
            "title": "売上推移",
            "chart": {
                "type": "bar",
                "data": {
                    "labels": ["2023", "2024", "2025"],
                    "series": [{"name": "売上", "values": [100, 150, 220]}],
                },
            },
            "key_points": ["成長率15%", "目標達成"],
        })
        assert len(slide.shapes) >= 3

    def test_render_full_width(self, prs, theme):
        slide = make_slide(prs)
        layout = ChartPageLayout()
        layout.render(slide, theme, {
            "title": "市場分析",
            "chart": {
                "type": "line",
                "data": {
                    "labels": ["Q1", "Q2", "Q3"],
                    "series": [{"name": "値", "values": [10, 20, 30]}],
                },
            },
        })
        assert len(slide.shapes) >= 2


class TestComparisonLayout:
    def test_render(self, prs, theme):
        slide = make_slide(prs)
        layout = ComparisonLayout()
        layout.render(slide, theme, {
            "title": "Before / After",
            "left_title": "Before",
            "left_components": [
                {"type": "bullets", "items": ["旧手法1", "旧手法2"]},
            ],
            "right_title": "After",
            "right_components": [
                {"type": "bullets", "items": ["新手法1", "新手法2"]},
            ],
        })
        assert len(slide.shapes) >= 4


class TestClosingLayout:
    def test_render_summary(self, prs, theme):
        slide = make_slide(prs)
        layout = ClosingLayout()
        layout.render(slide, theme, {
            "summary": ["要点1", "要点2", "要点3"],
            "next_steps": ["ステップ1", "ステップ2"],
        })
        assert len(slide.shapes) >= 2

    def test_render_thank_you(self, prs, theme):
        slide = make_slide(prs)
        layout = ClosingLayout()
        layout.render(slide, theme, {
            "type": "thank_you",
            "contact": "example@company.com",
        })
        assert len(slide.shapes) >= 1
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_layouts.py -v`
Expected: FAIL

- [ ] **Step 3: cover.pyを実装**

`src/layouts/cover.py`:
```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


class CoverLayout:
    def render(self, slide, theme, data):
        """表紙: 中央にタイトル、サブタイトル、左下にクライアント名、右下に日付"""
        sw = theme.slide_width
        sh = theme.slide_height

        # 背景塗り
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, sh)
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.background
        bg.line.fill.background()

        # 下部アクセントバー
        bar_height = Inches(0.08)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, sh - Inches(1.2), sw, bar_height,
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = theme.secondary
        bar.line.fill.background()

        # タイトル（中央）
        title_width = Inches(10.0)
        title_left = (sw - title_width) // 2
        title_top = sh // 2 - Inches(1.2)

        txBox = slide.shapes.add_textbox(title_left, title_top, title_width, Inches(1.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = data.get("title", "")
        run = p.runs[0]
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = theme.text_primary
        run.font.name = theme.font_title

        # サブタイトル
        subtitle = data.get("subtitle", "")
        if subtitle:
            sub_box = slide.shapes.add_textbox(
                title_left, title_top + Inches(1.2), title_width, Inches(0.6),
            )
            tf2 = sub_box.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            p2.text = subtitle
            run2 = p2.runs[0]
            run2.font.size = Pt(20)
            run2.font.color.rgb = theme.text_secondary
            run2.font.name = theme.font_body

        # クライアント名（左下）
        client = data.get("client", "")
        if client:
            cl_box = slide.shapes.add_textbox(
                theme.margin_left, sh - Inches(1.0), Inches(4.0), Inches(0.4),
            )
            cl_box.text_frame.paragraphs[0].text = client
            for run in cl_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(14)
                run.font.color.rgb = theme.text_secondary
                run.font.name = theme.font_body

        # 日付（右下）
        date = data.get("date", "")
        if date:
            dt_box = slide.shapes.add_textbox(
                sw - theme.margin_right - Inches(3.0), sh - Inches(1.0),
                Inches(3.0), Inches(0.4),
            )
            dt_box.text_frame.paragraphs[0].text = date
            dt_box.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            for run in dt_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(14)
                run.font.color.rgb = theme.text_secondary
                run.font.name = theme.font_body
```

- [ ] **Step 4: section_divider.pyを実装**

`src/layouts/section_divider.py`:
```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


class SectionDividerLayout:
    def render(self, slide, theme, data):
        """セクション区切り: primary色背景、中央にセクション番号+名前"""
        sw = theme.slide_width
        sh = theme.slide_height

        # 背景をprimary色で塗る
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, sh)
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.primary
        bg.line.fill.background()

        # セクション番号
        section_number = data.get("section_number", "")
        section_title = data.get("section_title", "")

        center_width = Inches(8.0)
        center_left = (sw - center_width) // 2

        if section_number:
            num_box = slide.shapes.add_textbox(
                center_left, sh // 2 - Inches(1.5), center_width, Inches(1.0),
            )
            num_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            num_box.text_frame.paragraphs[0].text = f"{section_number:02d}" if isinstance(section_number, int) else str(section_number)
            for run in num_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(60)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = theme.font_title

        # セクション名
        title_box = slide.shapes.add_textbox(
            center_left, sh // 2 - Inches(0.3), center_width, Inches(0.8),
        )
        title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_box.text_frame.paragraphs[0].text = section_title
        for run in title_box.text_frame.paragraphs[0].runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title

        # 下部アクセントライン
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            (sw - Inches(3.0)) // 2, sh // 2 + Inches(0.7),
            Inches(3.0), Pt(3),
        )
        line.fill.solid()
        line.fill.fore_color.rgb = theme.secondary
        line.line.fill.background()
```

- [ ] **Step 5: agenda.pyを実装**

`src/layouts/agenda.py`:
```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from src.components.text import add_title


class AgendaLayout:
    def render(self, slide, theme, data):
        """アジェンダ: タイトル + 番号付きリスト"""
        add_title(slide, theme, "Agenda", theme.margin_left, theme.margin_top)

        items = data.get("items", [])
        highlight = data.get("highlight", None)

        item_top = theme.content_area_top + Inches(0.2)
        item_height = Inches(0.7)

        for i, item_text in enumerate(items):
            y = item_top + item_height * i
            is_active = (highlight is not None and i == highlight)

            # 番号
            num_width = Inches(0.8)
            num_box = slide.shapes.add_textbox(
                theme.margin_left + Inches(0.5), y, num_width, item_height,
            )
            num_box.text_frame.paragraphs[0].text = f"{i + 1:02d}"
            num_box.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            for run in num_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(28)
                run.font.bold = True
                run.font.name = theme.font_title
                if is_active:
                    run.font.color.rgb = theme.secondary
                else:
                    run.font.color.rgb = theme.primary

            # テキスト
            text_box = slide.shapes.add_textbox(
                theme.margin_left + Inches(1.6), y + Inches(0.1),
                Inches(8.0), item_height,
            )
            text_box.text_frame.paragraphs[0].text = item_text
            for run in text_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(20)
                run.font.name = theme.font_body
                if is_active:
                    run.font.color.rgb = theme.text_primary
                    run.font.bold = True
                else:
                    run.font.color.rgb = theme.text_secondary

            # 区切り線
            if i < len(items) - 1:
                sep = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    theme.margin_left + Inches(1.6), y + item_height - Pt(1),
                    Inches(8.0), Pt(1),
                )
                sep.fill.solid()
                sep.fill.fore_color.rgb = theme.border
                sep.line.fill.background()
```

- [ ] **Step 6: テスト通過を確認（cover, section_divider, agendaのみ）**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_layouts.py::TestCoverLayout tests/test_layouts.py::TestSectionDividerLayout tests/test_layouts.py::TestAgendaLayout -v`
Expected: ALL PASS

- [ ] **Step 7: コミット**

```bash
git add src/layouts/cover.py src/layouts/section_divider.py src/layouts/agenda.py tests/test_layouts.py
git commit -m "feat: cover, section divider, and agenda layouts"
```

---

### Task 10: レイアウト — コンテンツ・チャートページ

**Files:**
- Create: `src/layouts/content.py`
- Create: `src/layouts/chart_page.py`

- [ ] **Step 1: content.pyを実装**

`src/layouts/content.py`:
```python
from pptx.util import Inches, Pt

from src.components.text import add_title, add_bullets, add_callout
from src.components.table import add_table
from src.components.shape import (
    add_matrix_2x2, add_pyramid, add_process_flow, add_cycle, add_org_chart,
)
from src.components.timeline import add_timeline, add_gantt
from src.components.icon import add_icon_row, add_kpi_card


class ContentLayout:
    def render(self, slide, theme, data):
        """汎用コンテンツ: 1/2/3カラム切り替え"""
        title = data.get("title", "")
        columns = data.get("columns", 1)
        components = data.get("components", [])

        if title:
            add_title(slide, theme, title, theme.margin_left, theme.margin_top)

        content_top = theme.content_area_top
        content_width = theme.content_width

        if columns == 1:
            self._render_components(
                slide, theme, components,
                theme.margin_left, content_top, content_width,
            )
        elif columns == 2:
            col_width = (content_width - Inches(0.4)) // 2
            # 左カラム: 最初のコンポーネント
            left_comps = components[:len(components)//2] if len(components) > 1 else components
            right_comps = components[len(components)//2:] if len(components) > 1 else []

            self._render_components(
                slide, theme, left_comps,
                theme.margin_left, content_top, col_width,
            )
            self._render_components(
                slide, theme, right_comps,
                theme.margin_left + col_width + Inches(0.4), content_top, col_width,
            )
        elif columns == 3:
            col_width = (content_width - Inches(0.8)) // 3
            third = max(1, len(components) // 3)
            for col_i in range(3):
                col_comps = components[col_i * third:(col_i + 1) * third]
                col_left = theme.margin_left + col_i * (col_width + Inches(0.4))
                self._render_components(
                    slide, theme, col_comps, col_left, content_top, col_width,
                )

    def _render_components(self, slide, theme, components, left, top, width):
        """コンポーネントリストを順番に描画"""
        current_top = top
        for comp in components:
            comp_type = comp.get("type", "")
            if comp_type == "bullets":
                items = comp.get("items", [])
                height = Inches(0.4 * len(items))
                add_bullets(slide, theme, items, left, current_top, width=width, height=height)
                current_top += height + Inches(0.2)

            elif comp_type == "callout":
                text = comp.get("text", "")
                add_callout(slide, theme, text, left, current_top, width=width)
                current_top += Inches(1.0)

            elif comp_type == "table":
                headers = comp.get("headers", [])
                rows = comp.get("rows", [])
                add_table(slide, theme, headers, rows, left, current_top, width=width)
                current_top += Inches(0.4 * (len(rows) + 1)) + Inches(0.2)

            elif comp_type == "matrix_2x2":
                add_matrix_2x2(
                    slide, theme,
                    x_axis=comp.get("x_axis", ""),
                    y_axis=comp.get("y_axis", ""),
                    quadrants=comp.get("quadrants", ["", "", "", ""]),
                    left=left, top=current_top,
                )
                current_top += Inches(5.5)

            elif comp_type == "pyramid":
                add_pyramid(slide, theme, comp.get("levels", []), left, current_top, width=width)
                current_top += Inches(5.0)

            elif comp_type == "process_flow":
                add_process_flow(slide, theme, comp.get("steps", []), left, current_top, width=width)
                current_top += Inches(2.0)

            elif comp_type == "cycle":
                add_cycle(slide, theme, comp.get("items", []), left, current_top)
                current_top += Inches(5.5)

            elif comp_type == "org_chart":
                add_org_chart(slide, theme, comp.get("data", {}), left, current_top, width=width)
                current_top += Inches(5.0)

            elif comp_type == "timeline":
                add_timeline(slide, theme, comp.get("milestones", []), left, current_top, width=width)
                current_top += Inches(3.5)

            elif comp_type == "gantt":
                add_gantt(
                    slide, theme,
                    comp.get("tasks", []), comp.get("phases", []),
                    left, current_top, width=width,
                )
                current_top += Inches(4.5)

            elif comp_type == "icon_row":
                add_icon_row(slide, theme, comp.get("items", []), left, current_top, width=width)
                current_top += Inches(2.0)

            elif comp_type == "kpi_cards":
                cards = comp.get("cards", [])
                card_width = (width - Inches(0.3) * (len(cards) - 1)) // len(cards)
                for ci, card in enumerate(cards):
                    card_left = left + ci * (card_width + Inches(0.3))
                    add_kpi_card(
                        slide, theme,
                        card.get("value", ""), card.get("unit", ""), card.get("label", ""),
                        card_left, current_top,
                        width=card_width, color_idx=ci,
                    )
                current_top += Inches(2.5)
```

- [ ] **Step 2: chart_page.pyを実装**

`src/layouts/chart_page.py`:
```python
from pptx.util import Inches, Pt

from src.components.text import add_title, add_bullets
from src.components.chart import add_bar_chart, add_line_chart, add_pie_chart, add_waterfall


CHART_FUNCTIONS = {
    "bar": add_bar_chart,
    "line": add_line_chart,
    "pie": add_pie_chart,
    "waterfall": add_waterfall,
}


class ChartPageLayout:
    def render(self, slide, theme, data):
        """チャート主体ページ: チャート + オプションでキーポイント"""
        title = data.get("title", "")
        chart_data = data.get("chart", {})
        key_points = data.get("key_points", [])

        if title:
            add_title(slide, theme, title, theme.margin_left, theme.margin_top)

        content_top = theme.content_area_top
        chart_type = chart_data.get("type", "bar")
        chart_func = CHART_FUNCTIONS.get(chart_type, add_bar_chart)
        unit = chart_data.get("unit", None)

        if key_points:
            # 左にチャート(65%) + 右にキーポイント(30%)
            chart_width = int(theme.content_width * 0.65)
            chart_func(
                slide, theme, chart_data["data"],
                theme.margin_left, content_top,
                width=chart_width, height=theme.content_height,
                **({"unit": unit} if unit else {}),
            )

            kp_left = theme.margin_left + chart_width + Inches(0.4)
            kp_width = int(theme.content_width * 0.30)
            add_bullets(
                slide, theme, key_points,
                kp_left, content_top + Inches(0.5),
                width=kp_width,
            )
        else:
            # チャート全面表示
            chart_func(
                slide, theme, chart_data["data"],
                theme.margin_left, content_top,
                width=theme.content_width, height=theme.content_height,
                **({"unit": unit} if unit else {}),
            )
```

- [ ] **Step 3: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_layouts.py::TestContentLayout tests/test_layouts.py::TestChartPageLayout -v`
Expected: ALL PASS

- [ ] **Step 4: コミット**

```bash
git add src/layouts/content.py src/layouts/chart_page.py
git commit -m "feat: content and chart page layouts"
```

---

### Task 11: レイアウト — 比較・まとめ

**Files:**
- Create: `src/layouts/comparison.py`
- Create: `src/layouts/closing.py`

- [ ] **Step 1: comparison.pyを実装**

`src/layouts/comparison.py`:
```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from src.components.text import add_title
from src.layouts.content import ContentLayout


class ComparisonLayout:
    def render(self, slide, theme, data):
        """比較ページ: 左右2分割、中央に区切り線"""
        title = data.get("title", "")
        left_title = data.get("left_title", "")
        right_title = data.get("right_title", "")
        left_components = data.get("left_components", [])
        right_components = data.get("right_components", [])

        if title:
            add_title(slide, theme, title, theme.margin_left, theme.margin_top)

        content_top = theme.content_area_top
        half_width = (theme.content_width - Inches(0.6)) // 2

        # 左側タイトル
        if left_title:
            lt_box = slide.shapes.add_textbox(
                theme.margin_left, content_top, half_width, Inches(0.5),
            )
            lt_box.text_frame.paragraphs[0].text = left_title
            lt_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            for run in lt_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = theme.primary
                run.font.name = theme.font_title

        # 右側タイトル
        right_left = theme.margin_left + half_width + Inches(0.6)
        if right_title:
            rt_box = slide.shapes.add_textbox(
                right_left, content_top, half_width, Inches(0.5),
            )
            rt_box.text_frame.paragraphs[0].text = right_title
            rt_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            for run in rt_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = theme.secondary
                run.font.name = theme.font_title

        # 中央区切り線
        divider_x = theme.margin_left + half_width + Inches(0.25)
        divider = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, divider_x, content_top, Pt(2), theme.content_height,
        )
        divider.fill.solid()
        divider.fill.fore_color.rgb = theme.border
        divider.line.fill.background()

        # 左側コンポーネント
        comp_top = content_top + Inches(0.7)
        content_layout = ContentLayout()
        content_layout._render_components(
            slide, theme, left_components,
            theme.margin_left, comp_top, half_width,
        )

        # 右側コンポーネント
        content_layout._render_components(
            slide, theme, right_components,
            right_left, comp_top, half_width,
        )
```

- [ ] **Step 2: closing.pyを実装**

`src/layouts/closing.py`:
```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from src.components.text import add_title, add_bullets


class ClosingLayout:
    def render(self, slide, theme, data):
        """まとめ/Next Steps or Thank You"""
        close_type = data.get("type", "summary")

        if close_type == "thank_you":
            self._render_thank_you(slide, theme, data)
        else:
            self._render_summary(slide, theme, data)

    def _render_summary(self, slide, theme, data):
        """まとめ + Next Steps"""
        summary = data.get("summary", [])
        next_steps = data.get("next_steps", [])

        # まとめセクション
        add_title(slide, theme, "Summary", theme.margin_left, theme.margin_top)

        if summary:
            add_bullets(
                slide, theme, summary,
                theme.margin_left, theme.content_area_top,
                width=theme.content_width,
                height=Inches(0.4 * len(summary)),
            )

        # Next Steps セクション
        if next_steps:
            ns_top = theme.content_area_top + Inches(0.4 * len(summary)) + Inches(0.8)

            # Next Stepsヘッダ
            ns_header = slide.shapes.add_textbox(
                theme.margin_left, ns_top - Inches(0.5),
                Inches(3.0), Inches(0.4),
            )
            ns_header.text_frame.paragraphs[0].text = "Next Steps"
            for run in ns_header.text_frame.paragraphs[0].runs:
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = theme.primary
                run.font.name = theme.font_title

            # 番号付きリスト
            for i, step in enumerate(next_steps):
                step_top = ns_top + Inches(0.5 * i)

                # 番号（丸囲み）
                num_size = Inches(0.35)
                num_shape = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    theme.margin_left, step_top, num_size, num_size,
                )
                num_shape.fill.solid()
                num_shape.fill.fore_color.rgb = theme.primary
                num_shape.line.fill.background()
                num_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                num_shape.text_frame.paragraphs[0].text = str(i + 1)
                for run in num_shape.text_frame.paragraphs[0].runs:
                    run.font.size = Pt(12)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.name = theme.font_title

                # テキスト
                step_box = slide.shapes.add_textbox(
                    theme.margin_left + Inches(0.6), step_top,
                    theme.content_width - Inches(0.6), Inches(0.4),
                )
                step_box.text_frame.paragraphs[0].text = step
                for run in step_box.text_frame.paragraphs[0].runs:
                    run.font.size = theme.font_size_body
                    run.font.color.rgb = theme.text_primary
                    run.font.name = theme.font_body

    def _render_thank_you(self, slide, theme, data):
        """Thank You ページ"""
        sw = theme.slide_width
        sh = theme.slide_height

        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, sh)
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.primary
        bg.line.fill.background()

        # Thank You テキスト
        ty_width = Inches(8.0)
        ty_left = (sw - ty_width) // 2
        ty_box = slide.shapes.add_textbox(ty_left, sh // 2 - Inches(1.0), ty_width, Inches(1.0))
        ty_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        ty_box.text_frame.paragraphs[0].text = "Thank You"
        for run in ty_box.text_frame.paragraphs[0].runs:
            run.font.size = Pt(48)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.name = theme.font_title

        # 連絡先
        contact = data.get("contact", "")
        if contact:
            ct_box = slide.shapes.add_textbox(ty_left, sh // 2 + Inches(0.3), ty_width, Inches(0.5))
            ct_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            ct_box.text_frame.paragraphs[0].text = contact
            for run in ct_box.text_frame.paragraphs[0].runs:
                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = theme.font_body
```

- [ ] **Step 3: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_layouts.py::TestComparisonLayout tests/test_layouts.py::TestClosingLayout -v`
Expected: ALL PASS

- [ ] **Step 4: コミット**

```bash
git add src/layouts/comparison.py src/layouts/closing.py
git commit -m "feat: comparison and closing layouts"
```

---

### Task 12: ジェネレーター

**Files:**
- Create: `src/generator.py`
- Create: `tests/test_generator.py`

- [ ] **Step 1: テストを書く**

`tests/test_generator.py`:
```python
import pytest
import os
import tempfile

from src.generator import generate_pptx


@pytest.fixture
def sample_config():
    return {
        "theme": "monotone",
        "title": "テスト資料",
        "slides": [
            {
                "layout": "cover",
                "title": "テスト資料",
                "subtitle": "テスト用",
                "date": "2026年4月",
            },
            {
                "layout": "agenda",
                "items": ["項目1", "項目2", "項目3"],
            },
            {
                "layout": "section_divider",
                "section_number": 1,
                "section_title": "セクション1",
            },
            {
                "layout": "content",
                "columns": 1,
                "title": "テストコンテンツ",
                "components": [
                    {"type": "bullets", "items": ["要点1", "要点2"]},
                ],
            },
            {
                "layout": "chart_page",
                "title": "テストチャート",
                "chart": {
                    "type": "bar",
                    "data": {
                        "labels": ["A", "B", "C"],
                        "series": [{"name": "値", "values": [10, 20, 30]}],
                    },
                },
            },
            {
                "layout": "closing",
                "summary": ["まとめ1", "まとめ2"],
                "next_steps": ["次1", "次2"],
            },
        ],
    }


class TestGeneratePptx:
    def test_generates_file(self, sample_config):
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            generate_pptx(sample_config, output_path)
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 0
        finally:
            os.unlink(output_path)

    def test_correct_slide_count(self, sample_config):
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(sample_config, output_path)
            assert len(prs.slides) == 6
        finally:
            os.unlink(output_path)

    def test_unknown_layout_skipped(self):
        config = {
            "theme": "monotone",
            "slides": [
                {"layout": "nonexistent_layout"},
                {"layout": "cover", "title": "テスト"},
            ],
        }
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(config, output_path)
            assert len(prs.slides) == 1  # unknownはスキップ
        finally:
            os.unlink(output_path)

    def test_dark_theme(self):
        config = {
            "theme": "dark",
            "slides": [
                {"layout": "cover", "title": "ダークテーマ"},
            ],
        }
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(config, output_path)
            assert len(prs.slides) == 1
        finally:
            os.unlink(output_path)

    def test_colorful_theme(self):
        config = {
            "theme": "colorful",
            "slides": [
                {"layout": "cover", "title": "カラフルテーマ"},
            ],
        }
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(config, output_path)
            assert len(prs.slides) == 1
        finally:
            os.unlink(output_path)
```

- [ ] **Step 2: テスト失敗を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_generator.py -v`
Expected: FAIL

- [ ] **Step 3: generator.pyを実装**

`src/generator.py`:
```python
import sys
from pptx import Presentation
from pptx.util import Inches

from src.themes import get_theme
from src.layouts import get_layout


def generate_pptx(config: dict, output_path: str) -> Presentation:
    """構成JSONを受け取り.pptxを生成する

    Args:
        config: スライド構成を定義するdict
        output_path: 出力先ファイルパス

    Returns:
        生成されたPresentationオブジェクト
    """
    theme_name = config.get("theme", "monotone")
    theme = get_theme(theme_name)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = config.get("slides", [])

    for slide_data in slides_data:
        layout_name = slide_data.get("layout", "")

        try:
            layout = get_layout(layout_name)
        except ValueError as e:
            print(f"Warning: {e} - skipping slide", file=sys.stderr)
            continue

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # スライド背景を設定
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = theme.background

        layout.render(slide, theme, slide_data)

    prs.save(output_path)
    return prs


if __name__ == "__main__":
    import json

    if len(sys.argv) < 2:
        print("Usage: python -m src.generator <config.json> [output.pptx]")
        sys.exit(1)

    config_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "output.pptx"

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    generate_pptx(config, output_path)
    print(f"Generated: {output_path}")
```

- [ ] **Step 4: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_generator.py -v`
Expected: ALL PASS

- [ ] **Step 5: コミット**

```bash
git add src/generator.py tests/test_generator.py
git commit -m "feat: generator entry point - JSON config to .pptx"
```

---

### Task 13: E2Eテスト

**Files:**
- Create: `tests/test_e2e.py`

- [ ] **Step 1: フルスペックのE2Eテストを書く**

`tests/test_e2e.py`:
```python
"""全レイアウト・全コンポーネントを使用したE2Eテスト"""
import os
import tempfile
import pytest

from src.generator import generate_pptx


FULL_CONFIG = {
    "theme": "monotone",
    "title": "E2Eテスト資料",
    "slides": [
        {
            "layout": "cover",
            "title": "DX推進戦略提案書",
            "subtitle": "2026年度計画",
            "client": "株式会社テスト",
            "date": "2026年4月",
        },
        {
            "layout": "agenda",
            "items": ["現状分析", "課題整理", "戦略提案", "実行計画"],
        },
        {
            "layout": "section_divider",
            "section_number": 1,
            "section_title": "現状分析",
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "売上は3年で2.2倍に成長",
            "components": [
                {"type": "bullets", "items": [
                    "2023年: 100億円",
                    "2024年: 150億円（前年比+50%）",
                    "2025年: 220億円（前年比+47%）",
                ]},
                {"type": "callout", "text": "CAGR 48%の高成長を実現"},
            ],
        },
        {
            "layout": "chart_page",
            "title": "市場規模は年平均15%で拡大",
            "chart": {
                "type": "bar",
                "data": {
                    "labels": ["2023", "2024", "2025"],
                    "series": [{"name": "市場規模", "values": [100, 150, 220]}],
                },
                "unit": "億円",
            },
            "key_points": ["CAGR 15%で成長", "2025年に200億円突破"],
        },
        {
            "layout": "chart_page",
            "title": "四半期別売上推移",
            "chart": {
                "type": "line",
                "data": {
                    "labels": ["Q1", "Q2", "Q3", "Q4"],
                    "series": [
                        {"name": "2024", "values": [30, 40, 35, 45]},
                        {"name": "2025", "values": [50, 55, 60, 55]},
                    ],
                },
            },
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "売上構成比",
            "components": [
                {"type": "kpi_cards", "cards": [
                    {"value": "220", "unit": "億円", "label": "総売上"},
                    {"value": "48", "unit": "%", "label": "CAGR"},
                    {"value": "3.2", "unit": "万人", "label": "顧客数"},
                ]},
            ],
        },
        {
            "layout": "section_divider",
            "section_number": 2,
            "section_title": "課題整理",
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "DX成熟度の現状評価",
            "components": [
                {
                    "type": "matrix_2x2",
                    "x_axis": "実行難易度",
                    "y_axis": "事業インパクト",
                    "quadrants": ["Quick Win", "戦略投資", "要検討", "後回し"],
                },
            ],
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "課題一覧",
            "components": [
                {
                    "type": "table",
                    "headers": ["課題", "重要度", "緊急度", "担当"],
                    "rows": [
                        ["レガシーシステム刷新", "高", "高", "IT部門"],
                        ["データ基盤構築", "高", "中", "DX推進室"],
                        ["人材育成", "中", "中", "人事部"],
                    ],
                },
            ],
        },
        {
            "layout": "section_divider",
            "section_number": 3,
            "section_title": "戦略提案",
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "3段階のDX推進アプローチ",
            "components": [
                {
                    "type": "pyramid",
                    "levels": ["デジタル変革", "デジタル最適化", "デジタイゼーション"],
                },
            ],
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "推進プロセス",
            "components": [
                {
                    "type": "process_flow",
                    "steps": ["現状把握", "計画策定", "パイロット", "全社展開"],
                },
            ],
        },
        {
            "layout": "comparison",
            "title": "Before / After",
            "left_title": "As-Is",
            "left_components": [
                {"type": "bullets", "items": ["手動オペレーション", "サイロ化されたデータ", "属人的な意思決定"]},
            ],
            "right_title": "To-Be",
            "right_components": [
                {"type": "bullets", "items": ["自動化されたプロセス", "統合データ基盤", "データドリブン経営"]},
            ],
        },
        {
            "layout": "section_divider",
            "section_number": 4,
            "section_title": "実行計画",
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "マイルストーン",
            "components": [
                {
                    "type": "timeline",
                    "milestones": [
                        {"date": "2026/Q1", "label": "キックオフ"},
                        {"date": "2026/Q2", "label": "要件定義完了"},
                        {"date": "2026/Q3", "label": "パイロット開始"},
                        {"date": "2027/Q1", "label": "全社展開"},
                    ],
                },
            ],
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "プロジェクトスケジュール",
            "components": [
                {
                    "type": "gantt",
                    "tasks": [
                        {"name": "要件定義", "start": 0, "duration": 2},
                        {"name": "設計", "start": 1, "duration": 3},
                        {"name": "開発", "start": 3, "duration": 4},
                        {"name": "テスト", "start": 6, "duration": 2},
                    ],
                    "phases": ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月"],
                },
            ],
        },
        {
            "layout": "content",
            "columns": 1,
            "title": "推進体制",
            "components": [
                {
                    "type": "org_chart",
                    "data": {
                        "name": "DX推進委員会",
                        "children": [
                            {"name": "IT部門"},
                            {"name": "DX推進室"},
                            {"name": "事業部門"},
                        ],
                    },
                },
            ],
        },
        {
            "layout": "closing",
            "summary": [
                "DX推進によりCGAR 48%の成長を加速",
                "3段階アプローチで着実に推進",
                "2027年Q1の全社展開を目指す",
            ],
            "next_steps": [
                "ステアリングコミッティの設置（4月中）",
                "パイロット対象部門の選定（5月中）",
                "予算承認プロセスの開始",
            ],
        },
    ],
}


class TestE2EGeneration:
    def test_full_presentation_monotone(self):
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(FULL_CONFIG, output_path)
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 0
            assert len(prs.slides) == len(FULL_CONFIG["slides"])
        finally:
            os.unlink(output_path)

    def test_full_presentation_dark(self):
        config = {**FULL_CONFIG, "theme": "dark"}
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(config, output_path)
            assert os.path.exists(output_path)
            assert len(prs.slides) == len(config["slides"])
        finally:
            os.unlink(output_path)

    def test_full_presentation_colorful(self):
        config = {**FULL_CONFIG, "theme": "colorful"}
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        try:
            prs = generate_pptx(config, output_path)
            assert os.path.exists(output_path)
            assert len(prs.slides) == len(config["slides"])
        finally:
            os.unlink(output_path)
```

- [ ] **Step 2: テスト通過を確認**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/test_e2e.py -v`
Expected: ALL PASS

- [ ] **Step 3: 全テスト実行**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/ -v`
Expected: ALL PASS

- [ ] **Step 4: コミット**

```bash
git add tests/test_e2e.py
git commit -m "feat: end-to-end tests covering all layouts, components, and themes"
```

---

### Task 14: スキル定義（skill.md）

**Files:**
- Create: `skill.md`

- [ ] **Step 1: skill.mdを作成**

`skill.md`:
````markdown
# PowerPoint資料生成スキル

コンサルティングファームレベルのPowerPoint資料を生成するスキル。

## 起動条件

ユーザーがPowerPoint、スライド、プレゼン資料、提案書の作成を依頼した場合に起動する。

## 処理フロー

### フェーズ1: ヒアリング

以下を簡潔に確認する（1-2回の質問で完了。ユーザーが十分な情報を提供済みなら質問をスキップ）:

1. **資料の目的**: 提案、報告、計画、分析、プレゼン等
2. **対象読者**: 経営層、現場担当者、クライアント等
3. **盛り込みたい内容**: キーメッセージ、含めたいデータや論点
4. **テーマ選択**: 以下から選択
   - `monotone`: 白背景 + 濃紺テキスト + 赤アクセント（端正・フォーマル）
   - `dark`: 濃紺背景 + 白テキスト + 橙アクセント（重厚・インパクト）
   - `colorful`: 白背景 + 青/緑/橙の3色（モダン・カジュアル）

### フェーズ2: 構成提案

スライド構成をMarkdown形式で提示する。

**構成設計の原則:**

- ストーリーラインを重視する。「状況 → 課題 → 原因 → 解決策 → 効果 → 実行計画」等の論理構造で組み立てる
- 1スライド1メッセージ原則を守る
- タイトルはアクションタイトルにする（結論を述べる形）
  - 良い例: 「売上は3年で2.2倍に成長」「レガシー刷新が最優先課題」
  - 悪い例: 「売上推移」「課題一覧」
- 箇条書きは3-5項目、各項目1-2行以内
- 絵文字は一切使わない
- データがない場合はプレースホルダー値を入れ「※サンプル値」と注記する

**構成の提示例:**

```
1. 表紙: 「DX推進戦略提案書」
2. アジェンダ: 4項目
3. セクション区切り: 1. 現状分析
4. コンテンツ: 「売上は3年で2.2倍に成長」- 箇条書き + 強調ボックス
5. チャートページ: 「市場規模は年平均15%で拡大」- 棒グラフ + キーポイント
6. セクション区切り: 2. 戦略提案
7. コンテンツ: 「3段階アプローチで推進」- プロセスフロー
8. 比較: 「Before / After」- As-Is vs To-Be
9. まとめ + Next Steps
```

ユーザーに承認を求める。修正があれば反映する。

### フェーズ3: 生成

承認後、以下の手順で生成する:

1. 構成をJSON形式に変換する
2. 以下のPythonコードを生成・実行する:

```python
import json
import sys
sys.path.insert(0, "/Users/kenichi/Desktop/project/ppt_skills")
from src.generator import generate_pptx

config = {構成JSON}

output_path = "{資料タイトル}.pptx"
generate_pptx(config, output_path)
print(f"生成完了: {output_path}")
```

3. 生成完了を報告し、ファイルパスを伝える

## 利用可能なレイアウト

| layout名 | 用途 | 主なパラメータ |
|---|---|---|
| `cover` | 表紙 | title, subtitle, client, date |
| `agenda` | アジェンダ | items, highlight(optional) |
| `section_divider` | セクション区切り | section_number, section_title |
| `content` | 汎用コンテンツ | title, columns(1/2/3), components |
| `chart_page` | チャート主体 | title, chart, key_points(optional) |
| `comparison` | 比較 | title, left_title, left_components, right_title, right_components |
| `closing` | まとめ/Thank You | summary, next_steps / type:"thank_you", contact |

## 利用可能なコンポーネント（contentレイアウト内）

| type | 用途 | パラメータ |
|---|---|---|
| `bullets` | 箇条書き | items: list[str or {text, level}] |
| `callout` | 強調ボックス | text: str |
| `table` | テーブル | headers: list[str], rows: list[list[str]] |
| `matrix_2x2` | 2x2マトリクス | x_axis, y_axis, quadrants: list[4] |
| `pyramid` | ピラミッド | levels: list[str] |
| `process_flow` | プロセスフロー | steps: list[str] |
| `cycle` | サイクル図 | items: list[str] |
| `org_chart` | 組織図 | data: {name, children: [{name, children}]} |
| `timeline` | タイムライン | milestones: [{date, label}] |
| `gantt` | ガント | tasks: [{name, start, duration}], phases: list[str] |
| `icon_row` | アイコン行 | items: [{icon, label}] |
| `kpi_cards` | KPIカード群 | cards: [{value, unit, label}] |

## チャートタイプ（chart_pageレイアウト内）

| type | 用途 | dataフォーマット |
|---|---|---|
| `bar` | 棒グラフ | {labels, series: [{name, values}]} |
| `line` | 折れ線 | {labels, series: [{name, values}]} |
| `pie` | 円グラフ | {labels, values} |
| `waterfall` | ウォーターフォール | {labels, values} |

## JSON構成例

```json
{
  "theme": "monotone",
  "title": "DX推進戦略提案書",
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
        {"type": "bullets", "items": ["2023年: 100億円", "2024年: 150億円", "2025年: 220億円"]},
        {"type": "callout", "text": "CAGR 48%の高成長を実現"}
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
      "layout": "closing",
      "summary": ["要点1", "要点2", "要点3"],
      "next_steps": ["ステップ1", "ステップ2", "ステップ3"]
    }
  ]
}
```
````

- [ ] **Step 2: コミット**

```bash
git add skill.md
git commit -m "feat: skill.md - Claude Code skill definition for PPT generation"
```

---

### Task 15: 全テスト実行 + 最終確認

- [ ] **Step 1: 全テスト実行**

Run: `cd /Users/kenichi/Desktop/project/ppt_skills && python -m pytest tests/ -v --tb=short`
Expected: ALL PASS

- [ ] **Step 2: サンプル資料を生成して動作確認**

Run:
```bash
cd /Users/kenichi/Desktop/project/ppt_skills && python -c "
import json, sys
sys.path.insert(0, '.')
from src.generator import generate_pptx

config = {
    'theme': 'monotone',
    'title': 'サンプル資料',
    'slides': [
        {'layout': 'cover', 'title': 'サンプル提案書', 'subtitle': '動作確認用', 'date': '2026年4月'},
        {'layout': 'agenda', 'items': ['概要', '分析', '提案']},
        {'layout': 'section_divider', 'section_number': 1, 'section_title': '概要'},
        {'layout': 'content', 'columns': 1, 'title': 'テストスライド', 'components': [{'type': 'bullets', 'items': ['項目1', '項目2', '項目3']}]},
        {'layout': 'closing', 'summary': ['まとめ1', 'まとめ2'], 'next_steps': ['次のステップ1']},
    ],
}
generate_pptx(config, 'sample_output.pptx')
print('Generated: sample_output.pptx')
"
```
Expected: `Generated: sample_output.pptx` — ファイルをPowerPointで開いて見た目を確認

- [ ] **Step 3: sample_output.pptxを削除してコミット**

```bash
rm -f sample_output.pptx
echo "*.pptx" >> .gitignore
git add .gitignore
git commit -m "chore: add .gitignore for generated pptx files"
```
