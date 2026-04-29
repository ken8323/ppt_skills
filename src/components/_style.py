"""テキスト書式の共通ヘルパ。paragraph 内の全 run に対して書式を一括適用する。
python-pptx で `paragraph.text = "a\nb"` のように改行を含むテキストを入れると
複数の run が生成されるため、`runs[0]` のみに書式を当てるバグを防ぐ目的で集約する。"""

from pptx.util import Pt
from pptx.dml.color import RGBColor


def style_runs(paragraph, *, size=None, color=None, name=None, bold=None, italic=None):
    """paragraph 配下の全 run に font 設定を適用。None はスキップ。"""
    for run in paragraph.runs:
        font = run.font
        if size is not None:
            font.size = size
        if color is not None:
            font.color.rgb = color
        if name is not None:
            font.name = name
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic


def set_paragraph_text(paragraph, text, *, size=None, color=None, name=None, bold=None, italic=None):
    """paragraph にテキストを設定し、生成された全 run に書式を適用する。"""
    paragraph.text = text
    style_runs(paragraph, size=size, color=color, name=name, bold=bold, italic=italic)
