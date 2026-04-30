"""ビジュアルリグレッションテスト。

初回実行: tests/golden/ にゴールデンPNGを生成 (--update-golden フラグ or 存在しない場合)。
以降の実行: 生成した PNG との差分ピクセル率が THRESHOLD 以内であることを確認する。

実行:
    pytest tests/test_visual_regression.py                  # 比較モード
    pytest tests/test_visual_regression.py --update-golden  # ゴールデン更新
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path

import pytest

GOLDEN_DIR = Path(__file__).parent / "golden"
DIFF_THRESHOLD = 0.02  # 2% 以内の差分を許容

try:
    from PIL import Image, ImageChops
    import numpy as np
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


def _pixel_diff_ratio(img_a: "Image.Image", img_b: "Image.Image") -> float:
    """2枚の画像の差分ピクセル率 (0.0–1.0) を返す。"""
    if img_a.size != img_b.size:
        img_b = img_b.resize(img_a.size, Image.LANCZOS)
    diff = ImageChops.difference(img_a.convert("RGB"), img_b.convert("RGB"))
    arr = np.array(diff)
    changed = np.any(arr > 10, axis=2)  # チャネル方向で 10 以上の差
    return float(changed.sum()) / changed.size


SNAPSHOT_CASES = [
    ("cover_monotone", {
        "theme": "monotone",
        "slides": [{"layout": "cover", "data": {"title": "Visual Regression", "subtitle": "Test"}}],
    }),
    ("content_colorful", {
        "theme": "colorful",
        "slides": [{
            "layout": "content",
            "data": {
                "title": "コンテンツスライド",
                "columns": 1,
                "components": [
                    {"type": "bullets", "items": ["項目A", "項目B", "項目C"]},
                ],
            },
        }],
    }),
    ("chart_dark", {
        "theme": "dark",
        "slides": [{
            "layout": "chart_page",
            "data": {
                "title": "売上推移",
                "source": "社内データ 2025年",
                "chart": {
                    "type": "bar",
                    "title": "月次売上",
                    "data": {
                        "labels": ["1月", "2月", "3月"],
                        "series": [{"name": "売上", "values": [100, 120, 90]}],
                    },
                },
            },
        }],
    }),
]


@pytest.mark.skipif(not PIL_AVAILABLE, reason="Pillow + numpy が必要")
@pytest.mark.parametrize("name,config", SNAPSHOT_CASES)
def test_visual_snapshot(name, config, tmp_path, request):
    from src.generator import generate_pptx
    from src.thumbnail import generate_thumbnails

    pptx_path = tmp_path / f"{name}.pptx"
    generate_pptx(config, str(pptx_path), lint=False)

    thumb_dir = tmp_path / "thumbs"
    thumbs = generate_thumbnails(pptx_path, thumb_dir)
    assert thumbs, f"{name}: サムネイルが生成されませんでした"

    actual = Image.open(thumbs[0]).convert("RGB")
    golden_path = GOLDEN_DIR / f"{name}.png"

    update = request.config.getoption("--update-golden", default=False)

    if update or not golden_path.exists():
        GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
        actual.save(golden_path)
        pytest.skip(f"ゴールデン画像を更新しました: {golden_path}")
        return

    golden = Image.open(golden_path).convert("RGB")
    ratio = _pixel_diff_ratio(actual, golden)
    assert ratio <= DIFF_THRESHOLD, (
        f"{name}: 差分ピクセル率 {ratio:.1%} が閾値 {DIFF_THRESHOLD:.1%} を超えています。"
        "意図的な変更なら --update-golden で更新してください。"
    )
