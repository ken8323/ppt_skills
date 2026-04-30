"""examples/ 配下の全 JSON が validate + generate できることを担保する。"""
import json
from pathlib import Path

import pytest

from src.generator import generate_pptx
from src.linter import lint_config
from src.validator import validate_config

EXAMPLES_DIR = Path(__file__).parent.parent / "examples"
EXAMPLE_FILES = sorted(EXAMPLES_DIR.glob("*.json"))


@pytest.mark.parametrize("path", EXAMPLE_FILES, ids=lambda p: p.stem)
def test_example_validates_and_generates(path, tmp_path):
    config = json.loads(path.read_text(encoding="utf-8"))
    config.pop("$comment", None)
    validate_config(config)
    assert lint_config(config) == [], f"lint warnings in {path.name}"
    out = tmp_path / f"{path.stem}.pptx"
    generate_pptx(config, str(out), lint=False)
    assert out.exists()
