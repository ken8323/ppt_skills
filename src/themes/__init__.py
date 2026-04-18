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
