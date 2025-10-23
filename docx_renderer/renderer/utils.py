"""Common helpers shared by renderer implementations."""
from __future__ import annotations

from typing import Dict


def style_to_css(style: Dict[str, object]) -> Dict[str, str]:
    """Convert style dictionary into CSS properties."""
    css: Dict[str, str] = {}
    if style.get("bold"):
        css["font-weight"] = "700"
    if style.get("italic"):
        css["font-style"] = "italic"
    return css
