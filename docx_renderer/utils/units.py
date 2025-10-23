"""Unit conversion helpers for WordprocessingML measurements."""
from __future__ import annotations

EMU_PER_INCH = 914400
TWIPS_PER_POINT = 20
POINTS_PER_INCH = 72


def emu_to_points(value: int) -> float:
    """Convert English Metric Units to typographic points."""
    return (value / EMU_PER_INCH) * POINTS_PER_INCH


def points_to_twips(value: float) -> int:
    """Convert points to twips (1/20th of a point)."""
    return int(round(value * TWIPS_PER_POINT))


def twips_to_points(value: int) -> float:
    """Convert twips to points."""
    return value / TWIPS_PER_POINT
