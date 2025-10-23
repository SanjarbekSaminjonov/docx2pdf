"""Aggregate model combining styles, structure, and layout."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from docx_renderer.model.elements import LayoutModel
from docx_renderer.model.style_model import StylesCatalog


@dataclass(slots=True)
class DocumentModel:
    """Flattened document representation that renderers consume."""

    styles: StylesCatalog
    layout: LayoutModel
    metadata: Optional[dict] = None
