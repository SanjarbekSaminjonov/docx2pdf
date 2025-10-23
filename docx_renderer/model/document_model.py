"""Aggregate model combining styles, structure, and layout."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, TYPE_CHECKING

from docx_renderer.model.elements import LayoutModel
from docx_renderer.model.numbering_model import NumberingCatalog
from docx_renderer.model.style_model import StylesCatalog

if TYPE_CHECKING:
    from docx_renderer.parser.media_extractor import MediaCatalog


@dataclass(slots=True)
class DocumentModel:
    """Flattened document representation that renderers consume."""

    styles: StylesCatalog
    layout: LayoutModel
    numbering: Optional[NumberingCatalog] = None
    media: Optional["MediaCatalog"] = None
    metadata: Optional[dict] = None
