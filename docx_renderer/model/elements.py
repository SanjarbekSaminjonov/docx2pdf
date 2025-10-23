"""In-memory representation of parsed document content and layout."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence


@dataclass(slots=True)
class RunFragment:
    """Represents a contiguous run of text with associated inline styling."""

    text: str
    style_id: Optional[str] = None
    properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class ParagraphElement:
    """High-level block element for paragraphs in the document body."""

    runs: List[RunFragment]
    style_id: Optional[str]
    properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class TableCell:
    """Single table cell container."""

    content: List["BlockElement"]
    properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class TableRow:
    """Row with a sequence of cells."""

    cells: List[TableCell]
    properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class TableElement:
    """Tabular structure extracted from Word tables."""

    rows: List[TableRow]
    style_id: Optional[str]
    properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class ImageElement:
    """An inline or anchored image reference."""

    r_id: str
    media_path: str
    width_emu: Optional[int] = None
    height_emu: Optional[int] = None
    properties: Dict[str, object] = field(default_factory=dict)


BlockElement = ParagraphElement | TableElement | ImageElement


@dataclass(slots=True)
class LayoutBox:
    """Absolute positioned box to be consumed by renderers."""

    element_type: str
    content: Dict[str, object]
    x: float
    y: float
    width: float
    height: float
    style: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class LayoutModel:
    """Ordered list of layout boxes with optional page segmentation."""

    boxes: Sequence[LayoutBox]
    pages: Sequence[Sequence[LayoutBox]] = field(default_factory=list)


@dataclass(slots=True)
class DocumentTree:
    """Structured representation prior to layout calculation."""

    blocks: List[BlockElement]
    metadata: Dict[str, object] = field(default_factory=dict)
