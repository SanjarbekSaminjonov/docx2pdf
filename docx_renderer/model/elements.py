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
    controls: List[Dict[str, object]] = field(default_factory=list)
    drawing: Optional["DrawingReference"] = None
    footnote_reference: Optional[int] = None
    endnote_reference: Optional[int] = None
    field_code: Optional[str] = None
    hyperlink_id: Optional[str] = None
    hyperlink_anchor: Optional[str] = None
    hyperlink_target: Optional[str] = None


@dataclass(slots=True)
class ParagraphElement:
    """High-level block element for paragraphs in the document body."""

    runs: List[RunFragment]
    style_id: Optional[str]
    properties: Dict[str, object] = field(default_factory=dict)
    numbering: Optional["NumberingInfo"] = None
    bookmarks: List["Bookmark"] = field(default_factory=list)
    annotations: Dict[str, object] = field(default_factory=dict)


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
    data: Optional[bytes] = None


@dataclass(slots=True)
class DrawingReference:
    """Represents a drawing (usually an image) embedded inside a run."""

    r_id: str
    target: Optional[str]
    description: Optional[str]
    width_emu: Optional[int]
    height_emu: Optional[int]
    inline: bool
    data: Optional[bytes] = None


@dataclass(slots=True)
class Bookmark:
    """Bookmark start marker embedded within a paragraph."""

    bookmark_id: int
    name: str


@dataclass(slots=True)
class NumberingInfo:
    """Resolved numbering reference applied to a paragraph."""

    num_id: int
    level: int
    abstract_num_id: Optional[int] = None
    start: Optional[int] = None
    format: Optional[str] = None
    level_text: Optional[str] = None
    alignment: Optional[str] = None


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
