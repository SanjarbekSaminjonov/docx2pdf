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
class SectionProperties:
    """Section-level configuration including headers, footers, and page setup."""

    page_width: Optional[int] = None
    page_height: Optional[int] = None
    margin_top: Optional[int] = None
    margin_bottom: Optional[int] = None
    margin_left: Optional[int] = None
    margin_right: Optional[int] = None
    margin_header: Optional[int] = None
    margin_footer: Optional[int] = None
    orientation: Optional[str] = None
    header_default: Optional["HeaderFooterContent"] = None
    header_first: Optional["HeaderFooterContent"] = None
    header_even: Optional["HeaderFooterContent"] = None
    footer_default: Optional["HeaderFooterContent"] = None
    footer_first: Optional["HeaderFooterContent"] = None
    footer_even: Optional["HeaderFooterContent"] = None
    title_page: bool = False
    raw_properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class HeaderFooterContent:
    """Content of a header or footer."""

    r_id: str
    blocks: List[BlockElement] = field(default_factory=list)
    raw_xml: Optional[str] = None


@dataclass(slots=True)
class DocumentSection:
    """Document section with content and section properties."""

    blocks: List[BlockElement]
    properties: SectionProperties


@dataclass(slots=True)
class DocumentTree:
    """Structured representation prior to layout calculation."""

    sections: List[DocumentSection] = field(default_factory=list)
    metadata: Dict[str, object] = field(default_factory=dict)

    @property
    def blocks(self) -> List[BlockElement]:
        """Backward compatibility: return all blocks from all sections."""
        all_blocks = []
        for section in self.sections:
            all_blocks.extend(section.blocks)
        return all_blocks


@dataclass(slots=True)
class MediaAsset:
    """Represents a media asset (image, font, etc.) from DOCX package."""

    relationship_id: str
    target_path: str
    media_type: str
    binary_data: bytes
    base64_data: str
    size: int
    metadata: Dict[str, object] = field(default_factory=dict)
