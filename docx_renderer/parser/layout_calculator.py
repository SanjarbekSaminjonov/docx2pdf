"""Convert structured document blocks into absolutely positioned layout boxes."""
from __future__ import annotations

from dataclasses import dataclass
from typing import List

from docx_renderer.model.elements import DocumentTree, LayoutBox, LayoutModel, ParagraphElement
from docx_renderer.model.style_model import StylesCatalog


@dataclass(slots=True)
class LayoutContext:
    """Mutable state while computing layout flow."""

    cursor_x: float = 0.0
    cursor_y: float = 0.0
    line_height: float = 14.0
    page_width: float = 595.0  # Default A4 width in points
    page_height: float = 842.0  # Default A4 height in points
    margin_left: float = 72.0
    margin_top: float = 72.0


class LayoutCalculator:
    """Transform document structure into renderer-friendly layout."""

    def __init__(self, styles: StylesCatalog) -> None:
        self._styles = styles

    def calculate(self, tree: DocumentTree) -> LayoutModel:
        context = LayoutContext()
        boxes: List[LayoutBox] = []
        for block in tree.blocks:
            if isinstance(block, ParagraphElement):
                boxes.append(self._layout_paragraph(block, context))
            else:
                boxes.append(self._layout_placeholder(block, context))
        return LayoutModel(boxes=boxes, pages=[boxes])

    def _layout_paragraph(self, paragraph: ParagraphElement, context: LayoutContext) -> LayoutBox:
        text_content = "".join(run.text for run in paragraph.runs)
        width = context.page_width - context.margin_left * 2
        height = context.line_height
        box = LayoutBox(
            element_type="paragraph",
            content={"text": text_content},
            x=context.margin_left,
            y=context.margin_top + context.cursor_y,
            width=width,
            height=height,
            style={"styleId": paragraph.style_id},
        )
        context.cursor_y += height
        return box

    def _layout_placeholder(self, block, context: LayoutContext) -> LayoutBox:
        width = context.page_width - context.margin_left * 2
        height = context.line_height
        box = LayoutBox(
            element_type="unsupported",
            content={"repr": repr(block)},
            x=context.margin_left,
            y=context.margin_top + context.cursor_y,
            width=width,
            height=height,
            style={},
        )
        context.cursor_y += height
        return box
