"""Convert structured document blocks into absolutely positioned layout boxes."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence

from docx_renderer.model.elements import (
    DocumentSection,
    DocumentTree,
    LayoutBox,
    LayoutModel,
    HeaderFooterContent,
    ParagraphElement,
    SectionProperties,
    TableElement,
    ImageElement,
    TableCell,
)
from docx_renderer.model.style_model import StyleDefinition, StylesCatalog


DEFAULT_PAGE_WIDTH_PT = 595.0  # ≈ 8.27" A4
DEFAULT_PAGE_HEIGHT_PT = 842.0  # ≈ 11.69" A4
DEFAULT_MARGIN_PT = 72.0  # 1 inch
DEFAULT_FONT_SIZE_PT = 11.0
DEFAULT_LINE_HEIGHT_PT = DEFAULT_FONT_SIZE_PT * 1.2
DEFAULT_TABLE_CELL_PADDING_PT = 4.0
DEFAULT_TABLE_BORDER_WIDTH_PT = 0.5
EMU_PER_POINT = 12700.0


@dataclass(slots=True)
class LayoutContext:
    """Mutable state while computing layout flow for a section."""

    page_width: float
    page_height: float
    margin_left: float
    margin_right: float
    margin_top: float
    margin_bottom: float
    cursor_x: float
    cursor_y: float
    header_margin: float = 0.0
    footer_margin: float = 0.0

    @property
    def available_width(self) -> float:
        """Content width within the current section, excluding margins."""

        return self.page_width - self.margin_left - self.margin_right


@dataclass(slots=True)
class ParagraphIndent:
    """Aggregated paragraph indentation values in points."""

    left: float = 0.0
    right: float = 0.0
    first_line: float = 0.0


@dataclass(slots=True)
class SpacingInfo:
    """Paragraph spacing configuration in points."""

    before: float = 0.0
    after: float = 0.0
    line: Optional[float] = None
    line_rule: Optional[str] = None


class LayoutCalculator:
    """Transform document structure into renderer-friendly layout."""

    def __init__(self, styles: StylesCatalog) -> None:
        self._styles = styles

    # ------------------------------------------------------------------
    # Public API
    def calculate(self, tree: DocumentTree) -> LayoutModel:
        """Return a LayoutModel with per-section pagination information."""

        boxes: List[LayoutBox] = []
        pages: List[Sequence[LayoutBox]] = []

        for section in self._iter_sections(tree):
            context = self._build_context(section.properties)
            section_boxes: List[LayoutBox] = []

            header_content = self._select_header_content(section.properties)
            if header_content:
                header_box = self._layout_header_footer(header_content, context, placement="header")
                if header_box:
                    boxes.append(header_box)
                    section_boxes.append(header_box)

            for block in section.blocks:
                box = self._layout_block(block, context)
                boxes.append(box)
                section_boxes.append(box)

            footer_content = self._select_footer_content(section.properties)
            if footer_content:
                footer_box = self._layout_header_footer(footer_content, context, placement="footer")
                if footer_box:
                    boxes.append(footer_box)
                    section_boxes.append(footer_box)

            pages.append(section_boxes)

        return LayoutModel(boxes=boxes, pages=pages)

    def _select_header_content(self, properties: SectionProperties | None) -> Optional[HeaderFooterContent]:
        if not isinstance(properties, SectionProperties):
            return None
        return properties.header_default or properties.header_first or properties.header_even

    def _select_footer_content(self, properties: SectionProperties | None) -> Optional[HeaderFooterContent]:
        if not isinstance(properties, SectionProperties):
            return None
        return properties.footer_default or properties.footer_first or properties.footer_even

    def _layout_header_footer(
        self,
        content: HeaderFooterContent,
        base_context: LayoutContext,
        placement: str,
    ) -> Optional[LayoutBox]:
        if not content.blocks:
            return None

        if placement not in {"header", "footer"}:
            return None

        if placement == "header":
            start = content_start = self._resolve_header_start(base_context)
        else:
            target_margin = base_context.footer_margin if base_context.footer_margin > 0 else base_context.margin_bottom
            start = content_start = max(base_context.page_height - target_margin, 0.0)

        working_context = LayoutContext(
            page_width=base_context.page_width,
            page_height=base_context.page_height,
            margin_left=base_context.margin_left,
            margin_right=base_context.margin_right,
            margin_top=base_context.margin_top,
            margin_bottom=base_context.margin_bottom,
            cursor_x=base_context.margin_left,
            cursor_y=content_start,
            header_margin=base_context.header_margin,
            footer_margin=base_context.footer_margin,
        )

        child_boxes: List[LayoutBox] = []
        for block in content.blocks:
            child_boxes.append(self._layout_block(block, working_context))

        total_height = max(working_context.cursor_y - content_start, 0.0)

        if placement == "header":
            container_y = self._clamp_header_top(base_context, start, total_height)
            offset = container_y - start
        else:
            container_y, offset = self._resolve_footer_position(base_context, start, total_height)

        if offset:
            for box in child_boxes:
                box.y += offset

        container = LayoutBox(
            element_type=placement,
            content={
                "boxes": child_boxes,
                "relationshipId": content.r_id,
                "rawXml": content.raw_xml,
            },
            x=base_context.margin_left,
            y=container_y,
            width=base_context.available_width,
            height=total_height,
            style={
                "type": placement,
                "rId": content.r_id,
            },
        )

        return container

    def _resolve_header_start(self, context: LayoutContext) -> float:
        default_start = context.margin_top * 0.5 if context.margin_top > 0 else 0.0
        desired = context.header_margin if context.header_margin > 0 else default_start
        max_within_margin = max(context.margin_top - DEFAULT_LINE_HEIGHT_PT, 0.0)
        if max_within_margin > 0:
            desired = min(desired, max_within_margin)
        return max(desired, 0.0)

    def _clamp_header_top(self, context: LayoutContext, start: float, height: float) -> float:
        top_limit = max(context.margin_top - height, 0.0)
        return min(start, top_limit) if top_limit > 0 else max(start, 0.0)

    def _resolve_footer_position(self, context: LayoutContext, start: float, height: float) -> tuple[float, float]:
        target_margin = context.footer_margin if context.footer_margin > 0 else context.margin_bottom
        usable_bottom = context.page_height - target_margin
        desired_top = usable_bottom - height
        min_top = context.margin_top
        max_top = context.page_height - context.margin_bottom
        if desired_top < min_top:
            desired_top = min_top
        elif desired_top > max_top:
            desired_top = max_top
        offset = desired_top - start
        return desired_top, offset

    # ------------------------------------------------------------------
    # Section helpers
    def _iter_sections(self, tree: DocumentTree) -> Iterable[DocumentSection]:
        """Yield sections, falling back to a synthetic section when needed."""

        if tree.sections:
            return tree.sections

        # Fallback for legacy trees without explicit sections
        return [DocumentSection(blocks=list(tree.blocks), properties=SectionProperties())]

    def _build_context(self, properties: SectionProperties | None) -> LayoutContext:
        """Create a layout context using section-specific page setup."""

        page_width = self._twips_to_points(getattr(properties, "page_width", None), DEFAULT_PAGE_WIDTH_PT)
        page_height = self._twips_to_points(getattr(properties, "page_height", None), DEFAULT_PAGE_HEIGHT_PT)

        orientation = getattr(properties, "orientation", None)
        if orientation and orientation.lower() == "landscape" and page_width < page_height:
            page_width, page_height = page_height, page_width

        margin_left = self._twips_to_points(getattr(properties, "margin_left", None), DEFAULT_MARGIN_PT)
        margin_right = self._twips_to_points(getattr(properties, "margin_right", None), DEFAULT_MARGIN_PT)
        margin_top = self._twips_to_points(getattr(properties, "margin_top", None), DEFAULT_MARGIN_PT)
        margin_bottom = self._twips_to_points(getattr(properties, "margin_bottom", None), DEFAULT_MARGIN_PT)

        header_margin = self._twips_to_points(getattr(properties, "margin_header", None), 0.0)
        footer_margin = self._twips_to_points(getattr(properties, "margin_footer", None), 0.0)

        return LayoutContext(
            page_width=page_width,
            page_height=page_height,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            cursor_x=margin_left,
            cursor_y=margin_top,
            header_margin=header_margin,
            footer_margin=footer_margin,
        )

    # ------------------------------------------------------------------
    # Block layout
    def _layout_block(self, block, context: LayoutContext) -> LayoutBox:
        if isinstance(block, ParagraphElement):
            return self._layout_paragraph(block, context)
        if isinstance(block, TableElement):
            return self._layout_table(block, context)
        if isinstance(block, ImageElement):
            return self._layout_image(block, context)

        return self._layout_placeholder(block, context)

    def _layout_paragraph(self, paragraph: ParagraphElement, context: LayoutContext) -> LayoutBox:
        text_content = "".join((run.text or "") for run in paragraph.runs)
        resolved_style = self._styles.get(paragraph.style_id)

        font_size = self._resolve_font_size(paragraph, resolved_style)
        spacing = self._resolve_spacing(paragraph, resolved_style)
        line_height = self._resolve_line_height(font_size, spacing)
        indent = self._resolve_paragraph_indent(paragraph, resolved_style)

        if spacing.before:
            context.cursor_y += spacing.before

        available_width = max(context.available_width - indent.left - indent.right, 0.0)
        lines = self._wrap_text(text_content, available_width, font_size)
        box_height = line_height * max(len(lines), 1)

        style: Dict[str, object] = {"styleId": paragraph.style_id}
        if isinstance(resolved_style, StyleDefinition):
            style["resolved"] = resolved_style.properties

        box = LayoutBox(
            element_type="paragraph",
            content={
                "text": text_content,
                "lines": lines,
                "firstLineIndent": indent.first_line,
            },
            x=context.margin_left + indent.left,
            y=context.cursor_y,
            width=available_width,
            height=box_height,
            style=style,
        )

        box.style["spacing"] = {
            "before": spacing.before,
            "after": spacing.after,
            "line": spacing.line,
            "lineRule": spacing.line_rule,
        }
        box.style["indent"] = {
            "left": indent.left,
            "right": indent.right,
            "firstLine": indent.first_line,
        }

        context.cursor_y += box_height + spacing.after
        return box

    def _layout_table(self, table: TableElement, context: LayoutContext) -> LayoutBox:
        available_width = context.available_width
        column_count = self._compute_table_column_count(table)
        min_widths = self._compute_table_min_widths(table, column_count, available_width)
        column_widths, table_width = self._resolve_column_widths(table, available_width, column_count, min_widths)

        cells_layout: List[List[Dict[str, object]]] = []
        row_heights: List[float] = []
        y_offset = 0.0
        table_x = context.margin_left
        table_y = context.cursor_y

        for row_idx, row in enumerate(table.rows):
            row_cells: List[Dict[str, object]] = []
            row_height = 0.0
            column_index = 0

            for cell in row.cells:
                span = self._resolve_grid_span(cell.properties)
                span = max(1, min(span, column_count - column_index))
                cell_width = sum(column_widths[column_index : column_index + span])
                merge_info = self._resolve_vertical_merge(cell.properties)

                cell_layout = self._layout_table_cell(
                    cell,
                    span,
                    cell_width,
                    context,
                    table.properties,
                )
                cell_layout["columnIndex"] = column_index
                cell_layout["rowIndex"] = row_idx
                if merge_info:
                    cell_layout["vMerge"] = merge_info

                row_height = max(row_height, cell_layout["height"])
                row_cells.append(cell_layout)
                column_index += span

            if column_index < column_count:
                remaining_width = sum(column_widths[column_index:])
                filler = self._create_empty_cell(remaining_width, column_index, column_count - column_index)
                filler["columnIndex"] = column_index
                filler["rowIndex"] = row_idx
                row_height = max(row_height, filler["height"])
                row_cells.append(filler)

            if row_height == 0.0:
                row_height = DEFAULT_LINE_HEIGHT_PT + 2 * (DEFAULT_TABLE_CELL_PADDING_PT + DEFAULT_TABLE_BORDER_WIDTH_PT)

            for cell_layout in row_cells:
                start = cell_layout.get("columnIndex", 0)
                span = cell_layout.get("colSpan", 1)
                width = sum(column_widths[start : start + span]) if column_count else cell_layout["width"]
                cell_layout["width"] = width
                cell_layout["height"] = row_height
                cell_layout["x"] = table_x + sum(column_widths[:start])
                cell_layout["y"] = table_y + y_offset
                cell_layout["baseRow"] = row_idx

            cells_layout.append(row_cells)
            row_heights.append(row_height)
            y_offset += row_height

        self._apply_vertical_merges(cells_layout, row_heights)

        if not row_heights:
            row_heights = [DEFAULT_LINE_HEIGHT_PT + 2 * (DEFAULT_TABLE_CELL_PADDING_PT + DEFAULT_TABLE_BORDER_WIDTH_PT)]
            cells_layout = [[self._create_empty_cell(available_width, 0, 1)]]
            column_widths = [available_width]
            table_width = available_width
            y_offset = row_heights[0]

        table_height = sum(row_heights)

        box = LayoutBox(
            element_type="table",
            content={
                "rows": len(cells_layout),
                "columns": len(column_widths),
                "columnWidths": column_widths,
                "rowHeights": row_heights,
                "cells": cells_layout,
            },
            x=table_x,
            y=table_y,
            width=table_width,
            height=table_height,
            style={"styleId": table.style_id, "properties": table.properties},
        )

        table_spacing = DEFAULT_LINE_HEIGHT_PT * 0.5
        context.cursor_y += table_height + table_spacing
        return box

    def _apply_vertical_merges(
        self,
        cells_layout: List[List[Dict[str, object]]],
        row_heights: List[float],
    ) -> None:
        if not cells_layout or not row_heights:
            return

        active_merges: Dict[tuple[int, int], Dict[str, object]] = {}

        for row_index, row in enumerate(cells_layout):
            for cell in row:
                merge_state = cell.get("vMerge")
                if merge_state not in {"restart", "continue"}:
                    cell["rowSpan"] = cell.get("rowSpan", 1) or 1
                    continue

                column_index = int(cell.get("columnIndex", 0))
                column_span = int(cell.get("colSpan", 1) or 1)
                key = (column_index, column_span)

                row_height = row_heights[row_index] if row_index < len(row_heights) else cell.get("height", 0.0)

                if merge_state == "restart":
                    state = {
                        "cell": cell,
                        "row_span": 1,
                        "height": row_height,
                        "start_row": row_index,
                        "last_row": row_index,
                    }
                    active_merges[key] = state
                    cell["rowSpan"] = 1
                    cell["baseRow"] = row_index
                    continue

                state = active_merges.get(key)
                if not state:
                    fallback_state = {
                        "cell": cell,
                        "row_span": 1,
                        "height": row_height,
                        "start_row": row_index,
                        "last_row": row_index,
                    }
                    active_merges[key] = fallback_state
                    cell["rowSpan"] = 1
                    cell["baseRow"] = row_index
                    continue

                state["row_span"] += 1
                state["height"] += row_height
                state["last_row"] = row_index

                base_cell = state["cell"]
                base_cell["rowSpan"] = state["row_span"]
                base_cell["height"] = state["height"]

                cell["rowSpan"] = 0
                cell["height"] = 0.0
                cell["contentHeight"] = 0.0
                cell["boxes"] = []
                cell["baseRow"] = state["start_row"]

            stale_keys = [
                key
                for key, merge_state in active_merges.items()
                if merge_state.get("last_row", -1) < row_index
            ]
            for key in stale_keys:
                active_merges.pop(key, None)

        active_merges.clear()

    def _layout_table_cell(
        self,
        cell: Optional[TableCell],
        span: int,
        column_width: float,
        parent_context: LayoutContext,
        table_properties: Optional[Dict[str, object]],
    ) -> Dict[str, object]:
        padding = self._resolve_table_cell_padding(cell.properties if cell else None, table_properties)
        borders = self._resolve_table_cell_borders(cell.properties if cell else None, table_properties)

        margin_left = padding["left"] + borders["left"]
        margin_right = padding["right"] + borders["right"]
        margin_top = padding["top"] + borders["top"]
        margin_bottom = padding["bottom"] + borders["bottom"]

        inner_context = LayoutContext(
            page_width=column_width,
            page_height=parent_context.page_height,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            cursor_x=margin_left,
            cursor_y=margin_top,
            header_margin=parent_context.header_margin,
            footer_margin=parent_context.footer_margin,
        )

        cell_boxes: List[LayoutBox] = []
        for block in (cell.content if cell else []):
            cell_boxes.append(self._layout_block(block, inner_context))

        content_height = max(inner_context.cursor_y - margin_top, 0.0)
        total_height = max(content_height + margin_top + margin_bottom, margin_top + margin_bottom)

        return {
            "width": column_width,
            "height": total_height,
            "padding": padding,
            "borders": borders,
            "contentHeight": content_height,
            "boxes": cell_boxes,
            "colSpan": span,
            "columnIndex": 0,
            "rowIndex": 0,
            "vMerge": None,
            "baseRow": None,
            "rowSpan": 1,
            "x": 0.0,
            "y": 0.0,
        }

    def _create_empty_cell(self, width: float, column_index: int, col_span: int) -> Dict[str, object]:
        padding_base = {
            "top": DEFAULT_TABLE_CELL_PADDING_PT,
            "bottom": DEFAULT_TABLE_CELL_PADDING_PT,
            "left": DEFAULT_TABLE_CELL_PADDING_PT,
            "right": DEFAULT_TABLE_CELL_PADDING_PT,
        }
        borders_base = {
            "top": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "bottom": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "left": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "right": DEFAULT_TABLE_BORDER_WIDTH_PT,
        }
        height = DEFAULT_LINE_HEIGHT_PT + 2 * (DEFAULT_TABLE_CELL_PADDING_PT + DEFAULT_TABLE_BORDER_WIDTH_PT)
        return {
            "width": width,
            "height": height,
            "padding": padding_base,
            "borders": borders_base,
            "contentHeight": 0.0,
            "boxes": [],
            "colSpan": col_span,
            "columnIndex": column_index,
            "rowIndex": 0,
            "vMerge": None,
            "baseRow": None,
            "rowSpan": 1,
            "x": 0.0,
            "y": 0.0,
        }

    def _layout_image(self, image: ImageElement, context: LayoutContext) -> LayoutBox:
        wrap_mode = self._resolve_image_wrap_mode(image)
        anchor = self._resolve_image_anchor(image, context)
        margins = self._resolve_image_margins(image)

        floating = wrap_mode != "inline" or anchor.get("anchored")

        if not floating and margins["top"]:
            context.cursor_y += margins["top"]

        max_inline_width = max(context.available_width - margins["left"] - margins["right"], 0.0)
        natural_width = self._emu_to_points(image.width_emu) if image.width_emu else max_inline_width or 96.0
        natural_height = self._emu_to_points(image.height_emu) if image.height_emu else natural_width * 0.75

        width = natural_width
        height = natural_height

        limit_width = max_inline_width if not floating else max(context.available_width - margins["left"] - margins["right"], 0.0)
        if limit_width and width > limit_width:
            scale = limit_width / width
            width = limit_width
            height *= scale

        if floating:
            x = self._compute_anchor_coordinate(
                anchor.get("offset_x"),
                context.margin_left + margins["left"],
                context.available_width - margins["left"] - margins["right"],
                width,
                anchor.get("relative_x"),
            )
            y = self._compute_anchor_coordinate(
                anchor.get("offset_y"),
                context.margin_top + margins["top"],
                context.page_height - context.margin_top - context.margin_bottom,
                height,
                anchor.get("relative_y"),
            )
        else:
            x = context.margin_left + margins["left"]
            y = context.cursor_y

        box = LayoutBox(
            element_type="image",
            content={
                "source": image.media_path,
                "resourceId": image.r_id,
            },
            x=x,
            y=y,
            width=width,
            height=height,
            style={
                "wrapStyle": wrap_mode,
                "inline": not floating,
                "anchor": anchor,
                "properties": image.properties,
                "margins": margins,
            },
        )

        if floating:
            if wrap_mode in {"square", "tight", "through"}:
                context.cursor_y = max(context.cursor_y, y + height + margins["bottom"])
            elif wrap_mode == "behind-text" or wrap_mode == "infront-of-text":
                # Floating behind/in front of text does not influence cursor
                pass
            else:
                context.cursor_y = max(context.cursor_y, y + height + margins["bottom"])
        else:
            context.cursor_y += height + margins["bottom"]

        return box

    def _layout_placeholder(self, block, context: LayoutContext) -> LayoutBox:
        height = DEFAULT_LINE_HEIGHT_PT
        box = LayoutBox(
            element_type="unsupported",
            content={"repr": repr(block)},
            x=context.margin_left,
            y=context.cursor_y,
            width=context.available_width,
            height=height,
            style={},
        )
        context.cursor_y += height
        return box

    # ------------------------------------------------------------------
    # Metric helpers
    def _resolve_font_size(self, paragraph: ParagraphElement, style: Optional[StyleDefinition]) -> float:
        font_sizes: List[float] = []
        for run in paragraph.runs:
            size = self._extract_font_size_from_properties(run.properties)
            if size is not None:
                font_sizes.append(size)

        if not font_sizes:
            style_size = self._extract_font_size_from_style(style)
            if style_size is not None:
                font_sizes.append(style_size)

        if font_sizes:
            return max(font_sizes)

        return DEFAULT_FONT_SIZE_PT

    def _resolve_spacing(self, paragraph: ParagraphElement, style: Optional[StyleDefinition]) -> SpacingInfo:
        direct = self._extract_spacing_info(paragraph.properties)
        styled = self._extract_spacing_info_from_style(style)

        before = self._coalesce_float(direct.before, styled.before, default=0.0)
        after = self._coalesce_float(direct.after, styled.after, default=0.0)
        line = self._coalesce_float(direct.line, styled.line, default=None)
        line_rule = direct.line_rule or styled.line_rule

        return SpacingInfo(before=before, after=after, line=line, line_rule=line_rule)

    def _resolve_line_height(self, font_size: float, spacing: SpacingInfo) -> float:
        base_height = max(font_size * 1.2, DEFAULT_LINE_HEIGHT_PT)

        if spacing.line is None:
            return base_height

        rule = (spacing.line_rule or "auto").lower()

        if rule == "exact":
            return spacing.line
        if rule == "atleast":
            return max(base_height, spacing.line)

        factor = spacing.line / 240.0 if spacing.line > 0 else 1.0
        return base_height * factor

    def _resolve_paragraph_indent(
        self, paragraph: ParagraphElement, style: Optional[StyleDefinition]
    ) -> ParagraphIndent:
        direct = self._extract_indent(paragraph.properties)
        styled = self._extract_indent_from_style(style)

        left = self._coalesce_float(direct.get("left"), styled.get("left"), default=0.0) or 0.0
        right = self._coalesce_float(direct.get("right"), styled.get("right"), default=0.0) or 0.0
        first_line = self._coalesce_float(direct.get("firstLine"), styled.get("firstLine"), default=0.0) or 0.0

        return ParagraphIndent(left=left, right=right, first_line=first_line)

    # ------------------------------------------------------------------
    # Property extraction
    def _extract_font_size_from_properties(self, properties) -> Optional[float]:
        if not properties:
            return None

        if isinstance(properties, dict) and "fontSize" in properties:
            size = properties["fontSize"]
            if isinstance(size, (int, float)):
                return self._normalise_font_size(size)

        for node in self._get_property_nodes(properties):
            tag = node.get("tag", "")
            if tag.endswith("sz") or tag.endswith("szCs"):
                value = self._parse_int_attribute(node, "val")
                if value is not None:
                    return self._normalise_font_size(value)

        return None

    def _extract_font_size_from_style(self, style: Optional[StyleDefinition]) -> Optional[float]:
        if not isinstance(style, StyleDefinition):
            return None

        if "fontSize" in style.properties:
            size = style.properties["fontSize"]
            if isinstance(size, (int, float)):
                return self._normalise_font_size(size)

        for node in style.properties.get("rPr", []):
            tag = node.get("tag", "")
            if tag.endswith("sz") or tag.endswith("szCs"):
                value = self._parse_int_attribute(node, "val")
                if value is not None:
                    return self._normalise_font_size(value)
        return None

    def _extract_spacing_info(self, properties) -> SpacingInfo:
        info = SpacingInfo(before=None, after=None, line=None, line_rule=None)
        if not properties:
            return info

        if isinstance(properties, dict):
            for key in ("spacing_before", "spacingBefore", "space_before"):
                value = properties.get(key)
                if isinstance(value, (int, float)):
                    info.before = value / 20.0
            for key in ("spacing_after", "spacingAfter", "space_after"):
                value = properties.get(key)
                if isinstance(value, (int, float)):
                    info.after = value / 20.0

        spacing_node = self._find_property_node(properties, "spacing")
        if spacing_node:
            before = self._parse_twips_attribute(spacing_node, "before")
            after = self._parse_twips_attribute(spacing_node, "after")
            line = self._parse_line_attribute(spacing_node)
            line_rule = self._get_attribute(spacing_node, "lineRule")

            if before is not None and not self._is_autospacing(spacing_node, "beforeAutospacing"):
                info.before = before
            if after is not None and not self._is_autospacing(spacing_node, "afterAutospacing"):
                info.after = after
            if line is not None:
                info.line = line
            if line_rule:
                info.line_rule = line_rule

        return info

    def _extract_spacing_info_from_style(self, style: Optional[StyleDefinition]) -> SpacingInfo:
        if not isinstance(style, StyleDefinition):
            return SpacingInfo(before=None, after=None, line=None, line_rule=None)
        nodes = style.properties.get("pPr")
        if not nodes:
            return SpacingInfo(before=None, after=None, line=None, line_rule=None)
        return self._extract_spacing_info(nodes)

    def _extract_indent(self, properties) -> Dict[str, Optional[float]]:
        indent: Dict[str, Optional[float]] = {"left": None, "right": None, "firstLine": None}

        if not properties:
            return indent

        if isinstance(properties, dict):
            if isinstance(properties.get("indent_left"), (int, float)):
                indent["left"] = properties["indent_left"] / 20.0
            if isinstance(properties.get("indent_right"), (int, float)):
                indent["right"] = properties["indent_right"] / 20.0
            if isinstance(properties.get("first_line_indent"), (int, float)):
                indent["firstLine"] = properties["first_line_indent"] / 20.0
            if isinstance(properties.get("hanging_indent"), (int, float)):
                indent["firstLine"] = -(properties["hanging_indent"] / 20.0)

        ind_node = self._find_property_node(properties, "ind")
        if ind_node:
            left = self._parse_twips_attribute(ind_node, "left")
            right = self._parse_twips_attribute(ind_node, "right")
            first_line = self._parse_twips_attribute(ind_node, "firstLine")
            hanging = self._parse_twips_attribute(ind_node, "hanging")

            if left is not None:
                indent["left"] = left
            if right is not None:
                indent["right"] = right
            if first_line is not None:
                indent["firstLine"] = first_line
            if hanging is not None:
                indent["firstLine"] = -hanging

        return indent

    def _extract_indent_from_style(self, style: Optional[StyleDefinition]) -> Dict[str, Optional[float]]:
        if not isinstance(style, StyleDefinition):
            return {"left": None, "right": None, "firstLine": None}
        nodes = style.properties.get("pPr")
        if not nodes:
            return {"left": None, "right": None, "firstLine": None}
        return self._extract_indent(nodes)

    # ------------------------------------------------------------------
    # Utility helpers
    def _get_property_nodes(self, properties) -> List[dict]:
        if properties is None:
            return []
        if isinstance(properties, dict):
            return properties.get("children", [])
        if isinstance(properties, list):
            return list(properties)
        return []

    def _find_property_node(self, properties, local_tag: str) -> Optional[dict]:
        for node in self._get_property_nodes(properties):
            tag = node.get("tag", "")
            if tag.endswith(local_tag):
                return node
        return None

    def _parse_twips_attribute(self, node: dict, local_name: str) -> Optional[float]:
        value = self._get_attribute(node, local_name)
        if value is None:
            return None
        try:
            return float(value) / 20.0
        except (TypeError, ValueError):
            return None

    def _parse_line_attribute(self, node: dict) -> Optional[float]:
        value = self._get_attribute(node, "line")
        if value is None:
            return None
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return None

        rule = self._get_attribute(node, "lineRule")
        if rule and rule.lower() in {"exact", "atleast"}:
            return numeric / 20.0
        return numeric

    def _get_attribute(self, node: dict, local_name: str) -> Optional[str]:
        for key, value in node.get("attributes", {}).items():
            if key.endswith(local_name):
                return value
        return None

    def _parse_int_attribute(self, node: dict, local_name: str) -> Optional[int]:
        value = self._get_attribute(node, local_name)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    def _is_autospacing(self, node: dict, local_name: str) -> bool:
        value = self._get_attribute(node, local_name)
        if value is None:
            return False
        return value in {"1", "true", "on"}

    def _coalesce_float(self, *values, default: Optional[float]) -> Optional[float]:
        for value in values:
            if value is not None:
                return value
        return default

    def _compute_table_column_count(self, table: TableElement) -> int:
        max_columns = 0
        for row in table.rows:
            column_index = 0
            for cell in row.cells:
                column_index += self._resolve_grid_span(cell.properties)
            max_columns = max(max_columns, column_index)

        if max_columns == 0:
            max_columns = max((len(row.cells) for row in table.rows), default=0)

        return max(1, max_columns)

    def _resolve_grid_span(self, cell_properties: Optional[Dict[str, object]]) -> int:
        if not cell_properties:
            return 1

        node = self._find_property_node(cell_properties, "gridSpan")
        if node:
            value = self._parse_int_attribute(node, "val")
            if value and value > 0:
                return value

        return 1

    def _resolve_vertical_merge(self, cell_properties: Optional[Dict[str, object]]) -> Optional[str]:
        if not cell_properties:
            return None

        node = None
        if isinstance(cell_properties, dict):
            potential = cell_properties.get("vMerge")
            if isinstance(potential, dict):
                node = potential
            else:
                node = self._find_property_node(cell_properties, "vMerge")
        elif isinstance(cell_properties, list):
            for item in cell_properties:
                if isinstance(item, dict) and item.get("tag", "").endswith("vMerge"):
                    node = item
                    break

        if not node:
            return None

        value = (self._get_attribute(node, "val") or "").lower()
        if value == "restart":
            return "restart"
        if value == "continue" or value == "":
            return "continue"
        return "continue"

    def _compute_table_min_widths(
        self,
        table: TableElement,
        column_count: int,
        available_width: float,
    ) -> List[float]:
        if column_count <= 0:
            return [available_width]

        min_widths = [0.0 for _ in range(column_count)]
        for row in table.rows:
            column_index = 0
            for cell in row.cells:
                span = self._resolve_grid_span(cell.properties)
                span = max(1, min(span, column_count - column_index))
                cell_min_width = self._calculate_cell_min_width(cell, table.properties, available_width, span)
                per_column = cell_min_width / span if span else cell_min_width
                for offset in range(span):
                    idx = column_index + offset
                    if idx < column_count:
                        min_widths[idx] = max(min_widths[idx], per_column)
                column_index += span

        baseline = DEFAULT_LINE_HEIGHT_PT + 2 * (DEFAULT_TABLE_CELL_PADDING_PT + DEFAULT_TABLE_BORDER_WIDTH_PT)
        for idx, width in enumerate(min_widths):
            if width == 0.0:
                min_widths[idx] = baseline

        return min_widths

    def _calculate_cell_min_width(
        self,
        cell: Optional[TableCell],
        table_properties: Optional[Dict[str, object]],
        available_width: float,
        span: int,
    ) -> float:
        if not cell:
            return 0.0

        padding = self._resolve_table_cell_padding(cell.properties, table_properties)
        borders = self._resolve_table_cell_borders(cell.properties, table_properties)
        margin_width = padding["left"] + padding["right"] + borders["left"] + borders["right"]

        content_width = self._measure_cell_content_min_width(cell)
        width_constraint = self._extract_cell_width(cell.properties, available_width)

        min_width = max(content_width, width_constraint or 0.0) + margin_width

        if span <= 0:
            return min_width

        return max(min_width, margin_width)

    def _extract_cell_width(
        self,
        cell_properties: Optional[Dict[str, object]],
        available_width: float,
    ) -> Optional[float]:
        if not cell_properties:
            return None

        node = self._find_property_node(cell_properties, "tcW")
        if not node:
            return None

        type_attr = (self._get_attribute(node, "type") or "").lower()
        raw_value = self._get_attribute(node, "w")
        if raw_value is None:
            return None

        try:
            numeric = float(raw_value)
        except (TypeError, ValueError):
            return None

        if type_attr in {"", "auto"}:
            return None
        if type_attr == "dxa":
            return numeric / 20.0
        if type_attr == "pct":
            return available_width * (numeric / 5000.0)

        return None

    def _measure_cell_content_min_width(self, cell: TableCell) -> float:
        if not cell.content:
            return 0.0

        widths = [self._measure_block_min_width(block) for block in cell.content]
        return max(widths, default=0.0)

    def _measure_block_min_width(self, block) -> float:
        if isinstance(block, ParagraphElement):
            return self._measure_paragraph_min_width(block)
        if isinstance(block, TableElement):
            return self._measure_table_min_width(block)
        if isinstance(block, ImageElement):
            return self._measure_image_min_width(block)
        return DEFAULT_LINE_HEIGHT_PT

    def _measure_paragraph_min_width(self, paragraph: ParagraphElement) -> float:
        resolved_style = self._styles.get(paragraph.style_id)
        font_size = self._resolve_font_size(paragraph, resolved_style)
        indent = self._resolve_paragraph_indent(paragraph, resolved_style)

        text_content = "".join((run.text or "") for run in paragraph.runs)
        width = 0.0
        words = [word for word in text_content.split(" ") if word]
        if words:
            width = max(self._estimate_text_width(word, font_size) for word in words)
        elif text_content:
            width = self._estimate_text_width(text_content, font_size)

        drawing_widths = [self._measure_drawing_min_width(getattr(run, "drawing", None)) for run in paragraph.runs]
        if drawing_widths:
            width = max(width, max(drawing_widths))

        width = max(width, font_size)
        width += max(indent.left, 0.0)
        width += max(indent.first_line, 0.0)
        return width

    def _measure_table_min_width(self, table: TableElement) -> float:
        column_count = self._compute_table_column_count(table)
        min_widths = self._compute_table_min_widths(table, column_count, DEFAULT_PAGE_WIDTH_PT)
        return sum(min_widths)

    def _measure_image_min_width(self, image: ImageElement) -> float:
        width = self._emu_to_points(image.width_emu) if image.width_emu else 0.0
        margins = self._resolve_image_margins(image)
        return width + margins["left"] + margins["right"]

    def _measure_drawing_min_width(self, drawing) -> float:
        if drawing is None:
            return 0.0

        width = self._emu_to_points(getattr(drawing, "width_emu", None))
        if width == 0.0:
            height = getattr(drawing, "height_emu", None)
            if height:
                width = self._emu_to_points(height) * (4.0 / 3.0)
        return width

    def _extract_table_grid_widths(
        self,
        table_properties: Optional[Dict[str, object]],
        column_count: int,
    ) -> List[float]:
        if not table_properties:
            return []

        node = self._find_property_node(table_properties, "tblGrid")
        if not node:
            return []

        widths: List[float] = []
        for child in self._get_property_nodes(node):
            tag = child.get("tag", "")
            if tag.endswith("gridCol"):
                width = self._parse_twips_attribute(child, "w")
                if width is not None:
                    widths.append(width)

        if not widths:
            return []

        if len(widths) < column_count:
            widths.extend([widths[-1]] * (column_count - len(widths)))
        elif len(widths) > column_count:
            widths = widths[:column_count]

        return widths

    def _scale_widths(self, widths: List[float], target_width: float) -> List[float]:
        if not widths:
            return []

        total = sum(widths)
        if total <= 0 or target_width <= 0:
            return list(widths)

        scale = target_width / total
        return [width * scale for width in widths]

    def _resolve_column_widths(
        self,
        table: TableElement,
        available_width: float,
        column_count: int,
        min_widths: List[float],
    ) -> tuple[List[float], float]:
        if column_count <= 0:
            return [available_width], available_width

        grid_widths = self._extract_table_grid_widths(table.properties, column_count)
        if grid_widths:
            column_widths = self._scale_widths(grid_widths, available_width)
        else:
            column_widths = list(min_widths)

        column_widths = [max(column_widths[i], min_widths[i]) for i in range(column_count)]
        total_width = sum(column_widths)

        if total_width < available_width:
            leftover = available_width - total_width
            weights = grid_widths if grid_widths else min_widths
            weight_sum = sum(weights)
            if weight_sum <= 0:
                weights = [1.0] * column_count
                weight_sum = float(column_count)
            for i in range(column_count):
                share = leftover * (weights[i] / weight_sum)
                column_widths[i] += share
            total_width = sum(column_widths)
        else:
            excess = total_width - available_width
            if excess > 0:
                adjustable = [column_widths[i] - min_widths[i] for i in range(column_count)]
                adjustable_sum = sum(value for value in adjustable if value > 0)
                if adjustable_sum > 0:
                    for i in range(column_count):
                        extra = adjustable[i]
                        if extra > 0:
                            reduction = min(extra, excess * (extra / adjustable_sum))
                            column_widths[i] -= reduction
                    total_width = sum(column_widths)

        total_width = max(total_width, sum(min_widths)) if min_widths else total_width
        return column_widths, total_width

    def _resolve_table_cell_padding(
        self,
        cell_properties: Optional[Dict[str, object]],
        table_properties: Optional[Dict[str, object]],
    ) -> Dict[str, float]:
        padding = {
            "top": DEFAULT_TABLE_CELL_PADDING_PT,
            "bottom": DEFAULT_TABLE_CELL_PADDING_PT,
            "left": DEFAULT_TABLE_CELL_PADDING_PT,
            "right": DEFAULT_TABLE_CELL_PADDING_PT,
        }

        if table_properties:
            node = self._find_property_node(table_properties, "tblCellMar")
            if node:
                self._apply_margin_node(node, padding)

        if cell_properties:
            node = self._find_property_node(cell_properties, "tcMar")
            if node:
                self._apply_margin_node(node, padding)

        return padding

    def _apply_margin_node(self, node: dict, target: Dict[str, float]) -> None:
        for child in self._get_property_nodes(node):
            tag = child.get("tag", "")
            local = tag.split("}")[-1]
            if local in target:
                value = self._parse_twips_attribute(child, "w")
                type_attr = (self._get_attribute(child, "type") or "").lower()
                if type_attr == "nil":
                    target[local] = 0.0
                elif value is not None:
                    target[local] = value

    def _resolve_table_cell_borders(
        self,
        cell_properties: Optional[Dict[str, object]],
        table_properties: Optional[Dict[str, object]],
    ) -> Dict[str, float]:
        borders = {
            "top": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "bottom": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "left": DEFAULT_TABLE_BORDER_WIDTH_PT,
            "right": DEFAULT_TABLE_BORDER_WIDTH_PT,
        }

        if table_properties:
            node = self._find_property_node(table_properties, "tblBorders")
            if node:
                self._apply_border_node(node, borders)

        if cell_properties:
            node = self._find_property_node(cell_properties, "tcBorders")
            if node:
                self._apply_border_node(node, borders)

        return borders

    def _apply_border_node(self, node: dict, target: Dict[str, float]) -> None:
        for child in self._get_property_nodes(node):
            tag = child.get("tag", "")
            local = tag.split("}")[-1]
            if local in target:
                width = self._parse_border_width(child)
                if width is not None:
                    target[local] = width

    def _parse_border_width(self, node: dict) -> Optional[float]:
        style = (self._get_attribute(node, "val") or "").lower()
        if style in {"nil", "none"}:
            return 0.0

        size_attr = self._get_attribute(node, "sz")
        if size_attr is None:
            return None
        try:
            size = float(size_attr)
        except (TypeError, ValueError):
            return None
        return size / 8.0

    def _resolve_image_margins(self, image: ImageElement) -> Dict[str, float]:
        properties = image.properties or {}
        return {
            "top": self._coerce_measurement(properties.get("margin_top")) or 0.0,
            "bottom": self._coerce_measurement(properties.get("margin_bottom")) or 0.0,
            "left": self._coerce_measurement(properties.get("margin_left")) or 0.0,
            "right": self._coerce_measurement(properties.get("margin_right")) or 0.0,
        }

    def _resolve_image_wrap_mode(self, image: ImageElement) -> str:
        properties = image.properties or {}
        value: Optional[str] = None
        for key in ("wrapStyle", "wrap_style", "wrap", "positioning"):
            candidate = properties.get(key)
            if isinstance(candidate, str):
                value = candidate.lower()
                break

        if not value:
            return "inline"

        mapping = {
            "inline": "inline",
            "square": "square",
            "tight": "tight",
            "through": "through",
            "behind": "behind-text",
            "behindtext": "behind-text",
            "behind-text": "behind-text",
            "infront": "infront-of-text",
            "infronttext": "infront-of-text",
            "infront-text": "infront-of-text",
            "front": "infront-of-text",
            "none": "inline",
        }

        return mapping.get(value, value)

    def _resolve_image_anchor(self, image: ImageElement, context: LayoutContext) -> Dict[str, object]:
        properties = image.properties or {}
        anchor_props = {}
        anchor_value = properties.get("anchor") or properties.get("position")
        if isinstance(anchor_value, dict):
            anchor_props = anchor_value

        merged = {**properties, **anchor_props}

        offset_x = self._coerce_measurement(
            merged.get("offset_x")
            or merged.get("offsetX")
            or merged.get("x")
            or merged.get("left")
        )
        offset_y = self._coerce_measurement(
            merged.get("offset_y")
            or merged.get("offsetY")
            or merged.get("y")
            or merged.get("top")
        )

        relative_x = merged.get("relative_x") or merged.get("relativeX") or merged.get("align")
        relative_y = merged.get("relative_y") or merged.get("relativeY")

        inline_flag = merged.get("inline")
        anchored = (inline_flag is False) or offset_x is not None or offset_y is not None

        return {
            "anchored": bool(anchored),
            "offset_x": offset_x,
            "offset_y": offset_y,
            "relative_x": relative_x,
            "relative_y": relative_y,
        }

    def _compute_anchor_coordinate(
        self,
        offset: Optional[float],
        origin: float,
        extent: float,
        size: float,
        relative: Optional[str],
    ) -> float:
        position = origin

        if isinstance(relative, str):
            key = relative.lower()
            if key in {"center", "centre", "middle"}:
                position = origin + max((extent - size) / 2.0, 0.0)
            elif key in {"right", "end"}:
                position = origin + max(extent - size, 0.0)
            elif key in {"bottom"}:
                position = origin + max(extent - size, 0.0)

        if offset is not None:
            position = origin + offset

        return position

    @staticmethod
    def _emu_to_points(value: Optional[int]) -> float:
        if not value:
            return 0.0
        return float(value) / EMU_PER_POINT

    # ------------------------------------------------------------------
    # Text wrapping
    def _wrap_text(self, text: str, max_width: float, font_size: float) -> List[str]:
        if not text:
            return [""]

        words = text.split(" ")
        if not words:
            return [text]

        lines: List[str] = []
        current_line: List[str] = []
        current_width = 0.0
        space_width = self._estimate_text_width(" ", font_size)

        for word in words:
            word_width = self._estimate_text_width(word, font_size)
            projected = current_width + (space_width if current_line else 0.0) + word_width

            if current_line and projected > max_width:
                lines.append(" ".join(current_line))
                current_line = [word]
                current_width = word_width
            else:
                if current_line:
                    current_width += space_width + word_width
                else:
                    current_width = word_width
                current_line.append(word)

        if current_line:
            lines.append(" ".join(current_line))

        return lines

    @staticmethod
    def _estimate_text_width(text: str, font_size: float) -> float:
        if not text:
            return 0.0

        average_width_factor = 0.5  # Approximation for Latin alphabets
        return len(text) * font_size * average_width_factor

    @staticmethod
    def _normalise_font_size(size: int | float) -> float:
        if size > 20:
            return float(size) / 2.0
        return float(size)

    @staticmethod
    def _twips_to_points(value: int | float | None, default: float) -> float:
        if value is None:
            return default
        return float(value) / 20.0

    def _coerce_measurement(self, value) -> Optional[float]:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            text = value.strip().lower()
            unit_map = {
                "pt": 1.0,
                "in": 72.0,
                "cm": 28.3465,
                "mm": 2.83465,
                "px": 0.75,
            }
            for unit, factor in unit_map.items():
                if text.endswith(unit):
                    try:
                        return float(text[:-len(unit)]) * factor
                    except ValueError:
                        return None
            try:
                numeric = float(text)
            except ValueError:
                return None
            if numeric > 1000:
                return numeric / 20.0
            return numeric
        return None

