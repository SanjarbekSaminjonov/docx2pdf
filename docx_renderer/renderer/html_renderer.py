"""Render the layout model into an HTML document."""
from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

from docx_renderer.model.document_model import DocumentModel
from docx_renderer.model.elements import LayoutBox, MediaAsset


class HtmlRenderer:
    """Produce an HTML representation of the calculated layout."""

    def __init__(self, output_path: Path) -> None:
        self._output_path = output_path

    def render(self, model: DocumentModel) -> None:
        html = self._build_html(model)
        self._output_path.write_text(html, encoding="utf-8")

    def _build_html(self, model: DocumentModel) -> str:
        pages = list(model.layout.pages)
        if not pages:
            pages = [list(model.layout.boxes)]

        rendered_pages = [self._render_page(page_index, page_boxes, model) for page_index, page_boxes in enumerate(pages)]
        body = "\n".join(rendered_pages)

        return """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>DOCX Preview</title>
  <style>
    :root { color-scheme: light; }
    body { margin: 0; padding: 24px; background: #f5f5f5; font-family: "Segoe UI", Arial, sans-serif; }
    .docx-document { margin: 0 auto; max-width: 900px; }
    .docx-page { position: relative; margin: 24px auto; background: #ffffff; box-shadow: 0 2px 18px rgba(0, 0, 0, 0.12); box-sizing: border-box; overflow: visible; }
    .docx-box { position: absolute; box-sizing: border-box; }
    .docx-box.docx-flow { position: relative; }
    .docx-paragraph { white-space: pre-wrap; line-height: 1.4; font-size: 12pt; }
    .docx-table table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    .docx-table td { vertical-align: top; }
    .docx-image img { display: block; object-fit: contain; }
  </style>
</head>
<body>
  <main class="docx-document">
{body}
  </main>
</body>
</html>
"""

    def _render_page(self, page_index: int, boxes: Sequence[LayoutBox], model: DocumentModel) -> str:
        width, height = self._measure_page_dimensions(boxes)
        style = self._style_to_string({
            "width": self._format_px(width),
            "height": self._format_px(height),
        })

        content = "\n".join(
            self._render_box(box, model, origin=(0.0, 0.0), absolute=True)
            for box in boxes
        )

        return f"    <section class=\"docx-page\" data-page-index=\"{page_index}\" style=\"{style}\">\n{content}\n    </section>"

    def _measure_page_dimensions(self, boxes: Sequence[LayoutBox]) -> Tuple[float, float]:
        if not boxes:
            return (595.0, 842.0)

        max_right = max(box.x + box.width for box in boxes)
        max_bottom = max(box.y + box.height for box in boxes)
        return max(max_right, 100.0), max(max_bottom, 100.0)

    def _render_box(
        self,
        box: LayoutBox,
        model: DocumentModel,
        *,
        origin: Tuple[float, float],
        absolute: bool,
    ) -> str:
        element = (box.element_type or "unknown").lower()

        if element in {"header", "footer"}:
            return self._render_header_footer(box, model, origin=origin)
        if element == "paragraph":
            return self._render_paragraph(box, origin=origin, absolute=absolute)
        if element == "table":
            return self._render_table(box, model, origin=origin, absolute=absolute)
        if element == "image":
            return self._render_image(box, model, origin=origin, absolute=absolute)

        return self._render_generic(box, origin=origin, absolute=absolute)

    def _render_header_footer(self, box: LayoutBox, model: DocumentModel, *, origin: Tuple[float, float]) -> str:
        outer_style = self._style_for_box(box, origin=origin, absolute=True)
        style_attr = self._style_to_string(outer_style)
        child_origin = (box.x, box.y)

        child_boxes: Iterable[LayoutBox] = box.content.get("boxes", [])  # type: ignore[assignment]
        children = "\n".join(
            self._render_box(child, model, origin=child_origin, absolute=True)
            for child in child_boxes
        )

        element_attr = escape(box.element_type)
        return f"      <div class=\"docx-box docx-{element_attr}\" data-element=\"{element_attr}\" style=\"{style_attr}\">\n{children}\n      </div>"

    def _render_paragraph(self, box: LayoutBox, *, origin: Tuple[float, float], absolute: bool) -> str:
        style = self._style_for_box(box, origin=origin, absolute=absolute)

        spacing: Dict[str, float] = box.style.get("spacing", {}) if isinstance(box.style.get("spacing"), dict) else {}
        indent: Dict[str, float] = box.style.get("indent", {}) if isinstance(box.style.get("indent"), dict) else {}

        if spacing.get("before"):
            style["margin-top"] = self._format_px(spacing["before"])
        if spacing.get("after"):
            style["margin-bottom"] = self._format_px(spacing["after"])
        if indent.get("firstLine"):
            style["text-indent"] = self._format_px(indent["firstLine"])

        style.setdefault("white-space", "pre-wrap")
        style.setdefault("line-height", "1.4")

        style_attr = self._style_to_string(style)

        lines: List[str] = box.content.get("lines", [])  # type: ignore[assignment]
        if not lines:
            lines = [box.content.get("text", "")]

        def _escape_line(value: str) -> str:
            return escape(value) if value else "&nbsp;"

        inner = "<br />".join(_escape_line(line) for line in lines)
        element_attr = escape(box.element_type)
        class_name = "docx-paragraph"
        extra_class = " docx-flow" if not absolute else ""
        return f"      <div class=\"docx-box{extra_class} {class_name}\" data-element=\"{element_attr}\" style=\"{style_attr}\">{inner}</div>"

    def _render_table(
        self,
        box: LayoutBox,
        model: DocumentModel,
        *,
        origin: Tuple[float, float],
        absolute: bool,
    ) -> str:
        wrapper_style = self._style_for_box(box, origin=origin, absolute=absolute)
        wrapper_style.setdefault("overflow", "visible")
        wrapper_attr = self._style_to_string(wrapper_style)

        content = box.content
        column_widths: Sequence[float] = content.get("columnWidths", [])  # type: ignore[assignment]
        rows: Sequence[Sequence[Dict[str, object]]] = content.get("cells", [])  # type: ignore[assignment]

        colgroup = "".join(
            f"          <col style=\"width: {self._format_px(width)}\" />"
            for width in column_widths
        )

        table_rows = []
        for row_cells in rows:
            cell_html = []
            for cell in row_cells:
                if cell.get("rowSpan", 1) == 0 or cell.get("height", 0.0) == 0.0:
                    continue

                colspan = int(cell.get("colSpan", 1) or 1)
                rowspan = int(cell.get("rowSpan", 1) or 1)

                attributes = ""
                if colspan > 1:
                    attributes += f" colspan=\"{colspan}\""
                if rowspan > 1:
                    attributes += f" rowspan=\"{rowspan}\""

                style_attr = self._style_to_string(self._table_cell_style(cell))
                inner_html = self._render_table_cell_content(cell, model)
                cell_html.append(f"        <td{attributes} style=\"{style_attr}\">{inner_html}</td>")

            table_rows.append("      <tr>\n" + "\n".join(cell_html) + "\n      </tr>")

        table_style = self._style_to_string({
            "width": "100%",
            "border-collapse": "collapse",
            "table-layout": "fixed",
        })

        table_markup = (
            "      <table class=\"docx-table-grid\" style=\""
            + table_style
            + "\">\n        <colgroup>\n"
            + colgroup
            + "\n        </colgroup>\n"
            + "\n".join(table_rows)
            + "\n      </table>"
        )

        return (
            "      <div class=\"docx-box docx-flow docx-table\" data-element=\"table\" style=\""
            + wrapper_attr
            + "\">\n"
            + table_markup
            + "\n      </div>"
        )

    def _render_table_cell_content(self, cell: Dict[str, object], model: DocumentModel) -> str:
        boxes: Iterable[LayoutBox] = cell.get("boxes", [])  # type: ignore[assignment]
        if not boxes:
            return "&nbsp;"

        fragments = [
            self._render_box(child, model, origin=(0.0, 0.0), absolute=False)
            for child in boxes
        ]
        return "".join(fragments)

    def _table_cell_style(self, cell: Dict[str, object]) -> Dict[str, str]:
        padding: Dict[str, float] = cell.get("padding", {}) if isinstance(cell.get("padding"), dict) else {}
        borders: Dict[str, float] = cell.get("borders", {}) if isinstance(cell.get("borders"), dict) else {}

        style: Dict[str, str] = {"vertical-align": "top", "box-sizing": "border-box"}
        style["padding-top"] = self._format_px(padding.get("top", 0.0))
        style["padding-right"] = self._format_px(padding.get("right", 0.0))
        style["padding-bottom"] = self._format_px(padding.get("bottom", 0.0))
        style["padding-left"] = self._format_px(padding.get("left", 0.0))

        style["border-top"] = self._border_css(borders.get("top"))
        style["border-right"] = self._border_css(borders.get("right"))
        style["border-bottom"] = self._border_css(borders.get("bottom"))
        style["border-left"] = self._border_css(borders.get("left"))

        if cell.get("width"):
            style["width"] = self._format_px(float(cell["width"]))
        if cell.get("height"):
            style["min-height"] = self._format_px(float(cell["height"]))

        return style

    def _render_image(
        self,
        box: LayoutBox,
        model: DocumentModel,
        *,
        origin: Tuple[float, float],
        absolute: bool,
    ) -> str:
        wrapper_style = self._style_for_box(box, origin=origin, absolute=absolute)
        wrapper_class = "docx-box docx-image" + (" docx-flow" if not absolute else "")
        style_attr = self._style_to_string(wrapper_style)

        img_style: Dict[str, str] = {
            "width": self._format_px(box.width),
            "height": self._format_px(box.height),
        }

        src = self._resolve_image_source(box, model)
        alt_text = escape(box.content.get("alt", "")) if isinstance(box.content, dict) else ""

        return (
            f"      <div class=\"{wrapper_class}\" data-element=\"image\" style=\"{style_attr}\">"
            f"<img src=\"{src}\" alt=\"{alt_text}\" style=\"{self._style_to_string(img_style)}\" />"
            "</div>"
        )

    def _resolve_image_source(self, box: LayoutBox, model: DocumentModel) -> str:
        if not isinstance(box.content, dict):
            return box.content.get("source", "") if hasattr(box.content, "get") else ""

        resource_id = box.content.get("resourceId")
        if not resource_id or not model.media:
            return escape(box.content.get("source", ""))

        asset = model.media.get_by_id(str(resource_id))
        if not isinstance(asset, MediaAsset):
            return escape(box.content.get("source", ""))

        mime_type = asset.media_type or "application/octet-stream"
        return f"data:{mime_type};base64,{asset.base64_data}"

    def _render_generic(self, box: LayoutBox, *, origin: Tuple[float, float], absolute: bool) -> str:
        style = self._style_for_box(box, origin=origin, absolute=absolute)
        content = box.content.get("text") or box.content.get("repr") or ""
        text = escape(str(content))
        style_attr = self._style_to_string(style)
        cls_suffix = escape(box.element_type or "unknown")
        extra_class = " docx-flow" if not absolute else ""
        return f"      <div class=\"docx-box{extra_class} docx-{cls_suffix}\" data-element=\"{cls_suffix}\" style=\"{style_attr}\">{text}</div>"

    def _style_for_box(
        self,
        box: LayoutBox,
        *,
        origin: Tuple[float, float],
        absolute: bool,
    ) -> Dict[str, str]:
        styles: Dict[str, str] = {"box-sizing": "border-box"}

        if absolute:
            left = box.x - origin[0]
            top = box.y - origin[1]
            styles["position"] = "absolute"
            styles["left"] = self._format_px(left)
            styles["top"] = self._format_px(top)
            styles["width"] = self._format_px(box.width)
            styles["height"] = self._format_px(box.height)
        else:
            styles["position"] = "relative"

        z_index = box.style.get("zIndex") if isinstance(box.style, dict) else None
        if isinstance(z_index, (int, float)):
            styles["z-index"] = str(int(z_index)) if float(z_index).is_integer() else f"{z_index:.2f}"

        return styles

    def _style_to_string(self, style: Dict[str, str]) -> str:
        return "; ".join(f"{key}: {value}" for key, value in style.items() if value)

    def _format_px(self, value: float) -> str:
        return f"{value:.2f}px"

    def _border_css(self, width: float | None) -> str:
        if not width or width <= 0:
            return "none"
        return f"{width:.2f}px solid #444"
