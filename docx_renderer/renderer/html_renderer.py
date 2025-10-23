"""Render the layout model into an HTML document."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

from docx_renderer.model.document_model import DocumentModel
from docx_renderer.model.elements import LayoutBox
from docx_renderer.renderer.utils import style_to_css


class HtmlRenderer:
    """Produce an absolutely positioned HTML representation of the document."""

    def __init__(self, output_path: Path) -> None:
        self._output_path = output_path

    def render(self, model: DocumentModel) -> None:
        boxes = model.layout.boxes
        html = self._build_html(boxes)
        self._output_path.write_text(html, encoding="utf-8")

    def _build_html(self, boxes: Iterable[LayoutBox]) -> str:
        elements = [self._box_to_div(box) for box in boxes]
        body = "\n".join(elements)
        return f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>DOCX Preview</title>
  <style>
    body {{ position: relative; margin: 0; padding: 0; }}
    .docx-box {{ position: absolute; white-space: pre-wrap; }}
  </style>
</head>
<body>
{body}
</body>
</html>
"""

    def _box_to_div(self, box: LayoutBox) -> str:
        style = {
            "left": f"{box.x}px",
            "top": f"{box.y}px",
            "width": f"{box.width}px",
            "height": f"{box.height}px",
        }
        style.update(style_to_css(box.style))
        style_str = "; ".join(f"{k}: {v}" for k, v in style.items())
        content = box.content.get("text", box.content.get("repr", ""))
        return f"  <div class=\"docx-box\" style=\"{style_str}\">{content}</div>"
