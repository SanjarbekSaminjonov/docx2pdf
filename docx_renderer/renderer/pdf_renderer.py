"""Render the layout model into a PDF file using ReportLab (planned)."""
from __future__ import annotations

from pathlib import Path

from docx_renderer.model.document_model import DocumentModel
from docx_renderer.utils.logger import get_logger

LOGGER = get_logger(__name__)


class PdfRenderer:
    """Placeholder PDF renderer that will integrate with ReportLab later."""

    def __init__(self, output_path: Path) -> None:
        self._output_path = output_path

    def render(self, model: DocumentModel) -> None:
        LOGGER.warning("PDF rendering not yet implemented; skipping %s", self._output_path.name)
