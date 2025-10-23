"""Entry-point for the docx renderer pipeline."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx_renderer.model.document_model import DocumentModel
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.document_parser import DocumentParser
from docx_renderer.parser.layout_calculator import LayoutCalculator
from docx_renderer.parser.styles_parser import StylesParser
from docx_renderer.renderer.html_renderer import HtmlRenderer
from docx_renderer.renderer.pdf_renderer import PdfRenderer
from docx_renderer.utils.debug import DebugDumper
from docx_renderer.utils.logger import get_logger

LOGGER = get_logger(__name__)


def build_document_model(docx_path: Path) -> DocumentModel:
    """Load a DOCX package, parse WordprocessingML, and build an internal model."""
    package = DocxPackage.load(docx_path)
    styles = StylesParser(package.require_styles_xml(), package.get_numbering_xml()).parse()
    document_tree = DocumentParser(package, styles).parse()
    layout_model = LayoutCalculator(styles).calculate(document_tree)
    return DocumentModel(styles=styles, layout=layout_model)


def render_outputs(model: DocumentModel, output_dir: Path, *, html: bool = True, pdf: bool = False) -> None:
    """Render the flattened model into the requested formats."""
    output_dir.mkdir(parents=True, exist_ok=True)
    if html:
        HtmlRenderer(output_dir / "document.html").render(model)
    if pdf:
        PdfRenderer(output_dir / "document.pdf").render(model)


def main(docx_file: str, output_dir: Optional[str] = None) -> None:
    """Run the DOCX → intermediate model → renderer pipeline."""
    docx_path = Path(docx_file).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    LOGGER.info("Building document model for %s", docx_path.name)
    model = build_document_model(docx_path)

    if output_dir is None:
        output_dir = docx_path.with_suffix("")

    output_path = Path(output_dir).resolve()
    LOGGER.info("Rendering outputs into %s", output_path)
    render_outputs(model, output_path)

    DebugDumper(output_path / "debug").dump(model)


if __name__ == "__main__":  # pragma: no cover
    import argparse

    parser = argparse.ArgumentParser(description="Render DOCX files into HTML/PDF using a flattened model")
    parser.add_argument("docx_file", help="Path to the input .docx file")
    parser.add_argument("--output", help="Directory to write generated artifacts")
    parser.add_argument("--pdf", action="store_true", help="Generate a PDF output as well as HTML")

    args = parser.parse_args()
    doc_model = build_document_model(Path(args.docx_file))
    render_outputs(doc_model, Path(args.output or Path(args.docx_file).with_suffix("")), html=True, pdf=args.pdf)
    DebugDumper(Path(args.output or Path(args.docx_file).with_suffix("")) / "debug").dump(doc_model)
