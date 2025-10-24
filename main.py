"""Entry-point for the docx renderer pipeline."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx_renderer.model.document_model import DocumentModel
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.document_parser import DocumentParser
from docx_renderer.parser.layout_calculator import LayoutCalculator
from docx_renderer.parser.media_extractor import extract_media_from_package
from docx_renderer.parser.numbering_parser import NumberingParser
from docx_renderer.parser.rels_parser import Relationships
from docx_renderer.parser.styles_parser import StylesParser
from docx_renderer.renderer.html_renderer import HtmlRenderer
from docx_renderer.renderer.pdf_renderer import PdfRenderer
from docx_renderer.utils.debug import DebugDumper
from docx_renderer.utils.logger import get_logger

LOGGER = get_logger(__name__)


def build_document_model(docx_path: Path) -> DocumentModel:
    """Load a DOCX package, parse WordprocessingML, and build an internal model."""
    package = DocxPackage.load(docx_path)
    relationships = Relationships.from_package(package.raw_parts)
    media_catalog = extract_media_from_package(package, relationships)
    styles = StylesParser(package.require_styles_xml(), package.get_numbering_xml()).parse()
    numbering = NumberingParser(package.get_numbering_xml()).parse()
    document_tree = DocumentParser(package, styles, numbering).parse()
    layout_model = LayoutCalculator(styles).calculate(document_tree)
    return DocumentModel(styles=styles, layout=layout_model, numbering=numbering, media=media_catalog)


def render_outputs(model: DocumentModel, output_dir: Path, *, html: bool = True, pdf: bool = False) -> None:
    """Render the flattened model into the requested formats."""
    output_dir.mkdir(parents=True, exist_ok=True)
    if html:
        HtmlRenderer(output_dir / "document.html").render(model)
    if pdf:
        PdfRenderer(output_dir / "document.pdf").render(model)


def run_pipeline(
    docx_file: Path | str,
    *,
    output_dir: Path | str | None = None,
    html: bool = True,
    pdf: bool = False,
    dump_debug: bool = True,
) -> Path:
    """Execute the end-to-end DOCX rendering pipeline and return output directory."""
    docx_path = Path(docx_file).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    LOGGER.info("Building document model for %s", docx_path)
    model = build_document_model(docx_path)

    target_dir: Path
    if output_dir is None:
        target_dir = docx_path.with_suffix("")
    else:
        target_dir = Path(output_dir).resolve()

    LOGGER.info("Rendering outputs into %s", target_dir)
    render_outputs(model, target_dir, html=html, pdf=pdf)

    if dump_debug:
        DebugDumper(target_dir / "debug").dump(model)

    return target_dir


def main(docx_file: str, output_dir: Optional[str] = None) -> None:
    """Run the DOCX → intermediate model → renderer pipeline."""
    run_pipeline(docx_file, output_dir=output_dir, html=True, pdf=False, dump_debug=True)


if __name__ == "__main__":  # pragma: no cover
    import importlib

    try:
        cli_main = importlib.import_module("docx2pdf.cli").main
    except ModuleNotFoundError:  # Running as a script without package install
        from cli import main as cli_main  # type: ignore
    raise SystemExit(cli_main())
