"""Parse document.xml into structured content blocks."""
from __future__ import annotations

from typing import List
from xml.etree import ElementTree as ET

from docx_renderer.model.elements import DocumentTree, ParagraphElement, RunFragment
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.media_extractor import MediaResolver
from docx_renderer.utils.logger import get_logger
from docx_renderer.utils.xml_utils import Namespaces

LOGGER = get_logger(__name__)


class DocumentParser:
    """Transforms Word body XML into model elements."""

    def __init__(self, package: DocxPackage, styles: StylesCatalog) -> None:
        self._package = package
        self._styles = styles
        self._media = MediaResolver(package.relationships, package.media)

    def parse(self) -> DocumentTree:
        """Parse the document body into high-level block elements."""
        root = self._package.document_xml.getroot()
        body = root.find("w:body", Namespaces.WORD)
        if body is None:
            LOGGER.warning("document.xml missing body element")
            return DocumentTree(blocks=[])

        blocks = []
        for child in list(body):
            tag = self._strip_namespace(child.tag)
            if tag == "p":
                blocks.append(self._parse_paragraph(child))
            elif tag == "tbl":
                blocks.append(self._parse_table(child))
            elif tag == "sectPr":
                # Section properties influence layout; stored as metadata later.
                continue
            else:
                LOGGER.debug("Skipping unsupported element: %s", tag)
        return DocumentTree(blocks=blocks)

    def _parse_paragraph(self, paragraph_el: ET.Element) -> ParagraphElement:
        runs: List[RunFragment] = []
        for run_el in paragraph_el.findall("w:r", Namespaces.WORD):
            text_el = run_el.find("w:t", Namespaces.WORD)
            text = text_el.text if text_el is not None else ""
            props = {}
            if run_el.find("w:b", Namespaces.WORD) is not None:
                props["bold"] = True
            runs.append(RunFragment(text=text, properties=props))
        style_id = self._get_style_id(paragraph_el)
        return ParagraphElement(runs=runs, style_id=style_id, properties={})

    def _parse_table(self, table_el: ET.Element):  # TODO: return TableElement
        LOGGER.debug("Table parsing is not yet implemented")
        return ParagraphElement(runs=[RunFragment(text="[Table placeholder]")], style_id=None, properties={})

    def _get_style_id(self, element: ET.Element) -> str | None:
        ppr = element.find("w:pPr", Namespaces.WORD)
        if ppr is None:
            return None
        style_el = ppr.find("w:pStyle", Namespaces.WORD)
        if style_el is None:
            return None
        attr_key = self._qualify("w:val")
        return style_el.attrib.get(attr_key)

    def _strip_namespace(self, tag: str) -> str:
        return tag.split("}", 1)[-1]

    def _qualify(self, attr_name: str) -> str:
        prefix, local = attr_name.split(":", 1)
        namespace = Namespaces.WORD[prefix]
        return f"{{{namespace}}}{local}"
