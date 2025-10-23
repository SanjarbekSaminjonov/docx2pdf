"""DOCX package loader responsible for unpacking XML parts and media."""
from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional
from xml.etree import ElementTree as ET

from docx_renderer.parser.rels_parser import Relationships
from docx_renderer.utils.logger import get_logger
from docx_renderer.utils.xml_utils import parse_xml

LOGGER = get_logger(__name__)


@dataclass(slots=True)
class DocxPackage:
    """Container for the XML parts and media extracted from a DOCX archive."""

    document_xml: ET.ElementTree
    styles_xml: ET.ElementTree
    numbering_xml: Optional[ET.ElementTree]
    theme_xml: Optional[ET.ElementTree]
    headers: Dict[str, ET.ElementTree]
    footers: Dict[str, ET.ElementTree]
    relationships: Relationships
    media: Dict[str, bytes]

    @classmethod
    def load(cls, docx_path: Path) -> "DocxPackage":
        """Open a DOCX archive and populate XML trees and related assets."""
        with zipfile.ZipFile(docx_path) as docx_zip:
            parts = {name: docx_zip.read(name) for name in docx_zip.namelist()}

        LOGGER.debug("Loaded %d parts from %s", len(parts), docx_path.name)

        document_xml = parse_xml(parts["word/document.xml"])
        styles_xml = parse_xml(parts["word/styles.xml"])
        numbering_xml = parse_optional_xml(parts, "word/numbering.xml")
        theme_xml = parse_optional_xml(parts, "word/theme/theme1.xml")

        headers = collect_related_parts(parts, prefix="word/header")
        footers = collect_related_parts(parts, prefix="word/footer")

        relationships = Relationships.from_package(parts)
        media = {name: data for name, data in parts.items() if name.startswith("word/media/")}

        return cls(
            document_xml=document_xml,
            styles_xml=styles_xml,
            numbering_xml=numbering_xml,
            theme_xml=theme_xml,
            headers=headers,
            footers=footers,
            relationships=relationships,
            media=media,
        )


def parse_optional_xml(parts: Dict[str, bytes], name: str) -> Optional[ET.ElementTree]:
    """Return an element tree if the part exists, else None."""
    if name not in parts:
        return None
    return parse_xml(parts[name])


def collect_related_parts(parts: Dict[str, bytes], prefix: str) -> Dict[str, ET.ElementTree]:
    """Collect header/footer XML parts that share a common prefix."""
    collected: Dict[str, ET.ElementTree] = {}
    for name, data in parts.items():
        if name.startswith(prefix) and name.endswith(".xml"):
            collected[name] = parse_xml(data)
    return collected
