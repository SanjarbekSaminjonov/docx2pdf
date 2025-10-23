"""DOCX package loader responsible for unpacking XML parts and media."""
from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Mapping, Optional
from xml.etree import ElementTree as ET

from docx_renderer.parser.rels_parser import Relationships
from docx_renderer.utils.logger import get_logger
from docx_renderer.utils.xml_utils import parse_xml

LOGGER = get_logger(__name__)

PACKAGE_REL_PATH = "_rels/.rels"
DOCUMENT_XML_PATH = "word/document.xml"
STYLES_XML_PATH = "word/styles.xml"
NUMBERING_XML_PATH = "word/numbering.xml"
DOCUMENT_RELS_PATH = "word/_rels/document.xml.rels"
FOOTNOTES_XML_PATH = "word/footnotes.xml"
ENDNOTES_XML_PATH = "word/endnotes.xml"
COMMENTS_XML_PATH = "word/comments.xml"
SETTINGS_XML_PATH = "word/settings.xml"
GLOSSARY_XML_PATH = "word/glossary/document.xml"
CORE_PROPS_PATH = "docProps/core.xml"
APP_PROPS_PATH = "docProps/app.xml"
CUSTOM_PROPS_PATH = "docProps/custom.xml"


@dataclass(slots=True)
class DocxPackage:
    """Container for the XML parts and media extracted from a DOCX archive."""

    raw_parts: Mapping[str, bytes]
    xml_cache: Dict[str, ET.ElementTree] = field(default_factory=dict)

    content_types_xml: ET.ElementTree | None = None
    package_rels_xml: ET.ElementTree | None = None

    document_xml: ET.ElementTree | None = None
    styles_xml: ET.ElementTree | None = None
    numbering_xml: Optional[ET.ElementTree] = None
    document_rels_xml: Optional[ET.ElementTree] = None
    footnotes_xml: Optional[ET.ElementTree] = None
    endnotes_xml: Optional[ET.ElementTree] = None
    comments_xml: Optional[ET.ElementTree] = None
    settings_xml: Optional[ET.ElementTree] = None
    glossary_xml: Optional[ET.ElementTree] = None

    core_properties_xml: Optional[ET.ElementTree] = None
    app_properties_xml: Optional[ET.ElementTree] = None
    custom_properties_xml: Optional[ET.ElementTree] = None

    headers: Dict[str, ET.ElementTree] = field(default_factory=dict)
    footers: Dict[str, ET.ElementTree] = field(default_factory=dict)
    theme_parts: Dict[str, ET.ElementTree] = field(default_factory=dict)

    relationships: Relationships = field(init=False)
    media: Dict[str, bytes] = field(default_factory=dict)

    @classmethod
    def load(cls, docx_path: Path) -> "DocxPackage":
        """Open a DOCX archive and populate XML trees and related assets."""
        with zipfile.ZipFile(docx_path) as docx_zip:
            parts = {name: docx_zip.read(name) for name in docx_zip.namelist()}

        LOGGER.debug("Loaded %d parts from %s", len(parts), docx_path.name)

        package = cls(raw_parts=parts)
        package._initialize_caches()
        return package

    # ------------------------------------------------------------------
    # Public helpers
    def require_document_xml(self) -> ET.ElementTree:
        if self.document_xml is None:
            raise ValueError("Primary document part missing from package")
        return self.document_xml

    def require_styles_xml(self) -> ET.ElementTree:
        if self.styles_xml is None:
            raise ValueError("Styles part missing from package")
        return self.styles_xml

    def get_numbering_xml(self) -> Optional[ET.ElementTree]:
        return self.numbering_xml

    def get_xml_part(self, name: str) -> Optional[ET.ElementTree]:
        if name in self.xml_cache:
            return self.xml_cache[name]
        data = self.raw_parts.get(name)
        if data is None:
            return None
        tree = parse_xml(data)
        self.xml_cache[name] = tree
        return tree

    # ------------------------------------------------------------------
    # Internal bootstrap
    def _initialize_caches(self) -> None:
        self.content_types_xml = self._parse_required("[Content_Types].xml")
        self.package_rels_xml = self._parse_optional(PACKAGE_REL_PATH)

        self.document_xml = self._parse_required(DOCUMENT_XML_PATH)
        self.styles_xml = self._parse_required(STYLES_XML_PATH)
        self.numbering_xml = self._parse_optional(NUMBERING_XML_PATH)
        self.document_rels_xml = self._parse_optional(DOCUMENT_RELS_PATH)
        self.footnotes_xml = self._parse_optional(FOOTNOTES_XML_PATH)
        self.endnotes_xml = self._parse_optional(ENDNOTES_XML_PATH)
        self.comments_xml = self._parse_optional(COMMENTS_XML_PATH)
        self.settings_xml = self._parse_optional(SETTINGS_XML_PATH)
        self.glossary_xml = self._parse_optional(GLOSSARY_XML_PATH)

        self.core_properties_xml = self._parse_optional(CORE_PROPS_PATH)
        self.app_properties_xml = self._parse_optional(APP_PROPS_PATH)
        self.custom_properties_xml = self._parse_optional(CUSTOM_PROPS_PATH)

        self.headers = self._collect_prefixed("word/header")
        self.footers = self._collect_prefixed("word/footer")
        self.theme_parts = self._collect_prefixed("word/theme/")

        self.relationships = Relationships.from_package(self.raw_parts)
        self.media = {
            name: data for name, data in self.raw_parts.items() if name.startswith("word/media/")
        }

    def _parse_required(self, name: str) -> ET.ElementTree:
        tree = self._parse_optional(name)
        if tree is None:
            raise KeyError(f"Required DOCX part missing: {name}")
        return tree

    def _parse_optional(self, name: str) -> Optional[ET.ElementTree]:
        if name in self.xml_cache:
            return self.xml_cache[name]
        data = self.raw_parts.get(name)
        if data is None:
            return None
        tree = parse_xml(data)
        self.xml_cache[name] = tree
        return tree

    def _collect_prefixed(self, prefix: str) -> Dict[str, ET.ElementTree]:
        collected: Dict[str, ET.ElementTree] = {}
        for name, data in self.raw_parts.items():
            if not name.startswith(prefix) or not name.endswith(".xml"):
                continue
            tree = parse_xml(data)
            self.xml_cache[name] = tree
            collected[name] = tree
        return collected

