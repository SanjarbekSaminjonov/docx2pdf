"""Helper functions to work with XML namespaces and parsing."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional
from xml.etree import ElementTree as ET


@dataclass(frozen=True)
class Namespaces:
    """Common OpenXML namespace prefixes used across parsers."""

    WORD: Dict[str, str] = None  # type: ignore[assignment]
    RELS: Dict[str, str] = None  # type: ignore[assignment]
    DRAWING: Dict[str, str] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:  # pragma: no cover
        raise RuntimeError("Namespaces should not be instantiated")


Namespaces.WORD = {  # type: ignore[attr-defined]
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
Namespaces.RELS = {  # type: ignore[attr-defined]
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
Namespaces.DRAWING = {  # type: ignore[attr-defined]
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


def parse_xml(data: bytes) -> ET.ElementTree:
    """Parse XML from raw bytes with sane defaults."""
    return ET.ElementTree(ET.fromstring(data))


def find_text(element: ET.Element, xpath: str, namespaces: Optional[Dict[str, str]] = None) -> Optional[str]:
    """Return trimmed text from the first element that matches the xpath."""
    found = element.find(xpath, namespaces or {})
    if found is None or found.text is None:
        return None
    return found.text.strip()
