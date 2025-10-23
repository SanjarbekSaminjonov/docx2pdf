"""Utilities for reading Open Packaging Convention relationship parts."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, Optional
from xml.etree import ElementTree as ET

from docx_renderer.utils.xml_utils import Namespaces, parse_xml


@dataclass(frozen=True)
class Relationship:
    """Represents a single OPC relationship."""

    r_id: str
    target: str
    rel_type: str


class Relationships:
    """Aggregated relationship mappings for the DOCX package."""

    def __init__(self, relationships: Dict[str, Relationship]) -> None:
        self._relationships = relationships

    @classmethod
    def from_package(cls, parts: Dict[str, bytes]) -> "Relationships":
        """Collect relationships from all known .rels parts within the package."""
        rels: Dict[str, Relationship] = {}
        for name, payload in parts.items():
            if not name.endswith(".rels"):
                continue
            tree = parse_xml(payload)
            rels.update(cls._parse_relationship_part(name, tree))
        return cls(rels)

    @classmethod
    def _parse_relationship_part(cls, part_name: str, tree: ET.ElementTree) -> Dict[str, Relationship]:
        result: Dict[str, Relationship] = {}
        for rel_el in tree.findall(".//rel:Relationship", Namespaces.RELS):
            r_id = rel_el.attrib["Id"]
            target = rel_el.attrib.get("Target", "")
            rel_type = rel_el.attrib.get("Type", "")
            key = f"{part_name}:{r_id}"
            result[key] = Relationship(r_id=r_id, target=target, rel_type=rel_type)
        return result

    def find(self, part_name: str, r_id: str) -> Optional[Relationship]:
        """Return a relationship by part and id if present."""
        return self._relationships.get(f"{part_name}:{r_id}")

    def iter_all(self) -> Iterable[Relationship]:
        """Iterate over all registered relationships."""
        return self._relationships.values()
