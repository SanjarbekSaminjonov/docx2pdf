"""Utilities for reading Open Packaging Convention relationship parts."""
from __future__ import annotations

import posixpath
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import Dict, Iterable, List, Mapping, Optional, Tuple
from xml.etree import ElementTree as ET

from docx_renderer.utils.xml_utils import Namespaces, parse_xml

WORD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

RELTYPE_IMAGE = f"{WORD_REL_NS}/image"
RELTYPE_HYPERLINK = f"{WORD_REL_NS}/hyperlink"
RELTYPE_HEADER = f"{WORD_REL_NS}/header"
RELTYPE_FOOTER = f"{WORD_REL_NS}/footer"
RELTYPE_NUMBERING = f"{WORD_REL_NS}/numbering"

MAIN_DOCUMENT_PART = "word/document.xml"


@dataclass(frozen=True)
class Relationship:
    """Represents a single OPC relationship."""

    source_part: str
    r_id: str
    target: str
    rel_type: str
    is_external: bool = False
    resolved_target: Optional[str] = None


@dataclass(frozen=True)
class DocumentRelationshipSummary:
    """Categorized relationship buckets for the main document part."""

    media: Dict[str, Relationship]
    headers: Dict[str, Relationship]
    footers: Dict[str, Relationship]
    numbering: Dict[str, Relationship]
    hyperlinks: Dict[str, Relationship]


class Relationships:
    """Aggregated relationship mappings for the DOCX package."""

    def __init__(self, relationships: Dict[str, Dict[str, Relationship]]) -> None:
        self._by_source = relationships

    @classmethod
    def from_package(cls, parts: Mapping[str, bytes]) -> "Relationships":
        """Collect relationships from all known .rels parts within the package."""
        by_source: Dict[str, Dict[str, Relationship]] = {}
        for name, payload in parts.items():
            if not name.endswith(".rels"):
                continue
            source, base_dir = cls._source_and_base_from_rel_part(name)
            tree = parse_xml(payload)
            parsed = cls._parse_relationship_part(source, base_dir, tree)
            if parsed:
                by_source[source] = parsed
        return cls(by_source)

    def find(self, part_name: str, r_id: str) -> Optional[Relationship]:
        """Return a relationship by part and id if present."""
        source = self._normalize_source(part_name)
        return self._by_source.get(source, {}).get(r_id)

    def for_source(self, part_name: str) -> Dict[str, Relationship]:
        """Return all relationships for a given source part."""
        source = self._normalize_source(part_name)
        rels = self._by_source.get(source, {})
        return dict(rels)

    def iter_all(self) -> Iterable[Relationship]:
        """Iterate over all registered relationships."""
        for rels in self._by_source.values():
            yield from rels.values()

    def document_summary(self) -> DocumentRelationshipSummary:
        """Return categorized relationships for the main document part."""
        doc_rels = self.for_source(MAIN_DOCUMENT_PART)
        return DocumentRelationshipSummary(
            media=self._filter_by_type(doc_rels, RELTYPE_IMAGE),
            headers=self._filter_by_type(doc_rels, RELTYPE_HEADER),
            footers=self._filter_by_type(doc_rels, RELTYPE_FOOTER),
            numbering=self._filter_by_type(doc_rels, RELTYPE_NUMBERING),
            hyperlinks=self._filter_by_type(doc_rels, RELTYPE_HYPERLINK),
        )
    
    def get_targets_by_type(self, rel_types: List[str]) -> Dict[str, str]:
        """Get relationship ID to target mapping for given relationship types."""
        result = {}
        for rel in self.iter_all():
            if rel.rel_type in rel_types:
                target = rel.resolved_target or rel.target
                result[rel.r_id] = target
        return result

    @classmethod
    def _parse_relationship_part(
        cls, source_part: str, base_dir: PurePosixPath, tree: ET.ElementTree
    ) -> Dict[str, Relationship]:
        result: Dict[str, Relationship] = {}
        for rel_el in tree.findall(".//rel:Relationship", Namespaces.RELS):
            r_id = rel_el.attrib["Id"]
            target = rel_el.attrib.get("Target", "")
            rel_type = rel_el.attrib.get("Type", "")
            is_external = rel_el.attrib.get("TargetMode") == "External"
            resolved_target = cls._resolve_target_path(base_dir, target, is_external)
            result[r_id] = Relationship(
                source_part=source_part,
                r_id=r_id,
                target=target,
                rel_type=rel_type,
                is_external=is_external,
                resolved_target=resolved_target,
            )
        return result

    @staticmethod
    def _source_and_base_from_rel_part(rel_part: str) -> Tuple[str, PurePosixPath]:
        rel_path = PurePosixPath(rel_part)
        base_dir = rel_path.parent
        if rel_part == "_rels/.rels":
            return "", base_dir
        if "/_rels/" in rel_part:
            folder, suffix = rel_part.split("/_rels/", 1)
            base = suffix[:-5]
            return f"{folder}/{base}", base_dir
        if rel_part.startswith("_rels/"):
            base = rel_part[len("_rels/") : -5]
            return base, base_dir
        return rel_part[:-5], base_dir

    @staticmethod
    def _resolve_target_path(base_dir: PurePosixPath, target: str, is_external: bool) -> Optional[str]:
        if not target:
            return None
        if is_external:
            return target
        resolved = base_dir.joinpath(target)
        normalized = posixpath.normpath(resolved.as_posix())
        normalized = normalized.replace("/_rels/", "/")
        if normalized.startswith("_rels/"):
            normalized = normalized[len("_rels/") :]
        return normalized

    @staticmethod
    def _filter_by_type(rels: Mapping[str, Relationship], rel_type: str) -> Dict[str, Relationship]:
        return {r_id: rel for r_id, rel in rels.items() if rel.rel_type == rel_type}

    @classmethod
    def _normalize_source(cls, part_name: str) -> str:
        if part_name.endswith(".rels"):
            source, _ = cls._source_and_base_from_rel_part(part_name)
            return source
        return part_name
