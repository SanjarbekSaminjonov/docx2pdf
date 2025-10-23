"""Resolve media references from document relationships."""
from __future__ import annotations

from typing import Dict, Optional

from docx_renderer.parser.rels_parser import Relationships


class MediaResolver:
    """Maps relationship identifiers to actual media payloads."""

    def __init__(self, relationships: Relationships, media_map: Dict[str, bytes]) -> None:
        self._relationships = relationships
        self._media_map = media_map

    def resolve_image(self, part_name: str, r_id: str) -> Optional[bytes]:
        """Return binary data for an image referenced by a relationship id."""
        rel = self._relationships.find(part_name, r_id)
        if rel is None:
            return None
        target = f"word/{rel.target}" if not rel.target.startswith("word/") else rel.target
        return self._media_map.get(target)
