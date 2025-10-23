"""Style model captures Word style definitions in a normalized form."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Mapping, Optional


@dataclass(slots=True)
class StyleDefinition:
    """Full style information after resolving inheritance."""

    style_id: str
    style_type: str
    name: Optional[str]
    properties: Dict[str, object] = field(default_factory=dict)
    based_on: Optional[str] = None
    next_style: Optional[str] = None
    linked_style: Optional[str] = None
    is_default: bool = False
    ui_priority: Optional[int] = None
    is_primary: bool = False
    aliases: Optional[str] = None


class StylesCatalog:
    """Collection of resolved styles keyed by identifier."""

    def __init__(self, styles: Mapping[str, StyleDefinition]):
        self._styles = dict(styles)

    def get(self, style_id: Optional[str]) -> Optional[StyleDefinition]:
        """Return the resolved style definition given its identifier."""
        if style_id is None:
            return None
        return self._styles.get(style_id)

    def all(self) -> Mapping[str, StyleDefinition]:
        """Return read-only view of resolved styles."""
        return dict(self._styles)

    def default_for(self, style_type: str) -> Optional[StyleDefinition]:
        """Return the default style for the given style type if defined."""
        for style in self._styles.values():
            if style.style_type == style_type and style.is_default:
                return style
        return None
