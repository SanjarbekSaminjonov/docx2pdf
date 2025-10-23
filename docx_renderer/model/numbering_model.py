"""Numbering model captures list definitions extracted from numbering.xml."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass(slots=True)
class NumberingLevel:
    """Defines numbering behavior for a specific indentation level."""

    level_index: int
    start: Optional[int]
    num_format: Optional[str]
    level_text: Optional[str]
    alignment: Optional[str]
    is_legal: Optional[bool]
    raw_properties: Dict[str, object] = field(default_factory=dict)
    paragraph_properties: List[Dict[str, object]] = field(default_factory=list)
    run_properties: List[Dict[str, object]] = field(default_factory=list)


@dataclass(slots=True)
class NumberingOverride:
    """Overrides applied to a numbering instance for specific levels."""

    level_index: int
    start_override: Optional[int]
    raw_properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class AbstractNumberingDefinition:
    """Template describing multi-level numbering behavior."""

    abstract_num_id: int
    multi_level_type: Optional[str]
    name: Optional[str]
    style_link: Optional[str]
    levels: Dict[int, NumberingLevel] = field(default_factory=dict)
    raw_properties: Dict[str, object] = field(default_factory=dict)


@dataclass(slots=True)
class NumberingInstance:
    """Concrete numbering instance bound to an abstract definition."""

    num_id: int
    abstract_num_id: int
    overrides: Dict[int, NumberingOverride] = field(default_factory=dict)


@dataclass(slots=True)
class NumberingCatalog:
    """Collection of abstract definitions and concrete numbering instances."""

    abstracts: Dict[int, AbstractNumberingDefinition]
    instances: Dict[int, NumberingInstance]

    def get_abstract(self, abstract_num_id: Optional[int]) -> Optional[AbstractNumberingDefinition]:
        if abstract_num_id is None:
            return None
        return self.abstracts.get(abstract_num_id)

    def get_instance(self, num_id: Optional[int]) -> Optional[NumberingInstance]:
        if num_id is None:
            return None
        return self.instances.get(num_id)
