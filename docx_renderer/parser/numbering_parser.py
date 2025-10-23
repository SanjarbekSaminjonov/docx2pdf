"""Parse numbering.xml into numbering model definitions."""
from __future__ import annotations

from copy import deepcopy
from typing import Dict, Optional
from xml.etree import ElementTree as ET

from docx_renderer.model.numbering_model import (
    AbstractNumberingDefinition,
    NumberingCatalog,
    NumberingInstance,
    NumberingLevel,
    NumberingOverride,
)
from docx_renderer.utils.xml_utils import Namespaces


class NumberingParser:
    """Parser for numbering definitions defined in numbering.xml."""

    def __init__(self, numbering_xml: Optional[ET.ElementTree]) -> None:
        self._numbering_xml = numbering_xml

    def parse(self) -> NumberingCatalog:
        if self._numbering_xml is None:
            return NumberingCatalog(abstracts={}, instances={})

        root = self._numbering_xml.getroot()
        abstracts = self._parse_abstract_nums(root)
        instances = self._parse_nums(root, abstracts)
        return NumberingCatalog(abstracts=abstracts, instances=instances)

    # ------------------------------------------------------------------
    def _parse_abstract_nums(self, root: ET.Element) -> Dict[int, AbstractNumberingDefinition]:
        abstracts: Dict[int, AbstractNumberingDefinition] = {}
        for abstract_el in root.findall("w:abstractNum", Namespaces.WORD):
            abstract_id = self._get_int_attr(abstract_el, None, "w:abstractNumId")
            if abstract_id is None:
                continue
            multi_level_type = self._get_attr(abstract_el, "w:multiLevelType", "w:val")
            name = self._get_attr(abstract_el, "w:name", "w:val")
            style_link = self._get_attr(abstract_el, "w:styleLink", "w:val")
            levels = self._parse_levels(abstract_el)
            raw = self._serialize_node(abstract_el)
            abstracts[abstract_id] = AbstractNumberingDefinition(
                abstract_num_id=abstract_id,
                multi_level_type=multi_level_type,
                name=name,
                style_link=style_link,
                levels=levels,
                raw_properties=raw,
            )
        return abstracts

    def _parse_levels(self, abstract_el: ET.Element) -> Dict[int, NumberingLevel]:
        levels: Dict[int, NumberingLevel] = {}
        for lvl_el in abstract_el.findall("w:lvl", Namespaces.WORD):
            level_index = self._get_int_attr(lvl_el, None, "w:ilvl")
            if level_index is None:
                continue
            start = self._get_int_attr(lvl_el, "w:start", "w:val")
            num_format = self._get_attr(lvl_el, "w:numFmt", "w:val")
            level_text = self._get_attr(lvl_el, "w:lvlText", "w:val")
            alignment = self._get_attr(lvl_el, "w:lvlJc", "w:val")
            is_legal_raw = self._get_attr(lvl_el, "w:isLgl", "w:val")
            is_legal = None
            if is_legal_raw is not None:
                is_legal = is_legal_raw == "1"
            raw = self._serialize_node(lvl_el)
            p_pr = self._collect_child_block(lvl_el, "w:pPr")
            r_pr = self._collect_child_block(lvl_el, "w:rPr")
            levels[level_index] = NumberingLevel(
                level_index=level_index,
                start=start,
                num_format=num_format,
                level_text=level_text,
                alignment=alignment,
                is_legal=is_legal,
                raw_properties=raw,
                paragraph_properties=p_pr,
                run_properties=r_pr,
            )
        return levels

    def _parse_nums(
        self,
        root: ET.Element,
        abstracts: Dict[int, AbstractNumberingDefinition],
    ) -> Dict[int, NumberingInstance]:
        instances: Dict[int, NumberingInstance] = {}
        for num_el in root.findall("w:num", Namespaces.WORD):
            num_id = self._get_int_attr(num_el, None, "w:numId")
            if num_id is None:
                continue
            abstract_num_id = self._get_int_attr(num_el, "w:abstractNumId", "w:val")
            if abstract_num_id is None:
                continue
            overrides = self._parse_overrides(num_el)
            instances[num_id] = NumberingInstance(
                num_id=num_id,
                abstract_num_id=abstract_num_id,
                overrides=overrides,
            )
        return instances

    def _parse_overrides(self, num_el: ET.Element) -> Dict[int, NumberingOverride]:
        overrides: Dict[int, NumberingOverride] = {}
        for override_el in num_el.findall("w:lvlOverride", Namespaces.WORD):
            level_index = self._get_int_attr(override_el, None, "w:ilvl")
            if level_index is None:
                continue
            start_override = self._get_int_attr(override_el, "w:startOverride", "w:val")
            raw = self._serialize_node(override_el)
            overrides[level_index] = NumberingOverride(
                level_index=level_index,
                start_override=start_override,
                raw_properties=raw,
            )
        return overrides

    # ------------------------------------------------------------------
    def _get_attr(self, element: ET.Element, child_name: Optional[str], attr_name: str) -> Optional[str]:
        target = element.find(child_name, Namespaces.WORD) if child_name else element
        if target is None:
            return None
        attr_key = self._qualify(attr_name)
        return target.attrib.get(attr_key)

    def _get_int_attr(self, element: ET.Element, child_name: Optional[str], attr_name: str) -> Optional[int]:
        value = self._get_attr(element, child_name, attr_name)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    def _collect_child_block(self, element: ET.Element, child_name: str):
        child = element.find(child_name, Namespaces.WORD)
        if child is None:
            return []
        return [self._serialize_node(node) for node in list(child)]

    def _serialize_node(self, node: ET.Element) -> Dict[str, object]:
        data: Dict[str, object] = {
            "tag": node.tag,
            "attributes": deepcopy(node.attrib),
        }
        if node.text and node.text.strip():
            data["text"] = node.text
        children = [self._serialize_node(child) for child in list(node)]
        if children:
            data["children"] = children
        return data

    def _qualify(self, attr_name: str) -> str:
        prefix, local = attr_name.split(":", 1)
        namespace = Namespaces.WORD[prefix]
        return f"{{{namespace}}}{local}"
