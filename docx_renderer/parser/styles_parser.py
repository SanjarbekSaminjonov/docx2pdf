"""Extract style definitions from styles.xml and produce a catalog."""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree as ET

from docx_renderer.model.style_model import StyleDefinition, StylesCatalog
from docx_renderer.utils.xml_utils import Namespaces

PropertyNode = Dict[str, Any]


class StylesParser:
    """Parse Word styles and resolve inheritance."""

    def __init__(self, styles_xml: ET.ElementTree, numbering_xml: Optional[ET.ElementTree] = None) -> None:
        self._styles_xml = styles_xml
        self._numbering_xml = numbering_xml

    def parse(self) -> StylesCatalog:
        """Parse the XML tree and return a resolved catalog."""
        raw_styles = self._collect_styles()
        resolved = self._resolve_inheritance(raw_styles)
        return StylesCatalog(resolved)

    def _collect_styles(self) -> Dict[str, StyleDefinition]:
        styles: Dict[str, StyleDefinition] = {}
        root = self._styles_xml.getroot()
        for style_el in root.findall("w:style", Namespaces.WORD):
            style_id = style_el.attrib.get(self._qualify("w:styleId"))
            if not style_id:
                continue
            style_type = style_el.attrib.get(self._qualify("w:type"), "paragraph")
            name = self._get_attr(style_el, "w:name", "w:val")
            based_on = self._get_attr(style_el, "w:basedOn", "w:val")
            next_style = self._get_attr(style_el, "w:next", "w:val")
            linked_style = self._get_attr(style_el, "w:link", "w:val")
            aliases = self._get_attr(style_el, "w:alias", "w:val")
            ui_priority = self._get_int_attr(style_el, "w:uiPriority", "w:val")
            is_default = style_el.attrib.get(self._qualify("w:default")) == "1"
            is_primary = style_el.find("w:qFormat", Namespaces.WORD) is not None
            properties = self._extract_properties(style_el)
            styles[style_id] = StyleDefinition(
                style_id=style_id,
                style_type=style_type,
                name=name,
                properties=properties,
                based_on=based_on,
                next_style=next_style,
                linked_style=linked_style,
                is_default=is_default,
                ui_priority=ui_priority,
                is_primary=is_primary,
                aliases=aliases,
            )
        return styles

    def _extract_properties(self, style_el: ET.Element) -> Dict[str, List[PropertyNode]]:
        properties: Dict[str, List[PropertyNode]] = {}
        for child_tag, key in (
            ("w:rPr", "rPr"),
            ("w:pPr", "pPr"),
            ("w:tblPr", "tblPr"),
            ("w:tblStylePr", "tblStylePr"),
            ("w:numPr", "numPr"),
        ):
            child = style_el.find(child_tag, Namespaces.WORD)
            if child is not None:
                properties[key] = self._serialize_property_block(child)
        return properties

    def _get_attr(self, element: ET.Element, child_name: str, attr_name: str) -> Optional[str]:
        child = element.find(child_name, Namespaces.WORD)
        if child is None:
            return None
        attr_key = self._qualify(attr_name)
        return child.attrib.get(attr_key)

    def _get_int_attr(self, element: ET.Element, child_name: str, attr_name: str) -> Optional[int]:
        value = self._get_attr(element, child_name, attr_name)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    def _qualify(self, attr_name: str) -> str:
        prefix, local = attr_name.split(":", 1)
        namespace = Namespaces.WORD[prefix]
        return f"{{{namespace}}}{local}"

    def _resolve_inheritance(self, raw_styles: Dict[str, StyleDefinition]) -> Dict[str, StyleDefinition]:
        resolved: Dict[str, StyleDefinition] = {}

        def resolve(style_id: str, stack: Optional[list[str]] = None) -> StyleDefinition:
            if style_id in resolved:
                return resolved[style_id]
            if stack is None:
                stack = []
            if style_id in stack:
                return raw_styles[style_id]
            stack.append(style_id)
            style = raw_styles[style_id]
            merged_props = deepcopy(style.properties)
            parent_style = None
            if style.based_on and style.based_on in raw_styles:
                parent_style = resolve(style.based_on, stack)
                merged_props = self._merge_properties(parent_style.properties, merged_props)
            resolved_style = StyleDefinition(
                style_id=style.style_id,
                style_type=style.style_type,
                name=style.name,
                properties=merged_props,
                based_on=style.based_on,
                next_style=style.next_style,
                linked_style=style.linked_style or (parent_style.linked_style if parent_style else None),
                is_default=style.is_default or (parent_style.is_default if parent_style else False),
                ui_priority=style.ui_priority if style.ui_priority is not None else (parent_style.ui_priority if parent_style else None),
                is_primary=style.is_primary or (parent_style.is_primary if parent_style else False),
                aliases=style.aliases or (parent_style.aliases if parent_style else None),
            )
            resolved[style_id] = resolved_style
            stack.pop()
            return resolved_style

        for style_id in raw_styles:
            resolve(style_id)
        return resolved

    def _merge_properties(
        self,
        parent_props: Dict[str, List[PropertyNode]],
        child_props: Dict[str, List[PropertyNode]],
    ) -> Dict[str, List[PropertyNode]]:
        merged = {key: deepcopy(value) for key, value in parent_props.items()}
        for block, entries in child_props.items():
            if block in merged:
                merged[block] = self._merge_property_entries(merged[block], entries)
            else:
                merged[block] = deepcopy(entries)
        return merged

    def _merge_property_entries(
        self, parent_entries: List[PropertyNode], child_entries: List[PropertyNode]
    ) -> List[PropertyNode]:
        merged: List[PropertyNode] = [deepcopy(entry) for entry in parent_entries]
        merged.extend(deepcopy(entry) for entry in child_entries)
        return merged

    def _serialize_property_block(self, element: ET.Element) -> List[PropertyNode]:
        return [self._serialize_node(child) for child in list(element)]

    def _serialize_node(self, node: ET.Element) -> PropertyNode:
        data: PropertyNode = {
            "tag": node.tag,
            "attributes": {attr: value for attr, value in node.attrib.items()},
        }
        if node.text and node.text.strip():
            data["text"] = node.text
        children = [self._serialize_node(child) for child in list(node)]
        if children:
            data["children"] = children
        return data
