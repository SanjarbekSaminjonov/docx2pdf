"""Extract style definitions from styles.xml and produce a catalog."""
from __future__ import annotations

from typing import Dict, Optional
from xml.etree import ElementTree as ET

from docx_renderer.model.style_model import StyleDefinition, StylesCatalog
from docx_renderer.utils.xml_utils import Namespaces


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
            style_id = style_el.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
            if not style_id:
                continue
            style_type = style_el.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "paragraph")
            name = self._get_attr(style_el, "w:name", "w:val")
            based_on = self._get_attr(style_el, "w:basedOn", "w:val")
            next_style = self._get_attr(style_el, "w:next", "w:val")
            properties = self._extract_properties(style_el)
            styles[style_id] = StyleDefinition(
                style_id=style_id,
                style_type=style_type,
                name=name,
                properties=properties,
                based_on=based_on,
                next_style=next_style,
            )
        return styles

    def _extract_properties(self, style_el: ET.Element) -> Dict[str, object]:
        # Placeholder for detailed property extraction; to be expanded incrementally.
        properties: Dict[str, object] = {}
        rpr = style_el.find("w:rPr", Namespaces.WORD)
        if rpr is not None and rpr.find("w:b", Namespaces.WORD) is not None:
            properties["bold"] = True
        if rpr is not None and rpr.find("w:i", Namespaces.WORD) is not None:
            properties["italic"] = True
        return properties

    def _get_attr(self, element: ET.Element, child_name: str, attr_name: str) -> Optional[str]:
        child = element.find(child_name, Namespaces.WORD)
        if child is None:
            return None
        attr_key = self._qualify(attr_name)
        return child.attrib.get(attr_key)

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
            merged_props = dict(style.properties)
            if style.based_on and style.based_on in raw_styles:
                parent = resolve(style.based_on, stack)
                parent_props = dict(parent.properties)
                parent_props.update(merged_props)
                merged_props = parent_props
            resolved_style = StyleDefinition(
                style_id=style.style_id,
                style_type=style.style_type,
                name=style.name,
                properties=merged_props,
                based_on=style.based_on,
                next_style=style.next_style,
            )
            resolved[style_id] = resolved_style
            stack.pop()
            return resolved_style

        for style_id in raw_styles:
            resolve(style_id)
        return resolved
