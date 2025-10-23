"""Parser for headers and footers linked to section properties."""
from __future__ import annotations

from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

from docx_renderer.model.elements import (
    BlockElement,
    DocumentSection,
    DocumentTree,
    HeaderFooterContent,
    ParagraphElement,
    RunFragment,
    SectionProperties,
    TableElement,
)
from docx_renderer.model.numbering_model import NumberingCatalog
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.utils.logger import get_logger
from docx_renderer.utils.xml_utils import Namespaces

LOGGER = get_logger(__name__)


class SectionParser:
    """Parse document sections including headers, footers, and section properties."""

    def __init__(self, package: DocxPackage, styles: StylesCatalog, numbering: NumberingCatalog) -> None:
        self._package = package
        self._styles = styles
        self._numbering = numbering

    def parse_sections(self, blocks: List[BlockElement]) -> DocumentTree:
        """Split blocks into sections based on section breaks and parse section properties."""
        sections: List[DocumentSection] = []
        current_blocks: List[BlockElement] = []
        
        # Get document-level section properties (final sectPr in document.xml)
        doc_tree = self._package.require_document_xml()
        body = doc_tree.getroot().find("w:body", Namespaces.WORD)
        final_sect_pr = body.find("w:sectPr", Namespaces.WORD) if body is not None else None
        
        # Process blocks and look for section breaks
        for block in blocks:
            if isinstance(block, ParagraphElement):
                # Check for section break in paragraph properties
                sect_pr = self._find_section_break(block)
                if sect_pr is not None:
                    # End current section
                    section_props = self._parse_section_properties(sect_pr)
                    sections.append(DocumentSection(blocks=current_blocks, properties=section_props))
                    current_blocks = []
                else:
                    current_blocks.append(block)
            else:
                current_blocks.append(block)
        
        # Handle final section
        if current_blocks or not sections:
            final_props = self._parse_section_properties(final_sect_pr) if final_sect_pr is not None else SectionProperties()
            sections.append(DocumentSection(blocks=current_blocks, properties=final_props))
        
        return DocumentTree(sections=sections)

    def _find_section_break(self, paragraph: ParagraphElement) -> Optional[ET.Element]:
        """Check if paragraph contains a section break."""
        if not isinstance(paragraph.properties.get("children"), list):
            return None
        
        for child in paragraph.properties["children"]:
            if isinstance(child, dict) and child.get("tag", "").endswith("sectPr"):
                # Convert back to ET.Element for processing
                return self._dict_to_element(child)
        return None

    def _parse_section_properties(self, sect_pr: Optional[ET.Element]) -> SectionProperties:
        """Parse section properties including page setup and header/footer references."""
        if sect_pr is None:
            return SectionProperties()

        # Parse page size
        pg_sz = sect_pr.find("w:pgSz", Namespaces.WORD)
        page_width = self._get_int_attr(pg_sz, None, "w:w") if pg_sz is not None else None
        page_height = self._get_int_attr(pg_sz, None, "w:h") if pg_sz is not None else None
        orientation = self._get_attr(pg_sz, None, "w:orient") if pg_sz is not None else None

        # Parse page margins
        pg_mar = sect_pr.find("w:pgMar", Namespaces.WORD)
        margin_top = self._get_int_attr(pg_mar, None, "w:top") if pg_mar is not None else None
        margin_bottom = self._get_int_attr(pg_mar, None, "w:bottom") if pg_mar is not None else None
        margin_left = self._get_int_attr(pg_mar, None, "w:left") if pg_mar is not None else None
        margin_right = self._get_int_attr(pg_mar, None, "w:right") if pg_mar is not None else None
        margin_header = self._get_int_attr(pg_mar, None, "w:header") if pg_mar is not None else None
        margin_footer = self._get_int_attr(pg_mar, None, "w:footer") if pg_mar is not None else None

        # Parse title page setting
        title_pg = sect_pr.find("w:titlePg", Namespaces.WORD)
        title_page = title_pg is not None

        # Parse header/footer references
        header_default = self._parse_header_footer_ref(sect_pr, "w:headerReference", "default")
        header_first = self._parse_header_footer_ref(sect_pr, "w:headerReference", "first")
        header_even = self._parse_header_footer_ref(sect_pr, "w:headerReference", "even")
        footer_default = self._parse_header_footer_ref(sect_pr, "w:footerReference", "default")
        footer_first = self._parse_header_footer_ref(sect_pr, "w:footerReference", "first")
        footer_even = self._parse_header_footer_ref(sect_pr, "w:footerReference", "even")

        # Serialize raw properties
        raw_properties = self._serialize_node(sect_pr)

        return SectionProperties(
            page_width=page_width,
            page_height=page_height,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_header=margin_header,
            margin_footer=margin_footer,
            orientation=orientation,
            header_default=header_default,
            header_first=header_first,
            header_even=header_even,
            footer_default=footer_default,
            footer_first=footer_first,
            footer_even=footer_even,
            title_page=title_page,
            raw_properties=raw_properties,
        )

    def _parse_header_footer_ref(self, sect_pr: ET.Element, ref_type: str, hf_type: str) -> Optional[HeaderFooterContent]:
        """Parse header or footer reference and load content."""
        for ref_el in sect_pr.findall(ref_type, Namespaces.WORD):
            ref_type_attr = self._get_attr(ref_el, None, "w:type")
            if ref_type_attr == hf_type or (hf_type == "default" and ref_type_attr is None):
                r_id = self._get_attr(ref_el, None, "r:id")
                if r_id:
                    return self._load_header_footer_content(r_id)
        return None

    def _load_header_footer_content(self, r_id: str) -> Optional[HeaderFooterContent]:
        """Load header/footer content from relationship."""
        rel = self._package.relationships.find("word/document.xml", r_id)
        if not rel or not rel.resolved_target:
            return None

        # Get header/footer XML
        header_footer_xml = self._package.get_xml_part(rel.resolved_target)
        if not header_footer_xml:
            return None

        # Parse content
        blocks = self._parse_header_footer_blocks(header_footer_xml)
        
        # Store raw XML for debugging
        raw_xml = ET.tostring(header_footer_xml.getroot(), encoding="unicode")

        return HeaderFooterContent(r_id=r_id, blocks=blocks, raw_xml=raw_xml)

    def _parse_header_footer_blocks(self, xml_tree: ET.ElementTree) -> List[BlockElement]:
        """Parse blocks within header/footer content."""
        blocks: List[BlockElement] = []
        root = xml_tree.getroot()

        for child in list(root):
            tag = self._strip_namespace(child.tag)
            if tag == "p":
                blocks.append(self._parse_paragraph(child))
            elif tag == "tbl":
                blocks.append(self._parse_table(child))
            else:
                LOGGER.debug("Skipping header/footer element: %s", tag)

        return blocks

    def _parse_paragraph(self, paragraph_el: ET.Element) -> ParagraphElement:
        """Simplified paragraph parsing for header/footer content."""
        runs: List[RunFragment] = []
        
        for run_el in paragraph_el.findall("w:r", Namespaces.WORD):
            text = ""
            for text_el in run_el.findall("w:t", Namespaces.WORD):
                if text_el.text:
                    text += text_el.text
            
            run_props = self._extract_run_properties(run_el)
            runs.append(RunFragment(text=text, properties=run_props))

        paragraph_props = self._extract_paragraph_properties(paragraph_el)
        style_id = self._get_style_id(paragraph_el)

        return ParagraphElement(
            runs=runs,
            style_id=style_id,
            properties=paragraph_props,
        )

    def _parse_table(self, table_el: ET.Element) -> TableElement:
        """Simplified table parsing for header/footer content."""
        from docx_renderer.model.elements import TableRow, TableCell
        
        rows = []
        for row_el in table_el.findall("w:tr", Namespaces.WORD):
            cells = []
            for cell_el in row_el.findall("w:tc", Namespaces.WORD):
                cell_content = []
                for p_el in cell_el.findall("w:p", Namespaces.WORD):
                    cell_content.append(self._parse_paragraph(p_el))
                
                cell_props = self._extract_cell_properties(cell_el)
                cells.append(TableCell(content=cell_content, properties=cell_props))
            
            row_props = self._extract_row_properties(row_el)
            rows.append(TableRow(cells=cells, properties=row_props))

        table_props = self._extract_table_properties(table_el)
        style_id = self._get_table_style_id(table_el)

        return TableElement(rows=rows, style_id=style_id, properties=table_props)

    # Helper methods for property extraction
    def _extract_paragraph_properties(self, paragraph_el: ET.Element) -> Dict[str, object]:
        ppr = paragraph_el.find("w:pPr", Namespaces.WORD)
        return self._serialize_node(ppr) if ppr is not None else {}

    def _extract_run_properties(self, run_el: ET.Element) -> Dict[str, object]:
        rpr = run_el.find("w:rPr", Namespaces.WORD)
        return self._serialize_node(rpr) if rpr is not None else {}

    def _extract_table_properties(self, table_el: ET.Element) -> Dict[str, object]:
        tbl_pr = table_el.find("w:tblPr", Namespaces.WORD)
        return self._serialize_node(tbl_pr) if tbl_pr is not None else {}

    def _extract_row_properties(self, row_el: ET.Element) -> Dict[str, object]:
        tr_pr = row_el.find("w:trPr", Namespaces.WORD)
        return self._serialize_node(tr_pr) if tr_pr is not None else {}

    def _extract_cell_properties(self, cell_el: ET.Element) -> Dict[str, object]:
        tc_pr = cell_el.find("w:tcPr", Namespaces.WORD)
        return self._serialize_node(tc_pr) if tc_pr is not None else {}

    def _get_style_id(self, element: ET.Element) -> Optional[str]:
        ppr = element.find("w:pPr", Namespaces.WORD)
        if ppr is None:
            return None
        style_el = ppr.find("w:pStyle", Namespaces.WORD)
        if style_el is None:
            return None
        return self._get_attr(style_el, None, "w:val")

    def _get_table_style_id(self, table_el: ET.Element) -> Optional[str]:
        tbl_pr = table_el.find("w:tblPr", Namespaces.WORD)
        if tbl_pr is None:
            return None
        style_el = tbl_pr.find("w:tblStyle", Namespaces.WORD)
        if style_el is None:
            return None
        return self._get_attr(style_el, None, "w:val")

    def _get_attr(self, element: Optional[ET.Element], child_name: Optional[str], attr_name: str) -> Optional[str]:
        if element is None:
            return None
        target = element.find(child_name, Namespaces.WORD) if child_name else element
        if target is None:
            return None
        attr_key = self._qualify(attr_name) if ":" in attr_name else attr_name
        return target.attrib.get(attr_key)

    def _get_int_attr(self, element: Optional[ET.Element], child_name: Optional[str], attr_name: str) -> Optional[int]:
        value = self._get_attr(element, child_name, attr_name)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    def _qualify(self, attr_name: str) -> str:
        prefix, local = attr_name.split(":", 1)
        if prefix == "w":
            namespace = Namespaces.WORD[prefix]
        elif prefix == "r":
            namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        else:
            return attr_name
        return f"{{{namespace}}}{local}"

    def _strip_namespace(self, tag: str) -> str:
        return tag.split("}", 1)[-1]

    def _serialize_node(self, node: ET.Element) -> Dict[str, object]:
        data = {
            "tag": node.tag,
            "attributes": dict(node.attrib),
        }
        if node.text and node.text.strip():
            data["text"] = node.text
        children = [self._serialize_node(child) for child in list(node)]
        if children:
            data["children"] = children
        return data

    def _dict_to_element(self, data: Dict[str, object]) -> ET.Element:
        """Convert serialized node back to ET.Element (simplified)."""
        tag = data["tag"]
        element = ET.Element(tag)
        
        if isinstance(data.get("attributes"), dict):
            element.attrib.update(data["attributes"])
        
        if "text" in data:
            element.text = str(data["text"])
            
        if isinstance(data.get("children"), list):
            for child_data in data["children"]:
                if isinstance(child_data, dict):
                    child_element = self._dict_to_element(child_data)
                    element.append(child_element)
        
        return element