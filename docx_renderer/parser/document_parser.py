"""Parse document.xml into structured content blocks."""
from __future__ import annotations

from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

from docx_renderer.model.elements import (
    BlockElement,
    Bookmark,
    DocumentTree,
    DrawingReference,
    ImageElement,
    NumberingInfo,
    ParagraphElement,
    RunFragment,
    TableCell,
    TableElement,
    TableRow,
)
from docx_renderer.model.numbering_model import NumberingCatalog
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.media_extractor import MediaResolver
from docx_renderer.utils.logger import get_logger
from docx_renderer.utils.xml_utils import Namespaces

LOGGER = get_logger(__name__)


class DocumentParser:
    """Transforms Word body XML into model elements."""

    def __init__(self, package: DocxPackage, styles: StylesCatalog, numbering: NumberingCatalog) -> None:
        self._package = package
        self._styles = styles
        self._numbering = numbering
        self._media = MediaResolver(package.relationships, package.media)

    def parse(self) -> DocumentTree:
        """Parse the document body into high-level block elements."""
        document_tree = self._package.require_document_xml()
        root = document_tree.getroot()
        body = root.find("w:body", Namespaces.WORD)
        if body is None:
            LOGGER.warning("document.xml missing body element")
            return DocumentTree(sections=[])

        blocks = []
        for child in list(body):
            tag = self._strip_namespace(child.tag)
            if tag == "p":
                blocks.append(self._parse_paragraph(child))
            elif tag == "tbl":
                blocks.append(self._parse_table(child))
            elif tag == "sectPr":
                # Section properties influence layout; stored as metadata later.
                continue
            else:
                LOGGER.debug("Skipping unsupported element: %s", tag)
        
        # Parse sections with headers/footers
        from docx_renderer.parser.section_parser import SectionParser
        section_parser = SectionParser(self._package, self._styles, self._numbering)
        return section_parser.parse_sections(blocks)

    def _parse_paragraph(self, paragraph_el: ET.Element) -> ParagraphElement:
        runs: List[RunFragment] = []
        bookmarks: List[Bookmark] = []
        
        # Parse paragraph-level properties
        paragraph_props = self._extract_paragraph_properties(paragraph_el)
        numbering = self._extract_numbering_info(paragraph_el)
        style_id = self._get_style_id(paragraph_el)
        
        # Process all child elements in order to maintain document flow
        for child in list(paragraph_el):
            tag = self._strip_namespace(child.tag)
            
            if tag == "r":  # Run element
                runs.extend(self._parse_run(child))
            elif tag == "bookmarkStart":
                bookmark = self._parse_bookmark_start(child)
                if bookmark:
                    bookmarks.append(bookmark)
            elif tag == "hyperlink":
                runs.extend(self._parse_hyperlink(child))
            elif tag in ["pPr", "sectPr", "bookmarkEnd"]:
                # Already processed or not needed in runs
                continue
            else:
                LOGGER.debug("Skipping paragraph child element: %s", tag)
        
        return ParagraphElement(
            runs=runs,
            style_id=style_id,
            properties=paragraph_props,
            numbering=numbering,
            bookmarks=bookmarks,
        )

    def _parse_table(self, table_el: ET.Element) -> TableElement:
        rows: List[TableRow] = []
        table_props = self._extract_table_properties(table_el)
        style_id = self._get_table_style_id(table_el)
        
        for row_el in table_el.findall("w:tr", Namespaces.WORD):
            cells: List[TableCell] = []
            row_props = self._extract_row_properties(row_el)
            
            for cell_el in row_el.findall("w:tc", Namespaces.WORD):
                cell_content: List[BlockElement] = []
                cell_props = self._extract_cell_properties(cell_el)
                vertical_merge = self._extract_vertical_merge(cell_el)
                
                # Parse cell content (paragraphs, tables, etc.)
                for child in list(cell_el):
                    tag = self._strip_namespace(child.tag)
                    if tag == "p":
                        cell_content.append(self._parse_paragraph(child))
                    elif tag == "tbl":
                        cell_content.append(self._parse_table(child))
                    elif tag not in ["tcPr"]:  # Skip cell properties
                        LOGGER.debug("Skipping cell child element: %s", tag)
                
                if vertical_merge:
                    cell_props["vMerge"] = vertical_merge

                cells.append(TableCell(content=cell_content, properties=cell_props))
            
            rows.append(TableRow(cells=cells, properties=row_props))
        
        return TableElement(rows=rows, style_id=style_id, properties=table_props)

    def _get_style_id(self, element: ET.Element) -> str | None:
        ppr = element.find("w:pPr", Namespaces.WORD)
        if ppr is None:
            return None
        style_el = ppr.find("w:pStyle", Namespaces.WORD)
        if style_el is None:
            return None
        attr_key = self._qualify("w:val")
        return style_el.attrib.get(attr_key)

    def _strip_namespace(self, tag: str) -> str:
        return tag.split("}", 1)[-1]

    def _qualify(self, attr_name: str) -> str:
        prefix, local = attr_name.split(":", 1)
        namespace = Namespaces.WORD[prefix]
        return f"{{{namespace}}}{local}"

    # ------------------------------------------------------------------
    # Enhanced parsing methods
    
    def _parse_run(self, run_el: ET.Element) -> List[RunFragment]:
        """Parse a run element, handling text, fields, drawings, etc."""
        fragments: List[RunFragment] = []
        run_props = self._extract_run_properties(run_el)
        
        # Collect text and special elements from the run
        current_text = ""
        
        for child in list(run_el):
            tag = self._strip_namespace(child.tag)
            
            if tag == "t":  # Text element
                if child.text:
                    current_text += child.text
            elif tag == "tab":  # Tab character
                current_text += "\t"
            elif tag == "br":  # Line break
                current_text += "\n"
            elif tag == "cr":  # Carriage return
                current_text += "\r"
            elif tag == "drawing":  # Drawing/image
                # Flush current text
                if current_text:
                    fragments.append(RunFragment(text=current_text, properties=run_props))
                    current_text = ""
                # Parse drawing
                drawing = self._parse_drawing(child)
                if drawing:
                    fragments.append(RunFragment(
                        text="",
                        properties=run_props,
                        drawing=drawing
                    ))
            elif tag == "fldChar":  # Field character
                field_type = self._get_attr(child, None, "w:fldCharType")
                if field_type == "begin":
                    current_text += "{"
                elif field_type == "end":
                    current_text += "}"
            elif tag == "instrText":  # Field instruction text
                if child.text:
                    fragments.append(RunFragment(
                        text="",
                        properties=run_props,
                        field_code=child.text.strip()
                    ))
            elif tag == "footnoteReference":
                footnote_id = self._get_int_attr(child, None, "w:id")
                if footnote_id is not None:
                    fragments.append(RunFragment(
                        text="",
                        properties=run_props,
                        footnote_reference=footnote_id
                    ))
            elif tag == "endnoteReference":
                endnote_id = self._get_int_attr(child, None, "w:id")
                if endnote_id is not None:
                    fragments.append(RunFragment(
                        text="",
                        properties=run_props,
                        endnote_reference=endnote_id
                    ))
            elif tag in ["rPr", "noBreakHyphen", "softHyphen", "lastRenderedPageBreak"]:
                # Skip already processed or layout-only elements
                continue
            else:
                LOGGER.debug("Skipping run child element: %s", tag)
        
        # Add final text fragment if any
        if current_text or not fragments:
            fragments.append(RunFragment(text=current_text, properties=run_props))
        
        return fragments

    def _parse_hyperlink(self, hyperlink_el: ET.Element) -> List[RunFragment]:
        """Parse hyperlink element containing runs."""
        fragments: List[RunFragment] = []
        r_id = self._get_attr(hyperlink_el, None, "r:id")
        anchor = self._get_attr(hyperlink_el, None, "w:anchor")
        
        # Resolve hyperlink target
        target = None
        if r_id:
            rel = self._package.relationships.find("word/document.xml", r_id)
            if rel and rel.is_external:
                target = rel.target
        
        # Parse runs within hyperlink
        for run_el in hyperlink_el.findall("w:r", Namespaces.WORD):
            run_fragments = self._parse_run(run_el)
            for fragment in run_fragments:
                fragment.hyperlink_id = r_id
                fragment.hyperlink_anchor = anchor
                fragment.hyperlink_target = target
                fragments.append(fragment)
        
        return fragments

    def _parse_drawing(self, drawing_el: ET.Element) -> Optional[DrawingReference]:
        """Parse drawing element to extract image information."""
        # Try inline drawing first
        inline = drawing_el.find(".//wp:inline", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
        if inline is not None:
            return self._parse_inline_drawing(inline, True)
        
        # Try anchored drawing
        anchor = drawing_el.find(".//wp:anchor", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
        if anchor is not None:
            return self._parse_inline_drawing(anchor, False)
        
        return None

    def _parse_inline_drawing(self, drawing_el: ET.Element, is_inline: bool) -> Optional[DrawingReference]:
        """Parse inline or anchored drawing element."""
        # Extract dimensions
        extent = drawing_el.find(".//wp:extent", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
        width_emu = self._get_int_attr(extent, None, "cx") if extent is not None else None
        height_emu = self._get_int_attr(extent, None, "cy") if extent is not None else None
        
        # Extract description
        doc_pr = drawing_el.find(".//wp:docPr", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
        description = self._get_attr(doc_pr, None, "descr") if doc_pr is not None else None
        
        # Extract relationship ID for image
        blip = drawing_el.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
        r_id = self._get_attr(blip, None, "r:embed") if blip is not None else None
        
        if not r_id:
            return None
        
        # Resolve image target
        rel = self._package.relationships.find("word/document.xml", r_id)
        target = rel.resolved_target if rel else None
        
        # Get image data
        data = self._media.resolve_image("word/document.xml", r_id)
        
        return DrawingReference(
            r_id=r_id,
            target=target,
            description=description,
            width_emu=width_emu,
            height_emu=height_emu,
            inline=is_inline,
            data=data
        )

    def _parse_bookmark_start(self, bookmark_el: ET.Element) -> Optional[Bookmark]:
        """Parse bookmark start element."""
        bookmark_id = self._get_int_attr(bookmark_el, None, "w:id")
        name = self._get_attr(bookmark_el, None, "w:name")
        
        if bookmark_id is not None and name:
            return Bookmark(bookmark_id=bookmark_id, name=name)
        return None

    # ------------------------------------------------------------------
    # Property extraction methods
    
    def _extract_paragraph_properties(self, paragraph_el: ET.Element) -> Dict[str, object]:
        """Extract paragraph-level properties."""
        props = {}
        ppr = paragraph_el.find("w:pPr", Namespaces.WORD)
        if ppr is not None:
            props = self._serialize_properties_block(ppr)
        return props

    def _extract_run_properties(self, run_el: ET.Element) -> Dict[str, object]:
        """Extract run-level properties."""
        props = {}
        rpr = run_el.find("w:rPr", Namespaces.WORD)
        if rpr is not None:
            props = self._serialize_properties_block(rpr)
        return props

    def _extract_table_properties(self, table_el: ET.Element) -> Dict[str, object]:
        """Extract table-level properties."""
        props = {}
        tbl_pr = table_el.find("w:tblPr", Namespaces.WORD)
        if tbl_pr is not None:
            props = self._serialize_properties_block(tbl_pr)
        return props

    def _extract_row_properties(self, row_el: ET.Element) -> Dict[str, object]:
        """Extract table row properties."""
        props = {}
        tr_pr = row_el.find("w:trPr", Namespaces.WORD)
        if tr_pr is not None:
            props = self._serialize_properties_block(tr_pr)
        return props

    def _extract_cell_properties(self, cell_el: ET.Element) -> Dict[str, object]:
        """Extract table cell properties."""
        props = {}
        tc_pr = cell_el.find("w:tcPr", Namespaces.WORD)
        if tc_pr is not None:
            props = self._serialize_properties_block(tc_pr)
        return props

    def _extract_vertical_merge(self, cell_el: ET.Element) -> Optional[Dict[str, object]]:
        tc_pr = cell_el.find("w:tcPr", Namespaces.WORD)
        if tc_pr is None:
            return None

        vmerge_el = tc_pr.find("w:vMerge", Namespaces.WORD)
        if vmerge_el is None:
            return None

        merged = self._serialize_node(vmerge_el)
        return merged

    def _extract_numbering_info(self, paragraph_el: ET.Element) -> Optional[NumberingInfo]:
        """Extract numbering information from paragraph properties."""
        ppr = paragraph_el.find("w:pPr", Namespaces.WORD)
        if ppr is None:
            return None
        
        num_pr = ppr.find("w:numPr", Namespaces.WORD)
        if num_pr is None:
            return None
        
        num_id = self._get_int_attr(num_pr, "w:numId", "w:val")
        level = self._get_int_attr(num_pr, "w:ilvl", "w:val")
        
        if num_id is None or level is None:
            return None
        
        # Resolve numbering details
        instance = self._numbering.get_instance(num_id)
        if not instance:
            return NumberingInfo(num_id=num_id, level=level)
        
        abstract = self._numbering.get_abstract(instance.abstract_num_id)
        if not abstract or level not in abstract.levels:
            return NumberingInfo(num_id=num_id, level=level, abstract_num_id=instance.abstract_num_id)
        
        level_def = abstract.levels[level]
        
        # Check for overrides
        start_override = None
        if level in instance.overrides:
            start_override = instance.overrides[level].start_override
        
        return NumberingInfo(
            num_id=num_id,
            level=level,
            abstract_num_id=instance.abstract_num_id,
            start=start_override or level_def.start,
            format=level_def.num_format,
            level_text=level_def.level_text,
            alignment=level_def.alignment
        )

    def _get_table_style_id(self, table_el: ET.Element) -> Optional[str]:
        """Extract table style ID."""
        tbl_pr = table_el.find("w:tblPr", Namespaces.WORD)
        if tbl_pr is None:
            return None
        style_el = tbl_pr.find("w:tblStyle", Namespaces.WORD)
        if style_el is None:
            return None
        return self._get_attr(style_el, None, "w:val")

    def _serialize_properties_block(self, element: ET.Element) -> Dict[str, object]:
        """Serialize property block while preserving structure."""
        return {
            "tag": element.tag,
            "attributes": dict(element.attrib),
            "children": [self._serialize_node(child) for child in list(element)]
        }

    def _serialize_node(self, node: ET.Element) -> Dict[str, object]:
        """Serialize XML node to dictionary."""
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

    def _get_attr(self, element: Optional[ET.Element], child_name: Optional[str], attr_name: str) -> Optional[str]:
        """Get attribute value from element or its child."""
        if element is None:
            return None
        
        target = element.find(child_name, Namespaces.WORD) if child_name else element
        if target is None:
            return None
        
        # Handle both qualified and unqualified attribute names
        if ":" in attr_name:
            attr_key = self._qualify(attr_name)
        else:
            attr_key = attr_name
            
        return target.attrib.get(attr_key)

    def _get_int_attr(self, element: Optional[ET.Element], child_name: Optional[str], attr_name: str) -> Optional[int]:
        """Get integer attribute value."""
        value = self._get_attr(element, child_name, attr_name)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None
