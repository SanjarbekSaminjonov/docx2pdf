"""Tests for section and header/footer parsing."""
import unittest
from xml.etree import ElementTree as ET

from docx_renderer.model.numbering_model import NumberingCatalog
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.rels_parser import Relationships
from docx_renderer.parser.section_parser import SectionParser


class MockDocxPackageWithHeaders:
    """Mock DOCX package with header/footer support."""

    def __init__(self, document_xml: str, headers: dict = None, relationships_data: dict = None):
        self.document_xml = ET.ElementTree(ET.fromstring(document_xml))
        self.headers = headers or {}
        self.footers = {}
        
        if relationships_data:
            self.relationships = Relationships(relationships_data)
        else:
            self.relationships = Relationships({})
        
        self.media = {}

    def require_document_xml(self):
        return self.document_xml

    def get_xml_part(self, name: str):
        if name in self.headers:
            return ET.ElementTree(ET.fromstring(self.headers[name]))
        return None


class SectionParserTest(unittest.TestCase):
    """Test section parsing with headers and footers."""

    def setUp(self) -> None:
        self.styles = StylesCatalog({})
        self.numbering = NumberingCatalog(abstracts={}, instances={})

    def test_parse_simple_section(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>Content</w:t>
              </w:r>
            </w:p>
            <w:sectPr>
              <w:pgSz w:w="12240" w:h="15840"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """
        
        package = MockDocxPackageWithHeaders(xml)
        parser = SectionParser(package, self.styles, self.numbering)
        
        # Mock paragraph element
        from docx_renderer.model.elements import ParagraphElement, RunFragment
        blocks = [ParagraphElement(runs=[RunFragment(text="Content")], style_id=None)]
        
        doc_tree = parser.parse_sections(blocks)
        
        self.assertEqual(len(doc_tree.sections), 1)
        section = doc_tree.sections[0]
        self.assertEqual(len(section.blocks), 1)
        self.assertEqual(section.properties.page_width, 12240)
        self.assertEqual(section.properties.page_height, 15840)
        self.assertEqual(section.properties.margin_top, 1440)

    def test_parse_section_with_headers(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:r>
                <w:t>Content</w:t>
              </w:r>
            </w:p>
            <w:sectPr>
              <w:headerReference w:type="default" r:id="rId1"/>
              <w:footerReference w:type="default" r:id="rId2"/>
              <w:pgSz w:w="12240" w:h="15840"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """
        
        header_xml = """
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p>
            <w:r>
              <w:t>Header Content</w:t>
            </w:r>
          </w:p>
        </w:hdr>
        """
        
        relationships_data = {
            "word/document.xml": {
                "rId1": type("MockRel", (), {
                    "r_id": "rId1",
                    "resolved_target": "word/header1.xml",
                    "is_external": False
                })(),
                "rId2": type("MockRel", (), {
                    "r_id": "rId2", 
                    "resolved_target": "word/footer1.xml",
                    "is_external": False
                })()
            }
        }
        
        package = MockDocxPackageWithHeaders(
            xml, 
            headers={"word/header1.xml": header_xml},
            relationships_data=relationships_data
        )
        parser = SectionParser(package, self.styles, self.numbering)
        
        # Mock paragraph element
        from docx_renderer.model.elements import ParagraphElement, RunFragment
        blocks = [ParagraphElement(runs=[RunFragment(text="Content")], style_id=None)]
        
        doc_tree = parser.parse_sections(blocks)
        
        self.assertEqual(len(doc_tree.sections), 1)
        section = doc_tree.sections[0]
        
        # Check header content
        self.assertIsNotNone(section.properties.header_default)
        self.assertEqual(section.properties.header_default.r_id, "rId1")
        self.assertEqual(len(section.properties.header_default.blocks), 1)
        self.assertEqual(section.properties.header_default.blocks[0].runs[0].text, "Header Content")

    def test_multiple_sections(self) -> None:
        """Test document with multiple sections."""
        from docx_renderer.model.elements import ParagraphElement, RunFragment
        
        # Create blocks representing content from different sections
        blocks = [
            ParagraphElement(runs=[RunFragment(text="Section 1 content")], style_id=None),
            # Section break would be detected in actual parsing
            ParagraphElement(runs=[RunFragment(text="Section 2 content")], style_id=None),
        ]
        
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:sectPr>
              <w:pgSz w:w="12240" w:h="15840"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """
        
        package = MockDocxPackageWithHeaders(xml)
        parser = SectionParser(package, self.styles, self.numbering)
        
        doc_tree = parser.parse_sections(blocks)
        
        # Should have one section with all blocks
        self.assertEqual(len(doc_tree.sections), 1)
        self.assertEqual(len(doc_tree.sections[0].blocks), 2)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()