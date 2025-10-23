"""Tests for document parser functionality."""
import unittest
from xml.etree import ElementTree as ET

from docx_renderer.model.numbering_model import NumberingCatalog
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.document_parser import DocumentParser
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.rels_parser import Relationships


class MockDocxPackage:
    """Mock DOCX package for testing."""

    def __init__(self, document_xml: str):
        self.document_xml = ET.ElementTree(ET.fromstring(document_xml))
        self.relationships = Relationships({})
        self.media = {}

    def require_document_xml(self):
        return self.document_xml


class DocumentParserTest(unittest.TestCase):
    """Test document parsing functionality."""

    def setUp(self) -> None:
        self.styles = StylesCatalog({})
        self.numbering = NumberingCatalog(abstracts={}, instances={})

    def test_parse_basic_paragraph(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>Hello World</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        package = MockDocxPackage(xml)
        parser = DocumentParser(package, self.styles, self.numbering)
        doc_tree = parser.parse()
        
        self.assertEqual(len(doc_tree.blocks), 1)
        paragraph = doc_tree.blocks[0]
        self.assertEqual(len(paragraph.runs), 1)
        self.assertEqual(paragraph.runs[0].text, "Hello World")

    def test_parse_paragraph_with_formatting(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Heading1"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:b/>
                  <w:i/>
                </w:rPr>
                <w:t>Formatted Text</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        package = MockDocxPackage(xml)
        parser = DocumentParser(package, self.styles, self.numbering)
        doc_tree = parser.parse()
        
        paragraph = doc_tree.blocks[0]
        self.assertEqual(paragraph.style_id, "Heading1")
        self.assertEqual(len(paragraph.runs), 1)
        self.assertEqual(paragraph.runs[0].text, "Formatted Text")
        self.assertIn("children", paragraph.runs[0].properties)

    def test_parse_table(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:tbl>
              <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
              </w:tblPr>
              <w:tr>
                <w:tc>
                  <w:p>
                    <w:r>
                      <w:t>Cell 1</w:t>
                    </w:r>
                  </w:p>
                </w:tc>
                <w:tc>
                  <w:p>
                    <w:r>
                      <w:t>Cell 2</w:t>
                    </w:r>
                  </w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
          </w:body>
        </w:document>
        """
        package = MockDocxPackage(xml)
        parser = DocumentParser(package, self.styles, self.numbering)
        doc_tree = parser.parse()
        
        self.assertEqual(len(doc_tree.blocks), 1)
        table = doc_tree.blocks[0]
        self.assertEqual(table.style_id, "TableGrid")
        self.assertEqual(len(table.rows), 1)
        self.assertEqual(len(table.rows[0].cells), 2)
        
        cell1 = table.rows[0].cells[0]
        self.assertEqual(len(cell1.content), 1)
        self.assertEqual(cell1.content[0].runs[0].text, "Cell 1")

    def test_parse_mixed_content(self) -> None:
        xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>First paragraph</w:t>
              </w:r>
            </w:p>
            <w:tbl>
              <w:tr>
                <w:tc>
                  <w:p>
                    <w:r>
                      <w:t>Table content</w:t>
                    </w:r>
                  </w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            <w:p>
              <w:r>
                <w:t>Second paragraph</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        package = MockDocxPackage(xml)
        parser = DocumentParser(package, self.styles, self.numbering)
        doc_tree = parser.parse()
        
        self.assertEqual(len(doc_tree.blocks), 3)
        self.assertEqual(doc_tree.blocks[0].runs[0].text, "First paragraph")
        self.assertEqual(len(doc_tree.blocks[1].rows), 1)
        self.assertEqual(doc_tree.blocks[2].runs[0].text, "Second paragraph")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()