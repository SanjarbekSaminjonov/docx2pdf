"""Unit tests for style parsing edge cases."""
import unittest
from xml.etree import ElementTree as ET

from docx_renderer.parser.styles_parser import StylesParser


class StylesParserTest(unittest.TestCase):
    """Ensure style inheritance merges expected properties."""

    def test_style_inheritance_merges_properties(self) -> None:
        xml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="Base">
            <w:rPr><w:b/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Derived">
            <w:basedOn w:val="Base"/>
            <w:rPr><w:i/></w:rPr>
          </w:style>
        </w:styles>
        """
        tree = ET.ElementTree(ET.fromstring(xml))
        catalog = StylesParser(tree).parse()
        derived = catalog.get("Derived")
        self.assertIsNotNone(derived)
        assert derived  # for type checkers
        self.assertTrue(derived.properties.get("bold"))
        self.assertTrue(derived.properties.get("italic"))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
