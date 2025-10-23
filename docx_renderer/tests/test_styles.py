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
        rpr = derived.properties.get("rPr")
        self.assertIsInstance(rpr, list)
        assert isinstance(rpr, list)
        tags = [entry["tag"].split("}")[-1] for entry in rpr]
        self.assertIn("b", tags)
        self.assertIn("i", tags)

    def test_metadata_and_defaults_propagate(self) -> None:
        xml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="character" w:styleId="CharBase" w:default="1">
            <w:name w:val="Char Base"/>
            <w:alias w:val="AliasOne"/>
            <w:uiPriority w:val="1"/>
            <w:qFormat/>
          </w:style>
          <w:style w:type="character" w:styleId="CharDerived">
            <w:basedOn w:val="CharBase"/>
            <w:link w:val="LinkedStyle"/>
          </w:style>
        </w:styles>
        """
        tree = ET.ElementTree(ET.fromstring(xml))
        catalog = StylesParser(tree).parse()
        derived = catalog.get("CharDerived")
        self.assertIsNotNone(derived)
        assert derived
        self.assertTrue(derived.is_default)
        self.assertTrue(derived.is_primary)
        self.assertEqual(derived.ui_priority, 1)
        self.assertEqual(derived.linked_style, "LinkedStyle")
        self.assertEqual(derived.aliases, "AliasOne")

    def test_table_properties_preserved(self) -> None:
        xml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="table" w:styleId="TableBase">
            <w:tblPr>
              <w:tblBorders>
                <w:top w:val="single" w:sz="8"/>
              </w:tblBorders>
            </w:tblPr>
          </w:style>
          <w:style w:type="table" w:styleId="TableDerived">
            <w:basedOn w:val="TableBase"/>
            <w:tblPr>
              <w:tblCellMar>
                <w:top w:w="100" w:type="dxa"/>
              </w:tblCellMar>
            </w:tblPr>
          </w:style>
        </w:styles>
        """
        tree = ET.ElementTree(ET.fromstring(xml))
        catalog = StylesParser(tree).parse()
        derived = catalog.get("TableDerived")
        assert derived
        tbl_pr = derived.properties.get("tblPr")
        assert isinstance(tbl_pr, list)
        tags = [entry["tag"].split("}")[-1] for entry in tbl_pr]
        self.assertIn("tblBorders", tags)
        self.assertIn("tblCellMar", tags)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
