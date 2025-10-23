"""Tests for numbering parser behavior."""
import unittest
from xml.etree import ElementTree as ET

from docx_renderer.parser.numbering_parser import NumberingParser


NUMBERING_XML = """
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="multilevel"/>
    <w:name w:val="List Bullet"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%2."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="5">
    <w:abstractNumId w:val="1"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="3"/>
    </w:lvlOverride>
  </w:num>
</w:numbering>
"""


class NumberingParserTest(unittest.TestCase):
    """Ensure numbering parser captures definitions correctly."""

    def setUp(self) -> None:
        self.tree = ET.ElementTree(ET.fromstring(NUMBERING_XML))
        self.catalog = NumberingParser(self.tree).parse()

    def test_abstracts_parsed(self) -> None:
        abstract = self.catalog.get_abstract(1)
        self.assertIsNotNone(abstract)
        assert abstract
        self.assertEqual(abstract.multi_level_type, "multilevel")
        self.assertEqual(abstract.name, "List Bullet")
        self.assertIn(0, abstract.levels)
        level0 = abstract.levels[0]
        self.assertEqual(level0.num_format, "bullet")
        self.assertEqual(level0.level_text, "\u2022")
        self.assertEqual(level0.alignment, "left")
        self.assertIsNotNone(level0.raw_properties.get("tag"))
        self.assertTrue(level0.paragraph_properties)

    def test_instances_and_overrides(self) -> None:
        instance = self.catalog.get_instance(5)
        self.assertIsNotNone(instance)
        assert instance
        self.assertEqual(instance.abstract_num_id, 1)
        self.assertIn(0, instance.overrides)
        override = instance.overrides[0]
        self.assertEqual(override.start_override, 3)
        self.assertEqual(override.raw_properties["tag"].split("}")[-1], "lvlOverride")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
