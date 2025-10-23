"""Tests for relationship parsing and indexing."""
import unittest

from docx_renderer.parser.rels_parser import RELTYPE_HEADER, RELTYPE_IMAGE, Relationships


doc_rels_xml = """
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>
"""

header_rels_xml = """
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>
"""


class RelationshipsTest(unittest.TestCase):
    """Validate relationship categorisation and target resolution."""

    def setUp(self) -> None:
        self.parts = {
            "word/_rels/document.xml.rels": doc_rels_xml.encode("utf-8"),
            "word/_rels/header1.xml.rels": header_rels_xml.encode("utf-8"),
        }

    def test_document_summary_groups_targets(self) -> None:
        relationships = Relationships.from_package(self.parts)
        summary = relationships.document_summary()

        self.assertIn("rId1", summary.headers)
        self.assertEqual(summary.headers["rId1"].resolved_target, "word/header1.xml")

        self.assertIn("rId3", summary.media)
        self.assertEqual(summary.media["rId3"].resolved_target, "word/media/image1.png")

        self.assertIn("rId4", summary.numbering)
        self.assertEqual(summary.numbering["rId4"].resolved_target, "word/numbering.xml")

        self.assertIn("rId5", summary.hyperlinks)
        self.assertTrue(summary.hyperlinks["rId5"].is_external)

    def test_part_lookup_normalizes_names(self) -> None:
        relationships = Relationships.from_package(self.parts)

        header_rels = relationships.for_source("word/header1.xml")
        self.assertIn("rId1", header_rels)
        self.assertEqual(header_rels["rId1"].resolved_target, "word/media/image2.png")

        rel = relationships.find("word/_rels/header1.xml.rels", "rId1")
        assert rel is not None
        self.assertEqual(rel.resolved_target, "word/media/image2.png")
        self.assertEqual(rel.rel_type, RELTYPE_IMAGE)

    def test_iter_all_includes_document_relationships(self) -> None:
        relationships = Relationships.from_package(self.parts)
        rel_types = {rel.rel_type for rel in relationships.iter_all()}
        self.assertIn(RELTYPE_HEADER, rel_types)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
