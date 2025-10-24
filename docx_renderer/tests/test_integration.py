"""
Integration tests for the complete DOCX parsing pipeline.

Tests the end-to-end parsing flow from DOCX package to final document model.
"""

import unittest
from unittest.mock import Mock, MagicMock
from xml.etree.ElementTree import Element, SubElement
import zipfile
import io

from main import build_document_model
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.model.document_model import DocumentModel


class IntegrationTest(unittest.TestCase):
    """Integration tests for complete DOCX processing pipeline."""
    
    def test_pipeline_with_minimal_document(self):
        """Test complete pipeline with minimal document structure."""
        # Create minimal DOCX structure in memory
        minimal_docx = self._create_minimal_docx_structure()
        
        # Mock the DocxPackage.load method to return our test structure
        original_load = DocxPackage.load
        DocxPackage.load = lambda path: minimal_docx
        
        try:
            # This would normally load from a real file path
            # but we're mocking the load method
            model = build_document_model("mock_path.docx")
            
            # Verify the model was created successfully
            self.assertIsInstance(model, DocumentModel)
            self.assertIsNotNone(model.styles)
            self.assertIsNotNone(model.layout)
            self.assertIsNotNone(model.numbering)
            self.assertIsNotNone(model.media)
            
            # Verify basic structure
            self.assertEqual(len(model.layout.boxes), 1)  # Should have one paragraph
            
        finally:
            # Restore original method
            DocxPackage.load = original_load
    
    def test_parser_components_integration(self):
        """Test that all parser components work together correctly."""
        from docx_renderer.parser.docx_loader import DocxPackage
        from docx_renderer.parser.rels_parser import Relationships  
        from docx_renderer.parser.styles_parser import StylesParser
        from docx_renderer.parser.numbering_parser import NumberingParser
        from docx_renderer.parser.document_parser import DocumentParser
        from docx_renderer.parser.media_extractor import extract_media_from_package
        
        # Create test package
        package = self._create_minimal_docx_structure()
        
        # Test each parser component
        relationships = Relationships.from_package(package.raw_parts)
        self.assertIsNotNone(relationships)
        
        media_catalog = extract_media_from_package(package, relationships)
        self.assertIsNotNone(media_catalog)
        
        styles = StylesParser(package.require_styles_xml(), package.get_numbering_xml()).parse()
        self.assertIsNotNone(styles)
        
        numbering = NumberingParser(package.get_numbering_xml()).parse()
        self.assertIsNotNone(numbering)
        
        document_tree = DocumentParser(package, styles, numbering).parse()
        self.assertIsNotNone(document_tree)
        
        # Verify integration
        self.assertEqual(len(document_tree.sections), 1)
        self.assertEqual(len(document_tree.sections[0].blocks), 1)
    
    def test_text_normalization_integration(self):
        """Test text normalization in the full pipeline."""
        from docx_renderer.utils.text_normalizer import normalize_docx_text
        
        # Test with various special characters
        test_text = "Hello\u00a0world\u2019s document\u2026"  # No soft hyphen to avoid removal
        normalized = normalize_docx_text(test_text)
        
        # Should normalize to clean text
        expected = "Hello world's document..."
        self.assertEqual(normalized, expected)
    
    def test_error_handling_in_pipeline(self):
        """Test error handling when components encounter issues."""
        # Test with malformed package
        malformed_package = DocxPackage(raw_parts={})
        
        # Should handle missing required parts gracefully
        with self.assertRaises((ValueError, KeyError)):
            malformed_package.require_document_xml()
        
        with self.assertRaises((ValueError, KeyError)):
            malformed_package.require_styles_xml()
    
    def _create_minimal_docx_structure(self) -> DocxPackage:
        """Create a minimal DOCX package structure for testing."""
        # Create minimal XML structures
        
        # Content Types
        content_types = Element("Types")
        content_types.set("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
        
        # Package relationships
        pkg_rels = Element("Relationships")
        pkg_rels.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        rel = SubElement(pkg_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        rel.set("Target", "word/document.xml")
        
        # Document
        document = Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document")
        body = SubElement(document, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
        para = SubElement(body, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
        run = SubElement(para, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        text = SubElement(run, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        text.text = "Hello World"
        
        # Styles
        styles = Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styles")
        default_style = SubElement(styles, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults")
        
        # Document relationships
        doc_rels = Element("Relationships")
        doc_rels.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        # Serialize to bytes
        import xml.etree.ElementTree as ET
        
        def serialize_element(elem):
            return ET.tostring(elem, encoding='utf-8')
        
        # Create the package structure
        parts = {
            "[Content_Types].xml": serialize_element(content_types),
            "_rels/.rels": serialize_element(pkg_rels),
            "word/document.xml": serialize_element(document),
            "word/styles.xml": serialize_element(styles),
            "word/_rels/document.xml.rels": serialize_element(doc_rels)
        }
        
        # Create and initialize package
        package = DocxPackage(raw_parts=parts)
        package._initialize_caches()
        
        return package


class PerformanceTest(unittest.TestCase):
    """Performance tests for the parsing pipeline."""
    
    def test_parser_performance_baseline(self):
        """Establish baseline performance for parser components."""
        import time
        
        # This is a placeholder for performance testing
        # In a real scenario, you'd test with larger documents
        start_time = time.time()
        
        # Simulate some parsing work
        from docx_renderer.utils.text_normalizer import TextNormalizer
        normalizer = TextNormalizer()
        
        # Normalize a bunch of text
        for _ in range(1000):
            normalizer.normalize_text("Test\u00a0text\u2019s\u00adcontent\u2026")
        
        elapsed = time.time() - start_time
        
        # Should complete reasonably quickly (less than 1 second for 1000 normalizations)
        self.assertLess(elapsed, 1.0)


if __name__ == '__main__':
    unittest.main()