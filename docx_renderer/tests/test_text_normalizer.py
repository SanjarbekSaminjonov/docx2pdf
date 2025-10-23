"""Test cases for text normalization functionality."""

import unittest
from xml.etree.ElementTree import Element, SubElement

from docx_renderer.utils.text_normalizer import (
    TextNormalizer, NamespaceStripper, normalize_docx_text
)


class TextNormalizerTest(unittest.TestCase):
    """Test text normalization utilities."""

    def setUp(self):
        """Set up test fixtures."""
        self.normalizer = TextNormalizer(preserve_whitespace=False)
        self.preserve_normalizer = TextNormalizer(preserve_whitespace=True)

    def test_special_character_replacement(self):
        """Test replacement of special Unicode characters."""
        test_cases = [
            ('\u00a0text', 'text'),   # Non-breaking space with text (leading space stripped)
            ('text\u2009word', 'text word'),  # Thin space between words
            ('\u200b', ''),       # Zero-width space
            ('\u00ad', ''),       # Soft hyphen
            ('\u2013', '–'),      # En dash
            ('\u2018\u2019', "''"),  # Smart quotes
            ('\u201c\u201d', '""'),  # Smart double quotes
            ('\u2026', '...'),    # Ellipsis
        ]
        
        for input_text, expected in test_cases:
            result = self.normalizer.normalize_text(input_text)
            self.assertEqual(result, expected, f"Failed for input: {repr(input_text)}")

    def test_whitespace_normalization(self):
        """Test whitespace collapsing and trimming."""
        test_cases = [
            ('  multiple   spaces  ', 'multiple spaces'),
            ('\t\ttabs\t\t', 'tabs'),
            ('\n\nnewlines\n\n', 'newlines'),
            ('   mixed \t\n\r  whitespace   ', 'mixed whitespace'),
            ('', ''),
            ('   ', ''),
            ('no extra spaces', 'no extra spaces'),
        ]
        
        for input_text, expected in test_cases:
            result = self.normalizer.normalize_text(input_text)
            self.assertEqual(result, expected, f"Failed for input: {repr(input_text)}")

    def test_preserve_whitespace_mode(self):
        """Test that whitespace preservation works correctly."""
        input_text = '  multiple   spaces  '
        
        # Normal mode collapses whitespace
        normal_result = self.normalizer.normalize_text(input_text)
        self.assertEqual(normal_result, 'multiple spaces')
        
        # Preserve mode keeps whitespace as-is (except control chars)
        preserve_result = self.preserve_normalizer.normalize_text(input_text)
        self.assertEqual(preserve_result, input_text)

    def test_control_character_removal(self):
        """Test removal of control characters."""
        # Include some control characters that should be removed
        input_text = 'text\x00with\x08control\x1fchars'
        expected = 'textwithcontrolchars'
        
        result = self.normalizer.normalize_text(input_text)
        self.assertEqual(result, expected)

    def test_element_text_extraction(self):
        """Test text extraction from XML elements."""
        # Create test element with mixed text content
        root = Element('root')
        root.text = 'Start text '
        
        child1 = SubElement(root, 'child1')
        child1.text = 'child content'
        child1.tail = ' after child1 '
        
        child2 = SubElement(root, 'child2')
        child2.text = 'more content'
        child2.tail = ' end text'
        
        result = self.normalizer.normalize_element_text(root)
        expected = 'Start text child content after child1 more content end text'
        self.assertEqual(result, expected)

    def test_plain_text_extraction(self):
        """Test plain text extraction using itertext."""
        root = Element('paragraph')
        root.text = 'This is '
        
        bold = SubElement(root, 'bold')
        bold.text = 'bold text'
        bold.tail = ' and '
        
        italic = SubElement(root, 'italic')
        italic.text = 'italic'
        italic.tail = ' text.'
        
        result = self.normalizer.extract_plain_text(root)
        expected = 'This is bold text and italic text.'
        self.assertEqual(result, expected)

    def test_empty_and_none_handling(self):
        """Test handling of empty strings and None values."""
        self.assertEqual(self.normalizer.normalize_text(''), '')
        self.assertEqual(self.normalizer.normalize_text(None), None)
        self.assertEqual(self.normalizer.normalize_element_text(None), '')
        self.assertEqual(self.normalizer.extract_plain_text(None), '')


class NamespaceStripperTest(unittest.TestCase):
    """Test namespace stripping functionality."""

    def setUp(self):
        """Set up test fixtures."""
        self.stripper = NamespaceStripper()

    def test_namespace_stripping_from_text(self):
        """Test stripping namespace prefixes from text."""
        test_cases = [
            ('w:document', 'document'),
            ('wp:docPr w:name', 'docPr name'),
            ('a:graphic pic:pic', 'graphic pic'),
            ('w14:textFill', 'textFill'),
            ('no:namespaces:here', 'no:namespaces:here'),  # Unknown prefixes not stripped
        ]
        
        for input_text, expected in test_cases:
            result = self.stripper.strip_namespaces(input_text)
            self.assertEqual(result, expected, f"Failed for input: {input_text}")

    def test_element_namespace_stripping(self):
        """Test in-place namespace stripping from XML elements."""
        # Create element with namespaced tag and attributes
        element = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}paragraph')
        element.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        element.set('normalAttr', 'value')
        
        self.stripper.strip_element_namespaces(element)
        
        # Check that namespace was stripped from tag
        self.assertEqual(element.tag, 'paragraph')
        
        # Check that namespace was stripped from attributes
        expected_attrib = {'space': 'preserve', 'normalAttr': 'value'}
        self.assertEqual(element.attrib, expected_attrib)

    def test_nested_element_namespace_stripping(self):
        """Test namespace stripping on nested elements."""
        root = Element('{http://example.com/ns}root')
        child = SubElement(root, '{http://example.com/ns}child')
        grandchild = SubElement(child, '{http://example.com/ns}grandchild')
        
        self.stripper.strip_element_namespaces(root)
        
        self.assertEqual(root.tag, 'root')
        self.assertEqual(child.tag, 'child')
        self.assertEqual(grandchild.tag, 'grandchild')


class NormalizeDocxTextTest(unittest.TestCase):
    """Test the convenience function for DOCX text normalization."""

    def test_string_normalization(self):
        """Test normalizing string input."""
        input_text = '  w:text\u00a0with\u2009special  '
        result = normalize_docx_text(input_text)
        expected = 'text with special'
        self.assertEqual(result, expected)

    def test_element_normalization(self):
        """Test normalizing XML element input."""
        element = Element('w:p')
        element.text = '  Text\u00a0content  '
        
        result = normalize_docx_text(element)
        expected = 'Text content'
        self.assertEqual(result, expected)

    def test_preserve_whitespace_option(self):
        """Test whitespace preservation option."""
        input_text = '  multiple   spaces  '
        
        # Default behavior normalizes whitespace
        normal = normalize_docx_text(input_text)
        self.assertEqual(normal, 'multiple spaces')
        
        # With preserve_whitespace=True
        preserved = normalize_docx_text(input_text, preserve_whitespace=True)
        self.assertEqual(preserved, '  multiple   spaces  ')

    def test_namespace_stripping_option(self):
        """Test namespace stripping option."""
        input_text = 'w:paragraph wp:content'
        
        # Default behavior strips namespaces
        stripped = normalize_docx_text(input_text)
        self.assertEqual(stripped, 'paragraph content')
        
        # With strip_namespaces=False
        not_stripped = normalize_docx_text(input_text, strip_namespaces=False)
        self.assertEqual(not_stripped, 'w:paragraph wp:content')

    def test_none_input(self):
        """Test handling of None input."""
        result = normalize_docx_text(None)
        self.assertEqual(result, '')


if __name__ == '__main__':
    unittest.main()