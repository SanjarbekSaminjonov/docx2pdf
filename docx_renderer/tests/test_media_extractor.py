"""Test cases for media extraction functionality."""

import unittest
from unittest.mock import Mock

from docx_renderer.parser.media_extractor import MediaExtractor, MediaCatalog
from docx_renderer.parser.docx_loader import DocxPackage
from docx_renderer.parser.rels_parser import Relationships


class MediaExtractorTest(unittest.TestCase):
    """Test media extraction from DOCX packages."""

    def setUp(self):
        """Set up test fixtures."""
        self.package = Mock(spec=DocxPackage)
        self.relationships = Mock(spec=Relationships)
        self.extractor = MediaExtractor(self.package, self.relationships)

    def test_extract_image_assets(self):
        """Test extraction of image assets."""
        # Mock relationships to return image relationship
        self.relationships.get_targets_by_type.return_value = {
            'rId1': 'media/image1.png'
        }
        
        # Mock package to return image data
        image_data = b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x01\x00\x00\x00\x01\x00'
        self.package.get_part_data.return_value = image_data
        
        catalog = self.extractor.extract_media_catalog()
        
        self.assertEqual(len(catalog.assets), 1)
        self.assertEqual(len(catalog.images), 1)
        
        asset = catalog.get_by_id('rId1')
        self.assertIsNotNone(asset)
        self.assertEqual(asset.relationship_id, 'rId1')
        self.assertEqual(asset.target_path, 'media/image1.png')
        self.assertEqual(asset.media_type, 'image/png')
        self.assertEqual(asset.binary_data, image_data)

    def test_extract_font_assets(self):
        """Test extraction of font assets."""
        # Mock relationships for fonts
        self.relationships.get_targets_by_type.side_effect = [
            {},  # No media assets
            {'rId2': 'fonts/arial.ttf'}  # Font assets
        ]
        
        # Mock package to return font data
        font_data = b'TTF_FONT_DATA'
        self.package.get_part_data.return_value = font_data
        
        catalog = self.extractor.extract_media_catalog()
        
        self.assertEqual(len(catalog.fonts), 1)
        self.assertIn('arial', catalog.fonts)
        
        font_asset = catalog.fonts['arial']
        self.assertEqual(font_asset.target_path, 'fonts/arial.ttf')
        self.assertEqual(font_asset.media_type, 'application/font-woff')

    def test_media_type_detection(self):
        """Test MIME type detection from file extensions."""
        test_cases = [
            ('image.png', 'image/png'),
            ('photo.jpg', 'image/jpeg'),
            ('document.pdf', 'application/pdf'),
            ('unknown.unknownext', 'application/octet-stream')  # Unknown extension falls back
        ]
        
        for filename, expected_type in test_cases:
            result = self.extractor._get_media_type(filename)
            self.assertEqual(result, expected_type)

    def test_png_dimensions_extraction(self):
        """Test PNG image dimensions extraction."""
        # Valid PNG header with 256x256 dimensions
        png_data = (
            b'\x89PNG\r\n\x1a\n'  # PNG signature
            b'\x00\x00\x00\rIHDR'  # IHDR chunk
            b'\x00\x00\x01\x00'  # Width: 256
            b'\x00\x00\x01\x00'  # Height: 256
            b'\x08\x02\x00\x00\x00'  # Rest of IHDR
        )
        
        dimensions = self.extractor._get_image_dimensions(png_data)
        self.assertEqual(dimensions['width'], 256)
        self.assertEqual(dimensions['height'], 256)

    def test_empty_catalog_creation(self):
        """Test creating empty media catalog."""
        self.relationships.get_targets_by_type.return_value = {}
        
        catalog = self.extractor.extract_media_catalog()
        
        self.assertEqual(len(catalog.assets), 0)
        self.assertEqual(len(catalog.fonts), 0)
        self.assertEqual(len(catalog.images), 0)
        self.assertEqual(len(catalog.audio), 0)
        self.assertEqual(len(catalog.video), 0)
        self.assertEqual(len(catalog.documents), 0)


if __name__ == '__main__':
    unittest.main()