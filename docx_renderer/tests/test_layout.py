"""Tests covering basic layout calculations."""
import unittest

from docx_renderer.model.elements import DocumentTree, ParagraphElement, RunFragment
from docx_renderer.model.style_model import StylesCatalog
from docx_renderer.parser.layout_calculator import LayoutCalculator


class LayoutCalculatorTest(unittest.TestCase):
    """Validate cursor flow across sequential blocks."""

    def test_blocks_increment_vertical_cursor(self) -> None:
        paragraph = ParagraphElement(runs=[RunFragment(text="Hello")], style_id=None, properties={})
        tree = DocumentTree(blocks=[paragraph, paragraph])
        catalog = StylesCatalog({})
        layout = LayoutCalculator(catalog).calculate(tree)
        self.assertEqual(len(layout.boxes), 2)
        self.assertLess(layout.boxes[0].y, layout.boxes[1].y)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
