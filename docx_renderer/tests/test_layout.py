"""Tests covering basic layout calculations."""
import unittest
from typing import Optional

from docx_renderer.model.elements import (
    DocumentSection,
    DocumentTree,
    ImageElement,
    ParagraphElement,
    RunFragment,
    SectionProperties,
    TableCell,
    TableElement,
    TableRow,
)
from docx_renderer.model.style_model import StyleDefinition, StylesCatalog
from docx_renderer.parser.layout_calculator import (
    DEFAULT_LINE_HEIGHT_PT,
    DEFAULT_PAGE_WIDTH_PT,
    DEFAULT_TABLE_BORDER_WIDTH_PT,
    DEFAULT_TABLE_CELL_PADDING_PT,
    LayoutCalculator,
)


class LayoutCalculatorTest(unittest.TestCase):
    """Validate cursor flow across sequential blocks."""

    def test_blocks_increment_vertical_cursor(self) -> None:
        paragraph = ParagraphElement(runs=[RunFragment(text="Hello")], style_id=None, properties={})
        section = DocumentSection(blocks=[paragraph, paragraph], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})
        layout = LayoutCalculator(catalog).calculate(tree)

        self.assertEqual(len(layout.boxes), 2)
        self.assertLess(layout.boxes[0].y, layout.boxes[1].y)
        self.assertEqual(len(layout.pages), 1)
        self.assertEqual(len(layout.pages[0]), 2)
        self.assertGreater(layout.boxes[0].height, 0)

    def test_section_margins_reflected_in_layout_coordinates(self) -> None:
        props = SectionProperties(margin_left=720, margin_top=1440)
        paragraph = ParagraphElement(runs=[RunFragment(text="Hello")], style_id=None, properties={})
        section = DocumentSection(blocks=[paragraph], properties=props)
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        box = layout.boxes[0]

        self.assertAlmostEqual(box.x, 36.0, places=2)
        self.assertAlmostEqual(box.y, 72.0, places=2)
        self.assertIn("lines", box.content)
        self.assertEqual(box.content["lines"], ["Hello"])

    def test_page_size_and_orientation_applied(self) -> None:
        props = SectionProperties(
            page_width=12240,
            page_height=15840,
            orientation="landscape",
        )
        paragraph = ParagraphElement(runs=[RunFragment(text="Hello")], style_id=None, properties={})
        section = DocumentSection(blocks=[paragraph], properties=props)
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        box = layout.boxes[0]

        # Landscape orientation swaps page dimensions; ensure available width reflects swap
        expected_width = (15840 / 20.0) - (72.0 * 2)
        self.assertAlmostEqual(box.width, expected_width, places=2)

    def test_paragraph_spacing_before_and_after(self) -> None:
        paragraph1 = ParagraphElement(
            runs=[RunFragment(text="Intro")],
            style_id=None,
            properties={"spacing_after": 240},  # 12pt
        )
        paragraph2 = ParagraphElement(
            runs=[RunFragment(text="Body")],
            style_id=None,
            properties={"spacing_before": 480},  # 24pt
        )
        section = DocumentSection(blocks=[paragraph1, paragraph2], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        first_box, second_box = layout.boxes

        self.assertGreater(second_box.y - (first_box.y + first_box.height), 20.0)

    def test_text_wrapping_creates_multiple_lines(self) -> None:
        narrow_props = SectionProperties(
            page_width=4000,
            margin_left=720,
            margin_right=720,
        )
        long_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit"
        paragraph = ParagraphElement(runs=[RunFragment(text=long_text)], style_id=None, properties={})
        section = DocumentSection(blocks=[paragraph], properties=narrow_props)
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        box = layout.boxes[0]

        self.assertGreater(len(box.content["lines"]), 1)
        per_line_height = box.height / len(box.content["lines"])
        self.assertGreater(per_line_height, DEFAULT_LINE_HEIGHT_PT - 1)

    def test_paragraph_indent_applied_from_style(self) -> None:
        paragraph = ParagraphElement(runs=[RunFragment(text="Indented")], style_id="Body", properties={})
        style = StyleDefinition(
            style_id="Body",
            style_type="paragraph",
            name="Body",
            properties={
                "pPr": [
                    {
                        "tag": "w:ind",
                        "attributes": {"w:left": "720", "w:firstLine": "360"},
                    }
                ]
            },
        )
        section = DocumentSection(blocks=[paragraph], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({"Body": style})

        layout = LayoutCalculator(catalog).calculate(tree)
        box = layout.boxes[0]

        self.assertAlmostEqual(box.x, 72.0 + 36.0, places=2)
        self.assertAlmostEqual(box.content["firstLineIndent"], 18.0, places=2)
        self.assertAlmostEqual(box.width, (DEFAULT_PAGE_WIDTH_PT - 144.0) - 36.0, places=2)

    def test_spacing_resolves_from_style_when_missing_on_paragraph(self) -> None:
        paragraph = ParagraphElement(runs=[RunFragment(text="Styled spacing")], style_id="Heading", properties={})
        style = StyleDefinition(
            style_id="Heading",
            style_type="paragraph",
            name="Heading",
            properties={
                "pPr": [
                    {
                        "tag": "w:spacing",
                        "attributes": {
                            "w:before": "480",
                            "w:after": "240",
                            "w:line": "360",
                            "w:lineRule": "exact",
                        },
                    }
                ]
            },
        )

        section = DocumentSection(blocks=[paragraph, paragraph], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({"Heading": style})

        layout = LayoutCalculator(catalog).calculate(tree)
        first_box, second_box = layout.boxes

        self.assertGreater(first_box.y - 72.0, 20.0)
        self.assertAlmostEqual(first_box.style["spacing"]["before"], 24.0, places=2)
        self.assertAlmostEqual(first_box.style["spacing"]["after"], 12.0, places=2)
        self.assertAlmostEqual(first_box.height, 18.0, places=1)
        self.assertAlmostEqual(second_box.y - (first_box.y + first_box.height), 36.0, places=1)

    def test_basic_table_layout_allocates_rows_and_columns(self) -> None:
        def make_paragraph(text: str) -> ParagraphElement:
            return ParagraphElement(runs=[RunFragment(text=text)], style_id=None, properties={})

        head = make_paragraph("Above table")
        row1_col1 = TableCell(content=[make_paragraph("R1C1")], properties={})
        row1_col2 = TableCell(content=[make_paragraph("R1C2 text")], properties={})
        row2_col1 = TableCell(content=[make_paragraph("R2C1 line"), make_paragraph("more")], properties={})
        row2_col2 = TableCell(content=[make_paragraph("R2C2")], properties={})

        table = TableElement(
            rows=[
                TableRow(cells=[row1_col1, row1_col2]),
                TableRow(cells=[row2_col1, row2_col2]),
            ],
            style_id="TableGrid",
            properties={"borders": "single"},
        )
        tail = make_paragraph("After table")

        section = DocumentSection(blocks=[head, table, tail], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        para_box, table_box, tail_box = layout.boxes

        self.assertEqual(table_box.element_type, "table")
        self.assertEqual(table_box.content["columns"], 2)
        self.assertEqual(len(table_box.content["rowHeights"]), 2)
        base_margin = 2 * (DEFAULT_TABLE_CELL_PADDING_PT + DEFAULT_TABLE_BORDER_WIDTH_PT)
        self.assertAlmostEqual(
            table_box.content["rowHeights"][0],
            DEFAULT_LINE_HEIGHT_PT + base_margin,
            places=1,
        )
        self.assertAlmostEqual(
            table_box.content["rowHeights"][1],
            DEFAULT_LINE_HEIGHT_PT * 2 + base_margin,
            places=1,
        )
        self.assertGreater(tail_box.y, table_box.y + table_box.height)
        self.assertAlmostEqual(table_box.width, DEFAULT_PAGE_WIDTH_PT - 144.0, places=1)

        column_widths = table_box.content["columnWidths"]
        self.assertEqual(len(column_widths), 2)
        self.assertAlmostEqual(sum(column_widths), table_box.width, places=2)

        first_row_cells = table_box.content["cells"][0]
        self.assertEqual(first_row_cells[0]["columnIndex"], 0)
        self.assertEqual(first_row_cells[0]["colSpan"], 1)
        self.assertGreaterEqual(first_row_cells[0]["contentHeight"], 0.0)
        self.assertEqual(len(first_row_cells[0]["boxes"]), 1)
        self.assertEqual(first_row_cells[0]["boxes"][0].element_type, "paragraph")

        second_row_cells = table_box.content["cells"][1]
        self.assertEqual(second_row_cells[0]["colSpan"], 1)
        self.assertEqual(second_row_cells[1]["columnIndex"], 1)

    def test_image_block_layout_respects_dimensions(self) -> None:
        image = ImageElement(
            r_id="rId1",
            media_path="word/media/image1.png",
            width_emu=914400,
            height_emu=457200,
            properties={"wrapStyle": "square"},
        )
        section = DocumentSection(blocks=[image], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        box = layout.boxes[0]

        self.assertEqual(box.element_type, "image")
        self.assertAlmostEqual(box.width, 72.0, places=1)
        self.assertAlmostEqual(box.height, 36.0, places=1)
        self.assertEqual(box.style["wrapStyle"], "square")

    def test_table_auto_fit_adjusts_column_widths(self) -> None:
        def make_paragraph(text: str) -> ParagraphElement:
            return ParagraphElement(runs=[RunFragment(text=text)], style_id=None, properties={})

        long_word = "Supercalifragilisticexpialidocious"
        table = TableElement(
            rows=[
                TableRow(
                    cells=[
                        TableCell(content=[make_paragraph("Short")], properties={}),
                        TableCell(content=[make_paragraph(long_word)], properties={}),
                    ]
                )
            ],
            style_id=None,
            properties={},
        )
        section = DocumentSection(blocks=[table], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        table_box = layout.boxes[0]

        widths = table_box.content["columnWidths"]
        self.assertEqual(len(widths), 2)
        self.assertGreater(widths[1], widths[0])
        self.assertAlmostEqual(table_box.width, sum(widths), places=5)

    def test_table_respects_grid_span_alignment(self) -> None:
        WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        def make_props(span: int) -> dict:
            return {
                "tag": f"{{{WORD_NS}}}tcPr",
                "attributes": {},
                "children": [
                    {
                        "tag": f"{{{WORD_NS}}}gridSpan",
                        "attributes": {f"{{{WORD_NS}}}val": str(span)},
                    }
                ],
            }

        def make_paragraph(text: str) -> ParagraphElement:
            return ParagraphElement(runs=[RunFragment(text=text)], style_id=None, properties={})

        row1 = TableRow(cells=[TableCell(content=[make_paragraph("Span")], properties=make_props(2))])
        row2 = TableRow(
            cells=[
                TableCell(content=[make_paragraph("A")], properties={}),
                TableCell(content=[make_paragraph("B")], properties={}),
            ]
        )
        table = TableElement(rows=[row1, row2], style_id=None, properties={})
        section = DocumentSection(blocks=[table], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        table_box = layout.boxes[0]

        self.assertEqual(table_box.content["columns"], 2)
        widths = table_box.content["columnWidths"]
        self.assertAlmostEqual(table_box.content["cells"][0][0]["width"], sum(widths), places=4)
        self.assertEqual(table_box.content["cells"][0][0]["colSpan"], 2)
        self.assertEqual(table_box.content["cells"][1][0]["columnIndex"], 0)
        self.assertEqual(table_box.content["cells"][1][1]["columnIndex"], 1)

    def test_vertical_merge_combines_rows(self) -> None:
        WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        def make_cell(text: str, merge: Optional[str]) -> TableCell:
            paragraph = ParagraphElement(runs=[RunFragment(text=text)], style_id=None, properties={})
            props = {
                "tag": f"{{{WORD_NS}}}tcPr",
                "attributes": {},
                "children": [],
            }
            if merge is not None:
                vmerge_node = {
                    "tag": f"{{{WORD_NS}}}vMerge",
                    "attributes": {},
                }
                if merge:
                    vmerge_node["attributes"][f"{{{WORD_NS}}}val"] = merge
                props["children"].append(vmerge_node)
                props["vMerge"] = vmerge_node
            return TableCell(content=[paragraph], properties=props)

        top_cell = make_cell("Top", "restart")
        bottom_cell = make_cell("Bottom", "continue")
        right_top = make_cell("Right", None)
        right_bottom = make_cell("Right2", None)

        table = TableElement(
            rows=[
                TableRow(cells=[top_cell, right_top]),
                TableRow(cells=[bottom_cell, right_bottom]),
            ],
            style_id=None,
            properties={},
        )
        section = DocumentSection(blocks=[table], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        table_box = layout.boxes[0]

        self.assertEqual(table_box.element_type, "table")
        first_row = table_box.content["cells"][0]
        second_row = table_box.content["cells"][1]

        merged = first_row[0]
        follower = second_row[0]
        row_heights = table_box.content["rowHeights"]

        self.assertEqual(merged["vMerge"], "restart")
        self.assertEqual(merged["rowSpan"], 2)
        self.assertAlmostEqual(merged["height"], sum(row_heights[:2]), places=4)
        self.assertEqual(follower["vMerge"], "continue")
        self.assertEqual(follower["rowSpan"], 0)
        self.assertEqual(follower["height"], 0.0)
        self.assertEqual(follower["boxes"], [])
        self.assertEqual(follower["baseRow"], merged["rowIndex"])

    def test_anchored_image_position_and_wrap(self) -> None:
        image = ImageElement(
            r_id="rId10",
            media_path="word/media/image10.png",
            width_emu=914400,
            height_emu=914400,
            properties={
                "wrapStyle": "behind",
                "anchor": {"offset_x": 144.0, "offset_y": 216.0},
                "inline": False,
            },
        )
        paragraph = ParagraphElement(runs=[RunFragment(text="After")], style_id=None, properties={})
        section = DocumentSection(blocks=[image, paragraph], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        image_box, paragraph_box = layout.boxes

        self.assertEqual(image_box.style["wrapStyle"], "behind-text")
        self.assertFalse(image_box.style["inline"])
        self.assertAlmostEqual(image_box.x, 72.0 + 144.0, places=2)
        self.assertAlmostEqual(image_box.y, 72.0 + 216.0, places=2)
        self.assertAlmostEqual(paragraph_box.y, 72.0, places=2)

    def test_square_wrapped_image_advances_flow(self) -> None:
        image = ImageElement(
            r_id="rId11",
            media_path="word/media/image11.png",
            width_emu=914400,
            height_emu=457200,
            properties={"wrapStyle": "square", "anchor": {"offset_x": 0.0, "offset_y": 0.0}, "inline": False},
        )
        paragraph = ParagraphElement(runs=[RunFragment(text="Below")], style_id=None, properties={})
        section = DocumentSection(blocks=[image, paragraph], properties=SectionProperties())
        tree = DocumentTree(sections=[section])
        catalog = StylesCatalog({})

        layout = LayoutCalculator(catalog).calculate(tree)
        image_box, paragraph_box = layout.boxes

        self.assertEqual(image_box.style["wrapStyle"], "square")
        self.assertGreaterEqual(paragraph_box.y, image_box.y + image_box.height)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
