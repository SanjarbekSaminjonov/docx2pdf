import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import mock

PROJECT_PARENT = Path(__file__).resolve().parents[3]
if str(PROJECT_PARENT) not in sys.path:
    sys.path.insert(0, str(PROJECT_PARENT))

from docx2pdf.cli import build_parser, main as cli_main


class CLITestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp_dir = TemporaryDirectory()
        self.addCleanup(self._temp_dir.cleanup)
        self.docx_path = Path(self._temp_dir.name) / "sample.docx"
        self.docx_path.write_bytes(b"PK\x03\x04")

    def test_parser_defaults(self) -> None:
        parser = build_parser()
        args = parser.parse_args([str(self.docx_path)])
        self.assertTrue(args.html)
        self.assertFalse(args.pdf)
        self.assertFalse(args.debug)

    def test_parser_boolean_overrides(self) -> None:
        parser = build_parser()
        args = parser.parse_args([str(self.docx_path), "--no-html", "--pdf", "--debug"])
        self.assertFalse(args.html)
        self.assertTrue(args.pdf)
        self.assertTrue(args.debug)

    @mock.patch("docx2pdf.cli.run_pipeline")
    def test_main_invokes_pipeline(self, mock_run_pipeline: mock.Mock) -> None:
        output_dir = Path(self._temp_dir.name) / "artifacts"
        output_dir.mkdir()
        mock_run_pipeline.return_value = output_dir

        exit_code = cli_main([str(self.docx_path), "--pdf", "--debug"])

        self.assertEqual(exit_code, 0)
        mock_run_pipeline.assert_called_once_with(
            str(self.docx_path),
            output_dir=None,
            html=True,
            pdf=True,
            dump_debug=True,
        )

    @mock.patch("docx2pdf.cli.run_pipeline")
    def test_main_requires_output_format(self, mock_run_pipeline: mock.Mock) -> None:
        mock_run_pipeline.return_value = Path(self._temp_dir.name)
        with self.assertRaises(SystemExit) as ctx:
            cli_main([str(self.docx_path), "--no-html", "--no-pdf"])
        self.assertEqual(ctx.exception.code, 2)
        mock_run_pipeline.assert_not_called()

    @mock.patch("docx2pdf.cli.webbrowser.open_new_tab")
    @mock.patch("docx2pdf.cli.run_pipeline")
    def test_main_optionally_opens_html(
        self,
        mock_run_pipeline: mock.Mock,
        mock_open_new_tab: mock.Mock,
    ) -> None:
        output_dir = Path(self._temp_dir.name) / "artifacts"
        output_dir.mkdir()
        html_path = output_dir / "document.html"
        html_path.write_text("<html></html>")
        mock_run_pipeline.return_value = output_dir

        exit_code = cli_main([str(self.docx_path), "--open"])

        self.assertEqual(exit_code, 0)
        mock_open_new_tab.assert_called_once_with(html_path.as_uri())


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
