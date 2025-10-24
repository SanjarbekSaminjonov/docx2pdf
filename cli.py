from __future__ import annotations

import argparse
import logging
import webbrowser
from pathlib import Path
from typing import Sequence

from docx_renderer.utils.logger import get_logger

try:
    from .main import run_pipeline
except ImportError:  # pragma: no cover - allows running as script
    from main import run_pipeline  # type: ignore

LOGGER = get_logger(__name__)

_LOG_LEVELS = {"CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"}


def _configure_logging(level_name: str) -> None:
    """Update the root logger level based on user input."""
    level = getattr(logging, level_name.upper(), None)
    if not isinstance(level, int):
        raise ValueError(f"Unsupported log level: {level_name}")
    logging.getLogger().setLevel(level)


def _existing_docx(path_str: str) -> str:
    """Ensure the provided argument points to an existing DOCX file."""
    candidate = Path(path_str)
    if not candidate.exists():
        raise argparse.ArgumentTypeError(f"DOCX file not found: {path_str}")
    if candidate.suffix.lower() != ".docx":
        raise argparse.ArgumentTypeError("Input file must have a .docx extension")
    return path_str


def build_parser() -> argparse.ArgumentParser:
    """Construct the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="docx-render",
        description="Render DOCX files into flattened layout outputs",
    )
    parser.add_argument("docx_file", type=_existing_docx, help="Path to the input .docx file")
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        dest="output",
        help="Write artifacts under this directory (defaults to alongside the DOCX)",
    )
    parser.add_argument(
        "--html",
        dest="html",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Generate HTML output (enabled by default)",
    )
    parser.add_argument(
        "--pdf",
        dest="pdf",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Generate PDF output",
    )
    parser.add_argument(
        "--debug",
        dest="debug",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Dump intermediate debug artifacts",
    )
    parser.add_argument(
        "--open",
        dest="open_output",
        action="store_true",
        help="Open the generated HTML in the default browser",
    )
    parser.add_argument(
        "--log-level",
        choices=sorted(_LOG_LEVELS),
        default="INFO",
        help="Override the logger level (default: INFO)",
    )
    return parser


def _validate_outputs(args: argparse.Namespace) -> None:
    if not args.html and not args.pdf:
        raise argparse.ArgumentTypeError("At least one output format (--html/--pdf) must be enabled")


def _open_html(output_dir: Path) -> None:
    html_path = output_dir / "document.html"
    if html_path.exists():
        webbrowser.open_new_tab(html_path.as_uri())
    else:
        LOGGER.warning("No HTML artifact found at %s", html_path)


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
        _validate_outputs(args)
        _configure_logging(args.log_level)
        output_path = run_pipeline(
            args.docx_file,
            output_dir=args.output,
            html=args.html,
            pdf=args.pdf,
            dump_debug=args.debug,
        )
        if args.open_output and args.html:
            _open_html(output_path)
        return 0
    except argparse.ArgumentTypeError as exc:
        parser.error(str(exc))
    except FileNotFoundError as exc:
        LOGGER.error(str(exc))
        return 1
    except Exception:  # pragma: no cover - defensive logging
        LOGGER.exception("Unexpected error while rendering document")
        return 2