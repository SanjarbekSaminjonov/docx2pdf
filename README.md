# DOCX Renderer Skeleton

This repository provides a modular foundation for a future DOCX → flattened model → PDF/HTML renderer pipeline. The goal is to ingest WordprocessingML documents, resolve styles and layout, and emit renderer-neutral structures that can be consumed by multiple backends.

## Architecture Overview

- **parser/** – loads the OPC package, parses XML, extracts styles/media, and emits structured blocks.
- **model/** – shared data structures for styles, content elements, and layout boxes.
- **renderer/** – targets that transform the flattened layout into HTML or PDF.
- **utils/** – reusable helpers for XML parsing, units, logging, and debug dumps.
- **tests/** – smoke tests and focused unit coverage for styles and layout logic.

## Getting Started

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

Run tests:

```bash
python -m unittest
```

## Next Steps

- Expand XML parsing coverage for tables, lists, images, and advanced features.
- Implement accurate layout calculations respecting Word metrics and pagination.
- Integrate ReportLab (or alternative) for PDF output and enhance HTML styling.
- Build debug visualizations from the `debug/` artifacts.
