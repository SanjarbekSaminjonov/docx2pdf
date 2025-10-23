"""Helpers to persist intermediate representations for debugging."""
from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from docx_renderer.model.document_model import DocumentModel


class DebugDumper:
    """Writes intermediate artifacts onto disk for inspection."""

    def __init__(self, directory: Path) -> None:
        self.directory = directory

    def dump(self, model: DocumentModel) -> None:
        """Persist the document model as JSON for offline analysis."""
        self.directory.mkdir(parents=True, exist_ok=True)
        payload = self._serialize(model)
        (self.directory / "document_model.json").write_text(json.dumps(payload, indent=2))

    def _serialize(self, value: Any) -> Any:
        if is_dataclass(value):
            return {k: self._serialize(v) for k, v in asdict(value).items()}
        if isinstance(value, dict):
            return {k: self._serialize(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [self._serialize(v) for v in value]
        return value
