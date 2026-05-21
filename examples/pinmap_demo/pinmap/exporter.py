"""Pinmap-side export wrapper: a thin shim over excel_toolkit.export_domain_xlsx."""

from __future__ import annotations

from pathlib import Path

from excel_toolkit import ExportHandler, export_domain_xlsx

from pinmap.models import PIN_FIELD_MAP, PinEntry


EXPORT_HANDLERS = [
    ExportHandler(
        pattern="Pinmap_*",
        field_map=PIN_FIELD_MAP,
        domain_cls=PinEntry,
    ),
]


def export_pinmap(session, output_path: str | Path) -> Path:
    """Round-trip the in-DB pinmap to xlsx, preserving original formatting."""
    return export_domain_xlsx(session, output_path, EXPORT_HANDLERS)
