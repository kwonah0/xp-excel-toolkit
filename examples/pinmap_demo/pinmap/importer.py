"""Pinmap-side import wrapper: a thin shim over excel_toolkit.import_xlsx."""

from __future__ import annotations

from pathlib import Path

from excel_toolkit import SheetConfig, import_xlsx

from pinmap.models import PIN_FIELD_MAP, PinEntry


SHEET_CONFIGS = {
    "Pinmap_*": SheetConfig(
        field_map=PIN_FIELD_MAP,
        domain_cls=PinEntry,
    ),
}


def import_pinmap(session, path: str | Path):
    """Import a pinmap xlsx into the given session."""
    return import_xlsx(session, path, sheet_configs=SHEET_CONFIGS)
