"""Public facade for pinmap.

Downstream applications import everything from here::

    from pinmap.api import init_db, PinEntry, import_pinmap, export_pinmap

The point is that downstream code never has to know that excel_toolkit
is the underlying infra. ``pinmap.api.Base`` is just an alias for
``excel_toolkit.Base`` — they share the same MetaData, so a single
``init_db()`` creates infra tables AND ``pin_entry`` together.

If pinmap ever moves off excel_toolkit (or wraps a different backend),
only this facade module and the importer/exporter modules need to
change; downstream code stays put.
"""

from __future__ import annotations

# Re-exports from excel_toolkit — callers see these as members of pinmap.api
# and don't have to import the underlying package by name. Importantly,
# pinmap.api.Base IS excel_toolkit.Base (same identity, same MetaData) — we
# don't subclass it because SQLAlchemy 2.0's DeclarativeBase doesn't allow
# an un-mapped intermediate subclass, and aliasing is the canonical way to
# rename a registry-owning Base across packages.
from excel_toolkit import (
    Base,
    ChangeLog,
    ExcelCell,
    ExcelMerge,
    ExcelSheet,
    ExcelWorkbook,
    MergeResolver,
    init_db,
)


# Domain + helpers — keep imports lazy-feeling but they're plain top-level
# so pyright/IDE jump-to-definition stays useful.
from pinmap.models import PIN_FIELD_MAP, PinEntry
from pinmap.importer import SHEET_CONFIGS, import_pinmap
from pinmap.exporter import EXPORT_HANDLERS, export_pinmap


__all__ = [
    # Setup
    "Base",
    "init_db",
    # Domain
    "PinEntry",
    "PIN_FIELD_MAP",
    # Workflow
    "import_pinmap",
    "export_pinmap",
    "SHEET_CONFIGS",
    "EXPORT_HANDLERS",
    # Cell-level escape hatches (audit log, raw cells, merge resolution)
    "ChangeLog",
    "ExcelWorkbook",
    "ExcelSheet",
    "ExcelCell",
    "ExcelMerge",
    "MergeResolver",
]
