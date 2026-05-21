"""pinmap — toy host package consuming excel_toolkit.

This ``__init__`` intentionally stays empty. The package's public surface
lives in :mod:`pinmap.api` (the *facade* module) — downstream applications
should ``from pinmap.api import ...`` and never reach into
:mod:`pinmap.models`, :mod:`pinmap.importer`, or :mod:`pinmap.exporter`
directly.

Keeping ``__init__.py`` empty also avoids running ``register_audit_target``
side effects (and the resulting SQLite trigger registration) until the
host actually imports the api facade.
"""
