"""Module-level configuration for xp-excel-toolkit.

Downstream callers (CLI tools, application code built on this toolkit)
override these before invoking any toolkit function.

Example::

    from xp_excel_toolkit import config

    config.CACHE_DIR = "/var/cache/myapp"
    config.LIBREOFFICE_PATH = "/usr/bin/soffice"
"""

from __future__ import annotations

from pathlib import Path

# Cache directory for converted xlsx + imported DB files.
# When None, falls back to <cwd>/.xltk_cache/.
CACHE_DIR: str | Path | None = None

# LibreOffice executable path. When None, auto-detection runs
# (searches PATH then common install locations).
LIBREOFFICE_PATH: str | None = None
