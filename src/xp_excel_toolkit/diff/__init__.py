"""Cell-level diff for two ExcelCell databases.

Domain-agnostic. Provides:
- DiffBase: shared SQLAlchemy declarative base for diff result tables
- DiffCell: one row per added/removed/changed/moved cell
- init_diff_db: create_all helper (creates every table registered on DiffBase)
- diff_cells: the smart cell diff algorithm
- load_cells_by_sheet, load_merge_ranges: helpers exposed for domain
  packages that want to reuse the loading layer.

Domain packages (e.g. dsm) define their own DiffRegister/DiffMemmap on the
same DiffBase. As long as those classes are imported before init_diff_db,
their tables are created in the same call.
"""

from xp_excel_toolkit.diff.models import (
    DiffBase,
    DiffCell,
    init_diff_db,
)
from xp_excel_toolkit.diff.engine import (
    cell_display_value,
    cell_formula,
    diff_cells,
    load_cells,
    load_cells_by_sheet,
    load_merge_ranges,
    row_signature,
)

__all__ = [
    "DiffBase",
    "DiffCell",
    "cell_display_value",
    "cell_formula",
    "diff_cells",
    "init_diff_db",
    "load_cells",
    "load_cells_by_sheet",
    "load_merge_ranges",
    "row_signature",
]
