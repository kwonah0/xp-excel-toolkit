"""Smart cell-level diff between two SQLite databases of ExcelCells.

Domain-agnostic. Uses difflib.SequenceMatcher on row signatures so that
insertions/deletions don't cascade into false positives. Cells that were
removed in one place and added in another with identical content are
post-classified as "moved".
"""

from __future__ import annotations

import difflib
import json as _json
from collections import defaultdict
from collections.abc import Callable

from sqlalchemy.orm import Session

from xp_excel_toolkit.diff.models import DiffCell
from xp_excel_toolkit.models import ExcelCell, ExcelMerge, ExcelSheet


# ── Loaders ────────────────────────────────────────────────────────

def load_cells(session: Session) -> dict[tuple, ExcelCell]:
    """Load all cells keyed by (sheet_name, row, col)."""
    cells: dict[tuple, ExcelCell] = {}
    sheets_by_id: dict[int, str] = {}
    for sheet in session.query(ExcelSheet).all():
        sheets_by_id[sheet.id] = sheet.name

    for cell in session.query(ExcelCell).all():
        sn = sheets_by_id.get(cell.sheet_id, "?")
        cells[(sn, cell.row, cell.col)] = cell
    return cells


def load_merge_ranges(session: Session) -> dict[int, str]:
    """Load merge ranges keyed by merge_id → 'R{min}C{min}:R{max}C{max}'."""
    ranges: dict[int, str] = {}
    for m in session.query(ExcelMerge).all():
        ranges[m.id] = f"R{m.min_row}C{m.min_col}:R{m.max_row}C{m.max_col}"
    return ranges


def load_cells_by_sheet(
    session: Session,
) -> dict[str, list[tuple[int, dict[int, ExcelCell]]]]:
    """Load cells grouped by sheet, ordered by row.

    Returns:
        {sheet_name: [(row_num, {col: ExcelCell}), ...]}
        Rows are sorted ascending by row number.
    """
    sheets_by_id: dict[int, str] = {}
    for sheet in session.query(ExcelSheet).all():
        sheets_by_id[sheet.id] = sheet.name

    nested: dict[str, dict[int, dict[int, ExcelCell]]] = defaultdict(
        lambda: defaultdict(dict)
    )
    for cell in session.query(ExcelCell).all():
        sn = sheets_by_id.get(cell.sheet_id, "?")
        nested[sn][cell.row][cell.col] = cell

    result: dict[str, list[tuple[int, dict[int, ExcelCell]]]] = {}
    for sn, rows_dict in nested.items():
        result[sn] = [(r, rows_dict[r]) for r in sorted(rows_dict)]
    return result


# ── Cell value helpers ─────────────────────────────────────────────

def cell_display_value(cell: ExcelCell) -> str | None:
    """Return effective display value: cached_value if available, else raw_value."""
    return cell.cached_value if cell.cached_value is not None else cell.raw_value


def cell_formula(cell: ExcelCell) -> str | None:
    """Return formula string if cell has one (indicated by cached_value existing)."""
    if cell.cached_value is not None and cell.raw_value and str(cell.raw_value).startswith("="):
        return cell.raw_value
    return None


def row_signature(cols: dict[int, ExcelCell]) -> tuple:
    """Convert a row's cells into a hashable tuple for SequenceMatcher."""
    if not cols:
        return ()
    max_col = max(cols)
    return tuple(
        (cell_display_value(cols[c]) if c in cols else None)
        for c in range(1, max_col + 1)
    )


# ── Smart diff ─────────────────────────────────────────────────────

def diff_cells(
    sheet_rows_a: dict[str, list[tuple[int, dict[int, ExcelCell]]]],
    sheet_rows_b: dict[str, list[tuple[int, dict[int, ExcelCell]]]],
    compare_comment: bool = False,
    compare_style: bool = False,
    compare_merge: bool = False,
    merges_a: dict[int, str] | None = None,
    merges_b: dict[int, str] | None = None,
    on_progress: Callable[[str], None] | None = None,
) -> list[DiffCell]:
    """Sequence-based smart cell diff using difflib.SequenceMatcher.

    Instead of positional (sheet, row, col) comparison, aligns rows by content
    so that insertions/deletions don't cascade into false positives.
    """
    if merges_a is None:
        merges_a = {}
    if merges_b is None:
        merges_b = {}

    diffs: list[DiffCell] = []
    all_sheets = sorted(set(sheet_rows_a) | set(sheet_rows_b))

    def _cell_extras(cell: ExcelCell, merges: dict[int, str]) -> dict[str, str | None]:
        extras: dict[str, str | None] = {}
        if compare_comment:
            extras["comment"] = cell.comment
        if compare_style:
            extras["style"] = _json.dumps(cell.style, ensure_ascii=False) if cell.style else None
        if compare_merge:
            extras["merge_range"] = merges.get(cell.merge_id) if cell.merge_id else None
        extras["formula"] = cell_formula(cell)
        return extras

    for sheet in all_sheets:
        rows_a = sheet_rows_a.get(sheet, [])
        rows_b = sheet_rows_b.get(sheet, [])

        if on_progress:
            on_progress(f"  Smart diff: {sheet} ({len(rows_a)} vs {len(rows_b)} rows)")

        sigs_a = [row_signature(cols) for (_row_num, cols) in rows_a]
        sigs_b = [row_signature(cols) for (_row_num, cols) in rows_b]

        sm = difflib.SequenceMatcher(None, sigs_a, sigs_b, autojunk=False)

        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                if not (compare_comment or compare_style or compare_merge):
                    continue
                for idx_a, idx_b in zip(range(i1, i2), range(j1, j2)):
                    old_row_num, old_cols = rows_a[idx_a]
                    new_row_num, new_cols = rows_b[idx_b]
                    all_cols = sorted(set(old_cols) | set(new_cols))
                    for col in all_cols:
                        ca = old_cols.get(col)
                        cb = new_cols.get(col)
                        if ca is None or cb is None:
                            continue
                        comment_diff = compare_comment and ca.comment != cb.comment
                        style_diff = compare_style and ca.style != cb.style
                        mr_a = merges_a.get(ca.merge_id) if (compare_merge and ca.merge_id) else None
                        mr_b = merges_b.get(cb.merge_id) if (compare_merge and cb.merge_id) else None
                        merge_diff = compare_merge and mr_a != mr_b
                        if comment_diff or style_diff or merge_diff:
                            ea = _cell_extras(ca, merges_a)
                            eb = _cell_extras(cb, merges_b)
                            diffs.append(DiffCell(
                                status="changed", sheet=sheet,
                                row=new_row_num, col=col,
                                old_row=old_row_num, new_row=new_row_num,
                                old_value=cell_display_value(ca), new_value=cell_display_value(cb),
                                old_comment=ea.get("comment"), new_comment=eb.get("comment"),
                                old_style=ea.get("style"), new_style=eb.get("style"),
                                old_merge_range=ea.get("merge_range"), new_merge_range=eb.get("merge_range"),
                                old_formula=ea.get("formula"), new_formula=eb.get("formula"),
                            ))
                continue

            elif tag == "delete":
                for idx in range(i1, i2):
                    old_row_num, old_cols = rows_a[idx]
                    for col, ca in sorted(old_cols.items()):
                        ea = _cell_extras(ca, merges_a)
                        diffs.append(DiffCell(
                            status="removed", sheet=sheet,
                            row=old_row_num, col=col,
                            old_row=old_row_num, new_row=None,
                            old_value=cell_display_value(ca), new_value=None,
                            old_comment=ea.get("comment"), new_comment=None,
                            old_style=ea.get("style"), new_style=None,
                            old_merge_range=ea.get("merge_range"), new_merge_range=None,
                            old_formula=ea.get("formula"), new_formula=None,
                        ))

            elif tag == "insert":
                for idx in range(j1, j2):
                    new_row_num, new_cols = rows_b[idx]
                    for col, cb in sorted(new_cols.items()):
                        eb = _cell_extras(cb, merges_b)
                        diffs.append(DiffCell(
                            status="added", sheet=sheet,
                            row=new_row_num, col=col,
                            old_row=None, new_row=new_row_num,
                            old_value=None, new_value=cell_display_value(cb),
                            old_comment=None, new_comment=eb.get("comment"),
                            old_style=None, new_style=eb.get("style"),
                            old_merge_range=None, new_merge_range=eb.get("merge_range"),
                            old_formula=None, new_formula=eb.get("formula"),
                        ))

            elif tag == "replace":
                old_block = rows_a[i1:i2]
                new_block = rows_b[j1:j2]
                paired = min(len(old_block), len(new_block))

                for k in range(paired):
                    old_row_num, old_cols = old_block[k]
                    new_row_num, new_cols = new_block[k]
                    all_cols = sorted(set(old_cols) | set(new_cols))

                    for col in all_cols:
                        ca = old_cols.get(col)
                        cb = new_cols.get(col)

                        if ca is None and cb is not None:
                            eb = _cell_extras(cb, merges_b)
                            diffs.append(DiffCell(
                                status="added", sheet=sheet,
                                row=new_row_num, col=col,
                                old_row=old_row_num, new_row=new_row_num,
                                old_value=None, new_value=cell_display_value(cb),
                                old_comment=None, new_comment=eb.get("comment"),
                                old_style=None, new_style=eb.get("style"),
                                old_merge_range=None, new_merge_range=eb.get("merge_range"),
                                old_formula=None, new_formula=eb.get("formula"),
                            ))
                        elif ca is not None and cb is None:
                            ea = _cell_extras(ca, merges_a)
                            diffs.append(DiffCell(
                                status="removed", sheet=sheet,
                                row=old_row_num, col=col,
                                old_row=old_row_num, new_row=new_row_num,
                                old_value=cell_display_value(ca), new_value=None,
                                old_comment=ea.get("comment"), new_comment=None,
                                old_style=ea.get("style"), new_style=None,
                                old_merge_range=ea.get("merge_range"), new_merge_range=None,
                                old_formula=ea.get("formula"), new_formula=None,
                            ))
                        elif ca is not None and cb is not None:
                            val_diff = cell_display_value(ca) != cell_display_value(cb)
                            comment_diff = compare_comment and ca.comment != cb.comment
                            style_diff = compare_style and ca.style != cb.style
                            mr_a = merges_a.get(ca.merge_id) if (compare_merge and ca.merge_id) else None
                            mr_b = merges_b.get(cb.merge_id) if (compare_merge and cb.merge_id) else None
                            merge_diff = compare_merge and mr_a != mr_b

                            if val_diff or comment_diff or style_diff or merge_diff:
                                ea = _cell_extras(ca, merges_a)
                                eb = _cell_extras(cb, merges_b)
                                diffs.append(DiffCell(
                                    status="changed", sheet=sheet,
                                    row=new_row_num, col=col,
                                    old_row=old_row_num, new_row=new_row_num,
                                    old_value=cell_display_value(ca), new_value=cell_display_value(cb),
                                    old_comment=ea.get("comment"), new_comment=eb.get("comment"),
                                    old_style=ea.get("style"), new_style=eb.get("style"),
                                    old_merge_range=ea.get("merge_range"), new_merge_range=eb.get("merge_range"),
                                    old_formula=ea.get("formula"), new_formula=eb.get("formula"),
                                ))

                for k in range(paired, len(old_block)):
                    old_row_num, old_cols = old_block[k]
                    for col, ca in sorted(old_cols.items()):
                        ea = _cell_extras(ca, merges_a)
                        diffs.append(DiffCell(
                            status="removed", sheet=sheet,
                            row=old_row_num, col=col,
                            old_row=old_row_num, new_row=None,
                            old_value=cell_display_value(ca), new_value=None,
                            old_comment=ea.get("comment"), new_comment=None,
                            old_style=ea.get("style"), new_style=None,
                            old_merge_range=ea.get("merge_range"), new_merge_range=None,
                            old_formula=ea.get("formula"), new_formula=None,
                        ))

                for k in range(paired, len(new_block)):
                    new_row_num, new_cols = new_block[k]
                    for col, cb in sorted(new_cols.items()):
                        eb = _cell_extras(cb, merges_b)
                        diffs.append(DiffCell(
                            status="added", sheet=sheet,
                            row=new_row_num, col=col,
                            old_row=None, new_row=new_row_num,
                            old_value=None, new_value=cell_display_value(cb),
                            old_comment=None, new_comment=eb.get("comment"),
                            old_style=None, new_style=eb.get("style"),
                            old_merge_range=None, new_merge_range=eb.get("merge_range"),
                            old_formula=None, new_formula=eb.get("formula"),
                        ))

    # ── Post-process: detect moved rows ──
    removed_rows: dict[tuple[str, int], dict] = {}
    added_rows: dict[tuple[str, int], dict] = {}
    for d in diffs:
        if d.status == "removed" and d.old_row is not None:
            key = (d.sheet, d.old_row)
            if key not in removed_rows:
                removed_rows[key] = {}
            removed_rows[key][d.col] = d.old_value
        elif d.status == "added" and d.new_row is not None:
            key = (d.sheet, d.new_row)
            if key not in added_rows:
                added_rows[key] = {}
            added_rows[key][d.col] = d.new_value

    def _cols_to_sig(cols: dict) -> tuple:
        if not cols:
            return ()
        max_col = max(cols)
        return tuple(cols.get(c) for c in range(1, max_col + 1))

    removed_sigs = {k: _cols_to_sig(v) for k, v in removed_rows.items()}
    added_sigs = {k: _cols_to_sig(v) for k, v in added_rows.items()}

    added_by_sheet_sig: dict[tuple[str, tuple], list[tuple[str, int]]] = defaultdict(list)
    for key, sig in added_sigs.items():
        added_by_sheet_sig[(key[0], sig)].append(key)

    moved_pairs: dict[tuple[str, int], int] = {}
    used_added: set[tuple[str, int]] = set()

    for rm_key, rm_sig in removed_sigs.items():
        sheet = rm_key[0]
        candidates = added_by_sheet_sig.get((sheet, rm_sig), [])
        for add_key in candidates:
            if add_key not in used_added:
                moved_pairs[rm_key] = add_key[1]
                used_added.add(add_key)
                break

    if moved_pairs:
        for d in diffs:
            if d.status == "removed" and d.old_row is not None:
                rm_key = (d.sheet, d.old_row)
                if rm_key in moved_pairs:
                    d.status = "moved"
                    d.new_row = moved_pairs[rm_key]
                    d.new_value = d.old_value
                    d.new_formula = d.old_formula
                    d.row = moved_pairs[rm_key]
            elif d.status == "added" and d.new_row is not None:
                add_key = (d.sheet, d.new_row)
                if add_key in used_added:
                    d.status = "moved"
                    for rm_key, new_r in moved_pairs.items():
                        if rm_key[0] == d.sheet and new_r == d.new_row:
                            d.old_row = rm_key[1]
                            d.old_value = d.new_value
                            d.old_formula = d.new_formula
                            break

        seen_moved: set[tuple] = set()
        deduped: list[DiffCell] = []
        for d in diffs:
            if d.status == "moved":
                key = (d.sheet, d.old_row, d.col)
                if key in seen_moved:
                    continue
                seen_moved.add(key)
            deduped.append(d)
        diffs = deduped

    return diffs
