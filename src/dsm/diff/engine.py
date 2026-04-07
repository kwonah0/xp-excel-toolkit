"""Diff engine — compare two DSM databases."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Callable
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime
from pathlib import Path

from sqlalchemy import insert
from sqlalchemy.orm import Session

from dsm.domain_models import REGMAP_FIELD_MAP, Register, MemoryMapEntry
from dsm.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook, init_db

from dsm.diff.models import (
    DiffBase, DiffCell, DiffMemmap, DiffMeta, DiffRegister, DiffResult,
    _MEMMAP_FIELDS, _REG_FIELDS, init_diff_db,
)


# ── Helpers ────────────────────────────────────────────────────────

def _reg_changes(dr: DiffRegister) -> list[tuple[str, str | None, str | None]]:
    """Return [(field, old, new), ...] for fields that differ in a DiffRegister."""
    return [
        (f, getattr(dr, f"old_{f}"), getattr(dr, f"new_{f}"))
        for f in _REG_FIELDS
        if getattr(dr, f"old_{f}") != getattr(dr, f"new_{f}")
    ]


def _mm_changes(dm: DiffMemmap) -> list[tuple[str, str | None, str | None]]:
    """Return [(field, old, new), ...] for fields that differ in a DiffMemmap."""
    return [
        (f, getattr(dm, f"old_{f}"), getattr(dm, f"new_{f}"))
        for f in _MEMMAP_FIELDS
        if getattr(dm, f"old_{f}") != getattr(dm, f"new_{f}")
    ]


def _reg_key(sheet_name: str, reg: Register) -> tuple:
    """Unique key for a register: (sheet, name, indx, page, para)."""
    return (sheet_name, reg.name, reg.indx, reg.page, reg.para)


def _memmap_key(entry: MemoryMapEntry) -> tuple:
    """Unique key for a memmap entry: (baseaddr, group)."""
    return (entry.baseaddr, entry.group)


# ── Loaders ────────────────────────────────────────────────────────

def _load_registers(session: Session) -> dict[tuple, tuple[str, Register]]:
    """Load all registers keyed by (sheet_name, name, indx, page, para)."""
    regs = {}
    sheets_by_id: dict[int, str] = {}
    for sheet in session.query(ExcelSheet).all():
        sheets_by_id[sheet.id] = sheet.name

    for reg in session.query(Register).order_by(Register.sheet_id, Register.excel_row).all():
        sn = sheets_by_id.get(reg.sheet_id, "?")
        key = _reg_key(sn, reg)
        regs[key] = (sn, reg)
    return regs


def _load_memmap(session: Session) -> dict[tuple, MemoryMapEntry]:
    """Load all memmap entries keyed by (baseaddr, group)."""
    entries = {}
    for entry in session.query(MemoryMapEntry).order_by(MemoryMapEntry.excel_row).all():
        key = _memmap_key(entry)
        entries[key] = entry
    return entries


def _load_cells(session: Session) -> dict[tuple, ExcelCell]:
    """Load all cells keyed by (sheet_name, row, col)."""
    cells: dict[tuple, ExcelCell] = {}
    sheets_by_id: dict[int, str] = {}
    for sheet in session.query(ExcelSheet).all():
        sheets_by_id[sheet.id] = sheet.name

    for cell in session.query(ExcelCell).all():
        sn = sheets_by_id.get(cell.sheet_id, "?")
        cells[(sn, cell.row, cell.col)] = cell
    return cells


def _load_merge_ranges(session: Session) -> dict[int, str]:
    """Load merge ranges keyed by merge_id → 'R{min}C{min}:R{max}C{max}'."""
    ranges: dict[int, str] = {}
    for m in session.query(ExcelMerge).all():
        ranges[m.id] = f"R{m.min_row}C{m.min_col}:R{m.max_row}C{m.max_col}"
    return ranges


def _load_cells_by_sheet(
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

    # sheet_name -> row_num -> col -> cell
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


def _row_signature(cols: dict[int, ExcelCell]) -> tuple:
    """Convert a row's cells into a hashable tuple for SequenceMatcher."""
    if not cols:
        return ()
    max_col = max(cols)
    return tuple(
        (cols[c].raw_value if c in cols else None)
        for c in range(1, max_col + 1)
    )


# ── Smart diff ─────────────────────────────────────────────────────

def _diff_cells_smart(
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
    import difflib
    import json as _json

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
        return extras

    for sheet in all_sheets:
        rows_a = sheet_rows_a.get(sheet, [])
        rows_b = sheet_rows_b.get(sheet, [])

        if on_progress:
            on_progress(f"  Smart diff: {sheet} ({len(rows_a)} vs {len(rows_b)} rows)")

        # Build signature sequences for SequenceMatcher
        sigs_a = [_row_signature(cols) for (_row_num, cols) in rows_a]
        sigs_b = [_row_signature(cols) for (_row_num, cols) in rows_b]

        sm = difflib.SequenceMatcher(None, sigs_a, sigs_b, autojunk=False)

        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                # Values are identical, but comment/style/merge may differ
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
                                old_value=ca.raw_value, new_value=cb.raw_value,
                                old_comment=ea.get("comment"), new_comment=eb.get("comment"),
                                old_style=ea.get("style"), new_style=eb.get("style"),
                                old_merge_range=ea.get("merge_range"), new_merge_range=eb.get("merge_range"),
                            ))
                continue

            elif tag == "delete":
                # Rows in old that are not in new (deleted rows)
                for idx in range(i1, i2):
                    old_row_num, old_cols = rows_a[idx]
                    for col, ca in sorted(old_cols.items()):
                        ea = _cell_extras(ca, merges_a)
                        diffs.append(DiffCell(
                            status="removed", sheet=sheet,
                            row=old_row_num, col=col,
                            old_row=old_row_num, new_row=None,
                            old_value=ca.raw_value, new_value=None,
                            old_comment=ea.get("comment"), new_comment=None,
                            old_style=ea.get("style"), new_style=None,
                            old_merge_range=ea.get("merge_range"), new_merge_range=None,
                        ))

            elif tag == "insert":
                # Rows in new that are not in old (inserted rows)
                for idx in range(j1, j2):
                    new_row_num, new_cols = rows_b[idx]
                    for col, cb in sorted(new_cols.items()):
                        eb = _cell_extras(cb, merges_b)
                        diffs.append(DiffCell(
                            status="added", sheet=sheet,
                            row=new_row_num, col=col,
                            old_row=None, new_row=new_row_num,
                            old_value=None, new_value=cb.raw_value,
                            old_comment=None, new_comment=eb.get("comment"),
                            old_style=None, new_style=eb.get("style"),
                            old_merge_range=None, new_merge_range=eb.get("merge_range"),
                        ))

            elif tag == "replace":
                # Rows exist in both but differ — compare cell by cell
                old_block = rows_a[i1:i2]
                new_block = rows_b[j1:j2]
                # Pair up rows: zip for matched length, remainder as add/remove
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
                                old_value=None, new_value=cb.raw_value,
                                old_comment=None, new_comment=eb.get("comment"),
                                old_style=None, new_style=eb.get("style"),
                                old_merge_range=None, new_merge_range=eb.get("merge_range"),
                            ))
                        elif ca is not None and cb is None:
                            ea = _cell_extras(ca, merges_a)
                            diffs.append(DiffCell(
                                status="removed", sheet=sheet,
                                row=old_row_num, col=col,
                                old_row=old_row_num, new_row=new_row_num,
                                old_value=ca.raw_value, new_value=None,
                                old_comment=ea.get("comment"), new_comment=None,
                                old_style=ea.get("style"), new_style=None,
                                old_merge_range=ea.get("merge_range"), new_merge_range=None,
                            ))
                        elif ca is not None and cb is not None:
                            # Both exist — check for changes
                            val_diff = ca.raw_value != cb.raw_value
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
                                    old_value=ca.raw_value, new_value=cb.raw_value,
                                    old_comment=ea.get("comment"), new_comment=eb.get("comment"),
                                    old_style=ea.get("style"), new_style=eb.get("style"),
                                    old_merge_range=ea.get("merge_range"), new_merge_range=eb.get("merge_range"),
                                ))

                # Extra old rows (deleted)
                for k in range(paired, len(old_block)):
                    old_row_num, old_cols = old_block[k]
                    for col, ca in sorted(old_cols.items()):
                        ea = _cell_extras(ca, merges_a)
                        diffs.append(DiffCell(
                            status="removed", sheet=sheet,
                            row=old_row_num, col=col,
                            old_row=old_row_num, new_row=None,
                            old_value=ca.raw_value, new_value=None,
                            old_comment=ea.get("comment"), new_comment=None,
                            old_style=ea.get("style"), new_style=None,
                            old_merge_range=ea.get("merge_range"), new_merge_range=None,
                        ))

                # Extra new rows (added)
                for k in range(paired, len(new_block)):
                    new_row_num, new_cols = new_block[k]
                    for col, cb in sorted(new_cols.items()):
                        eb = _cell_extras(cb, merges_b)
                        diffs.append(DiffCell(
                            status="added", sheet=sheet,
                            row=new_row_num, col=col,
                            old_row=None, new_row=new_row_num,
                            old_value=None, new_value=cb.raw_value,
                            old_comment=None, new_comment=eb.get("comment"),
                            old_style=None, new_style=eb.get("style"),
                            old_merge_range=None, new_merge_range=eb.get("merge_range"),
                        ))

    return diffs


# ── Main diff orchestrator ─────────────────────────────────────────

def diff_databases(
    db_path_a: Path,
    db_path_b: Path,
    on_progress: Callable[[str], None] | None = None,
    include_cells: bool = True,
    include_domain: bool = False,
    compare_comment: bool = False,
    compare_style: bool = False,
    compare_merge: bool = False,
    smart: bool = True,
) -> DiffResult:
    """Compare two SQLite databases and return differences.

    Args:
        db_path_a: Path to the "old" / base database.
        db_path_b: Path to the "new" / target database.
        include_cells: Include cell-level diff (default True).
        include_domain: Include domain-level diff — Register and MemoryMap (default False).

    Returns:
        DiffResult with added, removed, and changed items.
    """
    if on_progress:
        on_progress(f"Loading DB: {db_path_a.name}")
    SessionA = init_db(f"sqlite:///{db_path_a}")
    if on_progress:
        on_progress(f"Loading DB: {db_path_b.name}")
    SessionB = init_db(f"sqlite:///{db_path_b}")

    result = DiffResult()

    def _sort_key(k: tuple) -> tuple:
        return tuple(v if v is not None else "" for v in k)

    with SessionA() as sa, SessionB() as sb:
        # --- Domain-level diff (optional) ---
        if include_domain:
            if on_progress:
                on_progress("Loading registers from old DB...")
            regs_a = _load_registers(sa)
            if on_progress:
                on_progress(f"  {len(regs_a)} registers loaded")
                on_progress("Loading registers from new DB...")
            regs_b = _load_registers(sb)
            if on_progress:
                on_progress(f"  {len(regs_b)} registers loaded")
                on_progress("Comparing registers...")

            keys_a = set(regs_a.keys())
            keys_b = set(regs_b.keys())

            def _make_diff_reg(status: str, sheet: str,
                               old: Register | None, new: Register | None) -> DiffRegister:
                dr = DiffRegister(status=status, sheet=sheet)
                for f in _REG_FIELDS:
                    setattr(dr, f"old_{f}", getattr(old, f) if old else None)
                    setattr(dr, f"new_{f}", getattr(new, f) if new else None)
                return dr

            # Added
            for key in sorted(keys_b - keys_a, key=_sort_key):
                sn, reg = regs_b[key]
                result.registers.append(_make_diff_reg("added", sn, None, reg))

            # Removed
            for key in sorted(keys_a - keys_b, key=_sort_key):
                sn, reg = regs_a[key]
                result.registers.append(_make_diff_reg("removed", sn, reg, None))

            # Changed
            for key in sorted(keys_a & keys_b, key=_sort_key):
                sn_a, reg_a = regs_a[key]
                _sn_b, reg_b = regs_b[key]
                has_change = any(
                    getattr(reg_a, f) != getattr(reg_b, f) for f in _REG_FIELDS
                )
                if has_change:
                    result.registers.append(
                        _make_diff_reg("changed", sn_a, reg_a, reg_b)
                    )

            # --- MemoryMap diff ---
            if on_progress:
                on_progress("Comparing memory map...")
            mm_a = _load_memmap(sa)
            mm_b = _load_memmap(sb)

            mkeys_a = set(mm_a.keys())
            mkeys_b = set(mm_b.keys())

            def _make_diff_mm(status: str,
                              old: MemoryMapEntry | None,
                              new: MemoryMapEntry | None) -> DiffMemmap:
                dm = DiffMemmap(status=status)
                for f in _MEMMAP_FIELDS:
                    setattr(dm, f"old_{f}", getattr(old, f) if old else None)
                    setattr(dm, f"new_{f}", getattr(new, f) if new else None)
                return dm

            for key in sorted(mkeys_b - mkeys_a, key=_sort_key):
                result.memmap.append(_make_diff_mm("added", None, mm_b[key]))

            for key in sorted(mkeys_a - mkeys_b, key=_sort_key):
                result.memmap.append(_make_diff_mm("removed", mm_a[key], None))

            for key in sorted(mkeys_a & mkeys_b, key=_sort_key):
                ea, eb = mm_a[key], mm_b[key]
                has_change = any(
                    getattr(ea, f) != getattr(eb, f) for f in _MEMMAP_FIELDS
                )
                if has_change:
                    result.memmap.append(_make_diff_mm("changed", ea, eb))

        # --- Cell-level diff (default) ---
        if include_cells:
            # Load merge ranges if needed
            merges_a: dict[int, str] = {}
            merges_b: dict[int, str] = {}
            if compare_merge:
                if on_progress:
                    on_progress("Loading merge ranges...")
                merges_a = _load_merge_ranges(sa)
                merges_b = _load_merge_ranges(sb)

            if smart:
                # Sequence-based smart diff — avoids cascade on row insert/delete
                if on_progress:
                    on_progress("Loading cells by sheet (smart mode)...")
                sheet_rows_a = _load_cells_by_sheet(sa)
                sheet_rows_b = _load_cells_by_sheet(sb)
                total_a = sum(len(rows) for rows in sheet_rows_a.values())
                total_b = sum(len(rows) for rows in sheet_rows_b.values())
                if on_progress:
                    on_progress(f"  {total_a} rows (old), {total_b} rows (new)")
                    on_progress("Running sequence alignment diff...")

                result.cells = _diff_cells_smart(
                    sheet_rows_a, sheet_rows_b,
                    compare_comment=compare_comment,
                    compare_style=compare_style,
                    compare_merge=compare_merge,
                    merges_a=merges_a, merges_b=merges_b,
                    on_progress=on_progress,
                )
            else:
                # Positional diff — original (sheet, row, col) key comparison
                import json as _json

                if on_progress:
                    on_progress("Loading cells from old DB...")
                cells_a = _load_cells(sa)
                if on_progress:
                    on_progress(f"  {len(cells_a)} cells loaded")
                    on_progress("Loading cells from new DB...")
                cells_b = _load_cells(sb)
                if on_progress:
                    on_progress(f"  {len(cells_b)} cells loaded")
                    on_progress("Comparing cells...")

                ckeys_a = set(cells_a.keys())
                ckeys_b = set(cells_b.keys())

                def _cell_extras(cell: ExcelCell, merges: dict[int, str]) -> dict[str, str | None]:
                    """Build optional comment/style/merge fields for a DiffCell."""
                    extras: dict[str, str | None] = {}
                    if compare_comment:
                        extras["comment"] = cell.comment
                    if compare_style:
                        extras["style"] = _json.dumps(cell.style, ensure_ascii=False) if cell.style else None
                    if compare_merge:
                        extras["merge_range"] = merges.get(cell.merge_id) if cell.merge_id else None
                    return extras

                # Added cells
                for key in sorted(ckeys_b - ckeys_a, key=_sort_key):
                    cb = cells_b[key]
                    eb = _cell_extras(cb, merges_b)
                    result.cells.append(DiffCell(
                        status="added",
                        sheet=key[0], row=key[1], col=key[2],
                        old_value=None, new_value=cb.raw_value,
                        old_comment=None, new_comment=eb.get("comment"),
                        old_style=None, new_style=eb.get("style"),
                        old_merge_range=None, new_merge_range=eb.get("merge_range"),
                    ))

                # Removed cells
                for key in sorted(ckeys_a - ckeys_b, key=_sort_key):
                    ca = cells_a[key]
                    ea = _cell_extras(ca, merges_a)
                    result.cells.append(DiffCell(
                        status="removed",
                        sheet=key[0], row=key[1], col=key[2],
                        old_value=ca.raw_value, new_value=None,
                        old_comment=ea.get("comment"), new_comment=None,
                        old_style=ea.get("style"), new_style=None,
                        old_merge_range=ea.get("merge_range"), new_merge_range=None,
                    ))

                # Changed cells
                for key in sorted(ckeys_a & ckeys_b, key=_sort_key):
                    ca, cb = cells_a[key], cells_b[key]

                    # Check what changed
                    val_diff = ca.raw_value != cb.raw_value
                    comment_diff = compare_comment and ca.comment != cb.comment
                    style_diff = compare_style and ca.style != cb.style
                    merge_a = merges_a.get(ca.merge_id) if (compare_merge and ca.merge_id) else None
                    merge_b = merges_b.get(cb.merge_id) if (compare_merge and cb.merge_id) else None
                    merge_diff = compare_merge and merge_a != merge_b

                    if val_diff or comment_diff or style_diff or merge_diff:
                        ea = _cell_extras(ca, merges_a)
                        eb = _cell_extras(cb, merges_b)
                        result.cells.append(DiffCell(
                            status="changed",
                            sheet=key[0], row=key[1], col=key[2],
                            old_value=ca.raw_value, new_value=cb.raw_value,
                            old_comment=ea.get("comment"), new_comment=eb.get("comment"),
                            old_style=ea.get("style"), new_style=eb.get("style"),
                            old_merge_range=ea.get("merge_range"), new_merge_range=eb.get("merge_range"),
                        ))

            if on_progress:
                on_progress(f"  {len(result.cells)} cell differences found")

    return result


# ── Save diff to DB ────────────────────────────────────────────────

def _orm_to_dict(obj: object, table_cls: type) -> dict:
    """Extract column values from an ORM object as a dict (excluding 'id')."""
    return {
        c.name: getattr(obj, c.name)
        for c in table_cls.__table__.columns
        if c.name != "id"
    }


def save_diff_to_db(
    result: DiffResult,
    diff_db_path: Path,
    old_path: Path,
    new_path: Path,
) -> Path:
    """Save DiffResult into a SQLite DB for querying."""
    _BULK_CHUNK = 500

    added_regs = result._filter_regs("added")
    removed_regs = result._filter_regs("removed")
    changed_regs = result._filter_regs("changed")
    added_mm = result._filter_mm("added")
    removed_mm = result._filter_mm("removed")
    changed_mm = result._filter_mm("changed")

    DiffSession = init_diff_db(f"sqlite:///{diff_db_path}")

    with DiffSession() as session:
        meta = DiffMeta(
            old_path=str(old_path),
            new_path=str(new_path),
            created_at=datetime.now().isoformat(),
            added_regs=len(added_regs),
            removed_regs=len(removed_regs),
            changed_regs=len(changed_regs),
            added_memmap=len(added_mm),
            removed_memmap=len(removed_mm),
            changed_memmap=len(changed_mm),
        )
        session.add(meta)
        session.flush()

        # Registers
        if result.registers:
            reg_rows = [_orm_to_dict(r, DiffRegister) for r in result.registers]
            for i in range(0, len(reg_rows), _BULK_CHUNK):
                session.execute(insert(DiffRegister), reg_rows[i:i + _BULK_CHUNK])

        # MemoryMap
        if result.memmap:
            mm_rows = [_orm_to_dict(m, DiffMemmap) for m in result.memmap]
            for i in range(0, len(mm_rows), _BULK_CHUNK):
                session.execute(insert(DiffMemmap), mm_rows[i:i + _BULK_CHUNK])

        # Cells
        if result.cells:
            cell_rows = [_orm_to_dict(c, DiffCell) for c in result.cells]
            for i in range(0, len(cell_rows), _BULK_CHUNK):
                session.execute(insert(DiffCell), cell_rows[i:i + _BULK_CHUNK])

        session.commit()

    return diff_db_path


# ── Auto-import ────────────────────────────────────────────────────

def diff_with_auto_import(
    path_a: Path,
    path_b: Path,
    on_progress: Callable[[str], None] | None = None,
    include_cells: bool = True,
    include_domain: bool = False,
    compare_comment: bool = False,
    compare_style: bool = False,
    compare_merge: bool = False,
    smart: bool = True,
) -> DiffResult:
    """Diff two paths that can be .db or .xlsx files.

    If an xlsx is given, it is imported into a temp DB first.
    When both inputs need import, they are imported in parallel.
    """
    needs_a = path_a.suffix in (".xlsx", ".xls") and not path_a.with_suffix(".db").exists()
    needs_b = path_b.suffix in (".xlsx", ".xls") and not path_b.with_suffix(".db").exists()

    if needs_a and needs_b:
        if on_progress:
            on_progress(f"Auto-importing {path_a.name} and {path_b.name} in parallel...")
        with ProcessPoolExecutor(max_workers=2) as pool:
            fut_a = pool.submit(_resolve_db_worker, path_a)
            fut_b = pool.submit(_resolve_db_worker, path_b)
            db_a = fut_a.result()
            db_b = fut_b.result()
    else:
        db_a = _resolve_db(path_a, on_progress=on_progress)
        db_b = _resolve_db(path_b, on_progress=on_progress)

    return diff_databases(db_a, db_b, on_progress=on_progress,
                          include_cells=include_cells,
                          include_domain=include_domain,
                          compare_comment=compare_comment,
                          compare_style=compare_style,
                          compare_merge=compare_merge,
                          smart=smart)


def _resolve_db_worker(path: Path) -> Path:
    """Pickle-safe version of _resolve_db for multiprocessing (no callbacks)."""
    return _resolve_db(path)


def _resolve_db(
    path: Path,
    on_progress: Callable[[str], None] | None = None,
) -> Path:
    """If path is .xlsx, import to temp DB and return DB path.
    If .db, return as-is.
    """
    if path.suffix == ".db":
        return path

    if path.suffix in (".xlsx", ".xls"):
        # Check if a companion .db already exists
        companion_db = path.with_suffix(".db")
        if companion_db.exists():
            if on_progress:
                on_progress(f"Reusing existing DB: {companion_db.name}")
            return companion_db

        # Import into a new DB next to the xlsx
        if on_progress:
            on_progress(f"Auto-importing {path.name} into DB...")
        from dsm.xlsx_parser import import_xlsx
        Session = init_db(f"sqlite:///{companion_db}")
        with Session() as session:
            import_xlsx(session, path, on_progress=on_progress)
            session.commit()
        return companion_db

    raise ValueError(f"Unsupported file type: {path.suffix} (expected .db or .xlsx)")
