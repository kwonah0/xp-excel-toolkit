"""DSM CLI — Design Specification Manager."""

from __future__ import annotations

import sys
import time
from pathlib import Path

import click

from dsm.models import init_db


def _default_db(xlsx_path: Path) -> Path:
    """Derive default DB path from xlsx: same dir, same stem + .db"""
    return xlsx_path.with_suffix(".db")


@click.group()
@click.version_option(version="0.2.0", prog_name="dsm")
def main():
    """DSM — Design Specification Manager for register map Excel files."""


# -- import -----------------------------------------------------------------

def _import_cmd(xlsx_path: Path, db_path: Path | None):
    if db_path is None:
        db_path = _default_db(xlsx_path)

    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    t0 = time.perf_counter()

    from dsm.xlsx_parser import import_xlsx
    with Session() as session:
        sheets = import_xlsx(session, xlsx_path, on_progress=click.echo)
        click.echo("Committing to DB...")
        session.commit()
        sheet_info = [(s.name, s.header_row) for s in sheets]

    elapsed = time.perf_counter() - t0
    click.echo(f"Imported {len(sheet_info)} sheets into {db_path} ({elapsed:.1f}s)")
    for name, header_row in sheet_info:
        click.echo(f"  - {name} (header_row={header_row})")


@click.command("import")
@click.argument("xlsx_path", type=click.Path(exists=True, path_type=Path))
@click.option("--db", "db_path", type=click.Path(path_type=Path), default=None,
              help="SQLite DB path (default: <xlsx_stem>.db)")
def import_cmd(xlsx_path: Path, db_path: Path | None):
    """Import all sheets from an xlsx file into SQLite DB."""
    _import_cmd(xlsx_path, db_path)


main.add_command(import_cmd)


# -- split ------------------------------------------------------------------

@main.command()
@click.argument("xlsx_path", type=click.Path(exists=True, path_type=Path), required=False, default=None)
@click.option("--db", "db_path", type=click.Path(path_type=Path), default=None,
              help="SQLite DB path (default: <xlsx_stem>.db)")
@click.option("--output-dir", type=click.Path(path_type=Path), default=None,
              help="Output directory (default: <xlsx_stem>_split/)")
@click.option("--parallel/--no-parallel", default=False,
              help="Use multiprocessing for parallel split")
@click.option("--force-reimport", is_flag=True, default=False,
              help="Re-import even if DB already contains this workbook")
def split(xlsx_path: Path | None, db_path: Path | None, output_dir: Path | None,
          parallel: bool, force_reimport: bool):
    """Split register map by IP — one output file per level2 sheet.

    If --db is given without XLSX_PATH, uses existing DB directly (no import).
    If XLSX_PATH is given, imports first (or reuses if already in DB).
    """
    if xlsx_path is None and db_path is None:
        raise click.UsageError("Either XLSX_PATH or --db must be provided.")

    if db_path is None:
        db_path = _default_db(xlsx_path)

    if output_dir is None:
        if xlsx_path is not None:
            output_dir = xlsx_path.parent / f"{xlsx_path.stem}_split"
        else:
            output_dir = db_path.parent / f"{db_path.stem}_split"

    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    t0 = time.perf_counter()

    with Session() as session:
        from dsm.models import ExcelWorkbook

        if xlsx_path is not None:
            existing = session.query(ExcelWorkbook).filter_by(
                filename=xlsx_path.name
            ).first()

            if existing and not force_reimport:
                click.echo(f"Reusing existing DB import for {xlsx_path.name}")
            else:
                click.echo(f"Importing {xlsx_path.name}...")
                from dsm.xlsx_parser import import_xlsx
                import_xlsx(session, xlsx_path, on_progress=click.echo)
                session.commit()
            workbook_name = xlsx_path.name
        else:
            # DB-only mode: find the first (or only) workbook in the DB
            wb_obj = session.query(ExcelWorkbook).first()
            if not wb_obj:
                raise click.UsageError(f"DB '{db_path}' contains no imported workbooks.")
            workbook_name = wb_obj.filename
            click.echo(f"Using existing DB: {db_path} (workbook={workbook_name})")

        click.echo(f"Splitting into {output_dir}...")
        if parallel:
            from dsm.parallel import parallel_split_regmap
            # parallel_split_regmap needs xlsx_path for .name lookup;
            # pass workbook_name as Path so .name works
            results = parallel_split_regmap(session, Path(workbook_name), output_dir)
        else:
            from dsm.splitter import split_regmap_from_db
            results = split_regmap_from_db(session, workbook_name, output_dir,
                                           on_progress=click.echo)

        session.commit()

    elapsed = time.perf_counter() - t0
    click.echo(f"Split into {len(results)} files ({elapsed:.1f}s):")
    for src_name, out_path in results.items():
        click.echo(f"  {src_name} -> {out_path}")


# -- query ------------------------------------------------------------------

@main.group()
def query():
    """Query domain objects from an existing DB."""


@query.command()
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
def sheets(db_path: Path):
    """List all sheets in the DB."""
    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    with Session() as session:
        from dsm.models import ExcelSheet, ExcelWorkbook
        for s in session.query(ExcelSheet).all():
            wb = session.get(ExcelWorkbook, s.workbook_id)
            click.echo(f"  [{s.id}] {s.name}  (workbook={wb.filename if wb else '?'}, header_row={s.header_row})")


@query.command()
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.option("--sheet", "sheet_name", default=None, help="Filter by sheet name pattern (e.g. level2_%)")
def ips(db_path: Path, sheet_name: str | None):
    """List distinct IP names with register counts."""
    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    with Session() as session:
        from sqlalchemy import func
        from dsm.domain_models import Register
        from dsm.models import ExcelSheet

        q = (
            session.query(ExcelSheet.name, Register.name, func.count(Register.id))
            .join(ExcelSheet, Register.sheet_id == ExcelSheet.id)
        )
        if sheet_name:
            q = q.filter(ExcelSheet.name.like(sheet_name.replace("*", "%")))

        results = q.group_by(ExcelSheet.name, Register.name).all()
        current_sheet = None
        for sheet_n, ip_n, count in results:
            if sheet_n != current_sheet:
                click.echo(f"\n  [{sheet_n}]")
                current_sheet = sheet_n
            click.echo(f"    {ip_n or '(unnamed)':>15}: {count} registers")


@query.command()
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.option("--sheet", "sheet_name", default=None, help="Filter by sheet name pattern")
@click.option("--ip", "ip_name", default=None, help="Filter by IP name")
@click.option("--json", "as_json", is_flag=True, default=False, help="Output as JSON")
def registers(db_path: Path, sheet_name: str | None, ip_name: str | None, as_json: bool):
    """Query Register rows with optional filters."""
    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    with Session() as session:
        from dsm.domain_models import Register
        from dsm.models import ExcelSheet

        q = session.query(Register)
        if sheet_name:
            sheet_ids = [
                s.id for s in session.query(ExcelSheet)
                .filter(ExcelSheet.name.like(sheet_name.replace("*", "%")))
                .all()
            ]
            q = q.filter(Register.sheet_id.in_(sheet_ids))
        if ip_name:
            q = q.filter(Register.name == ip_name)

        regs = q.order_by(Register.sheet_id, Register.excel_row).all()

        if as_json:
            import json
            data = [
                {
                    "sheet_id": r.sheet_id, "excel_row": r.excel_row,
                    "type": r.type, "indx": r.indx, "page": r.page,
                    "para": r.para, "name": r.name,
                    "d7": r.d7, "d6": r.d6, "d5": r.d5, "d4": r.d4,
                    "d3": r.d3, "d2": r.d2, "d1": r.d1, "d0": r.d0,
                    "init": r.init,
                }
                for r in regs
            ]
            click.echo(json.dumps(data, indent=2))
        else:
            click.echo(f"{'row':>4} {'TYPE':>4} {'INDX':>5} {'PAGE':>5} "
                       f"{'PARA':>5} {'NAME':>12} {'INIT':>6}")
            click.echo("-" * 50)
            for r in regs:
                click.echo(f"{r.excel_row or 0:>4} {r.type or '-':>4} "
                           f"{r.indx or '-':>5} {r.page or '-':>5} "
                           f"{r.para or '-':>5} {r.name or '-':>12} "
                           f"{r.init or '-':>6}")
        click.echo(f"\n{len(regs)} registers found.")


@query.command()
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
def memmap(db_path: Path):
    """Query MemoryMapEntry rows."""
    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    with Session() as session:
        from dsm.domain_models import MemoryMapEntry
        entries = session.query(MemoryMapEntry).order_by(MemoryMapEntry.excel_row).all()
        click.echo(f"{'BASEADDR':>10} {'Group':>12} {'midgroup':>10} "
                   f"{'Comment':<30} {'special':<10}")
        click.echo("-" * 80)
        for e in entries:
            click.echo(f"{e.baseaddr or '-':>10} {e.group or '-':>12} "
                       f"{e.midgroup or '-':>10} {e.comment or '-':<30} "
                       f"{e.special or '-':<10}")
        click.echo(f"\n{len(entries)} entries found.")


# -- diff -------------------------------------------------------------------

@main.command()
@click.argument("path_a", type=click.Path(exists=True, path_type=Path))
@click.argument("path_b", type=click.Path(exists=True, path_type=Path))
@click.option("--db", "diff_db_path", type=click.Path(path_type=Path), default=None,
              help="Save diff results to SQLite DB (default: diff_<a>_<b>.db)")
@click.option("--verbose", "-v", is_flag=True, default=False,
              help="Show detailed bit-field info for added registers")
@click.option("--json", "as_json", is_flag=True, default=False,
              help="Output as JSON")
@click.option("--cells", is_flag=True, default=False,
              help="Include cell-level diff (compares raw_value only by default)")
@click.option("--comment", "compare_comment", is_flag=True, default=False,
              help="Compare cell comments (implies --cells)")
@click.option("--style", "compare_style", is_flag=True, default=False,
              help="Compare cell styles (implies --cells)")
@click.option("--merge-info", "compare_merge", is_flag=True, default=False,
              help="Compare merged cell ranges (implies --cells)")
@click.option("--all", "compare_all", is_flag=True, default=False,
              help="Enable all cell comparisons (cells + comment + style + merge)")
@click.option("--smart", is_flag=True, default=False,
              help="Use sequence-based smart diff (handles row insert/delete without cascade)")
def diff(path_a: Path, path_b: Path, diff_db_path: Path | None, verbose: bool,
         as_json: bool, cells: bool, compare_comment: bool, compare_style: bool,
         compare_merge: bool, compare_all: bool, smart: bool):
    """Compare two register map DBs or xlsx files.

    Accepts .db or .xlsx paths. If xlsx is given, auto-imports to DB first.
    Use --db to save results into a queryable SQLite DB.
    Use --cells to include cell-level comparison (raw_value only).
    Add --comment, --style, --merge-info for deeper comparison (each implies --cells).
    Use --all to enable all comparisons at once.
    Use --smart for sequence-based diff that handles row insert/delete correctly.

    \b
    Examples:
      dsm diff old.db new.db
      dsm diff old.db new.db --cells
      dsm diff old.db new.db --cells --smart
      dsm diff old.db new.db --all --smart
    """
    from dsm.diff import diff_with_auto_import, format_diff, save_diff_to_db

    # --all enables everything
    if compare_all:
        cells = True
        compare_comment = True
        compare_style = True
        compare_merge = True

    # --comment, --style, --merge-info each implies --cells
    if compare_comment or compare_style or compare_merge:
        cells = True

    # --smart implies --cells
    if smart:
        cells = True

    t0 = time.perf_counter()
    result = diff_with_auto_import(path_a, path_b, on_progress=click.echo,
                                    include_cells=cells,
                                    compare_comment=compare_comment,
                                    compare_style=compare_style,
                                    compare_merge=compare_merge,
                                    smart=smart)
    elapsed = time.perf_counter() - t0

    if as_json:
        import json
        from dsm.diff import _REG_FIELDS, _MEMMAP_FIELDS, _reg_changes, _mm_changes

        def _reg_to_json(r, side: str) -> dict:
            return {f: getattr(r, f"{side}_{f}") for f in _REG_FIELDS}

        def _mm_to_json(m, side: str) -> dict:
            return {f: getattr(m, f"{side}_{f}") for f in _MEMMAP_FIELDS}

        data = {
            "added_registers": [
                {"sheet": r.sheet, **_reg_to_json(r, "new")}
                for r in result._filter_regs("added")
            ],
            "removed_registers": [
                {"sheet": r.sheet, **_reg_to_json(r, "old")}
                for r in result._filter_regs("removed")
            ],
            "changed_registers": [
                {
                    "sheet": dr.sheet, "ip": dr.new_name,
                    "indx": dr.new_indx, "page": dr.new_page, "para": dr.new_para,
                    "changes": [
                        {"field": f, "old": old, "new": new}
                        for f, old, new in _reg_changes(dr)
                    ],
                }
                for dr in result._filter_regs("changed")
            ],
            "added_memmap": [_mm_to_json(m, "new") for m in result._filter_mm("added")],
            "removed_memmap": [_mm_to_json(m, "old") for m in result._filter_mm("removed")],
            "changed_memmap": [
                {
                    **_mm_to_json(dm, "old"),
                    "changes": {
                        f: {"old": old, "new": new}
                        for f, old, new in _mm_changes(dm)
                    },
                }
                for dm in result._filter_mm("changed")
            ],
        }
        if cells:
            cell_list = []
            for cd in result.cells:
                entry = {
                    "sheet": cd.sheet, "row": cd.row, "col": cd.col,
                    "status": cd.status,
                    "old_value": cd.old_value, "new_value": cd.new_value,
                }
                if compare_comment:
                    entry["old_comment"] = cd.old_comment
                    entry["new_comment"] = cd.new_comment
                if compare_style:
                    entry["old_style"] = cd.old_style
                    entry["new_style"] = cd.new_style
                if compare_merge:
                    entry["old_merge_range"] = cd.old_merge_range
                    entry["new_merge_range"] = cd.new_merge_range
                cell_list.append(entry)
            data["cell_diffs"] = cell_list
        click.echo(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        click.echo(format_diff(result, verbose=verbose))

    # Save to DB
    if diff_db_path is None:
        diff_db_path = Path(f"diff_{path_a.stem}_{path_b.stem}.db")
    save_diff_to_db(result, diff_db_path, path_a, path_b)
    click.echo(f"\nSaved to {diff_db_path}")

    click.echo(f"({elapsed:.1f}s)")


# -- merge ------------------------------------------------------------------

@main.command()
@click.option("--input-dir", type=click.Path(exists=True, path_type=Path), required=True,
              help="Directory containing split xlsx files")
@click.option("--output", type=click.Path(path_type=Path), default=None,
              help="Output xlsx path")
@click.option("--base", type=click.Path(exists=True, path_type=Path), default=None,
              help="Base DB or xlsx for patch merge (preserves original formatting)")
def merge(input_dir: Path, output: Path | None, base: Path | None):
    """Merge split xlsx files back into a single xlsx.

    Without --base: simple stack merge (combines IP tabs vertically).
    With --base: patch merge — applies only changed cells back onto the
    original xlsx, preserving all formatting, extra columns, and non-level2
    sheets.

    \b
    Examples:
      dsm merge --input-dir design_split/ --output merged.xlsx
      dsm merge --input-dir design_split/ --base original.db --output patched.xlsx
    """
    if base is not None:
        _do_patch_merge(input_dir, output, base)
    else:
        _do_stack_merge(input_dir, output)


def _do_stack_merge(input_dir: Path, output: Path | None):
    from dsm.splitter import merge_split_files

    if output is None:
        output = input_dir.parent / f"{input_dir.name}_merged.xlsx"

    t0 = time.perf_counter()
    click.echo(f"Stack merge from {input_dir}...")
    result = merge_split_files(input_dir, output, on_progress=click.echo)
    elapsed = time.perf_counter() - t0

    total_ips = sum(len(ips) for ips in result.values())
    click.echo(f"Merged {len(result)} sheets ({total_ips} IPs) into {output} ({elapsed:.1f}s)")
    for sheet_name, ip_names in result.items():
        click.echo(f"  {sheet_name}: {', '.join(ip_names)}")


def _do_patch_merge(input_dir: Path, output: Path | None, base: Path):
    from collections import defaultdict
    from dsm.diff import _resolve_db
    from dsm.patcher import patch_merge

    click.echo(f"Resolving base: {base}")
    db_path = _resolve_db(base, on_progress=click.echo)

    if output is None:
        output = input_dir.parent / f"{input_dir.name}_patched.xlsx"

    t0 = time.perf_counter()
    result = patch_merge(db_path, input_dir, output, on_progress=click.echo)
    elapsed = time.perf_counter() - t0

    click.echo(f"Patch merge: {len(result.changes)} cells changed ({elapsed:.1f}s)")
    click.echo(f"Output: {output}")

    if result.skipped_keys:
        click.echo(
            f"Warning: {len(result.skipped_keys)} registers in split "
            f"not found in original"
        )

    # Print change summary grouped by sheet/IP
    if result.changes:
        by_sheet_ip: dict[tuple, list] = defaultdict(list)
        for c in result.changes:
            by_sheet_ip[(c.sheet_name, c.ip_name)].append(c)

        for (sheet, ip), group in sorted(by_sheet_ip.items()):
            click.echo(f"  [{sheet}] {ip}: {len(group)} changes")
            for c in group[:5]:
                click.echo(
                    f"    row={c.row} {c.field}: {c.old_value!r} -> {c.new_value!r}"
                )
            if len(group) > 5:
                click.echo(f"    ... and {len(group) - 5} more")
