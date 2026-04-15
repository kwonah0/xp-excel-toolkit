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

def _ensure_xlsx(path: Path) -> Path:
    """If path is .xls, convert to .xlsx using LibreOffice (cached in __dsm_cache__/)."""
    from dsm.convert import ensure_xlsx_cached
    return ensure_xlsx_cached(path, on_progress=lambda msg: click.echo(msg))


def _import_cmd(xlsx_path: Path, db_path: Path | None, with_formulas: bool = False):
    xlsx_path = _ensure_xlsx(xlsx_path)

    if db_path is None:
        db_path = _default_db(xlsx_path)

    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    t0 = time.perf_counter()

    from dsm.xlsx_parser import import_xlsx
    with Session() as session:
        sheets = import_xlsx(session, xlsx_path, on_progress=click.echo,
                             with_formulas=with_formulas)
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
@click.option("--with-formulas", is_flag=True, default=False,
              help="Also store formula strings (loads workbook twice, results in cached_value)")
def import_cmd(xlsx_path: Path, db_path: Path | None, with_formulas: bool):
    """Import all sheets from an xlsx/xls file into SQLite DB.

    .xls files are auto-converted to .xlsx via LibreOffice before import.
    By default, stores calculated values (data_only mode).
    Use --with-formulas to also store formula strings.
    """
    _import_cmd(xlsx_path, db_path, with_formulas=with_formulas)


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

    if xlsx_path is not None:
        xlsx_path = _ensure_xlsx(xlsx_path)

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
              help="Save diff results to SQLite DB (path, implies --save-db)")
@click.option("--verbose", "-v", is_flag=True, default=False,
              help="Show detailed bit-field info for added registers")
@click.option("--json", "as_json", is_flag=True, default=False,
              help="Output as JSON (shortcut for --format json)")
@click.option("--format", "fmt", type=click.Choice(["text", "json", "daff", "csv"]),
              default=None, help="Output format (default: text)")
@click.option("--domain", "include_domain", is_flag=True, default=False,
              help="Include domain-level diff (Register and MemoryMap models)")
@click.option("--no-cells", is_flag=True, default=False,
              help="Disable cell-level diff (use with --domain for domain-only)")
@click.option("--comment", "compare_comment", is_flag=True, default=False,
              help="Compare cell comments")
@click.option("--style", "compare_style", is_flag=True, default=False,
              help="Compare cell styles")
@click.option("--merge-info", "compare_merge", is_flag=True, default=False,
              help="Compare merged cell ranges")
@click.option("--all", "compare_all", is_flag=True, default=False,
              help="Enable all comparisons (cells + domain + comment + style + merge)")
@click.option("--positional", is_flag=True, default=False,
              help="Use positional diff instead of smart diff (row/col based, may cascade on insert/delete)")
@click.option("--with-formulas", is_flag=True, default=False,
              help="Import xlsx with formulas and show formula strings in diff output")
@click.option("--limit", type=int, default=0,
              help="Max items per category in output (0 = show all, default: 0)")
@click.option("--output", "-o", type=click.Path(path_type=Path), default=None,
              help="Override auto-generated CSV filename (format auto-detected from extension)")
@click.option("--no-file", is_flag=True, default=False,
              help="Do not generate output file (stdout only)")
@click.option("--save-db", is_flag=True, default=False,
              help="Save diff results to a SQLite DB file")
def diff(path_a: Path, path_b: Path, diff_db_path: Path | None, verbose: bool,
         as_json: bool, fmt: str | None, include_domain: bool, no_cells: bool,
         compare_comment: bool, compare_style: bool,
         compare_merge: bool, compare_all: bool, positional: bool,
         with_formulas: bool, limit: int, output: Path | None,
         no_file: bool, save_db: bool):
    """Compare two register map DBs or xlsx files.

    Accepts .db or .xlsx paths. If xlsx is given, auto-imports to DB first.
    By default, compares all cells using smart (sequence-based) diff.

    \b
    Output: auto-generates a CSV file (diff_{stem}_{datetime}.csv) and
    prints a summary to stdout. Use -o to override filename, --no-file
    to skip file generation.

    \b
    Output formats (--format):
      text  — row-grouped human-readable (full output to stdout, no auto CSV)
      json  — structured JSON
      daff  — tabular diff with +++ / --- / -> markers
      csv   — CSV with status,sheet,row,col,old_value,new_value,...

    \b
    Examples:
      dsm diff old.db new.db
      dsm diff old.db new.db -o result.csv
      dsm diff old.db new.db --no-file
      dsm diff old.db new.db --format text
      dsm diff old.db new.db --all
      dsm diff old.db new.db --save-db
    """
    from dsm.diff import (
        diff_with_auto_import, format_csv, format_daff, format_diff,
        format_summary, save_diff_to_db,
    )

    # --all enables everything
    if compare_all:
        include_domain = True
        compare_comment = True
        compare_style = True
        compare_merge = True

    include_cells = not no_cells

    # smart is the default; --positional disables it
    smart = not positional

    # Resolve output format
    _EXT_FMT = {".json": "json", ".csv": "csv", ".daff": "daff", ".txt": "text"}
    explicit_format = fmt is not None or as_json
    if fmt is None:
        if as_json:
            fmt = "json"
        elif output and output.suffix in _EXT_FMT:
            fmt = _EXT_FMT[output.suffix]
        else:
            fmt = "csv"  # default file format is CSV

    t0 = time.perf_counter()
    result = diff_with_auto_import(path_a, path_b, on_progress=click.echo,
                                    include_cells=include_cells,
                                    include_domain=include_domain,
                                    compare_comment=compare_comment,
                                    compare_style=compare_style,
                                    compare_merge=compare_merge,
                                    smart=smart,
                                    with_formulas=with_formulas)
    elapsed = time.perf_counter() - t0

    def _format_output(chosen_fmt: str) -> str:
        if chosen_fmt == "json":
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
            if include_cells:
                from dsm.diff.formatter import _col_letter
                cell_list = []
                for cd in result.cells:
                    entry = {
                        "sheet": cd.sheet, "row": cd.row, "col": _col_letter(cd.col),
                        "status": cd.status,
                        "old_value": cd.old_value, "new_value": cd.new_value,
                    }
                    if cd.old_row is not None:
                        entry["old_row"] = cd.old_row
                    if cd.new_row is not None:
                        entry["new_row"] = cd.new_row
                    if compare_comment:
                        entry["old_comment"] = cd.old_comment
                        entry["new_comment"] = cd.new_comment
                    if compare_style:
                        entry["old_style"] = cd.old_style
                        entry["new_style"] = cd.new_style
                    if compare_merge:
                        entry["old_merge_range"] = cd.old_merge_range
                        entry["new_merge_range"] = cd.new_merge_range
                    if cd.old_formula or cd.new_formula:
                        entry["old_formula"] = cd.old_formula
                        entry["new_formula"] = cd.new_formula
                    cell_list.append(entry)
                data["cell_diffs"] = cell_list
            return json.dumps(data, indent=2, ensure_ascii=False)
        elif chosen_fmt == "daff":
            return format_daff(result)
        elif chosen_fmt == "csv":
            return format_csv(result)
        else:
            return format_diff(result, verbose=verbose, limit=limit)

    # When --format text is explicit: full output to stdout, no auto file
    if explicit_format and fmt == "text":
        output_text = _format_output("text")
        if output:
            output.write_text(output_text, encoding="utf-8")
            click.echo(f"Saved to {output}")
        else:
            click.echo(output_text)
    else:
        # Default: auto-generate file + summary to stdout
        output_text = _format_output(fmt)

        if not no_file:
            if output:
                out_path = output
            else:
                # Auto-generate: diff_{stem}_{datetime}.csv (or matching ext)
                from datetime import datetime
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                ext = {"csv": ".csv", "json": ".json", "daff": ".daff"}.get(fmt, ".csv")
                out_path = Path(f"diff_{path_a.stem}_{ts}{ext}")
            out_path.write_text(output_text, encoding="utf-8")
            click.echo(f"Saved to {out_path}")

        # Always print summary to stdout
        click.echo(format_summary(result))

    # Save to DB (only with --save-db or --db)
    if save_db or diff_db_path is not None:
        if diff_db_path is None:
            diff_db_path = Path(f"diff_{path_a.stem}_{path_b.stem}.db")
        save_diff_to_db(result, diff_db_path, path_a, path_b)
        click.echo(f"Diff DB saved to {diff_db_path}")

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
    db_path, _is_temp = _resolve_db(base, on_progress=click.echo)

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


# -- config -----------------------------------------------------------------

@main.group()
def config():
    """Manage sheet import configurations stored in the DB."""


@config.command("list")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
def config_list(db_path: Path):
    """List current sheet configurations."""
    import json
    from dsm.models import SheetConfigEntry
    from dsm.domain_models import seed_default_configs

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        seed_default_configs(session)
        session.commit()

        entries = session.query(SheetConfigEntry).all()
        if not entries:
            click.echo("No sheet configurations found.")
            return

        click.echo(f"{'ID':>3} {'Pattern':<20} {'Domain':<20} {'Header':<7} Field Map")
        click.echo("-" * 80)
        for e in entries:
            fm = ""
            if e.field_map_json:
                keys = list(json.loads(e.field_map_json).keys())
                fm = ", ".join(keys[:5])
                if len(keys) > 5:
                    fm += f" ... (+{len(keys) - 5})"
            click.echo(f"{e.id:>3} {e.pattern:<20} {e.domain_type or '-':<20} "
                       f"{e.header_row or 'auto':<7} {fm}")


@config.command("add")
@click.option("--db", "db_path", type=click.Path(path_type=Path), required=True)
@click.option("--pattern", required=True, help="Sheet name pattern (fnmatch, e.g. 'level2_*')")
@click.option("--domain", "domain_type", default=None,
              help="Domain type: 'register' or 'memorymap_entry'")
@click.option("--field-map", "field_map_json", default=None,
              help="Field map as JSON string (e.g. '{\"TYPE\":\"type\",...}')")
@click.option("--header-row", type=int, default=None,
              help="Explicit header row (default: auto-detect)")
def config_add(db_path: Path, pattern: str, domain_type: str | None,
               field_map_json: str | None, header_row: int | None):
    """Add a sheet configuration entry."""
    import json
    from dsm.models import SheetConfigEntry
    from dsm.domain_models import DOMAIN_REGISTRY, FIELD_MAP_REGISTRY

    if domain_type and domain_type not in DOMAIN_REGISTRY:
        raise click.UsageError(
            f"Unknown domain type '{domain_type}'. "
            f"Available: {', '.join(DOMAIN_REGISTRY.keys())}"
        )

    # If no field_map_json provided but domain_type is known, use default
    if field_map_json is None and domain_type:
        default_fm = FIELD_MAP_REGISTRY.get(domain_type)
        if default_fm:
            field_map_json = json.dumps(default_fm)
    elif field_map_json:
        # Validate JSON
        json.loads(field_map_json)

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        entry = SheetConfigEntry(
            pattern=pattern,
            domain_type=domain_type,
            field_map_json=field_map_json,
            header_row=header_row,
        )
        session.add(entry)
        session.commit()
        click.echo(f"Added config #{entry.id}: pattern='{pattern}' domain={domain_type}")


@config.command("remove")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.argument("config_id", type=int)
def config_remove(db_path: Path, config_id: int):
    """Remove a sheet configuration by ID."""
    from dsm.models import SheetConfigEntry

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        entry = session.get(SheetConfigEntry, config_id)
        if not entry:
            raise click.UsageError(f"Config #{config_id} not found.")
        click.echo(f"Removing config #{entry.id}: pattern='{entry.pattern}' domain={entry.domain_type}")
        session.delete(entry)
        session.commit()


@config.command("reset")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
def config_reset(db_path: Path):
    """Reset configurations to defaults."""
    from dsm.models import SheetConfigEntry
    from dsm.domain_models import seed_default_configs

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        session.query(SheetConfigEntry).delete()
        session.flush()
        seed_default_configs(session)
        session.commit()
        count = session.query(SheetConfigEntry).count()
        click.echo(f"Reset to {count} default configurations.")


# -- sql --------------------------------------------------------------------

@main.command("sql")
@click.argument("query_str")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
def sql_cmd(query_str: str, db_path: Path, as_json: bool):
    """Execute raw SQL against the DB.

    \b
    Examples:
      dsm sql "SELECT * FROM register" --db regmap.db
      dsm sql "SELECT * FROM register" --db regmap.db --json
      dsm sql "UPDATE register SET type='RW1' WHERE id=1" --db regmap.db
    """
    import json
    from sqlalchemy import text

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        result = session.execute(text(query_str))

        if result.returns_rows:
            columns = list(result.keys())
            rows = result.fetchall()

            if as_json:
                data = [dict(zip(columns, row)) for row in rows]
                click.echo(json.dumps(data, indent=2, default=str))
            else:
                col_widths = [len(c) for c in columns]
                for row in rows:
                    for i, val in enumerate(row):
                        col_widths[i] = max(col_widths[i], len(str(val) if val is not None else "NULL"))

                header = "  ".join(c.ljust(col_widths[i]) for i, c in enumerate(columns))
                click.echo(header)
                click.echo("  ".join("-" * w for w in col_widths))
                for row in rows:
                    line = "  ".join(
                        (str(v) if v is not None else "NULL").ljust(col_widths[i])
                        for i, v in enumerate(row)
                    )
                    click.echo(line)

                click.echo(f"\n({len(rows)} rows)")
        else:
            session.commit()
            click.echo(f"OK ({result.rowcount} rows affected)")


# -- log --------------------------------------------------------------------

@main.group()
def log():
    """View and manage change history (audit trail)."""


@log.command("show")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.option("--table", "table_name", default=None, help="Filter by table name")
@click.option("--last", "limit", type=int, default=20, help="Number of entries (default: 20)")
def log_show(db_path: Path, table_name: str | None, limit: int):
    """Show recent change history.

    \b
    Examples:
      dsm log show --db regmap.db
      dsm log show --db regmap.db --table register
      dsm log show --db regmap.db --last 50
    """
    from dsm.models import ChangeLog

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        q = session.query(ChangeLog).order_by(ChangeLog.id.desc())
        if table_name:
            q = q.filter(ChangeLog.table_name == table_name)
        entries = q.limit(limit).all()

        if not entries:
            click.echo("No changes recorded.")
            return

        click.echo(f"{'id':>4}  {'timestamp':<20} {'op':<6} {'table':<18} "
                   f"{'row':>4} {'column':<16} {'old':<20} {'new':<20}")
        click.echo("-" * 112)
        for e in reversed(entries):
            old = (e.old_value or "-")[:20].ljust(20)
            new = (e.new_value or "-")[:20].ljust(20)
            click.echo(
                f"{e.id:>4}  {e.timestamp:<20} {e.operation:<6} {e.table_name:<18} "
                f"{e.row_id:>4} {e.column_name:<16} {old} {new}"
            )
        click.echo(f"\n({len(entries)} entries)")


@log.command("undo")
@click.argument("log_id", type=int)
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
def log_undo(log_id: int, db_path: Path):
    """Revert a change by restoring the old value.

    \b
    Examples:
      dsm log undo 42 --db regmap.db
    """
    from sqlalchemy import text as sa_text
    from dsm.models import ChangeLog

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        entry = session.get(ChangeLog, log_id)
        if not entry:
            raise click.UsageError(f"Change log #{log_id} not found.")

        if entry.operation == "DELETE":
            click.echo("Cannot undo DELETE (row was removed). Use import to restore.")
            return

        click.echo(f"Reverting: {entry.table_name}.{entry.column_name} "
                    f"(row {entry.row_id}): {entry.new_value!r} -> {entry.old_value!r}")

        session.execute(
            sa_text(f"UPDATE {entry.table_name} SET {entry.column_name} = :val WHERE id = :rid"),
            {"val": entry.old_value, "rid": entry.row_id},
        )
        session.commit()
        click.echo("Done.")


@log.command("clear")
@click.option("--db", "db_path", type=click.Path(exists=True, path_type=Path), required=True)
@click.confirmation_option(prompt="Clear all change history?")
def log_clear(db_path: Path):
    """Clear all change history."""
    from dsm.models import ChangeLog

    Session = init_db(f"sqlite:///{db_path}")
    with Session() as session:
        count = session.query(ChangeLog).count()
        session.query(ChangeLog).delete()
        session.commit()
        click.echo(f"Cleared {count} entries.")
