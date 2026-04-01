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

def _import_cmd(xlsx_path: Path, db_path: Path | None, parallel: bool, workers: int | None):
    if db_path is None:
        db_path = _default_db(xlsx_path)

    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    t0 = time.perf_counter()

    if parallel:
        from dsm.parallel import parallel_import_xlsx
        with Session() as session:
            sheets = parallel_import_xlsx(session, xlsx_path, workers=workers)
            session.commit()
            sheet_info = [(s.name, s.header_row) for s in sheets]
    else:
        from dsm.xlsx_parser import import_xlsx
        with Session() as session:
            sheets = import_xlsx(session, xlsx_path)
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
@click.option("--parallel/--no-parallel", default=False,
              help="Use multiprocessing for parallel sheet import")
@click.option("--workers", type=int, default=None,
              help="Number of worker processes (default: cpu_count)")
def import_cmd(xlsx_path: Path, db_path: Path | None, parallel: bool, workers: int | None):
    """Import all sheets from an xlsx file into SQLite DB."""
    _import_cmd(xlsx_path, db_path, parallel, workers)


main.add_command(import_cmd)


# -- split ------------------------------------------------------------------

@main.command()
@click.argument("xlsx_path", type=click.Path(exists=True, path_type=Path))
@click.option("--db", "db_path", type=click.Path(path_type=Path), default=None,
              help="SQLite DB path (default: <xlsx_stem>.db)")
@click.option("--output-dir", type=click.Path(path_type=Path), default=None,
              help="Output directory (default: <xlsx_stem>_split/)")
@click.option("--parallel/--no-parallel", default=False,
              help="Use multiprocessing for parallel split")
@click.option("--force-reimport", is_flag=True, default=False,
              help="Re-import even if DB already contains this workbook")
def split(xlsx_path: Path, db_path: Path | None, output_dir: Path | None,
          parallel: bool, force_reimport: bool):
    """Split register map by IP — one output file per level2 sheet."""
    if db_path is None:
        db_path = _default_db(xlsx_path)
    if output_dir is None:
        output_dir = xlsx_path.parent / f"{xlsx_path.stem}_split"

    db_url = f"sqlite:///{db_path}"
    Session = init_db(db_url)

    t0 = time.perf_counter()

    with Session() as session:
        from dsm.models import ExcelWorkbook
        existing = session.query(ExcelWorkbook).filter_by(
            filename=xlsx_path.name
        ).first()

        if existing and not force_reimport:
            click.echo(f"Reusing existing DB import for {xlsx_path.name}")
        else:
            click.echo(f"Importing {xlsx_path.name}...")
            from dsm.xlsx_parser import import_xlsx
            import_xlsx(session, xlsx_path)
            session.commit()

        if parallel:
            from dsm.parallel import parallel_split_regmap
            results = parallel_split_regmap(session, xlsx_path, output_dir)
        else:
            from dsm.splitter import split_regmap_from_db
            results = split_regmap_from_db(session, xlsx_path.name, output_dir)

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
def diff(path_a: Path, path_b: Path, diff_db_path: Path | None, verbose: bool, as_json: bool):
    """Compare two register map DBs or xlsx files.

    Accepts .db or .xlsx paths. If xlsx is given, auto-imports to DB first.
    Use --db to save results into a queryable SQLite DB.

    \b
    Examples:
      dsm diff old.db new.db
      dsm diff old.db new.db --db diff_result.db
      dsm diff old.xlsx new.xlsx -v
    """
    from dsm.diff import diff_with_auto_import, format_diff, save_diff_to_db

    t0 = time.perf_counter()
    result = diff_with_auto_import(path_a, path_b)
    elapsed = time.perf_counter() - t0

    if as_json:
        import json
        data = {
            "added_registers": result.added_regs,
            "removed_registers": result.removed_regs,
            "changed_registers": [
                {
                    "sheet": rd.sheet, "ip": rd.ip,
                    "indx": rd.indx, "page": rd.page, "para": rd.para,
                    "changes": [
                        {"field": c.field, "old": c.old, "new": c.new}
                        for c in rd.changes
                    ],
                }
                for rd in result.changed_regs
            ],
            "added_memmap": result.added_memmap,
            "removed_memmap": result.removed_memmap,
            "changed_memmap": result.changed_memmap,
        }
        click.echo(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        click.echo(format_diff(result, verbose=verbose))

    # Save to DB
    if diff_db_path is None:
        diff_db_path = Path(f"diff_{path_a.stem}_{path_b.stem}.db")
    save_diff_to_db(result, diff_db_path, path_a, path_b)
    click.echo(f"\nSaved to {diff_db_path}")

    click.echo(f"({elapsed:.1f}s)")
