"""Compare two DSM databases (register maps and memory maps)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from sqlalchemy import Text, create_engine
from sqlalchemy.orm import (
    DeclarativeBase, Mapped, mapped_column, Session, sessionmaker,
)

from dsm.domain_models import REGMAP_FIELD_MAP, Register, MemoryMapEntry
from dsm.models import ExcelSheet, ExcelWorkbook, init_db


# Fields to compare (exclude internal tracking fields)
_REG_FIELDS = ["type", "indx", "page", "para", "name",
               "d7", "d6", "d5", "d4", "d3", "d2", "d1", "d0", "init"]
_MEMMAP_FIELDS = ["baseaddr", "group", "midgroup", "comment", "special"]


# ── Diff DB models ────────────────────────────────────────────────

class DiffBase(DeclarativeBase):
    pass


class DiffMeta(DiffBase):
    __tablename__ = "diff_meta"

    id: Mapped[int] = mapped_column(primary_key=True)
    old_path: Mapped[str] = mapped_column(Text)
    new_path: Mapped[str] = mapped_column(Text)
    created_at: Mapped[str] = mapped_column(Text)
    added_regs: Mapped[int] = mapped_column(default=0)
    removed_regs: Mapped[int] = mapped_column(default=0)
    changed_regs: Mapped[int] = mapped_column(default=0)
    added_memmap: Mapped[int] = mapped_column(default=0)
    removed_memmap: Mapped[int] = mapped_column(default=0)
    changed_memmap: Mapped[int] = mapped_column(default=0)


class DiffRegister(DiffBase):
    """One row per register: old/new values side-by-side."""
    __tablename__ = "diff_register"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed
    sheet: Mapped[str | None] = mapped_column(Text)
    # -- old values --
    old_type: Mapped[str | None] = mapped_column(Text)
    old_indx: Mapped[str | None] = mapped_column(Text)
    old_page: Mapped[str | None] = mapped_column(Text)
    old_para: Mapped[str | None] = mapped_column(Text)
    old_name: Mapped[str | None] = mapped_column(Text)
    old_d7: Mapped[str | None] = mapped_column(Text)
    old_d6: Mapped[str | None] = mapped_column(Text)
    old_d5: Mapped[str | None] = mapped_column(Text)
    old_d4: Mapped[str | None] = mapped_column(Text)
    old_d3: Mapped[str | None] = mapped_column(Text)
    old_d2: Mapped[str | None] = mapped_column(Text)
    old_d1: Mapped[str | None] = mapped_column(Text)
    old_d0: Mapped[str | None] = mapped_column(Text)
    old_init: Mapped[str | None] = mapped_column(Text)
    # -- new values --
    new_type: Mapped[str | None] = mapped_column(Text)
    new_indx: Mapped[str | None] = mapped_column(Text)
    new_page: Mapped[str | None] = mapped_column(Text)
    new_para: Mapped[str | None] = mapped_column(Text)
    new_name: Mapped[str | None] = mapped_column(Text)
    new_d7: Mapped[str | None] = mapped_column(Text)
    new_d6: Mapped[str | None] = mapped_column(Text)
    new_d5: Mapped[str | None] = mapped_column(Text)
    new_d4: Mapped[str | None] = mapped_column(Text)
    new_d3: Mapped[str | None] = mapped_column(Text)
    new_d2: Mapped[str | None] = mapped_column(Text)
    new_d1: Mapped[str | None] = mapped_column(Text)
    new_d0: Mapped[str | None] = mapped_column(Text)
    new_init: Mapped[str | None] = mapped_column(Text)


class DiffMemmap(DiffBase):
    __tablename__ = "diff_memmap"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed
    old_baseaddr: Mapped[str | None] = mapped_column(Text)
    old_group: Mapped[str | None] = mapped_column(Text)
    old_midgroup: Mapped[str | None] = mapped_column(Text)
    old_comment: Mapped[str | None] = mapped_column(Text)
    old_special: Mapped[str | None] = mapped_column(Text)
    new_baseaddr: Mapped[str | None] = mapped_column(Text)
    new_group: Mapped[str | None] = mapped_column(Text)
    new_midgroup: Mapped[str | None] = mapped_column(Text)
    new_comment: Mapped[str | None] = mapped_column(Text)
    new_special: Mapped[str | None] = mapped_column(Text)


def init_diff_db(db_url: str):
    engine = create_engine(db_url, echo=False)
    DiffBase.metadata.create_all(engine)
    return sessionmaker(bind=engine)


# ── Dataclasses (for in-memory results & text formatting) ─────────

@dataclass
class FieldChange:
    field: str
    old: str | None
    new: str | None


@dataclass
class RegisterDiff:
    """Diff result for a single register row."""
    sheet: str
    ip: str | None
    indx: str | None
    page: str | None
    para: str | None
    changes: list[FieldChange] = field(default_factory=list)


@dataclass
class DiffResult:
    """Full diff between two databases."""
    added_regs: list[dict] = field(default_factory=list)
    removed_regs: list[dict] = field(default_factory=list)
    changed_regs: list[RegisterDiff] = field(default_factory=list)
    added_memmap: list[dict] = field(default_factory=list)
    removed_memmap: list[dict] = field(default_factory=list)
    changed_memmap: list[dict] = field(default_factory=list)


def _reg_key(sheet_name: str, reg: Register) -> tuple:
    """Unique key for a register: (sheet, name, indx, page, para)."""
    return (sheet_name, reg.name, reg.indx, reg.page, reg.para)


def _reg_to_dict(sheet_name: str, reg: Register) -> dict:
    d = {"sheet": sheet_name}
    for f in _REG_FIELDS:
        d[f] = getattr(reg, f)
    return d


def _memmap_key(entry: MemoryMapEntry) -> tuple:
    """Unique key for a memmap entry: (baseaddr, group)."""
    return (entry.baseaddr, entry.group)


def _memmap_to_dict(entry: MemoryMapEntry) -> dict:
    return {f: getattr(entry, f) for f in _MEMMAP_FIELDS}


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


def diff_databases(
    db_path_a: Path,
    db_path_b: Path,
) -> DiffResult:
    """Compare two SQLite databases and return differences.

    Args:
        db_path_a: Path to the "old" / base database.
        db_path_b: Path to the "new" / target database.

    Returns:
        DiffResult with added, removed, and changed items.
    """
    SessionA = init_db(f"sqlite:///{db_path_a}")
    SessionB = init_db(f"sqlite:///{db_path_b}")

    result = DiffResult()

    with SessionA() as sa, SessionB() as sb:
        # --- Register diff ---
        regs_a = _load_registers(sa)
        regs_b = _load_registers(sb)

        keys_a = set(regs_a.keys())
        keys_b = set(regs_b.keys())

        # Added
        for key in sorted(keys_b - keys_a):
            sn, reg = regs_b[key]
            result.added_regs.append(_reg_to_dict(sn, reg))

        # Removed
        for key in sorted(keys_a - keys_b):
            sn, reg = regs_a[key]
            result.removed_regs.append(_reg_to_dict(sn, reg))

        # Changed
        for key in sorted(keys_a & keys_b):
            sn_a, reg_a = regs_a[key]
            sn_b, reg_b = regs_b[key]
            changes = []
            for f in _REG_FIELDS:
                va = getattr(reg_a, f)
                vb = getattr(reg_b, f)
                if va != vb:
                    changes.append(FieldChange(field=f, old=va, new=vb))
            if changes:
                result.changed_regs.append(RegisterDiff(
                    sheet=sn_a,
                    ip=reg_a.name,
                    indx=reg_a.indx,
                    page=reg_a.page,
                    para=reg_a.para,
                    changes=changes,
                ))

        # --- MemoryMap diff ---
        mm_a = _load_memmap(sa)
        mm_b = _load_memmap(sb)

        mkeys_a = set(mm_a.keys())
        mkeys_b = set(mm_b.keys())

        for key in sorted(mkeys_b - mkeys_a):
            result.added_memmap.append(_memmap_to_dict(mm_b[key]))

        for key in sorted(mkeys_a - mkeys_b):
            result.removed_memmap.append(_memmap_to_dict(mm_a[key]))

        for key in sorted(mkeys_a & mkeys_b):
            ea, eb = mm_a[key], mm_b[key]
            changes = {}
            for f in _MEMMAP_FIELDS:
                va = getattr(ea, f)
                vb = getattr(eb, f)
                if va != vb:
                    changes[f] = {"old": va, "new": vb}
            if changes:
                d = _memmap_to_dict(ea)
                d["changes"] = changes
                result.changed_memmap.append(d)

    return result


def save_diff_to_db(
    result: DiffResult,
    diff_db_path: Path,
    old_path: Path,
    new_path: Path,
) -> Path:
    """Save DiffResult into a SQLite DB for querying."""
    from sqlalchemy import insert

    DiffSession = init_diff_db(f"sqlite:///{diff_db_path}")

    with DiffSession() as session:
        # Meta
        meta = DiffMeta(
            old_path=str(old_path),
            new_path=str(new_path),
            created_at=datetime.now().isoformat(),
            added_regs=len(result.added_regs),
            removed_regs=len(result.removed_regs),
            changed_regs=len(result.changed_regs),
            added_memmap=len(result.added_memmap),
            removed_memmap=len(result.removed_memmap),
            changed_memmap=len(result.changed_memmap),
        )
        session.add(meta)
        session.flush()

        # Registers
        reg_rows = []
        for r in result.added_regs:
            row = {"status": "added", "sheet": r["sheet"]}
            for f in _REG_FIELDS:
                row[f"old_{f}"] = None
                row[f"new_{f}"] = r.get(f)
            reg_rows.append(row)

        for r in result.removed_regs:
            row = {"status": "removed", "sheet": r["sheet"]}
            for f in _REG_FIELDS:
                row[f"old_{f}"] = r.get(f)
                row[f"new_{f}"] = None
            reg_rows.append(row)

        for rd in result.changed_regs:
            row = {"status": "changed", "sheet": rd.sheet}
            # Start with the key fields (same in old/new)
            for f in _REG_FIELDS:
                row[f"old_{f}"] = None
                row[f"new_{f}"] = None
            # Key fields are the same for old and new
            for kf in ("name", "indx", "page", "para"):
                val = getattr(rd, {"name": "ip"}.get(kf, kf))
                row[f"old_{kf}"] = val
                row[f"new_{kf}"] = val
            # Apply changes
            for c in rd.changes:
                row[f"old_{c.field}"] = c.old
                row[f"new_{c.field}"] = c.new
            reg_rows.append(row)

        if reg_rows:
            session.execute(insert(DiffRegister), reg_rows)

        # MemoryMap
        mm_rows = []
        for e in result.added_memmap:
            row = {"status": "added"}
            for f in _MEMMAP_FIELDS:
                row[f"old_{f}"] = None
                row[f"new_{f}"] = e.get(f)
            mm_rows.append(row)

        for e in result.removed_memmap:
            row = {"status": "removed"}
            for f in _MEMMAP_FIELDS:
                row[f"old_{f}"] = e.get(f)
                row[f"new_{f}"] = None
            mm_rows.append(row)

        for e in result.changed_memmap:
            row = {"status": "changed"}
            for f in _MEMMAP_FIELDS:
                row[f"old_{f}"] = e.get(f)
                row[f"new_{f}"] = e.get(f)
            for fname, cv in e["changes"].items():
                row[f"old_{fname}"] = cv["old"]
                row[f"new_{fname}"] = cv["new"]
            mm_rows.append(row)

        if mm_rows:
            session.execute(insert(DiffMemmap), mm_rows)

        session.commit()

    return diff_db_path


def diff_with_auto_import(
    path_a: Path,
    path_b: Path,
) -> DiffResult:
    """Diff two paths that can be .db or .xlsx files.

    If an xlsx is given, it is imported into a temp DB first.
    """
    db_a = _resolve_db(path_a)
    db_b = _resolve_db(path_b)
    return diff_databases(db_a, db_b)


def _resolve_db(path: Path) -> Path:
    """If path is .xlsx, import to temp DB and return DB path.
    If .db, return as-is.
    """
    if path.suffix == ".db":
        return path

    if path.suffix in (".xlsx", ".xls"):
        # Check if a companion .db already exists
        companion_db = path.with_suffix(".db")
        if companion_db.exists():
            return companion_db

        # Import into a new DB next to the xlsx
        from dsm.xlsx_parser import import_xlsx
        Session = init_db(f"sqlite:///{companion_db}")
        with Session() as session:
            import_xlsx(session, path)
            session.commit()
        return companion_db

    raise ValueError(f"Unsupported file type: {path.suffix} (expected .db or .xlsx)")


def format_diff(result: DiffResult, verbose: bool = False) -> str:
    """Format a DiffResult as a human-readable string."""
    lines: list[str] = []

    total = (len(result.added_regs) + len(result.removed_regs) +
             len(result.changed_regs) + len(result.added_memmap) +
             len(result.removed_memmap) + len(result.changed_memmap))

    if total == 0:
        return "No differences found."

    # --- Registers ---
    if result.added_regs or result.removed_regs or result.changed_regs:
        lines.append("=== Registers ===")
        lines.append("")

    if result.added_regs:
        lines.append(f"  Added ({len(result.added_regs)}):")
        for r in result.added_regs:
            lines.append(f"    + [{r['sheet']}] {r['name']} "
                         f"indx={r['indx']} page={r['page']} para={r['para']}")
            if verbose:
                bits = " ".join(f"D{i}={r[f'd{i}']}" for i in range(7, -1, -1) if r.get(f"d{i}"))
                if bits:
                    lines.append(f"      {bits}  init={r['init']}")
        lines.append("")

    if result.removed_regs:
        lines.append(f"  Removed ({len(result.removed_regs)}):")
        for r in result.removed_regs:
            lines.append(f"    - [{r['sheet']}] {r['name']} "
                         f"indx={r['indx']} page={r['page']} para={r['para']}")
        lines.append("")

    if result.changed_regs:
        lines.append(f"  Changed ({len(result.changed_regs)}):")
        for rd in result.changed_regs:
            lines.append(f"    ~ [{rd.sheet}] {rd.ip} "
                         f"indx={rd.indx} page={rd.page} para={rd.para}")
            for c in rd.changes:
                lines.append(f"        {c.field}: {c.old!r} -> {c.new!r}")
        lines.append("")

    # --- MemoryMap ---
    if result.added_memmap or result.removed_memmap or result.changed_memmap:
        lines.append("=== MemoryMap ===")
        lines.append("")

    if result.added_memmap:
        lines.append(f"  Added ({len(result.added_memmap)}):")
        for e in result.added_memmap:
            lines.append(f"    + {e['baseaddr']} {e['group']} {e.get('comment', '')}")
        lines.append("")

    if result.removed_memmap:
        lines.append(f"  Removed ({len(result.removed_memmap)}):")
        for e in result.removed_memmap:
            lines.append(f"    - {e['baseaddr']} {e['group']} {e.get('comment', '')}")
        lines.append("")

    if result.changed_memmap:
        lines.append(f"  Changed ({len(result.changed_memmap)}):")
        for e in result.changed_memmap:
            lines.append(f"    ~ {e['baseaddr']} {e['group']}")
            for fname, cv in e["changes"].items():
                lines.append(f"        {fname}: {cv['old']!r} -> {cv['new']!r}")
        lines.append("")

    # Summary
    lines.append(f"Summary: "
                 f"+{len(result.added_regs)} -{len(result.removed_regs)} ~{len(result.changed_regs)} registers, "
                 f"+{len(result.added_memmap)} -{len(result.removed_memmap)} ~{len(result.changed_memmap)} memmap")

    return "\n".join(lines)
