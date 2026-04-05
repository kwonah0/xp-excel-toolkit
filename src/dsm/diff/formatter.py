"""Human-readable diff output formatter."""

from __future__ import annotations

from dsm.diff.models import DiffCell, DiffMemmap, DiffRegister, DiffResult, _REG_FIELDS, _MEMMAP_FIELDS
from dsm.diff.engine import _reg_changes, _mm_changes


def format_diff(result: DiffResult, verbose: bool = False) -> str:
    """Format a DiffResult as a human-readable string."""
    lines: list[str] = []

    added_regs = result._filter_regs("added")
    removed_regs = result._filter_regs("removed")
    changed_regs = result._filter_regs("changed")
    added_mm = result._filter_mm("added")
    removed_mm = result._filter_mm("removed")
    changed_mm = result._filter_mm("changed")

    total = (len(added_regs) + len(removed_regs) + len(changed_regs) +
             len(added_mm) + len(removed_mm) + len(changed_mm))

    if total == 0 and not result.cells:
        return "No differences found."

    # --- Registers ---
    if added_regs or removed_regs or changed_regs:
        lines.append("=== Registers ===")
        lines.append("")

    if added_regs:
        lines.append(f"  Added ({len(added_regs)}):")
        for r in added_regs:
            lines.append(f"    + [{r.sheet}] {r.new_name} "
                         f"indx={r.new_indx} page={r.new_page} para={r.new_para}")
            if verbose:
                bits = " ".join(
                    f"D{i}={getattr(r, f'new_d{i}')}"
                    for i in range(7, -1, -1)
                    if getattr(r, f"new_d{i}")
                )
                if bits:
                    lines.append(f"      {bits}  init={r.new_init}")
        lines.append("")

    if removed_regs:
        lines.append(f"  Removed ({len(removed_regs)}):")
        for r in removed_regs:
            lines.append(f"    - [{r.sheet}] {r.old_name} "
                         f"indx={r.old_indx} page={r.old_page} para={r.old_para}")
        lines.append("")

    if changed_regs:
        lines.append(f"  Changed ({len(changed_regs)}):")
        for dr in changed_regs:
            lines.append(f"    ~ [{dr.sheet}] {dr.new_name} "
                         f"indx={dr.new_indx} page={dr.new_page} para={dr.new_para}")
            for f, old, new in _reg_changes(dr):
                lines.append(f"        {f}: {old!r} -> {new!r}")
        lines.append("")

    # --- MemoryMap ---
    if added_mm or removed_mm or changed_mm:
        lines.append("=== MemoryMap ===")
        lines.append("")

    if added_mm:
        lines.append(f"  Added ({len(added_mm)}):")
        for m in added_mm:
            lines.append(f"    + {m.new_baseaddr} {m.new_group} {m.new_comment or ''}")
        lines.append("")

    if removed_mm:
        lines.append(f"  Removed ({len(removed_mm)}):")
        for m in removed_mm:
            lines.append(f"    - {m.old_baseaddr} {m.old_group} {m.old_comment or ''}")
        lines.append("")

    if changed_mm:
        lines.append(f"  Changed ({len(changed_mm)}):")
        for dm in changed_mm:
            lines.append(f"    ~ {dm.old_baseaddr} {dm.old_group}")
            for f, old, new in _mm_changes(dm):
                lines.append(f"        {f}: {old!r} -> {new!r}")
        lines.append("")

    # --- Cells ---
    if result.cells:
        added_cells = [c for c in result.cells if c.status == "added"]
        removed_cells = [c for c in result.cells if c.status == "removed"]
        changed_cells = [c for c in result.cells if c.status == "changed"]

        # Detect smart mode — smart diff populates old_row/new_row
        is_smart = any(c.old_row is not None or c.new_row is not None for c in result.cells)

        lines.append(f"=== Cells {'(smart)' if is_smart else ''} ===")
        lines.append("")

        def _cell_loc(c: DiffCell) -> str:
            """Format cell location, showing old_row->new_row for smart diff."""
            if is_smart and c.old_row is not None and c.new_row is not None and c.old_row != c.new_row:
                return f"[{c.sheet}] R{c.old_row}\u2192R{c.new_row}C{c.col}"
            return f"[{c.sheet}] R{c.row}C{c.col}"

        if added_cells:
            lines.append(f"  Added ({len(added_cells)}):")
            for c in added_cells[:20]:
                loc = f"[{c.sheet}] R{c.new_row or c.row}C{c.col}" if is_smart else f"[{c.sheet}] R{c.row}C{c.col}"
                lines.append(f"    + {loc}: {c.new_value!r}")
            if len(added_cells) > 20:
                lines.append(f"    ... and {len(added_cells) - 20} more")
            lines.append("")

        if removed_cells:
            lines.append(f"  Removed ({len(removed_cells)}):")
            for c in removed_cells[:20]:
                loc = f"[{c.sheet}] R{c.old_row or c.row}C{c.col}" if is_smart else f"[{c.sheet}] R{c.row}C{c.col}"
                lines.append(f"    - {loc}: {c.old_value!r}")
            if len(removed_cells) > 20:
                lines.append(f"    ... and {len(removed_cells) - 20} more")
            lines.append("")

        if changed_cells:
            lines.append(f"  Changed ({len(changed_cells)}):")
            for c in changed_cells[:20]:
                loc = _cell_loc(c)
                parts = [f"    ~ {loc}:"]
                if c.old_value != c.new_value:
                    parts.append(f" {c.old_value!r} -> {c.new_value!r}")
                lines.append("".join(parts))
                if c.old_comment != c.new_comment and (c.old_comment or c.new_comment):
                    lines.append(f"        comment: {c.old_comment!r} -> {c.new_comment!r}")
                if c.old_style != c.new_style and (c.old_style or c.new_style):
                    lines.append(f"        style changed")
                if c.old_merge_range != c.new_merge_range and (c.old_merge_range or c.new_merge_range):
                    lines.append(f"        merge: {c.old_merge_range} -> {c.new_merge_range}")
            if len(changed_cells) > 20:
                lines.append(f"    ... and {len(changed_cells) - 20} more")
            lines.append("")

    # Summary
    summary_parts = [
        f"+{len(added_regs)} -{len(removed_regs)} ~{len(changed_regs)} registers",
        f"+{len(added_mm)} -{len(removed_mm)} ~{len(changed_mm)} memmap",
    ]
    if result.cells:
        summary_parts.append(f"{len(result.cells)} cell diffs")
    lines.append(f"Summary: {', '.join(summary_parts)}")

    return "\n".join(lines)
