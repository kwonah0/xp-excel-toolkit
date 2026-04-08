"""Human-readable diff output formatter."""

from __future__ import annotations

from dsm.diff.models import DiffCell, DiffMemmap, DiffRegister, DiffResult, _REG_FIELDS, _MEMMAP_FIELDS
from dsm.diff.engine import _reg_changes, _mm_changes


def format_diff(result: DiffResult, verbose: bool = False, limit: int = 0) -> str:
    """Format a DiffResult as a human-readable string.

    Args:
        limit: Max items to show per category. 0 = no limit (show all).
    """
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

        def _group_by_row(cells: list[DiffCell]) -> dict[tuple[str, int], list[DiffCell]]:
            """Group cells by (sheet, row) for row-based display."""
            groups: dict[tuple[str, int], list[DiffCell]] = {}
            for c in cells:
                row_num = (c.new_row or c.row) if is_smart else c.row
                key = (c.sheet, row_num)
                groups.setdefault(key, []).append(c)
            return groups

        def _limit_groups(groups, n):
            if n <= 0:
                return groups
            result_g = {}
            count = 0
            for k, v in groups.items():
                if count >= n:
                    break
                result_g[k] = v
                count += 1
            return result_g

        if added_cells:
            groups = _group_by_row(added_cells)
            lines.append(f"  Added ({len(added_cells)} cells in {len(groups)} rows):")
            shown = _limit_groups(groups, limit)
            for (sheet, row_num), cells_in_row in shown.items():
                lines.append(f"    + [{sheet}] R{row_num}:")
                for c in sorted(cells_in_row, key=lambda x: x.col):
                    lines.append(f"        C{c.col}: {c.new_value!r}")
            if limit > 0 and len(groups) > limit:
                lines.append(f"    ... and {len(groups) - limit} more rows")
            lines.append("")

        if removed_cells:
            groups = _group_by_row(removed_cells)
            lines.append(f"  Removed ({len(removed_cells)} cells in {len(groups)} rows):")
            shown = _limit_groups(groups, limit)
            for (sheet, row_num), cells_in_row in shown.items():
                lines.append(f"    - [{sheet}] R{row_num}:")
                for c in sorted(cells_in_row, key=lambda x: x.col):
                    lines.append(f"        C{c.col}: {c.old_value!r}")
            if limit > 0 and len(groups) > limit:
                lines.append(f"    ... and {len(groups) - limit} more rows")
            lines.append("")

        if changed_cells:
            groups: dict[tuple[str, int, int | None], list[DiffCell]] = {}
            for c in changed_cells:
                row_num = c.row
                old_r = c.old_row if is_smart else None
                key = (c.sheet, row_num, old_r)
                groups.setdefault(key, []).append(c)

            lines.append(f"  Changed ({len(changed_cells)} cells in {len(groups)} rows):")
            shown_keys = list(groups.keys())
            if limit > 0:
                shown_keys = shown_keys[:limit]
            for sheet, row_num, old_r in shown_keys:
                if is_smart and old_r is not None and old_r != row_num:
                    lines.append(f"    ~ [{sheet}] R{old_r}\u2192R{row_num}:")
                else:
                    lines.append(f"    ~ [{sheet}] R{row_num}:")
                for c in sorted(groups[(sheet, row_num, old_r)], key=lambda x: x.col):
                    if c.old_value != c.new_value:
                        lines.append(f"        C{c.col}: {c.old_value!r} -> {c.new_value!r}")
                    if c.old_comment != c.new_comment and (c.old_comment or c.new_comment):
                        lines.append(f"        C{c.col} comment: {c.old_comment!r} -> {c.new_comment!r}")
                    if c.old_style != c.new_style and (c.old_style or c.new_style):
                        lines.append(f"        C{c.col} style changed")
                    if c.old_merge_range != c.new_merge_range and (c.old_merge_range or c.new_merge_range):
                        lines.append(f"        C{c.col} merge: {c.old_merge_range} -> {c.new_merge_range}")
            if limit > 0 and len(groups) > limit:
                lines.append(f"    ... and {len(groups) - limit} more rows")
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
