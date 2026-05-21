# pinmap_demo

End-to-end demo of how a host package (here called ``pinmap``) consumes
``excel_toolkit``.

## Files

| File | Role |
|---|---|
| [`pinmap.py`](./pinmap.py) | The host package, condensed into one module: domain model (`PinEntry`), `SheetConfig`, `ExportHandler`, plus thin `import_pinmap` / `export_pinmap` wrappers. |
| [`make_sample.py`](./make_sample.py) | Builds a tiny synthetic xlsx with merges, styles, and a comment so the demo runs without external files. |
| [`main.py`](./main.py) | End-to-end flow: build sample → `init_db` → import → inspect/modify → round-trip export → verify. |

## Run

From the repo root:

```bash
uv run python examples/pinmap_demo/main.py
```

Expected output (abridged):

```
── Imported PinEntry rows ──
  excel_row=3   A1    VDD  PWR
  excel_row=4   A2    GND  GND
  excel_row=5   A3    SDA  I/O
  excel_row=6   A4    SCL  I/O
── VDD cell metadata ──
  comment: 'Supply pin'
  style:   {'bg_color': ..., 'font_bold': ..., 'border_top': 'thin'}
── change_log entries ──
  UPDATE  pin_entry.name: 'SDA' → 'SDA_NEW'
  UPDATE  pin_entry.direction: 'I/O' → 'OUT'
── Round-tripped xlsx ──
  sheets:        ['Pinmap_A', 'Notes']
  SDA name cell: 'SDA_NEW'
  SDA dir  cell: 'OUT'
  merge count:   2 (title + Dir merge)
  header fill:   00FFFF00
  notes A1:      'Free-form notes'
OK
```

## What it demonstrates

1. **`PinEntry` subclasses `excel_toolkit.Base`** — so `init_db()` creates the
   infra tables AND the domain table in one `create_all()` pass.
2. **`register_audit_target("pin_entry", [...])`** runs at module import,
   *before* `init_db()`, so SQLite UPDATE/DELETE triggers are installed.
3. **`SheetConfig({"Pinmap_*": ...})`** maps the sheet pattern onto the
   domain class via `field_map`.
4. **Vertical merge fill** — `A4/SCL.direction == "I/O"` even though the
   cell is empty in the source (merge origin is row 5).
5. **`ChangeLog` audit** — mutating two columns produces two log rows.
6. **Round-trip via `export_domain_xlsx`** — original BLOB is reused, so the
   header fill, merges, comment, and the non-domain `Notes` sheet all
   survive untouched while only the mutated cells are overwritten.

## Note: vertical merges + per-row domain values

Multiple domain rows can share a single vertical-merged cell at import time
(merge fill — see A3/A4 inheriting `"I/O"`). On export, the flat round-trip
writes each row's value into its source coordinate, and writes that land
inside a merge range get redirected to the **merge origin**. The last
write wins.

If you mutate a per-row field that maps into a vertical merge, the result
will collapse to whichever row was written last. Avoid mutating fields
inside merge ranges through the flat handler — write a custom
``exporter_func`` if you need per-row control there, or break the merge
before exporting.

## What the equivalent real package looks like

For a real host package, split this single file out:

```
pinmap/
├── pyproject.toml          # "excel-toolkit @ git+..." dep
└── src/pinmap/
    ├── __init__.py
    ├── models.py           # PinEntry + register_audit_target
    ├── config.py           # PIN_FIELD_MAP, SHEET_CONFIGS, EXPORT_HANDLERS
    ├── importer.py         # import_pinmap wrapper
    ├── exporter.py         # export_pinmap wrapper
    ├── parsers/            # (optional) parser_func for non-flat sheets
    └── cli.py              # (optional) click — lazy imports
```

See [`docs/cookbook.md`](../../docs/cookbook.md) for the recipe index.
