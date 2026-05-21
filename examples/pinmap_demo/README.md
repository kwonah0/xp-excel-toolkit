# pinmap_demo

End-to-end demo of how a host package (here called ``pinmap``) consumes
``excel_toolkit``. The demo demonstrates **the facade pattern**: downstream
applications import only from ``pinmap.api`` and never have to touch
``excel_toolkit`` directly.

## Layout

```
examples/pinmap_demo/
├── README.md
├── make_sample.py          # builds the synthetic input xlsx
├── main.py                 # downstream-app simulation — imports only pinmap.api
└── pinmap/                 # the host package
    ├── __init__.py         # intentionally empty — facade lives in api.py
    ├── api.py              # ← public surface (Base, init_db, PinEntry, helpers)
    ├── models.py           # PinEntry domain (excel_toolkit.Base subclass) + audit registration
    ├── importer.py         # import_pinmap wrapper (SheetConfig knows "Pinmap_*")
    └── exporter.py         # export_pinmap wrapper (ExportHandler knows "Pinmap_*")
```

The split deliberately keeps the package's `__init__.py` empty. Downstream
code imports the explicit facade module:

```python
from pinmap.api import init_db, PinEntry, import_pinmap, export_pinmap
```

That single import line is what makes the facade real — there is no
backdoor by which `import pinmap` magically pulls in excel_toolkit
side-effects.

## Run

From the repo root:

```bash
uv run python examples/pinmap_demo/main.py
```

Expected output (abridged):

```
── Imported PinEntry rows ──
  excel_row=3   A1    VDD  PWR
  ...
── VDD cell metadata ──
  comment: 'Supply pin'
  style:   {'border_top': 'thin'}
── change_log entries ──
  UPDATE  pin_entry.direction: 'PWR' → 'CORE'
  UPDATE  pin_entry.name: 'SDA' → 'SDA_NEW'
── Round-tripped xlsx ──
  sheets:        ['Pinmap_A', 'Notes']
  VDD dir  cell: 'CORE'
  SDA name cell: 'SDA_NEW'
  merge count:   2 (title + Dir merge)
  header fill:   00FFFF00
  notes A1:      'Free-form notes'
OK
```

## What it demonstrates

1. **Facade isolation.** `main.py` does not import `excel_toolkit` even
   once. The only excel_toolkit-aware modules are `pinmap.api`,
   `pinmap.models`, `pinmap.importer`, `pinmap.exporter`. If pinmap ever
   moved off excel_toolkit, those four files are the only places that
   change.

2. **`pinmap.api.Base IS excel_toolkit.Base`** — same identity, same
   `MetaData`. So a single `init_db()` creates the infra tables
   (`excel_workbook`, `excel_sheet`, `excel_cell`, `excel_merge`,
   `change_log`, `sheet_config`) AND the domain table (`pin_entry`) in
   one `create_all()` pass.

   We re-export it by name rather than subclass because SQLAlchemy 2.0's
   `DeclarativeBase` doesn't accept an un-mapped intermediate subclass —
   aliasing is the canonical way to give a registry-owning Base a new
   name in a downstream package.

3. **`register_audit_target("pin_entry", [...])`** runs at module-import
   time, before `init_db()`, so the SQLite UPDATE/DELETE triggers for
   `pin_entry` are installed.

4. **`SheetConfig({"Pinmap_*": ...})`** maps the sheet pattern onto the
   domain class via `field_map`.

5. **Vertical merge fill** — `A4/SCL.direction == "I/O"` even though the
   cell is empty in the source (merge origin is row 5).

6. **`ChangeLog` audit** — mutating two columns produces two log rows.

7. **Round-trip via `export_domain_xlsx`** — original BLOB is reused, so
   the header fill, merges, comment, and the non-domain `Notes` sheet
   all survive untouched while only the mutated cells are overwritten.

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

## What the equivalent "real" host package looks like

Same layout, just with a `pyproject.toml`:

```
pinmap/
├── pyproject.toml          # "excel-toolkit @ git+..." dep
└── src/pinmap/
    ├── __init__.py         # empty
    ├── api.py              # facade
    ├── models.py
    ├── importer.py
    ├── exporter.py
    ├── parsers/            # (optional) parser_func for non-flat sheets
    └── cli.py              # (optional) click — lazy imports
```

See [`docs/cookbook.md`](../../docs/cookbook.md) for the full recipe index.
