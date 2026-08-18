# CSI SAFE → Excel VBA table extractor (`SAFE_Export.bas`)

A single VBA module to embed in Excel that **attaches to a running CSI SAFE
instance** and writes one or more SAFE database tables into a worksheet at a
tab + top-left coordinate you pass in as function parameters.

It is written against the COM API documented in
`CSI_API_SAFE_v1_html/` (`SAFEv1.dll`, SAFE 20) and its behaviour is informed
by the reference script `SAFE maximum moment search 2025-07-14.py` (whose
`get_table_view` / `get_table_edit` / `set_table` / `apply_table_edit`
workflow is replicated here).

---

## Files

| File              | Purpose                                                        |
|-------------------|----------------------------------------------------------------|
| `SAFE_Export.bas` | The VBA module. Import into Excel (`Alt+F11 → File → Import File…`). |
| `README.md`       | This document.                                                 |

## How to use

1. **Start SAFE** and open the model you want to read from. Leave SAFE running.
   (The script *attaches* — it never starts SAFE, never opens a model, and
   never closes it.)
2. **Add the SAFE reference in the VBA IDE** (once):
   `Alt+F11` → `Tools → References…` → tick `SAFEv1` (or `Browse…` and select
   `SAFEv1.tlb` from the SAFE installation folder). The reference is read from
   this IDE setting — **no file path is hardcoded** in the module.
3. **Import the module** into the Excel workbook:
   `Alt+F11` → `File → Import File…` → select `SAFE_Export.bas`.
4. **Run** the macro `DemoExport` (`F5`), or call the functions yourself
   (see below).

## Main function

```vba
n = ExportSAFETables(Tables, SheetName, StartCell, StackHorizontally, IncludeHeader)
```

| Parameter           | Meaning                                                              |
|---------------------|----------------------------------------------------------------------|
| `Tables`            | One table key (`String`), a comma-separated `String`, or an array of keys: `Array("…","…")` |
| `SheetName`         | Destination worksheet (“tab”) — created if it does not exist         |
| `StartCell`         | Top-left coordinate, e.g. `"B3"`                                     |
| `StackHorizontally` | `False` = stack tables downwards (default); `True` = side by side    |
| `IncludeHeader`     | Write the column-header row (default `True`)                         || `LoadCases`         | Optional: only these load cases appear in result tables. One name, a comma-separated string, or `Array("…","…")`. Empty/missing = all load cases || **Returns**         | Number of tables written; `-1` on a fatal error                      |

```vba
' Single table, own sheet:
ExportSAFETables "Element Forces - Area Shells", "Forces", "B2"

' Several tables stacked on one sheet:
ExportSAFETables Array("Point Object Connectivity", _
                      "Area Load Assignments - Uniform"), _
                 "SAFE Tables", "A1"

' Only include specific load cases in result tables (e.g. just LIVE):
ExportSAFETables "Element Forces - Area Shells", "Forces", "B2", LoadCases:="LIVE"
' ...or a list: LoadCases:=Array("LIVE", "DEAD")
```

### Other entry points

- `ListSAFETables(SheetName, StartCell)` — dumps **every available table key**
  (exact strings to use in `ExportSAFETables`).
- `WriteSAFETable(TableKey, Data, UnlockModel)` — bonus: writes a 2-D array
  **back into SAFE** and applies it (edit workflow, see below).
- `SAFEConnect()` / `SAFEDisconnect()` — attach / release the running SAFE.
- `ShowLog()` / `ClearLog()` / `GetLog()` — diagnostics.
- `DemoExport`, `DemoExportSingle`, `DemoListTables` — ready-made examples.

## Table keys

Use the exact strings shown in SAFE *Display → Show Tables*, for example:

- `Point Object Connectivity`
- `Area Load Assignments - Uniform`
- `Element Forces - Area Shells` (needs analysis results)
- `Joint Displacements`
- `Load Combination Definitions`

Run `DemoListTables` to see the exact keys available in your model.

## Error handling & CSI SAFE quirks handled

The script deliberately tolerates SAFE’s quirks instead of aborting:

1. **Nonzero return = “nothing to show”.** `GetTableForDisplayArray` returns a
   nonzero code when a table is empty / analysis hasn’t been run. This is
   treated as a **warning** (logged, a marker is written to the sheet) and the
   loop continues to the next table.
2. **Empty tables.** Headers-only or empty tables never crash the writer.
3. **Flattened data.** SAFE returns data as a *one-dimensional* array, row by
   row. Rows are rebuilt from the number of columns
   (`FieldsKeysIncluded`), with bounds guards in case the array is shorter
   than expected.
4. **All columns / all objects.** A single blank `FieldKeyList` entry requests
   all columns; `GroupName = ""` (or `"All"`) returns all objects.
5. **Reference from the VBA IDE.** The SAFE type-library reference is added in
   the Excel VBA IDE (`Tools → References…`); no file path is hardcoded. Only
   the COM ProgIDs are kept in the **Configuration Zone**.
6. **Attach ≠ close.** When attached, the script never calls
   `ApplicationExit` (that would close the user’s SAFE session).
7. **Edit workflow guards** (`WriteSAFETable`): `GroupName` is inactive in this
   SAFE release (pass `""`); the model must be unlocked
   (`SetModelIsLocked(False)`); the column count must match or SAFE rejects the
   write; `ApplyEditedTables` may leave the model corrupted on a fatal error,
   so the error/fatal counts are checked and `CancelTableEditing` is called to
   clear the pending-edit buffer. **Save the SAFE model before writing back.**
8. **Load-case filter** (`LoadCases` parameter): implemented with
   `SetLoadCasesSelectedForDisplay`, which only affects *result* tables.
   The previous selection is saved and restored afterwards (also on error).
   Empty parameter = all load cases (SAFE's default). Quirk: a single blank
   string selects *no* cases.

All warnings go to the VBA **Immediate window** (`Ctrl+G`) and can be shown
with `ShowLog()`.

## Notes / limitations

- Tested assumptions come from the SAFE 20 API docs (`SAFEv1`, v1.23.0.0) and
  the reference Python script; later SAFE versions may rename tables.
- The connect step tries `GetObject(, "CSI.SAFE.API.ETABSObject")` first,
  then the `CSI.SAFE.API.Helper` (via `New Helper`) as a fallback. If neither
  finds a running instance, make sure SAFE is open with a model and that the
  `SAFEv1` reference is ticked in `Tools → References…`.
- Reading tables works whether or not the model is locked; only the
  write-back (`WriteSAFETable`) needs it unlocked.
