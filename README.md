# CSI SAFE ↔ Excel VBA modules (`SAFE_Library.bas`, `SAFE_Use.bas`)

A VBA library module to embed in Excel that **attaches to a running CSI SAFE
instance**, **reads one table at a time into a 2-D array** so the data can be
processed in VBA, and **prints that array to a worksheet** at a tab + top-left
coordinate you pass in as function parameters — plus an optional companion module
of site-specific subs that call into it.

It is written against the COM API documented in
`CSI_API_SAFE_v1_html/` (`SAFEv1.dll`, SAFE 20) and its behaviour is informed
by the reference script `SAFE maximum moment search 2025-07-14.py` (whose
`get_table_view` / `get_table_edit` / `set_table` / `apply_table_edit`
workflow is replicated here).

---

## Files

- `SAFE_Library.bas` — the library module (import this one).
- `SAFE_Use.bas` — companion, site-specific subs (see below).
- `README.md` — this document.

## How to use

1. **Start SAFE** and open the model you want to read from. Leave SAFE running.
   (The script *attaches* — it never starts SAFE, never opens a model, and
   never closes it.)
2. **Add the SAFE reference in the VBA IDE** (once):
   `Alt+F11` → `Tools → References…` → tick `SAFEv1` (or `Browse…` and select
   `SAFEv1.tlb` from the SAFE installation folder). The reference is read from
   this IDE setting — **no file path is hardcoded** in the module.
3. **Import the modules** into the Excel workbook:
   `Alt+F11` → `File → Import File…` → select `SAFE_Library.bas` (and
   `SAFE_Use.bas` if you want the companion subs).
4. **Run** the macro `DemoExport` (`F5`), or call the functions yourself
   (see below).

## Main functions — two halves, called in chain

Reading from SAFE and writing to the sheet are separate now, so the data can be
processed in VBA in between. `ExportSAFETables` returns the table as a 2-D array;
`PrintTable` writes a 2-D array to a sheet. Existing callers are the two calls in
chain.

### PART 1 — `ExportSAFETables`: SAFE → 2-D array

```vba
data = ExportSAFETables(TableKey, Headers, [LoadCases], [LoadCombos], [Warning], [Failed])
```

- `TableKey` — **one** table key, e.g. `"Element Forces - Area Shells"`.
- `Headers()` — OUT: SAFE's column keys for the returned array, in SAFE's order
  (0-based `String` array). The DATA array never carries a header row.
- `LoadCases` — optional; only these load CASES appear in the returned rows of
  result tables. One name, a comma-separated list, or a `String()` array;
  empty = all of them.
- `LoadCombos` — optional; the same, for load COMBINATIONS. SAFE keeps load cases
  and load combinations as two separate display lists, so a case name belongs in
  `LoadCases` and a combination name in `LoadCombos`.
- `Warning` — OUT: the read-quirk text (e.g. `API code 1 (nonzero, but data was
  still returned)`); empty when there is nothing to report.
- `Failed` — OUT: `True` when the read could not be served.
- Returns a 1-based 2-D `Variant` array `[row, col]` of DATA, or `Empty` when
  nothing came back. **Branch on `Failed`, never on `Empty`** — an empty table is
  a normal answer, a failed read is not. Nothing is written to the sheet here.

```vba
' One table into an array — nothing lands on a sheet yet:
Dim hdrs() As String, data As Variant
data = ExportSAFETables("Element Forces - Area Shells", hdrs, LoadCases:="LIVE")
' ...or a list: LoadCases:=Array("LIVE", "DEAD"), or LoadCombos:="1.4DL+1.6LL"
```

### PART 2 — `PrintTable`: 2-D array → worksheet

```vba
rows = PrintTable(Data, SheetName, StartCell, [Title], [Headers], [WriteTitle], _
                  [WriteHeader], [StackHorizontally], [NextRow], [NextCol])
```

- `Data` — a 1-based 2-D array. Normally the array PART 1 returned, but **any**
  1-based 2-D array works (a `Range.Value` array included), so a computed or
  reshaped block is printable.
- `SheetName` / `StartCell` — destination worksheet (created if it does not exist)
  and its top-left coordinate, e.g. `"B3"`.
- `Title` — optional line above the block (normally the table key). It also names
  the block in the log and in the FAILED markers.
- `Headers` — optional `String()` array of column keys (normally `hdrs` from PART 1).
- `WriteTitle` / `WriteHeader` — `False` with `False` writes **neither** the title
  row nor the header row, so the DATA lands flush on `StartCell`.
- `StackHorizontally` — which way `NextRow`/`NextCol` point for the next block:
  `False` = downwards (default), `True` = to the right.
- `NextRow` / `NextCol` — OUT: the cursor for the next block. That is how several
  tables are stacked or placed side by side now that each call prints one table.
- Returns the number of DATA rows written (`0` = an empty block — title and/or
  header rows only, or a `no data returned` marker — and `-1` when the block could
  not be written; see *Error handling*).

```vba
' The chain, in full:
Dim hdrs() As String, data As Variant, warn As String, failed As Boolean

data = ExportSAFETables("Element Forces - Area Shells", hdrs, , , warn, failed)
If failed Then
    MarkTableFailed "Forces", "B2", "Element Forces - Area Shells", warn
Else
    PrintTable data, "Forces", "B2", "Element Forces - Area Shells", hdrs
End If
```

### Several tables on one sheet

Each call writes one table, so a loop runs the chain and threads the cursor
`NextRow`/`NextCol` back into the next `StartCell`:

```vba
Dim keys As Variant, i As Long, cell As String, nr As Long, nc As Long
Dim hdrs() As String, data As Variant, warn As String, failed As Boolean

keys = Array("Point Object Connectivity", "Area Load Assignments - Uniform")
ResetExportFailures                 ' count failures for the whole batch
cell = "A1"
For i = LBound(keys) To UBound(keys)
    data = ExportSAFETables(keys(i), hdrs, , , warn, failed)
    If failed Then
        MarkTableFailed "SAFE Tables", cell, keys(i), warn, , nr, nc
    Else
        PrintTable data, "SAFE Tables", cell, keys(i), hdrs, , , , nr, nc
    End If
    cell = ThisWorkbook.Worksheets("SAFE Tables").Cells(nr, nc).Address(False, False)
Next i
Debug.Print GetLastExportFailures() & " table(s) FAILED"
```

`DemoExport` in the library is this loop, `DemoProcessTable` shows the point of
the split (find the largest value of one column with `TableColumnIndex`, then
print the same array), and `DemoExportSingle` / `DemoExportFiltered` /
`DemoExportFilteredCombo` are the chain in its shortest form.

### Other entry points

- `ListSAFETables(SheetName, StartCell)` — dumps **every table SAFE reports**
  (`GetAllTables`: key | name | import type — a superset of the tables merely
  *available for display*; the log prints both counts so the difference is
  visible), the exact strings to use in `ExportSAFETables`.
- `WriteSAFETable(TableKey, Data, UnlockModel)` — bonus: writes a 2-D array
  **back into SAFE** and applies it (edit workflow, see below).
- `SAFEReadEditingTable(TableKey, Headers, Data)` — public bridge for companion
  modules: reads an **editing** table into `Headers()` (column keys, in SAFE’s
  order) plus a 1-based 2-D array **without** the header row — exactly the form
  `WriteSAFETable` takes back. Read-only; it never touches the model’s lock.
- `MarkTableFailed(SheetName, StartCell, TableKey, Reason, …)` — writes the bold
  red `Table '<key>' : FAILED - <reason>` marker for a table whose **read** failed
  (`ExportSAFETables`' `Failed` = `True`; pass its `Warning` as the reason), so the
  sheet shows that the table was asked for and did not come back instead of an
  empty gap. It returns the same `NextRow`/`NextCol` cursor as `PrintTable`, so the
  two can be swapped inside one loop. (`PrintTable` writes its own marker when the
  data *was* read but the worksheet could not hold the block.)
- `TableColumnIndex(Headers, ColumnKey)` — 1-based index of a column key in a
  `Headers()` array (`0` = the table does not report it), case- and
  space-insensitive. The same number indexes the data array, so a column is found
  **by name** instead of by position:
  `col = TableColumnIndex(hdrs, "M11")` then `data(row, col)`.
- `ResetExportFailures()` / `GetLastExportFailures()` — start and read the failure
  count for a batch of chains (0 = none). Each failed read and each block the
  worksheet could not hold counts once.
- `SAFEConnect()` / `SAFEDisconnect()` — attach / release the running SAFE.
- `ShowLog()` / `ClearLog()` / `GetLog()` / `LogMsg(msg)` — diagnostics; `LogMsg`
  is public so a companion module can write into the same log.
- `gSAFE` / `gSapModel` / `gDB` / `gConnected` — public module state, for API
  calls this library does not wrap (always call `SAFEConnect()` first).
- `DemoExport`, `DemoExportSingle`, `DemoExportFiltered`,
  `DemoExportFilteredCombo`, `DemoProcessTable`, `DemoListTables` — ready-made
  examples (the first five are the chain, in a loop / shortest form / with a
  filter / with a calculation).

## Companion subs (`SAFE_Use.bas`)

A separate module for site-specific subroutines that are called from Excel —
kept out of `SAFE_Library.bas` so re-importing an updated library cannot clobber
them. It needs `SAFE_Library.bas` in the same workbook.

`ApplyPileCoordinates1` is the Excel-facing entry point, and it does nothing but
pass a sheet and a range to the worker `SetPointCoordinates(SheetName,
RangeAddress)` — so adding another area is another two-line wrapper, and the
worker never names a sheet or a range. Every site-specific value is a **local
`Const` inside the wrapper sub**, so there is one place to edit per area and no
module-level name has to be visible project-wide. The worker reads a three-column table
(**Prefix | X | Y**) from the sheet/range it is given and sets the X and Y of
every point in SAFE’s `Point Object Connectivity` whose name **begins with** one
of those prefixes (case-insensitive; the first matching prefix in the sheet wins).

```vba
Public Sub ApplyPileCoordinates1()
    ' local constants: the ONLY hardcoded values, one place to edit per area
    Const PILE_COORD_SHEET As String = "Pile Coords"     ' worksheet
    Const PILE_COORD_RANGE As String = "A2:C50"          ' 3 columns: Prefix | X | Y

    SetPointCoordinates PILE_COORD_SHEET, PILE_COORD_RANGE
End Sub
```

Flow: read the sheet → `SAFEConnect` → `SAFEReadEditingTable` → one pass setting
X/Y per matching prefix → `WriteSAFETable` (applies the edit; unlocks **only** if
SAFE says that table needs it, and restores the lock) → read back and report
points written per prefix, **prefixes that matched nothing**, and any value that
does not match the requested coordinate.

Notes: **message boxes are off by default** (`MsgBoxLogging = False`) — every
message is written to the shared log instead, readable with `ShowLog`; set the
flag to `True` (Immediate window) to see them again. Coordinates are used
**as-is** (no unit conversion) and the units SAFE
reports for `X` are printed in the report; a blank row is ignored, a row with a
non-numeric X/Y is skipped with a log line (so a header row inside the range is
harmless); the edit makes existing analysis results **stale** — this module never
runs the analysis or saves the model.

### Result tables: load cases only (`ReadResultTableForCases`)

```vba
n = ReadResultTableForCases(TableKey, LoadCases, SheetName, TopLeftCell, [IncludeHeader])
```

A generic worker: it runs the two library halves in chain — `ExportSAFETables`
with the requested load cases, then `PrintTable` (or `MarkTableFailed` when the
read failed) — but **clears SAFE’s display load combinations first**
(a single blank list = select none) and puts them back afterwards — on the normal
path and on the error path — so a result table can never pick up combination rows by
accident and the SAFE session is left as it was found.

`LoadCases` has three states: **empty** (`""`, or omitted) reads **every load case
the model reports** — SAFE’s filter has no “all cases” value, so the names are
enumerated with `Analyze.GetRunCaseFlag` and passed on as an ordinary list; a
**non-empty argument with no usable entry** (`, ,` or `Array("","")`) is refused as a
mistake rather than quietly treated as “all”; and an argument with entries reads just
those cases (blank entries ignored).

**No title row, no header row — data only, by default.** `IncludeHeader` defaults to
`False` and suppresses the **whole** block header: neither the table key (title) nor
the column-header row is written, so the DATA lands **flush on `TopLeftCell`**
(`PrintTable` is called with `WriteTitle:=False` and `WriteHeader:=False`). Pass
`IncludeHeader:=True` for the labelled block — table key on the first row, column
keys on the second — which moves the first data row **two** rows down.

A site-specific entry point can also drive the two library halves itself, which is
what `WriteNodalReactions1` does: after clearing the display combinations it reads
`Joint Reactions` (verified against `reference/SAFE Input&Output Table Key
List.csv`, Import Type 0 — a result table, so it can be read but never written
back), **projects it onto a fixed column list**, drops every row whose `Fz` is
zero, and prints what is left. Its settings
are **local constants inside the sub** — `REACT_SHEET` / `REACT_TOPLEFT`
(**top-left corner only**; the block lands there flush — no title row, no header
row — and its columns extend right and down from there; a block too big for one
sheet continues on `<sheet>_2`, `_3`, …), `REACT_LOADCASES` (cases, never
combinations; empty = every load case the model reports) and `REACT_TABLE`. A key
that is wrong for the model writes nothing and marks the block bold red
(`MarkTableFailed`) instead of exporting plausible-looking data.

**Output columns.** The block that reaches the sheet holds **exactly** these
columns, in this order — every other column SAFE reports is left behind:

| Node | OutputCase | Fx | Fy | Fz | Mx | My | Mz |
| ---- | ---------- | -- | -- | -- | -- | -- | -- |

Two labelled blocks inside the sub are the single place to change the output:
**step 6a** holds the column list (`outHdrs` — the eight names written — and
`outSrc`, the SAFE key(s) each one is taken from) and the per-column divisor
(`outDiv`), and **step 6b** holds the row filter (`Fz = 0` rows are dropped; a
blank `Fz` cell is not a zero and is kept). Columns are resolved **by name**
against the keys SAFE returned, never by position, so SAFE may reorder its columns
freely; each `outSrc` entry is a `"a|b"` candidate list tried left to right, so a
key SAFE renames between versions still resolves — `Node` is this block's own name
for the point identifier and resolves to `UniqueName`, falling back to `Label`
(the substitution is logged). A required column the model does not report writes
**nothing** and says which key was looked for, rather than exporting a block with a
missing or misaligned column.

**Unit scaling.** The six force / moment columns (`Fx`, `Fy`, `Fz`, `Mx`, `My`,
`Mz`) are **divided by `FORCE_DIV`** as they are written — the constant sits in the
sub's *EDIT THESE* block, set to `1000` (so N → kN, kN → MN, N·mm → kN·m, …). The
columns are recognised by name (one letter `F`/`M` plus an axis letter), so adding
another force or moment column to the list scales it automatically, and the
division is applied during the projection in step 6b, never to the raw array — the
`Fz = 0` row filter therefore tests the value SAFE reported. A cell that does not
read as a number (a blank, `N/A`, a note) is written exactly as SAFE returned it:
no arithmetic is invented for it. Identifiers (`Node`, `OutputCase`) are written
as they come.

## Table keys

Use the exact strings shown in SAFE *Display → Show Tables*, for example:

- `Point Object Connectivity` (import type 2 — editable)
- `Area Load Assignments - Uniform` (2 — editable)
- `Load Combination Definitions` (2 — editable)
- `Element Forces - Area Shells` (0 — result, read-only)
- `Joint Displacements` (0 — result)
- `Joint Reactions` (0 — result); also `Base Reactions` and `Integrated Wall Reactions`

The full list for SAFE 20 ships in `reference/SAFE Input&Output Table Key List.csv`
(258 keys, each with its **Import Type**: 0 = not importable (a result table),
1 = importable but not interactively, 2 = interactively importable when the model
is unlocked, 3 = importable when locked or unlocked).
`reference/SAFE Input Table Key List.csv` is the same list.

Run `DemoListTables` to see the keys the *open model* actually reports — a model
can omit tables, so the CSV is a superset.

## Error handling & CSI SAFE quirks handled

The script deliberately tolerates SAFE’s quirks instead of aborting:

1. **Nonzero return.** `GetTableForDisplayArray` returns a nonzero code either
   when there is nothing to show *or* when the request cannot be served. The two
   are told apart by whether column headers came back: **nonzero with headers**
   is a warning (reported through the `Warning` out-argument and logged, the data
   is still used); **nonzero without headers** is a hard failure for that read —
   `Failed` comes back `True`, it is logged as an error, it counts by
   `GetLastExportFailures()`, and `MarkTableFailed` leaves a bold red `FAILED`
   marker in the sheet. The rest of a batch is unaffected.
2. **Empty tables.** Headers-only or empty tables never crash the writer. A block
   with nothing at all to write (no title, no header and no data row) gets a
   `no data returned` marker instead of an empty gap; the warning text for a
   nonzero-but-data-returned read is reported through `Warning` and the log
   rather than written beside the marker.
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
   SAFE release (pass `""`); the table's `ImportType` is preflighted with
   `GetAllTables`, and the model is unlocked **only** when that table needs it
   (ImportType 2) — 0 and 1 are refused before the lock is touched, 3 needs no
   unlock — with the lock state found on entry **restored on every exit path**.
   The column count must match or SAFE rejects the write; `TableVersion` is a
   returned item (pass 0); `ApplyEditedTables` is called with
   `FillImportLog:=True` (with `False` the log comes back empty) and may leave
   the model corrupted on a fatal error, so the error/fatal counts are checked
   and `CancelTableEditing` is called. **Save the SAFE model before writing
   back**, and re-run the analysis afterwards — an applied edit makes existing
   results stale.
8. **Display filters** (`LoadCases` / `LoadCombos`): load CASES and load
   COMBINATIONS are two separate SAFE lists
   (`SetLoadCasesSelectedForDisplay` / `SetLoadCombinationsSelectedForDisplay`),
   and each only affects *result* tables. Both selections are saved first and
   restored afterwards, on the normal exit and on the error path. Empty
   parameter = all of them. Quirk: a single blank string selects *no* cases.
   A name SAFE does not accept is **fatal** for that read (`Failed` = `True`,
   `Empty` returned) rather than handing back unfiltered force results.
9. **Excel limits (`PrintTable`).** A block larger than one worksheet is continued
   on `<SheetName>_2`, `_3`, … (title and header row repeated per sheet, when they
   are being written) and a warning is logged; a block too wide for the sheet, or a
   `StartCell` with no room left for the title/header rows (neither row exists when
   `WriteTitle` and `WriteHeader` are both `False`), is **refused** — bold red marker,
   `-1` returned, counted as a failure — rather than truncated silently.

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
  write-back (`WriteSAFETable`) may need the model unlocked, and only for a
  table SAFE reports as `ImportType` 2.
