"""xlwings driver for the Auto Piling workbook.

Runs the workbook's own SAFE macros (module SAFE_Use, which needs SAFE_Library)
through xlwings, so the workbook must ALREADY be open in Excel and SAFE must be
running with the model loaded. Intended order of calls:

    write_coords(coords)  ->  run_analysis()  ->  write_reactions()  ->  get_utilization()

The workbook name is a module constant; the worksheet, cells and macro names are
hardcoded at the top of each function.
"""

from __future__ import annotations

from typing import Any

import xlwings as xw

WORKBOOK = "Auto Piling v20260916-0421.xlsm"    # must already be open in Excel

_wb: xw.Book | None = None      # workbook found on the first call, kept for the session
_n_coords = 0                   # rows written by the last write_coords() call


# ---------------------------------------------------------------------------
# workbook / macro plumbing
# ---------------------------------------------------------------------------
def _book(workbook: str) -> xw.Book:
    """Return the ALREADY OPEN workbook called `workbook` (cached)."""
    global _wb
    if _wb is not None:
        try:
            _wb.name                        # cheap liveness probe: a closed book raises
            return _wb
        except Exception:
            _wb = None

    try:
        apps = list(xw.apps)
    except Exception as exc:
        raise RuntimeError(
            "no running Excel instance found - open the workbook first"
        ) from exc

    for app in apps:
        for book in app.books:
            if book.name.lower() == workbook.lower():
                _wb = book
                return _wb
    raise RuntimeError(
        f"workbook '{workbook}' is not open in Excel - open it and call again"
    )


def _run_macro(book: xw.Book, macro_name: str) -> None:
    """Run one VBA macro; a macro that cannot be run raises RuntimeError."""
    try:
        book.macro(macro_name)()
    except Exception as exc:
        raise RuntimeError(
            f"macro '{macro_name}' failed: {exc}. The SAFE_Use subs log their own "
            "errors instead of raising, so check the log in the workbook (ShowLog) "
            "for the reason."
        ) from exc


# ---------------------------------------------------------------------------
# value helpers
# ---------------------------------------------------------------------------
def _number(value: object, where: str) -> float:
    """float(value); a blank or non-numeric value raises ValueError."""
    if value is None or isinstance(value, bool):
        raise ValueError(f"{where} is not a number: {value!r}")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        raise ValueError(f"{where} is blank, expected a number")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"{where} is not a number: {value!r}") from exc


def _grid(values: Any, n_rows: int, n_cols: int) -> list[list[Any]]:
    """Normalize a rectangular xlwings .value read into a list of row lists."""
    if values is None:                      # whole range empty
        return [[None] * n_cols for _ in range(n_rows)]
    if n_rows == 1:
        return [list(values)] if n_cols > 1 else [[values]]
    return [list(row) for row in values]


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
def write_coords(coords: list[list[float]]) -> None:
    """Write (x, y) pairs to 'Pile Coords'!B3:C, then apply them in SAFE."""
    SHEET = "Pile Coords"
    FIRST_CELL = "B3"                                # x in B, y in C, downwards
    MACRO = "SAFE_Use.ApplyPileCoordinates1"         # reads A2:C50 (Prefix | X | Y)

    global _n_coords
    if not coords:
        raise ValueError("coords is empty - nothing to write")

    rows = []
    for i, pair in enumerate(coords, start=1):
        if not isinstance(pair, (list, tuple)) or len(pair) != 2:
            raise ValueError(f"coords[{i}] is not an (x, y) pair: {pair!r}")
        rows.append([_number(pair[0], f"coords[{i}][0]"),
                     _number(pair[1], f"coords[{i}][1]")])

    book = _book(WORKBOOK)
    book.sheets[SHEET].range(FIRST_CELL).resize(len(rows), 2).value = rows
    _run_macro(book, MACRO)
    _n_coords = len(rows)


def run_analysis() -> None:
    """Run SAFE's analysis for the open model."""
    MACRO = "SAFE_Use.RunAnalysis1"

    _run_macro(_book(WORKBOOK), MACRO)


def write_reactions() -> None:
    """Write SAFE's nodal reactions into the workbook."""
    MACRO = "SAFE_Use.WriteNodalReactions1"

    _run_macro(_book(WORKBOOK), MACRO)


def get_utilization() -> list[list[float]]:
    """Read the utilization pair of every pile just written, top to bottom."""
    SHEET = "Pile Coords"
    FIRST_CELL = "H3"                                # first utilization pair
    N_COLS = 2

    if _n_coords <= 0:
        raise RuntimeError(
            "write_coords() has not run yet, so the number of utilization rows is "
            "unknown - call write_coords() first"
        )

    rng = _book(WORKBOOK).sheets[SHEET].range(FIRST_CELL).resize(_n_coords, N_COLS)
    grid = _grid(rng.value, _n_coords, N_COLS)

    out = []
    for r in range(_n_coords):
        cell_row = rng.row + r
        out.append([_number(grid[r][c], xw.utils.col_name(rng.column + c) + str(cell_row))
                    for c in range(N_COLS)])
    return out