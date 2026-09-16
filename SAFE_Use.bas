Attribute VB_Name = "SAFE_Use"
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  SAFE_Use.bas  --  site-specific, Excel-facing subroutines
' ============================================================================
'  WHAT IT DOES
'    Each entry-point sub near the top is deliberately trivial: it carries its own
'    hardcoded settings as LOCAL constants and hands them to a worker.
'    ApplyPileCoordinates1 -> SetPointCoordinates(SheetName, RangeAddress), which
'    reads a small three-column table from the sheet/range it is given
'    (Prefix | X | Y) and then, in SAFE's "Point Object Connectivity" table,
'    changes the X and Y of EVERY point whose name BEGINS with one of those
'    prefixes - e.g. "Pile_1_3" and "Pile_1_7" both follow the prefix "Pile_1".
'    Matching is case-insensitive and the FIRST matching prefix in the sheet
'    wins, so an overlapping pair is resolved by putting the more specific
'    prefix first.
'
'  REQUIRES
'    - The companion module SAFE_Library.bas in the same workbook. It supplies
'      SAFEConnect, the table read (ExportSAFETables -> 2-D array) with its
'      sheet-side other half (PrintTable / MarkTableFailed), the editing-table
'      read (SAFEReadEditingTable), the write-back (WriteSAFETable: ImportType
'      preflight, unlock only if needed, edit + ApplyEditedTables, lock restore)
'      and the shared log (LogMsg).
'    - SAFE open with the model loaded, and the SAFEv1 reference ticked in the
'      VBA IDE (Tools > References...).
'
'  WORKFLOW - exactly what the sub does
'    1. read <SheetName>!<RangeAddress>   (Prefix | X | Y)
'    2. attach to SAFE (no-op when already attached)
'    3. read the whole "Point Object Connectivity" editing table, with its
'       column keys, through SAFE_Library
'    4. ONE pass over its rows: each point whose name starts with a prefix gets
'       that prefix's X and Y written into its X and Y cells
'    5. hand the rebuilt table to WriteSAFETable, which applies the edit to SAFE
'    6. read the table back and report: points written per prefix, values that do
'       NOT match the requested coordinates, and prefixes that matched nothing
'
'  NOTES
'    - MESSAGE BOXES ARE OFF by default (they are modal, and annoying in
'      production): every message is written to the shared log instead. Set the
'      module flag MsgBoxLogging = True to see them again, and read the history
'      afterwards with ShowLog.
'    - Coordinates are written AS GIVEN - no unit conversion. SAFE carries the
'      model's PRESENT units in its tables, and the units it reports for the X
'      column are printed in the report so the sheet can be checked against them.
'    - Applying the edit makes existing analysis results STALE. This sub does NOT
'      re-run the analysis and does NOT save the model - do that in SAFE before
'      using any results.
'    - Running it twice is harmless: the same coordinates are written again and
'      the report says how many values already matched.
' ============================================================================

' ---------------------------------------------------------------------------
' WHERE THE SETTINGS LIVE
' ---------------------------------------------------------------------------
' Every site-specific value - sheet, range, table key, load cases - is a LOCAL
' Const inside the entry-point sub that uses it (see ENTRY POINTS below). Nothing
' site-specific is declared at module level, so there is exactly ONE place to edit
' per area, a wrapper can live in any module (or behind a button) without hitting
' VBA's module-scoped Private Const rule, and no name is added to the project for
' another module to collide with.
'
' What does stay at module level is the WORKERS' internal detail - POINT_TABLE,
' the COL_* keys, the tolerances and MSG_TITLE - because only the code in this
' module refers to it.
'
' The coordinate table's format and its skip rules are documented inside
' ApplyPileCoordinates1, next to the two constants that name them.

' The SAFE table that holds the points (exact key - see ListSAFETables).
Private Const POINT_TABLE As String = "Point Object Connectivity"

' The columns of that table, looked up BY NAME - never by position, so SAFE is
' free to reorder its columns without breaking this sub.
Private Const COL_NAME As String = "UniqueName"
Private Const COL_X As String = "X"
Private Const COL_Y As String = "Y"

' Read-back tolerance, in model units: a value is accepted when it is within
' COORD_ABS_TOL, or within COORD_REL_TOL of the requested value, whichever is
' larger. The coordinate round-trips through SAFE's own text form, so anything
' beyond this is a real disagreement rather than rounding.
Private Const COORD_ABS_TOL As Double = 0.000001
Private Const COORD_REL_TOL As Double = 0.000000001

' ---------------------------------------------------------------------------
' MESSAGE BOXES - OFF BY DEFAULT
' ---------------------------------------------------------------------------
' A MsgBox is modal: handy while a module is being built, annoying in production
' (the run stops and every problem needs a click). This switch controls EVERY
' message this module raises, because they all go through the single Say() helper
' below. Turn them back on from the Immediate window (Ctrl+G):
'
'     MsgBoxLogging = True
'
' Nothing is lost while it is False: Say() always writes the same text to the
' shared log, so the messages are there afterwards in ShowLog (SAFE_Library).
Public MsgBoxLogging As Boolean       ' False = silent (VBA's default for a Boolean)

' Window title of every message this module raises.
Private Const MSG_TITLE As String = "Apply pile coordinates"

' ---------------------------------------------------------------------------
' SAY - this module's ONE MsgBox call site
' ---------------------------------------------------------------------------
' Logs the message (always - that is the record), then shows it only when
' MsgBoxLogging is True. AlsoLog:=False is for a caller that has already logged
' the same text itself (BuildReport uses LogReport for the indented version), so
' the log does not receive the block twice.
Private Sub Say( _
    ByVal Text As String, _
    Optional ByVal Icon As Long = vbInformation, _
    Optional ByVal Title As String = "", _
    Optional ByVal AlsoLog As Boolean = True)

    If AlsoLog Then LogMsg Text
    If Len(Title) = 0 Then Title = MSG_TITLE
    If MsgBoxLogging Then MsgBox Text, Icon, Title
End Sub

' ===========================================================================
' ENTRY POINTS - run these from Excel (Alt+F8)
' ===========================================================================
' A wrapper is deliberately trivial: its LOCAL constants are the only
' site-specific names anywhere, and it hands them to a worker. Copy one for any
' other area - any number of wrappers can drive the same worker.
Public Sub ApplyPileCoordinates1()
    ' ---- EDIT THESE --------------------------------------------------------
    ' Sheet and range that hold the coordinate table. The range must cover THREE
    ' columns, in this order, and is read from its FIRST row downwards:
    '
    '     Prefix     |    X     |    Y
    '     Pile_1     |   10.3   |   27.1
    '     Pile_2     |   11.4   |   27.8
    '
    ' A completely blank row is ignored (so an oversized range is harmless); a row
    ' with coordinates but no prefix, or with a non-numeric X or Y (a header line,
    ' a note), is SKIPPED with a log line; a repeated prefix keeps its FIRST value.
    Const PILE_COORD_SHEET As String = "Pile Coords"     ' <-- EDIT ME
    Const PILE_COORD_RANGE As String = "A2:C50"          ' <-- EDIT ME (3 columns)
    ' ------------------------------------------------------------------------

    SetPointCoordinates PILE_COORD_SHEET, PILE_COORD_RANGE
End Sub

' ===========================================================================
' WORKER - SetPointCoordinates(SheetName, RangeAddress)
' ===========================================================================
Public Sub SetPointCoordinates(ByVal SheetName As String, ByVal RangeAddress As String)

    Dim ws As Worksheet
    Dim vals As Variant
    Dim nRowsIn As Long
    Dim nColsIn As Long
    Dim pfx() As String
    Dim xs() As Double
    Dim ys() As Double
    Dim nPfx As Long
    Dim skipped As Long
    Dim hdrs() As String
    Dim data As Variant
    Dim ixName As Long, ixX As Long, ixY As Long
    Dim matched() As Long
    Dim nRows As Long
    Dim r As Long, p As Long
    Dim nm As String
    Dim touched As Long

    ' ---- 1. connect (no-op when already attached) --------------------------
    If Not SAFEConnect() Then
        Say "Could not attach to a running SAFE instance." & vbCrLf & vbCrLf & _
            "Start SAFE, open the model you want to edit, then run this again." & _
            vbCrLf & vbCrLf & "Detail: " & LastErrorText(), vbExclamation
        Exit Sub
    End If

    ' ---- 2. read the coordinate table from Excel ---------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo Fail

    If ws Is Nothing Then
        Say "Worksheet '" & SheetName & "' was not found in this workbook." & _
            vbCrLf & vbCrLf & "Check the PILE_COORD_SHEET constant in the wrapper sub " & _
            "that called this (or the argument it passes).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    vals = ws.Range(RangeAddress).Value
    On Error GoTo Fail
    If Err.Number <> 0 Then
        Say "Range '" & RangeAddress & "' could not be read from sheet '" & _
            SheetName & "'." & vbCrLf & vbCrLf & "Check the PILE_COORD_RANGE " & _
            "constant in the wrapper sub that called this (or the argument it " & _
            "passes).", vbExclamation
        Exit Sub
    End If

    If Not IsArray(vals) Then
        Say "Range '" & RangeAddress & "' on '" & SheetName & "' must cover at " & _
            "least one row and three columns (Prefix | X | Y).", vbExclamation
        Exit Sub
    End If

    nRowsIn = 0
    nColsIn = 0
    On Error Resume Next
    nRowsIn = UBound(vals, 1)
    nColsIn = UBound(vals, 2)
    On Error GoTo Fail
    If nRowsIn < 1 Or nColsIn < 3 Then
        Say "Range '" & RangeAddress & "' is " & nRowsIn & " row(s) x " & nColsIn & _
            " column(s); three columns are needed: Prefix | X | Y.", vbExclamation
        Exit Sub
    End If

    nPfx = ReadCoordTable(vals, pfx, xs, ys, skipped)
    If nPfx = 0 Then
        Say "No usable rows found in '" & SheetName & "'!" & RangeAddress & "." & _
            vbCrLf & vbCrLf & "Expected three columns: Prefix | X | Y." & vbCrLf & _
            "(" & skipped & " row(s) were skipped - ShowLog says why.)", vbExclamation
        Exit Sub
    End If

    LogMsg "SetPointCoordinates: read " & nPfx & " coordinate row(s) from '" & _
           SheetName & "'!" & RangeAddress & " (" & skipped & " row(s) skipped)."

    ' ---- 3. read SAFE's point table (editing form, columns by name) --------
    If Not SAFEReadEditingTable(POINT_TABLE, hdrs, data) Then
        Say "Could not read '" & POINT_TABLE & "' from SAFE." & vbCrLf & vbCrLf & _
            "Detail: " & LastErrorText() & vbCrLf & vbCrLf & _
            "ShowLog has the full log.", vbExclamation
        Exit Sub
    End If

    ixName = ColIndex(hdrs, COL_NAME)
    ixX = ColIndex(hdrs, COL_X)
    ixY = ColIndex(hdrs, COL_Y)
    If ixName = 0 Or ixX = 0 Or ixY = 0 Then
        Say "'" & POINT_TABLE & "' does not report the expected column(s)." & _
            vbCrLf & vbCrLf & "Looking for: " & COL_NAME & ", " & COL_X & ", " & _
            COL_Y & vbCrLf & "SAFE reports: " & Join(hdrs, ", "), vbExclamation
        Exit Sub
    End If

    nRows = 0
    If Not IsEmpty(data) Then
        On Error Resume Next
        nRows = UBound(data, 1)
        On Error GoTo Fail
    End If
    If nRows < 1 Then
        Say "'" & POINT_TABLE & "' has no rows, so there is nothing to change.", _
            vbInformation
        Exit Sub
    End If

    ' ---- 4. ONE pass: X / Y for every point whose name starts with a prefix -
    ReDim matched(1 To nPfx)
    For r = 1 To nRows
        nm = SafeText(data(r, ixName))
        If Len(nm) > 0 Then
            For p = 1 To nPfx
                If StartsWith(nm, pfx(p)) Then
                    ' First matching prefix wins: prefixes are applied in the
                    ' order they appear in the sheet.
                    data(r, ixX) = NumToText(xs(p))
                    data(r, ixY) = NumToText(ys(p))
                    matched(p) = matched(p) + 1
                    touched = touched + 1
                    Exit For
                End If
            Next p
        End If
    Next r

    If touched = 0 Then
        Say "Nothing to do: no point in '" & POINT_TABLE & "' starts with any of " & _
            "the " & nPfx & " prefix(es) read from '" & SheetName & "'!" & _
            RangeAddress & "." & vbCrLf & vbCrLf & _
            "Check the spelling of the prefixes against the point names " & _
            "(export the table, or run ListSAFETables, to see the names). " & _
            "Nothing was written to SAFE.", vbExclamation
        Exit Sub
    End If

    LogMsg "SetPointCoordinates: " & touched & " point(s) matched - applying the edit."

    ' ---- 5. apply the edit to SAFE ----------------------------------------
    If Not WriteSAFETable(POINT_TABLE, data) Then
        Say "SAFE did NOT accept the edit - '" & POINT_TABLE & "' was left " & _
            "unchanged." & vbCrLf & vbCrLf & _
            "Detail: " & LastErrorText() & vbCrLf & vbCrLf & _
            "If SAFE reports that the model may be corrupt, close it WITHOUT " & _
            "saving and reopen it (the saved copy is the safe fallback). " & _
            "ShowLog has the full log.", vbExclamation
        Exit Sub
    End If

    ' ---- 6. read back, verify, report -------------------------------------
    ' BuildReport logs the report itself (LogReport, indented), so Say must not
    ' log the same block a second time.
    Say BuildReport(pfx, xs, ys, nPfx, skipped, matched, touched, _
                    SheetName, RangeAddress), AlsoLog:=False
    Exit Sub

Fail:
    ' An unexpected error (a COM drop, a range that vanished, a Variant that is not
    ' an array ...). Report it rather than exiting silently - the user has to know
    ' whether the model was touched.
    Say "Apply pile coordinates stopped on an unexpected error:" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
        "Check the state of the model in SAFE before running this again. " & _
        "ShowLog has the detail.", vbExclamation
End Sub

' ===========================================================================
' REPORT - read the table back and verify what SAFE now holds
' ===========================================================================
Private Function BuildReport( _
    ByRef pfx() As String, _
    ByRef xs() As Double, _
    ByRef ys() As Double, _
    ByVal nPfx As Long, _
    ByVal Skipped As Long, _
    ByRef matched() As Long, _
    ByVal Touched As Long, _
    ByVal SheetName As String, _
    ByVal RangeAddress As String) As String

    Dim rep As String
    Dim hdrs() As String
    Dim back As Variant
    Dim ixName As Long, ixX As Long, ixY As Long
    Dim nRows As Long
    Dim r As Long, p As Long
    Dim nm As String
    Dim okCount() As Long
    Dim badCount() As Long
    Dim verified As Long, mismatched As Long, noMatch As Long
    Dim head As String

    rep = "Pile coordinates applied to '" & POINT_TABLE & "'." & vbCrLf & vbCrLf
    rep = rep & "X/Y units SAFE reports: " & TableFieldUnits(POINT_TABLE, COL_X, _
          "unknown") & "  (sheet values are used as-is, no conversion)" & vbCrLf
    rep = rep & "Source: '" & SheetName & "'!" & RangeAddress & vbCrLf
    rep = rep & "Prefix rows read: " & nPfx & "   (skipped rows: " & Skipped & ")" & vbCrLf
    rep = rep & "Points written:   " & Touched & vbCrLf & vbCrLf

    If Not SAFEReadEditingTable(POINT_TABLE, hdrs, back) Then
        rep = rep & "Read-back check could not be run: " & LastErrorText() & vbCrLf & _
              vbCrLf & StaleWarning()
        BuildReport = rep
        LogReport rep
        Exit Function
    End If

    ixName = ColIndex(hdrs, COL_NAME)
    ixX = ColIndex(hdrs, COL_X)
    ixY = ColIndex(hdrs, COL_Y)
    If ixName = 0 Or ixX = 0 Or ixY = 0 Then
        rep = rep & "Read-back check skipped: SAFE did not report the expected " & _
              "column key(s)." & vbCrLf & vbCrLf & StaleWarning()
        BuildReport = rep
        LogReport rep
        Exit Function
    End If

    nRows = 0
    If Not IsEmpty(back) Then
        On Error Resume Next
        nRows = UBound(back, 1)
        On Error GoTo 0
    End If

    ' Verify with the SAME rule the edit used - for each point, the FIRST prefix
    ' it matches.
    ReDim okCount(1 To nPfx)
    ReDim badCount(1 To nPfx)
    For r = 1 To nRows
        nm = SafeText(back(r, ixName))
        If Len(nm) > 0 Then
            For p = 1 To nPfx
                If StartsWith(nm, pfx(p)) Then
                    If NearlyEqual(NumFromText(SafeText(back(r, ixX))), xs(p)) And _
                       NearlyEqual(NumFromText(SafeText(back(r, ixY))), ys(p)) Then
                        okCount(p) = okCount(p) + 1
                    Else
                        badCount(p) = badCount(p) + 1
                        LogMsg "SetPointCoordinates: MISMATCH - point '" & nm & _
                               "' reads (" & SafeText(back(r, ixX)) & ", " & _
                               SafeText(back(r, ixY)) & "), expected (" & _
                               NumToText(xs(p)) & ", " & NumToText(ys(p)) & ") " & _
                               "for prefix '" & pfx(p) & "'."
                    End If
                    Exit For
                End If
            Next p
        End If
    Next r

    rep = rep & "Per prefix - points whose name starts with it:" & vbCrLf
    For p = 1 To nPfx
        head = "  " & pfx(p) & "  ->  " & Format$(xs(p), "0.###") & ", " & _
               Format$(ys(p), "0.###") & "   "
        verified = verified + okCount(p)
        mismatched = mismatched + badCount(p)
        If okCount(p) + badCount(p) = 0 Then
            noMatch = noMatch + 1
            rep = rep & head & "NO POINT MATCHED (written: " & matched(p) & _
                  ") - check the spelling" & vbCrLf
        Else
            rep = rep & head & (okCount(p) + badCount(p)) & " point(s), " & _
                  okCount(p) & " verified"
            If badCount(p) > 0 Then
                rep = rep & ", " & badCount(p) & " NOT at the requested value"
            End If
            rep = rep & vbCrLf
        End If
    Next p

    rep = rep & vbCrLf & "Read-back: " & verified & " point(s) verified"
    If mismatched > 0 Then rep = rep & ", " & mismatched & " MISMATCH(ES)"
    If noMatch > 0 Then rep = rep & "; " & noMatch & " prefix(es) matched nothing"
    rep = rep & "." & vbCrLf & vbCrLf & StaleWarning()

    BuildReport = rep
    LogReport rep
End Function

Private Function StaleWarning() As String
    StaleWarning = "NOTE: any existing analysis results are now STALE - re-run the " & _
                   "analysis in SAFE before using them. This sub does not run it " & _
                   "(and does not save the model)."
End Function

' Push the report into the shared log line by line, so ShowLog shows the same text
' the report carried (message box shown or not), with the line breaks intact.
Private Sub LogReport(ByVal rep As String)
    Dim lines() As String
    Dim i As Long

    lines = Split(rep, vbCrLf)
    LogMsg "SetPointCoordinates report:"
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then LogMsg "  " & lines(i)
    Next i
End Sub

' ===========================================================================
' RESULT TABLES - read ONE result table for LOAD CASES ONLY
' ===========================================================================
' ReadResultTableForCases is the generic worker: table key, load cases, sheet and
' top-left cell are all ARGUMENTS. It exists because SAFE keeps load CASES and
' load COMBINATIONS as two independent "selected for display" lists, and every use
' in this workbook wants cases only. It drives the library's TWO halves in chain:
'
'   * PART 1 - ExportSAFETables reads the table into a 2-D array with the requested
'     load cases applied (and puts the caller's case selection back before it
'     returns);
'   * PART 2 - PrintTable writes that array at the requested top-left cell, or
'     MarkTableFailed leaves a bold red marker there when the READ failed;
'   * the COMBINATION display selection is cleared here FIRST (a list holding a
'     single blank string is SAFE's documented way of selecting none) and put back
'     at the end, on the normal path AND on the error path - so the table cannot
'     pick up combination rows and the user's SAFE session is left as it was found.
'
' With no LoadCombos argument reaching ExportSAFETables, nothing can re-select
' combinations behind our back: combinations are EMPTY for the duration of the read.
'
' LoadCases has THREE states, and they are deliberately not interchangeable:
'   * EMPTY (omitted, "", "   ", or an array with no elements): read EVERY load
'     case the model reports. SAFE's display filter can only select NAMES - it has
'     no "all cases" value - so the names are enumerated from the model
'     (Analyze.GetRunCaseFlag) and passed on as an ordinary list.
'   * passed but with NO usable entry (", ," or Array("","")): nothing was asked
'     for while an argument WAS given - a mistake. Reported and refused (0), not
'     quietly reinterpreted as "all cases".
'   * passed with entries: only those cases are read (blank entries are ignored).
'
' Returns 1 when the table was dealt with, 0 when there was nothing to do at all
' or no usable case was named, -1 on a fatal error. Messages go through Say, so
' they are logged whether or not MsgBoxLogging is on.
Public Function ReadResultTableForCases( _
    ByVal TableKey As String, _
    ByVal LoadCases As Variant, _
    ByVal SheetName As String, _
    ByVal TopLeftCell As String, _
    Optional ByVal IncludeHeader As Boolean = False) As Long
    ' IncludeHeader is False by DEFAULT, and it suppresses the WHOLE block
    ' header: what lands on the sheet is the DATA alone, flush on TopLeftCell -
    ' no title row and no column-header row, because PrintTable is called with
    ' WriteTitle:=False and WriteHeader:=False. That is the layout this workbook
    ' wants for result tables. Pass True for the labelled block instead - the
    ' table key on the first row and SAFE's column keys on the second - which
    ' moves the first data row two rows further down.

    Dim savedCombos() As String
    Dim nSavedCombos As Long
    Dim none(0 To 0) As String
    Dim clearedCombos As Boolean
    Dim ret As Long
    Dim written As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim caseCount As Long
    Dim useAllCases As Boolean
    Dim resolvedCases As Variant
    Dim caseDesc As String
    Dim allNames() As String
    Dim allRuns() As Boolean
    Dim nAll As Long
    Dim parts As Variant
    Dim i As Long
    Dim reason As String
    Dim hdrs() As String
    Dim data As Variant
    Dim warn As String
    Dim failed As Boolean
    Dim blockResult As Long

    none(0) = ""                       ' SAFE: a single blank = select NOTHING

    ' ---- 0. WHICH load cases -----------------------------------------------
    ' EMPTY argument = every load case the model reports (section 1b, which needs
    ' the connection first). NOT empty but with no usable entry = a mistake: refuse
    ' rather than guessing, and do NOT quietly upgrade it to "all cases". A blank
    ' ENTRY inside an otherwise real list is still simply ignored.
    useAllCases = IsEmptyNameArgument(LoadCases)

    If Not useAllCases Then
        If IsArray(LoadCases) Then
            For i = LBound(LoadCases) To UBound(LoadCases)
                If Len(Trim$(CStr(LoadCases(i)))) > 0 Then caseCount = caseCount + 1
            Next i
        Else
            parts = Split(CStr(LoadCases), ",")
            For i = LBound(parts) To UBound(parts)
                If Len(Trim$(parts(i))) > 0 Then caseCount = caseCount + 1
            Next i
        End If

        If caseCount = 0 Then
            Say "ReadResultTableForCases: a load-case argument was passed, but no " & _
                "usable name was found in it, so nothing was read." & vbCrLf & vbCrLf & _
                "Names are separated by commas (""DEAD, LIVE"") or passed as an " & _
                "array. Blank entries are ignored, but an argument holding NOTHING " & _
                "but blanks is treated as a mistake rather than as ""all cases"" - " & _
                "leave the argument out entirely (or pass """") to read every load " & _
                "case the model reports.", vbExclamation
            ReadResultTableForCases = 0
            Exit Function
        End If

        resolvedCases = LoadCases
        caseDesc = "the " & caseCount & " load case(s) asked for"
    End If

    written = 0
    clearedCombos = False
    On Error GoTo Fail

    ' ---- 1. attach (no-op when already attached) ---------------------------
    If Not SAFEConnect() Then
        Say "ReadResultTableForCases: could not attach to a running SAFE " & _
            "instance, so nothing was read.", vbExclamation
        ReadResultTableForCases = -1
        Exit Function
    End If

    ' ---- 1b. "all load cases": ask the MODEL for its case names ------------
    ' SAFE's display filter can only select NAMES - there is no "all" value - so
    ' "all cases" has to be enumerated. Analyze.GetRunCaseFlag returns every
    ' analysis case the model has, which is exactly the list the setter wants.
    If useAllCases Then
        On Error Resume Next
        Err.Clear
        ret = gSapModel.Analyze.GetRunCaseFlag(nAll, allNames, allRuns)
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo Fail

        If errNum <> 0 Or ret <> 0 Then
            Say "ReadResultTableForCases: the model did not report its load case " & _
                "list (Analyze.GetRunCaseFlag returned " & ret & ", error " & _
                CStr(errNum) & " - " & errDesc & "), so nothing was read." & vbCrLf & _
                vbCrLf & "Name the cases explicitly instead, or check the model in " & _
                "SAFE.", vbExclamation
            ReadResultTableForCases = -1
            Exit Function
        End If

        caseCount = SelectedNameCount(allNames)
        If caseCount = 0 Then
            Say "ReadResultTableForCases: LoadCases was empty, but the model " & _
                "reports NO load case to read, so nothing was read." & vbCrLf & vbCrLf & _
                "Define the load cases (and run the analysis) in SAFE, then try " & _
                "again.", vbExclamation
            ReadResultTableForCases = 0
            Exit Function
        End If

        resolvedCases = allNames
        caseDesc = "every load case the model reports (" & caseCount & ")"
        LogMsg "ReadResultTableForCases: LoadCases was empty - reading ALL " & _
               caseCount & " load case(s) the model reports."
    End If

    ' ---- 2. remember the display combination selection ---------------------
    nSavedCombos = 0
    Erase savedCombos
    On Error Resume Next
    Err.Clear
    ret = gDB.GetLoadCombinationsSelectedForDisplay(nSavedCombos, savedCombos)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    If errNum <> 0 Or ret <> 0 Then
        ' Not fatal: combinations are cleared anyway. Only the caller's previous
        ' selection would be lost, so say that and carry on.
        LogMsg "ReadResultTableForCases: could not read the current display " & _
               "load-combination selection (" & CStr(errNum) & " / code " & ret & _
               " - " & errDesc & ") - it will be left as ""none selected""."
        nSavedCombos = 0
        Erase savedCombos
    End If

    ' ---- 3. clear combinations (cases are handled by ExportSAFETables) -----
    On Error Resume Next
    Err.Clear
    ret = gDB.SetLoadCombinationsSelectedForDisplay(none)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    If errNum <> 0 Or ret <> 0 Then
        ' The read still goes ahead, but the caller has to know the table may
        ' contain combination rows - that is a data question, not cosmetics.
        If errNum <> 0 Then
            reason = CStr(errNum) & " - " & errDesc
        Else
            reason = "return code " & ret
        End If
        Say "ReadResultTableForCases: SAFE would not clear the display load " & _
            "combinations (" & reason & ")." & vbCrLf & vbCrLf & _
            "The read will go ahead, but if the current selection in SAFE's Show " & _
            "Tables includes load combinations, the table will contain their rows " & _
            "as well as the load cases asked for.", vbExclamation
    Else
        clearedCombos = True
    End If

    ' ---- 4. read the table, then write it: the two library halves in chain --
    ' resolvedCases is the caller's list, or the model's full case list when the
    ' argument was empty (section 1b). Never LoadCombos: combinations must stay
    ' cleared for the read, which is the whole point of this worker.
    ' PART 1 - SAFE -> 2-D array. hdrs comes back with SAFE's column keys.
    ResetExportFailures            ' the count reported below belongs to THIS read
    data = ExportSAFETables(TableKey, hdrs, LoadCases:=resolvedCases, _
                            Warning:=warn, Failed:=failed)

    If failed Then
        ' PART 2, failure path - the red marker, so the sheet shows that this
        ' table was asked for and did not come back.
        MarkTableFailed SheetName, TopLeftCell, TableKey, warn
        written = -1
        Say "ReadResultTableForCases: reading '" & TableKey & "' failed - " & _
            "the log names the reason (" & warn & "), and LastErrorText() gives " & _
            "the last one." & vbCrLf & vbCrLf & _
            "If SAFE reports the table key is unknown, run ListSAFETables (or " & _
            "DemoListTables) to list the keys THIS model reports.", vbExclamation
    Else
        ' PART 2 - 2-D array -> sheet. IncludeHeader drives BOTH the title row
        ' and the column-header row of the block.
        blockResult = PrintTable(data, SheetName, TopLeftCell, TableKey, hdrs, _
                                 WriteTitle:=IncludeHeader, WriteHeader:=IncludeHeader)
        If blockResult < 0 Then
            written = -1
            Say "ReadResultTableForCases: '" & TableKey & "' was read, but could " & _
                "NOT be written to '" & SheetName & "' - see the bold red marker " & _
                "on that sheet and the log.", vbExclamation
        Else
            written = 1
            If IsEmpty(data) Then
                Say "ReadResultTableForCases: '" & TableKey & "' produced no data " & _
                    "for " & caseDesc & "." & vbCrLf & vbCrLf & _
                    "Check that the case names match SAFE exactly (casing included) " & _
                    "and that the analysis has been run.", vbExclamation
            Else
                LogMsg "ReadResultTableForCases: '" & TableKey & "' written to '" & _
                       SheetName & "'!" & TopLeftCell & " for " & caseDesc & ", load " & _
                       "cases only, combinations cleared for the read (" & _
                       GetLastExportFailures() & " table(s) failed)."
            End If
        End If
    End If

    GoTo CleanUp

Fail:
    ' Never leave SAFE carrying the display filter this worker changed: the
    ' combination selection is put back by CleanUp below.
    On Error Resume Next               ' the report below must not raise in turn
    reason = CStr(Err.Number) & " - " & Err.Description
    Say "ReadResultTableForCases stopped on an unexpected error: " & reason & _
        vbCrLf & vbCrLf & "ShowLog has the full log.", vbExclamation
    written = -1

CleanUp:
    ' ---- 5. put the display combination selection back ---------------------
    If clearedCombos Then
        On Error Resume Next
        Err.Clear
        If SelectedNameCount(savedCombos) > 0 Then
            ret = gDB.SetLoadCombinationsSelectedForDisplay(savedCombos)
        Else
            ret = gDB.SetLoadCombinationsSelectedForDisplay(none)
        End If
        errNum = Err.Number
        errDesc = Err.Description
        On Error Resume Next           ' the report below must not raise either

        If errNum <> 0 Or ret <> 0 Then
            Say "ReadResultTableForCases: the display load combinations could not " & _
                "be put back the way they were - check the selection in SAFE's " & _
                "Show Tables (" & CStr(errNum) & " / code " & ret & ").", vbExclamation
        ElseIf nSavedCombos > 0 Then
            LogMsg "ReadResultTableForCases: display load combinations restored."
        End If
    End If

    ReadResultTableForCases = written
End Function

' Number of NON-BLANK entries in a name list SAFE returned. A list holding a
' single blank entry is SAFE's way of saying "nothing selected", so that counts 0.
Private Function SelectedNameCount(ByRef names() As String) As Long
    Dim i As Long, n As Long, c As Long

    n = ItemCount(names)
    For i = 0 To n - 1
        If Len(Trim$(names(i))) > 0 Then c = c + 1
    Next i
    SelectedNameCount = c
End Function

' True when a name-list ARGUMENT was effectively not passed at all: Empty or Null,
' a string holding nothing but whitespace, or an array with no elements. An array
' that HAS elements - even blank ones - is NOT empty, and neither is a string with
' separators in it (", ,"): those are caller mistakes, which ReadResultTableForCases
' reports rather than reinterpreting as "all cases".
Private Function IsEmptyNameArgument(ByVal v As Variant) As Boolean
    Dim n As Long

    If IsArray(v) Then
        On Error Resume Next
        n = UBound(v) - LBound(v) + 1
        On Error GoTo 0
        IsEmptyNameArgument = (n <= 0)
        Exit Function
    End If

    IsEmptyNameArgument = (Len(SafeText(v)) = 0)
End Function

' ---------------------------------------------------------------------------
' NODAL REACTIONS - entry point; its settings are LOCAL constants inside it
' ---------------------------------------------------------------------------
' Only the TOP-LEFT cell of the destination is given: the table lands there
' FLUSH - no title row and no header row, because ReadResultTableForCases'
' IncludeHeader defaults to False - and its columns extend right and down from
' there. Pass IncludeHeader:=True for the labelled layout (table key + column
' keys). A table too big for one worksheet continues on <REACT_SHEET>_2, _3, ...
' (SAFE_Library does that, with a warning). Copy this sub - constants and all -
' for another area, e.g. WriteNodalReactions2.
Public Sub WriteNodalReactions1()
    ' ---- EDIT THESE --------------------------------------------------------
    Const REACT_SHEET As String = "Nodal Reactions"     ' <-- EDIT ME
    Const REACT_TOPLEFT As String = "B5"                ' <-- EDIT ME (corner only)

    ' Load CASES to include - NOT load combinations: combinations are cleared for
    ' the duration of the read, so a combination row cannot slip into this table.
    ' Blank entries and duplicates are ignored; casing must match SAFE exactly.
    Const REACT_LOADCASES As String = ""      ' <-- EDIT ME

    ' SAFE's nodal-reaction result table, verified against the shipped key list
    ' reference\SAFE Input&Output Table Key List.csv : "Joint Reactions", Import
    ' Type 0 - a RESULT table, so it can be read here but never written back.
    ' Keys are per model / SAFE version, so ListSAFETables is the way to check
    ' them elsewhere: a wrong key writes NOTHING and writes a bold red FAILED
    ' marker naming the key.
    Const REACT_TABLE As String = "Joint Reactions"     ' <-- from the key list CSV
    ' ------------------------------------------------------------------------

    ReadResultTableForCases REACT_TABLE, REACT_LOADCASES, REACT_SHEET, REACT_TOPLEFT
End Sub

' ===========================================================================
' HELPERS
' ===========================================================================

' Read the three-column range into pfx/xs/ys. Returns the number of usable rows,
' and logs everything it refused to use - a row is never silently absorbed into a
' coordinate. Skipped counts:
'   * a row with coordinates but no prefix;
'   * a row whose X or Y is not a number (a stray header line, a formula that
'     returned text, a cell error);
'   * a repeated prefix (same text ignoring case) - the FIRST occurrence wins,
'     which is the rule the edit pass uses too.
' A completely blank row is NOT counted, so an oversized range does not report a
' long tail of "skipped" rows.
Private Function ReadCoordTable( _
    ByVal vals As Variant, _
    ByRef pfx() As String, _
    ByRef xs() As Double, _
    ByRef ys() As Double, _
    ByRef Skipped As Long) As Long

    Dim nVals As Long
    Dim i As Long, j As Long
    Dim p As String
    Dim tx As String, ty As String
    Dim okX As Boolean, okY As Boolean
    Dim nOut As Long
    Dim dup As Boolean

    Skipped = 0
    nOut = 0
    If Not IsArray(vals) Then Exit Function

    On Error Resume Next
    nVals = UBound(vals, 1)
    On Error GoTo 0
    If nVals < 1 Then Exit Function

    ReDim pfx(1 To nVals)
    ReDim xs(1 To nVals)
    ReDim ys(1 To nVals)

    For i = 1 To nVals
        p = SafeText(vals(i, 1))
        tx = SafeText(vals(i, 2))
        ty = SafeText(vals(i, 3))

        ' IsNumeric is guarded: a cell holding a worksheet error value (#N/A,
        ' #DIV/0! ...) arrives as a Variant/Error and must not raise here.
        okX = False
        okY = False
        On Error Resume Next
        Err.Clear
        okX = IsNumeric(vals(i, 2))
        okY = IsNumeric(vals(i, 3))
        On Error GoTo 0

        If Len(p) = 0 And Len(tx) = 0 And Len(ty) = 0 Then
            ' Blank row inside an oversized range - ignored, not reported.
        ElseIf Len(p) = 0 Then
            Skipped = Skipped + 1
            LogMsg "SetPointCoordinates: SKIPPED row " & i & " of the range - " & _
                   "coordinates without a prefix ('" & tx & "', '" & ty & "')."
        ElseIf Not (okX And okY) Then
            Skipped = Skipped + 1
            LogMsg "SetPointCoordinates: SKIPPED row " & i & " of the range - '" & p & _
                   "' has a non-numeric X or Y ('" & tx & "', '" & ty & "'). " & _
                   "A header row inside the range is expected to land here."
        Else
            dup = False
            For j = 1 To nOut
                If StrComp(pfx(j), p, vbTextCompare) = 0 Then
                    dup = True
                    Exit For
                End If
            Next j

            If dup Then
                Skipped = Skipped + 1
                LogMsg "SetPointCoordinates: SKIPPED row " & i & " of the range - " & _
                       "the prefix '" & p & "' appears more than once; the FIRST " & _
                       "occurrence is used."
            Else
                nOut = nOut + 1
                pfx(nOut) = p
                xs(nOut) = ToDouble(vals(i, 2), tx)
                ys(nOut) = ToDouble(vals(i, 3), ty)
            End If
        End If
    Next i

    If nOut = 0 Then
        Erase pfx
        Erase xs
        Erase ys
    ElseIf nOut < nVals Then
        ReDim Preserve pfx(1 To nOut)
        ReDim Preserve xs(1 To nOut)
        ReDim Preserve ys(1 To nOut)
    End If

    ReadCoordTable = nOut
End Function

' 1-based index of a column key in the Headers array (0 when it is not there).
' Case- and space-insensitive.
Private Function ColIndex(ByRef hdrs() As String, ByVal wanted As String) As Long
    Dim i As Long, n As Long
    Dim probe As String

    probe = Replace(UCase$(Trim$(wanted)), " ", "")
    n = ItemCount(hdrs)
    For i = 0 To n - 1
        If Replace(UCase$(Trim$(hdrs(i))), " ", "") = probe Then
            ColIndex = i + 1                 ' 1-based, for data(row, col)
            Exit Function
        End If
    Next i
    ColIndex = 0
End Function

' Plain prefix test, case-insensitive: "Pile_1" also matches "Pile_10", which is
' why the caller applies the FIRST matching prefix.
Private Function StartsWith(ByVal Text As String, ByVal Prefix As String) As Boolean
    If Len(Prefix) = 0 Then Exit Function
    StartsWith = (StrComp(Left$(Text, Len(Prefix)), Prefix, vbTextCompare) = 0)
End Function

' SAFE parses the table text itself, so the decimal separator must be "." whatever
' the Windows locale is: under a comma locale CStr would give "10,3" and SAFE
' would read that as text (or cut it at the comma). Format with up to 12 decimals
' also avoids the scientific notation CStr can produce for small values.
Private Function NumToText(ByVal v As Double) As String
    NumToText = Replace$(Format$(v, "0.############"), ",", ".")
End Function

' Back from SAFE's text form to a number for the verification step. Val() always
' reads "." as the decimal separator, which is what SAFE writes.
Private Function NumFromText(ByVal t As String) As Double
    NumFromText = Val(Trim$(t))
End Function

' A numeric cell arrives as a Double already; a numeric-looking TEXT cell is
' converted with Val(), because CDbl on a String follows the locale and would
' fail on "10.3" under a comma decimal locale.
Private Function ToDouble(ByVal v As Variant, ByVal text As String) As Double
    On Error Resume Next
    Err.Clear
    ToDouble = 0
    If IsEmpty(v) Then Exit Function
    ToDouble = CDbl(v)
    If Err.Number <> 0 Then
        Err.Clear
        ToDouble = Val(text)
    End If
End Function

Private Function NearlyEqual(ByVal a As Double, ByVal b As Double) As Boolean
    Dim tol As Double

    tol = COORD_ABS_TOL
    If Abs(b) * COORD_REL_TOL > tol Then tol = Abs(b) * COORD_REL_TOL
    NearlyEqual = (Abs(a - b) <= tol)
End Function

' Cell value as text: Empty becomes "", a worksheet error value becomes "" (its
' CStr would raise), everything else is trimmed. Used everywhere a cell has to be
' turned into a string or tested for emptiness.
Private Function SafeText(ByVal v As Variant) As String
    SafeText = ""
    On Error Resume Next
    If Not IsEmpty(v) Then SafeText = Trim$(CStr(v))
End Function

' Best effort: the unit string SAFE reports for one column of a table (e.g.
' "mm"), read with GetAllFieldsInTable. Returns Fallback when it cannot be read -
' purely informational, and it must never stop the edit. SAFE carries the model's
' PRESENT units in its tables, which is what this reports.
Private Function TableFieldUnits( _
    ByVal TableKey As String, _
    ByVal FieldKey As String, _
    Optional ByVal Fallback As String = "unknown") As String

    Dim ret As Long
    Dim TableVersion As Long
    Dim nFields As Long
    Dim fKey() As String
    Dim fName() As String
    Dim fDesc() As String
    Dim fUnits() As String
    Dim fImportable() As Boolean
    Dim i As Long, n As Long

    TableFieldUnits = Fallback
    If gDB Is Nothing Then Exit Function

    On Error GoTo NoUnits
    ret = gDB.GetAllFieldsInTable(TableKey, TableVersion, nFields, _
                                  fKey, fName, fDesc, fUnits, fImportable)
    If ret <> 0 Then Exit Function

    n = ItemCount(fKey)
    If nFields > 0 And n > nFields Then n = nFields
    For i = 0 To n - 1
        If StrComp(Trim$(fKey(i)), Trim$(FieldKey), vbTextCompare) = 0 Then
            If ItemCount(fUnits) > i Then
                If Len(Trim$(fUnits(i))) > 0 Then TableFieldUnits = Trim$(fUnits(i))
            End If
            Exit Function
        End If
    Next i
    Exit Function

NoUnits:
    TableFieldUnits = Fallback
End Function

' Length of a dynamic String array; 0 when it is not dimensioned. This module is
' a separate VBA module, so it needs its own copy of the guard SAFE_Library uses
' internally (that one is private to SAFE_Library).
Private Function ItemCount(ByRef a() As String) As Long
    On Error GoTo NoArr
    ItemCount = UBound(a) - LBound(a) + 1
    Exit Function
NoArr:
    ItemCount = 0
End Function
