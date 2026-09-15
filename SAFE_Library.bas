Attribute VB_Name = "SAFE_Library"
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  SAFE_Library.bas  --  CSI SAFE database-table extractor for Excel / VBA
' ============================================================================
'  WHAT IT DOES
'    Attaches to a RUNNING CSI SAFE instance (the model already open in SAFE),
'    reads one or more database tables, and writes them into this workbook at a
'    worksheet ("tab") and top-left coordinate you pass in as parameters.
'
'  REQUIREMENTS
'    - SAFE must be OPEN with the desired model loaded.
'      This script ATTACHES to the running instance. It never starts SAFE,
'      never opens a model, and never closes SAFE (so your session is safe).
'    - A reference to SAFEv1.tlb must be added in the Excel VBA IDE:
'        Alt+F11 -> Tools > References... -> tick "SAFEv1" (or Browse... to
'        SAFEv1.tlb in the SAFE installation folder). The reference is read
'        from this VBA IDE setting - no file path is hardcoded in the script.
'
'  QUICK START
'    1. Open SAFE, load your model, leave it running.
'    2. In Excel: Alt+F11 -> Tools > References... -> add SAFEv1.tlb (once).
'    3. Alt+F11 -> File > Import File... -> choose this .bas
'    4. Run macro  DemoExport     (or call ExportSAFETables yourself).
'
'  MAIN API
'    ExportSAFETables(Tables, SheetName, StartCell, ...)
'        Tables     : a single table key (String) OR an array of keys (String())
'                     e.g. "Element Forces - Area Shells"
'                     or   Array("Point Object Connectivity", "Area Load Assignments - Uniform")
'                     or   "Table A, Table B"   (comma separated string)
'        SheetName  : destination worksheet ("tab") - created if missing
'        StartCell  : top-left cell, e.g. "B3"
'        StackHorizontally : False = tables stacked downwards (default),
'                            True  = tables placed side by side
'        IncludeHeader     : write the table's title row AND its column-header
'                            row (default True). False = DATA ONLY: neither the
'                            title nor the header row is written, so the data
'                            block starts ON StartCell, and the column keys are
'                            not even fetched from SAFE
'        LoadCases   (new) : optional load-CASE filter - only these load cases
'                            appear in result tables. Accepts one name, a
'                            comma-separated list, or a String() array.
'                            Empty / default = export ALL load cases.
'        LoadCombos  (new) : optional load-COMBINATION filter - same formats as
'                            LoadCases, but for combinations, e.g.
'                            "1.4DL+1.6LL" or "1.4(D+WH)" (the pile-cap sheet
'                            filters its results by COMBINATION, so this is
'                            the one it needs). SAFE keeps load cases and load
'                            combinations as two SEPARATE display lists, so a
'                            combination name belongs in LoadCombos and a case
'                            name in LoadCases - neither list accepts the
'                            other's names. Empty / default = ALL combinations.
'        FILTERS    : SAFE's display filter only selects what the RESULT tables
'                     show - non-result tables are unaffected.
'                     In a name list, blank entries and exact duplicates are
'                     ignored; everything else is sent to SAFE unchanged
'                     (SAFE matches names exactly, so casing is preserved).
'                     The selection in force BEFORE the export is saved first
'                     and put back afterwards, on the normal exit AND on the
'                     error path, so your SAFE session is never left filtered
'                     (a previously empty selection is restored as "none
'                     selected", the documented way to select nothing).
'                     An unknown / mis-spelled name is FATAL (-1): the export
'                     stops rather than quietly returning UNFILTERED force
'                     results. The log names the rejected names and the exact
'                     return code.
'        Returns    : number of tables written; -1 on a fatal error
'                     (see Immediate window / ShowLog for details)
'        FAILURES   : a table is marked FAILED - and is NOT counted in the
'                     number written - when its read cannot be served: a
'                     NONZERO code from GetTableForDisplayArray with NO column
'                     headers returned (the table key is not valid for this
'                     model / SAFE version, or the table is a result table
'                     whose analysis has not been run), or a COM error. A
'                     failed table writes a bold-red cell at the current
'                     cursor: "Table '<key>' : FAILED - <reason>", is logged
'                     as an ERROR, and does NOT stop the export - the
'                     remaining tables are still written. This is deliberately
'                     different from a valid key with no rows (an EMPTY table),
'                     which writes the plain "no data returned" marker.
'                     GetLastExportFailures() reports how many tables failed in
'                     the last run - a failed READ and a block the worksheet
'                     could not hold (a WRITE failure) are both counted.
'        LIMITS     : the data is written into ordinary worksheet cells, so
'                     EXCEL'S OWN LIMITS apply - and they are READ AT RUNTIME
'                     from the sheet (ws.Rows.Count / ws.Columns.Count), never
'                     hardcoded, because they depend on the file format
'                     (1,048,576 rows x 16,384 columns for .xlsx/.xlsm,
'                     65,536 x 256 for .xls). A table that does not fit on one
'                     worksheet is CONTINUED automatically on the sheets
'                     <SheetName>_2, <SheetName>_3, ... (the name of the sheet
'                     the table starts on, so a second spilled table nests off
'                     the sheet it starts on) with the title and the
'                     column-header row (whichever of the two is being written)
'                     REPEATED on every sheet, so each sheet is self-contained
'                     and can be read on its own; a WARNING
'                     is logged before the first continuation sheet, and the
'                     cursor returned to you moves to the sheet it ends on.
'                     NOTHING is ever truncated silently: a table WIDER than
'                     the sheet, or a StartCell so low that not even the title
'                     and header row fit, is refused with an ERROR in the log, a
'                     BOLD RED FAILED marker at the cursor, and is counted by
'                     GetLastExportFailures() - the same treatment a failed READ
'                     gets, and it is not counted as a table that was written.
'
'    ListSAFETables(SheetName, StartCell)   : dumps EVERY table SAFE reports
'                                             (key | name | import type) - the
'                                             GetAllTables superset, not just
'                                             the tables available for display
'                                             (the log also compares the two
'                                             counts)
'    WriteSAFETable(TableKey, Data, ...)    : writes a 2-D array BACK into SAFE
'                                             (edit + apply). Optional bonus.
'    SAFEReadEditingTable(TableKey, Headers, Data)
'                                           : PUBLIC read bridge for companion
'                                             modules - an EDITING-table read
'                                             into Headers() plus a 1-based 2-D
'                                             array (no header row), which is
'                                             exactly the form WriteSAFETable
'                                             accepts back. Read-only: it never
'                                             touches the model's lock state.
'    LogMsg(msgstring)                      : PUBLIC - appends a line to THIS
'                                             module's log, so a companion
'                                             module's messages appear in
'                                             ShowLog / GetLog beside the
'                                             library's own.
'    gSAFE / gSapModel / gDB / gConnected   : PUBLIC module state (see the state
'                                             block below) for any SAFE API call
'                                             this library does not wrap.
'    SAFEConnect / SAFEDisconnect           : attach / release the running SAFE
'    LastErrorNumber() / LastErrorDescription() / LastErrorContext() /
'    LastErrorText()                        : the last error that was recorded
'                                             by LogError / SetLastError (below)
'    GetLastExportFailures()                : how many tables FAILED (bold red
'                                             markers) in the most recent
'                                             ExportSAFETables call - 0 = none.
'                                             ClearLog / GetLog / ShowLog give
'                                             the full log text.
'
'  TABLE KEYS are the same strings shown in SAFE's "Display > Show Tables",
'  e.g. "Point Object Connectivity", "Area Load Assignments - Uniform",
'  "Element Forces - Area Shells", "Joint Displacements", ...
'  Run ListSAFETables to see the exact keys available in your model.
'
'  QUIRKS LEARNED (from the reference Python script + SAFE API docs)
'    - SAFE's COM main object is called "ETABSObject" (ETABS infrastructure).
'    - GetTableForDisplayArray: "nonzero means error OR nothing to show" -
'      the two are distinguished by whether column headers were returned.
'      Nonzero WITH headers is only a WARNING (data still came back / the table
'      is empty), and the export keeps going. Nonzero with NO headers at all is
'      a HARD FAILURE: CSI documents "if there is nothing to be shown in the
'      table then no data is returned", so a valid key would still have
'      reported its columns - the key is invalid for this model / SAFE
'      version, or analysis has not been run. That table is marked in BOLD RED
'      in the sheet, logged as an error, excluded from the function's count and
'      reported by GetLastExportFailures(), but the export still continues with
'      the remaining tables. A ZERO return with headers but no records is the
'      genuinely EMPTY table (plain "no data returned" marker). The reference
'      Python script treats ANY nonzero return as fatal; this module only does
'      that for the no-headers case, so one bad key cannot cost the whole run.
'    - FieldKeyList must be a single blank string to get ALL columns.
'    - GroupName "" (or "All") returns data for all objects in the model.
'    - The data array is FLATTENED row-by-row; rows are rebuilt using the
'      number of columns = number of FieldsKeysIncluded.
'    - Some cells come back as empty strings; written as blank cells.
'    - For editing tables, GroupName is documented as "not active in this
'      release" - pass "".
'    - Editing tables requires the model UNLOCKED - but only for the tables
'      SAFE reports as interactively importable while the model is unlocked
'      (GetAllTables ImportType = 2). WriteSAFETable preflights that import
'      type BEFORE it touches the lock, unlocks only when the preflight says it
'      must (or, for a table it could not preflight, when the UnlockModel
'      argument says so), and RESTORES the lock state it found on the success
'      path AND on every failure path - a write-back can never leave the model
'      unlocked. ImportType 3 is importable while locked or unlocked, so no
'      unlock is needed; 0 (not importable) and 1 (not INTERACTIVELY importable)
'      cannot be written back through the editing-table pair at all and are
'      refused without touching the lock.
'    - Applying an edit makes any existing analysis results STALE: re-run the
'      analysis in SAFE afterwards. This module deliberately never runs the
'      analysis (and never saves the model) for you.
'    - SetTableForEditingArray's TableVersion is a RETURNED (out) item, not an
'      input - the docs say "Returned Item: The version number of the specified
'      table", and the reference implementation passes 0 for it (and ignores
'      what comes back). Never feed the version read by GetTableForEditingArray
'      back in: pass 0 and capture the returned value; a mismatch is only
'      logged as information, never treated as a failure.
'    - ApplyEditedTables can corrupt the model on a fatal error - SAFE docs
'      recommend saving the model BEFORE calling it. We check the error counts:
'      a FAILURE is a nonzero return OR NumFatalErrors > 0. Error / warning /
'      info messages on their own are NOT a failure - they are logged as a
'      warning and the edit is still treated as applied.
'    - ApplyEditedTables' FillImportLog argument MUST be passed as True to get a
'      non-empty ImportLog: with False the counts still come back but the log
'      string is always empty (which is why a failure used to be reported with
'      no text). SAFE warns that the log MAY BE VERY LARGE, so this module
'      copies only the first IMPORT_LOG_MAX characters into its own log, with
'      a " ... [truncated]" marker; the full text stays available in SAFE.
'    - EXCEL LIMITS / SHEET SPILL: a table larger than one worksheet is written
'      across <SheetName>, <SheetName>_2, <SheetName>_3, ... (the sheets are
'      created on demand by GetWorksheet; "SheetName" is the sheet the table
'      STARTS on, so a second spilled table nests off the sheet it starts on and
'      cannot reuse a sheet that already carries rows). Excel's 31-character
'      sheet-name limit is respected by cutting the BASE name short and keeping
'      the "_N" suffix intact - truncating the whole combination could collapse
'      the name back onto a sheet that already holds rows. Every sheet
'      repeats the title and the header row (when they are being written), so a
'      sheet can be read on its own.
'      A WARNING is logged before the first continuation sheet ("<n> data rows
'      exceed the <limit>-row worksheet limit - continuing on additional
'      worksheets"), one line per continuation sheet, and the total number of
'      rows written when a spill happened. The limits are read AT RUNTIME from
'      the sheet itself (CurSheet.Rows.Count / CurSheet.Columns.Count) because
'      they differ between .xlsx/.xlsm (1,048,576 x 16,384) and .xls
'      (65,536 x 256). The ordinary case is untouched: it is still ONE bulk
'      Range.Value2 write of the whole array. Slicing (building a fresh 1-based
'      2-D Variant chunk, row by row) happens ONLY in the spill path, so a
'      normal export loses no speed. A table that is WIDER than the worksheet,
'      and a StartCell with no room left for the title + header rows, are
'      REFUSED (ERROR in the log + a bold red FAILED marker + reported as a
'      FAILED table by GetLastExportFailures, never as 0 rows written) rather
'      than truncated or allowed to raise a run-time error.
'    - When attaching to a running instance, NEVER call ApplicationExit
'      (it would close the user's SAFE session).
'    - CONNECTION ERRORS (Readme.txt troubleshooting): a "438 or 5" failure
'      usually means the SAFEv1 reference is not ticked, or SAFE is not open /
'      is busy analysing / is not locked. Every trappable error is recorded
'      with its number AND its 8-digit HRESULT by LogError, and can be read
'      back with LastErrorNumber / LastErrorDescription / LastErrorContext.
'    - SAFEConnect() PROVES the link is live before reporting success, by
'      calling SapModel.GetModelFilepath() - a cheap read-only call. A stale
'      Running Object Table entry (SAFE closed or crashed after connecting)
'      survives the GetObject attach and only fails on first use, so without
'      the probe a dead proxy could be cached as "connected".
' ============================================================================

' ---------------------------------------------------------------------------
' CONFIGURATION ZONE
' ---------------------------------------------------------------------------
' ProgID of the SAFE API main object (note: named after ETABS infrastructure).
Private Const SAFE_PROGID As String = "CSI.SAFE.API.ETABSObject"

' ProgID of the API Helper object (informational - early binding below uses
' the Helper coclass directly via "New Helper").
Private Const SAFE_HELPER_PROGID As String = "CSI.SAFE.API.Helper"

' ---------------------------------------------------------------------------
' [COMMENTED OUT - HARDCODED IMPORT SECTION]
' The SAFE type-library reference is now read from the Excel VBA IDE setting
' (Alt+F11 -> Tools > References... -> add SAFEv1.tlb), so no file path is
' hardcoded here any more. The old hardcoded paths are kept below for
' reference only.
' ---------------------------------------------------------------------------
'Private Const SAFE_TLB_PATH As String = _
'    "C:\Program Files\Computers and Structures\SAFE 20\SAFEv1.tlb"
'Private Const SAFE_DLL_PATH As String = _
'    "C:\Program Files\Computers and Structures\SAFE 20\SAFEv1.dll"

' Default destination used by the Demo macros.
Private Const DEF_SHEET As String = "SAFE Tables"
Private Const DEF_START As String = "A1"

' Maximum number of characters of SAFE's import log (ApplyEditedTables) that is
' copied into this module's log. SAFE documents the log as possibly "very
' large", so the text is truncated with an explicit marker and the full log is
' left to SAFE itself.
Private Const IMPORT_LOG_MAX As Long = 2000

' ---------------------------------------------------------------------------
' Module state (connection + log + display filter)
' ---------------------------------------------------------------------------
' The connection objects below are deliberately PUBLIC so that companion modules
' of site-specific subs (e.g. SAFE_Use.bas) can use the live SAFE
' objects directly - for API calls this library does not wrap, or to read a table
' it does not read - WITHOUT a second attach to SAFE and without a wrapper
' function for every call. VBA module scope gives no real protection here anyway:
' anything in the project can already obtain the same objects through GetObject
' or New Helper.
' Keep the "g" prefix when referencing them from another module (gSAFE /
' gSapModel / gDB / gConnected) so it stays obvious which module owns them and
' that they are only meaningful while gConnected is True - and always call
' SAFEConnect() first, which is idempotent.
Public gSAFE As cOAPI            ' SAFE API object
Public gSapModel As cSapModel    ' model object
Public gDB As cDatabaseTables    ' database tables (cached)
Public gConnected As Boolean     ' True between SAFEConnect and SAFEDisconnect
Private gLog As String

' Number of tables that FAILED in the most recent ExportSAFETables run (a
' nonzero read code with NO column headers, or a COM error - see Fix 4 in
' SAFETableToArray). Reset to 0 at the start of every run, incremented once per
' failed table, and read back with GetLastExportFailures(). Failed tables are
' marked in bold red in the destination sheet and are NOT part of the count
' ExportSAFETables returns.
Private mFailedTables As Long

' Display-filter state (the load cases / load combinations selected for table
' display). Captured by SaveDisplayFilterState before a filtered export and put
' back by RestoreDisplayFilterState on EVERY exit path, so a filtered export
' can never leave the user's SAFE session with a changed display selection.
Private mSavedCases() As String          ' selection that was in force
Private mSavedCaseCount As Long          ' 0 = nothing was selected
Private mSavedCombos() As String
Private mSavedComboCount As Long
Private mFilterCasesApplied As Boolean   ' the case setter actually changed it
Private mFilterCombosApplied As Boolean  ' the combo setter actually changed it
Private mFilterStateValid As Boolean     ' a saved state is waiting to be restored

' Last error recorded by LogError / SetLastError (see the LOGGING section).
Private gLastErrNumber As Long
Private gLastErrSource As String
Private gLastErrDesc As String
Private gLastErrContext As String

' ===========================================================================
' CONNECTION - attach to the running SAFE instance
' ===========================================================================

Public Function SAFEConnect() As Boolean
    Dim helper As cHelper
    Dim ret1 As Long, desc1 As String
    Dim ret2 As Long, desc2 As String
    Dim modelPath As String

    ' Strategy 1: attach to the running instance through the Running Object
    ' Table (VBA's GetObject with an empty path). The error NUMBER is recorded
    ' rather than discarded: "no active instance" is expected here, but shares
    ' its error numbers (429 / 0x800401E3) with genuine failures, so it has to
    ' be visible in the log.
    On Error Resume Next
    Err.Clear
    Set gSAFE = GetObject(, SAFE_PROGID)
    ret1 = Err.Number
    desc1 = Err.Description
    On Error GoTo 0

    ' Strategy 2: via the API Helper object (documented early-bound approach).
    ' Unlike GetObject it returns Nothing instead of raising when no instance is
    ' registered - any error is still recorded.
    If gSAFE Is Nothing Then
        On Error Resume Next
        Err.Clear
        Set helper = New Helper
        ret2 = Err.Number
        desc2 = Err.Description
        If ret2 = 0 And Not helper Is Nothing Then
            Set gSAFE = helper.GetObject(SAFE_PROGID)
            ret2 = Err.Number
            desc2 = Err.Description
        End If
        On Error GoTo 0
        Set helper = Nothing
    End If

    If gSAFE Is Nothing Then
        If ret1 <> 0 Then
            LogMsg "SAFEConnect: strategy 1 (GetObject via the ROT) failed - " & _
                   ErrorText(ret1, desc1)
        Else
            LogMsg "SAFEConnect: strategy 1 (GetObject via the ROT) found no instance."
        End If
        If ret2 <> 0 Then
            LogMsg "SAFEConnect: strategy 2 (API Helper) failed - " & _
                   ErrorText(ret2, desc2)
        End If
        LogMsg "SAFEConnect: no running SAFE instance found." & vbCrLf & _
               "Start SAFE, open the model, and try again." & vbCrLf & _
               "  " & ErrHint(438)
        SAFEConnect = False
        Exit Function
    End If

    ' Bind the model + database-table objects, then PROVE the connection is
    ' live with a cheap, read-only call. A stale Running Object Table entry
    ' still resolves here and only fails on first use.
    On Error GoTo Fail
    Set gSapModel = gSAFE.SapModel
    If gSapModel Is Nothing Then
        ' Connection-level failure: the Readme's 438 advice (reference ticked,
        ' SAFE open / idle / locked) is the right hint here.
        SetLastError "SAFEConnect", "the attached instance exposed no SapModel object " & _
                     "(stale or incompatible SAFE instance).", ErrHint(438)
        GoTo Fail
    End If

    modelPath = gSapModel.GetModelFilepath()      ' fast, read-only liveness probe

    Set gDB = gSapModel.DatabaseTables
    If gDB Is Nothing Then
        SetLastError "SAFEConnect", "the attached instance exposed no DatabaseTables object.", _
                     ErrHint(438)
        GoTo Fail
    End If

    gConnected = True
    ClearLastError
    If Len(modelPath) = 0 Then modelPath = "(untitled / unsaved model)"
    LogMsg "Attached to running SAFE instance (progid: " & SAFE_PROGID & ")." & vbCrLf & _
           "  Liveness probe OK (GetModelFilepath) - model: " & modelPath
    SAFEConnect = True
    Exit Function

Fail:
    ' LogError only records genuine COM failures; the explicit checks above
    ' reach this label with Err.Number = 0 and have already been logged.
    LogError "SAFEConnect", "Connection dropped (SAFE references released)."
    gConnected = False
    Set gDB = Nothing
    Set gSapModel = Nothing
    Set gSAFE = Nothing
    SAFEConnect = False
End Function

Public Sub SAFEDisconnect()
    ' Release COM references. Does NOT close SAFE.
    Set gDB = Nothing
    Set gSapModel = Nothing
    Set gSAFE = Nothing
    gConnected = False
    LogMsg "Disconnected from SAFE (SAFE left running)."
End Sub

' ===========================================================================
' MAIN FUNCTION - extract one or more tables and write them to Excel
' ===========================================================================

Public Function ExportSAFETables( _
    ByVal Tables As Variant, _
    ByVal SheetName As String, _
    ByVal StartCell As String, _
    Optional ByVal StackHorizontally As Boolean = False, _
    Optional ByVal IncludeHeader As Boolean = True, _
    Optional ByVal LoadCases As Variant = "", _
    Optional ByVal LoadCombos As Variant = "") As Long
    ' Returns the number of tables written, or -1 on a fatal error.
    ' LoadCases / LoadCombos: optional display filters - only these load CASES
    ' and load COMBINATIONS appear in result tables ("" / missing = export all
    ' of them; non-result tables are unaffected by SAFE's display filter).
    ' A name SAFE does not recognise is fatal: the export stops with -1 instead
    ' of silently writing unfiltered force results. Whatever filter is applied
    ' here is removed again on every exit path (RestoreDisplayFilterState).

    Dim savedScreen As Boolean
    savedScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fatal

    ' reset the failure counter HERE - before the first way out of
    ' this function. It used to be reset much further down, i.e. AFTER the "no
    ' table names" return, the invalid-StartCell return and a failed SAFEConnect,
    ' so those exits left GetLastExportFailures() reporting the PREVIOUS run's
    ' count. (The identical prologue in ListSAFETables deliberately does NOT
    ' touch it: listing tables is not an export run.)
    mFailedTables = 0

    If Not gConnected Then
        If Not SAFEConnect() Then
            Application.ScreenUpdating = savedScreen
            ExportSAFETables = -1
            Exit Function
        End If
    End If

    Dim names() As String
    Dim nTables As Long
    nTables = NormalizeTables(Tables, names)
    If nTables <= 0 Then
        LogMsg "ExportSAFETables: no table names were supplied."
        Application.ScreenUpdating = savedScreen
        ExportSAFETables = 0
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = GetWorksheet(ThisWorkbook, SheetName)

    Dim r As Long, c As Long
    If Not ParseStartCell(ws, StartCell, r, c) Then
        LogMsg "ExportSAFETables: invalid StartCell '" & StartCell & "'."
        Application.ScreenUpdating = savedScreen
        ExportSAFETables = -1
        Exit Function
    End If

    Dim i As Long
    Dim written As Long
    written = 0
    Dim hdrs() As String
    Dim warn As String
    Dim tblFailed As Boolean
    Dim data As Variant
    ' FIX 7: 'ws' is handed to WriteTableBlock ByRef and comes back pointing at
    ' the sheet the cursor ended up on (a table too large for one worksheet is
    ' continued on <SheetName>_2, _3, ...). Every marker below is therefore
    ' written through 'ws', and 'markerSheet' remembers where the LAST failure
    ' marker went, so the end-of-run message cannot name a continuation sheet
    ' that holds no marker just because a later table spilled onto it.
    Dim markerSheet As String

    ' --- optional load-case / load-combination display filter --------------
    ' The names are cleaned (trimmed, blank entries and exact duplicates
    ' dropped, casing preserved) because SAFE matches them exactly, and load
    ' CASES and load COMBINATIONS are two SEPARATE lists. The selection in force
    ' right now is saved before any setter runs, so RestoreDisplayFilterState can
    ' put it back on every exit path. A name SAFE does not recognise makes its
    ' setter fail, which is treated as fatal (see ApplyDisplayFilterState).
    Dim rawNames() As String
    Dim nRaw As Long
    Dim caseNames() As String
    Dim nCases As Long
    Dim comboNames() As String
    Dim nCombos As Long

    nRaw = NormalizeTables(LoadCases, rawNames)
    nCases = CleanNameList(rawNames, nRaw, caseNames)
    nRaw = NormalizeTables(LoadCombos, rawNames)
    nCombos = CleanNameList(rawNames, nRaw, comboNames)

    If nCases > 0 Or nCombos > 0 Then
        SaveDisplayFilterState
        If Not ApplyDisplayFilterState(caseNames, nCases, comboNames, nCombos) Then
            ' Fatal: the Fatal handler removes any filter that was already applied.
            GoTo Fatal
        End If
    End If

    For i = 0 To nTables - 1
        Dim key As String
        key = Trim(names(i))
        If Len(key) > 0 Then
            ' Clear any leftover error first, so the Err object can tell THIS
            ' table's COM failure apart from an earlier one in the same run.
            Err.Clear
            tblFailed = False
            ' ReturnHeaders follows IncludeHeader: with IncludeHeader = False the
            ' caller wants a bare data block, so the column keys are not fetched
            ' either. The normal-case test below still works without them, because
            ' a table that has rows has a non-empty data array, and WriteTableBlock
            ' takes its column count from that array.
            data = SAFETableToArray(key, hdrs, warn, tblFailed, ReturnHeaders:=IncludeHeader)

            If Not tblFailed And Len(warn) > 0 Then
                ' Not a failure: the API code was nonzero but data still came
                ' back, or the table genuinely has no rows. Log and carry on.
                LogMsg "[" & key & "] " & warn
            End If

            If tblFailed Then
                ' HARD FAILURE - deliberately NOT the empty-table path. A nonzero
                ' read code came back with NO column headers, so the request
                ' could not be served: the key is not valid for this model / SAFE
                ' version, or the table is a result table whose analysis has not
                ' been run. Counted separately, logged as an ERROR, marked
                ' unmistakably in the sheet, and the export CONTINUES with the
                ' next table (the failed table is not counted as written).
                mFailedTables = mFailedTables + 1

                ' LogError records a genuine COM failure with its number and
                ' 8-digit HRESULT. An API-level failure has no Err number at all,
                ' so SetLastError records the table key and the warning text
                ' (LogError would silently do nothing in that case).
                If Err.Number <> 0 Then
                    LogError "ExportSAFETables", "Table '" & key & "' : " & warn
                Else
                    ' API-level failure (no error number): the 438 / "check the
                    ' SAFEv1 reference" hint would point at the WRONG cause here,
                    ' so this call supplies its own hint (see SetLastError).
                    SetLastError "ExportSAFETables", "Table '" & key & "' : " & warn, _
                                 "Hint: run ListSAFETables (or DemoListTables) for the " & _
                                 "exact table keys of THIS model, and check that the " & _
                                 "analysis has been run for result tables."
                End If

                With ws.Cells(r, c)
                    .Value = "Table '" & key & "' : FAILED - " & warn
                    .Font.Bold = True
                    .Font.Color = vbRed
                End With
                markerSheet = ws.Name   ' for the summary message at the end
                r = r + 2              ' same cursor advance as the empty path
            ElseIf ArrLenStr(hdrs) > 0 Or Not IsEmpty(data) Then
                ' Normal case: write the block - title row and header row first
                ' when IncludeHeader asks for them (it is handed to WriteTableBlock
                ' twice, as WithTitle and as WithHeader), then the data.
                ' IncludeHeader:=False therefore suppresses BOTH, and the data
                ' block starts on the start cell.
                ' FIX 7: 'ws' is passed ByRef as the block's starting sheet and
                ' is updated to the sheet the cursor ends up on, so everything
                ' written after this point stays on the right sheet when the
                ' table was too large for one worksheet and was continued on
                ' <SheetName>_2, _3, ... The base name handed over is the name of
                ' the sheet the cursor is on NOW, so a second spilled table -
                ' which starts on the last sheet the first one used - derives its
                ' own continuation names from that sheet and can never reuse a
                ' sheet that already carries rows.
                ' DEFECT 2 FIX: a NEGATIVE return means the block could NOT be
                ' written in full - it is too wide for a worksheet, there is no
                ' room left for its title/header rows, or a continuation sheet ran
                ' out of room part-way through the spill. WriteTableBlock has
                ' already written a BOLD RED marker at the cursor and stepped past
                ' it, so this branch must NOT write a second marker: it counts the
                ' table as FAILED, records the sheet the marker went to, and logs
                ' it - exactly like a failed read (Fix 4). A return of 0 is NOT a
                ' failure - that is a valid empty table (title + header only, when
                ' they are being written).
                If WriteTableBlock(ThisWorkbook, ws, ws.Name, r, c, key, hdrs, data, _
                                   IncludeHeader, IncludeHeader, _
                                   StackHorizontally, r, c) < 0 Then
                    mFailedTables = mFailedTables + 1
                    SetLastError "ExportSAFETables", "Table '" & key & "' : the data " & _
                        "was read from SAFE but could NOT be written to the sheet - " & _
                        "see the bold red marker on sheet '" & ws.Name & "' and the " & _
                        "ERROR lines logged above.", _
                        "Hint: start the block closer to A1 (or on a sheet with more " & _
                        "room) so that the block and all of the table's columns " & _
                        "fit on the worksheet - with IncludeHeader:=False there is " & _
                        "no title row and no header row, so only the data has to fit."
                    markerSheet = ws.Name
                Else
                    written = written + 1
                End If
            Else
                ' Valid key, genuinely empty table - a visibly different marker,
                ' so an EMPTY table and a FAILED table cannot be confused.
                ' Written through the sheet the cursor is on now (unchanged from
                ' the starting sheet unless a previous table spilled).
                ws.Cells(r, c).Value = "Table '" & key & "' : no data returned"
                If Len(warn) > 0 Then ws.Cells(r, c + 1).Value = warn
                r = r + 2
            End If
        End If
    Next i

    ' Put the display selection back exactly as it was found. The restore is
    ' idempotent and the Fatal path calls it too, so a failed export cannot
    ' leave the user's SAFE session with a changed display filter.
    RestoreDisplayFilterState

    ' The run finished, but some tables did NOT: say so explicitly, name the
    ' sheet carrying the red markers, and point at the way to get valid keys.
    ' (Failed tables are not part of the returned count.)
    If mFailedTables > 0 Then
        ' 'markerSheet' - not 'ws', which may since have moved to a continuation
        ' sheet because a later table was too large for one worksheet.
        LogMsg "WARNING: ExportSAFETables: " & written & " table(s) written and " & _
               mFailedTables & " table(s) FAILED (marked in BOLD RED on sheet '" & _
               markerSheet & "'). A FAILED READ - a nonzero read code with no " & _
               "column headers - means the table key is not valid for this model " & _
               "/ SAFE version, or the table is a result table whose analysis has " & _
               "not been run (run ListSAFETables / DemoListTables for the valid " & _
               "keys). The same red marker names a WRITE failure when the data WAS " & _
               "read but the worksheet could not hold the whole block. " & _
               "See GetLastExportFailures()."
    End If

    Application.ScreenUpdating = savedScreen
    ExportSAFETables = written
    Exit Function

Fatal:
    RestoreDisplayFilterState
    Application.ScreenUpdating = savedScreen
    LogError "ExportSAFETables"
    ExportSAFETables = -1
End Function

' ===========================================================================
' DIAGNOSTIC - list EVERY table SAFE reports (exact strings to use above)
' ===========================================================================
' Uses GetAllTables, NOT GetAvailableTables: GetAllTables reports ALL of the
' tables SAFE knows about for this model / version, whereas GetAvailableTables
' reports only those "currently available for display". The two lists are not
' the same, and telling them apart is exactly what this listing is for - so the
' log also records BOTH counts side by side. The sheet holds the GetAllTables
' list. GetAllTables additionally returns an IsEmpty flag per table ("True means
' there is no data in the model to fill the table"); it is deliberately NOT
' written to the sheet here, so an empty table still appears in full.

Public Function ListSAFETables( _
    Optional ByVal SheetName As String = DEF_SHEET, _
    Optional ByVal StartCell As String = DEF_START) As Long
    ' Writes "Table Key | Table Name | Import Type" for EVERY table SAFE reports.
    ' Returns the number of tables listed, or -1 on error.

    Dim savedScreen As Boolean
    savedScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fatal

    If Not gConnected Then
        If Not SAFEConnect() Then
            Application.ScreenUpdating = savedScreen
            ListSAFETables = -1
            Exit Function
        End If
    End If

    Dim ws As Worksheet
    Set ws = GetWorksheet(ThisWorkbook, SheetName)

    Dim r As Long, c As Long
    If Not ParseStartCell(ws, StartCell, r, c) Then
        Application.ScreenUpdating = savedScreen
        ListSAFETables = -1
        Exit Function
    End If

    Dim NumberTables As Long
    Dim TableKey() As String
    Dim TableName() As String
    Dim ImportType() As Long
    ' GetAllTables' fifth argument: the per-table "no data in the model" flag.
    ' NOT written to the sheet (this listing is deliberately unfiltered), but it
    ' has to be dimensioned and passed - the API signature requires all five.
    Dim IsEmptyList() As Boolean
    Dim ret As Long

    ret = gDB.GetAllTables(NumberTables, TableKey, TableName, ImportType, IsEmptyList)
    If ret <> 0 Then
        LogMsg "ListSAFETables: GetAllTables returned " & ret
        Application.ScreenUpdating = savedScreen
        ListSAFETables = -1
        Exit Function
    End If

    ' Informational: also log how many tables the AVAILABLE-FOR-DISPLAY call
    ' reports, so both figures can be compared from a single run without writing
    ' a second sheet. Best effort throughout - nothing here can affect the list.
    Dim nAvail As Long
    Dim availKey() As String
    Dim availName() As String
    Dim availImport() As Long
    Dim availRet As Long
    Dim availErr As Long, availDesc As String
    Dim countNote As String

    nAvail = -1
    On Error Resume Next
    Err.Clear
    availRet = gDB.GetAvailableTables(nAvail, availKey, availName, availImport)
    availErr = Err.Number
    availDesc = Err.Description
    On Error GoTo Fatal

    If availErr <> 0 Then
        LogMsg "ListSAFETables: GetAvailableTables (comparison count only) raised - " & _
               ErrorText(availErr, availDesc)
    ElseIf availRet <> 0 Then
        LogMsg "ListSAFETables: GetAvailableTables (comparison count only) returned " & availRet
    Else
        If NumberTables = nAvail Then
            countNote = " (same count)"
        Else
            countNote = " (DIFFERENT count - that is the difference this listing exists to show)"
        End If
        LogMsg "ListSAFETables: GetAllTables reports " & NumberTables & " table(s); " & _
               "GetAvailableTables reports " & nAvail & " table(s) available for display" & _
               countNote & ". The sheet lists the GetAllTables set."
    End If

    ' Header row
    ws.Cells(r, c).Value = "Table Key"
    ws.Cells(r, c + 1).Value = "Table Name"
    ws.Cells(r, c + 2).Value = "Import Type"
    ws.Range(ws.Cells(r, c), ws.Cells(r, c + 2)).Font.Bold = True
    r = r + 1

    Dim i As Long
    Dim n As Long
    n = ArrLenStr(TableKey)
    For i = 0 To n - 1
        ws.Cells(r + i, c).Value = TableKey(i)
        If i < ArrLenStr(TableName) Then ws.Cells(r + i, c + 1).Value = TableName(i)
        If i < ArrLenLng(ImportType) Then ws.Cells(r + i, c + 2).Value = ImportType(i)
    Next i

    ws.Columns(c).AutoFit
    Application.ScreenUpdating = savedScreen
    ListSAFETables = n
    Exit Function

Fatal:
    Application.ScreenUpdating = savedScreen
    LogError "ListSAFETables"
    ListSAFETables = -1
End Function

' ===========================================================================
' EDITING-TABLE READ - PUBLIC bridge for companion modules
' ===========================================================================
' The display reader (SAFETableToArray, above) is private and shaped for
' EXPORTING: headers plus an array to write into a sheet. An EDIT needs the other
' shape - every column of the table, in SAFE's order, ready to hand back to
' SetTableForEditingArray - which is what this function returns.
'
' Contract
'   TableKey : a table SAFE reports as interactively editable, e.g.
'              "Point Object Connectivity"
'   Headers(): column keys in SAFE's order (0-based, as SAFE reports them).
'              Locate the column you want BY NAME from this array - a hardcoded
'              index is what breaks when SAFE reorders or renames a column.
'   Data     : 1-based 2-D Variant array [row, col], WITHOUT the header row -
'              exactly the form WriteSAFETable takes back. Empty when the table
'              has no rows, which is NOT an error.
'   Returns  : True when the read succeeded (an empty table included), False when
'              the request could not be served - an unknown key, a table that is
'              not interactively editable, or a COM failure. The reason is in the
'              log (ShowLog / GetLog / LastErrorText).
' READ-ONLY: the model's lock state is never touched here; only the write-back
' path unlocks, and only for a table that needs it.
Public Function SAFEReadEditingTable( _
    ByVal TableKey As String, _
    ByRef Headers() As String, _
    ByRef Data As Variant) As Boolean

    Dim ret As Long
    Dim TableVersion As Long
    Dim FieldsKeysIncluded() As String
    Dim NumberRecords As Long
    Dim TableData() As String
    Dim GroupName As String
    Dim nCols As Long
    Dim nRows As Long
    Dim i As Long, j As Long, k As Long
    Dim dataLen As Long
    Dim out() As Variant
    Dim errNum As Long, errDesc As String
    Dim extra As String

    On Error GoTo ErrHandler

    Erase Headers
    Data = Empty
    SAFEReadEditingTable = False

    If Not gConnected Then
        If Not SAFEConnect() Then Exit Function
    End If

    ' GroupName is documented as NOT ACTIVE for the editing tables in this release
    ' ("IMPORTANT NOTE: This parameter is not active in this release"), so an
    ' editing read is whole-table by definition - always pass it blank.
    GroupName = ""

    ret = gDB.GetTableForEditingArray( _
            TableKey, GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData)

    If ret <> 0 Then
        ' Unlike the DISPLAY read there is no documented "nonzero but here is the
        ' data anyway" case for the editing read, so a nonzero code is a failure.
        extra = ""
        If ArrLenStr(FieldsKeysIncluded) > 0 Then
            extra = " (it did report " & ArrLenStr(FieldsKeysIncluded) & " column(s))"
        End If
        SetLastError "SAFEReadEditingTable", "GetTableForEditingArray('" & TableKey & _
            "') returned " & ret & " - the table could not be read for editing" & extra & ".", _
            "Hint: run ListSAFETables (or DemoListTables) and check this table's " & _
            "Import Type - 0 = not importable, 1 = importable but NOT interactively " & _
            "editable through the editing-table pair."
        Exit Function
    End If

    nCols = ArrLenStr(FieldsKeysIncluded)
    If nCols <= 0 Then
        SetLastError "SAFEReadEditingTable", "GetTableForEditingArray('" & TableKey & _
            "') returned no column keys, so the table could not be rebuilt.", _
            "Hint: check the key against ListSAFETables (or DemoListTables)."
        Exit Function
    End If

    ' Column keys, in SAFE's order - the caller maps names to columns.
    ReDim Headers(0 To nCols - 1)
    For j = 0 To nCols - 1
        Headers(j) = FieldsKeysIncluded(j)
    Next j

    nRows = NumberRecords
    If nRows < 1 Then
        ' Valid key, no rows: NOT a failure. Data stays Empty.
        LogMsg "SAFEReadEditingTable: '" & TableKey & "' read - 0 row(s) x " & _
               nCols & " column(s): " & Join(Headers, ", ")
        SAFEReadEditingTable = True
        Exit Function
    End If

    ' Rebuild the flattened row-by-row array. Guarded the same way as the display
    ' read: TableData can come back shorter than rows x columns, and the missing
    ' cells are filled blank rather than raising mid-rebuild.
    ReDim out(1 To nRows, 1 To nCols)
    dataLen = ArrLenStr(TableData)
    k = 0
    For i = 1 To nRows
        For j = 1 To nCols
            If k < dataLen Then
                out(i, j) = TableData(k)
            Else
                out(i, j) = ""
            End If
            k = k + 1
        Next j
    Next i

    Data = out
    LogMsg "SAFEReadEditingTable: '" & TableKey & "' read - " & nRows & " row(s) x " & _
           nCols & " column(s): " & Join(Headers, ", ") & " (table version " & _
           TableVersion & ")"
    SAFEReadEditingTable = True
    Exit Function

ErrHandler:
    ' Capture the Err details in locals FIRST: Error$(...) inside ErrorText would
    ' overwrite the Err object before it has been reported.
    errNum = Err.Number
    errDesc = Err.Description
    SetLastError "SAFEReadEditingTable", "reading '" & TableKey & "' raised - " & _
        ErrorText(errNum, errDesc), ErrHint(errNum)
End Function

' ===========================================================================
' BONUS - write a 2-D array back INTO a SAFE editing table and apply it
'         (mirrors the reference Python script's set_table/apply_table_edit)
' ===========================================================================

Public Function WriteSAFETable( _
    ByVal TableKey As String, _
    ByVal Data As Variant, _
    Optional ByVal UnlockModel As Boolean = True) As Boolean
    ' Data: a 1-based 2-D array [row, col] WITHOUT the header row.
    '       You can pass a Range.Value array directly, e.g.
    '       WriteSAFETable "Point Object Connectivity", ws.Range("A2:D9").Value
    ' IMPORTANT: number of columns must match the table's columns.
    '            Save the SAFE model BEFORE calling (ApplyEditedTables can
    '            corrupt the model on a fatal error).
    ' LOCKING    : the model is unlocked ONLY when SAFE reports this table as
    '              interactively importable while the model is unlocked
    '              (GetAllTables ImportType = 2). ImportType = 3 needs no
    '              unlock at all, and ImportType 0 / 1 cannot be written back
    '              through the editing-table pair - those are refused BEFORE the
    '              lock is touched. Whatever lock state was in force when the
    '              call started is RESTORED on the way out, on the success path
    '              and on every failure path, so your SAFE session is left
    '              exactly as it was found.
    '              UnlockModel is now only a FALLBACK, used when the table's
    '              import type cannot be preflighted (the key is not in SAFE's
    '              table list, or that list could not be read): True = unlock
    '              the model anyway, False = leave the model alone.
    ' EDITING    : an applied edit makes any existing analysis results STALE.
    '              Re-run the analysis in SAFE afterwards - this module
    '              deliberately does NOT run the analysis (or save) for you.

    On Error GoTo Fail

    If Not gConnected Then
        If Not SAFEConnect() Then GoTo Fail
    End If

    ' -----------------------------------------------------------------------
    ' FIX 5 - LOCK HANDLING. Preflight the table's importability BEFORE the
    ' model's lock state is touched, so a table that cannot be written back is
    ' rejected without changing anything at all in the user's SAFE session.
    ' cDatabaseTables.GetAllTables reports an ImportType per table:
    '   0 = not importable
    '   1 = importable but NOT interactively importable
    '   2 = importable and interactively importable when the model is UNLOCKED
    '   3 = importable and interactively importable when the model is unlocked
    '       AND when it is locked
    ' The editing-table pair used below (Get/SetTableForEditingArray) is
    ' documented as needing an INTERACTIVELY editable table, so 0 and 1 can
    ' never work through this path, 2 needs the model unlocked for the write,
    ' and 3 needs no unlock at all.
    ' -----------------------------------------------------------------------
    Dim ret As Long
    Dim wasLocked As Boolean
    Dim unlockedByUs As Boolean
    Dim needUnlock As Boolean
    Dim preflighted As Boolean
    Dim importType As Long
    Dim NumberTables As Long
    Dim TableKeyList() As String
    Dim TableNameList() As String
    Dim ImportTypeList() As Long
    Dim IsEmptyList() As Boolean        ' short name used by the API (not VBA's IsEmpty)
    Dim gErrNum As Long, gErrDesc As String
    Dim keyFound As Boolean
    Dim t As Long, n As Long
    Dim probe As String
    Dim whyUnlock As String
    Dim lockWord As String

    unlockedByUs = False
    needUnlock = False
    preflighted = False

    On Error Resume Next
    Err.Clear
    ret = gDB.GetAllTables(NumberTables, TableKeyList, TableNameList, ImportTypeList, IsEmptyList)
    gErrNum = Err.Number
    gErrDesc = Err.Description
    On Error GoTo Fail

    If gErrNum <> 0 Then
        LogMsg "WARNING: WriteSAFETable: GetAllTables raised - " & _
               ErrorText(gErrNum, gErrDesc) & vbCrLf & _
               "  the import type of '" & TableKey & "' cannot be preflighted, so the " & _
               "UnlockModel argument decides the lock state (UnlockModel=" & UnlockModel & ")."
        needUnlock = UnlockModel
    ElseIf ret <> 0 Then
        LogMsg "WARNING: WriteSAFETable: GetAllTables returned " & ret & _
               " (nonzero) - the import type of '" & TableKey & "' cannot be preflighted, " & _
               "so the UnlockModel argument decides the lock state (UnlockModel=" & _
               UnlockModel & ")."
        needUnlock = UnlockModel
    Else
        ' Find the requested key: Trim'd and case-insensitive first, then - only
        ' if that finds nothing - with all spaces removed, so a minor
        ' key-formatting difference is not mistaken for "table not found".
        probe = Trim$(TableKey)
        n = ArrLenStr(TableKeyList)
        If NumberTables > 0 And NumberTables < n Then n = NumberTables
        For t = 0 To n - 1
            If StrComp(Trim$(TableKeyList(t)), probe, vbTextCompare) = 0 Then
                keyFound = True
                Exit For
            End If
        Next t
        If Not keyFound And Len(Replace(probe, " ", "")) > 0 Then
            Dim probeNoSpaces As String
            probeNoSpaces = Replace(probe, " ", "")
            For t = 0 To n - 1
                If StrComp(Replace(TableKeyList(t), " ", ""), probeNoSpaces, vbTextCompare) = 0 Then
                    keyFound = True
                    Exit For
                End If
            Next t
        End If

        If Not keyFound Then
            LogMsg "WARNING: WriteSAFETable: '" & TableKey & "' is not in SAFE's table " & _
                   "list (" & n & " key(s) reported), so its import type cannot be " & _
                   "preflighted - the UnlockModel argument decides the lock state " & _
                   "(UnlockModel=" & UnlockModel & "). Run ListSAFETables for the exact keys."
            needUnlock = UnlockModel
        Else
            ' ImportTypeList may hold fewer entries than the key list (SAFE only
            ' returns what it has), so this read is guarded.
            gErrNum = 0
            On Error Resume Next
            Err.Clear
            importType = ImportTypeList(t)
            gErrNum = Err.Number
            gErrDesc = Err.Description
            On Error GoTo Fail

            If gErrNum <> 0 Then
                LogMsg "WARNING: WriteSAFETable: GetAllTables reported no import type " & _
                       "for '" & TableKey & "' - the UnlockModel argument decides the " & _
                       "lock state (UnlockModel=" & UnlockModel & ")."
                needUnlock = UnlockModel
            Else
                preflighted = True
                Select Case importType
                    Case 0
                        ' Not importable at all: never touch the lock for this.
                        SetLastError "WriteSAFETable", "the table '" & TableKey & _
                            "' is NOT importable (GetAllTables ImportType = 0), so it " & _
                            "cannot be written back into this model. The model's lock " & _
                            "state was left untouched. Run ListSAFETables (or " & _
                            "DemoListTables) to see the keys this model does accept."
                        GoTo Fail
                    Case 1
                        ' Importable, but not through the INTERACTIVE get/set pair.
                        SetLastError "WriteSAFETable", "the table '" & TableKey & _
                            "' is importable but NOT interactively importable " & _
                            "(GetAllTables ImportType = 1). This routine edits through " & _
                            "the interactive pair Get/SetTableForEditingArray, which SAFE " & _
                            "documents as requiring an interactively editable table, so " & _
                            "the write-back is refused and the model's lock state was " & _
                            "left untouched. Run ListSAFETables (or DemoListTables) to " & _
                            "see the keys this model does accept."
                        GoTo Fail
                    Case 2
                        ' Needs the model unlocked for the duration of the write.
                        If Not UnlockModel Then
                            SetLastError "WriteSAFETable", "the table '" & TableKey & _
                                "' can only be edited interactively while the model is " & _
                                "UNLOCKED (GetAllTables ImportType = 2), but " & _
                                "UnlockModel:=False was passed, so this call will not " & _
                                "unlock the model and the edit cannot be applied as " & _
                                "requested. Pass UnlockModel:=True (the default): if the " & _
                                "model is already unlocked, the call then finds it " & _
                                "unlocked and changes nothing."
                            GoTo Fail
                        End If
                        needUnlock = True
                    Case 3
                        ' Importable while locked OR unlocked - never unlock for these.
                        needUnlock = False
                        LogMsg "WriteSAFETable: '" & TableKey & "' is importable with the " & _
                               "model locked or unlocked (GetAllTables ImportType = 3) - " & _
                               "no unlock is needed and the lock state is left alone."
                    Case Else
                        LogMsg "WARNING: WriteSAFETable: SAFE reported an undocumented " & _
                               "import type (" & importType & ") for '" & TableKey & "' - " & _
                               "the UnlockModel argument decides the lock state " & _
                               "(UnlockModel=" & UnlockModel & ")."
                        needUnlock = UnlockModel
                End Select
            End If
        End If
    End If

    ' Capture the lock state ONCE (this is only a read - it changes nothing). If
    ' it cannot be read we ASSUME the model is locked, which is SAFE's normal
    ' state while its API is in use (see the Readme troubleshooting note), so a
    ' model this call unlocks is always put back the way it was.
    wasLocked = True
    On Error Resume Next
    Err.Clear
    wasLocked = gSapModel.GetModelIsLocked()
    gErrNum = Err.Number
    gErrDesc = Err.Description
    On Error GoTo Fail
    If gErrNum <> 0 Then
        LogMsg "WARNING: WriteSAFETable: GetModelIsLocked raised - " & _
               ErrorText(gErrNum, gErrDesc) & vbCrLf & _
               "  assuming the model was LOCKED (it is re-locked on the way out if " & _
               "this call unlocks it)."
    End If
    If wasLocked Then lockWord = "locked" Else lockWord = "unlocked"

    ' Unlock ONLY when the preflight requires it (ImportType = 2), or when the
    ' fallback path applies and UnlockModel = True - and only when the model is
    ' actually locked.
    If needUnlock Then
        If preflighted Then
            whyUnlock = "SAFE reports it as interactively importable with the model " & _
                        "unlocked (GetAllTables ImportType = 2)"
        Else
            whyUnlock = "its import type could not be preflighted and UnlockModel:=True " & _
                        "was requested"
        End If

        If wasLocked Then
            ret = gSapModel.SetModelIsLocked(False)
            If ret <> 0 Then
                ' Not fatal on its own: if the model is in fact still locked, SAFE
                ' rejects the edit and that is reported where it happens.
                LogMsg "WARNING: WriteSAFETable: SetModelIsLocked(False) returned " & ret & _
                       " - the model may still be locked, in which case the write-back " & _
                       "will be rejected by SAFE."
            Else
                unlockedByUs = True
                LogMsg "WriteSAFETable: the model was LOCKED - unlocked it for this " & _
                       "write-back (" & whyUnlock & "). It is re-locked on the way out."
            End If
        Else
            LogMsg "WriteSAFETable: the model is already UNLOCKED - no lock change " & _
                   "needed (" & whyUnlock & ")."
        End If
    ElseIf preflighted Then
        LogMsg "WriteSAFETable: no unlock needed (GetAllTables ImportType = " & importType & _
               ") - the model's lock state (" & lockWord & ") is left as it is."
    Else
        LogMsg "WriteSAFETable: no unlock requested and no import type to preflight " & _
               "(UnlockModel=" & UnlockModel & ") - the model's lock state (" & _
               lockWord & ") is left as it is."
    End If

    ' Pull current table structure (headers + version + existing rows).
    ' TableVersion is an OUT item here (the version of the table as it is now) -
    ' see the QUIRK note at the SetTableForEditingArray call below before using it.
    Dim TableVersion As Long
    Dim FieldsKeysIncluded() As String
    Dim NumberRecords As Long
    Dim TableData() As String
    Dim GroupName As String
    GroupName = ""     ' quirk: GroupName is NOT active in this SAFE release

    ret = gDB.GetTableForEditingArray( _
            TableKey, GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData)
    If ret <> 0 Then
        LogMsg "WriteSAFETable: GetTableForEditingArray('" & TableKey & "') returned " & ret
        GoTo Fail
    End If

    Dim nCols As Long
    nCols = ArrLenStr(FieldsKeysIncluded)

    Dim nRows As Long
    Dim dataCols As Long
    On Error Resume Next
    nRows = UBound(Data, 1)
    dataCols = UBound(Data, 2)
    ' Restore the Fail handler, NOT "On Error GoTo 0": the lock-restore block
    ' at the bottom of this function is only reachable through "Fail:", and a
    ' model this call unlocked must be put back even when one of the SAFE calls
    ' below raises. Disabling trapping here would let a raised error leave the
    ' user's model UNLOCKED.
    On Error GoTo Fail
    If nRows < 1 Then nRows = 0

    ' Quirk: column count must match, otherwise SAFE rejects the write.
    If dataCols <> nCols Then
        LogMsg "WriteSAFETable: column mismatch - SAFE table has " & nCols & _
               " columns but data has " & dataCols & "."
        GoTo Fail
    End If

    ' Flatten the data row-by-row (the API only accepts a 1-D String array).
    Dim flat() As String
    ReDim flat(0 To nRows * nCols - 1)
    Dim i As Long, j As Long, k As Long
    k = 0
    For i = 1 To nRows
        For j = 1 To nCols
            flat(k) = CStr(Data(i, j))
            k = k + 1
        Next j
    Next i

    ' QUIRK: SetTableForEditingArray's TableVersion is a RETURNED (out) item -
    ' CSI documents it as "Returned Item: The version number of the specified
    ' table" - so the version that GetTableForEditingArray filled in above must
    ' NOT be fed back in. The reference implementation passes 0 here and ignores
    ' what comes back; this module does the same, and keeps the version SAFE
    ' reports in outTableVersion so a mismatch shows up in the log instead of
    ' silently rewriting a table the model has moved on from.
    Dim outTableVersion As Long
    outTableVersion = 0                    ' 0 = let SAFE use the current version
    ret = gDB.SetTableForEditingArray( _
            TableKey, outTableVersion, FieldsKeysIncluded, nRows, flat)
    If ret <> 0 Then
        LogMsg "WriteSAFETable: SetTableForEditingArray('" & TableKey & "') returned " & ret
        GoTo Fail
    End If
    ' Informational only: a different version is NOT a failure of the call.
    If outTableVersion <> TableVersion Then
        LogMsg "WriteSAFETable: SAFE reported table version " & outTableVersion & _
               " on write-back; GetTableForEditingArray had reported " & TableVersion & _
               " (informational - the edit was still accepted)."
    End If

    ' Apply all edited tables. Check the error counts - a nonzero return or
    ' fatal errors can leave the model in a bad state.
    ' QUIRK: FillImportLog must be True for ImportLog to be filled at all - with
    ' False the counts still come back but the log string is EMPTY, which is why
    ' a failure used to be reported with no detail. SAFE warns the log may be
    ' very large, so only the first IMPORT_LOG_MAX characters are logged here.
    Dim NumFatalErrors As Long, NumErrorMsgs As Long
    Dim NumWarnMsgs As Long, NumInfoMsgs As Long
    Dim ImportLog As String
    ret = gDB.ApplyEditedTables(True, NumFatalErrors, NumErrorMsgs, NumWarnMsgs, NumInfoMsgs, ImportLog)

    ' TRUE FAILURE = a nonzero return OR fatal errors. Either can leave the model
    ' in a bad state, so the edit buffer is cancelled and the call reports False.
    If ret <> 0 Or NumFatalErrors > 0 Then
        ' The four counts are logged FIRST and unconditionally - they are the
        ' evidence that the write-back went wrong even if the log is empty.
        LogMsg "WriteSAFETable: ApplyEditedTables returned " & ret & _
               " (fatal=" & NumFatalErrors & ", errors=" & NumErrorMsgs & _
               ", warnings=" & NumWarnMsgs & ", info=" & NumInfoMsgs & ")"
        LogImportLog ImportLog, NumErrorMsgs + NumWarnMsgs + NumInfoMsgs

        ' Preserve the corruption warning: the caller may now be holding a bad
        ' model, and SAFE recommends saving BEFORE this call for exactly this case.
        LogMsg "  The model may be in a corrupted state - if so, close it WITHOUT " & _
               "saving and reopen it (the saved copy is the safe fallback)."

        gDB.CancelTableEditing     ' clear the pending edit buffer
        GoTo Fail
    End If

    ' NOT a failure: messages (error / warning / info) on their own do NOT mean
    ' the write-back failed - ApplyEditedTables reports them for edits SAFE went
    ' on to accept. They are still logged, as a WARNING, with all four counts and
    ' the import log, but the call continues on the SUCCESS path and returns True.
    If NumErrorMsgs > 0 Or NumWarnMsgs > 0 Or NumInfoMsgs > 0 Then
        LogMsg "WARNING: WriteSAFETable: ApplyEditedTables returned " & ret & _
               " (fatal=" & NumFatalErrors & ", errors=" & NumErrorMsgs & _
               ", warnings=" & NumWarnMsgs & ", info=" & NumInfoMsgs & _
               ") - no fatal error, so the edit is treated as APPLIED."
        LogImportLog ImportLog, NumErrorMsgs + NumWarnMsgs + NumInfoMsgs
    End If

    ' Clear the internal edit buffer.
    gDB.CancelTableEditing
    WriteSAFETable = True
    GoTo RestoreLock            ' both exits share the lock-restore block below

Fail:
    ' Record a genuine COM failure with its number. Deliberate validation
    ' failures (bad table key, column mismatch, rejected edit, a table that
    ' cannot be written back) reach this label with Err.Number = 0 - they were
    ' already logged where they were detected.
    LogError "WriteSAFETable"
    WriteSAFETable = False
    ' (falls through: the lock-restore block runs on this path too)

RestoreLock:
    ' -----------------------------------------------------------------------
    ' FIX 5 - put the model's lock state back the way this call found it. This
    ' is the ONLY place the restore happens, and it is reached from the SUCCESS
    ' path AND from every failure path (the Fail label above falls into it), so
    ' the module can never leave the user's SAFE session unlocked because of a
    ' write-back. Clearing the flag as the block runs makes it idempotent (at
    ' most one restore per call), and the SetModelIsLocked call is wrapped in a
    ' minimal On Error Resume Next / On Error GoTo 0 pair so a failure to
    ' re-lock is reported but can never mask the error that brought us here.
    ' -----------------------------------------------------------------------
    If unlockedByUs Then
        unlockedByUs = False            ' idempotent: at most one restore per call
        On Error Resume Next
        ret = gSapModel.SetModelIsLocked(True)
        gErrNum = Err.Number
        gErrDesc = Err.Description
        On Error GoTo 0

        If gErrNum <> 0 Then
            LogMsg "WriteSAFETable: could not re-lock the model, which was LOCKED when " & _
                   "this call started (this call unlocked it) - " & _
                   ErrorText(gErrNum, gErrDesc)
        ElseIf ret <> 0 Then
            LogMsg "WriteSAFETable: SetModelIsLocked(True) returned " & ret & _
                   " while restoring the lock state this call found - the model may " & _
                   "still be unlocked."
        Else
            LogMsg "WriteSAFETable: model re-locked - it was LOCKED when this call " & _
                   "started and the write-back required it unlocked."
        End If
    End If
End Function

' Copy SAFE's ImportLog (ApplyEditedTables) into this module's log. Shared by the
' failure path and the warning path so both report the same evidence; the counts
' line is always written just BEFORE this is called, which is what the "counted
' above" wording refers to. SAFE documents the log as possibly very large, so
' only the first IMPORT_LOG_MAX characters are copied, with an explicit marker;
' NumMsgs > 0 with an empty log means FillImportLog was not honoured (see the
' quirk notes), which is worth saying out loud rather than losing the detail.
Private Sub LogImportLog(ByVal ImportLog As String, ByVal NumMsgs As Long)
    If Len(ImportLog) > 0 Then
        ' Truncate: the import log may be very large by design.
        If Len(ImportLog) > IMPORT_LOG_MAX Then
            LogMsg "ImportLog (first " & IMPORT_LOG_MAX & " of " & Len(ImportLog) & _
                   " characters; the full log is available in SAFE):" & vbCrLf & _
                   Left$(ImportLog, IMPORT_LOG_MAX) & " ... [truncated]"
        Else
            LogMsg "ImportLog (" & Len(ImportLog) & " characters):" & _
                   vbCrLf & ImportLog
        End If
    ElseIf NumMsgs > 0 Then
        ' FillImportLog=True was requested, yet nothing came back - note it so
        ' a future API change is visible instead of silently losing detail.
        LogMsg "ImportLog: requested (FillImportLog=True) but SAFE returned an " & _
               "EMPTY log even though it reported the messages counted above."
    End If
End Sub

' ===========================================================================
' CORE READ - pull one table into a 2-D array (quirk-tolerant)
' ===========================================================================

Public Function SAFETableToArray( _
    ByVal TableKey As String, _
    ByRef Headers() As String, _
    ByRef Warning As String, _
    ByRef Failed As Boolean, _
    Optional ByVal ReturnHeaders As Boolean = False) As Variant
    ' Returns a 1-based 2-D Variant array [row, col] of data (headers excluded),
    ' or Empty when there is nothing to write.
    '
    ' ReturnHeaders controls ONLY the Headers() out-array, and it is NOT the same
    ' question as "did SAFE report any columns" - that one is answered internally
    ' from FieldsKeysIncluded and still decides failure vs empty below:
    '   False (DEFAULT): Headers() is left EMPTY. Use this when the caller does
    '       not intend to write column headers at all (a header-less export); the
    '       data array still carries EVERY column, so a writer can take the column
    '       count from UBound(Data, 2) instead.
    '   True : Headers() is filled with SAFE's column keys, in SAFE's order.
    ' ExportSAFETables passes its own IncludeHeader through here, so the library
    ' behaves exactly as it always has when headers are wanted.
    '
    ' STATUS CONTRACT - Failed is False only for the non-failure cases, and the
    ' caller MUST branch on it (a failed read is never an empty table):
    '   Failed = True, Warning = text
    '       * a COM exception was raised: logged HERE with its error number and
    '         hint (the existing handler at the bottom), Empty returned; or
    '       * the API returned a NONZERO code AND no column headers came back at
    '         all (ArrLenStr(FieldsKeysIncluded) <= 0). CSI documents "if there
    '         is nothing to be shown in the table then no data is returned", so
    '         a table that exists still reports its columns: no headers means
    '         the request could not be served - the key is not valid for this
    '         model / SAFE version, or the table is a result table whose
    '         analysis has not been run. Logged here as an ERROR, Empty returned.
    '   Failed = False
    '       * nonzero code WITH headers: the data was still returned, so the
    '         existing warning text is kept (nonzero code, data still returned);
    '       * zero code WITH headers but NumberRecords <= 0: the legitimate
    '         EMPTY-table case, Warning stays "";
    '       * zero code with no headers: Warning says the table returned no
    '         columns (not the invalid-key failure above).
    On Error GoTo ErrHandler

    Erase Headers
    Warning = ""
    Failed = False

    Dim ret As Long
    Dim FieldKeyList() As String
    Dim GroupName As String
    Dim TableVersion As Long
    Dim FieldsKeysIncluded() As String
    Dim NumberRecords As Long
    Dim TableData() As String

    GroupName = ""                              ' "" / "All" = all objects
    ReDim FieldKeyList(0)
    FieldKeyList(0) = ""                        ' single blank = ALL columns

    ret = gDB.GetTableForDisplayArray( _
            TableKey, FieldKeyList, GroupName, TableVersion, _
            FieldsKeysIncluded, NumberRecords, TableData)

    ' QUIRK / FIX 4: a nonzero return means an error OR "nothing to show" -
    ' the two are told apart by whether the COLUMN HEADERS came back. A table
    ' that exists always reports its columns, so "nonzero AND no headers"
    ' cannot be an empty table: it is a failed request (invalid key for this
    ' model / SAFE version, or analysis not run). Nonzero WITH headers is only
    ' a warning - the data was still returned.
    Dim nCols As Long
    nCols = ArrLenStr(FieldsKeysIncluded)

    If ret <> 0 Then
        If nCols <= 0 Then
            Failed = True
            Warning = "API code " & ret & " and NO column headers were " & _
                      "returned - the table key is probably not valid for this " & _
                      "model / SAFE version, or the table has no data because " & _
                      "the analysis has not been run."
            LogMsg "ERROR [" & TableKey & "] read FAILED - " & Warning
            SAFETableToArray = Empty
            Exit Function
        End If
        Warning = "API code " & ret & " (nonzero, but data was still returned)"
    End If

    If nCols <= 0 Then
        ' Zero return and no columns either: nothing to write, but this is NOT
        ' the invalid-key / analysis-not-run failure handled above.
        Warning = "the table returned no columns"
        SAFETableToArray = Empty
        Exit Function
    End If

    ' Copy the column headers - but ONLY when the caller asked for them. The
    ' column count used from here on comes from FieldsKeysIncluded, NOT from
    ' Headers, so suppressing the copy does not touch the data array at all: every
    ' column is still returned.
    Dim j As Long
    If ReturnHeaders Then
        ReDim Headers(0 To nCols - 1)
        For j = 0 To nCols - 1
            Headers(j) = FieldsKeysIncluded(j)
        Next j
    End If

    Dim nRows As Long
    nRows = NumberRecords
    If nRows <= 0 Then
        ' Zero return WITH headers but no records: the legitimate "empty table"
        ' case. Failed stays False and Warning stays empty, so the caller uses
        ' its ordinary "no data returned" marker.
        SAFETableToArray = Empty
        Exit Function
    End If

    ' Rebuild the flattened, row-by-row data into a 2-D array.
    Dim out() As Variant
    ReDim out(1 To nRows, 1 To nCols)
    Dim i As Long, k As Long
    Dim dataLen As Long
    dataLen = ArrLenStr(TableData)
    k = 0
    For i = 1 To nRows
        For j = 1 To nCols
            If k < dataLen Then
                out(i, j) = TableData(k)
            Else
                out(i, j) = ""                  ' guard: shorter than expected
            End If
            k = k + 1
        Next j
    Next i

    SAFETableToArray = out
    Exit Function

ErrHandler:
    ' Capture the error details in locals before calling any helper, so the
    ' report cannot be invalidated by the Err object being overwritten.
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Failed = True
    Warning = ErrorText(errNum, errDesc)
    LogMsg "[" & TableKey & "] read failed - " & Warning
    If Len(ErrHint(errNum)) > 0 Then LogMsg "  " & ErrHint(errNum)
    SAFETableToArray = Empty
End Function

' ===========================================================================
' WRITE A TABLE BLOCK TO THE WORKSHEET (title row + header row, then data)
' ===========================================================================

Public Function WriteTableBlock( _
    ByVal Wb As Workbook, _
    ByRef CurSheet As Worksheet, _
    ByVal BaseSheetName As String, _
    ByVal r0 As Long, ByVal c0 As Long, _
    ByVal Title As String, _
    ByRef Headers() As String, _
    ByVal Data As Variant, _
    ByVal WithTitle As Boolean, _
    ByVal WithHeader As Boolean, _
    ByVal Horiz As Boolean, _
    ByRef NextRow As Long, ByRef NextCol As Long) As Long
    ' Writes the title row and the header row - each one only when WithTitle /
    ' WithHeader asks for it - followed by the data, starting at (r0, c0), and
    ' advances the cursor (NextRow/NextCol) so the caller can place the next table.
    ' Returns the number of DATA ROWS written - summed over every worksheet used -
    ' or -1 when the block could NOT be written in full (a table too wide for a
    ' worksheet, no room left for its title/header rows, or a spill that ran out
    ' of room on a continuation sheet). -1 is a FAILURE: the caller counts it with
    ' the failed tables and must NOT write a second marker, because this routine
    ' has already written a bold red one.
    '
    ' Wb            : workbook the continuation sheets are created in (passed to
    '                 GetWorksheet) - normally ThisWorkbook.
    ' CurSheet      : the sheet to START on. It is updated (ByRef) to the sheet
    '                 the cursor ENDS UP on, so the CALLER must write anything
    '                 that follows (the empty-table marker, the red FAILED
    '                 marker, the next table) through THIS variable: that is the
    '                 starting sheet unless the table spilled, in which case
    '                 writing through a reference captured before the call would
    '                 put the marker on the wrong sheet.
    ' BaseSheetName : the caller's sheet name - the name the continuation sheets
    '                 are derived from (see the spill note below).
    ' WithTitle     : write the title row (the table key) as the block's first
    '                 row. ExportSAFETables drives this AND WithHeader from its
    '                 single IncludeHeader switch, so IncludeHeader:=False puts
    '                 the DATA flush on (r0, c0): no title row, no header row.
    ' WithHeader    : write the column-header row below the title. Needs the
    '                 caller to have filled Headers() (SAFETableToArray's
    '                 ReturnHeaders); with no keys supplied the row is skipped
    '                 rather than written blank.
    '
    ' SPILL BEHAVIOUR (a table larger than one worksheet)
    '   Excel's limits are READ AT RUNTIME from the sheet (Rows.Count /
    '   Columns.Count) and never hardcoded, because they depend on the workbook
    '   format: 1,048,576 x 16,384 for .xlsx/.xlsm, 65,536 x 256 for .xls.
    '     * a table that FITS is written with ONE bulk Range.Value2 write of the
    '       whole Data array, exactly as before (no slicing at all in this path -
    '       that is what keeps a normal export fast);
    '     * a table that does NOT fit is continued on additional worksheets named
    '       <BaseSheetName>_2, <BaseSheetName>_3, ... created on demand by
    '       GetWorksheet. Excel allows at most 31 characters in a sheet name, so
    '       the base name is cut short when needed and the "_N" suffix is kept
    '       intact (truncating the whole combination could collapse the name back
    '       onto the base sheet and overwrite rows already written there). Every
    '       sheet - the first one included - carries the title AND the header
    '       row (whichever of the two is being written), so each sheet is
    '       self-contained; each sheet's rows are sliced
    '       into a fresh 1-based 2-D Variant chunk and bulk-written, so no row is
    '       skipped or duplicated. ONE warning is logged before the first
    '       continuation sheet, each continuation sheet is logged as it is used,
    '       and the total rows written is logged when the spill finishes.
    '   Two situations are REFUSED instead of being truncated silently - an ERROR
    '   is logged, a bold red marker cell is written at the cursor, the cursor is
    '   stepped past the marker, and -1 is returned (so the caller reports the
    '   table as FAILED, exactly like a failed read):
    '     * the table is WIDER than the sheet (c0 + nCols - 1 > Columns.Count);
    '     * the start cell leaves no room for the title + header rows, i.e. the
    '       capacity is < 1 (capacity = Rows.Count - (r0 - 1) - overhead rows).
    '   Cursor bookkeeping keeps its original meaning, but refers to the sheet
    '   the cursor lands on: Horiz = True -> NextCol = c0 + nCols + 1 and
    '   NextRow = r0; Horiz = False -> NextRow = the row after the last written
    '   row, plus one blank row, and NextCol = c0.

    Dim nCols As Long
    nCols = ArrLenStr(Headers)
    ' Headers are OPTIONAL (SAFETableToArray's ReturnHeaders, driven by
    ' ExportSAFETables' IncludeHeader): a caller that wants no header row does not
    ' fetch them. The column count still has to come from somewhere, or the block
    ' would be treated as zero-width and NOTHING would be written - so fall back to
    ' the width of the data array itself.
    If nCols <= 0 And Not IsEmpty(Data) Then
        On Error Resume Next
        nCols = UBound(Data, 2)
        On Error GoTo 0
        If nCols < 1 Then nCols = 0
    End If

    Dim nRows As Long
    nRows = 0
    If Not IsEmpty(Data) Then
        On Error Resume Next
        nRows = UBound(Data, 1)
        On Error GoTo 0
        If nRows < 1 Then nRows = 0
    End If

    ' --- runtime worksheet limits of the sheet this block starts on ----------
    Dim maxRow As Long, maxCol As Long
    maxRow = CurSheet.Rows.Count
    maxCol = CurSheet.Columns.Count

    ' Rows of "overhead" the block needs above the data on EVERY sheet it uses:
    ' the title row, plus the header row, each one only when it is being written.
    ' Both are 0 when the caller asked for a bare data block (IncludeHeader:=
    ' False), so the whole sheet below (r0 - 1) is available for data.
    Dim overheadRows As Long
    overheadRows = 0
    If WithTitle And Len(Title) > 0 Then overheadRows = overheadRows + 1
    If WithHeader And nCols > 0 And ArrLenStr(Headers) > 0 Then overheadRows = overheadRows + 1

    ' Data rows that fit on one sheet below (r0 - 1) rows of overhead. Every
    ' sheet used by a spill starts at the SAME top-left cell, so the same amount
    ' of room is available on each of them.
    Dim capacity As Long
    capacity = maxRow - (r0 - 1) - overheadRows

    Dim r As Long, c As Long, j As Long
    r = r0
    c = c0

    ' --- WIDTH CHECK FIRST: a table too wide is never truncated --------------
    If nCols > 0 And c0 + nCols - 1 > maxCol Then
        With CurSheet.Cells(r0, c0)
            .Value = "Table '" & Title & "' : FAILED - too wide for a worksheet (" & _
                     nCols & " column(s) starting at column " & c0 & ")"
            .Font.Bold = True
            .Font.Color = vbRed
        End With
        LogMsg "ERROR in WriteTableBlock: table '" & Title & "' needs " & nCols & _
               " column(s) starting at column " & c0 & ", but sheet '" & CurSheet.Name & _
               "' has only " & maxCol & " column(s) - nothing was written for this table."
        ' Step the cursor past the marker cell so the next table cannot land on
        ' it (Horiz = True moves one column right - the whole table would not fit
        ' to the right of it either - and Horiz = False leaves a blank row).
        If Horiz Then
            If c0 < maxCol Then NextCol = c0 + 1 Else NextCol = c0
            NextRow = r0
        Else
            NextRow = r0 + 2
            NextCol = c0
        End If
        ' DEFECT 2 FIX: -1, not 0. 0 means "wrote no rows" (a legitimate empty
        ' table) and the caller counts such a block as WRITTEN; -1 is the failure
        ' signal that makes the caller count this block with the failed tables
        ' instead of reporting it as a table it wrote.
        WriteTableBlock = -1
        Exit Function
    End If

    ' --- NO ROOM for the title + header rows: refuse rather than corrupt -----
    If capacity < 1 Then
        With CurSheet.Cells(r0, c0)
            .Value = "Table '" & Title & "' : FAILED - no room on this worksheet for " & _
                     overheadRows & " title/header row(s) (start row " & r0 & " of " & _
                     maxRow & ")"
            .Font.Bold = True
            .Font.Color = vbRed
        End With
        LogMsg "ERROR in WriteTableBlock: table '" & Title & "' starts at row " & r0 & _
               ", which leaves no room for its " & overheadRows & " title/header row(s) on " & _
               "sheet '" & CurSheet.Name & "' (" & maxRow & " row(s)) - nothing was written " & _
               "for this table."
        If Horiz Then
            If c0 < maxCol Then NextCol = c0 + 1 Else NextCol = c0
            NextRow = r0
        Else
            NextRow = r0 + 2
            NextCol = c0
        End If
        ' DEFECT 2 FIX: -1 (failure), not 0 ("wrote no rows") - see the width
        ' check above.
        WriteTableBlock = -1
        Exit Function
    End If

    ' --- NORMAL CASE: the whole table fits on this sheet ---------------------
    ' ONE bulk write of the entire Data array - unchanged from before the spill
    ' support, so an ordinary export loses no speed (no row-by-row copying here).
    If nRows <= capacity Then
        ' Title row - skipped when WithTitle is False, which is how a caller
        ' asking for a bare data block gets the data ON the start cell.
        If WithTitle And Len(Title) > 0 Then
            CurSheet.Cells(r, c).Value = Title
            CurSheet.Cells(r, c).Font.Bold = True
            r = r + 1
        End If

        ' Header row - also skipped when no header keys were supplied (asking for
        ' a header row AND suppressing the fetch would otherwise write blanks).
        If WithHeader And nCols > 0 And ArrLenStr(Headers) > 0 Then
            For j = 1 To nCols
                CurSheet.Cells(r, c + j - 1).Value = Headers(j - 1)
                CurSheet.Cells(r, c + j - 1).Font.Bold = True
                CurSheet.Cells(r, c + j - 1).Interior.Color = RGB(221, 235, 247)
            Next j
            r = r + 1
        End If

        ' Data block (bulk write for speed)
        If nRows > 0 And nCols > 0 Then
            CurSheet.Range(CurSheet.Cells(r, c), CurSheet.Cells(r + nRows - 1, c + nCols - 1)).Value2 = Data
            r = r + nRows
        End If

        ' Advance the cursor for the next table.
        If Horiz Then
            NextCol = c + nCols + 1
            NextRow = r0
        Else
            NextRow = r + 1
            NextCol = c
        End If

        WriteTableBlock = nRows
        Exit Function
    End If

    ' --- SPILL: the table does not fit on one worksheet ----------------------
    ' A zero-width data block cannot be sliced row by row (and there is nothing
    ' to write either). SAFETableToArray always reports the table's columns, so
    ' this is a guard against a run-time subscript error mid-spill, not a normal
    ' path; the cursor is left exactly where it was.
    If nCols < 1 Then
        ' Unreachable in practice (SAFETableToArray always reports the table's
        ' columns), but a zero-width block cannot be sliced row by row and there
        ' is nothing to write either. Reported as a FAILURE so rows SAFE returned
        ' are never dropped silently: the caller counts it with the failed tables,
        ' and the red marker below says where it happened.
        With CurSheet.Cells(r0, c0)
            .Value = "Table '" & Title & "' : FAILED - " & nRows & " data row(s) were " & _
                     "returned without any column(s), so they could not be written"
            .Font.Bold = True
            .Font.Color = vbRed
        End With
        LogMsg "ERROR in WriteTableBlock: table '" & Title & "' reported " & nRows & _
               " data row(s) but NO column(s), so no data block could be written."
        If Horiz Then
            If c0 < maxCol Then NextCol = c0 + 1 Else NextCol = c0
            NextRow = r0
        Else
            NextRow = r0 + 2
            NextCol = c0
        End If
        WriteTableBlock = -1
        Exit Function
    End If

    ' ONE warning before the first continuation sheet, so the extra sheets are
    ' never a surprise. This is the row-limit warning the caller relies on.
    LogMsg "WARNING: WriteTableBlock: table '" & Title & "': " & nRows & _
           " data rows exceed the " & maxRow & "-row worksheet limit - " & _
           "continuing on additional worksheets."

    Dim rowsLeft As Long
    Dim rowsThis As Long
    Dim offsetRows As Long
    Dim totalRows As Long
    Dim sheetCount As Long       ' worksheets this table has been written to
    Dim nameIdx As Long          ' index used to name the NEXT continuation sheet
    Dim sheetCapacity As Long
    Dim chunk() As Variant
    Dim ii As Long, jj As Long
    Dim contName As String
    Dim suffix As String
    Dim aborted As Boolean       ' a continuation sheet ran out of room mid-spill

    rowsLeft = nRows
    offsetRows = 0               ' data rows of this table already written
    totalRows = 0
    sheetCount = 0               ' 0 = the sheet the caller passed in is next
    nameIdx = 1                  ' the starting sheet counts as sheet 1

    Do While rowsLeft > 0
        If sheetCount > 0 Then
            ' Continuation sheet: <BaseSheetName>_2, <BaseSheetName>_3, ...
            ' (BaseSheetName is the name of the sheet this block STARTED on, so a
            ' second spilled table - which starts on the last sheet the first one
            ' used - nests off that sheet and cannot reuse a sheet that already
            ' carries rows.)
            ' The "_N" SUFFIX is the part that has to survive, so the BASE name
            ' is cut short when the combination would exceed Excel's
            ' 31-character sheet-name limit: truncating the whole combination
            ' instead could collapse the name back onto a sheet that is already
            ' holding rows. A 31-character base name that already ends in the
            ' same "_N" is the one case that can still rebuild the name of the
            ' sheet being written on, which would overwrite the chunk just
            ' written - so the index is advanced until the name differs (at most
            ' a couple of iterations).
            Do
                nameIdx = nameIdx + 1
                suffix = "_" & nameIdx
                If Len(suffix) >= 31 Then
                    contName = Left$(suffix, 31)      ' guard: cannot be reached
                Else
                    contName = Left$(BaseSheetName, 31 - Len(suffix)) & suffix
                End If
                If StrComp(contName, CurSheet.Name, vbTextCompare) <> 0 Then Exit Do
            Loop
            Set CurSheet = GetWorksheet(Wb, contName)
            LogMsg "WriteTableBlock: table '" & Title & "' continues on sheet '" & _
                   contName & "' (" & rowsLeft & " data row(s) still to write)."
        End If

        ' Re-read the limits: a continuation sheet is normally identical to the
        ' starting one (same start cell, same format), but this keeps the room
        ' calculation honest whatever the sheet turns out to be.
        maxRow = CurSheet.Rows.Count
        maxCol = CurSheet.Columns.Count
        sheetCapacity = maxRow - (r0 - 1) - overheadRows
        If sheetCapacity < 1 Then
            LogMsg "ERROR in WriteTableBlock: sheet '" & CurSheet.Name & "' has no room left " & _
                   "below row " & (r0 - 1) & " of its " & maxRow & " row(s) for the " & _
                   "title/header of table '" & Title & "' - " & rowsLeft & _
                   " data row(s) could not be written."
            ' Mark the sheet the spill died on, so the unwritten rows are visible
            ' in the workbook itself and not only in the log.
            With CurSheet.Cells(r0, c0)
                .Value = "Table '" & Title & "' : FAILED - " & rowsLeft & " data row(s) " & _
                         "could not be written (no room left on sheet '" & _
                         CurSheet.Name & "')"
                .Font.Bold = True
                .Font.Color = vbRed
            End With
            aborted = True
            Exit Do
        End If

        If rowsLeft < sheetCapacity Then
            rowsThis = rowsLeft
        Else
            rowsThis = sheetCapacity
        End If

        ' Slice this sheet's share out of Data into a fresh 1-based 2-D chunk and
        ' bulk-write that. Only the spill path copies rows; offsetRows is the
        ' number of rows already written, so no row is repeated or skipped.
        ReDim chunk(1 To rowsThis, 1 To nCols)
        For ii = 1 To rowsThis
            For jj = 1 To nCols
                chunk(ii, jj) = Data(offsetRows + ii, jj)
            Next jj
        Next ii

        ' Title + header are repeated on EVERY sheet, so each sheet stands alone
        ' (whichever of the two the caller asked for).
        r = r0
        c = c0
        If WithTitle And Len(Title) > 0 Then
            CurSheet.Cells(r, c).Value = Title
            CurSheet.Cells(r, c).Font.Bold = True
            r = r + 1
        End If
        If WithHeader And nCols > 0 And ArrLenStr(Headers) > 0 Then
            For j = 1 To nCols
                CurSheet.Cells(r, c + j - 1).Value = Headers(j - 1)
                CurSheet.Cells(r, c + j - 1).Font.Bold = True
                CurSheet.Cells(r, c + j - 1).Interior.Color = RGB(221, 235, 247)
            Next j
            r = r + 1
        End If
        CurSheet.Range(CurSheet.Cells(r, c), _
                       CurSheet.Cells(r + rowsThis - 1, c + nCols - 1)).Value2 = chunk
        r = r + rowsThis

        offsetRows = offsetRows + rowsThis
        rowsLeft = rowsLeft - rowsThis
        totalRows = totalRows + rowsThis
        sheetCount = sheetCount + 1

        LogMsg "WriteTableBlock: table '" & Title & "' - sheet '" & CurSheet.Name & _
               "' holds data row(s) " & (offsetRows - rowsThis + 1) & " to " & offsetRows & _
               " of " & nRows & "."
    Loop

    ' Cursor bookkeeping - same meaning as the single-sheet path, but referring
    ' to the sheet the cursor actually ended up on (CurSheet).
    If Horiz Then
        NextCol = c0 + nCols + 1
        NextRow = r0
    Else
        NextRow = r + 1
        NextCol = c0
    End If

    If aborted Then
        ' The block was only PARTLY written: report it as a FAILURE (-1) so the
        ' caller counts it with the failed tables and the run reads as "this table
        ' did not make it", never as a successful write of fewer rows.
        LogMsg "WriteTableBlock: table '" & Title & "' is INCOMPLETE - " & totalRows & _
               " of " & nRows & " data row(s) were written across " & sheetCount & _
               " worksheet(s); the remaining " & rowsLeft & " could not be written."
        ' Put the cursor past the red marker written on the sheet the spill died
        ' on, so a following table cannot land on it.
        If Horiz Then
            If c0 < maxCol Then NextCol = c0 + 1 Else NextCol = c0
            NextRow = r0
        Else
            NextRow = r0 + 2
            NextCol = c0
        End If
        WriteTableBlock = -1
        Exit Function
    End If

    LogMsg "WriteTableBlock: table '" & Title & "' was written across " & sheetCount & _
           " worksheet(s) - " & totalRows & " data row(s) in total."
    WriteTableBlock = totalRows
End Function

' ===========================================================================
' SMALL HELPERS
' ===========================================================================

Private Function NormalizeTables(ByVal Tables As Variant, ByRef names() As String) As Long
    ' Accepts a single key (String), a comma-separated String, or a String
    ' array. Fills names() and returns the count.
    On Error GoTo ErrH

    If IsArray(Tables) Then
        Dim i As Long, n As Long
        n = UBound(Tables) - LBound(Tables) + 1
        If n > 0 Then
            ReDim names(0 To n - 1)
            For i = 0 To n - 1
                names(i) = CStr(Tables(LBound(Tables) + i))
            Next i
        End If
        NormalizeTables = n
        Exit Function
    End If

    Dim s As String
    s = Trim(CStr(Tables))
    If Len(s) = 0 Then
        NormalizeTables = 0
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(s, ",")
    n = UBound(parts) - LBound(parts) + 1
    If n > 0 Then
        ReDim names(0 To n - 1)
        For i = 0 To n - 1
            names(i) = Trim(CStr(parts(LBound(parts) + i)))
        Next i
    End If
    NormalizeTables = n
    Exit Function

ErrH:
    LogError "NormalizeTables", "The table-name / load-case argument could not be read."
    NormalizeTables = 0
End Function

' ===========================================================================
' DISPLAY FILTER - load cases / load combinations (save / apply / restore)
' ===========================================================================
' SAFE keeps TWO independent "selected for table display" lists: load CASES and
' load COMBINATIONS (cDatabaseTables.Get/SetLoadCasesSelectedForDisplay and
' Get/SetLoadCombinationsSelectedForDisplay). They decide what the display
' (result) tables return, so the optional LoadCases / LoadCombos arguments of
' ExportSAFETables work by changing them - and by putting the previous
' selections back afterwards, on every path out of the function.
' Documented quirks: a list holding a single blank string selects NOTHING (there
' is no documented "select all"; a previously empty selection is put back that
' way), and a nonzero return from a setter means the names were not accepted
' (typically a mis-spelling). Load case names and load combination names are
' separate namespaces - each setter only accepts its own kind of name.

' Capture the CURRENT case and combo display selections so they can be put back
' later. A returned count of 0, or a list holding a single blank entry (SAFE's
' way of saying "nothing selected"), is recorded as "nothing selected".
Private Sub SaveDisplayFilterState()
    Dim n As Long
    Dim list() As String
    Dim ret As Long
    Dim errNum As Long, errDesc As String

    mSavedCaseCount = 0
    mSavedComboCount = 0
    Erase mSavedCases
    Erase mSavedCombos
    mFilterCasesApplied = False
    mFilterCombosApplied = False
    mFilterStateValid = True

    ' --- the load cases selected at the moment ---
    On Error Resume Next
    Err.Clear
    ret = gDB.GetLoadCasesSelectedForDisplay(n, list)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum <> 0 Then
        LogMsg "SaveDisplayFilterState: could not read the current display load-case " & _
               "selection - " & ErrorText(errNum, errDesc) & vbCrLf & _
               "  (it will be restored as ""none selected"")"
    ElseIf ret <> 0 Then
        LogMsg "SaveDisplayFilterState: GetLoadCasesSelectedForDisplay returned " & ret & _
               " (the previous display load-case selection will not be restored exactly)"
    Else
        mSavedCaseCount = CleanNameList(list, n, mSavedCases)
        LogMsg "SaveDisplayFilterState: previous display load-case selection = " & _
               DescribeNames(mSavedCases, mSavedCaseCount)
    End If

    ' --- the load combinations selected at the moment ---
    n = 0
    Erase list
    On Error Resume Next
    Err.Clear
    ret = gDB.GetLoadCombinationsSelectedForDisplay(n, list)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum <> 0 Then
        LogMsg "SaveDisplayFilterState: could not read the current display " & _
               "load-combination selection - " & ErrorText(errNum, errDesc) & vbCrLf & _
               "  (it will be restored as ""none selected"")"
    ElseIf ret <> 0 Then
        LogMsg "SaveDisplayFilterState: GetLoadCombinationsSelectedForDisplay returned " & _
               ret & " (the previous display load-combination selection will not be " & _
               "restored exactly)"
    Else
        mSavedComboCount = CleanNameList(list, n, mSavedCombos)
        LogMsg "SaveDisplayFilterState: previous display load-combination selection = " & _
               DescribeNames(mSavedCombos, mSavedComboCount)
    End If
End Sub

' Apply the requested display filters. mFilterCasesApplied / mFilterCombosApplied
' record which setter actually ran, so RestoreDisplayFilterState only undoes
' those. Returns False when a setter rejected the names (nonzero return, or a COM
' error) after logging what was rejected; the caller must then abort the export,
' because continuing would write UNFILTERED force results, which is far more
' dangerous for the engineering workflow than no export at all.
Private Function ApplyDisplayFilterState( _
    ByRef caseNames() As String, ByVal nCases As Long, _
    ByRef comboNames() As String, ByVal nCombos As Long) As Boolean

    Dim ret As Long
    Dim errNum As Long, errDesc As String

    ApplyDisplayFilterState = False       ' until every requested setter has run

    If nCases > 0 Then
        On Error Resume Next
        Err.Clear
        ret = gDB.SetLoadCasesSelectedForDisplay(caseNames)
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum <> 0 Or ret <> 0 Then
            LogMsg "ERROR in ApplyDisplayFilterState: SAFE would not accept the load " & _
                   "case name(s) " & DescribeNames(caseNames, nCases)
            If errNum <> 0 Then
                LogMsg "  the call raised " & ErrorText(errNum, errDesc)
            Else
                LogMsg "  SetLoadCasesSelectedForDisplay returned " & ret & _
                       " (nonzero = the name(s) were not accepted)"
            End If
            LogMsg "  Names must match SAFE exactly. LoadCases takes load CASE names " & _
                   "(e.g. ""LIVE""); a load COMBINATION name (e.g. ""1.4DL+1.6LL"") " & _
                   "belongs in LoadCombos."
            Exit Function
        End If

        mFilterCasesApplied = True
        LogMsg "ApplyDisplayFilterState: result tables restricted to load cases - " & _
               DescribeNames(caseNames, nCases)
    End If

    If nCombos > 0 Then
        On Error Resume Next
        Err.Clear
        ret = gDB.SetLoadCombinationsSelectedForDisplay(comboNames)
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum <> 0 Or ret <> 0 Then
            LogMsg "ERROR in ApplyDisplayFilterState: SAFE would not accept the load " & _
                   "combination name(s) " & DescribeNames(comboNames, nCombos)
            If errNum <> 0 Then
                LogMsg "  the call raised " & ErrorText(errNum, errDesc)
            Else
                LogMsg "  SetLoadCombinationsSelectedForDisplay returned " & ret & _
                       " (nonzero = the name(s) were not accepted)"
            End If
            LogMsg "  Names must match SAFE exactly. LoadCombos takes load COMBINATION " & _
                   "names (e.g. ""1.4DL+1.6LL""); a load CASE name (e.g. ""LIVE"") " & _
                   "belongs in LoadCases."
            Exit Function
        End If

        mFilterCombosApplied = True
        LogMsg "ApplyDisplayFilterState: result tables restricted to load " & _
               "combinations - " & DescribeNames(comboNames, nCombos)
    End If

    ApplyDisplayFilterState = True
End Function

' Put the display selections back the way SaveDisplayFilterState found them.
' Idempotent: the normal exit and the Fatal path both call it, and it clears
' mFilterStateValid when it is done, so a second call is a no-op. Only the
' selections that ApplyDisplayFilterState actually changed are restored, so a
' setter that never ran (or failed) is not overwritten with a stale copy.
' A failure here cannot be fatal - the export result is already decided - and it
' can never throw, but it IS logged with its error number, because it leaves the
' display filter changed in the user's SAFE session.
Private Sub RestoreDisplayFilterState()
    Dim ret As Long
    Dim errNum As Long, errDesc As String

    If Not mFilterStateValid Then Exit Sub     ' nothing saved, or already restored

    ' SAFE quirk: a list holding a single blank string means "select NOTHING" -
    ' this is how a previously empty selection is put back.
    Dim none(0 To 0) As String
    none(0) = ""

    If mFilterCasesApplied Then
        On Error Resume Next
        Err.Clear
        If mSavedCaseCount > 0 And ArrLenStr(mSavedCases) > 0 Then
            ret = gDB.SetLoadCasesSelectedForDisplay(mSavedCases)
        Else
            ret = gDB.SetLoadCasesSelectedForDisplay(none)
        End If
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum <> 0 Then
            LogMsg "RestoreDisplayFilterState: could not restore the display load-case " & _
                   "selection - " & ErrorText(errNum, errDesc)
        ElseIf ret <> 0 Then
            LogMsg "RestoreDisplayFilterState: SetLoadCasesSelectedForDisplay returned " & _
                   ret & " while restoring the previous display load-case selection."
        Else
            LogMsg "RestoreDisplayFilterState: display load cases restored - " & _
                   DescribeNames(mSavedCases, mSavedCaseCount)
        End If
    End If

    If mFilterCombosApplied Then
        On Error Resume Next
        Err.Clear
        If mSavedComboCount > 0 And ArrLenStr(mSavedCombos) > 0 Then
            ret = gDB.SetLoadCombinationsSelectedForDisplay(mSavedCombos)
        Else
            ret = gDB.SetLoadCombinationsSelectedForDisplay(none)
        End If
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum <> 0 Then
            LogMsg "RestoreDisplayFilterState: could not restore the display " & _
                   "load-combination selection - " & ErrorText(errNum, errDesc)
        ElseIf ret <> 0 Then
            LogMsg "RestoreDisplayFilterState: SetLoadCombinationsSelectedForDisplay " & _
                   "returned " & ret & " while restoring the previous display " & _
                   "load-combination selection."
        Else
            LogMsg "RestoreDisplayFilterState: display load combinations restored - " & _
                   DescribeNames(mSavedCombos, mSavedComboCount)
        End If
    End If

    ' Consume the saved state, so a second call (normal exit + Fatal, or a
    ' repeated call from the caller) does nothing.
    mFilterCasesApplied = False
    mFilterCombosApplied = False
    mFilterStateValid = False
    mSavedCaseCount = 0
    mSavedComboCount = 0
    Erase mSavedCases
    Erase mSavedCombos
End Sub

' Clean a list of names: Trim every entry, DROP empty entries and DROP EXACT
' duplicates (compared byte-wise, so "LIVE" and "live" stay distinct), keeping
' the original casing of everything that survives - SAFE matches these strings
' exactly. srcCount is the count the caller was given (NormalizeTables, or a
' Get...SelectedForDisplay call) and is authoritative: some API calls report more
' entries than the returned array actually holds.
Private Function CleanNameList( _
    ByRef src() As String, ByVal srcCount As Long, _
    ByRef clean() As String) As Long

    Dim i As Long, j As Long, n As Long
    Dim nm As String
    Dim outCount As Long
    Dim dup As Boolean

    Erase clean

    n = ArrLenStr(src)
    If n > srcCount Then n = srcCount      ' never walk past the reported count
    If n <= 0 Then
        CleanNameList = 0
        Exit Function
    End If

    ReDim clean(0 To n - 1)
    outCount = 0
    For i = 0 To n - 1
        nm = Trim$(src(i))
        If Len(nm) > 0 Then
            dup = False
            For j = 0 To outCount - 1
                If StrComp(clean(j), nm, vbBinaryCompare) = 0 Then
                    dup = True
                    Exit For
                End If
            Next j
            If Not dup Then
                clean(outCount) = nm
                outCount = outCount + 1
            End If
        End If
    Next i

    If outCount = 0 Then
        Erase clean
    ElseIf outCount < n Then
        ReDim Preserve clean(0 To outCount - 1)
    End If

    CleanNameList = outCount
End Function

' Render a name list for the log: "2 name(s): DEAD, LIVE", or "nothing selected".
Private Function DescribeNames(ByRef names() As String, ByVal count As Long) As String
    Dim i As Long, n As Long
    Dim buf As String

    n = ArrLenStr(names)
    If count <= 0 Or n = 0 Then
        DescribeNames = "nothing selected"
        Exit Function
    End If
    If n > count Then n = count

    For i = 0 To n - 1
        If Len(buf) > 0 Then buf = buf & ", "
        buf = buf & names(i)
    Next i
    DescribeNames = count & " name(s): " & buf
End Function

Private Function GetWorksheet(ByVal Wb As Workbook, ByVal SheetName As String) As Worksheet
    ' Returns the worksheet, creating it (at the end) if it does not exist.
    On Error Resume Next
    Set GetWorksheet = Wb.Worksheets(SheetName)
    On Error GoTo 0
    If GetWorksheet Is Nothing Then
        Set GetWorksheet = Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.Count))
        GetWorksheet.Name = SheetName
    End If
End Function

Private Function ParseStartCell( _
    ByVal ws As Worksheet, ByVal StartCell As String, _
    ByRef r As Long, ByRef c As Long) As Boolean
    ' Resolve "B3" style coordinates against the target worksheet.
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(StartCell)
    On Error GoTo 0
    If rng Is Nothing Then
        ParseStartCell = False
        Exit Function
    End If
    r = rng.Row
    c = rng.Column
    ParseStartCell = True
End Function

Private Function ArrLenStr(ByRef arr() As String) As Long
    ' Length of a dynamic String array; 0 if not dimensioned (avoids
    ' "Subscript out of range" when SAFE returns nothing).
    On Error GoTo NoArr
    ArrLenStr = UBound(arr) - LBound(arr) + 1
    Exit Function
NoArr:
    ArrLenStr = 0
End Function

Private Function ArrLenLng(ByRef arr() As Long) As Long
    ' The same guard for the Long arrays the API returns (ImportType()). It has
    ' to be a SECOND function: a VBA array parameter is checked on its element
    ' type, so a Long() array cannot be passed to ArrLenStr's String()
    ' parameter (and VBA has no generic array type to fall back on).
    On Error GoTo NoArr
    ArrLenLng = UBound(arr) - LBound(arr) + 1
    Exit Function
NoArr:
    ArrLenLng = 0
End Function

' ===========================================================================
' LOGGING
' ===========================================================================

' PUBLIC so companion modules (e.g. SAFE_Use.bas) can append to the SAME
' log: ShowLog / GetLog then show one continuous story for a session instead of
' the library's messages landing in one place and theirs in another.
Public Sub LogMsg(ByVal msg As String)
    gLog = gLog & msg & vbCrLf
    Debug.Print msg
End Sub

' Record the current Err object and write it to the log.
' The number is logged BOTH in decimal and as an 8-digit hexadecimal HRESULT,
' because the decimal number VBA reports for a failed COM call depends on how
' the call was marshalled (a stale or disconnected SAFE proxy surfaces as 5,
' 438, 462 or -2147417846/-2147417848 depending on the path taken), while the
' HRESULT identifies the real cause.
' Nothing is recorded when Err.Number is 0, so callers can share a "Fail" label
' between real errors and deliberate validation failures.
Private Sub LogError(ByVal Context As String, Optional ByVal Extra As String = "")
    If Err.Number = 0 Then Exit Sub

    ' Capture the details in LOCALS first: the helper calls below are allowed to
    ' raise (Error$(...)), and that would overwrite the Err object before we
    ' have finished reporting on it.
    Dim n As Long, d As String, src As String
    n = Err.Number
    d = Err.Description
    src = Err.Source

    gLastErrNumber = n
    gLastErrContext = Context
    gLastErrSource = src
    gLastErrDesc = d

    Dim msg As String, hint As String
    msg = "ERROR in " & Context & ": " & ErrorText(n, d)
    If Len(src) > 0 Then msg = msg & " [source: " & src & "]"
    If Len(Extra) > 0 Then msg = msg & vbCrLf & "  " & Extra
    hint = ErrHint(n)
    If Len(hint) > 0 Then msg = msg & vbCrLf & "  " & hint
    LogMsg msg
End Sub

' Record + log a failure that did NOT come from the Err object (for example an
' API call that returned Nothing instead of raising). LastErrorNumber() is set
' to 0 because there genuinely is no error number.
Private Sub SetLastError( _
    ByVal Context As String, _
    ByVal Text As String, _
    Optional ByVal Hint As String = "")
    ' Hint is the extra guidance line printed under the message, and the CALLER
    ' now chooses it. This routine used to append the "438 / check the SAFEv1
    ' reference" hint UNCONDITIONALLY, which aimed the reader at the reference
    ' library for failures that had nothing to do with it (an unknown table key,
    ' an analysis that has not been run, a table that is not importable, a block
    ' too wide for the worksheet). Pass ErrHint(438) where that advice really
    ' applies, a specific hint where the cause is known, or nothing at all (the
    ' default) when the message already says what to do.

    gLastErrNumber = 0
    gLastErrContext = Context
    gLastErrSource = ""
    gLastErrDesc = Text

    If Len(Hint) > 0 Then
        LogMsg "ERROR in " & Context & ": " & Text & vbCrLf & "  " & Hint
    Else
        LogMsg "ERROR in " & Context & ": " & Text
    End If
End Sub

Private Sub ClearLastError()
    gLastErrNumber = 0
    gLastErrSource = ""
    gLastErrDesc = ""
    gLastErrContext = ""
End Sub

' "438 (0x000001B6) - Object doesn't support this property or method"
Private Function ErrorText(ByVal Number As Long, ByVal Description As String) As String
    Dim d As String
    d = Trim$(Description)
    If Len(d) = 0 And Number > 0 And Number <= 65535 Then
        ' Fall back to VBA's built-in text for standard errors. Only for valid
        ' positive codes - Error$() raises for anything else (and would clobber
        ' Err.) which is why it is guarded here.
        On Error Resume Next
        d = Trim$(Error$(Number))
        On Error GoTo 0
    End If

    ErrorText = CStr(Number) & " (0x" & ErrHex(Number) & ")"
    If Len(d) > 0 Then ErrorText = ErrorText & " - " & d
End Function

' 8-digit two's-complement hex for an error number: -2147417846 -> "8001010A".
' Done with Double-based integer maths so it does not depend on how VBA's Hex()
' treats negative values, and so the high bit does not overflow a Long.
Private Function ErrHex(ByVal Number As Long) As String
    Dim u As Double
    Dim hi As Long, lo As Long

    If Number < 0 Then
        u = 4294967296# + Number          ' 2^32 + negative = unsigned 32-bit value
    Else
        u = Number
    End If

    hi = CLng(Int(u / 65536#))
    lo = CLng(u - Int(u / 65536#) * 65536#)
    ErrHex = Right$("000" & Hex$(hi), 4) & Right$("000" & Hex$(lo), 4)
End Function

' Guidance for the error numbers seen when a COM call to SAFE fails - mirrors
' the troubleshooting in the pile-cap Readme.txt ("if it returns 438 or 5
' error, check if the reference library for SAFE API is active... then check if
' the SAFE application is open, not running analysis and locked so that the API
' is active").
Private Function ErrHint(ByVal Number As Long) As String
    Select Case Number
        Case 5, 438
            ErrHint = "Hint: check that the 'SAFEv1' reference is ticked " & _
                      "(VBA IDE > Tools > References...), then that SAFE is open " & _
                      "with a model loaded, idle (not running analysis) and " & _
                      "locked, so that the API is active."
        Case 429, 424
            ErrHint = "Hint: no live SAFE automation object was found. Is SAFE " & _
                      "running in the same Windows session and at the same " & _
                      "integrity level (elevation) as Excel?"
        Case 462
            ErrHint = "Hint: the SAFE instance stopped responding or was closed " & _
                      "while the call was in flight. Call SAFEDisconnect, then " & _
                      "SAFEConnect."
        Case -2147417846              ' &H8001010A RPC_E_CALL_REJECTED
            ErrHint = "Hint: SAFE rejected the call because it is busy (analysing, " & _
                      "or showing a modal dialog). Wait until it is idle and retry."
        Case -2147417848              ' &H80010108 RPC_E_DISCONNECTED
            ErrHint = "Hint: the SAFE instance this module is attached to has " & _
                      "exited. Call SAFEDisconnect, then SAFEConnect."
    End Select
End Function

' ---------------------------------------------------------------------------
' Last-error accessors (populated by LogError / SetLastError)
' ---------------------------------------------------------------------------

Public Function LastErrorNumber() As Long
    ' Decimal VBA error number, or 0 if the failure had no error number.
    LastErrorNumber = gLastErrNumber
End Function

Public Function LastErrorDescription() As String
    LastErrorDescription = gLastErrDesc
End Function

Public Function LastErrorSource() As String
    LastErrorSource = gLastErrSource
End Function

Public Function LastErrorContext() As String
    ' Which procedure recorded the error, e.g. "SAFEConnect".
    LastErrorContext = gLastErrContext
End Function

Public Function LastErrorText() As String
    ' One-line summary, e.g. "SAFEConnect: 438 (0x000001B6) - Object doesn't "
    ' "support this property or method"
    If gLastErrNumber = 0 And Len(gLastErrDesc) = 0 Then Exit Function
    LastErrorText = ErrorText(gLastErrNumber, gLastErrDesc)
    If Len(gLastErrContext) > 0 Then _
        LastErrorText = gLastErrContext & ": " & LastErrorText
End Function

Public Sub ClearLog()
    gLog = ""
    ClearLastError
End Sub

Public Function GetLog() As String
    GetLog = gLog
End Function

Public Sub ShowLog()
    If Len(gLog) = 0 Then
        MsgBox "(log is empty)", vbInformation, "SAFE Export Log"
    Else
        MsgBox gLog, vbInformation, "SAFE Export Log"
    End If
End Sub

Public Function GetLastExportFailures() As Long
    ' Number of tables that FAILED in the last ExportSAFETables call (0 = none).
    ' A failed table is one whose read could not be served - a nonzero
    ' GetTableForDisplayArray code with NO column headers returned (the table key
    ' is not valid for this model / SAFE version, or analysis has not been run),
    ' or a COM error. Such a table is written to the destination sheet as a bold
    ' red "Table '<key>' : FAILED - <reason>" marker, logged as an error, and is
    ' NOT included in the number ExportSAFETables returns. A valid key that
    ' simply has no rows is an EMPTY table, not a failure, and is not counted
    ' here either - it gets the plain "no data returned" marker.
    ' A WRITE failure is counted too: a table whose data WAS read but could not
    ' be written in full (too wide for the worksheet, no room left for its
    ' title/header rows, or a continuation sheet running out of room mid-spill)
    ' gets the same bold red marker and is counted here.
    GetLastExportFailures = mFailedTables
End Function

' ===========================================================================
' [COMMENTED OUT - HARDCODED IMPORT SECTION]
' The SAFEv1.tlb reference is now added through the Excel VBA IDE and read
' from that IDE setting, so nothing is hardcoded here any more.
'
' To add the reference:
'   1. Alt+F11 to open the VBA IDE.
'   2. Tools > References...
'   3. Tick "SAFEv1" if listed, otherwise Browse... and select SAFEv1.tlb
'      from the SAFE installation folder (e.g. ...\SAFE 20\SAFEv1.tlb).
'
' (The old code that added the reference programmatically from a hardcoded
'  path is commented out below. It required "Trust access to the VBA project
'  object model" and a hardcoded path - neither is needed any more.)
' ===========================================================================
'
'Public Sub SAFESetupReference()
'    Dim ref As Object
'    Dim found As Boolean
'    found = False
'    For Each ref In ThisWorkbook.VBProject.References
'        If InStr(1, ref.Name, "SAFEv1", vbTextCompare) > 0 Then
'            found = True
'            Exit For
'        End If
'    Next ref
'
'    If found Then
'        MsgBox "SAFEv1 reference already present: " & ref.Name, vbInformation
'        Exit Sub
'    End If
'
'    On Error GoTo RefFail
'    ThisWorkbook.VBProject.References.AddFromFile ""   ' path was hardcoded
'    MsgBox "Added SAFEv1 reference.", vbInformation
'    Exit Sub
'
'RefFail:
'    MsgBox "Could not add reference:" & vbCrLf & Err.Description, vbExclamation
'End Sub

' ===========================================================================
' DEMOS
' ===========================================================================

' Example: export several tables to a sheet, stacked downwards from A1.
Public Sub DemoExport()
    Dim Tables() As String
    ReDim Tables(0 To 2)
    Tables(0) = "Point Object Connectivity"
    Tables(1) = "Area Load Assignments - Uniform"
    Tables(2) = "Element Forces - Area Shells"   ' requires analysis results

    Dim n As Long
    n = ExportSAFETables(Tables, DEF_SHEET, DEF_START, False, True)
    If n >= 0 Then
        MsgBox "Exported " & n & " of " & (UBound(Tables) - LBound(Tables) + 1) & _
               " table(s) to sheet '" & DEF_SHEET & "' at " & DEF_START & ".", _
               vbInformation
    Else
        MsgBox "Export failed - see Immediate window (Ctrl+G) / ShowLog.", vbExclamation
    End If
End Sub

' Example: a single table, at a specific tab + coordinate.
Public Sub DemoExportSingle()
    Dim n As Long
    n = ExportSAFETables("Element Forces - Area Shells", "SAFE Forces", "B2")
    If n < 0 Then
        MsgBox "Export failed - see Immediate window (Ctrl+G) / ShowLog.", vbExclamation
    End If
End Sub

' Example: only include specific load cases in the result tables.
' ("LIVE, DEAD" or Array("LIVE","DEAD") also work.)
' For load COMBINATIONS (e.g. "1.4DL+1.6LL") use the LoadCombos argument instead
' - see DemoExportFilteredCombo below. The two filters are independent, and SAFE
' keeps load cases and load combinations as separate lists.
Public Sub DemoExportFiltered()
    Dim n As Long
    n = ExportSAFETables("Element Forces - Area Shells", "SAFE Forces LIVE", "B2", _
                         LoadCases:="LIVE")
    If n < 0 Then
        MsgBox "Export failed - see Immediate window (Ctrl+G) / ShowLog.", vbExclamation
    End If
End Sub

' Example: only include specific load COMBINATIONS in the result tables - the
' names SAFE shows under "Load Combinations" (the pile-cap sheet filters by
' combination, e.g. "1.4DL+1.6LL" / "1.4(D+WH)"). A combination name that SAFE
' does not know makes the export fail with -1 rather than exporting unfiltered
' force results - the log names the rejected string.
Public Sub DemoExportFilteredCombo()
    Dim n As Long
    n = ExportSAFETables("Element Forces - Area Shells", "SAFE Forces COMBO", "B2", _
                         LoadCombos:="1.4DL+1.6LL")
    If n < 0 Then
        MsgBox "Export failed - see Immediate window (Ctrl+G) / ShowLog.", vbExclamation
    End If
End Sub

' Example: dump all available table keys so you can pick exact names.
Public Sub DemoListTables()
    Dim n As Long
    n = ListSAFETables("SAFE Table Keys", "A1")
    If n >= 0 Then
        MsgBox "Listed " & n & " table(s) SAFE reports (GetAllTables) on sheet " & _
               "'SAFE Table Keys'.", vbInformation
    End If
End Sub
