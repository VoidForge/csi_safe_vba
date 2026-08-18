Attribute VB_Name = "SAFE_Export"
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  SAFE_Export.bas  --  CSI SAFE database-table extractor for Excel / VBA
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
'        IncludeHeader     : write the column-header row (default True)
'        LoadCases   (new) : optional load-case filter - only these load cases
'                            appear in result tables. Accepts one name, a
'                            comma-separated list, or a String() array.
'                            Empty / default = export ALL load cases.
'                            (Non-result tables are not affected by SAFE's
'                             display load-case filter.)
'        Returns    : number of tables written; -1 on a fatal error
'                     (see Immediate window / ShowLog for details)
'
'    ListSAFETables(SheetName, StartCell)   : dumps every available table key
'                                             (handy for finding exact names)
'    WriteSAFETable(TableKey, Data, ...)    : writes a 2-D array BACK into SAFE
'                                             (edit + apply). Optional bonus.
'    SAFEConnect / SAFEDisconnect           : attach / release the running SAFE
'
'  TABLE KEYS are the same strings shown in SAFE's "Display > Show Tables",
'  e.g. "Point Object Connectivity", "Area Load Assignments - Uniform",
'  "Element Forces - Area Shells", "Joint Displacements", ...
'  Run ListSAFETables to see the exact keys available in your model.
'
'  QUIRKS LEARNED (from the reference Python script + SAFE API docs)
'    - SAFE's COM main object is called "ETABSObject" (ETABS infrastructure).
'    - GetTableForDisplayArray returns NONZERO when there is "nothing to show"
'      (empty table / analysis not run yet). We treat that as a WARNING, write
'      the table title + note, and keep going instead of aborting.
'    - FieldKeyList must be a single blank string to get ALL columns.
'    - GroupName "" (or "All") returns data for all objects in the model.
'    - The data array is FLATTENED row-by-row; rows are rebuilt using the
'      number of columns = number of FieldsKeysIncluded.
'    - Some cells come back as empty strings; written as blank cells.
'    - For editing tables, GroupName is documented as "not active in this
'      release" - pass "".
'    - Editing tables requires the model UNLOCKED (SetModelIsLocked False).
'    - ApplyEditedTables can corrupt the model on a fatal error - SAFE docs
'      recommend saving the model BEFORE calling it. We check the error counts.
'    - When attaching to a running instance, NEVER call ApplicationExit
'      (it would close the user's SAFE session).
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

' ---------------------------------------------------------------------------
' Module state (connection + log)
' ---------------------------------------------------------------------------
Private gSAFE As cOAPI           ' SAFE API object
Private gSapModel As cSapModel   ' model object
Private gDB As cDatabaseTables   ' database tables (cached)
Private gConnected As Boolean
Private gLog As String

' ===========================================================================
' CONNECTION - attach to the running SAFE instance
' ===========================================================================

Public Function SAFEConnect() As Boolean
    Dim helper As cHelper

    ' Strategy 1: attach to the running instance through the ROT.
    On Error Resume Next
    Set gSAFE = GetObject(, SAFE_PROGID)
    On Error GoTo 0

    ' Strategy 2: via the API Helper object (documented early-bound approach).
    If gSAFE Is Nothing Then
        On Error Resume Next
        Set helper = New Helper
        If Not helper Is Nothing Then
            Set gSAFE = helper.GetObject(SAFE_PROGID)
        End If
        On Error GoTo 0
    End If

    If gSAFE Is Nothing Then
        LogMsg "SAFEConnect: no running SAFE instance found." & vbCrLf & _
               "Start SAFE, open the model, and try again."
        SAFEConnect = False
        Exit Function
    End If

    On Error GoTo Fail
    Set gSapModel = gSAFE.SapModel
    If gSapModel Is Nothing Then GoTo Fail
    Set gDB = gSapModel.DatabaseTables
    gConnected = True
    LogMsg "Attached to running SAFE instance (progid: " & SAFE_PROGID & ")."
    SAFEConnect = True
    Exit Function

Fail:
    LogMsg "SAFEConnect failed: " & Err.Description & " (" & Err.Number & ")"
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
    Optional ByVal LoadCases As Variant = "") As Long
    ' Returns the number of tables written, or -1 on a fatal error.
    ' LoadCases: optional filter - only these load cases appear in result
    ' tables ("" / missing = export ALL load cases; non-result tables are
    ' unaffected by SAFE's display load-case filter).

    Dim savedScreen As Boolean
    savedScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fatal

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
    Dim data As Variant

    ' --- optional load-case filter (saved and restored afterwards) ---
    Dim savedCases() As String
    Dim savedCount As Long
    Dim filterApplied As Boolean
    filterApplied = False

    Dim caseNames() As String
    Dim nCases As Long
    nCases = NormalizeTables(LoadCases, caseNames)
    If nCases > 0 Then
        Dim getRet As Long
        getRet = gDB.GetLoadCasesSelectedForDisplay(savedCount, savedCases)
        Dim setRet As Long
        setRet = gDB.SetLoadCasesSelectedForDisplay(caseNames)
        If setRet <> 0 Then
            LogMsg "ExportSAFETables: SetLoadCasesSelectedForDisplay returned " & setRet
        Else
            filterApplied = True
            LogMsg "ExportSAFETables: restricting result tables to " & nCases & _
                   " load case(s): " & Join(caseNames, ", ")
        End If
    End If

    For i = 0 To nTables - 1
        Dim key As String
        key = Trim(names(i))
        If Len(key) > 0 Then
            data = SAFETableToArray(key, hdrs, warn)
            If Len(warn) > 0 Then LogMsg "[" & key & "] " & warn

            If ArrLenStr(hdrs) > 0 Or Not IsEmpty(data) Then
                ' Normal case: write title + headers + data block.
                WriteTableBlock ws, r, c, key, hdrs, data, IncludeHeader, StackHorizontally, r, c
                written = written + 1
            Else
                ' Completely empty table - leave a visible marker and carry on.
                ws.Cells(r, c).Value = "Table '" & key & "' : no data returned"
                If Len(warn) > 0 Then ws.Cells(r, c + 1).Value = warn
                r = r + 2
            End If
        End If
    Next i

    ' restore the load-case display selection that existed before the export
    RestoreLoadCaseFilter savedCases, savedCount, filterApplied

    Application.ScreenUpdating = savedScreen
    ExportSAFETables = written
    Exit Function

Fatal:
    RestoreLoadCaseFilter savedCases, savedCount, filterApplied
    Application.ScreenUpdating = savedScreen
    LogMsg "ExportSAFETables fatal error: " & Err.Description & " (" & Err.Number & ")"
    ExportSAFETables = -1
End Function

' ===========================================================================
' DIAGNOSTIC - list all available table keys (exact strings to use above)
' ===========================================================================

Public Function ListSAFETables( _
    Optional ByVal SheetName As String = DEF_SHEET, _
    Optional ByVal StartCell As String = DEF_START) As Long
    ' Writes "Table Key | Table Name | Import Type" for every available table.
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
    Dim ret As Long

    ret = gDB.GetAvailableTables(NumberTables, TableKey, TableName, ImportType)
    If ret <> 0 Then
        LogMsg "ListSAFETables: GetAvailableTables returned " & ret
        Application.ScreenUpdating = savedScreen
        ListSAFETables = -1
        Exit Function
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
        If i < ArrLenStr(ImportType) Then ws.Cells(r + i, c + 2).Value = ImportType(i)
    Next i

    ws.Columns(c).AutoFit
    Application.ScreenUpdating = savedScreen
    ListSAFETables = n
    Exit Function

Fatal:
    Application.ScreenUpdating = savedScreen
    LogMsg "ListSAFETables fatal error: " & Err.Description & " (" & Err.Number & ")"
    ListSAFETables = -1
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

    On Error GoTo Fail

    If Not gConnected Then
        If Not SAFEConnect() Then GoTo Fail
    End If

    ' Editing tables requires an unlocked model (quirk).
    Dim ret As Long
    If UnlockModel Then
        ret = gSapModel.SetModelIsLocked(False)
        If ret <> 0 Then LogMsg "WriteSAFETable: SetModelIsLocked(False) returned " & ret
    End If

    ' Pull current table structure (headers + version + existing rows).
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
    On Error GoTo 0
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

    ret = gDB.SetTableForEditingArray(TableKey, TableVersion, FieldsKeysIncluded, nRows, flat)
    If ret <> 0 Then
        LogMsg "WriteSAFETable: SetTableForEditingArray('" & TableKey & "') returned " & ret
        GoTo Fail
    End If

    ' Apply all edited tables. Check the error counts - a nonzero return or
    ' fatal errors can leave the model in a bad state.
    Dim NumFatalErrors As Long, NumErrorMsgs As Long
    Dim NumWarnMsgs As Long, NumInfoMsgs As Long
    Dim ImportLog As String
    ret = gDB.ApplyEditedTables(False, NumFatalErrors, NumErrorMsgs, NumWarnMsgs, NumInfoMsgs, ImportLog)
    If ret <> 0 Or NumFatalErrors > 0 Then
        LogMsg "WriteSAFETable: ApplyEditedTables returned " & ret & _
               " (fatal=" & NumFatalErrors & ", errors=" & NumErrorMsgs & _
               ", warnings=" & NumWarnMsgs & ")"
        If NumErrorMsgs + NumFatalErrors > 0 Then LogMsg "ImportLog: " & ImportLog
        gDB.CancelTableEditing     ' clear the pending edit buffer
        GoTo Fail
    End If

    ' Clear the internal edit buffer.
    gDB.CancelTableEditing
    WriteSAFETable = True
    Exit Function

Fail:
    WriteSAFETable = False
End Function

' ===========================================================================
' CORE READ - pull one table into a 2-D array (quirk-tolerant)
' ===========================================================================

Private Function SAFETableToArray( _
    ByVal TableKey As String, _
    ByRef Headers() As String, _
    ByRef Warning As String) As Variant
    ' Returns a 1-based 2-D Variant array [row, col] of data (headers excluded),
    ' or Empty when the table has nothing to write. Headers are filled in.
    On Error GoTo ErrHandler

    Erase Headers
    Warning = ""

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

    ' QUIRK: SAFE returns nonzero when there is nothing to show
    ' (empty table, analysis not run, etc.). Not necessarily fatal.
    If ret <> 0 Then
        Warning = "API code " & ret & " (table empty / not yet populated?)"
    End If

    Dim nCols As Long
    nCols = ArrLenStr(FieldsKeysIncluded)
    If nCols <= 0 Then
        SAFETableToArray = Empty
        Exit Function
    End If

    ' Copy the column headers.
    ReDim Headers(0 To nCols - 1)
    Dim j As Long
    For j = 0 To nCols - 1
        Headers(j) = FieldsKeysIncluded(j)
    Next j

    Dim nRows As Long
    nRows = NumberRecords
    If nRows <= 0 Then
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
    Warning = "COM error " & Err.Number & ": " & Err.Description
    SAFETableToArray = Empty
End Function

' ===========================================================================
' WRITE A TABLE BLOCK TO THE WORKSHEET (title + headers + data)
' ===========================================================================

Private Function WriteTableBlock( _
    ByVal ws As Worksheet, _
    ByVal r0 As Long, ByVal c0 As Long, _
    ByVal Title As String, _
    ByRef Headers() As String, _
    ByVal Data As Variant, _
    ByVal WithHeader As Boolean, _
    ByVal Horiz As Boolean, _
    ByRef NextRow As Long, ByRef NextCol As Long) As Long
    ' Writes Title / headers / data starting at (r0, c0) and advances the
    ' cursor (NextRow/NextCol) so the caller can place the next table.
    ' Returns the number of data rows written.

    Dim nCols As Long
    nCols = ArrLenStr(Headers)

    Dim nRows As Long
    nRows = 0
    If Not IsEmpty(Data) Then
        On Error Resume Next
        nRows = UBound(Data, 1)
        On Error GoTo 0
        If nRows < 1 Then nRows = 0
    End If

    Dim r As Long, c As Long, j As Long
    r = r0
    c = c0

    ' Title row
    If Len(Title) > 0 Then
        ws.Cells(r, c).Value = Title
        ws.Cells(r, c).Font.Bold = True
        r = r + 1
    End If

    ' Header row
    If WithHeader And nCols > 0 Then
        For j = 1 To nCols
            ws.Cells(r, c + j - 1).Value = Headers(j - 1)
            ws.Cells(r, c + j - 1).Font.Bold = True
            ws.Cells(r, c + j - 1).Interior.Color = RGB(221, 235, 247)
        Next j
        r = r + 1
    End If

    ' Data block (bulk write for speed)
    If nRows > 0 And nCols > 0 Then
        ws.Range(ws.Cells(r, c), ws.Cells(r + nRows - 1, c + nCols - 1)).Value2 = Data
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
    NormalizeTables = 0
End Function

' Restore the load-case display selection that was in place before a filtered
' export (best effort - the return code is ignored).
' Note (SAFE quirk): passing a single blank string selects NO load cases.
Private Sub RestoreLoadCaseFilter( _
    ByRef savedCases() As String, _
    ByVal savedCount As Long, _
    ByVal filterWasApplied As Boolean)

    If Not filterWasApplied Then Exit Sub

    Dim none(0 To 0) As String
    none(0) = ""

    On Error Resume Next
    If savedCount > 0 And ArrLenStr(savedCases) > 0 Then
        gDB.SetLoadCasesSelectedForDisplay savedCases
    Else
        gDB.SetLoadCasesSelectedForDisplay none
    End If
    On Error GoTo 0
End Sub

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

' ===========================================================================
' LOGGING
' ===========================================================================

Private Sub LogMsg(ByVal msg As String)
    gLog = gLog & msg & vbCrLf
    Debug.Print msg
End Sub

Public Sub ClearLog()
    gLog = ""
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
Public Sub DemoExportFiltered()
    Dim n As Long
    n = ExportSAFETables("Element Forces - Area Shells", "SAFE Forces LIVE", "B2", _
                         LoadCases:="LIVE")
    If n < 0 Then
        MsgBox "Export failed - see Immediate window (Ctrl+G) / ShowLog.", vbExclamation
    End If
End Sub

' Example: dump all available table keys so you can pick exact names.
Public Sub DemoListTables()
    Dim n As Long
    n = ListSAFETables("SAFE Table Keys", "A1")
    If n >= 0 Then
        MsgBox "Listed " & n & " table(s) on sheet 'SAFE Table Keys'.", vbInformation
    End If
End Sub
