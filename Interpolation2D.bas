Attribute VB_Name = "Interpolation2D"
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  Interpolation2D.bas  --  linear interpolation on an UNSTRUCTURED 2-D mesh
' ============================================================================
'  WHAT IT DOES
'    Reads a cloud of scattered (X, Y, Value...) points from a worksheet, builds
'    a Delaunay triangulation of them ONCE, and then interpolates any query
'    point by finding the triangle that contains it and blending that triangle's
'    three corner values with barycentric weights. Inside a triangle the result
'    is therefore exactly LINEAR in X and Y, and the surface is continuous across
'    the edges between triangles.
'    The mesh does NOT have to be a grid, and it does not have to be ordered:
'    any set of points that is not (nearly) collinear will do.
'
'  STANDALONE
'    This module calls NOTHING outside itself - no reference beyond Excel
'    itself. It carries its own log and its own sheet writer, so it can be
'    dropped into any workbook.
'
'  SHEET FORMATS  (all ranges are read from their FIRST row downwards - there is
'                  no header row, and a BLANK row ends the table)
'    mesh point table      X | Y | Value          (3 columns)
'                          X | Y | V1 | V2 | ...  (3+ columns: several value
'                                                  columns are interpolated in
'                                                  one pass, X and Y must be the
'                                                  first two columns)
'    sample point table    Name | X | Y           (3+ columns - the name column
'                          X | Y                   is optional: with exactly 2
'                                                  column it is X | Y)
'    output table          Name? | X | Y | V1 | V2 | ...   written with NO title
'                          and NO header row, flush on the top-left corner given
'
'  MESH AND SAMPLE DATA LIVE IN THIS MODULE
'    No arguments carry the clouds around: ReadMeshPoints parks the mesh (and its
'    triangulation) in module-level arrays, ReadSamplePoints parks the query
'    points AND builds their values through the same interpolate subr, and the
'    writer/accessors read that state back. So the calls go in ONE order:
'
'        ReadMeshPoints      (mesh first - it is what interpolation needs)
'        ReadSamplePoints    (samples second - this builds the values)
'        WriteSampleTable    (prints them)
'
'    RunInterpolation does exactly those three in that order, and re-running any
'    step is harmless: ReadSamplePoints rebuilds the values against whatever mesh
'    is currently loaded, and reading a NEW mesh automatically re-interpolates the
'    sample points that are already there.
'
'  ALGORITHM NOTES
'    - Triangulation: Bowyer-Watson incremental Delaunay, seeded with a super
'      triangle that is deleted again at the end. Points on a common circle, or
'      repeated points, are the classic trouble spots, so:
'        * duplicated / near-coincident mesh points are DROPPED (first one wins),
'          reported in the log;
'        * every geometric test carries a tolerance relative to the size of the
'          coordinates involved (see GeomTol), so a degenerate or collinear input
'          degrades into "nearest point" instead of producing garbage triangles.
'    - Point location: a plain loop over the triangles. That is O(triangles) per
'      query, which is deliberate - the data here is small, and a walk or a grid
'      index would be far more code for no measurable gain.
'    - OUTSIDE THE MESH the answer is the NEAREST mesh point's value (a flat
'      extrapolation), because that is what was asked for. It is NOT an
'      interpolation, so those rows are counted, listed in the log (first 10),
'      and flagged per row by SampleUsedNearest.
'
'  QUICK START
'    1. Alt+F11 -> File > Import File... -> choose this .bas
'    2. Edit the constants inside InterpolationDemo (the bottom of the file's
'       ENTRY POINTS section) and run it, or call RunInterpolation directly.
'    3. Read the log afterwards with InterpShowLog (or the Immediate window).
'
'  MESSAGE BOXES ARE OFF by default - every message goes to this module's log
'  instead (InterpShowLog / InterpGetLog). Set InterpMsgBoxLogging = True to see
'  them again.
' ============================================================================

' ---------------------------------------------------------------------------
' MODULE STATE - the mesh, the triangulation, the samples and the log
' ---------------------------------------------------------------------------
' Only the geometry that has to survive between calls is kept here. Everything
' that is just a working value stays a local in the sub that uses it.

' ---- mesh points: mZ(i, f) is value column f of mesh point i --------------
Private mN As Long                    ' number of mesh points in use
Private mNFields As Long              ' number of value columns per point
Private mX() As Double                ' 1 To mN
Private mY() As Double                ' 1 To mN
Private mZ() As Double                ' 1 To mN, 1 To mNFields

' ---- the triangulation of those points ------------------------------------
Private mTri() As Long                ' mTri(1 To mNTri, 1 To 3) - mesh point indices
Private mNTri As Long                 ' 0 = no triangle (degenerate mesh)

' ---- sample (query) points and their interpolated values ------------------
Private mNS As Long                   ' number of sample points in use
Private mSHasName As Boolean          ' the sample table had a name column
Private mSName() As String            ' 1 To mNS
Private mSX() As Double               ' 1 To mNS
Private mSY() As Double               ' 1 To mNS
Private mSV() As Double               ' mSV(1 To mNS, 1 To mNFields)
Private mSOut() As Boolean            ' True = that row fell outside the mesh
Private mSValuesReady As Boolean      ' mSV / mSOut are dimensioned and filled

' ---------------------------------------------------------------------------
' NUMERIC CONSTANTS
' ---------------------------------------------------------------------------
' 1E-12 relative is roughly ten times the rounding error of a 3x3 or 4x4
' determinant expansion, so anything under it is arithmetic noise, not geometry.
Private Const GEOM_EPS_REL As Double = 0.000000000001

' Barycentric coordinates are dimensionless, so this one needs no scaling: a
' weight of -1E-9 is "on the edge, one rounding step outside it".
Private Const BARY_EPS As Double = 0.000000001

' Two mesh points are the same point when they agree to this fraction of the
' mesh's own extent (the table is usually a copy-paste, so the real duplicates
' are exact - this tolerance only catches the near ones).
Private Const DUP_EPS_REL As Double = 0.000000001

' How many individual rows of a run the log lists by name before it switches to
' counting only (a 5,000-row sample table must not produce a 5,000-line log).
Private Const DETAIL_LIMIT As Long = 10

' ---------------------------------------------------------------------------
' LOGGING - this module's own
' ---------------------------------------------------------------------------
Private Const MSG_TITLE As String = "2-D interpolation"
Private mLog As String
Private mDupMsgs As Long              ' duplicate mesh points listed so far
Private mOutMsgs As Long              ' outside-the-mesh points listed so far

' A MsgBox is modal, which is fine while a module is being built and annoying in
' production, so every message of this module goes through InterpSay and this
' switch decides whether it is also SHOWN. Nothing is lost while it is False:
' the text always reaches the log.
Public InterpMsgBoxLogging As Boolean     ' False = silent

Private Sub InterpSay( _
    ByVal Text As String, _
    Optional ByVal Icon As Long = vbInformation, _
    Optional ByVal Title As String = "", _
    Optional ByVal AlsoLog As Boolean = True)

    If AlsoLog Then InterpLogMsg Text
    If Len(Title) = 0 Then Title = MSG_TITLE
    If InterpMsgBoxLogging Then MsgBox Text, Icon, Title
End Sub

' Append one line to this module's log. PUBLIC so a wrapper living in another
' module can register its own steps in the same story.
Public Sub InterpLogMsg(ByVal msg As String)
    mLog = mLog & msg & vbCrLf
    Debug.Print msg
End Sub

Public Function InterpGetLog() As String
    InterpGetLog = mLog
End Function

' NOTE: a MsgBox truncates a long string (about 1,000 characters), so this is for
' a quick look; InterpGetLog is the one to write to a file or a cell.
Public Sub InterpShowLog()
    If Len(mLog) = 0 Then
        MsgBox "(the interpolation log is empty)", vbInformation, MSG_TITLE & " - log"
    Else
        MsgBox mLog, vbInformation, MSG_TITLE & " - log"
    End If
End Sub

Public Sub InterpClearLog()
    mLog = ""
End Sub

' ===========================================================================
' 1. INPUT - ReadMeshPoints(SheetName, RangeAddress)
' ===========================================================================
' Reads the mesh point table (X | Y | Value | Value ...) off one worksheet into
' this module, then builds the triangulation.
'
' Reading walks down from the first row of the range and STOPS at the first
' blank row, so the range only has to START at the first data row.
'
' Rows are SKIPPED, with a line in the log, when X or Y (or any value cell) is
' not a number, and duplicated points are dropped (the first one wins). A mesh
' with fewer than three usable points, or one that is entirely collinear, still
' loads: there is simply no triangle, and every query then returns the nearest
' mesh point's value.
Public Sub ReadMeshPoints( _
    ByVal SheetName As String, _
    ByVal RangeAddress As String)

    Dim vals As Variant
    Dim nRowsIn As Long, nColsIn As Long
    Dim nF As Long
    Dim r As Long, f As Long, i As Long, j As Long
    Dim nBad As Long, nUse As Long, nDup As Long, nOut As Long
    Dim allBlank As Boolean, badVal As Boolean, dup As Boolean
    Dim rx() As Double, ry() As Double, rz() As Double
    Dim fx() As Double, fy() As Double, fz() As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim span As Double, dupTol As Double

    ' ---- 1. read the range --------------------------------------------------
    If Not ReadRange(SheetName, RangeAddress, "the MESH point table", vals) Then Exit Sub

    nRowsIn = 0
    nColsIn = 0
    On Error Resume Next
    nRowsIn = UBound(vals, 1)
    nColsIn = UBound(vals, 2)
    On Error GoTo 0

    If nRowsIn < 1 Or nColsIn < 3 Then
        InterpSay "The mesh range '" & SheetName & "'!" & RangeAddress & " is " & _
            nRowsIn & " row(s) x " & nColsIn & " column(s); at least three columns " & _
            "are needed: X | Y | Value (extra columns = further value columns).", _
            vbExclamation
        Exit Sub
    End If

    nF = nColsIn - 2                  ' the value columns are 3..nColsIn

    ' ---- 2. parse the rows --------------------------------------------------
    ' Arrays are sized for the WORST case (every row usable) and shortened by the
    ' counters; VBA's ReDim Preserve can only grow the LAST dimension, so the
    ' compacted copy is made at the end instead of shrinking in place.
    ReDim rx(1 To nRowsIn)
    ReDim ry(1 To nRowsIn)
    ReDim rz(1 To nRowsIn, 1 To nF)

    For r = 1 To nRowsIn
        allBlank = True
        For f = 0 To nColsIn - 1
            If Not IsBlankCell(vals(r, 1 + f)) Then
                allBlank = False
                Exit For
            End If
        Next f

        If allBlank Then
            InterpLogMsg "ReadMeshPoints: reading stopped at the blank row " & r & _
                         " of '" & SheetName & "'!" & RangeAddress & "."
            Exit For
        ElseIf Not (IsNumCell(vals(r, 1)) And IsNumCell(vals(r, 2))) Then
            nBad = nBad + 1
            InterpLogMsg "ReadMeshPoints: SKIPPED row " & r & " of '" & SheetName & _
                         "'!" & RangeAddress & " - X or Y is not a number ('" & _
                         SafeText(vals(r, 1)) & "', '" & SafeText(vals(r, 2)) & "')."
        Else
            badVal = False
            For f = 1 To nF
                If Not IsNumCell(vals(r, 2 + f)) Then badVal = True
            Next f

            If badVal Then
                nBad = nBad + 1
                InterpLogMsg "ReadMeshPoints: SKIPPED row " & r & " of '" & SheetName & _
                             "'!" & RangeAddress & " - at least one of its " & nF & _
                             " value column(s) holds something that is not a number."
            Else
                nUse = nUse + 1
                rx(nUse) = ToDbl(vals(r, 1))
                ry(nUse) = ToDbl(vals(r, 2))
                For f = 1 To nF
                    rz(nUse, f) = ToDbl(vals(r, 2 + f))
                Next f
            End If
        End If
    Next r

    If nUse = 0 Then
        ClearMesh
        InterpSay "No usable mesh point was found in '" & SheetName & "'!" & _
            RangeAddress & "." & vbCrLf & vbCrLf & "Expected X | Y | Value, with " & _
            "numeric X and Y, as the very first row of the range." & vbCrLf & "(" & _
            nBad & " row(s) skipped as non-numeric - InterpShowLog says why.)", _
            vbExclamation
        Exit Sub
    End If

    ' ---- 3. drop duplicated points (the FIRST one wins) --------------------
    minX = rx(1): maxX = rx(1): minY = ry(1): maxY = ry(1)
    For i = 2 To nUse
        If rx(i) < minX Then minX = rx(i)
        If rx(i) > maxX Then maxX = rx(i)
        If ry(i) < minY Then minY = ry(i)
        If ry(i) > maxY Then maxY = ry(i)
    Next i
    span = maxX - minX
    If maxY - minY > span Then span = maxY - minY
    dupTol = span * DUP_EPS_REL

    mDupMsgs = 0
    ReDim fx(1 To nUse)
    ReDim fy(1 To nUse)
    ReDim fz(1 To nUse, 1 To nF)
    For i = 1 To nUse
        dup = False
        For j = 1 To nOut
            If Abs(fx(j) - rx(i)) <= dupTol And Abs(fy(j) - ry(i)) <= dupTol Then
                dup = True
                Exit For
            End If
        Next j

        If dup Then
            nDup = nDup + 1
            mDupMsgs = mDupMsgs + 1
            If mDupMsgs <= DETAIL_LIMIT Then
                InterpLogMsg "ReadMeshPoints: SKIPPED duplicated mesh point at (" & _
                             NumText(rx(i)) & ", " & NumText(ry(i)) & ") - the FIRST " & _
                             "copy keeps its value(s)."
            ElseIf mDupMsgs = DETAIL_LIMIT + 1 Then
                InterpLogMsg "ReadMeshPoints: further duplicated points are counted " & _
                             "but not listed one by one."
            End If
        Else
            nOut = nOut + 1
            fx(nOut) = rx(i)
            fy(nOut) = ry(i)
            For f = 1 To nF
                fz(nOut, f) = rz(i, f)
            Next f
        End If
    Next i

    ' ---- 4. store the mesh in this module ----------------------------------
    mN = nOut
    mNFields = nF
    ReDim mX(1 To mN)
    ReDim mY(1 To mN)
    ReDim mZ(1 To mN, 1 To mNFields)
    For i = 1 To mN
        mX(i) = fx(i)
        mY(i) = fy(i)
        For f = 1 To mNFields
            mZ(i, f) = fz(i, f)
        Next f
    Next i

    InterpLogMsg "ReadMeshPoints: " & mN & " mesh point(s) x " & mNFields & _
                 " value column(s) read from '" & SheetName & "'!" & RangeAddress & _
                 " (" & nBad & " non-numeric row(s) skipped, " & nDup & _
                 " duplicate point(s) dropped)."

    ' ---- 5. triangulate once ----------------------------------------------
    BuildTriangulation

    ' ---- 6. a new mesh invalidates the values already interpolated ----------
    If mNS > 0 Then
        InterpLogMsg "ReadMeshPoints: re-interpolating the " & mNS & _
                     " sample point(s) already loaded against the new mesh."
        BuildInterpolatedTable
    End If
End Sub

' ===========================================================================
' 2. INTERPOLATE - Interpolate(X, Y, [FieldIndex])
' ===========================================================================
' The value at ONE point, by linear interpolation on the triangle that contains
' it. FieldIndex picks the value column of a multi-column mesh (1 = the first
' Value column, which is the whole story for an X | Y | Value mesh).
'
' OUTSIDE the mesh the NEAREST mesh point's value is returned (flat
' extrapolation) - use InterpolateEx if the caller needs to know that happened.
Public Function Interpolate( _
    ByVal X As Double, _
    ByVal Y As Double, _
    Optional ByVal FieldIndex As Long = 1) As Double

    Dim outside As Boolean

    Interpolate = InterpolateEx(X, Y, FieldIndex, outside)
End Function

' The same interpolation, plus OUT: OutsideMesh = True when the point was not
' inside any triangle and the nearest mesh point's value was substituted.
Public Function InterpolateEx( _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal FieldIndex As Long, _
    ByRef OutsideMesh As Boolean) As Double

    Dim t As Long
    Dim w1 As Double, w2 As Double, w3 As Double
    Dim f As Long
    Dim d As Double

    OutsideMesh = False

    If mN < 1 Then
        InterpSay "Interpolate: there is no mesh to interpolate from - read the " & _
            "mesh point table first (ReadMeshPoints). 0 is returned.", vbExclamation
        InterpolateEx = 0#
        Exit Function
    End If

    f = FieldIndex
    If f < 1 Then f = 1
    If f > mNFields Then
        InterpLogMsg "Interpolate: value column " & FieldIndex & " was asked for but " & _
                     "the mesh has only " & mNFields & " - column " & mNFields & _
                     " is used instead."
        f = mNFields
    End If

    ' ---- 1. inside the mesh: blend the containing triangle's corners -------
    If mNTri > 0 Then
        t = LocateTriangle(X, Y, w1, w2, w3)
        If t > 0 Then
            InterpolateEx = w1 * mZ(mTri(t, 1), f) _
                          + w2 * mZ(mTri(t, 2), f) _
                          + w3 * mZ(mTri(t, 3), f)
            Exit Function
        End If
    End If

    ' ---- 2. outside the mesh (or no triangle exists): nearest point --------
    OutsideMesh = True
    t = NearestPoint(X, Y, d)
    InterpolateEx = mZ(t, f)

    mOutMsgs = mOutMsgs + 1
    If mOutMsgs <= DETAIL_LIMIT Then
        InterpLogMsg "Interpolate: (" & NumText(X) & ", " & NumText(Y) & ") is outside " & _
                     "the mesh - the value of the NEAREST mesh point (" & NumText(mX(t)) & _
                     ", " & NumText(mY(t)) & ", " & NumText(d) & " away) was used."
    ElseIf mOutMsgs = DETAIL_LIMIT + 1 Then
        InterpLogMsg "Interpolate: further points outside the mesh are counted but " & _
                     "not listed one by one."
    End If
End Function

' ===========================================================================
' 3. INPUT - ReadSamplePoints(SheetName, RangeAddress)
' ===========================================================================
' Reads the sample (query) point table - Name | X | Y, or X | Y when the range
' has exactly two columns - into this module and, in the same call, runs every
' one of them through Interpolate to build the table of interpolated values.
' Reading stops at the first blank row, exactly as in ReadMeshPoints.
' That is why the MESH has to be read first; if it has not been, the points are
' still stored and the log says so, and BuildInterpolatedTable can fill the
' values in later.
Public Sub ReadSamplePoints( _
    ByVal SheetName As String, _
    ByVal RangeAddress As String)

    Dim vals As Variant
    Dim nRowsIn As Long, nColsIn As Long
    Dim r As Long, i As Long
    Dim nBad As Long, nUse As Long
    Dim allBlank As Boolean
    Dim named As Boolean
    Dim ixName As Long, ixX As Long, ixY As Long
    Dim ts() As String, tx() As Double, ty() As Double

    If Not ReadRange(SheetName, RangeAddress, "the SAMPLE point table", vals) Then Exit Sub

    nRowsIn = 0
    nColsIn = 0
    On Error Resume Next
    nRowsIn = UBound(vals, 1)
    nColsIn = UBound(vals, 2)
    On Error GoTo 0

    If nRowsIn < 1 Or nColsIn < 2 Then
        InterpSay "The sample range '" & SheetName & "'!" & RangeAddress & " is " & _
            nRowsIn & " row(s) x " & nColsIn & " column(s); two columns (X | Y) or " & _
            "three (Name | X | Y) are needed.", vbExclamation
        Exit Sub
    End If

    ' Two columns are X | Y; three or more are Name | X | Y, because a sample
    ' table that carries a label is how a person reads the output back.
    named = (nColsIn >= 3)
    If named Then
        ixName = 1
        ixX = 2
        ixY = 3
    Else
        ixName = 0
        ixX = 1
        ixY = 2
    End If

    If nColsIn > 3 Then
        InterpLogMsg "ReadSamplePoints: '" & SheetName & "'!" & RangeAddress & " has " & _
                     nColsIn & " columns - the name, X and Y are used and the " & _
                     (nColsIn - 3) & " further column(s) are ignored."
    End If

    ' ---- parse the rows (worst-case sizing, compacted at the end) ----------
    ReDim ts(1 To nRowsIn)
    ReDim tx(1 To nRowsIn)
    ReDim ty(1 To nRowsIn)

    For r = 1 To nRowsIn
        allBlank = True
        For i = 1 To nColsIn
            If Not IsBlankCell(vals(r, i)) Then
                allBlank = False
                Exit For
            End If
        Next i

        If allBlank Then
            InterpLogMsg "ReadSamplePoints: reading stopped at the blank row " & r & _
                         " of '" & SheetName & "'!" & RangeAddress & "."
            Exit For
        ElseIf Not (IsNumCell(vals(r, ixX)) And IsNumCell(vals(r, ixY))) Then
            nBad = nBad + 1
            InterpLogMsg "ReadSamplePoints: SKIPPED row " & r & " of '" & SheetName & _
                         "'!" & RangeAddress & " - X or Y is not a number ('" & _
                         SafeText(vals(r, ixX)) & "', '" & SafeText(vals(r, ixY)) & "')."
        Else
            nUse = nUse + 1
            If named Then ts(nUse) = SafeText(vals(r, ixName))
            tx(nUse) = ToDbl(vals(r, ixX))
            ty(nUse) = ToDbl(vals(r, ixY))
        End If
    Next r

    If nUse = 0 Then
        mNS = 0
        mSValuesReady = False
        InterpSay "No usable sample point was found in '" & SheetName & "'!" & _
            RangeAddress & "." & vbCrLf & vbCrLf & "Expected X | Y, or Name | X | Y, " & _
            "as the very first row of the range." & vbCrLf & "(" & nBad & _
            " row(s) skipped as non-numeric.)", vbExclamation
        Exit Sub
    End If

    ' ---- store the samples in this module ---------------------------------
    mNS = nUse
    mSHasName = named
    mSValuesReady = False
    ReDim mSName(1 To mNS)
    ReDim mSX(1 To mNS)
    ReDim mSY(1 To mNS)
    For i = 1 To mNS
        mSName(i) = ts(i)
        mSX(i) = tx(i)
        mSY(i) = ty(i)
    Next i

    InterpLogMsg "ReadSamplePoints: " & mNS & " sample point(s) read from '" & _
                 SheetName & "'!" & RangeAddress & " (" & nBad & _
                 " non-numeric row(s) skipped)."

    ' ---- and interpolate them straight away --------------------------------
    If mN < 1 Then
        InterpSay "The " & mNS & " sample point(s) were read, but there is NO mesh " & _
            "loaded yet, so their values could not be built." & vbCrLf & vbCrLf & _
            "Read the mesh point table first (ReadMeshPoints), then call " & _
            "BuildInterpolatedTable - or use RunInterpolation, which does all of it " & _
            "in the right order.", vbExclamation
        Exit Sub
    End If

    mOutMsgs = 0
    BuildInterpolatedTable
End Sub

' ===========================================================================
' WORKER - BuildInterpolatedTable
' ===========================================================================
' Runs EVERY loaded sample point through Interpolate and keeps the results in
' this module, one column per mesh value column. ReadSamplePoints calls it
' automatically; call it again by hand after reading a new mesh, or to rebuild
' the values if the mesh state changed underneath them.
Public Sub BuildInterpolatedTable()

    Dim s As Long, f As Long
    Dim outsideNow As Boolean
    Dim nOutside As Long

    If mNS < 1 Then
        InterpSay "BuildInterpolatedTable: no sample point is loaded - read the " & _
            "sample point table first (ReadSamplePoints).", vbExclamation
        Exit Sub
    End If

    If mN < 1 Then
        InterpSay "BuildInterpolatedTable: no mesh point is loaded - read the mesh " & _
            "point table first (ReadMeshPoints).", vbExclamation
        Exit Sub
    End If

    mOutMsgs = 0
    ReDim mSV(1 To mNS, 1 To mNFields)
    ReDim mSOut(1 To mNS)

    For s = 1 To mNS
        For f = 1 To mNFields
            outsideNow = False
            mSV(s, f) = InterpolateEx(mSX(s), mSY(s), f, outsideNow)
            If outsideNow Then mSOut(s) = True
        Next f
        If mSOut(s) Then nOutside = nOutside + 1
    Next s

    mSValuesReady = True
    InterpLogMsg "BuildInterpolatedTable: " & mNS & " sample point(s) x " & _
                 mNFields & " value column(s) interpolated on " & mNTri & _
                 " triangle(s); " & nOutside & " point(s) fell OUTSIDE the mesh and " & _
                 "used the nearest mesh point."
End Sub

' ===========================================================================
' 4. OUTPUT - WriteSampleTable(SheetName, StartCell, [WriteName])
' ===========================================================================
' Prints the sample point table WITH its interpolated values, flush on the
' top-left cell given: no title row, no header row. The columns are
'
'     Name | X | Y | V1 | V2 | ...      (Name only when the sample table had one
'                                        and WriteName is True)
'
' The sheet is created (at the end of the workbook) when it does not exist yet.
Public Sub WriteSampleTable( _
    ByVal SheetName As String, _
    ByVal StartCell As String, _
    Optional ByVal WriteName As Boolean = True)

    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim nCols As Long, nOffset As Long
    Dim s As Long, f As Long
    Dim data As Variant

    If mNS < 1 Then
        InterpSay "WriteSampleTable: no sample point is loaded - nothing to write.", _
            vbExclamation
        Exit Sub
    End If

    If Not mSValuesReady Then
        InterpSay "WriteSampleTable: the sample point values have NOT been built " & _
            "yet, so there is nothing to write." & vbCrLf & vbCrLf & _
            "Call ReadSamplePoints (which builds them) or BuildInterpolatedTable " & _
            "first.", vbExclamation
        Exit Sub
    End If

    Set ws = WriteSheet(SheetName)
    If ws Is Nothing Then
        InterpSay "WriteSampleTable: worksheet '" & SheetName & "' could not be " & _
            "opened for writing.", vbExclamation
        Exit Sub
    End If

    If Not ParseCell(ws, StartCell, r, c) Then
        InterpSay "WriteSampleTable: '" & StartCell & "' is not a usable top-left " & _
            "cell on sheet '" & SheetName & "'.", vbExclamation
        Exit Sub
    End If

    ' ---- 1. shape the block ------------------------------------------------
    nCols = 2 + mNFields
    nOffset = 0
    If WriteName Then
        If mSHasName Then
            nCols = nCols + 1
            nOffset = 1
        Else
            InterpLogMsg "WriteSampleTable: WriteName:=True, but the sample table " & _
                         "has no name column - X | Y | values are written."
        End If
    End If

    ' ---- 2. fill it --------------------------------------------------------
    ReDim data(1 To mNS, 1 To nCols)
    For s = 1 To mNS
        If nOffset = 1 Then data(s, 1) = mSName(s)
        data(s, 1 + nOffset) = mSX(s)
        data(s, 2 + nOffset) = mSY(s)
        For f = 1 To mNFields
            data(s, 2 + nOffset + f) = mSV(s, f)
        Next f
    Next s

    ' ---- 3. write it in one go ---------------------------------------------
    On Error GoTo Fail
    ws.Cells(r, c).Resize(mNS, nCols).Value = data
    On Error GoTo 0

    InterpLogMsg "WriteSampleTable: " & mNS & " row(s) x " & nCols & " column(s) " & _
                 "written to '" & ws.Name & "' at " & StartCell & " (no title, no header)."
    Exit Sub

Fail:
    InterpSay "WriteSampleTable: writing to '" & SheetName & "' at " & StartCell & _
        " failed:" & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ===========================================================================
' 5. WRAPPER - RunInterpolation(...)
' ===========================================================================
' The whole job in one call: read the mesh point table, read the sample point
' table (which interpolates it), and print the result. Every table is given by
' sheet name + range, so nothing is hardcoded here - a site-specific entry point
' only has to hold the constants (see InterpolationDemo at the bottom of this
' section).
'
'   MeshSheet     sheet holding the mesh point table (X | Y | Value...)
'   MeshRange     the range on that sheet, e.g. "A2:C2000"
'   SampleSheet   sheet holding the sample point table (Name | X | Y)
'   SampleRange   the range on that sheet
'   OutSheet      sheet the results go to (created when it does not exist)
'   OutCell       top-left cell of the result block, e.g. "B3"
'   WriteName     write the sample point's name as the first output column
Public Sub RunInterpolation( _
    ByVal MeshSheet As String, _
    ByVal MeshRange As String, _
    ByVal SampleSheet As String, _
    ByVal SampleRange As String, _
    ByVal OutSheet As String, _
    ByVal OutCell As String, _
    Optional ByVal WriteName As Boolean = True)

    InterpLogMsg "RunInterpolation: BEGIN - mesh '" & MeshSheet & "'!" & MeshRange & _
                 ", samples '" & SampleSheet & "'!" & SampleRange & " -> '" & _
                 OutSheet & "' at " & OutCell & "."

    ' ---- 1. the mesh (interpolation is impossible without it) --------------
    ReadMeshPoints MeshSheet, MeshRange
    If mN < 1 Then
        InterpSay "RunInterpolation stopped: no usable mesh point was read, so " & _
            "nothing was interpolated and nothing was written." & vbCrLf & vbCrLf & _
            "InterpShowLog has the full log.", vbExclamation
        Exit Sub
    End If

    ' ---- 2. the sample points - this also builds their values -------------
    ReadSamplePoints SampleSheet, SampleRange
    If mNS < 1 Then
        InterpSay "RunInterpolation stopped: no usable sample point was read, so " & _
            "nothing was written." & vbCrLf & vbCrLf & "InterpShowLog has the full log.", _
            vbExclamation
        Exit Sub
    End If

    ' ---- 3. print the result ----------------------------------------------
    WriteSampleTable OutSheet, OutCell, WriteName
    If Not mSValuesReady Then Exit Sub

    ' ---- 4. report --------------------------------------------------------
    InterpLogMsg "RunInterpolation: DONE - " & mN & " mesh point(s) / " & mNTri & _
                 " triangle(s), " & mNFields & " value column(s), " & mNS & _
                 " sample point(s) -> '" & OutSheet & "' starting at " & OutCell & "."
    InterpLogMsg "   outside the mesh (nearest point used) for " & OutsideCount() & _
                 " of the " & mNS & " sample point(s) - see SampleUsedNearest for " & _
                 "the exact rows."
End Sub

' ===========================================================================
' ENTRY POINT - run this one from Excel (Alt+F8)
' ===========================================================================
' All the site-specific values live here as LOCAL constants - this is the ONLY
' place to edit, and the wrapper below stays reusable.
Public Sub InterpolationDemo()
    ' ---- EDIT THESE --------------------------------------------------------
    Const MESH_SHEET As String = "Mesh Points"       ' <-- EDIT ME
    Const MESH_RANGE As String = "A2:C200"           ' <-- EDIT ME (X | Y | Value...)

    Const SAMPLE_SHEET As String = "Sample Points"   ' <-- EDIT ME
    Const SAMPLE_RANGE As String = "A2:C500"         ' <-- EDIT ME (Name | X | Y)

    Const OUT_SHEET As String = "Interpolated"       ' <-- EDIT ME
    Const OUT_CELL As String = "A1"                  ' <-- EDIT ME (top-left cell)

    Const WRITE_NAME As Boolean = True               ' False = Name | X | Y -> X | Y
    ' ------------------------------------------------------------------------

    RunInterpolation MESH_SHEET, MESH_RANGE, _
                     SAMPLE_SHEET, SAMPLE_RANGE, _
                     OUT_SHEET, OUT_CELL, WRITE_NAME

    If InterpMsgBoxLogging Then InterpShowLog
End Sub

' ===========================================================================
' STATE ACCESSORS - for a caller in another module, or the Immediate window
' ===========================================================================

Public Function MeshIsLoaded() As Boolean
    MeshIsLoaded = (mN > 0)
End Function

Public Function MeshPointCount() As Long
    MeshPointCount = mN
End Function

Public Function MeshTriangleCount() As Long
    MeshTriangleCount = mNTri
End Function

Public Function MeshValueCount() As Long
    MeshValueCount = mNFields
End Function

Public Function SamplePointCount() As Long
    SamplePointCount = mNS
End Function

' True when sample row RowIndex fell outside the mesh, i.e. its value is the
' NEAREST mesh point's value rather than an interpolation.
Public Function SampleUsedNearest(ByVal RowIndex As Long) As Boolean
    If Not mSValuesReady Then Exit Function
    If RowIndex < 1 Or RowIndex > mNS Then Exit Function
    SampleUsedNearest = mSOut(RowIndex)
End Function

' The interpolated value of sample row RowIndex, value column FieldIndex.
Public Function SampleValueAt(ByVal RowIndex As Long, ByVal FieldIndex As Long) As Double
    If Not mSValuesReady Then Exit Function
    If RowIndex < 1 Or RowIndex > mNS Then Exit Function
    If FieldIndex < 1 Or FieldIndex > mNFields Then Exit Function
    SampleValueAt = mSV(RowIndex, FieldIndex)
End Function

' Number of sample rows that used the nearest mesh point.
Public Function OutsideCount() As Long
    Dim s As Long, k As Long

    If Not mSValuesReady Then Exit Function
    For s = 1 To mNS
        If mSOut(s) Then k = k + 1
    Next s
    OutsideCount = k
End Function

' FORGET the mesh, the samples and the triangulation. The log is kept - clear
' that with InterpClearLog.
Public Sub ClearInterpState()
    ClearMesh
    mNS = 0
    mSHasName = False
    mSValuesReady = False
    Erase mSName
    Erase mSX
    Erase mSY
    Erase mSV
    Erase mSOut
    InterpLogMsg "ClearInterpState: mesh, triangulation and sample points dropped."
End Sub

Private Sub ClearMesh()
    mN = 0
    mNFields = 0
    mNTri = 0
    Erase mX
    Erase mY
    Erase mZ
    Erase mTri

    ' No mesh means no values: the sample POINTS stay (they can be re-interpolated
    ' once a mesh is loaded again) but their values are marked as not built, so a
    ' writer that runs afterwards refuses instead of printing stale numbers.
    mSValuesReady = False
    Erase mSV
    Erase mSOut
End Sub

' ===========================================================================
' TRIANGULATION - Bowyer-Watson incremental Delaunay
' ===========================================================================
' One super triangle encloses every mesh point; the points are then inserted one
' at a time. Inserting a point means deleting every triangle whose circumcircle
' contains it and re-filling the hole with triangles from the point to the
' HOLE'S BOUNDARY - the edges that belonged to exactly one deleted triangle
' (an edge shared by two deleted triangles is inside the hole and disappears).
' At the end the triangles that still touch a super corner are dropped, and what
' is left is the Delaunay triangulation of the mesh points.
'
' The interior boundary edges are found by cancelling pairs, which is why the
' edge list is a plain unsorted array: with a handful of hundred points the
' linear scan is faster than any cleverer structure, and it cannot go wrong.
Private Sub BuildTriangulation()

    Dim n As Long
    Dim ex() As Double, ey() As Double
    Dim ta() As Long, tb() As Long, tc() As Long
    Dim capTri As Long, nTri As Long
    Dim bad() As Boolean
    Dim ea() As Long, eb() As Long, capEdge As Long, nEdge As Long
    Dim i As Long, t As Long, j As Long, e As Long, w As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim midX As Double, midY As Double, dmax As Double
    Dim a As Long, b As Long, tmp As Long
    Dim cor(0 To 2) As Long
    Dim found As Boolean

    Erase mTri
    mNTri = 0
    n = mN

    If n < 3 Then
        If n > 0 Then
            InterpLogMsg "BuildTriangulation: only " & n & " mesh point(s), so no " & _
                         "triangle can be formed - every query falls back to the " & _
                         "NEAREST mesh point."
        End If
        Exit Sub
    End If

    ' ---- 1. the point list plus the three super-triangle corners -----------
    ReDim ex(1 To n + 3)
    ReDim ey(1 To n + 3)
    minX = mX(1): maxX = mX(1): minY = mY(1): maxY = mY(1)
    For i = 1 To n
        ex(i) = mX(i)
        ey(i) = mY(i)
        If mX(i) < minX Then minX = mX(i)
        If mX(i) > maxX Then maxX = mX(i)
        If mY(i) < minY Then minY = mY(i)
        If mY(i) > maxY Then maxY = mY(i)
    Next i

    midX = (minX + maxX) / 2#
    midY = (minY + maxY) / 2#
    dmax = maxX - minX
    If maxY - minY > dmax Then dmax = maxY - minY
    If dmax <= 0# Then dmax = 1#              ' all points identical (guarded anyway)

    ' A wide super triangle: far enough away that no real circumcircle reaches a
    ' corner, but not so far that its own circumcircle test loses precision.
    ex(n + 1) = midX - 20# * dmax: ey(n + 1) = midY - dmax
    ex(n + 2) = midX: ey(n + 2) = midY + 20# * dmax
    ex(n + 3) = midX + 20# * dmax: ey(n + 3) = midY - dmax

    ' ---- 2. insert the points one by one ----------------------------------
    ' A triangulation of k points has at most 2k - 5 triangles, and inserting one
    ' point adds at most two, so 2n + 16 is a safe starting capacity (it is grown
    ' below if a degenerate input still manages to exceed it).
    capTri = 2 * n + 16
    ReDim ta(1 To capTri)
    ReDim tb(1 To capTri)
    ReDim tc(1 To capTri)
    ReDim bad(1 To capTri)
    nTri = 1
    ta(1) = n + 1: tb(1) = n + 2: tc(1) = n + 3

    capEdge = 3 * capTri
    ReDim ea(1 To capEdge)
    ReDim eb(1 To capEdge)

    For i = 1 To n
        ' --- 2a. which triangles have point i inside their circumcircle? ----
        nEdge = 0
        For t = 1 To nTri
            bad(t) = InCircle(ex(ta(t)), ey(ta(t)), ex(tb(t)), ey(tb(t)), _
                              ex(tc(t)), ey(tc(t)), ex(i), ey(i))
            If bad(t) Then
                cor(0) = ta(t): cor(1) = tb(t): cor(2) = tc(t)
                For e = 0 To 2
                    a = cor(e)
                    b = cor((e + 1) Mod 3)
                    If a > b Then
                        tmp = a: a = b: b = tmp
                    End If

                    found = False
                    For j = 1 To nEdge
                        If ea(j) = a And eb(j) = b Then
                            ' the SECOND bad triangle owning this edge proves the
                            ' edge is interior to the hole - cancel both
                            ea(j) = ea(nEdge)
                            eb(j) = eb(nEdge)
                            nEdge = nEdge - 1
                            found = True
                            Exit For
                        End If
                    Next j

                    If Not found Then
                        nEdge = nEdge + 1
                        ea(nEdge) = a
                        eb(nEdge) = b
                    End If
                Next e
            End If
        Next t

        ' --- 2b. drop the bad triangles (in-place compaction) ---------------
        w = 0
        For t = 1 To nTri
            If Not bad(t) Then
                w = w + 1
                ta(w) = ta(t): tb(w) = tb(t): tc(w) = tc(t)
            End If
        Next t
        nTri = w

        ' --- 2c. re-fill the hole from the surviving boundary edges ---------
        If nTri + nEdge > capTri Then
            capTri = (nTri + nEdge) * 2
            capEdge = 3 * capTri
            ReDim Preserve ta(1 To capTri)
            ReDim Preserve tb(1 To capTri)
            ReDim Preserve tc(1 To capTri)
            ReDim Preserve bad(1 To capTri)
            ReDim Preserve ea(1 To capEdge)
            ReDim Preserve eb(1 To capEdge)
        End If

        For j = 1 To nEdge
            nTri = nTri + 1
            ta(nTri) = ea(j)
            tb(nTri) = eb(j)
            tc(nTri) = i
        Next j
    Next i

    ' ---- 3. drop every triangle that still uses a super corner -------------
    w = 0
    For t = 1 To nTri
        If ta(t) <= n And tb(t) <= n And tc(t) <= n Then
            w = w + 1
            ta(w) = ta(t): tb(w) = tb(t): tc(w) = tc(t)
        End If
    Next t
    nTri = w

    If nTri = 0 Then
        InterpLogMsg "BuildTriangulation: no triangle survived - the " & n & _
                     " mesh point(s) are all coincident or collinear. Every query " & _
                     "falls back to the NEAREST mesh point."
        Exit Sub
    End If

    mNTri = nTri
    ReDim mTri(1 To mNTri, 1 To 3)
    For t = 1 To mNTri
        mTri(t, 1) = ta(t)
        mTri(t, 2) = tb(t)
        mTri(t, 3) = tc(t)
    Next t

    InterpLogMsg "BuildTriangulation: " & n & " mesh point(s) -> " & mNTri & _
                 " triangle(s)."
End Sub

' Twice the signed area of the triangle a-b-c. Positive when a->b->c turns
' counter-clockwise. Everything geometric in this module is built on it.
Private Function Orient2( _
    ByVal ax As Double, ByVal ay As Double, _
    ByVal bx As Double, ByVal by As Double, _
    ByVal cx As Double, ByVal cy As Double) As Double

    Orient2 = (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)
End Function

' Is point d inside the circumcircle of a-b-c? The 4x4 determinant holds
' |d - a|^2 and friends, and the answer has to be flipped when a-b-c is
' clockwise - hence the orientation test in front of it. Collinear corners have
' no meaningful circumcircle, so they answer False rather than a division by ~0.
Private Function InCircle( _
    ByVal ax As Double, ByVal ay As Double, _
    ByVal bx As Double, ByVal by As Double, _
    ByVal cx As Double, ByVal cy As Double, _
    ByVal dx As Double, ByVal dy As Double) As Boolean

    Dim o As Double, det As Double
    Dim adx As Double, ady As Double, bdx As Double, bdy As Double
    Dim cdx As Double, cdy As Double
    Dim alift As Double, blift As Double, clift As Double
    Dim s As Double

    s = TriScale(ax, ay, bx, by, cx, cy)
    o = Orient2(ax, ay, bx, by, cx, cy)
    If Abs(o) <= GeomTol(s, 2) Then Exit Function     ' collinear: not "inside"

    adx = ax - dx: ady = ay - dy
    bdx = bx - dx: bdy = by - dy
    cdx = cx - dx: cdy = cy - dy
    alift = adx * adx + ady * ady
    blift = bdx * bdx + bdy * bdy
    clift = cdx * cdx + cdy * cdy

    det = adx * (bdy * clift - blift * cdy) _
        - ady * (bdx * clift - blift * cdx) _
        + alift * (bdx * cdy - bdy * cdx)

    If o > 0# Then
        InCircle = (det > GeomTol(s, 4))
    Else
        InCircle = (det < -GeomTol(s, 4))
    End If
End Function

' The size of the coordinates a test is working with: the largest span between
' the corners, but never smaller than the distance from the origin, because with
' coordinates far from the origin the DIFFERENCES above carry less precision.
Private Function TriScale( _
    ByVal ax As Double, ByVal ay As Double, _
    ByVal bx As Double, ByVal by As Double, _
    ByVal cx As Double, ByVal cy As Double) As Double

    Dim s As Double

    s = Abs(bx - ax)
    If Abs(by - ay) > s Then s = Abs(by - ay)
    If Abs(cx - ax) > s Then s = Abs(cx - ax)
    If Abs(cy - ay) > s Then s = Abs(cy - ay)
    If Abs(ax) > s Then s = Abs(ax)
    If Abs(ay) > s Then s = Abs(ay)
    If Abs(bx) > s Then s = Abs(bx)
    If Abs(by) > s Then s = Abs(by)
    If Abs(cx) > s Then s = Abs(cx)
    If Abs(cy) > s Then s = Abs(cy)

    TriScale = s
End Function

' The tolerance for a determinant whose terms are of order Scale^Power:
' GEOM_EPS_REL relative to it. Clamped at both ends so a pathological Scale
' (0, or something close to the Double limit) cannot produce 0 or an overflow -
' a comparison tolerance on an Infinity is not a tolerance.
Private Function GeomTol(ByVal Scale As Double, ByVal Power As Long) As Double

    Dim t As Double
    Dim i As Long

    If Scale <= 0# Then Scale = 1#
    If Scale > 1E+100 Then Scale = 1E+100

    t = GEOM_EPS_REL
    For i = 1 To Power
        t = t * Scale
        If t > 1E+300 Then t = 1E+300
    Next i

    If t <= 0# Then t = 1E-300
    GeomTol = t
End Function

' Which triangle contains (px, py)? 0 = none of them, and then the weights are
' the barycentric coordinates of the point inside that triangle. BARY_EPS lets a
' point sitting exactly ON an edge (the common case of a sample point landing on
' a mesh node) count as inside; when several triangles claim it, the one the
' point is DEEPEST inside wins, so the choice does not depend on triangle order.
Private Function LocateTriangle( _
    ByVal px As Double, _
    ByVal py As Double, _
    ByRef w1 As Double, _
    ByRef w2 As Double, _
    ByRef w3 As Double) As Long

    Dim t As Long, i1 As Long, i2 As Long, i3 As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim d As Double, a1 As Double, a2 As Double, a3 As Double, m As Double
    Dim best As Long, bestMin As Double

    LocateTriangle = 0
    bestMin = -1#
    w1 = 0#: w2 = 0#: w3 = 0#

    For t = 1 To mNTri
        i1 = mTri(t, 1): i2 = mTri(t, 2): i3 = mTri(t, 3)
        x1 = mX(i1): y1 = mY(i1)
        x2 = mX(i2): y2 = mY(i2)
        x3 = mX(i3): y3 = mY(i3)

        d = Orient2(x1, y1, x2, y2, x3, y3)
        If Abs(d) > GeomTol(TriScale(x1, y1, x2, y2, x3, y3), 2) Then
            ' weights are the three sub-areas over the whole area, so they sum to
            ' 1 for ANY point - "inside" is simply "none of them negative"
            a1 = Orient2(px, py, x2, y2, x3, y3) / d
            a2 = Orient2(x1, y1, px, py, x3, y3) / d
            a3 = Orient2(x1, y1, x2, y2, px, py) / d

            m = a1
            If a2 < m Then m = a2
            If a3 < m Then m = a3

            If m >= -BARY_EPS Then
                If m > bestMin Then
                    bestMin = m
                    best = t
                    w1 = a1: w2 = a2: w3 = a3
                End If
                If m > BARY_EPS Then Exit For   ' clearly inside - done
            End If
        End If
    Next t

    LocateTriangle = best
End Function

' Index of the mesh point closest to (px, py); 0 when there is no mesh.
' OutDist comes back with that distance.
Private Function NearestPoint( _
    ByVal px As Double, _
    ByVal py As Double, _
    Optional ByRef OutDist As Double = 0#) As Long

    Dim i As Long
    Dim dx As Double, dy As Double, d2 As Double, bestD2 As Double

    NearestPoint = 0
    OutDist = 0#
    If mN < 1 Then Exit Function

    bestD2 = -1#
    For i = 1 To mN
        dx = mX(i) - px
        dy = mY(i) - py
        d2 = dx * dx + dy * dy
        If bestD2 < 0# Or d2 < bestD2 Then
            bestD2 = d2
            NearestPoint = i
        End If
    Next i

    OutDist = Sqr(bestD2)
End Function

' ===========================================================================
' SHEET HELPERS - reading, writing, cell parsing
' ===========================================================================

' Reads one range into a 1-based 2-D Variant array. False = already reported,
' and the caller must stop. ForWhat only names the table in the message.
Private Function ReadRange( _
    ByVal SheetName As String, _
    ByVal RangeAddress As String, _
    ByVal ForWhat As String, _
    ByRef vals As Variant) As Boolean

    Dim ws As Worksheet
    Dim rng As Range
    Dim errNum As Long

    ReadRange = False
    vals = Empty

    On Error Resume Next
    Err.Clear
    Set ws = ThisWorkbook.Worksheets(SheetName)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Or ws Is Nothing Then
        InterpSay "Worksheet '" & SheetName & "' was not found in this workbook " & _
            "(" & ForWhat & ")." & vbCrLf & vbCrLf & "Check the sheet name that was " & _
            "passed in - the constants of the wrapper sub that called this, or the " & _
            "arguments of RunInterpolation.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set rng = ws.Range(RangeAddress)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Or rng Is Nothing Then
        InterpSay "Range '" & RangeAddress & "' could not be resolved on sheet '" & _
            SheetName & "' (" & ForWhat & ")." & vbCrLf & vbCrLf & "Check the range " & _
            "address that was passed in.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    vals = rng.Value
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        InterpSay "Range '" & SheetName & "'!" & RangeAddress & " could not be read (" & _
            ForWhat & "):" & vbCrLf & vbCrLf & errNum & " - " & Error$(errNum), _
            vbExclamation
        Exit Function
    End If

    ' A single cell comes back as a bare value, not an array: that can never be
    ' one of the tables this module reads, so it is an error, not a special case.
    If Not IsArray(vals) Then
        InterpSay "'" & SheetName & "'!" & RangeAddress & " is a single cell (" & _
            ForWhat & ")." & vbCrLf & vbCrLf & "A range of at least one row by two " & _
            "(or three) columns is needed.", vbExclamation
        Exit Function
    End If

    ReadRange = True
End Function

' The sheet to WRITE to, created at the end of the workbook when it is missing.
' Nothing is returned (and Set returns Nothing) when even that fails.
Private Function WriteSheet(ByVal SheetName As String) As Worksheet

    Dim errNum As Long

    On Error Resume Next
    Err.Clear
    Set WriteSheet = ThisWorkbook.Worksheets(SheetName)
    errNum = Err.Number
    On Error GoTo 0
    If errNum = 0 And Not WriteSheet Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set WriteSheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    WriteSheet.Name = SheetName
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        InterpLogMsg "WriteSheet: could not create worksheet '" & SheetName & "': " & _
                     errNum & " - " & Error$(errNum)
        Set WriteSheet = Nothing
    End If
End Function

' Resolves an "A1" style address against a sheet into row / column numbers.
Private Function ParseCell( _
    ByVal ws As Worksheet, _
    ByVal StartCell As String, _
    ByRef r As Long, _
    ByRef c As Long) As Boolean

    Dim rng As Range
    Dim errNum As Long

    ParseCell = False
    r = 0
    c = 0

    On Error Resume Next
    Err.Clear
    Set rng = ws.Range(StartCell)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Or rng Is Nothing Then Exit Function

    r = rng.Row
    c = rng.Column
    ParseCell = True
End Function

' ===========================================================================
' CELL / NUMBER HELPERS
' ===========================================================================

' Cell value as text: Empty becomes "", a worksheet error value becomes "" (its
' CStr would raise), everything else is trimmed.
Private Function SafeText(ByVal v As Variant) As String
    SafeText = ""
    On Error Resume Next
    If Not IsEmpty(v) Then SafeText = Trim$(CStr(v))
End Function

' A cell with nothing in it - Empty, or the "" a formula returns.
Private Function IsBlankCell(ByVal v As Variant) As Boolean
    IsBlankCell = (Len(SafeText(v)) = 0)
End Function

' IsNumeric, guarded: a cell holding #N/A, #DIV/0! ... arrives as a Variant/Error
' and raises rather than answering.
Private Function IsNumCell(ByVal v As Variant) As Boolean

    Dim errNum As Long

    IsNumCell = False
    On Error Resume Next
    Err.Clear
    IsNumCell = IsNumeric(v)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then IsNumCell = False
End Function

' The numeric value of a cell. A numeric cell is a Double already; a text cell
' that merely LOOKS numeric falls back to Val, which reads "." as the decimal
' separator whatever the Windows locale is (CDbl on a String would follow the
' locale and fail on "10.3" under a comma-decimal locale).
Private Function ToDbl(ByVal v As Variant) As Double

    Dim errNum As Long

    ToDbl = 0#
    If IsEmpty(v) Then Exit Function

    On Error Resume Next
    Err.Clear
    ToDbl = CDbl(v)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then ToDbl = Val(SafeText(v))
End Function

' A number as text for the log: "." as the decimal separator, whatever the
' Windows locale is, and no scientific notation for ordinary values.
Private Function NumText(ByVal v As Double) As String
    NumText = Replace$(Format$(v, "0.############"), ",", ".")
End Function
