Attribute VB_Name = "modManning"
Option Explicit

'============================================================
' modManning (Optimised, same functionality)
'============================================================

'--- existing globals (kept) ---
Public vGoAll As Boolean
Public vStopAll As Boolean

Public arrLookahead As Variant
Public LookaheadColRange As String
Public LookaheadRowRange As String
Public LookaheadSheetName As String

'============================================================
' App state guard for StopAll/GoAll
'============================================================
Private Type TAppState
    Calc As XlCalculation
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayStatusBar As Boolean
End Type

Private mSaved As TAppState
Private mHasSaved As Boolean

'============================================================
' Small helpers
'============================================================
Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = vbNullString
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function DateKey(ByVal d As Date) As String
    DateKey = Format$(d, "yyyymmdd")
End Function

Public Function FormatDate(v As Variant) As String
    'kept signature
    FormatDate = Format$(CDate(v), "yyyymmdd")
End Function

Private Function ToDateFloor(ByVal v As Variant) As Date
    'strip time
    ToDateFloor = DateSerial(Year(CDate(v)), Month(CDate(v)), Day(CDate(v)))
End Function

Private Function GetManningWS(Optional ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set GetManningWS = wb.Worksheets("Manning")
End Function

Private Function GetLookahead(Optional ByVal wb As Workbook) As ListObject
    Set GetLookahead = GetManningWS(wb).ListObjects("tblLookahead")
End Function

'============================================================
' CondFormats wrapper (your original referenced an unknown symbol)
'============================================================
Public Function GetCondFormats() As Boolean
    'Your original: GetCondFormats = CondFormats (undefined here)
    'Keeping signature, but make it deterministic:
    On Error Resume Next
    GetCondFormats = CondFormats
    On Error GoTo 0
End Function

'============================================================
' StopAll / GoAll (safe + restore)
'============================================================
Public Sub StopAll()
    If Not mHasSaved Then
        mSaved.Calc = Application.Calculation
        mSaved.ScreenUpdating = Application.ScreenUpdating
        mSaved.EnableEvents = Application.EnableEvents
        mSaved.DisplayStatusBar = Application.DisplayStatusBar
        mHasSaved = True
    End If

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    vStopAll = True
    vGoAll = False
End Sub

Public Sub GoAll()
    Dim ws As Worksheet
    On Error Resume Next

    If mHasSaved Then
        Application.Calculation = mSaved.Calc
        Application.ScreenUpdating = mSaved.ScreenUpdating
        Application.EnableEvents = mSaved.EnableEvents
        Application.DisplayStatusBar = mSaved.DisplayStatusBar
        mHasSaved = False
    Else
        'fallback
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
    End If

    'your preference switch
    Set ws = ThisWorkbook.Worksheets("Manning")
    If ws.Range("rngCalc").Value = True Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If

    vGoAll = True
    vStopAll = False
End Sub

'============================================================
' CheckLeaveAndFillColumn (optimised)
'============================================================
Public Sub CheckLeaveAndFillColumn()
    Dim wsLookahead As Worksheet, wsLeave As Worksheet
    Dim lo As ListObject
    Dim leavePID As Variant, leaveStart As Variant, leaveEnd As Variant
    Dim tblData As Variant

    Dim leaveDict As Object
    Dim i As Long, j As Long
    Dim pid As String

    Dim fDate As Date, eDate As Date
    Dim isOnLeave As Boolean
    Dim ranges As Collection
    Dim r As Variant

    On Error GoTo ErrorHandle

    If vStopAll = False Then StopAll
    Application.EnableEvents = False

    Set wsLookahead = ThisWorkbook.Worksheets("Manning")
    Set lo = wsLookahead.ListObjects("tblLookahead")

    Set wsLeave = ThisWorkbook.Worksheets("tbl_Vista_HR_Leave")
    leavePID = wsLeave.Range("rngLeavePID").Value2
    leaveStart = wsLeave.Range("rngLeaveSDate").Value2
    leaveEnd = wsLeave.Range("rngLeaveFDate").Value2

    If lo.DataBodyRange Is Nothing Then GoTo ExitSub
    tblData = lo.DataBodyRange.Value2

    fDate = ToDateFloor(wsLookahead.Range("rngManningFDate").Value)
    eDate = DateAdd("d", CLng(wsLookahead.Range("rngManningDateRangeTo").Value) - 1, fDate)

    'Build leave ranges per PID
    Set leaveDict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(leavePID, 1)
        If IsDate(leaveStart(i, 1)) And IsDate(leaveEnd(i, 1)) Then
            pid = UCase$(Trim$(NzStr(leavePID(i, 1))))
            If Len(pid) > 0 Then
                If Not leaveDict.Exists(pid) Then
                    leaveDict.Add pid, New Collection
                End If
                leaveDict(pid).Add Array(ToDateFloor(leaveStart(i, 1)), ToDateFloor(leaveEnd(i, 1)))
            End If
        End If
    Next i

    'Write output in-memory first
    Dim outArr() As Variant
    ReDim outArr(1 To UBound(tblData, 1), 1 To 1)

    For i = 1 To UBound(tblData, 1)
        pid = UCase$(Trim$(NzStr(tblData(i, 8)))) 'col 8 = PID
        isOnLeave = False

        If Len(pid) > 0 And leaveDict.Exists(pid) Then
            Set ranges = leaveDict(pid)
            For Each r In ranges
                'overlap test: [r(0), r(1)] overlaps [fDate, eDate]
                If r(0) <= eDate And r(1) >= fDate Then
                    isOnLeave = True
                    Exit For
                End If
            Next r
        End If

        outArr(i, 1) = isOnLeave
    Next i

    'Column 11 in table
    If lo.ListColumns.Count >= 11 Then
        lo.ListColumns(11).DataBodyRange.Value = outArr
    End If

ExitSub:
    If vGoAll = False Then GoAll
    Application.EnableEvents = True
    Exit Sub

ErrorHandle:
    MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical, "CheckLeaveAndFillColumn"
    Resume ExitSub
End Sub

Private Function TryHeaderDateKey(ByVal v As Variant, ByRef outKey As String) As Boolean
    'Returns True if v can be interpreted as a date; outputs yyyymmdd key.
    'Handles:
    '  - Excel serial (Value2 as Double)
    '  - real Date
    '  - date-like strings ("12/02/2026", etc.)

    On Error GoTo Fail

    If IsEmpty(v) Or IsNull(v) Then GoTo Fail

    '1) Excel serial (most common with Value2)
    If IsNumeric(v) Then
        Dim d As Date
        d = DateSerial(1899, 12, 30) + Int(CDbl(v)) 'floor time
        outKey = Format$(d, "yyyymmdd")
        TryHeaderDateKey = True
        Exit Function
    End If

    '2) Date variant or parsable string
    If IsDate(v) Then
        Dim d2 As Date
        d2 = DateSerial(Year(CDate(v)), Month(CDate(v)), Day(CDate(v))) 'floor time
        outKey = Format$(d2, "yyyymmdd")
        TryHeaderDateKey = True
        Exit Function
    End If

Fail:
    TryHeaderDateKey = False
End Function

Private Function BuildHeaderDict_FromTable(ByVal lo As ListObject, _
                                          ByVal firstCalAbsCol As Long) As Object
    'Maps yyyymmdd -> absolute worksheet column index
    Dim dict As Object
    Dim c As Long
    Dim v As Variant
    Dim k As String

    Set dict = CreateObject("Scripting.Dictionary")

    'Use table header row as source of truth
    For c = 1 To lo.HeaderRowRange.Columns.Count
        Dim absCol As Long
        absCol = lo.HeaderRowRange.Cells(1, c).Column

        'Only consider calendar columns (skip left meta columns)
        If absCol >= firstCalAbsCol Then
            v = lo.HeaderRowRange.Cells(1, c).Value2
            If TryHeaderDateKey(v, k) Then
                dict(k) = absCol
            End If
        End If
    Next c

    Set BuildHeaderDict_FromTable = dict
End Function


'============================================================
' CheckClashesAndFillColumn (optimised)
'============================================================
Public Sub CheckClashesAndFillColumn()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim data As Variant
    Dim i As Long, c As Long

    Dim fDate As Date, eDate As Date
    Dim startAbsCol As Long, endAbsCol As Long
    Dim tableStartAbsCol As Long
    Dim relStart As Long, relEnd As Long

    Dim headerDict As Object
    Dim counts As Object ' key: emp|relCol -> count

    Dim emp As String
    Dim shiftVal As String
    Dim k As String

    On Error GoTo ErrorHandle

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")
    If lo.DataBodyRange Is Nothing Then GoTo ExitSub

    data = lo.DataBodyRange.Value2

    fDate = ToDateFloor(ws.Range("rngManningFDate").Value)
    eDate = DateAdd("d", CLng(ws.Range("rngManningDateRangeTo").Value) - 1, fDate)

    'Build date header map from TABLE HEADER ROW (robust)
    Set headerDict = BuildHeaderDict_FromTable(lo, vManCalCol)

    If Not headerDict.Exists(DateKey(fDate)) Then
        MsgBox "Start date header not found in tblLookahead: " & Format$(fDate, "dd/mm/yyyy"), vbExclamation
        GoTo ExitSub
    End If
    If Not headerDict.Exists(DateKey(eDate)) Then
        MsgBox "End date header not found in tblLookahead: " & Format$(eDate, "dd/mm/yyyy"), vbExclamation
        GoTo ExitSub
    End If

    startAbsCol = headerDict(DateKey(fDate))
    endAbsCol = headerDict(DateKey(eDate))

    tableStartAbsCol = lo.Range.Column
    relStart = startAbsCol - tableStartAbsCol + 1
    relEnd = endAbsCol - tableStartAbsCol + 1

    'Clamp into table
    If relStart < 1 Then relStart = 1
    If relEnd > lo.ListColumns.Count Then relEnd = lo.ListColumns.Count
    If relStart > relEnd Then GoTo ExitSub

    '---------------------------------------
    ' PASS 1: count allocations per emp/date
    '---------------------------------------
    Set counts = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)
        emp = Trim$(NzStr(data(i, 4))) 'Personnel/Employee column
        If Len(emp) = 0 Then GoTo NextRow1

        'Preserve your original skip: ignore if emp == PID (col 8)
        If StrComp(emp, Trim$(NzStr(data(i, 8))), vbTextCompare) = 0 Then GoTo NextRow1

        For c = relStart To relEnd
            shiftVal = Trim$(NzStr(data(i, c)))
            If Len(shiftVal) > 0 Then
                k = emp & "|" & CStr(c)
                If counts.Exists(k) Then
                    counts(k) = counts(k) + 1
                Else
                    counts.Add k, 1
                End If
            End If
        Next c

NextRow1:
    Next i

    '---------------------------------------
    ' PASS 2: flag per row (old behaviour)
    '---------------------------------------
    Dim outArr() As Variant
    Dim hasClash As Boolean
    Dim hasAnyShiftInWindow As Boolean

    ReDim outArr(1 To UBound(data, 1), 1 To 1)

    For i = 1 To UBound(data, 1)
        emp = Trim$(NzStr(data(i, 4)))
        hasClash = False
        hasAnyShiftInWindow = False

        If Len(emp) > 0 Then
            If StrComp(emp, Trim$(NzStr(data(i, 8))), vbTextCompare) <> 0 Then

                For c = relStart To relEnd
                    shiftVal = Trim$(NzStr(data(i, c)))
                    If Len(shiftVal) > 0 Then
                        hasAnyShiftInWindow = True
                        k = emp & "|" & CStr(c)
                        If counts.Exists(k) Then
                            If counts(k) > 1 Then
                                hasClash = True
                                Exit For
                            End If
                        End If
                    End If
                Next c
            End If
        End If

        'Old behaviour: only TRUE when there is a clash on a date where this row has a shift
        outArr(i, 1) = (hasAnyShiftInWindow And hasClash)
    Next i

    lo.ListColumns(11).DataBodyRange.Value2 = outArr

ExitSub:
    If vGoAll = False Then GoAll
    Exit Sub

ErrorHandle:
    MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical, "CheckClashesAndFillColumn"
    Resume ExitSub
End Sub


'============================================================
' IsClash (UDF) - kept but bug-fixed
'============================================================
Public Function IsClash( _
    EmpRng As Range, _
    LstEmp As Range, _
    ShiftRng As Range, _
    ManningFDate As Variant, _
    ManningDateRangeTo As Variant) As Boolean

    Dim sht As Worksheet
    Dim vDate As Date, vEndDate As Date
    Dim lCol As Long
    Dim startCol As Variant, endCol As Variant
    Dim nRng As Range
    Dim arrData As Variant
    Dim i As Long
    Dim colAbs As Long
    Dim dateVal As Date
    Dim rngCol As Range
    Dim lngSRow As Long, lngERow As Long
    Dim empVal As Variant

    On Error GoTo Exit_Function
    Application.Volatile True 'kept (expensive but preserves UDF behaviour)

    Set sht = EmpRng.Worksheet
    empVal = EmpRng.Value2

    lngSRow = LstEmp.row
    lngERow = lngSRow + LstEmp.Rows.Count - 1

    vDate = ToDateFloor(ManningFDate)
    vEndDate = DateAdd("d", CLng(ManningDateRangeTo) - 1, vDate)

    lCol = sht.Cells(2, sht.Columns.Count).End(xlToLeft).Column

    'Find start/end columns by matching dates in row 2 from col 13 onward
    startCol = Application.Match(vDate, sht.Range(sht.Cells(2, 13), sht.Cells(2, lCol)), 0)
    endCol = Application.Match(vEndDate, sht.Range(sht.Cells(2, 13), sht.Cells(2, lCol)), 0)

    If Not IsError(startCol) And Not IsError(endCol) Then
        Set nRng = sht.Range(sht.Cells(ShiftRng.row, CLng(startCol) + 12), sht.Cells(ShiftRng.row, CLng(endCol) + 12))
    Else
        Set nRng = ShiftRng
    End If

    If nRng Is Nothing Then Set nRng = ShiftRng

    If nRng.Cells.Count = 1 Then
        ReDim arrData(1 To 1, 1 To 1)
        arrData(1, 1) = nRng.Value2
    Else
        arrData = nRng.Value2
    End If

    IsClash = False

    For i = 1 To UBound(arrData, 2)
        If Trim$(NzStr(arrData(1, i))) <> "" Then
            colAbs = nRng.Cells(1, i).Column
            If IsDate(sht.Cells(2, colAbs).Value2) Then
                dateVal = ToDateFloor(sht.Cells(2, colAbs).Value2)
                If dateVal >= vDate And dateVal <= vEndDate Then
                    Set rngCol = sht.Range(sht.Cells(lngSRow, colAbs), sht.Cells(lngERow, colAbs))
                    'BUG FIX: EmpRng must be value, not range object
                    If Application.WorksheetFunction.CountIfs(LstEmp, empVal, rngCol, "<>") > 1 Then
                        IsClash = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

Exit_Function:
    Set rngCol = Nothing
    Set sht = Nothing
End Function
Private Function GetTableWindowColumns(ByVal lo As ListObject, _
                                       ByVal fDate As Date, ByVal eDate As Date, _
                                       ByRef relStart As Long, ByRef relEnd As Long) As Boolean
    Dim c As Long
    Dim v As Variant
    Dim k As String
    Dim d As Date

    relStart = 0
    relEnd = 0

    For c = 1 To lo.HeaderRowRange.Columns.Count
        v = lo.HeaderRowRange.Cells(1, c).Value2

        'coerce header to a real Date (robust)
        If IsNumeric(v) Then
            d = DateSerial(1899, 12, 30) + Int(CDbl(v))
        ElseIf IsDate(v) Then
            d = ToDateFloor(CDate(v))
        Else
            GoTo NextC
        End If

        If relStart = 0 Then
            If d >= fDate Then relStart = c
        End If

        If d <= eDate Then
            relEnd = c
        ElseIf d > eDate Then
            Exit For
        End If

NextC:
    Next c

    GetTableWindowColumns = (relStart > 0 And relEnd > 0 And relStart <= relEnd)
End Function

'============================================================
' HideCols (optimised: array-based emptiness test)
'============================================================
Public Sub HideCols(Optional FiltPM As Boolean = True, _
                    Optional FiltRole As Boolean = True, _
                    Optional FiltEmp As String)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long
    Dim fDate As Date, eDate As Date
    Dim days As Long

    Dim data As Variant
    Dim outFlag() As Variant

    Dim relStart As Long, relEnd As Long
    Dim r As Long, c As Long

    Dim strPM As String, strRole As String
    Dim ok As Boolean, hasData As Boolean

    On Error GoTo CleanFail

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    fDate = ToDateFloor(ws.Range("rngManningFDate").Value)
    days = CLng(ws.Range("rngManningDateRangeTo").Value)
    eDate = DateAdd("d", days - 1, fDate)

    'Find calendar window columns INSIDE the table
    If Not GetTableWindowColumns(lo, fDate, eDate, relStart, relEnd) Then
        MsgBox "Could not locate date window in tblLookahead headers for " & _
               Format$(fDate, "dd/mm/yyyy") & " to " & Format$(eDate, "dd/mm/yyyy"), vbExclamation
        GoTo CleanExit
    End If

    data = lo.DataBodyRange.Value2
    ReDim outFlag(1 To UBound(data, 1), 1 To 1)

    strPM = Trim$(NzStr(ws.Range("rngManningManager").Value2))
    strRole = Trim$(NzStr(ws.Range("rngManningRole").Value2))

    'Row-by-row flagging (same behaviour as original: row must have any data in the window)
    For r = 1 To UBound(data, 1)
        ok = True

        If FiltPM And Len(strPM) > 0 Then
            If Trim$(NzStr(data(r, 1))) <> strPM Then ok = False
        End If

        If ok And FiltRole And Len(strRole) > 0 Then
            If Trim$(NzStr(data(r, 10))) <> strRole Then ok = False
        End If

        If ok And Len(FiltEmp) > 0 Then
            If Trim$(NzStr(data(r, 4))) <> FiltEmp Then ok = False
        End If

        If ok Then
            hasData = False
            For c = relStart To relEnd
                If Trim$(NzStr(data(r, c))) <> "" Then
                    hasData = True
                    Exit For
                End If
            Next c
            ok = hasData
        End If

        outFlag(r, 1) = ok
    Next r

    'Write flags to table column 11 (K) in one go
    lo.ListColumns(11).DataBodyRange.Value2 = outFlag

    'Apply filter: show only TRUE rows
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo CleanFail
    lo.Range.AutoFilter Field:=11, Criteria1:=True

    ws.Range("A3").Select

CleanExit:
    If vGoAll = False Then GoAll
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

Public Sub ResetManning()

    Call ResetSort

End Sub

'============================================================
' ResetSort (qualified + safe)
'============================================================
Public Sub ResetSort()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lCol As Long, lRow As Long, r As Long, c As Long

    On Error GoTo ExitSub

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0

    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=lo.ListColumns("PM").Range, Order:=xlAscending
        .SortFields.Add key:=lo.ListColumns("Jnumber").Range, Order:=xlAscending
        .SortFields.Add key:=lo.ListColumns("Job").Range, Order:=xlAscending
        .SortFields.Add key:=lo.ListColumns("Personnel").Range, Order:=xlAscending
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Unhide columns (your original starts 20)
    lCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    For c = 20 To lCol
        If ws.Columns(c).Hidden Then ws.Columns(c).Hidden = False
    Next c

    lRow = ws.Range("rngManningTotals").row - 1
    For r = vManFirstRow To lRow
        If ws.Rows(r).Hidden Then ws.Rows(r).Hidden = False
    Next r

    ws.Outline.ShowLevels ColumnLevels:=1

ExitSub:
    If vGoAll = False Then GoAll
End Sub

'============================================================
' SelectPMData / SelectPMRoster (fixed Union trap, returns blocks)
'============================================================
Public Function SelectPMData(Optional wbk As Workbook, Optional strManager As String) As Range
    Dim ws As Worksheet, lo As ListObject
    Dim firstHit As Range, lastHit As Range
    Dim colPM As Range

    If wbk Is Nothing Then Set wbk = ActiveWorkbook
    Set ws = wbk.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")
    Set colPM = lo.ListColumns(1).DataBodyRange

    'Find first/last (assumes contiguous blocks after ResetSort)
    Set firstHit = colPM.Find(What:=strManager, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If firstHit Is Nothing Then
        Set SelectPMData = Nothing
        Exit Function
    End If

    'Last hit: search backwards
    Set lastHit = colPM.Find(What:=strManager, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    Set SelectPMData = ws.Range(firstHit, lastHit).Resize(RowSize:=lastHit.row - firstHit.row + 1, ColumnSize:=6)
End Function

Public Function SelectPMRoster(Optional wbk As Workbook, Optional strManager As String) As Range
    Dim ws As Worksheet, lo As ListObject
    Dim tWs As Worksheet
    Dim firstHit As Range, lastHit As Range
    Dim pasteCol As Long
    Dim dt As Date

    If wbk Is Nothing Then Set wbk = ActiveWorkbook
    Set ws = wbk.Worksheets("Manning")
    Set tWs = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    dt = ToDateFloor(tWs.Range("rngManningFDate").Value)
    pasteCol = GetHeaderMatchColumnIndex(lo, dt)
    If pasteCol = 0 Then
        Set SelectPMRoster = Nothing
        Exit Function
    End If

    'Find first/last PM (assumes contiguous blocks after ResetSort)
    Set firstHit = lo.ListColumns(1).DataBodyRange.Find(What:=strManager, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)
    If firstHit Is Nothing Then
        Set SelectPMRoster = Nothing
        Exit Function
    End If
    Set lastHit = lo.ListColumns(1).DataBodyRange.Find(What:=strManager, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)

    Dim colCount As Long
    colCount = lo.ListColumns.Count - pasteCol + 1

    Set SelectPMRoster = lo.DataBodyRange.Rows(firstHit.row - lo.DataBodyRange.row + 1) _
        .Resize(lastHit.row - firstHit.row + 1, colCount) _
        .Offset(0, pasteCol - 1)
End Function

'============================================================
' Header match helper (kept)
'============================================================
Private Function GetHeaderMatchColumnIndex(ByVal TBL As ListObject, ByVal dt As Date) As Long
    Dim hdr As Range
    Dim target As String
    target = Format$(dt, "dd/mm/yy")

    GetHeaderMatchColumnIndex = 0

    For Each hdr In TBL.HeaderRowRange.Cells
        If Len(hdr.Value2) > 0 Then
            If IsDate(hdr.Value2) Then
                If Format$(ToDateFloor(hdr.Value2), "dd/mm/yy") = target Then
                    GetHeaderMatchColumnIndex = hdr.Column - TBL.Range.Cells(1, 1).Column + 1
                    Exit Function
                End If
            End If
        End If
    Next hdr
End Function

'============================================================
' Legacy + missing procedures (kept for compatibility)
'============================================================

Public Sub HideColsOrig(Optional FiltPM As Boolean = True)
    'Legacy entrypoint - keep behaviour by delegating to HideCols
    On Error GoTo Exit_Function
    HideCols FiltPM:=FiltPM, FiltRole:=True, FiltEmp:=vbNullString
Exit_Function:
End Sub

Public Sub RolesOnly()
    'Your newer RolesOnly is fine. Just qualify refs + app guard.
    On Error GoTo ErrorHandle

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dataArr As Variant
    Dim i As Long, j As Long
    Dim roleCol As Long, personCol As Long
    Dim firstDateCol As Long, lastDateCol As Long
    Dim flagCol As Long
    Dim allEmpty As Boolean

    If vStopAll = False Then StopAll
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    If lo.DataBodyRange Is Nothing Then GoTo ExitSub

    dataArr = lo.DataBodyRange.Value2

    roleCol = GetTableColumnIndex(lo, "Role")
    personCol = GetTableColumnIndex(lo, "Personnel")

    'Your original math is fragile; keep your intent:
    'First date column = vManCalCol + (FDate - headerDateAtvManCalCol)
    firstDateCol = vManCalCol + (ws.Range("rngManningFDate").Value - ws.Cells(2, vManCalCol).Value)
    lastDateCol = firstDateCol + (ws.Range("rngManningDateRangeTo").Value - 1)

    flagCol = 11

    If roleCol = 0 Or personCol = 0 Then
        MsgBox "Required columns not found in tblLookahead", vbExclamation
        GoTo ExitSub
    End If

    If flagCol > lo.ListColumns.Count Then
        MsgBox "Column 11 does not exist in tblLookahead", vbExclamation
        GoTo ExitSub
    End If

    'Clamp date cols into table width
    If firstDateCol < 1 Then firstDateCol = 1
    If lastDateCol > lo.ListColumns.Count Then lastDateCol = lo.ListColumns.Count

    Dim flagArr() As Variant
    ReDim flagArr(1 To UBound(dataArr, 1), 1 To 1)

    For i = 1 To UBound(dataArr, 1)
        allEmpty = True
        For j = firstDateCol To lastDateCol
            If Trim$(CStr(dataArr(i, j))) <> "" Then
                allEmpty = False
                Exit For
            End If
        Next j

        If allEmpty Then
            flagArr(i, 1) = False
        ElseIf StrComp(CStr(dataArr(i, roleCol)), CStr(dataArr(i, personCol)), vbTextCompare) <> 0 _
            Or LCase$(CStr(dataArr(i, roleCol))) = "subcontractor" Then
            flagArr(i, 1) = False
        Else
            flagArr(i, 1) = True
        End If
    Next i

    lo.ListColumns(flagCol).DataBodyRange.Value2 = flagArr
    lo.Range.AutoFilter Field:=flagCol, Criteria1:=True
    ws.Range("A3").Select

ExitSub:
    If vGoAll = False Then GoAll
    Application.EnableEvents = True
    Exit Sub

ErrorHandle:
    MsgBox "RolesOnly error: " & Err.description, vbCritical
    Resume ExitSub
End Sub

Public Sub RolesOnlyOld()
    'Legacy entrypoint - delegate to RolesOnly
    On Error Resume Next
    RolesOnly
End Sub

Public Function GetTableColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, headerName, vbTextCompare) = 0 Then
            GetTableColumnIndex = i
            Exit Function
        End If
    Next i
    GetTableColumnIndex = 0
End Function

Public Sub FilterAndExportPM()
    'Kept mostly as-is but qualified.
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim TBL As ListObject
    Dim savePath As String
    Dim fd As fileDialog
    Dim originalFileName As String
    Dim pmValue As String
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim col As Long
    Dim formulaDS As String, formulaNS As String
    Dim columnLetter As String

    On Error GoTo CleanExit

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set TBL = ws.ListObjects("tblLookahead")

    pmValue = CStr(ws.Range("rngManningManager").Value2)
    If Len(pmValue) = 0 Then
        MsgBox "The PM value is empty (rngManningManager).", vbExclamation
        GoTo CleanExit
    End If

    On Error Resume Next
    TBL.Range.AutoFilter Field:=1, Criteria1:=pmValue
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Error applying AutoFilter for PM value.", vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo CleanExit

    Set newWb = Workbooks.Add
    Set newWs = newWb.Worksheets(1)

    TBL.Range.SpecialCells(xlCellTypeVisible).Copy
    newWs.Range("A1").PasteSpecial xlPasteAll

    lastRow = newWs.Cells(newWs.Rows.Count, 1).End(xlUp).row + 1
    newWs.Cells(lastRow, 1).Value = "Total DS"
    newWs.Cells(lastRow + 1, 1).Value = "Total NS"

    For col = 7 To newWs.Cells(1, newWs.Columns.Count).End(xlToLeft).Column
        columnLetter = Split(newWs.Cells(1, col).Address, "$")(1)

        formulaDS = "=SUMPRODUCT((" & columnLetter & "$2:" & columnLetter & "$" & (lastRow - 1) & _
                    "=""DS"")*(SUBTOTAL(103,OFFSET(" & columnLetter & "$2,ROW(" & columnLetter & "$2:" & _
                    columnLetter & "$" & (lastRow - 1) & ")-MIN(ROW(" & columnLetter & "$2:" & columnLetter & _
                    "$" & (lastRow - 1) & ")),0))))"

        formulaNS = Replace(formulaDS, "=""DS""", "=""NS""")

        newWs.Cells(lastRow, col).Formula = formulaDS
        newWs.Cells(lastRow + 1, col).Formula = formulaNS
    Next col

    newWs.Range("A" & lastRow & ":B" & lastRow + 1).Font.Bold = True
    newWs.Range("A" & lastRow & ":B" & lastRow + 1).HorizontalAlignment = xlCenter
    newWs.Rows(1).AutoFilter

    On Error Resume Next
    TBL.AutoFilter.ShowAllData
    On Error GoTo 0

    originalFileName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    newWs.Columns("A:F").AutoFit

    newWs.Range("G2").Select
    newWs.Application.ActiveWindow.FreezePanes = True

    Set fd = Application.fileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save Filtered Data"
        .InitialFileName = originalFileName & "_" & pmValue & ".xlsx"
        If .Show = -1 Then
            savePath = .SelectedItems(1)
        Else
            MsgBox "Export canceled.", vbInformation
            GoTo CleanExit
        End If
    End With

    newWb.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
    MsgBox "Filtered data exported successfully to " & savePath, vbInformation

CleanExit:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub RefreshAllQueries()
    Dim cn As WorkbookConnection
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For Each cn In ThisWorkbook.Connections
        If cn.Type = xlConnectionTypeOLEDB Or cn.Type = xlConnectionTypeODBC Then
            cn.Refresh
        End If
    Next cn

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    MsgBox "All queries and pivot tables have been refreshed!", vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub FilterDuplicates()
    'Kept, but now it depends on optimised CheckClashesAndFillColumn
    Dim msgStr As String
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo ExitSub

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    msgStr = "Filtering Duplicates for " & ws.Range("rngManningDateRangeTo").Value & _
             " days from " & ws.Range("rngManningFDate").Value

    ShowTemporaryMessage "Filter Duplicates", msgStr, 0
    DoEvents

    If Application.Calculation = xlCalculationManual Then
        Application.Calculate
        DoEvents
    End If

    CheckClashesAndFillColumn

    If vStopAll = False Then StopAll

    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0

    lo.Range.AutoFilter Field:=vManFiltCol, Criteria1:="TRUE"

    lo.Sort.SortFields.Clear
    lo.Sort.SortFields.Add key:=lo.ListColumns("Personnel").Range, Order:=xlAscending

    With lo.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Application.Goto ws.Range("A3"), True

ExitSub:
    CloseTemporaryMessage
    If vGoAll = False Then GoAll
End Sub

Private Sub SafeClearRowRange(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long)
    Dim rng As Range
    On Error GoTo Fallback

    Set rng = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    rng.ClearContents
    Exit Sub

Fallback:
    Dim c As Long
    On Error Resume Next
    For c = c1 To c2
        ws.Cells(r, c).ClearContents
    Next c
    On Error GoTo 0
End Sub

Public Function GetTable(ByVal tableName As String, Optional ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set GetTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not GetTable Is Nothing Then Exit Function
    Next ws

    Set GetTable = Nothing
End Function

Public Function GetTableStrict(ByVal tableName As String, Optional ByVal wb As Workbook) As ListObject
    Dim lo As ListObject
    Set lo = GetTable(tableName, wb)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 513, , "Table '" & tableName & "' not found in workbook."
    End If
    Set GetTableStrict = lo
End Function

Private Function GetRolesRange(ByVal ws As Worksheet) As Range
    On Error GoTo TryName
    Set GetRolesRange = ws.ListObjects("tblRoles").DataBodyRange
    If Not GetRolesRange Is Nothing Then Exit Function

TryName:
    On Error GoTo Fail
    Set GetRolesRange = ws.Range("tblRoles")
    Exit Function

Fail:
    Set GetRolesRange = ws.Range("A1")
End Function

Public Sub CleanupManning(Optional ByVal manageAppState As Boolean = True, _
                         Optional ByVal forceCalc As Boolean = False)

    'Kept (your logic), but with one critical micro-optimisation:
    'role check uses a dictionary instead of scanning tblRoles for every row.

    Dim ws As Worksheet
    Dim loRoles As ListObject
    Dim rolesRng As Range
    Dim roleDict As Object

    Dim lRows As Long, lCols As Long
    Dim rCnt As Long, rwCnt As Long, cCnt As Long
    Dim strCrit As String

    On Error GoTo ExitSub

    ShowTemporaryMessage "Manning", "Cleaning Up Manning", 0
    DoEvents

    If manageAppState Then
        If vStopAll = False Then StopAll
        Application.EnableEvents = False
    End If

    If forceCalc Then Application.Calculate

    If IsGod = False Then
        If UnlockCode = False Then GoTo ExitSub
    End If

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set loRoles = GetTableStrict("tblRoles")
    Set rolesRng = loRoles.DataBodyRange

    'Build role dictionary once
    Set roleDict = CreateObject("Scripting.Dictionary")
    Dim c As Range
    For Each c In rolesRng.Cells
        If Len(Trim$(CStr(c.Value2))) > 0 Then roleDict(Trim$(CStr(c.Value2))) = True
    Next c

    With ws
        lCols = .Cells(3, .Columns.Count).End(xlToLeft).Column
        lRows = .Range("rngManningTotals").row - 1

        For rCnt = lRows To vManFirstRow Step -1

            If .Cells(rCnt, 4).Value2 = "" Then GoTo nextRow
            If roleDict.Exists(CStr(.Cells(rCnt, 4).Value2)) Then GoTo nextRow

            strCrit = .Cells(rCnt, 1).Value2 & .Cells(rCnt, 2).Value2 & .Cells(rCnt, 3).Value2 & .Cells(rCnt, 4).Value2

            For rwCnt = rCnt + 1 To lRows
                If .Cells(rwCnt, 1).Value2 & .Cells(rwCnt, 2).Value2 & .Cells(rwCnt, 3).Value2 & .Cells(rwCnt, 4).Value2 = strCrit Then

                    For cCnt = 11 To lCols
                        If .Cells(rwCnt, cCnt).Value2 <> "" Then
                            If .Cells(rCnt, cCnt).Value2 = "" Then
                                .Cells(rCnt, cCnt).Value2 = .Cells(rwCnt, cCnt).Value2
                                .Cells(rwCnt, cCnt).ClearContents
                            End If
                        End If
                    Next cCnt

                    SafeClearRowRange ws, rwCnt, vManCalCol, lCols
                    SafeClearRowRange ws, rwCnt, 1, 6
                End If
            Next rwCnt
nextRow:
        Next rCnt
    End With

    ResetSort

ExitSub:
    CloseTemporaryMessage
    If manageAppState Then
        Application.EnableEvents = True
        If vGoAll = False Then GoAll
    End If
End Sub

Public Sub CleanBlanks()
    Dim ws As Worksheet
    Dim lRow As Long, lCol As Long
    Dim NRow As Long
    Dim rng As Range
    Dim delCells As Range, a As Range

    On Error GoTo Exit_Function

    If IsGod = False Then
        If UnlockCode = False Then GoTo Exit_Function
    End If

    ResetSort

    If vStopAll = False Then StopAll
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("Manning")

    With ws
        lCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
        lRow = .Range("rngManningTotals").row - 1

        For NRow = lRow To vManFirstRow Step -1
            If .Cells(NRow, 1).Value2 <> "" Then
                Set rng = .Range(.Cells(NRow, vManCalCol), .Cells(NRow, lCol))
                If Application.WorksheetFunction.CountA(rng) = 0 Then
                    If delCells Is Nothing Then
                        Set delCells = .Cells(NRow, 1)
                    Else
                        Set delCells = Union(delCells, .Cells(NRow, 1))
                    End If
                End If
            End If
        Next NRow

        If Not delCells Is Nothing Then
            For Each a In delCells.Areas
                a.EntireRow.Delete
            Next a
        End If
    End With

    Clean_tblLookahead_BlankRows

    ws.Range("A3").Select

Exit_Function:
    If vGoAll = False Then GoAll
    Application.EnableEvents = True
End Sub

Public Sub FillRNR()
    Dim rng As Range, rowRng As Range
    Dim c As Range, i As Long, j As Long
    Dim lastCol As Long

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Set rng = Selection
    For Each rowRng In rng.Rows
        lastCol = rowRng.Columns.Count
        For i = 1 To lastCol
            Set c = rowRng.Cells(1, i)
            If Trim$(CStr(c.Value2)) = "" Then
                If i > 1 And Trim$(CStr(rowRng.Cells(1, i - 1).Value2)) <> "" Then
                    For j = i + 1 To lastCol
                        If Trim$(CStr(rowRng.Cells(1, j).Value2)) <> "" Then
                            c.Value2 = "RNR"
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
    Next rowRng

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub FillRNR2()
    Dim rng As Range, c As Range
    Application.Calculation = xlCalculationManual
    Set rng = Selection

    For Each c In rng.Cells
        If Trim$(CStr(c.Value2)) = "" Then c.Value2 = "RNR"
    Next c

    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub FilterNonGladstone()
    On Error GoTo ErrorHandle

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim data As Variant
    Dim outFlag() As Variant

    Dim fDate As Date, eDate As Date
    Dim days As Long
    Dim relStart As Long, relEnd As Long

    Dim r As Long, c As Long
    Dim classVal As String, pohVal As Variant
    Dim hasData As Boolean, ok As Boolean

    Dim lastAbsCol As Long, absStartCol As Long
    Dim tableAbsCol As Long

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    If lo.DataBodyRange Is Nothing Then GoTo ExitSub

    '--- reset sort + clear filters (like your pattern)
    ResetSort
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo ErrorHandle

    '--- dates
    fDate = ToDateFloor(ws.Range("rngManningFDate").Value)
    days = CLng(ws.Range("rngManningDateRangeTo").Value)
    eDate = DateAdd("d", days - 1, fDate)

    '--- find date window columns INSIDE the table
    If Not GetTableWindowColumns(lo, fDate, eDate, relStart, relEnd) Then
        MsgBox "Could not locate date window in tblLookahead headers for " & _
               Format$(fDate, "dd/mm/yyyy") & " to " & Format$(eDate, "dd/mm/yyyy"), vbExclamation
        GoTo ExitSub
    End If

    data = lo.DataBodyRange.Value2
    ReDim outFlag(1 To UBound(data, 1), 1 To 1)

    '--- build boolean flag array in memory
    For r = 1 To UBound(data, 1)

        classVal = Trim$(NzStr(data(r, vManClassCol)))  'Class column index (your constant)
        pohVal = data(r, vManPOHCol)                    'POH column index (your constant)

        ok = True

        'Class must be present
        If Len(classVal) = 0 Then ok = False

        'POH must be non-Gladstone and not 0
        If ok Then
            If VarType(pohVal) = vbString Then
                If Trim$(CStr(pohVal)) = "" Then ok = False
                If StrComp(Trim$(CStr(pohVal)), "Gladstone", vbTextCompare) = 0 Then ok = False
                If Trim$(CStr(pohVal)) = "0" Then ok = False
            Else
                'Numeric or other type
                If pohVal = 0 Then ok = False
            End If
        End If

        'Must have at least one allocation in the date window
        If ok Then
            hasData = False
            For c = relStart To relEnd
                If Trim$(NzStr(data(r, c))) <> "" Then
                    hasData = True
                    Exit For
                End If
            Next c
            ok = hasData
        End If

        outFlag(r, 1) = ok
    Next r

    '--- write flag back (Column 11) and filter
    lo.ListColumns(11).DataBodyRange.Value2 = outFlag
    lo.Range.AutoFilter Field:=11, Criteria1:=True

    '--- hide columns prior to FDate for readability (optional but matches old behaviour)
    ws.Columns.Hidden = False

    tableAbsCol = lo.Range.Column
    absStartCol = tableAbsCol + relStart - 1
    lastAbsCol = lo.Range.Column + lo.ListColumns.Count - 1

    Dim cc As Long
    For cc = vManCalCol To absStartCol - 1
        ws.Columns(cc).Hidden = True
    Next cc

    ws.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ws.Range("A3").Select

ExitSub:
    If vGoAll = False Then GoAll
    Exit Sub

ErrorHandle:
    Resume ExitSub
End Sub



Public Sub FilterRole()
    On Error GoTo ExitSub
    Dim ws As Worksheet
    Dim vStr As String

    If vStopAll = False Then StopAll
    Set ws = ThisWorkbook.Worksheets("Manning")

    vStr = CStr(ws.Range("rngManningRole").Value2)
    If Len(vStr) = 0 Then GoTo ExitSub

    ResetSort
    ws.ListObjects("tblLookahead").Range.AutoFilter Field:=vManRoleCol, Criteria1:=vStr
    ws.Range("A3").Select

ExitSub:
    If vGoAll = False Then GoAll
End Sub

Public Sub DeletePMForecast()
    'Clears the selected PM forecast from the selected start date (Cells(2,2)) to the end of tblLookahead columns.
    'PM read from Cells(1,1).
    'Clears blocks where column A = PM, from matching date header to last table column.
    'Prompts + UnlockCode behaviour preserved.

    On Error GoTo ErrorHandle

    Dim ws As Worksheet
    Dim lo As ListObject

    Dim pm As String
    Dim vDate As Date

    Dim dAbsCol As Long
    Dim tFirstRow As Long, tLastRow As Long
    Dim tLastAbsCol As Long

    Dim r As Long
    Dim inBlock As Boolean, blockStart As Long
    Dim delRng As Range

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    If lo.DataBodyRange Is Nothing Then GoTo ExitSub

    'PM from A1
    pm = Trim$(CStr(ws.Cells(1, 1).Value2))
    If Len(pm) = 0 Then
        MsgBox "No PM Selected", vbExclamation
        GoTo ExitSub
    End If

    'Date from B2
    If ws.Cells(2, 2).Value2 <> 0 Then
        vDate = ToDateFloor(ws.Cells(2, 2).Value2)
    Else
        MsgBox "No date selected (cell B2).", vbExclamation
        GoTo ExitSub
    End If

    If MsgBox("Do you want to clear the PM Forecast for PM " & pm & _
              " from " & Format$(vDate, "dd/mm/yyyy") & "?", _
              vbYesNo + vbQuestion, "Clear Forecast") = vbNo Then
        GoTo ExitSub
    End If

    If IsGod = False Then
        If UnlockCode = False Then GoTo ExitSub
    End If

    'Match date column in the TABLE header row
    dAbsCol = FindDateAbsColInLookahead(lo, vDate)
    If dAbsCol = 0 Then
        MsgBox "No matching date column for " & Format$(vDate, "dd/mm/yyyy"), vbExclamation
        GoTo ExitSub
    End If

    'Table bounds (ONLY clear within tblLookahead area)
    tFirstRow = lo.DataBodyRange.row
    tLastRow = lo.DataBodyRange.row + lo.DataBodyRange.Rows.Count - 1
    tLastAbsCol = lo.Range.Column + lo.Range.Columns.Count - 1

    'Build union of PM blocks to clear (table rows only)
    inBlock = False
    blockStart = 0

    For r = tFirstRow To tLastRow + 1 'sentinel
        Dim isMatch As Boolean
        If r <= tLastRow Then
            isMatch = (Trim$(CStr(ws.Cells(r, 1).Value2)) = pm)
        Else
            isMatch = False
        End If

        If isMatch Then
            If Not inBlock Then
                inBlock = True
                blockStart = r
            End If
        Else
            If inBlock Then
                If delRng Is Nothing Then
                    Set delRng = ws.Range(ws.Cells(blockStart, dAbsCol), ws.Cells(r - 1, tLastAbsCol))
                Else
                    Set delRng = Union(delRng, ws.Range(ws.Cells(blockStart, dAbsCol), ws.Cells(r - 1, tLastAbsCol)))
                End If
                inBlock = False
            End If
        End If
    Next r

    If Not delRng Is Nothing Then delRng.ClearContents

    ws.Range("A3").Select

ExitSub:
    If vGoAll = False Then GoAll
    Exit Sub

ErrorHandle:
    Resume ExitSub
End Sub


'============================================================
' Helper: finds the absolute worksheet column in tblLookahead
' whose header matches the given date (floored).
'============================================================
Private Function FindDateAbsColInLookahead(ByVal lo As ListObject, ByVal dt As Date) As Long
    Dim c As Long
    Dim v As Variant
    Dim d As Date

    FindDateAbsColInLookahead = 0

    For c = 1 To lo.HeaderRowRange.Columns.Count
        v = lo.HeaderRowRange.Cells(1, c).Value2

        If IsNumeric(v) Then
            d = DateSerial(1899, 12, 30) + Int(CDbl(v))
        ElseIf IsDate(v) Then
            d = ToDateFloor(CDate(v))
        Else
            GoTo NextC
        End If

        If d = dt Then
            FindDateAbsColInLookahead = lo.HeaderRowRange.Cells(1, c).Column
            Exit Function
        End If

NextC:
    Next c
End Function


Public Sub DeletePMForecastPass(Optional ByVal strManager As String = "", _
                                Optional ByVal intDate As Date = 0, _
                                Optional ByVal skipPass As Boolean = False)
    'Clears a PM forecast from a start date to the end of tblLookahead columns.
    'PM defaults to rngManningManager when not provided.
    'Date defaults to Cells(2,2) when not provided.
    'Optional confirmation (skipPass=True skips prompt).
    'UnlockCode behaviour preserved.
    'Clears PM blocks (col A) within tblLookahead DataBodyRange only.

    On Error GoTo ErrorHandle

    Dim ws As Worksheet
    Dim lo As ListObject

    Dim pm As String
    Dim vDate As Date

    Dim dAbsCol As Long
    Dim tFirstRow As Long, tLastRow As Long
    Dim tLastAbsCol As Long

    Dim r As Long
    Dim inBlock As Boolean, blockStart As Long
    Dim delRng As Range

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    Set lo = ws.ListObjects("tblLookahead")

    If lo.DataBodyRange Is Nothing Then GoTo ExitSub

    'PM
    If Len(Trim$(strManager)) = 0 Then
        pm = Trim$(CStr(ws.Range("rngManningManager").Value2))
    Else
        pm = Trim$(strManager)
    End If
    If Len(pm) = 0 Then
        MsgBox "No PM Selected", vbExclamation
        GoTo ExitSub
    End If

    'Date
    If intDate = 0 Then
        If ws.Cells(2, 2).Value2 <> 0 Then
            vDate = ToDateFloor(ws.Cells(2, 2).Value2)
        Else
            MsgBox "No date selected (cell B2).", vbExclamation
            GoTo ExitSub
        End If
    Else
        vDate = ToDateFloor(intDate)
    End If

    'Confirm
    If Not skipPass Then
        If MsgBox("Do you want to clear the PM Forecast for PM " & pm & _
                  " from " & Format$(vDate, "dd/mm/yyyy") & "?", _
                  vbYesNo + vbQuestion, "Clear Forecast") = vbNo Then
            GoTo ExitSub
        End If
    End If

    If IsGod = False Then
        If UnlockCode = False Then GoTo ExitSub
    End If

    'Match date column in the TABLE header row
    dAbsCol = FindDateAbsColInLookahead(lo, vDate)
    If dAbsCol = 0 Then
        MsgBox "No matching date column for " & Format$(vDate, "dd/mm/yyyy"), vbExclamation
        GoTo ExitSub
    End If

    'Table bounds
    tFirstRow = lo.DataBodyRange.row
    tLastRow = lo.DataBodyRange.row + lo.DataBodyRange.Rows.Count - 1
    tLastAbsCol = lo.Range.Column + lo.Range.Columns.Count - 1

    'Build union of PM blocks to clear
    inBlock = False
    blockStart = 0

    For r = tFirstRow To tLastRow + 1 'sentinel
        Dim isMatch As Boolean
        If r <= tLastRow Then
            isMatch = (Trim$(CStr(ws.Cells(r, 1).Value2)) = pm)
        Else
            isMatch = False
        End If

        If isMatch Then
            If Not inBlock Then
                inBlock = True
                blockStart = r
            End If
        Else
            If inBlock Then
                If delRng Is Nothing Then
                    Set delRng = ws.Range(ws.Cells(blockStart, dAbsCol), ws.Cells(r - 1, tLastAbsCol))
                Else
                    Set delRng = Union(delRng, ws.Range(ws.Cells(blockStart, dAbsCol), ws.Cells(r - 1, tLastAbsCol)))
                End If
                inBlock = False
            End If
        End If
    Next r

    If Not delRng Is Nothing Then delRng.ClearContents

    ws.Range("A3").Select

ExitSub:
    If vGoAll = False Then GoAll
    Exit Sub

ErrorHandle:
    Resume ExitSub
End Sub

Public Sub ImportPMData()
    Dim st As appState
    Dim wbk As Workbook
    Dim wbOpened As Boolean

    Dim strManager As String
    Dim intDate As Date

    Dim ws As Worksheet
    Dim TBL As ListObject

    Dim selectedWbName As String
    Dim selectedWbPath As String
    Dim frm As frmSelect

    Dim PMDataRng As Range, PMRosterRng As Range
    Dim dataArr As Variant, rosterArr As Variant
    Dim dataRows As Long, dataCols As Long
    Dim rosterRows As Long, rosterCols As Long

    Dim pasteCol As Long
    Dim startRow As Long
    Dim oldBodyRows As Long
    Dim searchValue As String

    On Error GoTo Fail

    '========================
    ' App Guard (your standard)
    '========================
    PushAppState st, manualCalc:=True, noScreen:=True, noEvents:=True, noAlerts:=True

    If vStopAll = False Then StopAll

    Set ws = ThisWorkbook.Worksheets("Manning")
    With ws
        strManager = CStr(.Range("rngManningManager").Value)
        intDate = .Range("rngManningFDate").Value
        Set TBL = .ListObjects("tblLookahead")
    End With

    If Len(strManager) = 0 Or intDate = 0 Then
        MsgBox "Manager and Date not selected", vbCritical
        GoTo CleanExit
    End If

    If MsgBox("Do you want to update the PM Forecast for PM: " & strManager & _
              " from: " & Format$(intDate, "dd/mm/yyyy") & "?", _
              vbYesNo + vbQuestion, "Update Forecast") = vbNo Then
        GoTo CleanExit
    End If

    '========================
    ' Workbook selection
    '========================
    If MsgBox("Do you want to select a closed workbook?", vbYesNo + vbQuestion, "Workbook Type") = vbYes Then

        selectedWbPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
        If selectedWbPath = "False" Then
            MsgBox "No workbook selected. Exiting.", vbExclamation
            GoTo CleanExit
        End If

        Set wbk = Workbooks.Open( _
                    Filename:=selectedWbPath, _
                    ReadOnly:=True, _
                    Notify:=False, _
                    UpdateLinks:=0, _
                    IgnoreReadOnlyRecommended:=True)

        wbOpened = True

    Else
        Set frm = New frmSelect
        frm.FrmType = 1
        frm.LoadCombo
        frm.Show

        selectedWbName = frm.SelectedWorkbookName
        Unload frm
        Set frm = Nothing

        If Len(selectedWbName) = 0 Then
            MsgBox "No workbook selected. Exiting.", vbExclamation
            GoTo CleanExit
        End If

        On Error Resume Next
        Set wbk = Workbooks(selectedWbName)
        On Error GoTo Fail

        If wbk Is Nothing Then
            MsgBox "Selected workbook is not open. Please open it first.", vbExclamation
            GoTo CleanExit
        End If

        wbOpened = False
    End If

    '========================
    ' Ensure table not filtered (filters + slicers can explode performance)
    '========================
    On Error Resume Next
    If TBL.AutoFilter.FilterMode Then TBL.AutoFilter.ShowAllData
    On Error GoTo Fail

    '========================
    ' Pause pivots/slicers (prevents "Running Slicer Operation")
    '========================
    PivotSlicerGuard_Begin ThisWorkbook

    '========================
    ' Delete existing PM data
    ' (reassert events OFF in case downstream code flips them)
    '========================
    Application.EnableEvents = False
    DeletePMForecastPass strManager, intDate, True
    Application.EnableEvents = False

    '========================
    ' Pull PM data (source workbook)
    '========================
    Set PMDataRng = SelectPMData(wbk, strManager)
    If PMDataRng Is Nothing Then
        MsgBox "No PM data found for " & strManager, vbInformation
        GoTo CleanExit
    End If

    dataArr = PMDataRng.Value
    dataRows = UBound(dataArr, 1)
    dataCols = UBound(dataArr, 2)

    If dataRows < 1 Or dataCols < 1 Then
        MsgBox "No PM data found for " & strManager, vbInformation
        GoTo CleanExit
    End If

    '========================
    ' Grow table ONCE (no ListRows.Add loop)
    '========================
    oldBodyRows = 0
    If Not TBL.DataBodyRange Is Nothing Then oldBodyRows = TBL.DataBodyRange.Rows.Count

    startRow = oldBodyRows + 1   '1-based row index within DataBodyRange after expansion

    EnsureTableHasBodyRows TBL, oldBodyRows + dataRows

    '========================
    ' Bulk write PM data rows
    '========================
    TBL.DataBodyRange.Cells(startRow, 1).Resize(dataRows, dataCols).Value = dataArr

    '========================
    ' Find matching date column in headers
    '========================
    pasteCol = GetHeaderMatchColumnIndex(TBL, intDate)
    If pasteCol = 0 Then
        MsgBox "Date not found in table headers.", vbExclamation
        GoTo CleanExit
    End If

    searchValue = Format$(intDate, "dd/mm/yy")

    '========================
    ' Pull & write roster to matching column
    '========================
    Set PMRosterRng = SelectPMRoster(wbk, strManager)
    If Not PMRosterRng Is Nothing Then
        rosterArr = PMRosterRng.Value
        rosterRows = UBound(rosterArr, 1)
        rosterCols = UBound(rosterArr, 2)

        If rosterRows > 0 And rosterCols > 0 Then
            TBL.DataBodyRange.Cells(startRow, pasteCol).Resize(rosterRows, rosterCols).Value = rosterArr
        End If
    Else
        MsgBox "No roster data found for " & strManager, vbInformation
    End If

    '========================
    ' Post-cleanup (keep events off)
    '========================
    Application.EnableEvents = False
    CleanupManning
    Application.EnableEvents = False
    CleanBlanks
    Application.EnableEvents = False

    '========================
    ' Resume pivots/slicers with ONE refresh
    '========================
    PivotSlicerGuard_End ThisWorkbook, doRefresh:=True

    MsgBox "Import for " & strManager & " from the " & searchValue & " Complete", vbInformation

CleanExit:
    On Error Resume Next
    If wbOpened Then If Not wbk Is Nothing Then wbk.Close SaveChanges:=False
    On Error GoTo 0

    If vGoAll = False Then GoAll

    PopAppState st
    Exit Sub

Fail:
    Dim msg As String
    msg = "ImportPMData failed." & vbCrLf & _
          "Error " & Err.Number & ": " & Err.description

    On Error Resume Next
    If wbOpened Then If Not wbk Is Nothing Then wbk.Close SaveChanges:=False
    PivotSlicerGuard_End ThisWorkbook, doRefresh:=False
    If vGoAll = False Then GoAll
    PopAppState st
    On Error GoTo 0

    MsgBox msg, vbCritical
End Sub

Private Sub PivotSlicerGuard_Begin(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In wb.SlicerCaches
        sc.EnableRefresh = False
    Next sc
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            pt.ManualUpdate = True
        Next pt
    Next ws
    On Error GoTo 0
End Sub

Private Sub PivotSlicerGuard_End(ByVal wb As Workbook, ByVal doRefresh As Boolean)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In wb.SlicerCaches
        sc.EnableRefresh = True
    Next sc

    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            pt.ManualUpdate = False
            If doRefresh Then pt.RefreshTable
        Next pt
    Next ws
    On Error GoTo 0
End Sub
Private Sub EnsureTableHasBodyRows(ByVal lo As ListObject, ByVal requiredBodyRows As Long)
    'Expands a table to requiredBodyRows by:
    '  1) inserting worksheet rows to create space (avoids overlap with other tables)
    '  2) resizing the ListObject once
    '
    'Fixes: Error 1004 "affects cells inside and outside of a table / multiple tables"

    Dim ws As Worksheet
    Dim curBodyRows As Long
    Dim addBody As Long
    Dim totalCols As Long
    Dim headerRow As Long
    Dim curTotalRows As Long
    Dim needTotalRows As Long
    Dim totalsRowExtra As Long

    Dim curBottom As Long
    Dim newBottom As Long
    Dim rowsToInsert As Long

    Dim topLeft As Range, newRange As Range
    Dim lo2 As ListObject
    Dim insertAtRow As Long
    Dim overlapFound As Boolean

    Set ws = lo.Parent

    curBodyRows = 0
    If Not lo.DataBodyRange Is Nothing Then curBodyRows = lo.DataBodyRange.Rows.Count
    If requiredBodyRows <= curBodyRows Then Exit Sub

    totalCols = lo.Range.Columns.Count
    headerRow = lo.HeaderRowRange.row

    totalsRowExtra = IIf(lo.ShowTotals, 1, 0)

    curTotalRows = 1 + curBodyRows + totalsRowExtra
    needTotalRows = 1 + requiredBodyRows + totalsRowExtra

    curBottom = headerRow + curTotalRows - 1
    newBottom = headerRow + needTotalRows - 1

    rowsToInsert = newBottom - curBottom
    If rowsToInsert <= 0 Then GoTo DoResize

    '------------------------------------------------------------
    ' Create room:
    ' Prefer inserting directly below the table,
    ' but if another ListObject is immediately below / overlaps columns,
    ' insert above that next table instead.
    '------------------------------------------------------------
    insertAtRow = curBottom + 1
    overlapFound = False

    For Each lo2 In ws.ListObjects
        If lo2.Name <> lo.Name Then
            'If the other table starts at or below where we'd insert, and within same column band, treat as conflict.
            If lo2.Range.row >= insertAtRow Then
                If RangesOverlap(lo.Range.Columns(1).Resize(1, totalCols), lo2.Range.Columns(1).Resize(1, lo2.Range.Columns.Count)) Then
                    overlapFound = True
                    insertAtRow = lo2.Range.row 'insert ABOVE next table
                    Exit For
                End If
            End If
        End If
    Next lo2

    ws.Rows(insertAtRow).Resize(rowsToInsert).Insert Shift:=xlDown

DoResize:
    Set topLeft = lo.Range.Cells(1, 1)
    Set newRange = topLeft.Resize(needTotalRows, totalCols)
    lo.Resize newRange
End Sub

Private Function RangesOverlap(ByVal r1 As Range, ByVal r2 As Range) As Boolean
    On Error GoTo NoOverlap
    RangesOverlap = Not (Application.Intersect(r1, r2) Is Nothing)
    Exit Function
NoOverlap:
    RangesOverlap = False
End Function
Public Sub NewRole()
Attribute NewRole.VB_ProcData.VB_Invoke_Func = "C\n14"
    ' Keyboard Shortcut: Ctrl+Shift+C

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim selectedRng As Range
    Dim selectedRow As Long
    Dim selectedCol As Long

    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayStatusBar As Boolean

    On Error GoTo Cleanup

    Set wb = ThisWorkbook

    If TypeName(Application.Selection) <> "Range" Then Exit Sub
    Set selectedRng = Application.Selection
    If Not selectedRng.Parent.Parent Is wb Then Exit Sub

    Set ws = selectedRng.Worksheet
    selectedRow = selectedRng.row
    selectedCol = selectedRng.Column

    If vStopAll = False Then StopAll

    ' Fast mode (save + disable)
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevDisplayStatusBar = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual

    ws.Rows(selectedRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Select Case selectedCol
        Case 1 ' Copy PM (A)
            ws.Cells(selectedRow, 1).Value = ws.Cells(selectedRow + 1, 1).Value

        Case 2 ' Copy A:B and F
            ws.Cells(selectedRow, 1).Value = ws.Cells(selectedRow + 1, 1).Value
            ws.Cells(selectedRow, 2).Value = ws.Cells(selectedRow + 1, 2).Value
            ws.Cells(selectedRow, 6).Value = ws.Cells(selectedRow + 1, 6).Value

        Case 3 ' Copy A:C and F
            ws.Cells(selectedRow, 1).Value = ws.Cells(selectedRow + 1, 1).Value
            ws.Cells(selectedRow, 2).Value = ws.Cells(selectedRow + 1, 2).Value
            ws.Cells(selectedRow, 3).Value = ws.Cells(selectedRow + 1, 3).Value
            ws.Cells(selectedRow, 6).Value = ws.Cells(selectedRow + 1, 6).Value

        Case 4 ' Copy D only
            ws.Cells(selectedRow, 4).Value = ws.Cells(selectedRow + 1, 4).Value

        Case Is > vManCalCol
            ' Copy A:C and E:F
            ws.Cells(selectedRow, 1).Value = ws.Cells(selectedRow + 1, 1).Value
            ws.Cells(selectedRow, 2).Value = ws.Cells(selectedRow + 1, 2).Value
            ws.Cells(selectedRow, 3).Value = ws.Cells(selectedRow + 1, 3).Value
            ws.Cells(selectedRow, 5).Value = ws.Cells(selectedRow + 1, 5).Value
            ws.Cells(selectedRow, 6).Value = ws.Cells(selectedRow + 1, 6).Value

            ' Role J -> D
            ws.Cells(selectedRow, 4).Value = ws.Cells(selectedRow + 1, 10).Value

            ' Move current cell content up one row and clear original
            ws.Cells(selectedRow, selectedCol).Value = selectedRng.Value
            selectedRng.ClearContents

        Case Else
            ' Copy A:C and E:F, Role J -> D
            ws.Range(ws.Cells(selectedRow, 1), ws.Cells(selectedRow, 3)).Value = _
                ws.Range(ws.Cells(selectedRow + 1, 1), ws.Cells(selectedRow + 1, 3)).Value
            ws.Range(ws.Cells(selectedRow, 5), ws.Cells(selectedRow, 6)).Value = _
                ws.Range(ws.Cells(selectedRow + 1, 5), ws.Cells(selectedRow + 1, 6)).Value
            ws.Cells(selectedRow, 4).Value = ws.Cells(selectedRow + 1, 10).Value
    End Select

Cleanup:
    Application.CutCopyMode = False

    ' Restore fast mode settings (even if error)
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.DisplayStatusBar = prevDisplayStatusBar
    On Error GoTo 0

    If vGoAll = False Then GoAll
End Sub


'================================================================================
' Procedure: UpdateFilterCol_UnapprovedLeave
'
' Purpose:
'   Updates tblLookahead[Filter Col] to TRUE where the employee has
'   unapproved leave (Leave Status = "N") overlapping the Manning date window.
'
' Manning Window:
'   Start = Manning!rngManningFDate
'   End   = Start + Manning!rngManningDateRangeTo - 1
'
' Matching Key:
'   tblLookahead[Employee No.]  <->  Leave rngLeavePID
'
' Business Rule:
'   Leave Status = "N" corresponds to COLOR_RULE_12_PURPLE in
'   FormatRosterAll_Optimized.
'
' Behaviour:
'   - TRUE  = employee has at least one unapproved leave record overlapping
'             the Manning date window
'   - FALSE = no overlapping unapproved leave found
'
' Assumptions:
'   - Sheet "Manning" contains table "tblLookahead"
'   - Sheet "tbl_Vista_HR_Leave" contains named ranges:
'         rngLeavePID
'         rngLeaveSDate
'         rngLeaveFDate
'         rngLeaveStatus
'   - tblLookahead contains headers:
'         "Employee No."
'         "Filter Col"
'================================================================================
Public Sub UpdateFilterCol_UnapprovedLeaveold()

    Const PROC_NAME        As String = "UpdateFilterCol_UnapprovedLeave"
    Const WS_MANNING       As String = "Manning"
    Const WS_LEAVE         As String = "tbl_Vista_HR_Leave"
    Const TBL_LOOKAHEAD    As String = "tblLookahead"
    Const COL_EMP_NO       As String = "Employee No."
    Const STATUS_UNAPPROVED As String = "N"

    Dim wsMan As Worksheet
    Dim wsLeave As Worksheet
    Dim lo As ListObject

    Dim empColIndex As Long
    Dim filterColIndex As Long

    Dim dataArr As Variant
    Dim outArr() As Variant

    Dim leavePID As Variant
    Dim leaveStart As Variant
    Dim leaveEnd As Variant
    Dim leaveStatus As Variant

    Dim dictUnapproved As Object

    Dim rowCount As Long
    Dim i As Long

    Dim sDate As Date
    Dim eDate As Date
    Dim daysToShow As Long

    Dim empKey As String
    Dim statusKey As String
    Dim lvStart As Date
    Dim lvEnd As Date

    On Error GoTo ErrorHandle

    '----------------------------------------------------------------------
    ' Resolve objects
    '----------------------------------------------------------------------
    Set wsMan = ThisWorkbook.Worksheets(WS_MANNING)
    Set wsLeave = ThisWorkbook.Worksheets(WS_LEAVE)
    Set lo = wsMan.ListObjects(TBL_LOOKAHEAD)

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    '----------------------------------------------------------------------
    ' Resolve Manning date window
    '----------------------------------------------------------------------
    sDate = CDate(wsMan.Range("rngManningFDate").Value)
    daysToShow = CLng(wsMan.Range("rngManningDateRangeTo").Value)

    If daysToShow <= 0 Then Exit Sub

    eDate = DateAdd("d", daysToShow - 1, sDate)

    '----------------------------------------------------------------------
    ' Resolve required table columns
    '----------------------------------------------------------------------
    empColIndex = lo.ListColumns(COL_EMP_NO).Index
    filterColIndex = lo.ListColumns(vManFiltCol).Index

    '----------------------------------------------------------------------
    ' Read Lookahead table and Leave ranges into memory
    '----------------------------------------------------------------------
    dataArr = lo.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    leavePID = wsLeave.Range("rngLeavePID").Value
    leaveStart = wsLeave.Range("rngLeaveSDate").Value
    leaveEnd = wsLeave.Range("rngLeaveFDate").Value
    leaveStatus = wsLeave.Range("rngLeaveStatus").Value

    '----------------------------------------------------------------------
    ' Build dictionary of employees with overlapping unapproved leave
    '----------------------------------------------------------------------
    Set dictUnapproved = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(leavePID, 1)

        empKey = Trim$(CStr(leavePID(i, 1)))
        statusKey = UCase$(Trim$(CStr(leaveStatus(i, 1))))

        If Len(empKey) > 0 Then
            If statusKey = STATUS_UNAPPROVED Then
                If IsDate(leaveStart(i, 1)) And IsDate(leaveEnd(i, 1)) Then

                    lvStart = CDate(leaveStart(i, 1))
                    lvEnd = CDate(leaveEnd(i, 1))

                    '------------------------------------------------------
                    ' Overlap test:
                    ' leaveStart <= windowEnd AND leaveEnd >= windowStart
                    '------------------------------------------------------
                    If lvStart <= eDate And lvEnd >= sDate Then
                        If Not dictUnapproved.Exists(empKey) Then
                            dictUnapproved.Add empKey, True
                        End If
                    End If

                End If
            End If
        End If

    Next i

    '----------------------------------------------------------------------
    ' Build output array for Filter Col
    '----------------------------------------------------------------------
    ReDim outArr(1 To rowCount, 1 To 1)

    For i = 1 To rowCount

        empKey = Trim$(CStr(dataArr(i, empColIndex)))

        If Len(empKey) > 0 Then
            outArr(i, 1) = dictUnapproved.Exists(empKey)
        Else
            outArr(i, 1) = False
        End If

    Next i

    '----------------------------------------------------------------------
    ' Write results back in one hit
    '----------------------------------------------------------------------
    lo.ListColumns(filterColIndex).DataBodyRange.Value = outArr

    ' Apply filter to show only TRUE rows
    On Error Resume Next
    If wsMan.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo ErrorHandle '

    lo.Range.AutoFilter Field:=vManFiltCol, Criteria1:="TRUE"

    Exit Sub

ErrorHandle:
    MsgBox PROC_NAME & " failed: " & Err.description, vbExclamation

End Sub


'================================================================================
' Procedure: UpdateFilterCol_LeaveStatusFilter
'
' Purpose:
'   Updates tblLookahead[Filter Col] to TRUE where:
'       1. the employee has leave overlapping the Manning date window, and
'       2. the tblLookahead row contains data in at least one date column
'          within that same Manning window.
'
' Leave Mode Prompt:
'   1 = Unapproved Leave Only
'   2 = Approved Leave Only
'   3 = Both
'
' Default:
'   - Blank / Cancel / invalid input = 3 (Both)
'
' Manning Window:
'   Start = Manning!rngManningFDate
'   End   = Start + Manning!rngManningDateRangeTo - 1
'
' Matching Key:
'   tblLookahead[Employee No.] <-> Leave rngLeavePID
'
' Assumptions:
'   - Sheet "Manning" contains table "tblLookahead"
'   - Sheet "tbl_Vista_HR_Leave" contains named ranges:
'         rngLeavePID
'         rngLeaveSDate
'         rngLeaveFDate
'         rngLeaveStatus
'   - tblLookahead contains headers:
'         "Employee No."
'         and the filter column referenced by vManFiltCol
'   - Date columns in tblLookahead have header values that are actual dates
'     (or are convertible to dates)
'
' Behaviour:
'   - TRUE  = employee has matching leave in selected mode overlapping the
'             Manning window AND the row has data in at least one date column
'             inside the Manning window
'   - FALSE = otherwise
'
' Dependencies:
'   - AppGuard_Begin / AppGuard_End
'   - RangeTo2DArray
'   - LeaveStatusMatchesMode
'================================================================================
Public Sub UpdateFilterCol_LeaveStatusFilter()

    '-------------------------------------------------------
    ' Constants
    '-------------------------------------------------------
    Const PROC_NAME As String = "UpdateFilterCol_LeaveStatusFilter"
    Const WS_MANNING As String = "Manning"
    Const WS_LEAVE As String = "tbl_Vista_HR_Leave"
    Const TBL_LOOKAHEAD As String = "tblLookahead"
    Const COL_EMP_NO As String = "Employee No."

    Const MODE_UNAPPROVED As Long = 1
    Const MODE_APPROVED As Long = 2
    Const MODE_BOTH As Long = 3

    Const STATUS_UNAPPROVED As String = "N"
    Const STATUS_APPROVED As String = "A"

    '-------------------------------------------------------
    ' Workbook / table objects
    '-------------------------------------------------------
    Dim wsMan As Worksheet
    Dim wsLeave As Worksheet
    Dim lo As ListObject

    '-------------------------------------------------------
    ' Table column indexes
    '-------------------------------------------------------
    Dim empColIndex As Long
    Dim filterColIndex As Long

    '-------------------------------------------------------
    ' Manning window
    '-------------------------------------------------------
    Dim sDate As Date
    Dim eDate As Date
    Dim daysToShow As Long

    '-------------------------------------------------------
    ' Arrays
    '-------------------------------------------------------
    Dim dataArr As Variant
    Dim outArr() As Variant

    Dim leavePID As Variant
    Dim leaveStart As Variant
    Dim leaveEnd As Variant
    Dim leaveStatus As Variant

    '-------------------------------------------------------
    ' Working variables
    '-------------------------------------------------------
    Dim dictMatched As Object
    Dim rowCount As Long
    Dim leaveRowCount As Long
    Dim i As Long
    Dim j As Long

    Dim empKey As String
    Dim statusKey As String
    Dim lvStart As Date
    Dim lvEnd As Date

    Dim userInput As String
    Dim leaveMode As Long
    Dim modeDescription As String

    Dim appGuardStarted As Boolean

    '-------------------------------------------------------
    ' Date-column detection
    '-------------------------------------------------------
    Dim hdrArr As Variant
    Dim dateCols() As Long
    Dim dateColCount As Long
    Dim hdrVal As Variant
    Dim hdrDate As Long

    '-------------------------------------------------------
    ' Row data test
    '-------------------------------------------------------
    Dim hasDataInWindow As Boolean
    Dim cellVal As Variant
    Dim cellText As String

    On Error GoTo ErrorHandle

    '-------------------------------------------------------
    ' Ask user which leave mode to apply
    '-------------------------------------------------------
    userInput = Application.InputBox( _
                    Prompt:="Enter leave filter mode:" & vbCrLf & vbCrLf & _
                            "1 - Unapproved Leave Only" & vbCrLf & _
                            "2 - Approved Leave Only" & vbCrLf & _
                            "3 - Both" & vbCrLf & vbCrLf & _
                            "Leave blank for default = 3 (Both)", _
                    Title:="Leave Filter Mode", _
                    Default:="3", _
                    Type:=2)

    If userInput = "False" Then
        leaveMode = MODE_BOTH
    ElseIf Len(Trim$(userInput)) = 0 Then
        leaveMode = MODE_BOTH
    ElseIf IsNumeric(userInput) Then
        Select Case CLng(userInput)
            Case MODE_UNAPPROVED, MODE_APPROVED, MODE_BOTH
                leaveMode = CLng(userInput)
            Case Else
                leaveMode = MODE_BOTH
        End Select
    Else
        leaveMode = MODE_BOTH
    End If

    Select Case leaveMode
        Case MODE_UNAPPROVED: modeDescription = "Unapproved Leave Only"
        Case MODE_APPROVED:   modeDescription = "Approved Leave Only"
        Case Else:            modeDescription = "Both"
    End Select

    '-------------------------------------------------------
    ' Begin application guard
    '-------------------------------------------------------
    AppGuard_Begin True, "Updating Manning Filter Col - " & modeDescription & "...", True
    appGuardStarted = True

    '-------------------------------------------------------
    ' Resolve objects
    '-------------------------------------------------------
    Set wsMan = ThisWorkbook.Worksheets(WS_MANNING)
    Set wsLeave = ThisWorkbook.Worksheets(WS_LEAVE)
    Set lo = wsMan.ListObjects(TBL_LOOKAHEAD)

    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    '-------------------------------------------------------
    ' Resolve Manning date window
    '-------------------------------------------------------
    sDate = CDate(wsMan.Range("rngManningFDate").Value)
    daysToShow = CLng(wsMan.Range("rngManningDateRangeTo").Value)

    If daysToShow <= 0 Then GoTo CleanExit

    eDate = DateAdd("d", daysToShow - 1, sDate)

    '-------------------------------------------------------
    ' Resolve required columns
    '-------------------------------------------------------
    empColIndex = lo.ListColumns(COL_EMP_NO).Index
    filterColIndex = lo.ListColumns(vManFiltCol).Index

    '-------------------------------------------------------
    ' Read lookahead table into memory
    '-------------------------------------------------------
    dataArr = lo.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    '-------------------------------------------------------
    ' Identify all tblLookahead columns whose headers fall
    ' within the Manning date window
    '-------------------------------------------------------
    hdrArr = lo.HeaderRowRange.Value
    ReDim dateCols(1 To lo.ListColumns.Count)
    dateColCount = 0

    For j = 1 To lo.ListColumns.Count
        hdrVal = hdrArr(1, j)

        If IsDate(hdrVal) Then
            hdrDate = CLng(CDate(hdrVal))

            If hdrDate >= CLng(sDate) And hdrDate <= CLng(eDate) Then
                dateColCount = dateColCount + 1
                dateCols(dateColCount) = j
            End If
        End If
    Next j

    If dateColCount = 0 Then
        Err.Raise vbObjectError + 1001, PROC_NAME, _
                  "No tblLookahead date columns were found within the Manning date window."
    End If

    ReDim Preserve dateCols(1 To dateColCount)

    '-------------------------------------------------------
    ' Read leave named ranges into memory
    '-------------------------------------------------------
    leavePID = RangeTo2DArray(wsLeave.Range("rngLeavePID").Value)
    leaveStart = RangeTo2DArray(wsLeave.Range("rngLeaveSDate").Value)
    leaveEnd = RangeTo2DArray(wsLeave.Range("rngLeaveFDate").Value)
    leaveStatus = RangeTo2DArray(wsLeave.Range("rngLeaveStatus").Value)

    '-------------------------------------------------------
    ' Validate row alignment of leave ranges
    '-------------------------------------------------------
    leaveRowCount = UBound(leavePID, 1)

    If UBound(leaveStart, 1) <> leaveRowCount _
    Or UBound(leaveEnd, 1) <> leaveRowCount _
    Or UBound(leaveStatus, 1) <> leaveRowCount Then
        Err.Raise vbObjectError + 1000, PROC_NAME, _
                  "Leave named ranges are not aligned to the same row count."
    End If

    '-------------------------------------------------------
    ' Build dictionary of employees with matching overlapping leave
    '
    ' Overlap rule:
    '   leaveStart <= windowEnd AND leaveEnd >= windowStart
    '-------------------------------------------------------
    Set dictMatched = CreateObject("Scripting.Dictionary")

    For i = 1 To leaveRowCount

        empKey = Trim$(CStr(leavePID(i, 1)))
        statusKey = UCase$(Trim$(CStr(leaveStatus(i, 1))))

        If Len(empKey) > 0 Then
            If LeaveStatusMatchesMode(statusKey, leaveMode, STATUS_UNAPPROVED, STATUS_APPROVED) Then
                If IsDate(leaveStart(i, 1)) And IsDate(leaveEnd(i, 1)) Then

                    lvStart = CDate(leaveStart(i, 1))
                    lvEnd = CDate(leaveEnd(i, 1))

                    If lvStart <= eDate And lvEnd >= sDate Then
                        If Not dictMatched.Exists(empKey) Then
                            dictMatched.Add empKey, True
                        End If
                    End If

                End If
            End If
        End If

    Next i

    '-------------------------------------------------------
    ' Build output array:
    ' TRUE only where:
    '   - employee is matched in leave dictionary, and
    '   - row has data in at least one Manning date column
    '-------------------------------------------------------
    ReDim outArr(1 To rowCount, 1 To 1)

    For i = 1 To rowCount

        empKey = Trim$(CStr(dataArr(i, empColIndex)))
        hasDataInWindow = False

        'Check date-window cells for real content
        For j = 1 To dateColCount
            cellVal = dataArr(i, dateCols(j))

            If Not IsError(cellVal) Then
                cellText = Trim$(CStr(cellVal))

                If Len(cellText) > 0 Then
                    hasDataInWindow = True
                    Exit For
                End If
            End If
        Next j

        If Len(empKey) > 0 Then
            outArr(i, 1) = (dictMatched.Exists(empKey) And hasDataInWindow)
        Else
            outArr(i, 1) = False
        End If

    Next i

    '-------------------------------------------------------
    ' Write results back in one hit
    '-------------------------------------------------------
    lo.ListColumns(filterColIndex).DataBodyRange.Value = outArr

    '-------------------------------------------------------
    ' Apply filter to show only TRUE rows
    '-------------------------------------------------------
    On Error Resume Next
    If wsMan.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo ErrorHandle

    lo.Range.AutoFilter Field:=filterColIndex, Criteria1:="=TRUE"

CleanExit:
    If appGuardStarted Then AppGuard_End
    Exit Sub

ErrorHandle:
    If appGuardStarted Then
        On Error Resume Next
        AppGuard_End
        On Error GoTo 0
    End If

    MsgBox PROC_NAME & " failed: " & Err.description, vbExclamation, PROC_NAME
End Sub


'================================================================================
' Function: LeaveStatusMatchesMode
'
' Purpose:
'   Determines whether a leave status should be included for the selected mode.
'
' Modes:
'   1 = Unapproved only
'   2 = Approved only
'   3 = Both
'
' Notes:
'   This keeps the main loop clean and makes the business rule obvious.
'================================================================================
Private Function LeaveStatusMatchesMode(ByVal statusKey As String, _
                                        ByVal leaveMode As Long, _
                                        ByVal statusUnapproved As String, _
                                        ByVal statusApproved As String) As Boolean

    Select Case leaveMode
        Case 1
            LeaveStatusMatchesMode = (statusKey = statusUnapproved)

        Case 2
            LeaveStatusMatchesMode = (statusKey = statusApproved)

        Case 3
            LeaveStatusMatchesMode = (statusKey = statusUnapproved Or statusKey = statusApproved)

        Case Else
            LeaveStatusMatchesMode = (statusKey = statusUnapproved Or statusKey = statusApproved)
    End Select

End Function


'================================================================================
' Function: RangeTo2DArray
'
' Purpose:
'   Normalises a worksheet value into a guaranteed 2D, 1-based array.
'
' Why needed:
'   VBA returns a scalar for a single-cell range and a 2D array for a
'   multi-cell range. This helper standardises both forms.
'================================================================================
Private Function RangeTo2DArray(ByVal v As Variant) As Variant

    Dim arr(1 To 1, 1 To 1) As Variant

    If IsArray(v) Then
        RangeTo2DArray = v
    Else
        arr(1, 1) = v
        RangeTo2DArray = arr
    End If

End Function



'-------------------------------------------------------
' FilterLookaheadByRequiredInductions
'
' Purpose:
'   Builds a dictionary of employees from tbl_HR_Inductions
'   where docType is SBSB, SASA, or SISI, then marks the
'   coinciding rows in tblLookahead[Filter Col] as True and
'   filters tblLookahead to show only True rows.
'
' Requirements:
'   - Table "tbl_HR_Inductions" exists
'   - Table "tblLookahead" exists
'   - tbl_HR_Inductions contains:
'       * a docType column
'       * an employee key column
'   - tblLookahead contains:
'       * an employee key column
'   - modGuardsAndTables provides:
'       * AppGuard_Begin / AppGuard_End
'       * SheetGuard_Begin / SheetGuard_End
'       * LogError
'
' Notes:
'   - Employee matching is done using a normalised text key.
'   - Filter Col is created if missing.
'   - This procedure is written for production-grade use on
'     large table datasets using array-based reads/writes.
'-------------------------------------------------------
Public Sub FilterLookaheadByRequiredInductions()

    '-------------------------------------------------------
    ' Variable declarations
    '-------------------------------------------------------
    Dim appSt           As TAppGuardState
    Dim shStLook        As TSheetGuardState
    Dim shStHR          As TSheetGuardState

    Dim wsLook          As Worksheet
    Dim wsHR            As Worksheet
    Dim loLook          As ListObject
    Dim loHR            As ListObject

    Dim idxHRDocType    As Long
    Dim idxHRKey        As Long
    Dim idxLookKey      As Long
    Dim idxFilterCol    As Long

    Dim arrHR           As Variant
    Dim arrLookKeys     As Variant
    Dim arrOut          As Variant

    Dim dictReq         As Object
    Dim allowedDocs     As Object

    Dim r               As Long
    Dim keyVal          As String
    Dim docVal          As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Guard initialisation
    '-------------------------------------------------------
    AppGuard_Begin

    '-------------------------------------------------------
    ' Resolve source and target tables
    '-------------------------------------------------------
    Set loHR = FindTableByName(ThisWorkbook, "tbl_HR_Inductions")
    Set loLook = FindTableByName(ThisWorkbook, "tblLookahead")

    If loHR Is Nothing Then
        Err.Raise vbObjectError + 1000, "FilterLookaheadByRequiredInductions", _
                  "Table 'tbl_HR_Inductions' was not found."
    End If

    If loLook Is Nothing Then
        Err.Raise vbObjectError + 1001, "FilterLookaheadByRequiredInductions", _
                  "Table 'tblLookahead' was not found."
    End If

    Set wsHR = loHR.Parent
    Set wsLook = loLook.Parent

    shStHR = SheetGuard_Begin(wsHR)
    If wsLook Is wsHR Then
        shStLook = shStHR
    Else
        shStLook = SheetGuard_Begin(wsLook)
    End If

    '-------------------------------------------------------
    ' Resolve required columns and validate structure
    '-------------------------------------------------------
    idxHRDocType = GetColumnIndexByCandidates(loHR, Array("docType", "Doc Type", "Document Type"))
    idxHRKey = GetColumnIndexByCandidates(loHR, Array("Employee", "Employee Name", "Name", "Full Name", "PID", "PersonID"))
    idxLookKey = GetColumnIndexByCandidates(loLook, Array("Employee No.", "Employee", "Employee Name", "Name", "Full Name", "PID"))
    idxFilterCol = EnsureTableColumn(loLook, "Filter Col")

    If idxHRDocType = 0 Then
        Err.Raise vbObjectError + 1002, "FilterLookaheadByRequiredInductions", _
                  "Could not find docType column in 'tbl_HR_Inductions'."
    End If

    If idxHRKey = 0 Then
        Err.Raise vbObjectError + 1003, "FilterLookaheadByRequiredInductions", _
                  "Could not find employee key column in 'tbl_HR_Inductions'."
    End If

    If idxLookKey = 0 Then
        Err.Raise vbObjectError + 1004, "FilterLookaheadByRequiredInductions", _
                  "Could not find employee key column in 'tblLookahead'."
    End If

    '-------------------------------------------------------
    ' Build allowed document type lookup
    '-------------------------------------------------------
    Set allowedDocs = CreateObject("Scripting.Dictionary")
    allowedDocs.CompareMode = vbTextCompare
    allowedDocs("SBSB") = True
    allowedDocs("SASA") = True
    allowedDocs("SISI") = True

    '-------------------------------------------------------
    ' Build dictionary of employees with required inductions
    '-------------------------------------------------------
    Set dictReq = CreateObject("Scripting.Dictionary")
    dictReq.CompareMode = vbTextCompare

    If Not loHR.DataBodyRange Is Nothing Then
        arrHR = loHR.DataBodyRange.Value2

        If IsArray(arrHR) Then
            For r = 1 To UBound(arrHR, 1)
                docVal = UCase$(Trim$(SafeText(arrHR(r, idxHRDocType))))
                keyVal = NormaliseKey(arrHR(r, idxHRKey))

                If Len(keyVal) > 0 Then
                    If allowedDocs.Exists(docVal) Then
                        If Not dictReq.Exists(keyVal) Then
                            dictReq.Add keyVal, True
                        End If
                    End If
                End If
            Next r
        End If
    End If

    '-------------------------------------------------------
    ' Prepare output for tblLookahead[Filter Col]
    '-------------------------------------------------------
    If loLook.DataBodyRange Is Nothing Then
        GoTo SafeFilterApply
    End If

    arrLookKeys = loLook.ListColumns(idxLookKey).DataBodyRange.Value2
    ReDim arrOut(1 To UBound(arrLookKeys, 1), 1 To 1)

    '-------------------------------------------------------
    ' Main processing logic
    '-------------------------------------------------------
    For r = 1 To UBound(arrLookKeys, 1)
        keyVal = NormaliseKey(arrLookKeys(r, 1))

        If Len(keyVal) > 0 Then
            arrOut(r, 1) = dictReq.Exists(keyVal)
        Else
            arrOut(r, 1) = False
        End If
    Next r

    '-------------------------------------------------------
    ' Table writes
    '-------------------------------------------------------
    loLook.ListColumns(idxFilterCol).DataBodyRange.Value2 = arrOut

SafeFilterApply:
    '-------------------------------------------------------
    ' Apply filter on Filter Col = True
    '-------------------------------------------------------
    On Error Resume Next
    If loLook.AutoFilter.FilterMode Then loLook.AutoFilter.ShowAllData
    On Error GoTo ErrHandler

    loLook.Range.AutoFilter Field:=idxFilterCol, Criteria1:=True

Cleanup:
    '-------------------------------------------------------
    ' Cleanup
    '-------------------------------------------------------
    On Error Resume Next

    If Not wsLook Is Nothing Then
        SheetGuard_End wsLook, shStLook
    End If

    If Not wsHR Is Nothing Then
        If Not (wsHR Is wsLook) Then
            SheetGuard_End wsHR, shStHR
        End If
    End If

    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    '-------------------------------------------------------
    ' Error handling
    '-------------------------------------------------------
    On Error Resume Next
    LogError "FilterLookaheadByRequiredInductions", Err.Number, Err.description
    On Error GoTo 0

    Resume Cleanup

End Sub


'-------------------------------------------------------
' FindTableByName
'
' Purpose:
'   Returns the first ListObject in the workbook matching
'   the supplied table name.
'-------------------------------------------------------
Public Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws

End Function


'-------------------------------------------------------
' GetColumnIndexByCandidates
'
' Purpose:
'   Resolves a column index from a list of possible header
'   names. Returns 0 if none are found.
'-------------------------------------------------------
Public Function GetColumnIndexByCandidates(ByVal lo As ListObject, ByVal candidates As Variant) As Long

    Dim i As Long
    Dim idx As Long

    For i = LBound(candidates) To UBound(candidates)
        idx = GetColumnIndex(lo, CStr(candidates(i)))
        If idx > 0 Then
            GetColumnIndexByCandidates = idx
            Exit Function
        End If
    Next i

End Function


'-------------------------------------------------------
' GetColumnIndex
'
' Purpose:
'   Returns the ListColumn index for an exact header match.
'   Returns 0 if not found.
'-------------------------------------------------------
Public Function GetColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerName), vbTextCompare) = 0 Then
            GetColumnIndex = lc.Index
            Exit Function
        End If
    Next lc

End Function


'-------------------------------------------------------
' EnsureTableColumn
'
' Purpose:
'   Ensures the specified column exists in the table.
'   Returns its column index. Creates it if missing.
'-------------------------------------------------------
Public Function EnsureTableColumn(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim idx As Long
    Dim lc As ListColumn

    idx = GetColumnIndex(lo, headerName)
    If idx > 0 Then
        EnsureTableColumn = idx
        Exit Function
    End If

    Set lc = lo.ListColumns.Add
    lc.Name = headerName
    EnsureTableColumn = lc.Index

End Function


'-------------------------------------------------------
' NormaliseKey
'
' Purpose:
'   Converts a raw value to a trimmed, upper-case key for
'   deterministic dictionary matching.
'-------------------------------------------------------
Public Function NormaliseKey(ByVal v As Variant) As String
    NormaliseKey = UCase$(Trim$(SafeText(v)))
End Function


'-------------------------------------------------------
' SafeText
'
' Purpose:
'   Safely converts any value to text, handling Empty,
'   Null, and Excel error values.
'-------------------------------------------------------
Public Function SafeText(ByVal v As Variant) As String

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    SafeText = CStr(v)

End Function

