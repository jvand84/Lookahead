Attribute VB_Name = "modGuardsAndTables"
Option Explicit

'============================================================
' modGuardsAndTables
'
' Guard + Table Helper Module Template (paste-ready)
'============================================================

'--------------------------------
' Sheet protection guard (optional)
'--------------------------------
Public Type TSheetGuardState
    WasProtected As Boolean
    sheetName As String
End Type

'-----------------------------
' Application Guard (AppGuard)
'-----------------------------
Public Type TAppGuardState
    Calc As XlCalculation
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayStatusBar As Boolean
    StatusBarText As Variant
    Cursor As XlMousePointer
End Type

Private mAppSaved As TAppGuardState
Private mAppHasSaved As Boolean
Private mAppGuardDepth As Long

' Begin an application guard.
Public Sub AppGuard_Begin(Optional ByVal showStatus As Boolean = False, _
                         Optional ByVal statusText As String = vbNullString, _
                         Optional ByVal setCalcManual As Boolean = True)

    On Error GoTo ErrorHandle

    mAppGuardDepth = mAppGuardDepth + 1

    ' Only capture state on the outermost guard.
    If mAppGuardDepth = 1 Then
        mAppSaved.Calc = Application.Calculation
        mAppSaved.ScreenUpdating = Application.ScreenUpdating
        mAppSaved.EnableEvents = Application.EnableEvents
        mAppSaved.DisplayStatusBar = Application.DisplayStatusBar
        mAppSaved.StatusBarText = Application.StatusBar
        mAppSaved.Cursor = Application.Cursor

        mAppHasSaved = True

        ' Apply "safe performance mode"
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = True
        Application.Cursor = xlWait

        If setCalcManual Then
            Application.Calculation = xlCalculationManual
        End If
    End If

    If showStatus Then
        If Len(statusText) > 0 Then Application.StatusBar = statusText
    End If

    Exit Sub

ErrorHandle:
    ' If AppGuard_Begin fails, do not leave Excel broken.
    On Error Resume Next
    mAppGuardDepth = mAppGuardDepth - 1
    If mAppGuardDepth <= 0 Then
        mAppGuardDepth = 0
        If mAppHasSaved Then AppGuard_End
    End If
End Sub

' End an application guard.
Public Sub AppGuard_End(Optional ByVal clearStatus As Boolean = True)
    On Error GoTo ErrorHandle

    If mAppGuardDepth <= 0 Then Exit Sub

    mAppGuardDepth = mAppGuardDepth - 1

    ' Only restore on the outermost end.
    If mAppGuardDepth = 0 Then
        If mAppHasSaved Then
            Application.Calculation = mAppSaved.Calc
            Application.ScreenUpdating = mAppSaved.ScreenUpdating
            Application.EnableEvents = mAppSaved.EnableEvents
            Application.DisplayStatusBar = mAppSaved.DisplayStatusBar
            Application.Cursor = mAppSaved.Cursor

            If clearStatus Then
                Application.StatusBar = False
            Else
                Application.StatusBar = mAppSaved.StatusBarText
            End If
        End If

        mAppHasSaved = False
    End If

    Exit Sub

ErrorHandle:
    ' Last-ditch safety restore attempt
    On Error Resume Next
    mAppGuardDepth = 0
    If mAppHasSaved Then
        Application.Calculation = mAppSaved.Calc
        Application.ScreenUpdating = mAppSaved.ScreenUpdating
        Application.EnableEvents = mAppSaved.EnableEvents
        Application.DisplayStatusBar = mAppSaved.DisplayStatusBar
        Application.Cursor = mAppSaved.Cursor
        Application.StatusBar = False
    End If
    mAppHasSaved = False
End Sub

'-------------------------------------------------------
' SheetGuard_Begin
'
' Purpose:
'   Unprotects a worksheet if currently protected and
'   returns its protection state.
'
' Outputs:
'   - TSheetGuardState containing:
'       * sheetName
'       * WasProtected
'
' Notes:
'   - Must be paired with SheetGuard_End
'   - Safe for repeated calls
'-------------------------------------------------------
Public Function SheetGuard_Begin(ByVal ws As Worksheet, _
                                 Optional ByVal password As String = vbNullString) As TSheetGuardState

    Dim st As TSheetGuardState
    On Error GoTo ErrorHandle

    If ws Is Nothing Then
        Err.Raise vbObjectError + 2000, "SheetGuard_Begin", "Worksheet is Nothing."
    End If

    st.sheetName = ws.Name
    st.WasProtected = ws.ProtectContents

    If st.WasProtected Then
        ws.Unprotect password:=password
    End If

    SheetGuard_Begin = st
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "SheetGuard_Begin -> " & ws.Name, Err.description
End Function

'-------------------------------------------------------
' SheetGuard_End
'
' Purpose:
'   Re-protects a worksheet if it was originally protected
'-------------------------------------------------------
Public Sub SheetGuard_End(ByVal ws As Worksheet, _
                          ByRef st As TSheetGuardState, _
                          Optional ByVal password As String = vbNullString)

    On Error GoTo ErrorHandle

    If ws Is Nothing Then Exit Sub

    If st.WasProtected Then
        ws.Protect password:=password
    End If

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "SheetGuard_End -> " & ws.Name, Err.description
End Sub

'-----------------------------
' Table / ListObject Helpers
'-----------------------------

Public Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo ErrorHandle
    Set GetWorksheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandle:
    Err.Raise 5, "GetWorksheet", "Worksheet not found: '" & sheetName & "' in workbook '" & wb.Name & "'"
End Function

Public Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error GoTo ErrorHandle
    Set GetTable = ws.ListObjects(tableName)
    Exit Function
ErrorHandle:
    Err.Raise 5, "GetTable", "Table not found: '" & tableName & "' on sheet '" & ws.Name & "'"
End Function

Public Function GetWorksheetOfTable(ByVal wb As Workbook, ByVal tableName As String) As Worksheet
    Dim ws As Worksheet
    On Error GoTo ErrorHandle

    For Each ws In wb.Worksheets
        If HasTable(ws, tableName) Then
            Set GetWorksheetOfTable = ws
            Exit Function
        End If
    Next ws

    Err.Raise 5, "GetWorksheetOfTable", "Table '" & tableName & "' not found in workbook '" & wb.Name & "'"
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "GetWorksheetOfTable", Err.description
End Function

Public Function HasTable(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    HasTable = Not (lo Is Nothing)
End Function

Public Function GetTableColIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long, h As String, target As String
    target = NormalizeHeader(headerName)

    For i = 1 To lo.ListColumns.Count
        h = NormalizeHeader(lo.ListColumns(i).Name)
        If h = target Then
            GetTableColIndex = i
            Exit Function
        End If
    Next i

    Err.Raise 5, "GetTableColIndex", "Column not found: '" & headerName & "' in table '" & lo.Name & "'"
End Function

Public Function GetTableDataColRange(ByVal lo As ListObject, ByVal headerName As String) As Range
    Dim idx As Long
    idx = GetTableColIndex(lo, headerName)

    If lo.DataBodyRange Is Nothing Then Exit Function
    Set GetTableDataColRange = lo.ListColumns(idx).DataBodyRange
End Function

' Clear a table to header-only (keeps formatting and total row).
' Deterministic: deletes all ListRows safely.
Public Sub ClearTableToHeaderOnly(ByVal lo As ListObject)
    On Error GoTo ErrorHandle

    If lo.ListRows.Count = 0 Then Exit Sub

    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ClearTableToHeaderOnly(" & lo.Name & ")", Err.description
End Sub

Public Sub ClearTableRowsToHeaderOnly(ByVal lo As ListObject)
    On Error GoTo ErrorHandle

    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ClearTableRowsToHeaderOnly(" & lo.Name & ")", Err.description
End Sub

Public Sub ResizeListObjectRowsExact(ByVal lo As ListObject, ByVal nRows As Long)
    Dim cur As Long
    On Error GoTo ErrorHandle

    If nRows < 0 Then Err.Raise 5, "ResizeListObjectRowsExact", "nRows cannot be negative."

    cur = lo.ListRows.Count

    Do While cur < nRows
        lo.ListRows.Add
        cur = cur + 1
    Loop

    Do While cur > nRows
        lo.ListRows(cur).Delete
        cur = cur - 1
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ResizeListObjectRowsExact(" & lo.Name & ")", Err.description
End Sub

Public Function IsCalculatedColumnFast(ByVal lo As ListObject, ByVal colIndex As Long) As Boolean
    On Error GoTo ErrorHandle

    If lo.ListRows.Count = 0 Then
        IsCalculatedColumnFast = False
        Exit Function
    End If

    Dim r As Range
    Set r = lo.ListColumns(colIndex).DataBodyRange.Cells(1, 1)

    IsCalculatedColumnFast = (Len(r.Formula) > 0 And Left$(r.Formula, 1) = "=")
    Exit Function

ErrorHandle:
    IsCalculatedColumnFast = False
End Function

Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeHeader = LCase$(t)
End Function

Public Function TableToArray(ByVal lo As ListObject) As Variant
    On Error GoTo ErrorHandle

    If lo.DataBodyRange Is Nothing Then
        TableToArray = Empty
        Exit Function
    End If

    TableToArray = lo.DataBodyRange.Value2
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "TableToArray(" & lo.Name & ")", Err.description
End Function

Public Sub ArrayToTable(ByVal lo As ListObject, ByVal data As Variant, _
                        Optional ByVal skipCalculatedColumns As Boolean = True)

    Dim r As Long, c As Long, nR As Long, nC As Long
    Dim writeArr As Variant
    Dim colIsCalc() As Boolean

    On Error GoTo ErrorHandle

    If IsEmpty(data) Then
        ClearTableRowsToHeaderOnly lo
        Exit Sub
    End If

    nR = UBound(data, 1)
    nC = UBound(data, 2)

    ResizeListObjectRowsExact lo, nR

    ReDim colIsCalc(1 To lo.ListColumns.Count)
    If skipCalculatedColumns Then
        For c = 1 To lo.ListColumns.Count
            colIsCalc(c) = IsCalculatedColumnFast(lo, c)
        Next c
    End If

    ReDim writeArr(1 To nR, 1 To lo.ListColumns.Count)

    For r = 1 To nR
        For c = 1 To lo.ListColumns.Count
            If c <= nC Then
                If skipCalculatedColumns And colIsCalc(c) Then
                    ' Leave formula column untouched
                Else
                    writeArr(r, c) = data(r, c)
                End If
            End If
        Next c
    Next r

    lo.DataBodyRange.Value2 = writeArr
    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ArrayToTable(" & lo.Name & ")", Err.description
End Sub


'================================================================================
' Function  : FindListObjectByName
' Purpose   : Returns a ListObject by name across the workbook
' Notes     :
'   - Case-insensitive match
'   - Returns Nothing if not found
'   - Does NOT activate sheets
'================================================================================
Public Function FindListObjectByName(ByVal tableName As String) As ListObject
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    On Error GoTo ErrHandler
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then
            For Each lo In ws.ListObjects
                If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                    Set FindListObjectByName = lo
                    Exit Function
                End If
            Next lo
        End If
    Next ws
    
    ' Not found ? return Nothing
    Set FindListObjectByName = Nothing
    Exit Function

ErrHandler:
    Set FindListObjectByName = Nothing
End Function


