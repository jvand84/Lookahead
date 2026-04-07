Attribute VB_Name = "modAwayRoster"
Option Explicit

Sub ChangeAwayRoster()
    Dim frm As frmSelect
    Dim selectedWbName As String
    Dim wbFound As Boolean
    Dim wb As Workbook
    Dim wbk As Workbook
    
    ' Show UserForm
    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show

    selectedWbName = frm.SelectedWorkbookName
    Unload frm

    If selectedWbName = "" Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        GoTo ExitSub
    End If
    
    Let ThisWorkbook.Worksheets("Map").Range("F4") = selectedWbName
    
    ' Find and set the selected workbook
    wbFound = False
    For Each wb In Application.Workbooks
        If wb.Name = selectedWbName Then
            Set wbk = wb
            wbFound = True
            Exit For
        End If
    Next wb

    If Not wbFound Then
        MsgBox "The workbook '" & selectedWbName & "' was not found among open workbooks.", vbExclamation
        GoTo ExitSub
    End If
    
    ChangeLA wbk

ExitSub:

    Exit Sub

ErrorHandle:

    Resume ExitSub
    
End Sub

Function ItemExistsInCombo(cmb As ComboBox, item As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = item Then
            ItemExistsInCombo = True
            Exit Function
        End If
    Next i
    ItemExistsInCombo = False
End Function

Sub ChangeLA(wbk As Workbook)
    Dim st As appState
    Dim ws As Worksheet, mWs As Worksheet
    Dim hRow As Long, lRow As Long, x As Long
    Dim siteCode As String
    Dim mapIdx As Object ' key: code -> array(PM, JobNum, Job)
    Dim key As String, info As Variant
    Dim needInsert As Boolean

    PushAppState st, manualCalc:=True, noScreen:=True, noEvents:=True

    Set ws = wbk.Worksheets("Roster")
    Set mWs = ThisWorkbook.Worksheets("Map")
    siteCode = "GRM"

    ' Build index from Map!tblMap: assume col1=Code, col2=PM, col3=JobNum, col4=Job
    Set mapIdx = BuildMapIndex(mWs.ListObjects("tblMap"), 1, Array(2, 3, 4))

    hRow = 7
    lRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row

    ' Insert headers only if missing
    needInsert = (ws.Cells(hRow, 6).Value <> "PM") _
              Or (ws.Cells(hRow, 11).Value <> "Site")
    If needInsert Then
        ws.Columns("F:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(hRow, 6).Resize(1, 6).Value = Array("PM", "Job Num", "Job", "Name", "Position Req.", "Site")
    End If

    ' Current group’s metadata (copied down while names are listed)
    Dim curPM As String, curJobNum As String, curJob As String

    For x = hRow + 1 To lRow
        ' Row marking a new group: col A has code, col D empty
        If Len(ws.Cells(x, 1).Value2) > 0 And Len(ws.Cells(x, 4).Value2) = 0 Then
            key = CStr(ws.Cells(x, 1).Value2)
            If mapIdx.Exists(key) Then
                info = mapIdx(key)
                curPM = info(0): curJobNum = info(1): curJob = info(2)
            Else
                curPM = "": curJobNum = "": curJob = ""
            End If
        End If

        ' Data row to fill: col D has a name
        If Len(ws.Cells(x, 4).Value2) > 0 Then
            ws.Cells(x, 6).Value2 = curPM
            ws.Cells(x, 7).Value2 = curJobNum
            ws.Cells(x, 8).Value2 = curJob
            ws.Cells(x, 9).Value2 = ws.Cells(x, 4).Value2      ' Name
            ws.Cells(x, 10).Value2 = ws.Cells(x, 5).Value2     ' Position
            ws.Cells(x, 11).Value2 = siteCode
        End If
    Next x

    PopAppState st
End Sub

Private Function BuildMapIndex(TBL As ListObject, keyCol As Long, valueCols As Variant) As Object
    Dim d As Object, r As Range, i As Long, arr, rowArr, vals(), k As String
    Set d = CreateObject("Scripting.Dictionary")
    If TBL.DataBodyRange Is Nothing Then Set BuildMapIndex = d: Exit Function

    arr = TBL.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        k = CStr(arr(i, keyCol))
        If Len(k) > 0 Then
            ReDim vals(LBound(valueCols) To UBound(valueCols))
            Dim j As Long
            For j = LBound(valueCols) To UBound(valueCols)
                vals(j) = arr(i, valueCols(j))
            Next
            d(k) = vals
        End If
    Next
    Set BuildMapIndex = d
End Function

' ---------- ChangeData (robust to array/collection/range/scalar) ----------
Sub ChangeData(Optional wbk As Workbook, Optional rng As Range)
    Dim st As appState
    Dim ws As Worksheet, mWs As Worksheet
    Dim mode As Long
    Dim mapDetail As Object            ' key -> (Value, Color) in some shape
    Dim vis As Range
    Dim c As Range, k As String
    Dim valOut As Variant, colorOut As Variant
    Dim haveVal As Boolean, haveColor As Boolean

    On Error GoTo ErrHandle
    PushAppState st, manualCalc:=True

    Set mWs = ThisWorkbook.Worksheets("Map")

    If rng Is Nothing Then
        If wbk Is Nothing Then
            Set ws = ActiveWorkbook.Worksheets("Roster")
        Else
            Set ws = wbk.Worksheets("Roster")
        End If

        If TypeName(Selection) = "Range" Then
            Set rng = Selection
        Else
            Err.Raise 5, , "ChangeData: No target range selected."
        End If
    Else
        If wbk Is Nothing Then
            Set ws = rng.Parent
        Else
            Set ws = wbk.Worksheets("Roster")
        End If
    End If

    ' mode: 1 = Value+Colour, 2 = Value only, 3 = Colour only
    mode = 1

    Set mapDetail = BuildDetailIndex(mWs.ListObjects("tblDetail"), mode)

    ' Use visible cells if filtered; gracefully fall back if not
    On Error Resume Next
    Set vis = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandle
    If vis Is Nothing Then Set vis = rng

    For Each c In vis.Cells
        k = DetailKey(c, mode)
        If Len(k) > 0 Then
            If mapDetail.Exists(k) Then
                haveVal = False: haveColor = False
                TryExtractValColor mapDetail(k), valOut, colorOut, haveVal, haveColor

                If haveVal Then
                    If Not IsEmpty(valOut) Then c.Value2 = valOut
                End If

                If haveColor Then
                    If Not IsEmpty(colorOut) Then
                        Dim lngCol As Variant
                        lngCol = ToLongOrEmpty(colorOut)
                        If Not IsEmpty(lngCol) Then c.Interior.Color = lngCol
                    End If
                End If
            End If
        End If
    Next c

CleanExit:
    PopAppState st
    Exit Sub

ErrHandle:
    ' Ensure we always restore app state
    On Error Resume Next
    PopAppState st
    On Error GoTo 0
    MsgBox "ChangeData error: " & Err.description, vbExclamation
End Sub


' -------- Helpers --------

' Tries to read (Value, Color) from any of:
'   - 0/1-based array with 1 or 2 elements
'   - Collection (Item(1) = Value, Item(2) = Color if present)
'   - Range (first cell as Value; Color not inferred)
'   - Scalar (taken as Value)
Private Sub TryExtractValColor(ByVal v As Variant, _
                               ByRef valOut As Variant, ByRef colorOut As Variant, _
                               ByRef haveVal As Boolean, ByRef haveColor As Boolean)

    Dim t As String: t = TypeName(v)

    haveVal = False: haveColor = False
    valOut = Empty: colorOut = Empty

    Select Case t
        Case "Variant()", "String()", "Double()", "Long()", "Integer()", "Boolean()"
            ' Generic array case
            Dim lb As Long, ub As Long
            On Error GoTo NotArray
            lb = LBound(v): ub = UBound(v)
            On Error GoTo 0
            If ub >= lb Then
                valOut = v(lb): haveVal = True
                If ub >= lb + 1 Then
                    colorOut = v(lb + 1): haveColor = True
                End If
            End If
            Exit Sub
NotArray:
            ' Fall through to scalar handling

        Case "Collection"
            If v.Count >= 1 Then
                valOut = v.item(1): haveVal = True
            End If
            If v.Count >= 2 Then
                colorOut = v.item(2): haveColor = True
            End If

        Case "Range"
            If v.Cells.CountLarge > 0 Then
                valOut = v.Cells(1, 1).Value
                haveVal = True
            End If

        Case Else
            If IsObject(v) Then
                ' Unknown object (e.g., Dictionary): leave as Empty
            Else
                ' Scalar
                valOut = v
                haveVal = True
            End If
    End Select
End Sub

' Coerces numeric-like to Long for .Interior.Color; returns Empty if not numeric.
Private Function ToLongOrEmpty(ByVal v As Variant) As Variant
    If IsError(v) Then
        ToLongOrEmpty = Empty
    ElseIf IsNumeric(v) Then
        ToLongOrEmpty = CLng(v)
    Else
        ' Allow hex strings like "&HFF0000"
        If VarType(v) = vbString Then
            Dim s As String: s = Trim$(v)
            If Len(s) > 0 Then
                If LCase$(Left$(s, 2)) = "&h" Then
                    ToLongOrEmpty = CLng("&H" & Mid$(s, 3))
                    Exit Function
                End If
                If s Like "#*" Then
                    On Error Resume Next
                    ToLongOrEmpty = CLng(CDec(s))
                    If Err.Number <> 0 Then ToLongOrEmpty = Empty
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
        ToLongOrEmpty = Empty
    End If
End Function


Sub ChangeFormats()
    Dim frm As frmSelect
    Dim selectedWbName As String
    Dim wbk As Workbook, wb As Workbook
    Dim ws As Worksheet
    Dim lRow As Long, lCol As Long, fCol As Long
    Dim rng As Range
    Dim st As appState

    PushAppState st, manualCalc:=True

    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show
    selectedWbName = frm.SelectedWorkbookName
    Unload frm
    If Len(selectedWbName) = 0 Then GoTo ExitSub

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, selectedWbName, vbTextCompare) = 0 Then
            Set wbk = wb
            Exit For
        End If
    Next
    If wbk Is Nothing Then
        MsgBox "Workbook '" & selectedWbName & "' not found.", vbCritical
        GoTo ExitSub
    End If

    Set ws = wbk.Worksheets(1) ' or specific
    lRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
    lCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column

    fCol = ColLetterToNumber(ThisWorkbook.Worksheets("Map").Range("F7").Value2)
    If fCol <= 0 Then
        MsgBox "Map!F7 doesn't look like a column letter.", vbCritical
        GoTo ExitSub
    End If

    Set rng = ws.Range(ws.Cells(9, fCol), ws.Cells(lRow, lCol))
    ChangeData wbk, rng

ExitSub:
    PopAppState st
End Sub


' Col letters -> number (A=1, Z=26, AA=27, ...). Ignores trailing row digits if present (e.g., "I1" -> 9).
Function ColLetterToNumber(colLetter As String) As Long
    Dim s As String, i As Long, ch As Long, n As Long
    
    s = UCase$(Trim$(colLetter))
    If s = "" Then Exit Function
    
    ' Strip any digits (handles inputs like "I1" or "AA23")
    For i = Len(s) To 1 Step -1
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit For ' stop at last non-digit
        s = Left$(s, i - 1)
    Next i
    If s = "" Then Exit Function
    
    ' Reject anything not A–Z
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then Exit Function
        n = n * 26 + (ch - 64)
    Next i
    
    ColLetterToNumber = n
End Function





' ---------- Levenshtein (unchanged) ----------
Function Levenshtein(ByVal s1 As String, ByVal s2 As String) As Long
    Dim i As Long, j As Long
    Dim l1 As Long, l2 As Long
    Dim d() As Long
    Dim cost As Long

    l1 = Len(s1)
    l2 = Len(s2)
    ReDim d(0 To l1, 0 To l2)

    For i = 0 To l1
        d(i, 0) = i
    Next i

    For j = 0 To l2
        d(0, j) = j
    Next j

    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = Application.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    Levenshtein = d(l1, l2)
End Function


' ---------- Detail map (composite key fixed) ----------
Private Function BuildDetailIndex(TBL As ListObject, mode As Long) As Object
    ' mode 1: key = Value & "|" & Color
    ' mode 2: key = Value
    ' mode 3: key = Color
    Dim d As Object, arr, i As Long, k As String
    Dim valueCol As Long: valueCol = 4
    Dim colorCol As Long: colorCol = 5

    Set d = CreateObject("Scripting.Dictionary")
    If TBL.DataBodyRange Is Nothing Then Set BuildDetailIndex = d: Exit Function

    arr = TBL.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        Select Case mode
            Case 1: k = CStr(arr(i, 2)) & "|" & CStr(arr(i, 3))
            Case 2: k = CStr(arr(i, 1))
            Case 3: k = CStr(arr(i, colorCol))
        End Select
        If Len(k) > 0 Then
            d(k) = Array(arr(i, valueCol), arr(i, colorCol))
        End If
    Next i
    Set BuildDetailIndex = d
End Function

Private Function DetailKey(c As Range, mode As Long) As String
    Select Case mode
        Case 1: DetailKey = CStr(c.Value2) & "|" & CStr(c.Interior.Color)
        Case 2: DetailKey = CStr(c.Value2)
        Case 3: DetailKey = CStr(c.Interior.Color)
    End Select
End Function




Function ReturnValuefromTable(wsName As String, tableName As String, searchValue As String, searchCol As Integer, vCol As Integer) As String
    Dim ws As Worksheet
    Dim TBL As ListObject
    Dim dataArr As Variant
    Dim i As Long
    
    ' Set worksheet (modify if needed)
    Set ws = ThisWorkbook.Sheets(wsName) ' Change "Sheet1" to your sheet name
    
    ' Set table reference
    On Error Resume Next
    Set TBL = ws.ListObjects(tableName)
    On Error GoTo 0
    
    ' Exit if table is not found
    If TBL Is Nothing Then
        ReturnValuefromTable = ""
        Exit Function
    End If
    
    ' Load table into array (excluding headers)
    dataArr = TBL.DataBodyRange.Value
    
    ' Loop through Column 1 of the array
    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, searchCol) = searchValue Then
            ReturnValuefromTable = dataArr(i, vCol)
            Exit Function
        End If
    Next i
    
    ' Value not found
    ReturnValuefromTable = ""
End Function

Sub MatchNames()

    Dim frm As frmSelect
    Dim selectedWbName As String
    Dim wbFound As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim NCol As String

    Application.Calculation = xlCalculationManual
    
    ' Show UserForm
    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show

    selectedWbName = frm.SelectedWorkbookName
    Unload frm

    If selectedWbName = "" Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        GoTo ExitSub
    End If

    ' Attempt to find the selected workbook
    wbFound = False
    For Each wb In Application.Workbooks
        If wb.Name = selectedWbName Then
            wbFound = True
            Exit For
        End If
    Next wb

    If Not wbFound Then
        MsgBox "Workbook '" & selectedWbName & "' not found.", vbCritical
        GoTo ExitSub
    End If
    
    ' Get starting column from Map sheet F7
    NCol = ThisWorkbook.Worksheets("map").Range("F12").Value
    
    If IsError(NCol) Then
        MsgBox "Could not find matching column from 'map'!F12.", vbCritical
        GoTo ExitSub
    End If
    
    Call MatchAndCorrectNames_Optimized(wb, NCol)

ExitSub:
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandle:
    
    Resume ExitSub
    
End Sub

' ---------- Match & Correct Names ----------
Sub MatchAndCorrectNames_Optimized(wbk As Workbook, strCol As String)
    Dim st As appState
    Dim wsRoster As Worksheet
    Dim tblEmployees As ListObject
    Dim empSet As Object               ' exact matches (Upper -> True)
    Dim byFirst As Object              ' letter -> Collection of names
    Dim mapAliases As Object           ' Map!tblNames col1->col2
    Dim lastRow As Long, r As Long
    Dim colIdx As Long
    Dim inputName As String, uName As String
    Dim best As String, bestSim As Double
    Dim names As Variant

    On Error GoTo Clean

    PushAppState st, manualCalc:=True, noScreen:=True, noEvents:=True

    Set wsRoster = wbk.Worksheets("Roster")
    Set tblEmployees = ThisWorkbook.Worksheets("Employees").ListObjects("tblEmployees")
    colIdx = ColLetterToNumber(strCol)
    If colIdx <= 0 Then Err.Raise 5, , "Invalid column: " & strCol

    Set empSet = CreateObject("Scripting.Dictionary")
    Set byFirst = CreateObject("Scripting.Dictionary")
    BuildEmployeeIndexes tblEmployees, empSet, byFirst

    Set mapAliases = BuildAliasMap(ThisWorkbook.Worksheets("Map").ListObjects("tblNames"))

    lastRow = wsRoster.Cells(wsRoster.Rows.Count, colIdx).End(xlUp).row
    names = wsRoster.Range(wsRoster.Cells(2, colIdx), wsRoster.Cells(lastRow, colIdx)).Value2

    For r = 1 To UBound(names, 1)
        inputName = Trim$(CStr(names(r, 1)))
        If Len(inputName) = 0 Then GoTo Skip

        ' 1) alias map
        If mapAliases.Exists(UCase$(inputName)) Then
            wsRoster.Cells(r + 2 - 1, colIdx).Value2 = mapAliases(UCase$(inputName))
            wsRoster.Cells(r + 1, colIdx).Interior.Color = RGB(255, 255, 153)
            GoTo Skip
        End If

        ' 2) exact match set
        uName = UCase$(inputName)
        If empSet.Exists(uName) Then
            wsRoster.Cells(r + 1, colIdx).Interior.Color = RGB(144, 238, 144)
            GoTo Skip
        End If

        ' 3) fuzzy within first-letter bucket
        best = ""
        bestSim = 0#
        Dim first As String: first = Left$(uName, 1)
        If byFirst.Exists(first) Then
            Dim coll As Collection, i As Long, cand As String, dist As Long
            Dim denom As Double, sim As Double
            Set coll = byFirst(first)
            For i = 1 To coll.Count
                cand = coll(i)
                dist = Levenshtein(inputName, cand)
                denom = MaxDbl(Len(inputName), Len(cand))
                If denom > 0# Then
                    sim = 1# - (dist / denom)
                    If sim > bestSim Then bestSim = sim: best = cand
                End If
            Next
        End If

        If bestSim >= 0.7 And Len(best) > 0 Then
            wsRoster.Cells(r + 1, colIdx).Value2 = best
            wsRoster.Cells(r + 1, colIdx).Interior.Color = RGB(255, 255, 153)
        Else
            wsRoster.Cells(r + 1, colIdx).Interior.Color = RGB(255, 102, 102)
        End If
Skip:
    Next r

    MsgBox "Matching complete", vbInformation

Clean:
    PopAppState st
    If Err.Number <> 0 Then MsgBox "Error: " & Err.description, vbExclamation
End Sub


' ---------- Employee & Alias Index Builders ----------
Private Sub BuildEmployeeIndexes(TBL As ListObject, empSet As Object, byFirst As Object)
    Dim arr, i As Long, nameCol As Long, nm As String, key As String, first As String
    nameCol = TBL.ListColumns("Person").Index
    If TBL.DataBodyRange Is Nothing Then Exit Sub
    arr = TBL.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        nm = Trim$(CStr(arr(i, nameCol)))
        If Len(nm) > 0 Then
            key = UCase$(nm)
            empSet(key) = True
            first = Left$(key, 1)
            If Not byFirst.Exists(first) Then
                Dim c As New Collection
                byFirst.Add first, c
            End If
            byFirst(first).Add nm
        End If
    Next
End Sub

Private Function BuildAliasMap(TBL As ListObject) As Object
    Dim d As Object, arr, i As Long, src As String, dst As String
    Set d = CreateObject("Scripting.Dictionary")
    If TBL.DataBodyRange Is Nothing Then Set BuildAliasMap = d: Exit Function
    arr = TBL.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        src = UCase$(Trim$(CStr(arr(i, 1))))
        dst = Trim$(CStr(arr(i, 2)))
        If Len(src) > 0 And Len(dst) > 0 Then d(src) = dst
    Next
    Set BuildAliasMap = d
End Function



Function IsInArray(valueToFind As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(valueToFind, arr(i), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function GetValueFromTblNames(lookupValue As String) As Variant
    Dim ws As Worksheet
    Dim TBL As ListObject
    Dim i As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets("Map") ' <-- Change to your sheet name
    Set TBL = ws.ListObjects("tblNames")

    For i = 1 To TBL.ListRows.Count
        If Trim(TBL.DataBodyRange(i, 1).Value) = Trim(lookupValue) Then
            GetValueFromTblNames = Trim(TBL.DataBodyRange(i, 2).Value)
            Exit Function
        End If
    Next i

    GetValueFromTblNames = CVErr(xlErrNA) ' Not found
    Exit Function

ErrHandler:
    GetValueFromTblNames = CVErr(xlErrValue)
End Function



Private Function MaxDbl(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxDbl = a Else MaxDbl = b
End Function

