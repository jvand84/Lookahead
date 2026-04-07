Attribute VB_Name = "modFormatting"
Option Explicit

Dim ws As Worksheet
Dim wbk As Workbook

Dim COLOR_RULE_2_PURPLE As Long
Dim COLOR_RULE_3_YELLOW As Long
Dim COLOR_RULE_4_ORANGE_DEEP As Long
Dim COLOR_RULE_5_BLUE As Long
Dim COLOR_RULE_6_PURPLE As Long
Dim COLOR_RULE_7_CYAN As Long
Dim COLOR_RULE_8_YELLOW As Long
Dim COLOR_RULE_9_RED As Long
Dim COLOR_RULE_10_BLUE As Long
Dim COLOR_RULE_12_PURPLE As Long
Dim COLOR_RULE_11_RED As Long
Dim COLOR_GREEN As Long
Dim COLOR_LBLUE As Long

Global empDict As Object
Global leaveDict As Object
Global trainingDict As Object

Sub GetColours()
    COLOR_RULE_2_PURPLE = RGB(204, 192, 218) 'Non Local
    COLOR_RULE_3_YELLOW = RGB(255, 255, 153) 'Trades Only
    COLOR_RULE_4_ORANGE_DEEP = RGB(252, 213, 180) 'Non Employee
    COLOR_RULE_5_BLUE = RGB(51, 153, 255) 'Tafe
    COLOR_RULE_6_PURPLE = RGB(204, 153, 255) 'Training
    COLOR_RULE_7_CYAN = RGB(102, 204, 255) 'Leave
    COLOR_RULE_12_PURPLE = RGB(153, 102, 255) 'Leave Unapproved
    COLOR_RULE_8_YELLOW = RGB(255, 255, 153) 'Trades
    COLOR_RULE_9_RED = RGB(255, 124, 128) 'Double up
    COLOR_RULE_10_BLUE = RGB(102, 204, 255) 'Crane
    COLOR_GREEN = RGB(204, 255, 204) 'DS
    COLOR_LBLUE = RGB(204, 255, 255) 'NS
    COLOR_RULE_11_RED = RGB(255, 0, 102)
End Sub

Sub GetCellRGB()
    Dim cell As Range
    Dim colorValue As Long
    Dim r As Long, g As Long, b As Long

    ' Set the target cell
    Set cell = Selection

    ' Get the color value of the cell fill
    colorValue = cell.DisplayFormat.Interior.Color

    ' Extract RGB components
    r = colorValue Mod 256
    g = (colorValue \ 256) Mod 256
    b = (colorValue \ 65536) Mod 256

    ' Output to Immediate Window (Ctrl + G to view)
    Debug.Print "=RGB(" & r & "," & g & "," & b & ")"
End Sub

'================================================================================
' Procedure: FormatRosterAll_Optimized
'
' Purpose:
'   Formats the Manning "Lookahead" roster table (tblLookahead) using fast array-
'   based reads and dictionary lookups. Applies colour rules for:
'       - Trade / Local / Crane row classification
'       - Employee validation (unknown people)
'       - Shift code validation (DS / NS / other valid codes / invalid codes)
'       - Leave overlays (now status-aware: Approved vs Not Approved)
'       - Training overlays (Training vs TAFE)
'       - Conflict overlays (multiple allocations for same PID/date)
'       - Optional fatigue analysis via CheckFatigue
'
' Performance Strategy:
'   - Reads table once into arrays (dataArr/headerArr)
'   - Uses dictionaries to avoid repeated worksheet scans
'   - Writes formatting directly to cells in the table range
'
' Optional Flags:
'   bOnClose     – reserved / not currently used
'   bCleanLeave  – if True, clears roster content on leave days prior to filtdate
'   bFindFatigue – if True, runs CheckFatigue after formatting
'
' Dependencies / Globals:
'   - empDict, leaveDict, trainingDict (module/global scope)
'   - datavalDict (local to this procedure)
'   - vStopAll, vGoAll, StopAll, GoAll
'   - GetColours (populates COLOR_* constants)
'   - ShowTemporaryMessage, CloseTemporaryMessage, frmMsgRef (progress UI)
'   - Named ranges:
'       Manning:   rngManningFDate, rngManningDateRangeTo
'       Training:  rngTrainingPID, rngTrainingDate, rngTrainingFDate, rngDocType
'       Leave:     rngLeavePID, rngLeaveSDate, rngLeaveFDate, rngLeaveStatus
'       Constants: rngVal
'
' Leave Status Behaviour:
'   - Status "A" ? Approved ? COLOR_RULE_7_CYAN
'   - Status "N" ? Not Approved / Pending ? COLOR_RULE_12_PURPLE
'   - Anything else defaults to COLOR_RULE_7_CYAN (change if required)
'================================================================================
Public Sub FormatRosterAll_Optimized(Optional bOnClose As Boolean, _
                                     Optional bCleanLeave As Boolean, _
                                     Optional bFindFatigue As Boolean)

    Dim ws As Worksheet, wsEmp As Worksheet, wsTrn As Worksheet, wsLv As Worksheet, wsConst As Worksheet
    Dim TBL As ListObject, tblEmp As ListObject

    Dim r As Long, c As Long, i As Long
    Dim vPid As Variant, headerDate As Variant
    Dim vLoc, vPersonnel, vRole, empVal
    Dim matchFound As Boolean, hasTAFE As Boolean, hasTraining As Boolean
    Dim conflictDict As Object
    Dim datavalDict As Object ' local alias to avoid confusion with any global naming
    
    Dim cell As Range, rowRange As Range, targetCell As Range

    Dim trainingPID, trainingStart, trainingEnd, trainingType
    Dim leavePID, leaveStart, leaveEnd, leaveStatus
    Dim dataVal

    Dim colCount As Long, rowCount As Long
    Dim dataArr As Variant, headerArr As Variant
    Dim isTrade As Boolean

    Dim pidKey As String, DateKey As String
    Dim leaveRange As Variant, trainingRange As Variant

    Dim startTime As Double
    Dim endTime As Double

    Dim sDate As Date, filtdate As Date

    On Error GoTo ErrorHandle

    'startTime = Timer

    '-------------------------------------------------------------------------
    ' Progress UI
    '-------------------------------------------------------------------------
    ShowTemporaryMessage "Lookahead Formatting", "Formatting Cells", 0
    DoEvents

    '-------------------------------------------------------------------------
    ' Optional interactive prompt (only if caller didn't pass a value)
    ' Note: Optional Boolean params default to False when omitted; this IsEmpty
    ' check is legacy and may never trigger. Kept for backward compatibility.
    '-------------------------------------------------------------------------
    If IsEmpty(bCleanLeave) Then
        bCleanLeave = False
        If MsgBox("Do you want to clear the leave allocations on all?", vbYesNo) = vbYes Then
            bCleanLeave = True
        End If
    End If

    '-------------------------------------------------------------------------
    ' Disable heavy Excel behaviours (screen updating, events, calc etc.)
    '-------------------------------------------------------------------------
    If vStopAll = False Then StopAll

    '-------------------------------------------------------------------------
    ' Resolve references
    '-------------------------------------------------------------------------
    Set ws = ThisWorkbook.Sheets("Manning")
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsTrn = ThisWorkbook.Sheets("Training Bookings")
    Set wsLv = ThisWorkbook.Sheets("tbl_Vista_HR_Leave")
    Set wsConst = ThisWorkbook.Sheets("Constants")

    Set TBL = ws.ListObjects("tblLookahead")
    Set tblEmp = wsEmp.ListObjects("tblEmployees")

    '-------------------------------------------------------------------------
    ' Pull table values into arrays for speed
    '-------------------------------------------------------------------------
    dataArr = TBL.DataBodyRange.Value
    headerArr = TBL.HeaderRowRange.Value
    colCount = UBound(dataArr, 2)
    rowCount = UBound(dataArr, 1)

    '-------------------------------------------------------------------------
    ' Employee dictionary: valid employee names (uppercased)
    ' Used to highlight unknown people in the roster.
    '-------------------------------------------------------------------------
    If empDict Is Nothing Then
        Set empDict = CreateObject("Scripting.Dictionary")
        For Each cell In tblEmp.ListColumns("Person").DataBodyRange
            If Not IsEmpty(cell.Value) Then
                empDict(UCase$(CStr(cell.Value))) = True
            End If
        Next cell
    End If

    '-------------------------------------------------------------------------
    ' Date window (Manning)
    '   start date: rngManningFDate
    '   end date:   start + rngManningDateRangeTo - 1
    ' filtdate is used as a threshold for clearing content when bCleanLeave is True
    '-------------------------------------------------------------------------
    sDate = ws.Range("rngManningFDate").Value
    filtdate = sDate + ws.Range("rngManningDateRangeTo").Value - 1

    '-------------------------------------------------------------------------
    ' Training arrays (single read)
    '-------------------------------------------------------------------------
    trainingPID = wsTrn.Range("rngTrainingPID").Value
    trainingStart = wsTrn.Range("rngTrainingDate").Value
    trainingEnd = wsTrn.Range("rngTrainingFDate").Value
    trainingType = wsTrn.Range("rngDocType").Value

    '-------------------------------------------------------------------------
    ' Leave arrays (single read) - NOW INCLUDES STATUS
    '-------------------------------------------------------------------------
    leavePID = wsLv.Range("rngLeavePID").Value
    leaveStart = wsLv.Range("rngLeaveSDate").Value
    leaveEnd = wsLv.Range("rngLeaveFDate").Value
    leaveStatus = wsLv.Range("rngLeaveStatus").Value

    '-------------------------------------------------------------------------
    ' Validation list of allowed roster codes (Constants)
    '-------------------------------------------------------------------------
    dataVal = wsConst.Range("rngVal").Value

    '-------------------------------------------------------------------------
    ' Populate colour constants
    '-------------------------------------------------------------------------
    Call GetColours

    '-------------------------------------------------------------------------
    ' Build conflict dictionary: [pid|date] ? count of non-blank allocations
    ' Used to highlight multiple allocations for the same person on the same day.
    '-------------------------------------------------------------------------
    Set conflictDict = CreateObject("Scripting.Dictionary")

    For r = 1 To rowCount
        vPid = UCase$(CStr(dataArr(r, 8))) ' PID column (table-relative)
        For c = 13 To colCount
            If IsDate(headerArr(1, c)) Then
                If dataArr(r, c) <> "" Then
                    DateKey = vPid & "|" & CDate(headerArr(1, c))
                    conflictDict(DateKey) = conflictDict(DateKey) + 1
                End If
            End If
        Next c
    Next r

    '-------------------------------------------------------------------------
    ' Validation dictionary (allowed roster codes)
    ' If you already maintain a global datavalDict, keep using it.
    ' Here we create/populate a local dictionary to avoid name clashes.
    '-------------------------------------------------------------------------
    If datavalDict Is Nothing Then
        Set datavalDict = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(dataVal, 1)
            If Not datavalDict.Exists(dataVal(i, 1)) Then
                datavalDict.Add dataVal(i, 1), True
            End If
        Next i
    End If

    '-------------------------------------------------------------------------
    ' Leave dictionary: leaveDict(pid) = Collection of Array(Start, End, Status)
    ' Status expected values:
    '   "A" Approved
    '   "N" Not approved / pending
    '-------------------------------------------------------------------------
    If leaveDict Is Nothing Then
        Set leaveDict = CreateObject("Scripting.Dictionary")

        For i = 1 To UBound(leavePID, 1)

            pidKey = CStr(leavePID(i, 1))

            If IsDate(leaveStart(i, 1)) And IsDate(leaveEnd(i, 1)) Then

                If Not leaveDict.Exists(pidKey) Then
                    Set leaveDict(pidKey) = New Collection
                End If

                leaveDict(pidKey).Add Array( _
                    leaveStart(i, 1), _
                    leaveEnd(i, 1), _
                    UCase$(Trim$(CStr(leaveStatus(i, 1)))) _
                )

            End If
        Next i
    End If

    '-------------------------------------------------------------------------
    ' Training dictionary: trainingDict(pid) = Collection of Array(Start, End, Type)
    ' Type "TB" is used as TAFE marker (business rule).
    '-------------------------------------------------------------------------
    If trainingDict Is Nothing Then
        Set trainingDict = CreateObject("Scripting.Dictionary")

        For i = 1 To UBound(trainingPID, 1)

            pidKey = CStr(trainingPID(i, 1))

            If IsDate(trainingStart(i, 1)) And IsDate(trainingEnd(i, 1)) Then

                If Not trainingDict.Exists(pidKey) Then
                    Set trainingDict(pidKey) = New Collection
                End If

                trainingDict(pidKey).Add Array(trainingStart(i, 1), trainingEnd(i, 1), trainingType(i, 1))
            End If
        Next i
    End If

    '-------------------------------------------------------------------------
    ' Clear formatting on non-formula columns
    ' Assumption:
    '   Columns 7..12 contain formulas or should retain formatting.
    '   Everything outside this range gets reset prior to applying rules.
    '-------------------------------------------------------------------------
    For c = 1 To TBL.Range.Columns.Count
        If c < 7 Or c > 12 Then
            With TBL.DataBodyRange.Columns(c)
                .Interior.ColorIndex = xlNone
                .Font.Size = 10
                .Font.Bold = False
                .ShrinkToFit = True
            End With
        End If
    Next c

    '-------------------------------------------------------------------------
    ' Main formatting loop
    '-------------------------------------------------------------------------
    For r = 1 To rowCount

        Set rowRange = TBL.ListRows(r).Range

        vPid = Trim(UCase$(CStr(dataArr(r, 8))))

        vLoc = dataArr(r, 12)
        vPersonnel = Trim(UCase$(CStr(dataArr(r, 4))))
        vRole = UCase$(CStr(dataArr(r, 10)))
        empVal = vPersonnel
        isTrade = False

        '-------------------------------------------------------------
        ' Row classification rules (trade/local/crane)
        ' Trade-only rows use PID=Personnel logic (existing convention)
        '-------------------------------------------------------------
        If vPid = vPersonnel Then
            If UCase$(CStr(vLoc)) = "ROLE" Then
                isTrade = True
                rowRange.Cells(1, 4).Interior.Color = COLOR_RULE_3_YELLOW
            ElseIf UCase$(CStr(vLoc)) = "CRANE" Then
                rowRange.Cells(1, 4).Interior.Color = COLOR_RULE_10_BLUE
            End If
        End If

        ' Non-local with role mismatch (non-trade)
        If UCase$(CStr(vLoc)) <> "LOCAL" And UCase$(CStr(vLoc)) <> "LOCAL APP" And vRole <> vPersonnel And Not isTrade Then
            rowRange.Cells(1, 4).Interior.Color = COLOR_RULE_2_PURPLE
        End If

        ' Unknown personnel (not in employee list) (non-trade)
        If Not empDict.Exists(empVal) And Not isTrade Then
            rowRange.Cells(1, 4).Interior.Color = COLOR_RULE_4_ORANGE_DEEP
        End If

        '-------------------------------------------------------------
        ' Per-date logic (columns 13..end) driven by header row dates
        '-------------------------------------------------------------
        For c = 13 To colCount

            If Not IsDate(headerArr(1, c)) Then GoTo SkipColumn
            headerDate = CDate(headerArr(1, c))

            Set targetCell = rowRange.Cells(1, c)

            ' Trade overrides: any non-blank in a trade row is coloured as trade
            If isTrade And dataArr(r, c) <> "" Then
                targetCell.Interior.Color = COLOR_RULE_3_YELLOW
                GoTo SkipColumn
            End If

            '---------------------------------------------------------
            ' Shift code validation / primary colouring
            '---------------------------------------------------------
            If Len(Trim$(CStr(targetCell.Value))) > 0 Then

                If Not datavalDict.Exists(targetCell.Value) Then
                    ' Invalid code (not in Constants rngVal)
                    targetCell.Interior.Color = vbYellow
                Else
                    Select Case CStr(targetCell.Value)
                        Case "DS"
                            targetCell.Interior.Color = COLOR_GREEN
                        Case "NS"
                            targetCell.Interior.Color = COLOR_LBLUE
                        Case Else
                            ' Valid but not DS/NS (leave other valid codes yellow rule)
                            targetCell.Interior.Color = COLOR_RULE_8_YELLOW
                    End Select
                End If

            End If

            '---------------------------------------------------------
            ' Leave overlay (status-aware)
            ' leaveRange = Array(StartDate, EndDate, Status)
            '---------------------------------------------------------
            matchFound = False

            If leaveDict.Exists(vPid) Then
                For Each leaveRange In leaveDict(vPid)

                    If headerDate >= leaveRange(0) And headerDate <= leaveRange(1) Then

                        matchFound = True

                        Select Case leaveRange(2)
                            Case "A"
                                targetCell.Interior.Color = COLOR_RULE_7_CYAN
                            Case "N"
                                targetCell.Interior.Color = COLOR_RULE_12_PURPLE
                            Case Else
                                ' Unknown status ? default to approved colour
                                targetCell.Interior.Color = COLOR_RULE_7_CYAN
                        End Select

                        ' Clear contents in leave window if requested (only before filtdate)
                        If bCleanLeave Then
                            If headerDate < filtdate Then
                                targetCell.ClearContents
                            End If
                        End If

                        Exit For
                    End If

                Next leaveRange
            End If

            If matchFound Then GoTo SkipColumn

            '---------------------------------------------------------
            ' Training / TAFE overlay
            ' Training type "TB" is treated as TAFE.
            '---------------------------------------------------------
            hasTAFE = False
            hasTraining = False

            If trainingDict.Exists(vPid) Then
                For Each trainingRange In trainingDict(vPid)
                    If headerDate >= trainingRange(0) And headerDate <= trainingRange(1) Then
                        hasTraining = True
                        If UCase$(CStr(trainingRange(2))) = "TB" Then hasTAFE = True
                        Exit For
                    End If
                Next trainingRange
            End If

            If hasTAFE Then
                targetCell.Interior.Color = COLOR_RULE_5_BLUE
                GoTo SkipColumn
            ElseIf hasTraining Then
                targetCell.Interior.Color = COLOR_RULE_6_PURPLE
                GoTo SkipColumn
            End If

            '---------------------------------------------------------
            ' Conflict overlay
            ' Only applied where personnel != role and cell is non-blank
            ' conflictDict(pid|date) > 1 means multiple allocations.
            '---------------------------------------------------------
            If vPersonnel <> vRole And dataArr(r, c) <> "" Then
                DateKey = vPid & "|" & headerDate
                If conflictDict.Exists(DateKey) Then
                    If conflictDict(DateKey) > 1 Then
                        targetCell.Interior.Color = COLOR_RULE_9_RED
                    End If
                End If
            End If

SkipColumn:
        Next c

        '-------------------------------------------------------------
        ' Progress update (keeps UI responsive during long runs)
        '-------------------------------------------------------------
        With frmMsgRef.lbl1
            .Caption = "Formatting row " & r & " of " & rowCount & "..."
            DoEvents
        End With

    Next r

    '-------------------------------------------------------------------------
    ' Optional fatigue check (separate policy engine)
    '-------------------------------------------------------------------------
    If bFindFatigue Then
        CheckFatigue
    End If

CleanExit:
    Application.StatusBar = False
    CloseTemporaryMessage

    If vGoAll = False Then GoAll

    endTime = Timer
    'Debug.Print "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
    Exit Sub

ErrorHandle:
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.description, vbExclamation
    End If
    Resume CleanExit

End Sub



Sub CheckLAFatigue()
    CheckFatigue True
End Sub

'==========================================================
' CHECKFATIGUE – POLICY-DRIVEN FATIGUE CHECK
'
' HOW TO UPDATE IF POLICY CHANGES:
'   1) Shift codes (DS/NS)         -> Update GetShiftType() Select Case
'   2) Off codes (RNR/Leave/etc)   -> Update IsOffShift() Select Case
'   3) Max consecutive limits      -> Update MaxAllowedConsecutive()
'   4) Minimum R&R rules (off days)-> Update RequiredRR()
'
' NOTES:
'   - Consolidation is by PID across multiple rows.
'   - If duplicates exist for same PID/day:
'       NS overrides DS (policy decision: “worst case”)
'   - Minimum R&R is enforced ONLY if there is another worked shift
'     after the rest window (i.e., not at end of horizon).
'==========================================================

Sub CheckFatigue(Optional vFilt As Boolean)

    Dim ws As Worksheet
    Dim TBL As ListObject
    Dim dataArr As Variant

    Dim shiftDict As Object, rowMap As Object
    Dim shiftArr() As Variant

    Dim empRow As Variant, shiftOffset As Long
    Dim colCount As Long, pid As Variant
    Dim shiftVal As String, maxShifts As Long
    Dim shiftCol As Long, i As Long

    GetColours
    If Not vStopAll Then StopAll

    Set ws = Worksheets("Manning")
    Set TBL = ws.ListObjects("tblLookahead")

    Set shiftDict = CreateObject("Scripting.Dictionary")
    Set rowMap = CreateObject("Scripting.Dictionary")

    dataArr = TBL.DataBodyRange.Value
    colCount = UBound(dataArr, 2)

    '======================================================
    ' CONFIG ASSUMPTION:
    ' Shifts start at column 13 (M) in tblLookahead.
    ' If the table structure changes, update this constant.
    '======================================================
    maxShifts = colCount - 12

    ' Reset Filter column if requested (Field 11 is your filter flag)
    If vFilt Then
        For i = 1 To UBound(dataArr, 1)
            TBL.DataBodyRange.Cells(i, 11).Value = False
        Next i
    End If

    '======================================================
    ' 1) CONSOLIDATE SHIFTS PER PID
    ' Policy behaviour:
    '   - Off-codes do not populate shiftArr
    '   - If duplicates exist for same PID/day:
    '       NS overrides DS
    '
    ' If that override rule changes, edit the small block
    ' under “Night overrides Day”.
    '======================================================
    For empRow = 1 To UBound(dataArr, 1)

        pid = CStr(dataArr(empRow, 8)) ' PID in Col 8
        If pid <> "" And IsNumeric(pid) Then

            If Not shiftDict.Exists(pid) Then
                ReDim shiftArr(1 To maxShifts)
                shiftDict.Add pid, shiftArr
                rowMap.Add pid, New Collection
            Else
                shiftArr = shiftDict(pid)
            End If

            rowMap(pid).Add empRow

            For shiftOffset = 1 To maxShifts
                shiftCol = shiftOffset + 12
                shiftVal = UCase$(Trim$(CStr(dataArr(empRow, shiftCol))))

                If Not IsOffShift(shiftVal) Then

                    '-------------------------------
                    ' DUPLICATE DAY RESOLUTION RULE
                    ' Policy decision:
                    '   NS overrides DS
                    ' If this policy changes, edit here.
                    '-------------------------------
                    If shiftArr(shiftOffset) = "" Then
                        shiftArr(shiftOffset) = shiftVal
                    ElseIf shiftArr(shiftOffset) = "DS" And shiftVal = "NS" Then
                        shiftArr(shiftOffset) = "NS"
                    End If

                End If
            Next shiftOffset

            shiftDict(pid) = shiftArr
        End If
    Next empRow

    '======================================================
    ' 2) APPLY FATIGUE POLICY:
    '   - Detect consecutive streaks of Day or Night shifts
    '   - Check against MaxAllowedConsecutive()
    '   - Then enforce minimum R&R from RequiredRR()
    '
    ' If policy changes, do NOT touch this engine.
    ' Update the policy functions below instead.
    '======================================================
    Dim idx As Long, n As Long
    Dim t As String, t0 As String
    Dim streakStart As Long, streakEnd As Long
    Dim consec As Long, offDays As Long
    Dim j As Long, reqRR As Long

    For Each pid In shiftDict.Keys
        shiftArr = shiftDict(pid)
        n = UBound(shiftArr)

        ' Skip if this person has no shift data at all
        Dim hasData As Boolean
        hasData = False
        For idx = 1 To n
            If shiftArr(idx) <> "" Then
                hasData = True
                Exit For
            End If
        Next idx
        If Not hasData Then GoTo NextPID

        idx = 1
        Do While idx <= n

            t = GetShiftType(CStr(shiftArr(idx)))

            If t = "" Then
                idx = idx + 1

            Else
                ' Start streak
                t0 = t
                streakStart = idx
                consec = 0

                ' Count consecutive same-type shifts
                Do While idx <= n And GetShiftType(CStr(shiftArr(idx))) = t0
                    consec = consec + 1
                    idx = idx + 1
                Loop
                streakEnd = idx - 1

                ' Count consecutive off-days after streak (VBA-safe loop)
                offDays = 0
                j = idx
                Do While j <= n
                    If GetShiftType(CStr(shiftArr(j))) = "" Then
                        offDays = offDays + 1
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop

                ' A) Always flag if the max consecutive limit is exceeded
                If consec > MaxAllowedConsecutive(t0) Then
                    HighlightIfRowHasShiftData TBL, rowMap(pid), streakStart, streakEnd, dataArr, vFilt

                Else
                    ' B) Only enforce minimum R&R if there is another worked shift ahead
                    If j <= n Then
                        reqRR = RequiredRR(consec, t0)
                        If offDays < reqRR Then
                            HighlightIfRowHasShiftData TBL, rowMap(pid), streakStart, streakEnd, dataArr, vFilt
                        End If
                    End If
                End If

                ' Continue after off-window
                If j > idx Then idx = j

            End If
        Loop

NextPID:
    Next pid

    '======================================================
    ' 3) POST FILTER/SORT
    ' (Does not affect fatigue policy; leave as-is.)
    '======================================================
    With TBL
        On Error Resume Next
        .AutoFilter.ShowAllData
        On Error GoTo 0

        .Range.AutoFilter Field:=vManFiltCol, Criteria1:="TRUE"

        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 _
            key:=.ListColumns("Personnel").Range, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        With .Sort
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

ExitSub:
    If Not vGoAll Then GoAll
    Exit Sub

End Sub

'==========================================================
' POLICY DEFINITIONS (EDIT HERE WHEN POLICY CHANGES)
'==========================================================

' Off-days / non-work codes:
' If HR adds/removes a code that counts as an off-day, edit here.
Private Function IsOffShift(ByVal v As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))

    Select Case s
        Case "", "LEAVE", "TVL", "RNR", "PH", "DI", "DO", "FI", "FO", "TRN", "SD", "TAFE"
            IsOffShift = True
        Case Else
            IsOffShift = False
    End Select
End Function

' Shift typing:
' If shift codes change (e.g., DS/NS become D12/N12), edit here.
' Returns:
'   ""  = off / non-work (counts toward R&R)
'   "D" = day shift
'   "N" = night shift
Private Function GetShiftType(ByVal v As String) As String
    Dim s As String
    s = UCase$(Trim$(CStr(v)))

    If IsOffShift(s) Then
        GetShiftType = ""
        Exit Function
    End If

    Select Case s
        Case "DS": GetShiftType = "D"
        Case "NS": GetShiftType = "N"
        Case Else
            ' Unknown worked code: conservative default = Day
            GetShiftType = "D"
    End Select
End Function

' Max consecutive shifts allowed by type:
' If maximums change (e.g., Day 12, Night 8), edit here.
Private Function MaxAllowedConsecutive(ByVal t As String) As Long
    If t = "N" Then
        MaxAllowedConsecutive = 7
    Else
        MaxAllowedConsecutive = 10
    End If
End Function

' Minimum R&R (off-days) required after a consecutive streak:
' This is the main policy table.
' Update the mapping below if the fatigue policy changes.
Private Function RequiredRR(ByVal consec As Long, ByVal t As String) As Long
    If consec <= 0 Then
        RequiredRR = 0
        Exit Function
    End If

    If t = "N" Then
        ' Night shift R&R policy table
        Select Case consec
            Case Is <= 4: RequiredRR = 1
            Case 5:       RequiredRR = 2
            Case 6:       RequiredRR = 3
            Case Else:    RequiredRR = 4 ' 7+
        End Select
    Else
        ' Day shift R&R policy table
        Select Case consec
            Case Is <= 6: RequiredRR = 1
            Case 7:       RequiredRR = 2
            Case 8 To 9:  RequiredRR = 3
            Case Else:    RequiredRR = 4 ' 10+
        End Select
    End If
End Function

Private Sub HighlightIfRowHasShiftData(TBL As ListObject, empRows As Collection, startCol As Long, endCol As Long, _
                                       dataArr As Variant, vFilt As Boolean)

    Dim empRow As Variant, shiftOffset As Long
    Dim maxShifts As Long: maxShifts = UBound(dataArr, 2) - 12

    For Each empRow In empRows
        Dim hasShiftData As Boolean
        hasShiftData = False

        ' Only check shift columns (column 13 onwards in dataArr)
        For shiftOffset = startCol To endCol
            If Trim(dataArr(empRow, shiftOffset + 12)) <> "" Then
                hasShiftData = True
                Exit For
            End If
        Next shiftOffset

        If hasShiftData Then
            Dim i As Long
            For i = startCol To endCol
                TBL.DataBodyRange.Cells(empRow, i + 12).Interior.Color = COLOR_RULE_11_RED
                TBL.DataBodyRange.Cells(empRow, i + 12).Font.Bold = True
            Next i

            If vFilt Then
                TBL.DataBodyRange.Cells(empRow, 11).Value = True ' Column 11 = filter marker
            End If
        End If
    Next empRow
End Sub



Sub FormatRoster_CellOnly(ByVal targetCell As Range)
    Dim ws As Worksheet, wsEmp As Worksheet, wsTrn As Worksheet, wsLv As Worksheet
    Dim TBL As ListObject, tblEmp As ListObject
    Dim vPid As Variant, headerDate As Variant
    Dim vLoc, vPersonnel, vRole, empVal
    Dim empDict As Object, conflictDict As Object, leaveDict As Object, trainingDict As Object
    Dim cell As Range
    Dim trainingPID, trainingStart, trainingEnd, trainingType
    Dim leavePID, leaveStart, leaveEnd
    Dim pidKey As String, DateKey As String
    Dim leaveRange As Variant, trainingRange As Variant
    Dim r As Long, c As Long
    Dim dataArr As Variant, headerArr As Variant
    Dim rowCount As Long, colCount As Long
    Dim matchFound As Boolean, hasTAFE As Boolean, hasTraining As Boolean
    Dim isTrade As Boolean
    Dim startTime As Double, endTime As Double
    Dim i As Long, j As Long
    
    startTime = Timer
    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Sheets("Manning")
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsTrn = ThisWorkbook.Sheets("Training Bookings")
    Set wsLv = ThisWorkbook.Sheets("tbl_Vista_HR_Leave")
    Set TBL = ws.ListObjects("tblLookahead")
    Set tblEmp = wsEmp.ListObjects("tblEmployees")

    If Intersect(targetCell, TBL.DataBodyRange) Is Nothing Then
        MsgBox "Selected cell is not within 'tblLookahead'", vbExclamation
        GoTo CleanExit
    End If

    r = targetCell.row - TBL.DataBodyRange.Cells(1, 1).row + 1
    c = targetCell.Column - TBL.DataBodyRange.Cells(1, 1).Column + 1

    dataArr = TBL.DataBodyRange.Value
    headerArr = TBL.HeaderRowRange.Value
    rowCount = UBound(dataArr, 1)
    colCount = UBound(dataArr, 2)

    
    Set conflictDict = CreateObject("Scripting.Dictionary")
    
    Set trainingDict = CreateObject("Scripting.Dictionary")

    If IsEmpty(empDict) Then
        Set empDict = CreateObject("Scripting.Dictionary")
        For Each cell In tblEmp.ListColumns("Person").DataBodyRange
            If Not IsEmpty(cell.Value) Then empDict(UCase(CStr(cell.Value))) = True
        Next cell
    End If

    trainingPID = wsTrn.Range("rngTrainingPID").Value
    trainingStart = wsTrn.Range("rngTrainingDate").Value
    trainingEnd = wsTrn.Range("rngTrainingFDate").Value
    trainingType = wsTrn.Range("rngDocType").Value
    leavePID = wsLv.Range("rngLeavePID").Value
    leaveStart = wsLv.Range("rngLeaveSDate").Value
    leaveEnd = wsLv.Range("rngLeaveFDate").Value

    Call GetColours

    vPid = UCase(dataArr(r, 8))
    For i = 1 To rowCount
        If UCase(dataArr(i, 8)) = vPid Then
            For j = 13 To colCount
                If IsDate(headerArr(1, j)) Then
                    If dataArr(i, j) <> "" Then
                        DateKey = vPid & "|" & CDate(headerArr(1, j))
                        conflictDict(DateKey) = conflictDict(DateKey) + 1
                    End If
                End If
            Next j
        End If
    Next i

    For i = 1 To UBound(leavePID, 1)
        pidKey = CStr(leavePID(i, 1))
        If IsDate(leaveStart(i, 1)) And IsDate(leaveEnd(i, 1)) Then
            If Not leaveDict.Exists(pidKey) Then Set leaveDict(pidKey) = New Collection
            leaveDict(pidKey).Add Array(leaveStart(i, 1), leaveEnd(i, 1))
        End If
    Next i

    For i = 1 To UBound(trainingPID, 1)
        pidKey = CStr(trainingPID(i, 1))
        If IsDate(trainingStart(i, 1)) And IsDate(trainingEnd(i, 1)) Then
            If Not trainingDict.Exists(pidKey) Then Set trainingDict(pidKey) = New Collection
            trainingDict(pidKey).Add Array(trainingStart(i, 1), trainingEnd(i, 1), trainingType(i, 1))
        End If
    Next i

    If Not targetCell.HasFormula Then targetCell.Interior.ColorIndex = xlNone

    vLoc = dataArr(r, 12)
    vPersonnel = UCase(dataArr(r, 4))
    vRole = UCase(dataArr(r, 10))
    empVal = vPersonnel
    isTrade = (vPid = vPersonnel And UCase(vLoc) = "LOCAL")

    If Not IsDate(headerArr(1, c)) Then GoTo CleanExit
    headerDate = CDate(headerArr(1, c))

    If isTrade And dataArr(r, c) <> "" Then
        targetCell.Interior.Color = COLOR_RULE_3_YELLOW
        GoTo CleanExit
    End If

    If leaveDict.Exists(vPid) Then
        For Each leaveRange In leaveDict(vPid)
            If headerDate >= leaveRange(0) And headerDate <= leaveRange(1) Then
                targetCell.Interior.Color = COLOR_RULE_7_CYAN
                GoTo CleanExit
            End If
        Next
    End If

    hasTAFE = False: hasTraining = False
    If trainingDict.Exists(vPid) Then
        For Each trainingRange In trainingDict(vPid)
            If headerDate >= trainingRange(0) And headerDate <= trainingRange(1) Then
                hasTraining = True
                If UCase(trainingRange(2)) = "TAFEB" Then hasTAFE = True
                Exit For
            End If
        Next
    End If
    If hasTAFE Then
        targetCell.Interior.Color = COLOR_RULE_5_BLUE
        GoTo CleanExit
    ElseIf hasTraining Then
        targetCell.Interior.Color = COLOR_RULE_6_PURPLE
        GoTo CleanExit
    End If

    If vPersonnel <> vRole And dataArr(r, c) <> "" Then
        DateKey = vPid & "|" & headerDate
        If conflictDict.Exists(DateKey) Then
            If conflictDict(DateKey) > 1 Then
                targetCell.Interior.Color = COLOR_RULE_9_RED
            End If
        End If
    End If

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    endTime = Timer
    '.Print "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
End Sub

Sub Build()
    BuildRosterLegend
End Sub

Public Sub BuildRosterLegend(Optional ByVal targetSheetName As String = "Manning")

    Dim ws As Worksheet
    Dim startRow As Long, startCol As Long
    Dim r As Long

    ' Ensure colour constants are loaded
    Call GetColours

    Set ws = ThisWorkbook.Sheets(targetSheetName)

    ' Where the legend will be placed (adjust if needed)
    startRow = ws.Range("rngManningTotals").row + 2
    startCol = 2

    ' Clear existing legend area (safe buffer)
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 30, startCol + 3)).Clear

    ' Headers
    ws.Cells(startRow, startCol).Value = "Roster Colour Legend"
    ws.Cells(startRow, startCol).Font.Bold = True

    r = startRow + 2

    ' Helper to write a legend row
    Dim AddLegendRow As Object
    Set AddLegendRow = _
        CreateObject("Scripting.Dictionary") ' dummy, just to scope helper idea

    '-------------------------------------------------------------
    ' Legend entries (order = visual priority)
    '-------------------------------------------------------------
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_3_YELLOW, "Trade / Local Trade Row"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_10_BLUE, "Crane / Crane Trade Row"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_2_PURPLE, "Non-Local / Role Mismatch"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_4_ORANGE_DEEP, "Unknown Employee"): r = r + 1

    r = r + 1 ' spacer

    Call WriteLegendRow(ws, r, startCol, COLOR_GREEN, "Day Shift (DS)"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_LBLUE, "Night Shift (NS)"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_8_YELLOW, "Valid Code (Other than DS/NS)"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, vbYellow, "Invalid / Unrecognised Code"): r = r + 1

    r = r + 1 ' spacer

    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_7_CYAN, "Leave – Approved (Status A)"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_12_PURPLE, "Leave – Not Approved / Pending (Status N)"): r = r + 1

    r = r + 1 ' spacer

    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_5_BLUE, "TAFE"): r = r + 1
    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_6_PURPLE, "Training"): r = r + 1

    r = r + 1 ' spacer

    Call WriteLegendRow(ws, r, startCol, COLOR_RULE_9_RED, "Conflict (Multiple Allocations Same Day)"): r = r + 1

    ' Formatting
    ws.Columns(startCol).ColumnWidth = 3
    ws.Columns(startCol + 1).ColumnWidth = 38

End Sub

Private Sub WriteLegendRow(ws As Worksheet, _
                           ByVal rowNum As Long, _
                           ByVal colNum As Long, _
                           ByVal fillColor As Long, _
                           ByVal description As String)

    With ws.Cells(rowNum, colNum)
        .Interior.Color = fillColor
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Cells(rowNum, colNum + 1)
        .Value = description
        .VerticalAlignment = xlCenter
    End With

End Sub


