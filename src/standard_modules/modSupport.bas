Attribute VB_Name = "modSupport"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Public Const VK_CONTROL As Long = &H11
Public Const VK_SHIFT As Long = &H10
Public Const VK_MENU As Long = &H12   'Alt

Global FormulasEnabled As Boolean
Global CondFormats As Boolean
Global ShtMan As Worksheet
Global lRow As Long, lCol As Long
Global pssword As String
Global IsGod As Boolean
Global LastSelection As Range

Public Const vManPersonCol As Long = 4 'Personnel Column
Public Const vManPJCol As Long = 2 'PJ Number Col
Public Const vManCalCol As Long = 13 'First Calendar Column
Public Const vManLocalCol As Long = 12 'Local Column
Public Const vManFiltCol As Long = 11 'Double up Column
Public Const vManFirstRow As Long = 4 'First Manning Row
Public Const vManEmpNumCol As Long = 8 'Employee Number col
Public Const vManRoleCol As Long = 10 'Role Column
Public Const vManPOHCol As Long = 9 'POH Column
Public Const vManClassCol As Long = 7 'Classification Column
Public Const vPassword As String = "Lookahead2023"


Public frmMsgRef As frmMessages

' ===== App state guard =====
Public Type appState
    Calc As XlCalculation
    scr As Boolean
    evt As Boolean
    alerts As Boolean
End Type

Sub SaveActiveSheetAsPDF_WithSaveDialog(Optional strName As String, Optional userpick As Variant)
    ' Prompts user for a save location & filename, then saves the ActiveSheet as PDF.
    Dim suggestedName As String
    Dim fullPath As String

    On Error GoTo ErrHandler

    If strName <> "" Then
    
        strName = strName & ".pdf"
        suggestedName = strName
    Else
    ' Suggest a filename based on sheet name + timestamp
        suggestedName = ActiveSheet.Name & " - " & Format(Now, "yyyy-mm-dd hhmmss") & ".pdf"
    End If

    ' Ask user where to save (GetSaveAsFilename doesn't actually save; it only returns the chosen path)
    If userpick = "" Then
        userpick = Application.GetSaveAsFilename(InitialFileName:=suggestedName, _
                    FileFilter:="PDF Files (*.pdf), *.pdf", _
                    Title:="Save active sheet as PDF")
    End If

    ' If user cancelled, GetSaveAsFilename returns False
    If userpick = False Then Exit Sub

    fullPath = CStr(userpick)
    ' Ensure .pdf extension
    If LCase$(Right$(fullPath, 4)) <> ".pdf" Then fullPath = fullPath & ".pdf"

    ' If file exists, confirm overwrite (optional)
    If Dir(fullPath) <> "" Then
        If MsgBox("File already exists. Overwrite?", vbExclamation + vbYesNo, "Confirm Overwrite") <> vbYes Then
            Exit Sub
        End If
    End If

    ' Export the active sheet
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=fullPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    'MsgBox "Saved PDF to:" & vbCrLf & fullPath, vbInformation

    Exit Sub

ErrHandler:
    MsgBox "Error saving PDF: " & Err.description, vbExclamation
End Sub

Sub TimeSheetCalculations()
    Dim ws As Worksheet
    Dim t As Double
    Dim msg As String

    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        t = Timer
        ws.Calculate
        msg = msg & ws.Name & " - " & Format(Timer - t, "0.000") & " seconds" & vbNewLine
    Next ws

    Application.Calculation = xlCalculationAutomatic
    MsgBox msg, vbInformation, "Sheet Calculation Times"
End Sub


Sub ShowTemporaryMessage(strCaption As String, strMess As String, secondsToShow As Double)
    Set frmMsgRef = New frmMessages
    frmMsgRef.InitializeMessage strCaption, strMess, secondsToShow
End Sub

Sub CloseTemporaryMessage()
    On Error Resume Next
    If Not frmMsgRef Is Nothing Then
        If frmMsgRef.TimeElapsed >= frmMsgRef.DurationSeconds Then
            Unload frmMsgRef
            Set frmMsgRef = Nothing
        Else
            ' Reschedule next poll in 0.5 seconds
            Application.OnTime Now + (0.5 / 86400), "CloseTemporaryMessage"
        End If
    End If
End Sub

Function UnlockCode() As Boolean
    If IsGod = False Then
        FrmPassword.Show
        If pssword <> vPassword Then
            MsgBox "Wrong Password", vbCritical, "Enter Password"
            UnlockCode = False
        Else
            UnlockCode = True
            IsGod = True
        End If
    End If
    
End Function

Sub ChangeGreenCellsToDS()
    Dim cell As Range
    'Green 5296274
    'Yellow 65535
    'Blue 15773696
    Application.Calculation = xlCalculationManual
    ' Loop through each selected cell
    For Each cell In Selection
        ' Check if the cell's color is green (RGB value for green can vary)
        Select Case cell.Interior.Color
        
        Case 5296274 'Green 5296274
            cell.Value = "DS"
        Case 65535 'Yellow 65535
            cell.Value = "RNR"
        Case 15773696 'Blue 15773696
            cell.Value = "NS"
        End Select
    Next cell
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub ChangeDStoGreen()
    Dim cell As Range
    'Green 5296274
    'Yellow 65535
    'Blue 15773696
    Application.Calculation = xlCalculationManual
    ' Loop through each selected cell
    For Each cell In Selection
        ' Check if the cell's color is green (RGB value for green can vary)
        Select Case cell.Value
        
        Case ""
            cell.Interior.Color = 12566463
        Case "DS"
         'Green 5296274
            cell.Interior.Color = 5296274
            cell.Font.Size = 12
        Case "RNR"
         'Yellow 65535
            cell.Interior.Color = 65535
            cell.Font.Size = 8
        Case "NS"
        'Blue 15773696
            cell.Interior.Color = 15773696
            cell.Font.Size = 12
        End Select
    Next cell
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Public Function getColour(target As Range)
    
        getColour = target.Interior.Color

End Function

Public Sub PasteValuesOnlyIfCopied()
Attribute PasteValuesOnlyIfCopied.VB_ProcData.VB_Invoke_Func = "V\n14"
    ' Check if a copy or cut operation is active
    If Application.CutCopyMode = xlCopy Then
        If TypeName(Selection) = "Range" Then
            Selection.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Else
            MsgBox "Please select a valid range to paste into.", vbExclamation
        End If
    Else
        MsgBox "Nothing has been copied. Please copy a range first.", vbInformation
    End If
End Sub



Public Sub PushAppState(st As appState, Optional manualCalc As Boolean = True, _
                         Optional noScreen As Boolean = True, Optional noEvents As Boolean = True, _
                         Optional noAlerts As Boolean = True)
    With Application
        st.Calc = .Calculation
        st.scr = .ScreenUpdating
        st.evt = .EnableEvents
        st.alerts = .DisplayAlerts
        If manualCalc Then .Calculation = xlCalculationManual
        If noScreen Then .ScreenUpdating = False
        If noEvents Then .EnableEvents = False
        If noAlerts Then .DisplayAlerts = False
    End With
End Sub

Public Sub PopAppState(st As appState)
    With Application
        .Calculation = st.Calc
        .ScreenUpdating = st.scr
        .EnableEvents = st.evt
        .DisplayAlerts = st.alerts
    End With
End Sub

Public Function ColLetterToNumber(colLetter As String) As Long
    ColLetterToNumber = Range(UCase$(Trim$(colLetter)) & "1").Column
End Function

Public Sub DisableAllConnectionAutoRefresh()

    Dim c As WorkbookConnection

    For Each c In ThisWorkbook.Connections
        
        On Error Resume Next
        
        ' Not all connections support this
        c.RefreshWithRefreshAll = False
        
        ' OLEDB connections (SQL Server, etc.)
        If c.Type = xlConnectionTypeOLEDB Then
            c.OLEDBConnection.BackgroundQuery = False
            c.OLEDBConnection.EnableRefresh = False
            c.OLEDBConnection.RefreshOnFileOpen = False
        End If

        ' ODBC connections
        If c.Type = xlConnectionTypeODBC Then
            c.ODBCConnection.BackgroundQuery = False
            c.ODBCConnection.EnableRefresh = False
            c.ODBCConnection.RefreshOnFileOpen = False
        End If
        
        On Error GoTo 0
        
    Next c

End Sub

Public Sub DisableAllQueryTables()
    Dim qt As QueryTable
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        For Each qt In ws.QueryTables
            qt.RefreshOnFileOpen = False
            qt.BackgroundQuery = False
            qt.EnableRefresh = False
        Next qt
    Next ws
End Sub

Public Sub Timeline_ThisWeek(ByVal timelineCacheName As String, Optional wb As Workbook)
    Dim today As Date, startDate As Date, endDate As Date

    today = DateSerial(Year(Date), Month(Date), Day(Date))
    startDate = today - Weekday(today, vbMonday) + 1
    endDate = startDate + 6

    SetTimelineDateRange timelineCacheName, startDate, endDate, True, True, wb
End Sub

Public Sub Timeline_NextWeek(ByVal timelineCacheName As String, Optional wb As Workbook)
    Dim today As Date, startDate As Date, endDate As Date

    today = DateSerial(Year(Date), Month(Date), Day(Date))
    startDate = today - Weekday(today, vbMonday) + 8   ' Next Monday
    endDate = startDate + 6                            ' Next Sunday

    SetTimelineDateRange timelineCacheName, startDate, endDate, True, True, wb
End Sub

'===========================================================
' Generic Timeline Date Range Setter (locale-safe, module-safe)
'===========================================================
Public Sub SetTimelineDateRange( _
    ByVal timelineCacheName As String, _
    ByVal dStart As Date, _
    ByVal dEnd As Date, _
    Optional ByVal forceDaysLevel As Boolean = True, _
    Optional ByVal refreshConnectedPivots As Boolean = True, _
    Optional ByVal wb As Workbook _
)
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim pt As PivotTable
    Dim vStart As Double, vEnd As Double

    ' Default workbook context
    If wb Is Nothing Then Set wb = ActiveWorkbook

    Set sc = wb.SlicerCaches(timelineCacheName)

    ' Guard rails
    If sc.SlicerCacheType <> xlTimeline Then
        Err.Raise vbObjectError + 513, "SetTimelineDateRange", _
            "SlicerCache '" & timelineCacheName & "' is not a Timeline."
    End If

    If dEnd < dStart Then
        Err.Raise vbObjectError + 514, "SetTimelineDateRange", _
            "End date cannot be before start date."
    End If

    ' Force Excel serial numbers (kills dd/mm vs mm/dd issues)
    vStart = CDbl(DateSerial(Year(dStart), Month(dStart), Day(dStart)))
    vEnd = CDbl(DateSerial(Year(dEnd), Month(dEnd), Day(dEnd)))

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Optional: force UI to Days level so week ranges display correctly
    If forceDaysLevel Then
        Set sl = sc.Slicers(1)
        On Error Resume Next
        sl.TimelineViewState.Level = xlTimelineLevelDays
        On Error GoTo 0
    End If

    sc.ClearManualFilter
    sc.TimelineState.SetFilterDateRange startDate:=vStart, endDate:=vEnd

    If refreshConnectedPivots Then
        For Each pt In sc.PivotTables
            pt.RefreshTable
        Next pt
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'-------------------------------------------------------
' HeaderMapCached
'
' Purpose:
'   Returns a dictionary mapping header names to their
'   column index within a ListObject.
'
'   Example:
'       dict("Role") = 1
'       dict("Local / Away") = 2
'
' Features:
'   - Cached per table for performance
'   - Case-insensitive lookup
'   - Ignores blank headers
'   - First duplicate header wins
'
' Inputs:
'   - lo: target ListObject
'
' Outputs:
'   - Scripting.Dictionary (late-bound)
'
' Notes:
'   - Cache is keyed by Workbook + Worksheet + Table Name
'   - Safe for repeated calls in large procedures
'-------------------------------------------------------
Public Function HeaderMapCached(ByVal lo As ListObject) As Object

    On Error GoTo ErrHandler

    Static dictCache As Object   ' cache across calls

    Dim dictHeaders As Object
    Dim cacheKey As String

    Dim hdrArr As Variant
    Dim c As Long
    Dim headerName As String

    Const PROC_NAME As String = "HeaderMapCached"

    '-------------------------------------------------------
    ' Validate input
    '-------------------------------------------------------
    If lo Is Nothing Then
        Err.Raise vbObjectError + 1200, PROC_NAME, "ListObject is Nothing."
    End If

    If lo.HeaderRowRange Is Nothing Then
        Err.Raise vbObjectError + 1201, PROC_NAME, "ListObject has no HeaderRowRange."
    End If

    '-------------------------------------------------------
    ' Initialise cache store
    '-------------------------------------------------------
    If dictCache Is Nothing Then
        Set dictCache = CreateObject("Scripting.Dictionary")
        dictCache.CompareMode = vbTextCompare
    End If

    '-------------------------------------------------------
    ' Build unique cache key
    '-------------------------------------------------------
    cacheKey = lo.Parent.Parent.Name & "|" & lo.Parent.Name & "|" & lo.Name

    '-------------------------------------------------------
    ' Return cached version if exists
    '-------------------------------------------------------
    If dictCache.Exists(cacheKey) Then
        Set HeaderMapCached = dictCache(cacheKey)
        Exit Function
    End If

    '-------------------------------------------------------
    ' Build new header map
    '-------------------------------------------------------
    Set dictHeaders = CreateObject("Scripting.Dictionary")
    dictHeaders.CompareMode = vbTextCompare

    hdrArr = lo.HeaderRowRange.Value2

    ' HeaderRowRange is always 1 row
    For c = 1 To UBound(hdrArr, 2)

        headerName = Trim$(NzText(hdrArr(1, c)))

        If Len(headerName) > 0 Then
            ' First instance wins (ignore duplicates)
            If Not dictHeaders.Exists(headerName) Then
                dictHeaders.Add headerName, c
            End If
        End If

    Next c

    '-------------------------------------------------------
    ' Store in cache
    '-------------------------------------------------------
    dictCache.Add cacheKey, dictHeaders

    '-------------------------------------------------------
    ' Return result
    '-------------------------------------------------------
    Set HeaderMapCached = dictHeaders
    Exit Function

ErrHandler:
    Err.Raise Err.Number, PROC_NAME, Err.description

End Function


'-------------------------------------------------------
' ClearHeaderMapCache
'
' Purpose:
'   Clears the cached header maps.
'
' Use When:
'   - Headers are modified
'   - Tables are resized/rebuilt
'-------------------------------------------------------
Public Sub ClearHeaderMapCache()

    Static dictCache As Object

    On Error Resume Next

    If Not dictCache Is Nothing Then
        dictCache.RemoveAll
    End If

    On Error GoTo 0

End Sub

'-------------------------------------------------------
' NzText
'
' Purpose:
'   Safely converts any Variant to a trimmed String.
'
' Behaviour:
'   - Returns "" for:
'       * Empty
'       * Null
'       * Error values (e.g. #N/A, #VALUE!)
'   - Converts numbers/dates to String
'   - Trims leading/trailing whitespace
'
' Inputs:
'   - v: Variant value
'
' Outputs:
'   - String (never raises an error)
'
' Notes:
'   - Designed for dictionary keys, joins, and comparisons
'   - Avoids Type Mismatch errors from CStr on Error values
'-------------------------------------------------------
Public Function NzText(ByVal v As Variant) As String

    On Error GoTo SafeExit

    '-------------------------------------------------------
    ' Handle invalid/empty states first (fast exit)
    '-------------------------------------------------------
    If IsError(v) Then GoTo SafeExit
    If IsNull(v) Then GoTo SafeExit
    If IsEmpty(v) Then GoTo SafeExit

    '-------------------------------------------------------
    ' Convert to string and trim
    '-------------------------------------------------------
    NzText = Trim$(CStr(v))
    Exit Function

SafeExit:
    NzText = vbNullString

End Function

'-------------------------------------------------------
' LogError
'
' Purpose:
'   Logs errors to a central worksheet "zz_Log"
'
' Captures:
'   - Timestamp
'   - Procedure Name
'   - Error Number
'   - Error Description
'   - Workbook Name
'   - Worksheet Name (if available)
'
' Inputs:
'   - procName: Name of procedure raising/logging error
'   - errNumber: Err.Number
'   - errDescription: Err.Description
'
' Notes:
'   - Safe: will not crash calling code
'   - Auto-creates log sheet if missing
'   - Append-only logging
'-------------------------------------------------------
Public Sub LogError(ByVal procName As String, _
                    ByVal errNumber As Long, _
                    ByVal errDescription As String)

    On Error GoTo SafeExit

    Dim wsLog As Worksheet
    Dim nextRow As Long

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Const LOG_SHEET_NAME As String = "zz_Log"

    '-------------------------------------------------------
    ' Get or create log sheet
    '-------------------------------------------------------
    Set wsLog = GetOrCreateLogSheet(wb, LOG_SHEET_NAME)

    '-------------------------------------------------------
    ' Find next row
    '-------------------------------------------------------
    If Application.WorksheetFunction.CountA(wsLog.Cells) = 0 Then
        nextRow = 2
    Else
        nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row + 1
    End If

    '-------------------------------------------------------
    ' Write log entry
    '-------------------------------------------------------
    With wsLog
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = procName
        .Cells(nextRow, 3).Value = errNumber
        .Cells(nextRow, 4).Value = errDescription
        .Cells(nextRow, 5).Value = wb.Name

        ' Attempt to capture active sheet (best effort only)
        On Error Resume Next
        .Cells(nextRow, 6).Value = Application.ActiveSheet.Name
        On Error GoTo 0
    End With

SafeExit:
    ' Never raise errors from logger
End Sub


'-------------------------------------------------------
' GetOrCreateLogSheet
'
' Purpose:
'   Returns the log worksheet, creating it if missing
'
' Notes:
'   - Creates standard headers
'   - Keeps sheet hidden (optional)
'-------------------------------------------------------
Private Function GetOrCreateLogSheet(ByVal wb As Workbook, _
                                     ByVal sheetName As String) As Worksheet

    On Error GoTo CreateSheet

    Set GetOrCreateLogSheet = wb.Worksheets(sheetName)
    Exit Function

CreateSheet:
    On Error GoTo 0

    Set GetOrCreateLogSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

    With GetOrCreateLogSheet
        .Name = sheetName

        ' Headers
        .Cells(1, 1).Value = "Timestamp"
        .Cells(1, 2).Value = "Procedure"
        .Cells(1, 3).Value = "Error Number"
        .Cells(1, 4).Value = "Description"
        .Cells(1, 5).Value = "Workbook"
        .Cells(1, 6).Value = "Worksheet"

        ' Basic formatting
        .Rows(1).Font.Bold = True

        ' Optional: keep hidden from users
        .Visible = xlSheetVeryHidden
    End With

End Function

'-------------------------------------------------------
' HeaderAsDDMMYY
'
' Purpose:
'   Converts a header value into strict text format DD/MM/YY
'
' Behaviour:
'   - If value is a real Excel date, formats as DD/MM/YY
'   - If value is already text resembling a date, attempts conversion
'   - Otherwise returns trimmed text unchanged
'
' Notes:
'   - Output is always text, not a numeric Excel date
'   - Prevents header mismatch such as:
'       01/10/26 <> 10/01/2026
'-------------------------------------------------------
Public Function HeaderAsDDMMYY(ByVal v As Variant) As String

    On Error GoTo SafeExit

    Dim s As String
    Dim d As Date
    Dim parts() As String

    If IsError(v) Then GoTo SafeExit
    If IsNull(v) Then GoTo SafeExit
    If IsEmpty(v) Then GoTo SafeExit

    '-------------------------------------------------------
    ' If Excel is storing it as a real date/serial, format directly
    '-------------------------------------------------------
    If IsDate(v) Then
        HeaderAsDDMMYY = Format$(CDate(v), "dd/mm/yy")
        Exit Function
    End If

    '-------------------------------------------------------
    ' Try to interpret text dates safely as DD/MM/YY or DD/MM/YYYY
    '-------------------------------------------------------
    s = Trim$(CStr(v))
    If Len(s) = 0 Then GoTo SafeExit

    s = Replace$(s, "-", "/")
    parts = Split(s, "/")

    If UBound(parts) = 2 Then
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            d = DateSerial(NormaliseYear(CLng(parts(2))), CLng(parts(1)), CLng(parts(0)))
            HeaderAsDDMMYY = Format$(d, "dd/mm/yy")
            Exit Function
        End If
    End If

    '-------------------------------------------------------
    ' Fallback: return cleaned text unchanged
    '-------------------------------------------------------
    HeaderAsDDMMYY = s
    Exit Function

SafeExit:
    HeaderAsDDMMYY = vbNullString

End Function


'-------------------------------------------------------
' NormaliseYear
'
' Purpose:
'   Converts 2-digit years to 4-digit years for DateSerial
'
' Notes:
'   - 00 to 29 => 2000 to 2029
'   - 30 to 99 => 1930 to 1999
'   Adjust if your workbook requires a different rule
'-------------------------------------------------------
Public Function NormaliseYear(ByVal y As Long) As Long

    If y < 100 Then
        If y <= 29 Then
            NormaliseYear = 2000 + y
        Else
            NormaliseYear = 1900 + y
        End If
    Else
        NormaliseYear = y
    End If

End Function

'-------------------------------------------------------
' IsValueBlank
'
' Purpose:
'   Returns True when a value should be treated as blank.
'-------------------------------------------------------
Public Function IsValueBlank(ByVal v As Variant) As Boolean

    On Error GoTo BlankValue

    If IsError(v) Then GoTo BlankValue
    If IsNull(v) Then GoTo BlankValue
    If IsEmpty(v) Then GoTo BlankValue

    IsValueBlank = (Len(Trim$(CStr(v))) = 0)
    Exit Function

BlankValue:
    IsValueBlank = True

End Function

'-------------------------------------------------------
' HeaderToDate
'
' Purpose:
'   Converts a tblLookahead header value into a real Excel date.
'
' Behaviour:
'   - If already a real date/serial, returns Date value
'   - If text in dd/mm/yy or dd/mm/yyyy form, converts safely
'   - Raises an error if header cannot be converted
'
' Notes:
'   - Intended for AU-style day/month/year headers
'   - Returns a true Date, not text
'-------------------------------------------------------
Public Function HeaderToDate(ByVal v As Variant) As Date

    On Error GoTo ErrHandler

    Dim s As String
    Dim parts() As String
    Dim dd As Long
    Dim mm As Long
    Dim yy As Long

    Const PROC_NAME As String = "HeaderToDate"

    If IsError(v) Then Err.Raise vbObjectError + 1200, PROC_NAME, "Header contains an error value."
    If IsNull(v) Then Err.Raise vbObjectError + 1201, PROC_NAME, "Header contains Null."
    If IsEmpty(v) Then Err.Raise vbObjectError + 1202, PROC_NAME, "Header is empty."

    ' Real Excel date / serial
    If IsDate(v) Then
        HeaderToDate = DateValue(CDate(v))
        Exit Function
    End If

    ' Try explicit dd/mm/yy or dd/mm/yyyy text parsing
    s = Trim$(CStr(v))
    s = Replace$(s, "-", "/")

    If Len(s) = 0 Then
        Err.Raise vbObjectError + 1203, PROC_NAME, "Header text is blank."
    End If

    parts = Split(s, "/")
    If UBound(parts) <> 2 Then
        Err.Raise vbObjectError + 1204, PROC_NAME, "Header '" & s & "' is not a valid dd/mm/yy date."
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        Err.Raise vbObjectError + 1205, PROC_NAME, "Header '" & s & "' contains non-numeric date parts."
    End If

    dd = CLng(parts(0))
    mm = CLng(parts(1))
    yy = NormaliseYear(CLng(parts(2)))

    HeaderToDate = DateSerial(yy, mm, dd)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, PROC_NAME, Err.description
End Function

'-------------------------------------------------------
' ValueToDateSerial
'
' Purpose:
'   Converts a value to an Excel date serial (Long).
'
' Handles:
'   - Real Excel date serial numbers
'   - VBA Date values
'   - Date-like text such as dd/mm/yy or dd/mm/yyyy
'-------------------------------------------------------
Public Function ValueToDateSerial(ByVal v As Variant) As Long

    On Error GoTo ErrHandler

    Dim s As String
    Dim parts() As String
    Dim dd As Long
    Dim mm As Long
    Dim yy As Long

    If IsError(v) Then Err.Raise vbObjectError + 3000, "ValueToDateSerial", "Error value cannot be converted to date."
    If IsNull(v) Then Err.Raise vbObjectError + 3001, "ValueToDateSerial", "Null value cannot be converted to date."
    If IsEmpty(v) Then Err.Raise vbObjectError + 3002, "ValueToDateSerial", "Empty value cannot be converted to date."

    '-------------------------------------------------------
    ' If already numeric, treat as Excel serial date
    '-------------------------------------------------------
    If IsNumeric(v) Then
        ValueToDateSerial = CLng(v)
        Exit Function
    End If

    '-------------------------------------------------------
    ' If VBA recognises it as a date, convert to serial
    '-------------------------------------------------------
    If IsDate(v) Then
        ValueToDateSerial = CLng(DateValue(CDate(v)))
        Exit Function
    End If

    '-------------------------------------------------------
    ' Otherwise parse as dd/mm/yy or dd/mm/yyyy text
    '-------------------------------------------------------
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        Err.Raise vbObjectError + 3003, "ValueToDateSerial", "Blank text cannot be converted to date."
    End If

    s = Replace$(s, "-", "/")
    parts = Split(s, "/")

    If UBound(parts) <> 2 Then
        Err.Raise vbObjectError + 3004, "ValueToDateSerial", "Invalid date text: " & s
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        Err.Raise vbObjectError + 3005, "ValueToDateSerial", "Non-numeric date text: " & s
    End If

    dd = CLng(parts(0))
    mm = CLng(parts(1))
    yy = CLng(parts(2))

    If yy < 100 Then
        If yy <= 29 Then
            yy = 2000 + yy
        Else
            yy = 1900 + yy
        End If
    End If

    ValueToDateSerial = CLng(DateSerial(yy, mm, dd))
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "ValueToDateSerial", Err.description

End Function

Public Function RoleKeyText(ByVal v As Variant) As String
    RoleKeyText = UCase$(Trim$(NzText(v)))
End Function

