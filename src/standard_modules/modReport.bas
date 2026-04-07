Attribute VB_Name = "modReport"
Option Explicit
Option Compare Text

Sub GotoPerson()
Attribute GotoPerson.VB_ProcData.VB_Invoke_Func = "g\n14"
    ' Enable with Ctrl + G
    Dim sht As Worksheet
    Dim vStr As String
    Dim vRow As Long
    Dim cnt As Long
    Dim vRng As Range
    Dim actSheet As String
    Dim selectedColumn As Integer
    Dim selectedRow As Integer
    Dim vGoto As Boolean

    On Error GoTo ErrorHandle
    'Application.EnableEvents = False
    
    If vStopAll = False Then StopAll
    
    ' Set worksheet reference
    Set sht = Worksheets("Manning")

    ' Exit if selection is empty
    If IsEmpty(Application.Selection) Then Exit Sub
    
    ' Get active sheet name and selection details
    actSheet = Application.ActiveSheet.Name
    vStr = Application.Selection.Value
    selectedColumn = Application.Selection.Column
    selectedRow = Application.Selection.row

    ' Main decision structure
    Select Case actSheet
        Case "Report"
            Select Case selectedColumn
                Case 1
                    Call ResetSort
                    sht.Activate
                    sht.ListObjects("tblLookahead").Range.AutoFilter Field:=4, Criteria1:=vStr
                Case 2
                    vStr = Application.Selection.Offset(0, 1).Value
                    With Worksheets("Employee Inductions")
                        .Activate
                        .ListObjects("Table5").Range.AutoFilter.ShowAllData
                        .ListObjects("Table5").Range.AutoFilter Field:=2, Criteria1:=vStr
                    End With
                Case 3
                    With Worksheets("Employees")
                        .Activate
                        .ListObjects("tblEmployees").Range.AutoFilter Field:=17, Criteria1:=vStr
                    End With
                Case 30
                    With Worksheets("EMP_Inductions")
                        
                        Dim pt As PivotTable
                        Dim pf As PivotField
                        Dim val As String
                        val = Cells(Selection.row, 3).Value
                        .Activate
                        ' Reference the pivot table
                        Set pt = ActiveSheet.PivotTables("pvtEmpInductions")
                        Set pf = pt.PivotFields("[Employees].[PersonID].[PersonID]")
                        
                        ' Clear any existing filters
                        pf.ClearAllFilters
                        
                        ' Set the current page using the variable
                        pf.CurrentPageName = "[Employees].[PersonID].&[" & val & "]"
                        
                    End With
            End Select
        
        Case "Manning"
            vGoto = True
            Call ResetSort
            sht.Activate
            Select Case selectedColumn
                Case 1 To 6
                    sht.ListObjects("tblLookahead").Range.AutoFilter Field:=selectedColumn, Criteria1:=vStr
                
                Case 8 'Employee Table
                    With Worksheets("Employees")
                        .Activate
                        .ListObjects("tblEmployees").Range.AutoFilter Field:=17, Criteria1:=vStr
                    End With
                    vGoto = False
            End Select
            If vGoto = True Then Application.Goto LastSelection, True
            
        Case "Employees"
            If selectedColumn = 17 Then
                With Worksheets("Constants")
                    .Activate
                    .ListObjects("tblLocal").Range.AutoFilter Field:=1, Criteria1:=vStr
                    cnt = 0
                    For Each vRng In .ListObjects("tblLocal").DataBodyRange.Columns(1).Cells
                        If Not vRng.EntireRow.Hidden Then cnt = cnt + 1
                    Next vRng
                    If cnt = 0 Then
                        .ListObjects("tblLocal").DataBodyRange.Cells(.ListObjects("tblLocal").ListRows.Count + 1, 1).Value = "'" & vStr
                    End If
                End With
            End If
        
        Case "Project Report"
            Call ResetSort
            vStr = Worksheets("Project Report").Cells(selectedRow, 1).Value
            With sht.ListObjects("tblLookahead").Range
                .AutoFilter Field:=1, Criteria1:=vStr
                If selectedColumn >= 2 Then .AutoFilter Field:=2, Criteria1:=Worksheets("Project Report").Cells(selectedRow, 2).Value
                If selectedColumn >= 3 Then .AutoFilter Field:=3, Criteria1:=Worksheets("Project Report").Cells(selectedRow, 3).Value
            End With
            sht.Activate
    End Select

ExitSub:
    ' Cleanup
    Set sht = Nothing
    If vGoAll = False Then GoAll
    
    Application.EnableEvents = True
    Exit Sub
ErrorHandle:

    Resume ExitSub
End Sub


Sub PopulateReportTable()
    Const COL_EMP_ID As Long = 3
    Const COL_FIRST_DATE As Long = 4 ' Report table: column D
 
    Const LOOKAHEAD_EMP_COL As Long = 8 ' Column H in tblLookahead

    Dim wsReport As Worksheet, wsManning As Worksheet
    Dim tblReport As ListObject, tblLookahead As ListObject
    Dim dataReport As Variant, dataManning As Variant, empIDs As Variant
    Dim headerDates() As Variant, headerDatesFlat() As Variant
    Dim leaveDict As Object, trainingDict As Object
    Dim r As Long, c As Long, i As Long, colOffset As Long
    Dim reportDate As Date, empID As Variant
    Dim rowIndex As Long, result As Long, training As Long
    Dim manningFDateCol As Range
    Dim COL_LAST_DATE As Long  ' Report table: column
    Dim numDateCols As Long
    Dim parts() As String
    Dim lvKey As String, trnKey As String
    Dim cellVal As Variant
    Dim pmFlag As Boolean
    Dim rptWks As Long
    Dim RptSitePM As String
    Dim allManning As Variant

    On Error GoTo ErrorHandle
    
    If vStopAll = False Then StopAll

    ShowTemporaryMessage "Report Formatting", "Generating Report", 0
    DoEvents

    Application.EnableEvents = False
    Application.StatusBar = "Initializing..."

    Set wsReport = ThisWorkbook.Sheets("Report")
    Set wsManning = ThisWorkbook.Sheets("Manning")
    Set tblReport = wsReport.ListObjects("tblReport")
    Set tblLookahead = wsManning.ListObjects("tblLookahead")
    Set manningFDateCol = wsManning.Range("rngManningFDate")
    
    rptWks = wsReport.Range("rngRptWeeks").Value
    RptSitePM = wsReport.Range("RptSitePM").Value
    
    If rptWks >= 1 And rptWks <= 3 Then
        COL_LAST_DATE = 3 + (rptWks * 7)
    Else
        MsgBox "Invalid report weeks value. Please ensure it's set to 1, 2, or 3.", vbExclamation
        Exit Sub
    End If
    
    
    ' Load Report table
    dataReport = tblReport.DataBodyRange.Value

    ' Clear Data first
    Dim dataRange As Range
    Set dataRange = tblReport.DataBodyRange
    
    dataRange.Columns(COL_FIRST_DATE).Resize(, 24 - COL_FIRST_DATE + 1).ClearContents

    ' Header parsing
    Dim firstDateCol As Long
    firstDateCol = Application.Match(Format(manningFDateCol.Value, "dd/mm/yy"), tblLookahead.HeaderRowRange, 0)
    If IsError(firstDateCol) Then
        MsgBox "Start date not found in tblLookahead headers.", vbCritical
        GoTo ExitSub
    End If

    numDateCols = tblLookahead.ListColumns.Count - firstDateCol + 1
    ReDim headerDates(1 To 1, 1 To numDateCols)
    ReDim headerDatesFlat(1 To numDateCols)

    For i = 1 To numDateCols
        parts = Split(tblLookahead.HeaderRowRange.Cells(1, firstDateCol + i - 1).Value, "/")
        If UBound(parts) = 2 Then
            headerDates(1, i) = DateSerial(2000 + parts(2), parts(1), parts(0))
            headerDatesFlat(i) = headerDates(1, i)
        End If
    Next i

    ' Load dataManning and empIDs from tblLookahead
    dataManning = tblLookahead.DataBodyRange.Columns(firstDateCol).Resize(tblLookahead.DataBodyRange.Rows.Count, numDateCols).Value
    empIDs = tblLookahead.DataBodyRange.Columns(LOOKAHEAD_EMP_COL).Value

    ' Load full Lookahead table into array for PM/Site lookups
    allManning = tblLookahead.DataBodyRange.Value

    ' Cache Leave data
    Set leaveDict = CreateObject("Scripting.Dictionary")
    With ThisWorkbook.Sheets("tbl_Vista_HR_Leave").ListObjects("tbl_Vista_HR_Leave")
        Dim lvData As Variant: lvData = .DataBodyRange.Value
        Dim lvIDCol As Long: lvIDCol = .ListColumns("WalzAppID").Index
        Dim lvDateCol As Long: lvDateCol = .ListColumns("Date").Index
        For r = 1 To UBound(lvData)
            lvKey = lvData(r, lvIDCol) & "|" & Format(lvData(r, lvDateCol), "yyyymmdd")
            leaveDict(lvKey) = True
        Next r
    End With

    ' Cache Training data
    Set trainingDict = CreateObject("Scripting.Dictionary")
    With ThisWorkbook.Sheets("Training Bookings").ListObjects("Training_Bookings")
        Dim trnData As Variant: trnData = .DataBodyRange.Value
        Dim trnIDCol As Long: trnIDCol = .ListColumns("PersonID").Index
        Dim trnStartCol As Long: trnStartCol = .ListColumns("BookingDate").Index
        Dim trnEndCol As Long: trnEndCol = .ListColumns("Finish Date").Index
        For r = 1 To UBound(trnData)
            For i = trnData(r, trnStartCol) To trnData(r, trnEndCol)
                trnKey = trnData(r, trnIDCol) & "|" & Format(i, "yyyymmdd")
                trainingDict(trnKey) = True
            Next i
        Next r
    End With

    ' Loop through report rows
    
    Dim totalRows As Long
    totalRows = UBound(dataReport, 1)
    
    For r = 1 To totalRows
        empID = dataReport(r, COL_EMP_ID)
   
        
        pmFlag = (wsReport.Range("RptSitePM").Value = "PM")

        For c = COL_FIRST_DATE To COL_LAST_DATE
            reportDate = CDate(manningFDateCol.Value + c - COL_FIRST_DATE)

            ' Find matching column in header
            colOffset = 0
            For i = 1 To numDateCols
                If IsDate(headerDatesFlat(i)) And headerDatesFlat(i) = reportDate Then
                    colOffset = i
                    Exit For
                End If
            Next i
            If colOffset = 0 Then GoTo SkipCell

            ' Scan dataManning once to get rowIndex and count
            rowIndex = 0: result = 0
            For i = 1 To UBound(dataManning)
                If dataManning(i, colOffset) <> "" And empIDs(i, 1) = empID Then
                    result = result + 1
                    If rowIndex = 0 Then rowIndex = i
                End If
            Next i

            Dim keyDate As String: keyDate = Format(reportDate, "yyyymmdd")
            Dim lvKeyCheck As String: lvKeyCheck = empID & "|" & keyDate
            Dim trnKeyCheck As String: trnKeyCheck = empID & "|" & keyDate
            training = IIf(trainingDict.Exists(trnKeyCheck), 1, 0)

            ' Logic based on result and training
            Select Case True
                Case result = 0
                    If training > 0 Then
                        cellVal = "TAFE / TRAINING"
                    ElseIf leaveDict.Exists(lvKeyCheck) Then
                        cellVal = "Leave"
                    Else
                        cellVal = False
                    End If
                Case result > 1
                    cellVal = "Double"
                Case rowIndex = 0
                    cellVal = IIf(training = 0, "Leave", "Training")
                Case Else
                    Select Case RptSitePM
                        Case "PM"
                            ' Column 6 in tblLookahead = PM
                            cellVal = allManning(rowIndex, 1)
                    
                        Case "Site"
                            ' Column 1 in tblLookahead = Site
                            cellVal = allManning(rowIndex, 6)
                    
                        Case "Shift"
                            ' Use cached date/value block for shifts
                            cellVal = dataManning(rowIndex, colOffset)
                    End Select                    'cellVal = IIf(pmFlag, wsManning.Cells(rowIndex + 3, 1).Value, wsManning.Cells(rowIndex + 3, 6).Value)
            End Select

            ' Set value
            tblReport.DataBodyRange.Cells(r, c).Value = cellVal

SkipCell:
        Next c
    

        ' Update progress
        If r Mod 5 = 0 Or r = totalRows Then
            Application.StatusBar = "Populating report... " & Format(r / totalRows, "0%")
            DoEvents
        End If
    
    Next r

    Application.Calculate

    With ActiveWorkbook.SlicerCaches("Slicer_Blanks")
        .SlicerItems("FALSE").Selected = True
        .SlicerItems("TRUE").Selected = False
    End With

ExitSub:
    CloseTemporaryMessage
    Application.EnableEvents = True
    If vGoAll = False Then GoAll
    Application.StatusBar = "Report table populated successfully."
    If vGoAll = False Then GoAll
    'MsgBox "Report table populated successfully.", vbInformation
    Exit Sub

ErrorHandle:
    MsgBox "Error: " & Err.description, vbExclamation
    Resume ExitSub
End Sub

