Attribute VB_Name = "modRegressionSmoke"
Option Explicit

' Lightweight regression smoke suite for public entry-point Subs.
' - Logs PASS / FAIL / SKIP to the Immediate window.
' - Avoids interactive/destructive actions unless includeInteractive:=True.

Private Type TRegressionResult
    ProcName As String
    Outcome As String
    Details As String
End Type

Private mResults() As TRegressionResult
Private mCount As Long

Public Sub RunPublicEntryPointRegression(Optional ByVal includeInteractive As Boolean = False)
    ResetResults

    ' No-arg entry points
    ExecNoArg "modRoleGraph.RebuildRoleGraphFromLookahead"
    ExecNoArg "modManning.StopAll"
    ExecNoArg "modManning.GoAll"
    ExecNoArg "modManning.CheckLeaveAndFillColumn"
    ExecNoArg "modManning.CheckClashesAndFillColumn"
    ExecNoArg "modManning.HideCols"
    ExecNoArg "modManning.ResetManning"
    ExecNoArg "modManning.ResetSort"
    ExecNoArg "modManning.HideColsOrig"
    ExecNoArg "modManning.RolesOnly"
    ExecNoArg "modManning.RolesOnlyOld"
    ExecNoArg "modManning.FilterAndExportPM"
    ExecNoArg "modManning.RefreshAllQueries"
    ExecNoArg "modManning.FilterDuplicates"
    ExecNoArg "modManning.CleanupManning"
    ExecNoArg "modManning.CleanBlanks"
    ExecNoArg "modManning.FillRNR"
    ExecNoArg "modManning.FillRNR2"
    ExecNoArg "modManning.FilterNonGladstone"
    ExecNoArg "modManning.FilterRole"
    ExecNoArg "modManning.DeletePMForecast"
    ExecNoArg "modManning.DeletePMForecastPass"
    ExecNoArg "modManning.ImportPMData"
    ExecNoArg "modManning.NewRole"
    ExecNoArg "modManning.UpdateFilterCol_UnapprovedLeaveold"
    ExecNoArg "modManning.UpdateFilterCol_LeaveStatusFilter"
    ExecNoArg "modManning.FilterLookaheadByRequiredInductions"
    ExecNoArg "modFormatting.FormatRosterAll_Optimized"
    ExecNoArg "modFormatting.BuildRosterLegend"
    ExecNoArg "modCleanLookahead.Clean_tblLookahead_BlankRows"
    ExecNoArg "modSupport.PasteValuesOnlyIfCopied"
    ExecNoArg "modSupport.DisableAllConnectionAutoRefresh"
    ExecNoArg "modSupport.DisableAllQueryTables"
    ExecNoArg "modSupport.ClearHeaderMapCache"

    ' Potentially interactive procedures
    If includeInteractive Then
        ExecNoArg "modExport.ExportVBAModules"
        ExecNoArg "modExport.ImportVBAModules"
    Else
        RecordSkip "modExport.ExportVBAModules", "Skipped by default (interactive/path-dependent)."
        RecordSkip "modExport.ImportVBAModules", "Skipped by default (interactive/path-dependent)."
    End If

    ' Parameterized procedures
    TestAppAndSheetGuards
    TestTableHelpers
    TestSupportStateAndTimeline
    TestSupportLogError
    TestFormEntryPoints

    PrintSummary
End Sub

Private Sub TestAppAndSheetGuards()
    Dim ws As Worksheet
    Dim shState As TSheetGuardState

    On Error GoTo FailCase

    Set ws = EnsureScratchSheet()

    AppGuard_Begin False, "Regression smoke"
    AppGuard_End True

    shState = SheetGuard_Begin(ws, False)
    SheetGuard_End ws, shState

    RecordPass "modGuardsAndTables.AppGuard_Begin"
    RecordPass "modGuardsAndTables.AppGuard_End"
    RecordPass "modGuardsAndTables.SheetGuard_End"
    Exit Sub

FailCase:
    RecordFail "modGuardsAndTables.AppGuard_Begin", Err.Description
    RecordFail "modGuardsAndTables.AppGuard_End", Err.Description
    RecordFail "modGuardsAndTables.SheetGuard_End", Err.Description
    Err.Clear
End Sub

Private Sub TestTableHelpers()
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo FailCase

    Set ws = EnsureScratchSheet()
    Set lo = EnsureScratchTable(ws)

    ClearTableToHeaderOnly lo
    ClearTableRowsToHeaderOnly lo
    ResizeListObjectRowsExact lo, 1
    ArrayToTable lo, BuildSample2DArray(), False

    RecordPass "modGuardsAndTables.ClearTableToHeaderOnly"
    RecordPass "modGuardsAndTables.ClearTableRowsToHeaderOnly"
    RecordPass "modGuardsAndTables.ResizeListObjectRowsExact"
    RecordPass "modGuardsAndTables.ArrayToTable"
    Exit Sub

FailCase:
    RecordFail "modGuardsAndTables.ClearTableToHeaderOnly", Err.Description
    RecordFail "modGuardsAndTables.ClearTableRowsToHeaderOnly", Err.Description
    RecordFail "modGuardsAndTables.ResizeListObjectRowsExact", Err.Description
    RecordFail "modGuardsAndTables.ArrayToTable", Err.Description
    Err.Clear
End Sub

Private Sub TestSupportStateAndTimeline()
    Dim st As appState
    Dim timelineName As String

    On Error GoTo FailState

    PushAppState st, True, False, False
    PopAppState st

    RecordPass "modSupport.PushAppState"
    RecordPass "modSupport.PopAppState"

    timelineName = FirstTimelineCacheName(ThisWorkbook)
    If Len(timelineName) = 0 Then
        RecordSkip "modSupport.Timeline_ThisWeek", "No timeline cache available in workbook."
        RecordSkip "modSupport.Timeline_NextWeek", "No timeline cache available in workbook."
        RecordSkip "modSupport.SetTimelineDateRange", "No timeline cache available in workbook."
    Else
        Timeline_ThisWeek timelineName, ThisWorkbook
        Timeline_NextWeek timelineName, ThisWorkbook
        SetTimelineDateRange timelineName, Date, Date + 7, ThisWorkbook

        RecordPass "modSupport.Timeline_ThisWeek"
        RecordPass "modSupport.Timeline_NextWeek"
        RecordPass "modSupport.SetTimelineDateRange"
    End If
    Exit Sub

FailState:
    RecordFail "modSupport.PushAppState", Err.Description
    RecordFail "modSupport.PopAppState", Err.Description
    Err.Clear
End Sub

Private Sub TestSupportLogError()
    On Error GoTo FailCase

    LogError "RunPublicEntryPointRegression", "Smoke call", "N/A"
    RecordPass "modSupport.LogError"
    Exit Sub

FailCase:
    RecordFail "modSupport.LogError", Err.Description
    Err.Clear
End Sub

Private Sub TestFormEntryPoints()
    Dim frm As frmMessages

    On Error GoTo FailCase

    Set frm = New frmMessages
    frm.InitializeMessage "Regression", "Smoke", 1
    Unload frm

    RecordPass "frmMessages.InitializeMessage"
    Exit Sub

FailCase:
    RecordFail "frmMessages.InitializeMessage", Err.Description
    Err.Clear
End Sub

Private Sub ExecNoArg(ByVal procName As String)
    On Error GoTo FailCase

    Application.Run procName
    RecordPass procName
    Exit Sub

FailCase:
    RecordFail procName, Err.Description
    Err.Clear
End Sub

Private Function EnsureScratchSheet() As Worksheet
    Const SHEET_NAME As String = "zz_regression_scratch"

    On Error Resume Next
    Set EnsureScratchSheet = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If EnsureScratchSheet Is Nothing Then
        Set EnsureScratchSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureScratchSheet.Name = SHEET_NAME
    End If

    EnsureScratchSheet.Cells.Clear
    EnsureScratchSheet.Range("A1:B1").Value = Array("Key", "Value")
    EnsureScratchSheet.Range("A2:B3").Value = BuildSample2DArray()
End Function

Private Function EnsureScratchTable(ByVal ws As Worksheet) As ListObject
    Const TABLE_NAME As String = "tblRegressionScratch"

    Dim lo As ListObject

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)
        lo.Name = TABLE_NAME
    Else
        lo.Resize ws.Range("A1:B3")
    End If

    Set EnsureScratchTable = lo
End Function

Private Function BuildSample2DArray() As Variant
    Dim data(1 To 2, 1 To 2) As Variant

    data(1, 1) = "A"
    data(1, 2) = 1
    data(2, 1) = "B"
    data(2, 2) = 2

    BuildSample2DArray = data
End Function

Private Function FirstTimelineCacheName(ByVal wb As Workbook) As String
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In wb.SlicerCaches
        FirstTimelineCacheName = sc.Name
        Exit Function
    Next sc
    On Error GoTo 0
End Function

Private Sub ResetResults()
    ReDim mResults(1 To 1)
    mCount = 0
End Sub

Private Sub RecordPass(ByVal procName As String)
    RecordResult procName, "PASS", ""
End Sub

Private Sub RecordSkip(ByVal procName As String, ByVal reason As String)
    RecordResult procName, "SKIP", reason
End Sub

Private Sub RecordFail(ByVal procName As String, ByVal errText As String)
    RecordResult procName, "FAIL", errText
End Sub

Private Sub RecordResult(ByVal procName As String, ByVal outcome As String, ByVal details As String)
    mCount = mCount + 1
    ReDim Preserve mResults(1 To mCount)

    mResults(mCount).ProcName = procName
    mResults(mCount).Outcome = outcome
    mResults(mCount).Details = details
End Sub

Private Sub PrintSummary()
    Dim i As Long

    Debug.Print String(80, "-")
    Debug.Print "Public entry-point regression smoke summary";
    Debug.Print " (" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ")"
    Debug.Print String(80, "-")

    For i = 1 To mCount
        Debug.Print mResults(i).Outcome & " | " & mResults(i).ProcName & IIf(Len(mResults(i).Details) > 0, " | " & mResults(i).Details, "")
    Next i

    Debug.Print String(80, "-")
End Sub
