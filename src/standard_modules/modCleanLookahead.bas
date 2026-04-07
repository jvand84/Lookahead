Attribute VB_Name = "modCleanLookahead"
Option Explicit

'================================================================================
' Procedure : Clean_tblLookahead_BlankRows
' Purpose   : Remove all tblLookahead rows where Column 1 is blank
'             (including whitespace-only values), using staged delete logic.
'
' Rules     :
'   - Blank detection is based only on Column 1
'   - Rows are collected first, then deleted in reverse order
'   - Header row is always preserved
'   - Final state = valid data rows + exactly one blank spare row
'
' Dependencies:
'   - AppGuard_Begin / AppGuard_End
'   - SheetGuard_Begin / SheetGuard_End
'   - FindListObjectByName
'   - ResizeListObjectRowsExact
'   - LogError
'================================================================================
Public Sub Clean_tblLookahead_BlankRows()

    Const PROC_NAME As String = "Clean_tblLookahead_BlankRows"
    
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim shState As TSheetGuardState
    
    Dim col1Arr As Variant
    Dim deleteIdx() As Long
    
    Dim rowCount As Long
    Dim delCount As Long
    Dim keepCount As Long
    Dim targetRows As Long
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    '--------------------------------------------------------------------------
    ' Begin application guard
    '--------------------------------------------------------------------------
    AppGuard_Begin True, "Cleaning tblLookahead blank rows...", True
    
    '--------------------------------------------------------------------------
    ' Locate table
    '--------------------------------------------------------------------------
    Set lo = FindListObjectByName("tblLookahead")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 1000, PROC_NAME, _
                  "Table 'tblLookahead' was not found."
    End If
    
    Set ws = lo.Parent
    
    '--------------------------------------------------------------------------
    ' Begin sheet guard
    '--------------------------------------------------------------------------
    shState = SheetGuard_Begin(ws)
    
    '--------------------------------------------------------------------------
    ' If table has no data rows, force exactly one blank spare row
    '--------------------------------------------------------------------------
    If lo.DataBodyRange Is Nothing Then
        ResizeListObjectRowsExact lo, 1
        
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Rows(1).ClearContents
        End If
        
        GoTo SafeExit
    End If
    
    '--------------------------------------------------------------------------
    ' Read Column 1 only into memory for efficient scan
    '--------------------------------------------------------------------------
    col1Arr = lo.ListColumns(1).DataBodyRange.Value2
    rowCount = UBound(col1Arr, 1)
    
    ReDim deleteIdx(1 To rowCount)
    delCount = 0
    keepCount = 0
    
    '--------------------------------------------------------------------------
    ' Stage 1: collect indexes of rows to delete
    ' Do not delete during scan
    '--------------------------------------------------------------------------
    For i = 1 To rowCount
        If Trim$(CStr(col1Arr(i, 1))) = vbNullString Then
            delCount = delCount + 1
            deleteIdx(delCount) = i
        Else
            keepCount = keepCount + 1
        End If
    Next i
    
    '--------------------------------------------------------------------------
    ' Stage 2: delete collected rows in reverse order
    '--------------------------------------------------------------------------
    If delCount > 0 Then
        For i = delCount To 1 Step -1
            lo.ListRows(deleteIdx(i)).Delete
        Next i
    End If
    
    '--------------------------------------------------------------------------
    ' Stage 3: resize to valid rows + one blank spare row
    ' If all rows were blank, result becomes one blank row
    '--------------------------------------------------------------------------
    targetRows = keepCount + 1
    If targetRows < 1 Then targetRows = 1
    
    ResizeListObjectRowsExact lo, targetRows
    
    '--------------------------------------------------------------------------
    ' Ensure final row is the single spare blank row
    '--------------------------------------------------------------------------
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Rows(lo.DataBodyRange.Rows.Count).ClearContents
    End If

SafeExit:
    On Error Resume Next
    
    SheetGuard_End ws, shState
    AppGuard_End
    
    On Error GoTo 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    
    'LogError PROC_NAME, Err.Number, Err.description
    
    SheetGuard_End ws, shState
    AppGuard_End
    
    On Error GoTo 0
End Sub

