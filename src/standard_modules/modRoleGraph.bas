Attribute VB_Name = "modRoleGraph"
Option Explicit

'-------------------------------------------------------
' RebuildRoleGraphFromLookahead
'
' Purpose:
'   Rebuilds tblRoleGraph as a flat table with columns:
'       1. Role
'       2. Local Qty
'       3. Lookahead Qty
'       4. Date
'
' Functional Rules:
'   - Role list is sourced from tblEmployees
'   - Local Qty for each Role/Date is:
'         Base Local Employees for that Role
'         minus Local Leave for that same Role/Date
'   - Lookahead Qty is:
'         total non-blank allocations in tblLookahead
'         for that Role/Date
'   - Date output is written as an actual Excel date
'
' Source Columns:
'   tblEmployees:
'       - Role                = "Role"
'       - Local / Away        = "Local / Away"
'
'   tbl_Vista_HR_Leave:
'       - Role                = "Employees.Role"
'       - Local / Away        = "tblLocal.Local / Away"
'       - Date                = "Date"
'
' Notes:
'   - Uses deterministic dictionary keys:
'         UCase(Trim(Role)) & "|" & DateSerial
'   - This avoids text/date mismatch issues that stop
'     leave subtraction from applying.
'   - Assumes helper functions exist in another module:
'       * FindListObjectByName
'       * HeaderMapCached
'       * NzText
'       * HeaderToDate
'       * IsValueBlank
'       * LogError
'       * SheetGuard_Begin / SheetGuard_End
'       * AppGuard_Begin / AppGuard_End
'-------------------------------------------------------
Public Sub RebuildRoleGraphFromLookahead()

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim loEmployees As ListObject
    Dim loLookahead As ListObject
    Dim loRoleGraph As ListObject
    Dim loLeave As ListObject

    Dim wsEmployees As Worksheet
    Dim wsLookahead As Worksheet
    Dim wsRoleGraph As Worksheet
    Dim wsLeave As Worksheet

    Dim sgEmployees As TSheetGuardState
    Dim sgLookahead As TSheetGuardState
    Dim sgRoleGraph As TSheetGuardState
    Dim sgLeave As TSheetGuardState

    Dim appGuardStarted As Boolean

    Dim empHdrMap As Object
    Dim leaveHdrMap As Object

    Dim dictRoleOrder As Object
    Dim dictLocalEmpCount As Object
    Dim dictLookaheadCounts As Object
    Dim dictLeaveCounts As Object

    Dim arrEmp As Variant
    Dim arrLook As Variant
    Dim arrLeave As Variant
    Dim arrHeadersLook As Variant
    Dim arrOut() As Variant

    Dim roles() As String
    Dim dates() As Date

    Dim idxEmpRole As Long
    Dim idxEmpLocalAway As Long

    Dim idxLeaveRole As Long
    Dim idxLeaveLocalAway As Long
    Dim idxLeaveDate As Long

    Dim totalLookCols As Long
    Dim totalDateCols As Long
    Dim roleCount As Long
    Dim outRowCount As Long

    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim roleIdx As Long
    Dim dateIdx As Long
    Dim dateOffset As Long

    Dim roleKey As String
    Dim localAway As String
    Dim compositeKey As String
    Dim localQty As Long
    Dim leaveSerial As Long
    Dim vKey As Variant

    Const PROC_NAME As String = "RebuildRoleGraphFromLookahead"

    '-------------------------------------------------------
    ' Guard Initialisation
    '-------------------------------------------------------
    AppGuard_Begin
    appGuardStarted = True

    '-------------------------------------------------------
    ' Object / Table Resolution
    '-------------------------------------------------------
    Set loEmployees = FindListObjectByName("tblEmployees")
    Set loLookahead = FindListObjectByName("tblLookahead")
    Set loRoleGraph = FindListObjectByName("tblRoleGraph")
    Set loLeave = FindListObjectByName("tbl_Vista_HR_Leave")

    If loEmployees Is Nothing Then Err.Raise vbObjectError + 1000, PROC_NAME, "Table 'tblEmployees' was not found."
    If loLookahead Is Nothing Then Err.Raise vbObjectError + 1001, PROC_NAME, "Table 'tblLookahead' was not found."
    If loRoleGraph Is Nothing Then Err.Raise vbObjectError + 1002, PROC_NAME, "Table 'tblRoleGraph' was not found."
    If loLeave Is Nothing Then Err.Raise vbObjectError + 1003, PROC_NAME, "Table 'tbl_Vista_HR_Leave' was not found."

    Set wsEmployees = loEmployees.Parent
    Set wsLookahead = loLookahead.Parent
    Set wsRoleGraph = loRoleGraph.Parent
    Set wsLeave = loLeave.Parent

    sgEmployees = SheetGuard_Begin(wsEmployees)
    sgLookahead = SheetGuard_Begin(wsLookahead)
    sgRoleGraph = SheetGuard_Begin(wsRoleGraph)
    sgLeave = SheetGuard_Begin(wsLeave)

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    Set empHdrMap = HeaderMapCached(loEmployees)
    Set leaveHdrMap = HeaderMapCached(loLeave)

    idxEmpRole = GetHeaderIndexSafe(empHdrMap, "Role")
    If idxEmpRole = 0 Then
        Err.Raise vbObjectError + 1004, PROC_NAME, "Header 'Role' was not found in tblEmployees."
    End If

    idxEmpLocalAway = GetHeaderIndexSafe(empHdrMap, "Local / Away")
    If idxEmpLocalAway = 0 Then
        Err.Raise vbObjectError + 1005, PROC_NAME, "Header 'Local / Away' was not found in tblEmployees."
    End If

    idxLeaveRole = GetHeaderIndexSafe(leaveHdrMap, "Employees.Role")
    If idxLeaveRole = 0 Then
        Err.Raise vbObjectError + 1006, PROC_NAME, "Header 'Employees.Role' was not found in tbl_Vista_HR_Leave."
    End If

    idxLeaveLocalAway = GetHeaderIndexSafe(leaveHdrMap, "tblLocal.Local / Away")
    If idxLeaveLocalAway = 0 Then
        Err.Raise vbObjectError + 1007, PROC_NAME, "Header 'tblLocal.Local / Away' was not found in tbl_Vista_HR_Leave."
    End If

    idxLeaveDate = GetHeaderIndexSafe(leaveHdrMap, "Date")
    If idxLeaveDate = 0 Then
        Err.Raise vbObjectError + 1008, PROC_NAME, "Header 'Date' was not found in tbl_Vista_HR_Leave."
    End If

    totalLookCols = loLookahead.ListColumns.Count

    If vManRoleCol <= 0 Or vManRoleCol > totalLookCols Then
        Err.Raise vbObjectError + 1009, PROC_NAME, "vManRoleCol is outside the bounds of tblLookahead."
    End If

    If vManCalCol <= 0 Or vManCalCol > totalLookCols Then
        Err.Raise vbObjectError + 1010, PROC_NAME, "vManCalCol is outside the bounds of tblLookahead."
    End If

    totalDateCols = totalLookCols - vManCalCol + 1
    arrHeadersLook = loLookahead.HeaderRowRange.Value2

    '-------------------------------------------------------
    ' Initialise dictionaries
    '-------------------------------------------------------
    Set dictRoleOrder = CreateObject("Scripting.Dictionary")
    Set dictLocalEmpCount = CreateObject("Scripting.Dictionary")
    Set dictLookaheadCounts = CreateObject("Scripting.Dictionary")
    Set dictLeaveCounts = CreateObject("Scripting.Dictionary")

    dictRoleOrder.CompareMode = vbTextCompare
    dictLocalEmpCount.CompareMode = vbTextCompare
    dictLookaheadCounts.CompareMode = vbTextCompare
    dictLeaveCounts.CompareMode = vbTextCompare

    '-------------------------------------------------------
    ' Build from tblEmployees:
    '   1. Unique role list
    '   2. Base Local employee count per role
    '-------------------------------------------------------
    If Not loEmployees.DataBodyRange Is Nothing Then
        arrEmp = loEmployees.DataBodyRange.Value2

        For r = 1 To UBound(arrEmp, 1)
            roleKey = RoleKeyText(arrEmp(r, idxEmpRole))
            localAway = UCase$(Trim$(NzText(arrEmp(r, idxEmpLocalAway))))

            If Len(roleKey) > 0 Then
                If Not dictRoleOrder.Exists(roleKey) Then
                    dictRoleOrder.Add roleKey, dictRoleOrder.Count + 1
                End If

                If localAway = "LOCAL" Then
                    If dictLocalEmpCount.Exists(roleKey) Then
                        dictLocalEmpCount(roleKey) = CLng(dictLocalEmpCount(roleKey)) + 1
                    Else
                        dictLocalEmpCount.Add roleKey, 1
                    End If
                End If
            End If
        Next r
    End If

    roleCount = dictRoleOrder.Count

    '-------------------------------------------------------
    ' Build ordered role array
    '-------------------------------------------------------
    If roleCount > 0 Then
        ReDim roles(1 To roleCount)

        For Each vKey In dictRoleOrder.Keys
            roles(CLng(dictRoleOrder(vKey))) = CStr(vKey)
        Next vKey
    End If

    '-------------------------------------------------------
    ' Build date array from tblLookahead headers as real dates
    '-------------------------------------------------------
    If totalDateCols > 0 Then
        ReDim dates(1 To totalDateCols)

        For c = vManCalCol To totalLookCols
            dates(c - vManCalCol + 1) = HeaderToDate(arrHeadersLook(1, c))
        Next c
    End If

    '-------------------------------------------------------
    ' Count lookahead allocations from tblLookahead
    ' Group key = UCase(Role)|DateSerial
    '-------------------------------------------------------
    If Not loLookahead.DataBodyRange Is Nothing Then
        arrLook = loLookahead.DataBodyRange.Value2

        For r = 1 To UBound(arrLook, 1)
            roleKey = RoleKeyText(arrLook(r, vManRoleCol))

            If Len(roleKey) > 0 Then
                For c = vManCalCol To totalLookCols
                    If Not IsValueBlank(arrLook(r, c)) Then
                        dateOffset = c - vManCalCol + 1
                        compositeKey = roleKey & "|" & CStr(CLng(dates(dateOffset)))

                        If dictLookaheadCounts.Exists(compositeKey) Then
                            dictLookaheadCounts(compositeKey) = CLng(dictLookaheadCounts(compositeKey)) + 1
                        Else
                            dictLookaheadCounts.Add compositeKey, 1
                        End If
                    End If
                Next c
            End If
        Next r
    End If

    '-------------------------------------------------------
    ' Count local leave by Role + Date directly from leave table
    ' Group key = UCase(Role)|DateSerial
    '-------------------------------------------------------
    If Not loLeave.DataBodyRange Is Nothing Then
        arrLeave = loLeave.DataBodyRange.Value2

        For r = 1 To UBound(arrLeave, 1)
            roleKey = RoleKeyText(arrLeave(r, idxLeaveRole))
            localAway = UCase$(Trim$(NzText(arrLeave(r, idxLeaveLocalAway))))

            If Len(roleKey) > 0 Then
                If localAway = "LOCAL" Then
                    leaveSerial = ValueToDateSerial(arrLeave(r, idxLeaveDate))
                    compositeKey = roleKey & "|" & CStr(leaveSerial)

                    If dictLeaveCounts.Exists(compositeKey) Then
                        dictLeaveCounts(compositeKey) = CLng(dictLeaveCounts(compositeKey)) + 1
                    Else
                        dictLeaveCounts.Add compositeKey, 1
                    End If
                End If
            End If
        Next r
    End If

    '-------------------------------------------------------
    ' Build output array:
    '   Role | Local Qty | Lookahead Qty | Date
    ' One row for every Role x Date
    '-------------------------------------------------------
    outRowCount = roleCount * totalDateCols

    If outRowCount > 0 Then
        ReDim arrOut(1 To outRowCount, 1 To 4)

        outRow = 0

        For roleIdx = 1 To roleCount
            roleKey = roles(roleIdx)

            For dateIdx = 1 To totalDateCols
                outRow = outRow + 1
                compositeKey = roleKey & "|" & CStr(CLng(dates(dateIdx)))

                arrOut(outRow, 1) = roleKey

                If dictLocalEmpCount.Exists(roleKey) Then
                    localQty = CLng(dictLocalEmpCount(roleKey))
                Else
                    localQty = 0
                End If

                If dictLeaveCounts.Exists(compositeKey) Then
                    localQty = localQty - CLng(dictLeaveCounts(compositeKey))
                    If localQty < 0 Then localQty = 0
                End If

                arrOut(outRow, 2) = localQty

                If dictLookaheadCounts.Exists(compositeKey) Then
                    arrOut(outRow, 3) = CLng(dictLookaheadCounts(compositeKey))
                Else
                    arrOut(outRow, 3) = 0
                End If

                arrOut(outRow, 4) = dates(dateIdx)
            Next dateIdx
        Next roleIdx
    End If

    '-------------------------------------------------------
    ' Resize tblRoleGraph to exact dimensions:
    '   4 columns
    '   Role x Date rows
    '-------------------------------------------------------
    ResizeListObjectToExactDimensions loRoleGraph, outRowCount, 4

    '-------------------------------------------------------
    ' Write fixed headers
    '-------------------------------------------------------
    loRoleGraph.HeaderRowRange.Cells(1, 1).Value = "Role"
    loRoleGraph.HeaderRowRange.Cells(1, 2).Value = "Local Qty"
    loRoleGraph.HeaderRowRange.Cells(1, 3).Value = "Lookahead Qty"
    loRoleGraph.HeaderRowRange.Cells(1, 4).Value = "Date"

    '-------------------------------------------------------
    ' Write output rows
    '-------------------------------------------------------
    If outRowCount > 0 Then
        loRoleGraph.DataBodyRange.Value2 = arrOut
        loRoleGraph.ListColumns(4).DataBodyRange.NumberFormat = "dd/mm/yy"
    Else
        If Not loRoleGraph.DataBodyRange Is Nothing Then
            loRoleGraph.DataBodyRange.Delete
        End If
    End If

CleanExit:
    '-------------------------------------------------------
    ' Cleanup
    '-------------------------------------------------------
    On Error Resume Next

    SheetGuard_End wsRoleGraph, sgRoleGraph
    SheetGuard_End wsLookahead, sgLookahead
    SheetGuard_End wsEmployees, sgEmployees
    SheetGuard_End wsLeave, sgLeave

    If appGuardStarted Then
        AppGuard_End
    End If

    On Error GoTo 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    LogError PROC_NAME, Err.Number, Err.description
    On Error GoTo 0
    Resume CleanExit

End Sub

'-------------------------------------------------------
' ResizeListObjectToExactDimensions
'
' Purpose:
'   Resizes a ListObject to the exact required number of
'   data rows and columns.
'-------------------------------------------------------
Private Sub ResizeListObjectToExactDimensions(ByVal lo As ListObject, _
                                              ByVal dataRowCount As Long, _
                                              ByVal colCount As Long)

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim topLeft As Range
    Dim newRange As Range
    Dim totalRows As Long

    Const PROC_NAME As String = "ResizeListObjectToExactDimensions"

    If lo Is Nothing Then Err.Raise vbObjectError + 1100, PROC_NAME, "Target ListObject is Nothing."
    If colCount < 1 Then Err.Raise vbObjectError + 1101, PROC_NAME, "colCount must be at least 1."
    If dataRowCount < 0 Then Err.Raise vbObjectError + 1102, PROC_NAME, "dataRowCount cannot be negative."

    Set ws = lo.Parent
    Set topLeft = lo.Range.Cells(1, 1)

    totalRows = 1 + dataRowCount
    Set newRange = ws.Range(topLeft, topLeft.Offset(totalRows - 1, colCount - 1))

    lo.Resize newRange
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, PROC_NAME, Err.description
End Sub

'-------------------------------------------------------
' GetHeaderIndexSafe
'
' Purpose:
'   Returns a header index from a header map dictionary,
'   or 0 if not found.
'-------------------------------------------------------
Private Function GetHeaderIndexSafe(ByVal hdrMap As Object, ByVal headerName As String) As Long

    If hdrMap Is Nothing Then Exit Function
    If Len(headerName) = 0 Then Exit Function

    If hdrMap.Exists(headerName) Then
        GetHeaderIndexSafe = CLng(hdrMap(headerName))
    Else
        GetHeaderIndexSafe = 0
    End If

End Function

