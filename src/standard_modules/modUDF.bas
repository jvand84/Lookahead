Attribute VB_Name = "modUDF"
Option Explicit

'====================================================================
' VisibleUniqueList
'
' Purpose:
'   Returns a vertical dynamic array of unique visible values only,
'   excluding blanks.
'
' Usage in worksheet:
'   =VisibleUniqueList(tblLookahead[Role])
'
' Optional:
'   =SORT(VisibleUniqueList(tblLookahead[Role]))
'
' Notes:
'   - Respects filtered/hidden rows
'   - Excludes blanks
'   - Returns a spill range (Excel 365 / dynamic arrays)
'   - If no visible values exist, returns ""
'   - Uses late-bound Scripting.Dictionary (no reference required)
'
' Important:
'   Excel does not always recalculate UDFs when filters change.
'   If needed, press F9 after filtering, or use the optional Trigger arg:
'
'   =VisibleUniqueList(tblLookahead[Role],SUBTOTAL(103,tblLookahead[Role]))
'
'====================================================================

Public Function VisibleUniqueList(ByVal rng As Range, Optional ByVal Trigger As Variant) As Variant
    
    Dim dict As Object
    Dim visRng As Range
    Dim area As Range
    Dim cell As Range
    Dim v As Variant
    Dim arrOut() As Variant
    Dim i As Long
    
    On Error GoTo CleanFail
    
    Application.Volatile
    
    If rng Is Nothing Then
        VisibleUniqueList = vbNullString
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare   ' Case-insensitive
    
    ' Get visible cells only
    On Error Resume Next
    Set visRng = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo CleanFail
    
    If visRng Is Nothing Then
        VisibleUniqueList = vbNullString
        Exit Function
    End If
    
    ' Loop visible cells only
    For Each area In visRng.Areas
        For Each cell In area.Cells
            
            If Not cell.EntireRow.Hidden Then
                
                v = cell.Value2
                
                ' Skip errors
                If Not IsError(v) Then
                    
                    ' Convert once for consistency
                    Dim s As String
                    s = Trim$(CStr(v))
                    
                    ' --- FILTER LOGIC ---
                    ' Exclude:
                    '   - blanks ("")
                    '   - zero values (0 or "0")
                    If Len(s) > 0 Then
                        If s <> "0" Then
                            
                            If Not dict.Exists(s) Then
                                dict.Add s, v
                            End If
                            
                        End If
                    End If
                    
                End If
                
            End If
            
        Next cell
    Next area
    
    If dict.Count = 0 Then
        VisibleUniqueList = vbNullString
        Exit Function
    End If
    
    ReDim arrOut(1 To dict.Count, 1 To 1)
    
    For i = 0 To dict.Count - 1
        arrOut(i + 1, 1) = dict.Items()(i)
    Next i
    
    VisibleUniqueList = arrOut
    Exit Function

CleanFail:
    VisibleUniqueList = vbNullString
End Function
