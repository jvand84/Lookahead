Attribute VB_Name = "modGraph"
Option Explicit

Sub SetHistoLocalSlicer()
    ApplySlicerFilter "Slicer_Local___Away2", Array("Local", "Local App", "Role")
End Sub

Sub SetHistoRoleSlicer()
    ApplySlicerFilter "Slicer_Role2", Array("Crane Operator", "Dogman", "Rigger", "Slew Crane Operator", "Franna Operator")
End Sub


Sub ListSlicerCaches()
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        Debug.Print sc.Name
    Next sc
End Sub

Private Sub ApplySlicerFilter(ByVal slicerCacheName As String, ByVal selectedItems As Variant)
    Dim sc As SlicerCache
    Dim si As SlicerItem
    Dim keepItems As Object
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    Set sc = ThisWorkbook.SlicerCaches(slicerCacheName)
    Set keepItems = CreateObject("Scripting.Dictionary")
    keepItems.CompareMode = vbTextCompare
    
    For i = LBound(selectedItems) To UBound(selectedItems)
        keepItems(CStr(selectedItems(i))) = True
    Next i
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    sc.ClearManualFilter
    For Each si In sc.SlicerItems
        si.Selected = keepItems.Exists(si.Name)
    Next si
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error setting slicer filter: " & Err.Description, vbExclamation
End Sub

