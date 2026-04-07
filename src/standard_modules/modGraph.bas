Attribute VB_Name = "modGraph"
Option Explicit

Sub SetHistoLocalSlicer()

    Dim sc As SlicerCache
    Dim si As SlicerItem
    Dim keepItems As Object
    
    On Error GoTo ErrHandler
    
    '---------------------------------------------
    ' Get slicer cache by name
    '---------------------------------------------
    Set sc = ThisWorkbook.SlicerCaches("Slicer_Local___Away2")
    
    '---------------------------------------------
    ' Items to keep selected
    '---------------------------------------------
    Set keepItems = CreateObject("Scripting.Dictionary")
    keepItems.CompareMode = vbTextCompare
    keepItems("Local") = True
    keepItems("Local App") = True
    keepItems("Role") = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '---------------------------------------------
    ' Clear any existing filter first
    '---------------------------------------------
    sc.ClearManualFilter
    
    '---------------------------------------------
    ' Select only required items
    '---------------------------------------------
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
    MsgBox "Error setting slicer filter: " & Err.description, vbExclamation
End Sub

Sub SetHistoRoleSlicer()

    Dim sc As SlicerCache
    Dim si As SlicerItem
    Dim keepItems As Object
    
    On Error GoTo ErrHandler
    
    '---------------------------------------------
    ' Get slicer cache by name
    '---------------------------------------------
    Set sc = ThisWorkbook.SlicerCaches("Slicer_Role2")
    
    '---------------------------------------------
    ' Items to keep selected
    '---------------------------------------------
    Set keepItems = CreateObject("Scripting.Dictionary")
    keepItems.CompareMode = vbTextCompare
    keepItems("Crane Operator") = True
    keepItems("Dogman") = True
    keepItems("Rigger") = True
    keepItems("Slew Crane Operator") = True
    keepItems("Franna Operator") = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '---------------------------------------------
    ' Clear any existing filter first
    '---------------------------------------------
    sc.ClearManualFilter
    
    '---------------------------------------------
    ' Select only required items
    '---------------------------------------------
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
    MsgBox "Error setting slicer filter: " & Err.description, vbExclamation
End Sub


Sub ListSlicerCaches()
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        Debug.Print sc.Name
    Next sc
End Sub

