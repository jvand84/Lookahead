Attribute VB_Name = "modExport"
Option Explicit

' ===== CONFIG =====
Const EXPORT_PATH As String = "C:\Users\jvand\OneDrive\Walz\998 VBA Projects\Lookahead\src\"
' ==================

' Export all VBA modules, classes, and forms to the EXPORT_PATH folder
Public Sub ExportVBAModules()
    Dim vbComp As Object
    Dim filePath As String
    
    If Not ExportPathExists() Then
        MsgBox "Export folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        filePath = BuildComponentExportPath(vbComp)
        If Len(filePath) > 0 Then vbComp.Export filePath
    Next vbComp
    
    MsgBox "Modules exported to: " & EXPORT_PATH, vbInformation
End Sub

' Import all .bas/.cls/.frm files from EXPORT_PATH into the workbook
Public Sub ImportVBAModules()
    Dim fso As Object, folder As Object, file As Object
    
    If Not ExportPathExists() Then
        MsgBox "Import folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(EXPORT_PATH)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "bas" _
        Or LCase(fso.GetExtensionName(file.Name)) Like "cls" _
        Or LCase(fso.GetExtensionName(file.Name)) Like "frm" Then
            ThisWorkbook.VBProject.VBComponents.Import file.Path
        End If
    Next file
    
    MsgBox "Modules imported from: " & EXPORT_PATH, vbInformation
End Sub

' Validate export/import base path once.
Private Function ExportPathExists() As Boolean
    ExportPathExists = (Dir(EXPORT_PATH, vbDirectory) <> "")
End Function

' Build full export target path for a component; returns "" for unsupported types.
Private Function BuildComponentExportPath(ByVal vbComp As Object) As String
    Dim exportFolder As String
    
    exportFolder = ComponentFolder(vbComp.Type)
    If Len(exportFolder) = 0 Then Exit Function
    
    BuildComponentExportPath = EXPORT_PATH & exportFolder & "\" & vbComp.Name & FileExtension(vbComp.Type)
End Function

' Map VB component type to folder name.
Private Function ComponentFolder(ByVal compType As Long) As String
    Select Case compType
        Case 1: ComponentFolder = "standard_modules"
        Case 2: ComponentFolder = "class_modules"
        Case 3: ComponentFolder = "userforms"
        Case Else: ComponentFolder = ""
    End Select
End Function

' Helper function to get correct file extension
Private Function FileExtension(compType As Long) As String
    Select Case compType
        Case 1: FileExtension = ".bas" ' Module
        Case 2: FileExtension = ".cls" ' Class
        Case 3: FileExtension = ".frm" ' Form
        Case Else: FileExtension = ".txt"
    End Select
End Function



