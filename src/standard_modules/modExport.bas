Attribute VB_Name = "modExport"
Option Explicit

' ===== CONFIG =====
Const EXPORT_PATH As String = "C:\Users\jvand\OneDrive\Walz\998 VBA Projects\Lookahead\src\"
' ==================

' Export all VBA modules, classes, and forms to the EXPORT_PATH folder
Public Sub ExportVBAModules()
    Dim vbComp As Object
    Dim filePath As String
    Dim newExport_Path As String
    
    If Dir(EXPORT_PATH, vbDirectory) = "" Then
        MsgBox "Export folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard module, Class, Form
                newExport_Path = EXPORT_PATH & "standard_modules\"
                filePath = newExport_Path & vbComp.Name & FileExtension(vbComp.Type)
                vbComp.Export filePath
            Case 2
                newExport_Path = EXPORT_PATH & "class_modules\"
                filePath = newExport_Path & vbComp.Name & FileExtension(vbComp.Type)
                vbComp.Export filePath
            Case 3
                newExport_Path = EXPORT_PATH & "userforms\"
                filePath = newExport_Path & vbComp.Name & FileExtension(vbComp.Type)
                vbComp.Export filePath
        End Select
    Next vbComp
    
    MsgBox "Modules exported to: " & EXPORT_PATH, vbInformation
End Sub

' Import all .bas/.cls/.frm files from EXPORT_PATH into the workbook
Public Sub ImportVBAModules()
    Dim fso As Object, folder As Object, file As Object
    
    If Dir(EXPORT_PATH, vbDirectory) = "" Then
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

' Helper function to get correct file extension
Private Function FileExtension(compType As Long) As String
    Select Case compType
        Case 1: FileExtension = ".bas" ' Module
        Case 2: FileExtension = ".cls" ' Class
        Case 3: FileExtension = ".frm" ' Form
        Case Else: FileExtension = ".txt"
    End Select
End Function



