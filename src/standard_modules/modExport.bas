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

Public Sub ImportVBAModules_Safe()

    Dim fso As Object, folder As Object
    Dim vbComps As Object
    Dim okCount As Long, failCount As Long, skipCount As Long
    Dim logText As String, logPath As String
    Dim ts As Object

    Const CT_STD As Long = 1
    Const CT_CLASS As Long = 2
    Const CT_MSFORM As Long = 3
    Const CT_DOC As Long = 100

    On Error GoTo HardFail

    '--- sanity: export folder present?
    If Len(EXPORT_PATH) = 0 Or Dir(EXPORT_PATH, vbDirectory) = "" Then
        MsgBox "Import folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(EXPORT_PATH)

    '--- sanity: VBProject access
    On Error Resume Next
    Set vbComps = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Then
        Dim errNum As Long, eMsg As String
        errNum = Err.Number: eMsg = Err.description
        On Error GoTo 0
        MsgBox "Cannot access VBA project (" & errNum & "): " & eMsg & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings... > " & _
               "Macro Settings > check 'Trust access to the VBA project object model'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    logText = "Import started: " & Now & vbCrLf & _
              "Source: " & EXPORT_PATH & vbCrLf & String(60, "-") & vbCrLf

    '--- RECURSIVE CALL
    Call ProcessFolderRecursive(folder, fso, vbComps, _
                                okCount, failCount, skipCount, logText)

    '--- write log
    logPath = fso.BuildPath(EXPORT_PATH, "ImportLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")

    On Error Resume Next
    Set ts = fso.OpenTextFile(logPath, 2, True)
    ts.Write logText
    ts.Close
    On Error GoTo 0

    '--- summary
    MsgBox "VBA import finished." & vbCrLf & vbCrLf & _
           "Imported: " & okCount & vbCrLf & _
           "Failed:   " & failCount & vbCrLf & _
           "Skipped:  " & skipCount & " (document modules not removed)" & vbCrLf & vbCrLf & _
           "Log: " & logPath, _
           IIf(failCount > 0, vbExclamation, vbInformation)

    Exit Sub

HardFail:
    MsgBox "Unexpected error (" & Err.Number & "): " & Err.description, vbCritical

End Sub


Private Sub ProcessFolderRecursive( _
    ByVal folder As Object, _
    ByVal fso As Object, _
    ByVal vbComps As Object, _
    ByRef okCount As Long, _
    ByRef failCount As Long, _
    ByRef skipCount As Long, _
    ByRef logText As String)

    Dim file As Object, subFolder As Object
    Dim vbComp As Object
    Dim compName As String, ext As String

    Const CT_DOC As Long = 100

    '--- process files in this folder
    For Each file In folder.Files
        ext = LCase$(fso.GetExtensionName(file.Name))

        If ext = "bas" Or ext = "cls" Or ext = "frm" Then

            compName = fso.GetBaseName(file.Name)

            ' Find existing
            Set vbComp = Nothing
            On Error Resume Next
            Set vbComp = vbComps(compName)
            On Error GoTo 0

            ' Remove if exists (except document modules)
            If Not vbComp Is Nothing Then
                If vbComp.Type = CT_DOC Then
                    skipCount = skipCount + 1
                    logText = logText & "SKIP remove (Document): " & compName & vbCrLf
                Else
                    On Error Resume Next
                    vbComps.Remove vbComp
                    If Err.Number <> 0 Then
                        failCount = failCount + 1
                        logText = logText & "FAIL remove " & compName & _
                                  " -> " & Err.Number & ": " & Err.description & vbCrLf
                        Err.Clear
                    Else
                        logText = logText & "OK   removed: " & compName & vbCrLf
                    End If
                    On Error GoTo 0
                End If
            End If

            ' Import
            On Error Resume Next
            vbComps.Import file.Path
            If Err.Number <> 0 Then
                failCount = failCount + 1
                logText = logText & "FAIL import " & file.Path & _
                          " -> " & Err.Number & ": " & Err.description & vbCrLf
                Err.Clear
            Else
                okCount = okCount + 1
                logText = logText & "OK   import " & file.Path & vbCrLf
            End If
            On Error GoTo 0
        End If
    Next file

    '--- recurse into subfolders
    For Each subFolder In folder.SubFolders
        ProcessFolderRecursive subFolder, fso, vbComps, _
                               okCount, failCount, skipCount, logText
    Next subFolder

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



