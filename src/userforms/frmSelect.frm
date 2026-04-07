VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "Select Workbook"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelect.frx":0000
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public vSelection As String
Public FrmType As Integer

Public Property Get SelectedWorkbookName() As String
    SelectedWorkbookName = vSelection
End Property


Private Sub UserForm_Initialize()
    Dim xlLeft As Long, xlTop As Long
    Dim xlWidth As Long, xlHeight As Long
    Dim frmWidth As Long, frmHeight As Long

    ' Get Excel window position and size
    With Application
        xlLeft = .Left
        xlTop = .Top
        xlWidth = .Width
        xlHeight = .Height
    End With

    ' Get form size in points (approximate)
    frmWidth = Me.Width
    frmHeight = Me.Height

    ' Ensure the form is on the current screen
    Dim x As Long, y As Long
    Dim scrW As Long, scrH As Long
    x = Application.Left + (Application.Width - Me.Width) / 2
    y = Application.Top + (Application.Height - Me.Height) / 2

    ' Get screen dimensions
    scrW = Application.UsableWidth
    scrH = Application.UsableHeight

    ' Keep the form on screen
    If x + Me.Width > Application.Left + scrW Then x = Application.Left + scrW - Me.Width
    If y + Me.Height > Application.Top + scrH Then y = Application.Top + scrH - Me.Height
    If x < Application.Left Then x = Application.Left
    If y < Application.Top Then y = Application.Top

    Me.StartUpPosition = 0
    Me.Left = x
    Me.Top = y
    
    If cmbSelection.ListCount > 0 Then
        cmbSelection.ListIndex = 0 ' Select first item by default
    End If
    
End Sub

Sub LoadCombo()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TBL As ListObject
    Dim i As Long
    Dim val1 As String, val2 As String
    
    Select Case FrmType
    
    Case 1 ' Populate ComboBox with open workbooks (excluding ThisWorkbook)
        For Each wb In Application.Workbooks
            If wb.Name <> ThisWorkbook.Name Then
                cmbSelection.AddItem wb.Name
            End If
        Next wb
        lbl1.Caption = "Select Workbook to Get Data from."
        frmSelect.Caption = "Select Workbook"
    Case 2 ' Populate ComboBox with PCC Claims
        
        Set ws = ThisWorkbook.Sheets("PC Register")
        Set TBL = ws.ListObjects("tblPCC")
        ' Clear and fill combo box
        cmbSelection.Clear
        For i = 1 To TBL.ListRows.Count
            cmbSelection.AddItem TBL.DataBodyRange.Cells(i, 1).Value
        Next i
        lbl1.Caption = "Select Progress Claim to populate."
        frmSelect.Caption = "Select Progress Claim"
    Case 3
        Set ws = ThisWorkbook.Sheets("Table Names Summary")
        Set TBL = ws.ListObjects("tblTables")
        With cmbSelection
            .Clear
            
            For i = 1 To TBL.ListRows.Count
                val1 = TBL.DataBodyRange.Cells(i, 1).Value ' Column 1
                If TBL.DataBodyRange.Cells(i, 6) = True Then
                    If Not ItemExistsInCombo(cmbSelection, val1) Then
                        .AddItem
                        .List(.ListCount - 1, 0) = val1
                    End If
                End If
            Next i
        End With
        lbl1.Caption = "Select Sheet."
        frmSelect.Caption = "Select Sheet to Navigate to."
    End Select
    
    If cmbSelection.ListCount > 0 Then
        cmbSelection.ListIndex = 0 ' Select first item by default
    End If
End Sub

Private Sub btnOK_Click()
    If cmbSelection.ListIndex = -1 Then
        MsgBox "Please select.", vbExclamation
        Exit Sub
    End If
    vSelection = cmbSelection.Value
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    vSelection = ""
    Me.Hide
End Sub
