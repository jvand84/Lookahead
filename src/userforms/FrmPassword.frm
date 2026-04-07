VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPassword 
   Caption         =   "Password"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmPassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    pssword = ""
    Unload FrmPassword
End Sub

Private Sub cmdok_Click()
    pssword = txtBoxPW.Value
    Unload FrmPassword
End Sub

Private Sub UserForm_Initialize()
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
End Sub
