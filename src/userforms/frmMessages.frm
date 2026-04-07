VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMessages 
   Caption         =   "Important Message"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMessages.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public closeTime As Double
Public strCaption As String
Public strMess As String
Public durationInSeconds As Double
Private pStartTime As Single
Private pDurationSeconds As Double

Public Sub InitializeMessage(strCaption As String, strMess As String, secondsToShow As Double)
    Const maxWidth As Double = 300
    Const padding As Double = 40

    Me.Caption = strCaption

    ' Setup the measuring label
    With Me.lblMeasure
        .Width = maxWidth
        .Caption = strMess
    End With

    ' Setup the main label
    With Me.lbl1
        .WordWrap = True
        .AutoSize = False
        .Width = maxWidth
        .Caption = strMess
        .Height = Me.lblMeasure.Height
    End With

    ' Resize form to fit message
    Me.Height = Me.lbl1.Top + Me.lbl1.Height + padding
    Me.Width = Me.lbl1.Left + Me.lbl1.Width + padding

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

    ' Show the form
    pStartTime = Timer
    pDurationSeconds = secondsToShow

    Me.Show vbModeless

    If secondsToShow > 0 Then
        Application.OnTime Now + (0.5 / 86400), "CloseTemporaryMessage"
    End If
End Sub

Public Property Get TimeElapsed() As Single
    TimeElapsed = Timer - pStartTime
End Property

Public Property Get DurationSeconds() As Single
    DurationSeconds = pDurationSeconds
End Property

