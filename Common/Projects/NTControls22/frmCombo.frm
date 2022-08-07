VERSION 5.00
Begin VB.Form frmCombo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmCombo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstMatch 
      CausesValidation=   0   'False
      Height          =   1035
      Left            =   -30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   -15
      Width           =   2865
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public HookHWnd As Long

Public srhText

Public Sub ResizeBox()
    On Error Resume Next
    lstMatch.Top = -30
    lstMatch.Left = -30
    lstMatch.Width = Me.ScaleWidth + (60)
    Err.Clear
End Sub

Private Sub SetText(ByVal KeyAscii As Integer)
    If lstMatch.ListIndex >= 0 Then
        SetFocusAPI srhText.hWnd
        Me.Visible = False
        srhText.Text = lstMatch.List(lstMatch.ListIndex)
        srhText.SelStart = Len(srhText.Text)
        If KeyAscii > 0 Then
            SendMessageLong srhText.hWnd, &H102, KeyAscii, 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook
End Sub

Private Sub lstMatch_DblClick()
    SetText 0
End Sub

Private Sub lstMatch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If frmCombo.lstMatch.ListIndex = 0 Then
                SetFocusAPI srhText.hWnd
                srhText.SelStart = Len(srhText.Text)
                lstMatch.ListIndex = -1
            End If
        Case 13
            SetText 0
            HideWindow
    End Select
End Sub

Private Sub lstMatch_KeyPress(KeyAscii As Integer)
    SetText KeyAscii
End Sub

Private Sub lstMatch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetText 0
        HideWindow
    End If
End Sub

Attribute 