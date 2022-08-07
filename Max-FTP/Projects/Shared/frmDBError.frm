VERSION 5.00
Begin VB.Form frmDBError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Accessing"
   ClientHeight    =   2736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6780
   Icon            =   "frmDBError.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2736
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Retry"
      Height          =   420
      Index           =   2
      Left            =   3915
      TabIndex        =   0
      Top             =   2175
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   420
      Index           =   0
      Left            =   5370
      TabIndex        =   1
      Top             =   2175
      Width           =   1305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   165
      X2              =   6660
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   165
      X2              =   6660
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   165
      X2              =   6660
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   165
      X2              =   6660
      Y1              =   165
      Y2              =   165
   End
   Begin VB.Label Label2 
      Height          =   885
      Left            =   240
      TabIndex        =   3
      Top             =   825
      Width           =   6165
   End
   Begin VB.Label Label1 
      Caption         =   "There was a problem trying to access the settings database."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   405
      Width           =   4380
   End
End
Attribute VB_Name = "frmDBError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Integer

Public Sub ShowError(ByVal ErrorValue As String)

    Label1.Caption = "There was a problem trying to access the settings database."
    Label2.Caption = ErrorValue
    
    Command1(2).Visible = True
    Command1(0).Caption = "&Close"
    
    Timer1.Tag = CInt(9)
    Command1(2).Caption = "Retry (" & Trim(Timer1.Tag) & ")"
    Timer1.enabled = True
    
    Me.Show
End Sub

Private Sub Command1_Click(Index As Integer)
    IsOk = Index
    Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        IsOk = 0
        Me.Hide
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Tag = CInt(Timer1.Tag) - 1
    If Timer1.Tag = 0 Then
        Timer1.enabled = False
        Command1_Click 2
    Else
        Command1(2).Caption = "Retry (" & Trim(Timer1.Tag) & ")"
    End If
End Sub


