VERSION 5.00
Begin VB.Form frmNetConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Connection"
   ClientHeight    =   996
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6564
   HelpContextID   =   14
   Icon            =   "frmNetConnection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   996
   ScaleWidth      =   6564
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      HelpContextID   =   14
      Index           =   1
      Left            =   5580
      TabIndex        =   3
      Top             =   540
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      HelpContextID   =   14
      Index           =   0
      Left            =   5580
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Height          =   900
      HelpContextID   =   14
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   5355
      Begin VB.ComboBox Combo1 
         Height          =   315
         HelpContextID   =   14
         Left            =   3495
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1740
      End
      Begin VB.TextBox Text1 
         Height          =   315
         HelpContextID   =   14
         Left            =   135
         TabIndex        =   0
         Top             =   420
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Drive Letter:"
         Height          =   225
         Left            =   3495
         TabIndex        =   6
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Share Name:"
         Height          =   225
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmNetConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean
Public ShareName As String
Public DriveLetter As String

Private Sub Command1_Click(Index As Integer)

Select Case Index
    Case 0

            If Trim(Len(Combo1.Text)) <> 2 Then
                MsgBox "You must choose a drive letter that isn't already used by a network connection.", 64, AppName
            Else
                ShareName = Text1.Text
                DriveLetter = Left(Combo1.Text, 1)
                IsOk = True
                Me.Visible = False
                
            End If

    Case 1
        IsOk = False
        Me.Visible = False
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then
    Cancel = True
    IsOk = False
    Me.Visible = False
    End If

End Sub
