


VERSION 5.00
Begin VB.Form frmWildCards 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Wildcard Selection"
   ClientHeight    =   2460
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   4332
   HelpContextID   =   8
   Icon            =   "frmWildCards.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "wildcards"
   Begin VB.Frame Frame1 
      Height          =   1905
      HelpContextID   =   8
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   4110
      Begin VB.OptionButton Option1 
         Caption         =   "Files and Folders"
         Height          =   240
         HelpContextID   =   8
         Index           =   2
         Left            =   2400
         TabIndex        =   2
         Top             =   540
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Folders"
         Height          =   240
         HelpContextID   =   8
         Index           =   1
         Left            =   1215
         TabIndex        =   1
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Files"
         Height          =   240
         HelpContextID   =   8
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   540
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox txtWildCard 
         Height          =   304
         HelpContextID   =   8
         Left            =   195
         TabIndex        =   3
         Top             =   1365
         Width           =   3705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   135
         X2              =   3975
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   135
         X2              =   3975
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Select:"
         Height          =   225
         Left            =   195
         TabIndex        =   8
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Wildcard Pattern."
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1050
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      HelpContextID   =   8
      Index           =   1
      Left            =   3315
      TabIndex        =   5
      Top             =   2025
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      HelpContextID   =   8
      Index           =   0
      Left            =   2250
      TabIndex        =   4
      Top             =   2025
      Width           =   915
   End
End
Attribute VB_Name = "frmWildCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public WildCard As String
Public WildOption As Integer
Public IsOk As Boolean

Private Sub Command1_Click(Index As Integer)

    Select Case Index
        Case 0
            IsOk = True
            WildCard = txtWildCard.Text
            If Option1(0).Value Then
                WildOption = 0
            ElseIf Option1(1).Value Then
                WildOption = 1
            ElseIf Option1(2).Value Then
                WildOption = 2
            End If
                
        Case 1
            IsOk = False
    End Select
    Me.Hide
    
End Sub

Private Sub Form_Load()
    SetIcecue Line1, "icecue_shadow"
    SetIcecue Line2, "icecue_hilite"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        IsOk = False
        Me.Hide
        Cancel = True
    End If

End Sub