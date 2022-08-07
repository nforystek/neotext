VERSION 5.00
Begin VB.Form frmDragOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drag and Drop Options"
   ClientHeight    =   1548
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   6084
   HelpContextID   =   8
   Icon            =   "frmDragOption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1548
   ScaleWidth      =   6084
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1360
      HelpContextID   =   8
      Left            =   96
      TabIndex        =   5
      Top             =   64
      Width           =   3200
      Begin VB.OptionButton Option2 
         Caption         =   "Copy the media being dragged"
         Height          =   400
         HelpContextID   =   8
         Left            =   160
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   2848
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cut the media being dragged"
         Height          =   336
         HelpContextID   =   8
         Left            =   160
         TabIndex        =   1
         Top             =   810
         Width           =   2912
      End
   End
   Begin VB.CheckBox Check1 
      Height          =   720
      HelpContextID   =   8
      Left            =   3488
      TabIndex        =   4
      Top             =   656
      Width           =   2528
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   352
      HelpContextID   =   8
      Index           =   1
      Left            =   3435
      TabIndex        =   2
      Top             =   150
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   352
      HelpContextID   =   8
      Index           =   0
      Left            =   4770
      TabIndex        =   3
      Top             =   144
      Width           =   1200
   End
End
Attribute VB_Name = "frmDragOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean
Public DragOption As Integer

Public Sub PromptDragOption()
    Me.Show
    Do Until Me.Visible = False
        modCommon.DoTasks
    Loop
    If Option1.Value Then DragOption = 2
    If Option2.Value Then DragOption = 1
    If IsOk Then
        If Check1.Value Then
            dbSettings.SetClientSetting "DragOption", DragOption
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 1
            IsOk = True
        Case 0
            IsOk = False
    End Select
    Me.Hide
End Sub

Private Sub Form_Load()
    Check1.Caption = "Always Copy media when drag and dropping."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        IsOk = False
        Me.Hide
    End If
End Sub

Private Sub Option1_Click()
    Check1.Caption = "Always Cut media when drag and dropping."
End Sub

Private Sub Option2_Click()
    Check1.Caption = "Always Copy media when drag and dropping."
End Sub
