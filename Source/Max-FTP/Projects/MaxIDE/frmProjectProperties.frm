VERSION 5.00
Begin VB.Form frmProjectProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Properties"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmProjectProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   90
      TabIndex        =   2
      Top             =   -15
      Width           =   3390
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1155
         TabIndex        =   4
         Top             =   645
         Width           =   2115
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   90
         X2              =   3285
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   90
         X2              =   3285
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Project Name:"
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Visual Basic Project"
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   180
         Width           =   3165
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1650
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   1515
      TabIndex        =   0
      Top             =   1650
      Width           =   930
   End
End
Attribute VB_Name = "frmProjectProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        
        Case 0 'ok
            If ValidNameSpace(Text1.Text) Then
            
                'frmMainIDE.SetProjectName Text1.Text
                
                Unload Me
            End If
        Case 1 'cancel
            Unload Me
    End Select
End Sub

Private Sub Form_Load()

    Label3.Caption = "Language: Visual Basic Script"

    
    'Text1.Text = Project.ProjectName
End Sub

