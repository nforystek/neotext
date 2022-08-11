VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goto Line Number"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   1980
      TabIndex        =   3
      Top             =   735
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   180
      TabIndex        =   2
      Top             =   735
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   240
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Line Number:"
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   285
      Width           =   1080
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Command1_Click()
    If IsNumeric(Text1.Text) Then
        fMain.GotoLine Text1.Text
        Unload Me
    Else
        MsgBox "Invalid line number.", vbInformation, App.Title
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fMain.StayOnTop Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fMain.NotStayOnTop Me
End Sub
