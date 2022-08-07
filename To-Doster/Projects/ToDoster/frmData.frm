
VERSION 5.00
Begin VB.Form frmData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Database"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Setup"
      Height          =   315
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   2265
      MaxLength       =   65534
      TabIndex        =   3
      Top             =   105
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   315
      Index           =   1
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   315
      Index           =   0
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   555
      Width           =   870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
      Height          =   210
      Left            =   1485
      TabIndex        =   2
      Top             =   135
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   -15
      Picture         =   "frmData.frx":2CFA
      Top             =   -120
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmData.frx":372B
      Height          =   1050
      Left            =   1485
      TabIndex        =   4
      Top             =   510
      Width           =   4065
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0 'ok
            Settings.Location = Text2.Text
            SaveSettings
            If TestConnection Then
                Unload Me
            Else
                MsgBox "Invalid Database or Access.", vbInformation
            End If
    End Select
End Sub

Private Sub Form_Load()
    LoadSettings
    Text2.Text = Settings.Location
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

