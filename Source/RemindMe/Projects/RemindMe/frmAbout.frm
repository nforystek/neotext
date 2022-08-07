VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   45
      ScaleHeight     =   3165
      ScaleWidth      =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   6420
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5355
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2760
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   15
         Picture         =   "frmAbout.frx":0442
         Top             =   -15
         Width           =   6225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sosouix.net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   2250
         Width           =   6345
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   -45
         TabIndex        =   2
         Top             =   1530
         Width           =   6420
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me

    End Select
End Sub

Private Sub Label4_Click()
    RunFile AppPath & "SoSouiX.net.url"
End Sub

Private Sub Form_Load()

    Me.Caption = "About RemindMe " & App.Major & "." & App.Minor & "." & App.Revision
    
    Label2.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    Label4.Caption = WebSite

End Sub



