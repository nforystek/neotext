VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   45
      ScaleHeight     =   3675
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   45
      Width           =   6150
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Purchase"
         Height          =   315
         Index           =   1
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3285
         Width           =   1080
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   315
         Index           =   0
         Left            =   4935
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3285
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   1110
         Picture         =   "frmAbout.frx":08CA
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.neotextsoftware.com"
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
         Left            =   0
         TabIndex        =   3
         Top             =   2280
         Width           =   6150
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   465
         Left            =   0
         TabIndex        =   2
         Top             =   1410
         Width           =   6150
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   1005
         Width           =   6150
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


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            OpenWebsite AppPurchase
    End Select
End Sub

Private Sub Form_Load()
    'Label1.Caption = ProgramName
    Label2.Caption = "Version: " + AppName & " v" & AppVersion
    Label3.Caption = AppCopyRight
    Label4.Caption = AppWebsite
End Sub

Private Sub Label4_Click()
    OpenWebsite Me, WebSite, False
End Sub
