
VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neotext Basic"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmHelp.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   492
      Top             =   345
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   5748
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   8760
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/signonly or /s and /timeonly or /t projectname"
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   150
         TabIndex        =   29
         Top             =   4545
         Width           =   1695
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":23E2
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   1995
         TabIndex        =   28
         Top             =   4545
         Width           =   6570
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":24F0
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   26
         Top             =   5205
         Width           =   6570
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/open or /o projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   25
         Top             =   5205
         Width           =   1695
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":258F
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   24
         Top             =   4065
         Width           =   6570
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/signmake or /sm projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   23
         Top             =   4065
         Width           =   1695
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":2649
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   22
         Top             =   3600
         Width           =   6570
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/sign or /s projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   21
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":26FE
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   20
         Top             =   3150
         Width           =   6570
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/mdi or /sdi"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   19
         Top             =   3150
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":27B2
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   18
         Top             =   2685
         Width           =   6570
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/cmd or /c argument"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   17
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":285B
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   16
         Top             =   2250
         Width           =   6570
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/d or /D const=value..."
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   15
         Top             =   2250
         Width           =   1695
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Specifies a directory path to place all output filesin when using /make."
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1995
         TabIndex        =   14
         Top             =   1995
         Width           =   6570
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/outdir path"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   165
         TabIndex        =   13
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":2913
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   12
         Top             =   1560
         Width           =   6570
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/out filename"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":29B0
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   10
         Top             =   1125
         Width           =   6570
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/make or /m projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   9
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tells Visual Basic to compile projectname and run it.  VIsual Basic will exit when the projest returns to deign mode."
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   8
         Top             =   690
         Width           =   6570
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/runexit projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   7
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":2A3A
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   6
         Top             =   255
         Width           =   6570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/run or /r projectname"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   5
         Top             =   255
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   4245
      TabIndex        =   1
      Top             =   7224
      Width           =   1170
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   $"frmHelp.frx":2AE8
      Height          =   888
      Left            =   840
      TabIndex        =   27
      Top             =   6348
      Width           =   8148
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VBN[.EXE]"
      Height          =   336
      Left            =   240
      TabIndex        =   4
      Top             =   96
      Width           =   876
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2CA6
      Height          =   588
      Left            =   1320
      TabIndex        =   3
      Top             =   96
      Width           =   7572
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmHelp.frx":2D83
      Top             =   330
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   1005
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   6240
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Binary


Private Sub Command1_Click()
    Unload Me
End Sub

