VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neotext Basic"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmHelp.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Height          =   360
      Left            =   5610
      TabIndex        =   31
      Top             =   7830
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2760
      TabIndex        =   30
      Top             =   7830
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   492
      Top             =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   4185
      TabIndex        =   1
      Top             =   7830
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   6360
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   8760
      Begin VB.Label Label30 
         Caption         =   $"frmHelp.frx":23E2
         Height          =   645
         Left            =   2010
         TabIndex        =   33
         Top             =   5655
         Width           =   6465
      End
      Begin VB.Label Label29 
         Caption         =   "/copy sourcefolder destfolder"
         Height          =   405
         Left            =   165
         TabIndex        =   32
         Top             =   5655
         Width           =   1635
      End
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
         Caption         =   $"frmHelp.frx":24E3
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
         Caption         =   $"frmHelp.frx":25F1
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
         Caption         =   $"frmHelp.frx":2690
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
         Caption         =   $"frmHelp.frx":274A
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
         Caption         =   $"frmHelp.frx":27FF
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
         Caption         =   $"frmHelp.frx":28B3
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
         Caption         =   $"frmHelp.frx":295C
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
         Caption         =   $"frmHelp.frx":2A14
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
         Caption         =   $"frmHelp.frx":2AB1
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
         Caption         =   $"frmHelp.frx":2B3B
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
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   $"frmHelp.frx":2BE9
      Height          =   885
      Left            =   780
      TabIndex        =   27
      Top             =   6945
      Width           =   8145
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VBN[.EXE]"
      Height          =   330
      Left            =   195
      TabIndex        =   4
      Top             =   90
      Width           =   870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2DA7
      Height          =   585
      Left            =   1215
      TabIndex        =   3
      Top             =   90
      Width           =   7665
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmHelp.frx":2E94
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

Private Sub Command1_Click() ' _

Attribute Command1_Click.VB_Description = ""
    If Command1.Caption = "&No" Then
        Me.Tag = vbNo
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub Command2_Click() ' _

Attribute Command2_Click.VB_Description = ""
    Me.Tag = vbYes
    Me.Hide
End Sub

Private Sub Command3_Click() ' _

Attribute Command3_Click.VB_Description = ""
    Me.Tag = vbCancel
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) ' _

Attribute Form_QueryUnload.VB_Description = ""
    If (Command1.Caption = "&No" And UnloadMode = 0) Then
        Cancel = True
        Me.Tag = vbCancel
        Me.Hide
    End If
    
End Sub

