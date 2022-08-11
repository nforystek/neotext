VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Neotext (a front end Visual Basic handler)"
   ClientHeight    =   7716
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmIcon.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7716
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   492
      Top             =   225
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   5748
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/signonly or /s and /timeonly or /t projectname"
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   156
         TabIndex        =   29
         Top             =   4584
         Width           =   1692
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":014A
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   1992
         TabIndex        =   28
         Top             =   4584
         Width           =   6576
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":0258
         ForeColor       =   &H80000008&
         Height          =   408
         Left            =   1992
         TabIndex        =   26
         Top             =   5220
         Width           =   6576
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/open or /o projectname"
         ForeColor       =   &H80000008&
         Height          =   408
         Left            =   168
         TabIndex        =   25
         Top             =   5220
         Width           =   1692
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":02F7
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   24
         Top             =   4140
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
         Top             =   4140
         Width           =   1695
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":03B1
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   22
         Top             =   3705
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
         Top             =   3705
         Width           =   1695
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":0466
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   20
         Top             =   3270
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
         Top             =   3270
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":051A
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   18
         Top             =   2835
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
         Top             =   2835
         Width           =   1695
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":05C3
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   16
         Top             =   2400
         Width           =   6570
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "/d or /D const-value..."
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   15
         Top             =   2400
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
         Top             =   2055
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
         Top             =   2055
         Width           =   1695
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":067B
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   12
         Top             =   1575
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
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":0718
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   10
         Top             =   1140
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
         Top             =   1140
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
         Top             =   705
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
         Top             =   705
         Width           =   1695
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIcon.frx":07A2
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1995
         TabIndex        =   6
         Top             =   270
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
         Top             =   270
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
      Caption         =   $"frmIcon.frx":0850
      Height          =   888
      Left            =   840
      TabIndex        =   27
      Top             =   6348
      Width           =   8148
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VB6[.EXE]"
      Height          =   336
      Left            =   240
      TabIndex        =   4
      Top             =   96
      Width           =   876
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIcon.frx":0A14
      Height          =   588
      Left            =   1320
      TabIndex        =   3
      Top             =   96
      Width           =   7572
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   228
      Picture         =   "frmIcon.frx":0AF1
      Top             =   336
      Visible         =   0   'False
      Width           =   384
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
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Binary

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private pVBInstance As VBIDE.VBE

Public Function GetUserLoginName() As String

    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        GetUserLoginName = Left$(sBuffer, lSize)
    End If
End Function

Public Function GetMachineName() As String

    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetCompName(sBuffer, lSize)
    If lSize > 0 Then
        GetMachineName = Left$(sBuffer, lSize)
    End If

End Function

Friend Property Get VBInstance() As VBIDE.VBE
    Set VBInstance = pVBInstance
End Property
Friend Property Set VBInstance(ByRef newVal As VBIDE.VBE)
    Set pVBInstance = newVal
End Property

Public Sub ShowHelp(ByVal msgtext As String)
    App_Load = True
    Label3.Visible = True
    Frame1.Visible = True
    Label2.Visible = True
    Label6.Visible = True
    Label1.Visible = False
    Image1.Visible = False

    Label1.Left = (40 * Screen.TwipsPerPixelX) + Image1.Width
    Label1.Width = Me.TextWidth(msgtext)
    Label1.Caption = msgtext
    Command1.Left = (Me.ScaleWidth / 2) - (Command1.Width / 2)

    Label1.Height = (57 + 382) * Screen.TwipsPerPixelX

    Me.Caption = "Basic Neotext"
    Me.Show
    DoEvents
    TopMostForm Me, True, False
    TopMostForm Me, False, True
End Sub

Public Sub ShowMessage(ByVal msgtext As String)
    App_Load = True
    Label3.Visible = False
    Frame1.Visible = False
    Label2.Visible = False
    Label6.Visible = False
    Label1.Visible = True
    Image1.Visible = True
    Me.Height = (148 * Screen.TwipsPerPixelY)
    Me.Width = (60 * Screen.TwipsPerPixelX) + Image1.Width + Me.TextWidth(msgtext)

    Label1.Left = (40 * Screen.TwipsPerPixelX) + Image1.Width
    Label1.Width = Me.TextWidth(msgtext)
    Label1.Height = 57 * Screen.TwipsPerPixelX
    Label1.Caption = msgtext
    Command1.Top = Label1.Top + Label1.Height + (10 * Screen.TwipsPerPixelY)
    'Command1.Left = 212 * Screen.TwipsPerPixelX
    Command1.Left = (Me.ScaleWidth / 2) - (Command1.Width / 2)
    Me.Caption = "Error"
    Me.Show
    DoEvents
    TopMostForm Me, True, False
    TopMostForm Me, False, True
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    App_Path = GetFilePath(AppEXE(False)) & "\"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Load()
    App_Load = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
    Else
        Set VBInstance = Nothing
        App_Load = False
    End If
End Sub

Private Sub Timer1_Timer()

    On Error GoTo exitthis
    
    If (((VBNProjects.Count = 0) And Not Me.Visible) And Not _
        ((Me.Caption = "Basic Neotext") And Me.Visible)) Then

        Unload Me

    End If

    If frmIcon.Timer1.Tag = "POSTCOMPILE" Then
        Dim inst As Project
        For Each inst In VBNProjects.Items
            If (Not (VBInstance Is Nothing)) Then
                If (InStr(VBInstance.MainWindow.Caption, "[design]") > 0) And (ProcessRunning("C2.EXE") = 0) Then
                    If FindWindow("#32770", "Microsoft Visual Basic") = 0 And FindWindow("#32770", "Make Project") = 0 Then
                        frmIcon.Timer1.Tag = ""
                        PostCompileTend GetProject(VBInstance.ActiveVBProject.FileName, False, Me)
                        CoFreeUnusedLibraries
                    End If
                End If
            End If
        Next
    End If

    Exit Sub
exitthis:
    Err.Clear

End Sub
