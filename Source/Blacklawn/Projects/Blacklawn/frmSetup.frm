
VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklawn Options"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check6 
      Caption         =   "Clear Multiplayer Score"
      Height          =   300
      Left            =   90
      TabIndex        =   9
      Top             =   825
      Width           =   2025
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Intro"
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   480
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Music"
      Height          =   195
      Left            =   1860
      TabIndex        =   7
      Top             =   480
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sound"
      Height          =   240
      Left            =   975
      TabIndex        =   6
      Top             =   465
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "2D Version"
      Height          =   240
      Left            =   2805
      TabIndex        =   5
      Top             =   465
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   360
      Left            =   3165
      TabIndex        =   4
      Top             =   810
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   360
      Left            =   2250
      TabIndex        =   3
      Top             =   810
      Width           =   810
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Full Screen"
      Height          =   225
      Left            =   2805
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Resolution:"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   885
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Play() As Boolean
    Play = (Me.Tag = "Play")
End Property

Private Sub Check1_Click()
    FullScreen = (Check1.Value = 1)
    db.dbQuery "UPDATE Settings SET Windowed = " & IIf(Check1.Value = 1, "No", "Yes") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Check2_Click()
    Version2D = (Check2.Value = 1)
    db.dbQuery "UPDATE Settings SET Version2D = " & IIf(Check2.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    Check1.Enabled = (Check2.Value = 0)
    Check3.Enabled = (Check2.Value = 0)
    Check5.Enabled = (Check2.Value = 0)
    Check6.Enabled = (Check2.Value = 0)
End Sub

Private Sub Check5_Click()
    ViewIntro = (Check5.Value = 1)
    db.dbQuery "UPDATE Settings SET ViewIntro = " & IIf(Check5.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Check3_Click()
    PlaySound = (Check3.Value = 1)
    db.dbQuery "UPDATE Settings SET SoundEnabled = " & IIf(Check3.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Check4_Click()
    PlayMusic = (Check4.Value = 1)
    db.dbQuery "UPDATE Settings SET MusicEnabled = " & IIf(Check4.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Check6_Click()
    ClearScore = (Check6.Value = 1)
    db.dbQuery "UPDATE Settings SET ClearScore = " & IIf(Check6.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Combo1_Click()
    Resolution = Combo1.List(Combo1.ListIndex)
    db.dbQuery "UPDATE Settings SET Resolution = '" & Resolution & "' WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Sub Command1_Click()
    Me.Tag = "Play"
    Me.Visible = False
End Sub

Private Sub Command2_Click()
    Me.Tag = "Close"
    Me.Visible = False
End Sub

Private Sub EnumDispModes()

    Combo1.Clear
    
    Dim dx As DirectX8
    Dim D3D As Direct3D8
    
    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    
    Dim i As Integer
    For i = 0 To D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
        D3D.EnumAdapterModes D3DADAPTER_DEFAULT, i, Display
        '32 bit pixel format
        If Display.Format = D3DFMT_R8G8B8 Or Display.Format = D3DFMT_X8R8G8B8 Or Display.Format = D3DFMT_A8R8G8B8 Then
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.width & "x" & Display.height)) Then
                Combo1.addItem Display.width & "x" & Display.height
            End If
            '16 bit pixel format
          Else
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.width & "x" & Display.height)) Then
                Combo1.addItem Display.width & "x" & Display.height
            End If
        End If
    Next i
    
    Set dx = Nothing
    Set D3D = Nothing
    
End Sub

Private Sub Form_Load()

    EnumDispModes
    
    ViewIntro = (Check5.Value = 1)
    Resolution = "1024x768"
    FullScreen = (Check1.Value = 1)
    Version2D = (Check2.Value = 1)
    PlaySound = (Check3.Value = 1)
    PlayMusic = (Check4.Value = 1)
    ClearScore = (Check6.Value = 1)
    
    Dim cnt As Long
    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    If Not db.rsEnd(rs) Then
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = rs("Resolution") Then
                Combo1.ListIndex = cnt
            End If
        Next
        Check1.Value = -CInt(Not CBool(rs("Windowed")))
        Check2.Value = -CInt(CBool(rs("Version2D")))
        Check3.Value = -CInt(CBool(rs("SoundEnabled")))
        Check4.Value = -CInt(CBool(rs("MusicEnabled")))
        Check5.Value = -CInt(CBool(rs("ViewIntro")))
        Check6.Value = -CInt(CBool(rs("ClearScore")))
        
    Else
        db.dbQuery "INSERT INTO Settings (Username, SoundEnabled, MusicEnabled, ViewIntro, Resolution, ClearScore, Windowed, Version2D, Score1, Score2, Score3, Score1st, Score2nd) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "', Yes, Yes, Yes, '1024x768', No, Yes, No, 0, 0, 0, '1st 0(s)', '2nd 0(s) 0(s)');"
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = Resolution Then
                Combo1.ListIndex = cnt
            End If
        Next
    End If
    If (Combo1.Text = "") And (Combo1.ListCount > 0) Then Combo1.ListIndex = 0
    db.rsClose rs
End Sub
