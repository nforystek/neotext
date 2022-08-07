VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check4 
      Caption         =   "Music"
      Height          =   195
      Left            =   2820
      TabIndex        =   5
      Top             =   495
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   360
      Left            =   2865
      TabIndex        =   4
      Top             =   810
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   360
      Left            =   1725
      TabIndex        =   3
      Top             =   810
      Width           =   1065
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

Private Sub Check4_Click()
    PlayMusic = (Check4.Value = 1)
    db.dbQuery "UPDATE Settings SET MusicEnabled = " & IIf(Check4.Value = 1, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
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
                Combo1.AddItem Display.width & "x" & Display.height
            End If
            '16 bit pixel format
          Else
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.width & "x" & Display.height)) Then
                Combo1.AddItem Display.width & "x" & Display.height
            End If
        End If
    Next i
    
    Set dx = Nothing
    Set D3D = Nothing
    
End Sub

Private Sub Form_Load()

    EnumDispModes
    
    Resolution = "1024x768"
    FullScreen = (Check1.Value = 1)
    PlayMusic = (Check4.Value = 1)
    
    Dim cnt As Long
    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    If Not db.rsEnd(rs) Then
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = rs("Resolution") Then
                Combo1.ListIndex = cnt
            End If
        Next
        Check1.Value = -CInt(Not CBool(rs("Windowed")))
        Check4.Value = -CInt(CBool(rs("MusicEnabled")))
    Else
        db.dbQuery "INSERT INTO Settings (Username, Resolution, Windowed, MusicEnabled) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "', '1024x768', Yes, Yes);"
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = Resolution Then
                Combo1.ListIndex = cnt
            End If
        Next
    End If
    If (Combo1.text = "") And (Combo1.ListCount > 0) Then Combo1.ListIndex = 0
    db.rsClose rs
End Sub

