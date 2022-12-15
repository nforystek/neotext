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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox Check3 
      Caption         =   "Silent Mode"
      Height          =   195
      Left            =   2745
      TabIndex        =   5
      Top             =   525
      Value           =   1  'Checked
      Width           =   1320
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
      Default         =   -1  'True
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
Option Compare Binary

Public Property Get Play() As Boolean
    Play = (Me.Tag = "Play")
End Property
Private Sub Command1_Click()

    If FullScreen <> (Check1.Value = 1) Then
        FullScreen = (Check1.Value = 1)
        SaveSetting AppEXE(True, True), "System", "FullScreen", FullScreen
        Me.Tag = "Play"
    End If

    If SilentMode <> (Check3.Value = 1) Then
        SilentMode = (Check3.Value = 1)
        SaveSetting AppEXE(True, True), "System", "SilentMode", SilentMode
        Me.Tag = "Play"
    End If

    If Resolution <> Combo1.List(Combo1.ListIndex) Then
        Resolution = Combo1.List(Combo1.ListIndex)
        SaveSetting AppEXE(True, True), "System", "Resolution", Resolution
        Me.Tag = "Play"
    End If

    Me.Visible = False
End Sub

Private Sub Command2_Click()
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
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
                Combo1.AddItem Display.Width & "x" & Display.Height
            End If
            '16 bit pixel format
          Else
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
                Combo1.AddItem Display.Width & "x" & Display.Height
            End If
        End If
    Next i
    
    Set dx = Nothing
    Set D3D = Nothing
    
End Sub

Private Sub Form_Load()
    Me.Tag = "Close"
    
    EnumDispModes
   
    If Combo1.ListCount > 0 Then
        Dim cnt As Integer
        
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = Resolution Then
                Combo1.ListIndex = cnt
            End If
        Next
    End If
    Check1.Value = -CInt(FullScreen)
    Check3.Value = -CInt(SilentMode)

    If (Combo1.Text = "") And (Combo1.ListCount > 0) Then Combo1.ListIndex = 0
End Sub
