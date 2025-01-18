VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "KadPatch Slpash Window"
   ClientHeight    =   5160
   ClientLeft      =   15
   ClientTop       =   -60
   ClientWidth     =   11430
   ClipControls    =   0   'False
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSetup.frx":000C
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   762
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   10605
      Top             =   3645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSetup.frx":C02CE
      Height          =   210
      Left            =   330
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   9915
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Private Sub EnumDispModes()
'
'    Combo1.Clear
'
'    Dim dx As DirectX8
'    Dim D3D As Direct3D8
'
'    Set dx = New DirectX8
'    Set D3D = dx.Direct3DCreate
'
'    Dim i As Integer
'    For i = 0 To D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
'        D3D.EnumAdapterModes D3DADAPTER_DEFAULT, i, Display
'        '32 bit pixel format
'        If Display.Format = D3DFMT_R8G8B8 Or Display.Format = D3DFMT_X8R8G8B8 Or Display.Format = D3DFMT_A8R8G8B8 Then
'            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
'                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
'                Combo1.AddItem Display.Width & "x" & Display.Height
'            End If
'            '16 bit pixel format
'          Else
'            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
'                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
'                Combo1.AddItem Display.Width & "x" & Display.Height
'            End If
'        End If
'    Next i
'
'    Set dx = Nothing
'    Set D3D = Nothing
'
'End Sub
Public Sub ShowAbout()
    Label1.Visible = True


    Timer1.Enabled = False
    Me.Show
End Sub


Private Sub Form_Load()

'    EnumDispModes
'
'    Resolution = "1024x768"
'
'    Dim cnt As Long
'    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
'    If Not db.rsEnd(rs) Then
'        For cnt = 0 To Combo1.ListCount - 1
'            If Combo1.List(cnt) = rs("Resolution") Then
'                Combo1.ListIndex = cnt
'            End If
'        Next
'    Else
'        db.dbQuery "INSERT INTO Settings (Username, Resolution, Windowed) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "', '1024x768', Yes);"
'        For cnt = 0 To Combo1.ListCount - 1
'            If Combo1.List(cnt) = Resolution Then
'                Combo1.ListIndex = cnt
'            End If
'        Next
'    End If
'    If (Combo1.Text = "") And (Combo1.ListCount > 0) Then Combo1.ListIndex = 0
'    db.rsClose rs
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        Unload Me
End Sub

Private Sub Timer1_Timer()

    Unload Me
End Sub




