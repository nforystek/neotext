VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklawn 3D"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmSplash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4785
      Left            =   15
      Picture         =   "frmSplash.frx":0E42
      ScaleHeight     =   4725
      ScaleWidth      =   9240
      TabIndex        =   0
      Top             =   15
      Width           =   9300
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      Height          =   4800
      Left            =   510
      Picture         =   "frmSplash.frx":8F7A4
      ScaleHeight     =   4740
      ScaleWidth      =   9240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5520
      Width           =   9300
   End
   Begin VB.PictureBox PicMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   10740
      Picture         =   "frmSplash.frx":11E106
      ScaleHeight     =   4140
      ScaleWidth      =   9240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4995
      Width           =   9240
   End
   Begin VB.PictureBox PicText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   10650
      Picture         =   "frmSplash.frx":147D68
      ScaleHeight     =   4140
      ScaleWidth      =   9240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   9240
   End
   Begin VB.PictureBox PicShow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4800
      Left            =   0
      Picture         =   "frmSplash.frx":1719CA
      ScaleHeight     =   4740
      ScaleWidth      =   9240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9300
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8700
         Top             =   3615
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const NOTSRCCOPY = &H330008

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private DrawAt As RECT

Public Skip As Boolean

Public Sub PlayIntro()
    Picture1.Visible = False
    PicShow.Visible = True
    Timer1.Enabled = True
    PicShow.SetFocus
    Do While frmSplash.Visible
        DoTasks
    Loop
    Picture1.Visible = True
    PicShow.Visible = False
    Me.Show
    DoTasks
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicShow.Visible Then
        FinishSplash
    Else
        Skip = True
    End If
End Sub

Private Sub Form_Load()
    DrawAt.Top = PicShow.height
    DrawAt.Left = 0
    DrawAt.Bottom = Me.ScaleHeight
    DrawAt.Right = Me.ScaleWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Me.Hide
End Sub

Private Sub PicShow_KeyDown(KeyCode As Integer, Shift As Integer)
    FinishSplash
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Skip = True
End Sub

Private Sub Timer1_Timer()

    If Not Skip Then
    
        PicShow.Cls
        PicBack.Cls
    
        BitBlt PicShow.hDC, DrawAt.Left, DrawAt.Top, DrawAt.Right, DrawAt.Bottom - DrawAt.Top, PicMask.hDC, 0, 0, NOTSRCCOPY
        BitBlt PicShow.hDC, DrawAt.Left, DrawAt.Top, DrawAt.Right, DrawAt.Bottom - DrawAt.Top, PicText.hDC, 0, 0, SRCCOPY
        BitBlt PicBack.hDC, DrawAt.Left, DrawAt.Top, DrawAt.Right, DrawAt.Bottom - DrawAt.Top, PicMask.hDC, 0, 0, SRCAND
        
        PicBack.Refresh
        
        BitBlt PicShow.hDC, 0, 0, PicShow.width, PicShow.height, PicBack.hDC, 0, 0, SRCPAINT
    
        DrawAt.Top = DrawAt.Top - 1
        If (DrawAt.Top + DrawAt.Bottom) < 0 Then
            FinishSplash
        End If
    Else
        FinishSplash
    End If
End Sub

Private Sub FinishSplash()
    Timer1.Enabled = False
    Me.Hide
End Sub
