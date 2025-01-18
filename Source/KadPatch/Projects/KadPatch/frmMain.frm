VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "KadPatch"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13320
   ClipControls    =   0   'False
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":23D2
   MousePointer    =   1  'Arrow
   ScaleHeight     =   9210
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      ScaleHeight     =   1740
      ScaleWidth      =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   3720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    frmStudio.KeyDown KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmStudio.KeyPress KeyAscii
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmStudio.MouseMove Button, Shift, X, Y
End Sub

Friend Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmStudio.MouseMove Button, Shift, X, Y
End Sub

