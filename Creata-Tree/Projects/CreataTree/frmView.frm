VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image View"
   ClientHeight    =   2100
   ClientLeft      =   6585
   ClientTop       =   3735
   ClientWidth     =   2220
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   2220
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   555
      Left            =   420
      Stretch         =   -1  'True
      Top             =   375
      Width           =   1575
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Function LoadImage(ByVal FileName As String)
    
    Image1.Picture = LoadPicture(FileName)
    
    Form_Resize
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Image1.Top = 0
    Image1.Left = 0
    Image1.Height = Me.ScaleHeight
    Image1.Width = Me.ScaleWidth
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

