VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "House Of Glass"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2565
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
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   885
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Form_Click()
    TrapMouse = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    StopGame = True
End Sub
