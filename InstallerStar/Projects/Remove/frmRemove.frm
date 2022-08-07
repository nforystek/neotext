VERSION 5.00
Begin VB.Form frmRemove 
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2910
   Icon            =   "frmRemove.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   930
      Top             =   180
   End
End
Attribute VB_Name = "frmRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Sub Timer1_Timer()
    Unload Me
End Sub
