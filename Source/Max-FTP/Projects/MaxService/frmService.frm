VERSION 5.00
Begin VB.Form frmService 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Max-FTP Schedule Service"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "DDEServer"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Form_Load()
    Me.Caption = ServiceFormCaption
End Sub


