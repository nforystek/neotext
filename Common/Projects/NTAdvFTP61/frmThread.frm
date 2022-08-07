VERSION 5.00
Begin VB.Form frmThread 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NTAdvFTP61.Socket"
   ClientHeight    =   1755
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "frmThread.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private sck As ISocket


Public Property Get Socket() As ISocket
    Set Socket = sck
End Property

Public Property Set Socket(ByRef RHS As ISocket)
    Set sck = RHS
End Property
