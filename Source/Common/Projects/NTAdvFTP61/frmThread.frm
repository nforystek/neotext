VERSION 5.00
Begin VB.Form frmThread 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NTAdvFTP61.Socket.Form"
   ClientHeight    =   540
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   2280
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2280
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

