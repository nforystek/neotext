VERSION 5.00
Begin VB.Form frmEvents 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3744
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2256
      Top             =   456
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pParent As Driver

Public Property Get Parent() As Driver
    Set Parent = pParent
End Property
Public Property Set Parent(ByRef RHS As Driver)
    Set pParent = RHS
    Timer1.Enabled = True
End Property

Private Sub Timer1_Timer()
    pParent.RaiseCallBack
End Sub
