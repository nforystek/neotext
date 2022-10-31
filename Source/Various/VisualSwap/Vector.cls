VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pX As Long
Private pY As Long
Public Property Get X() As Long
    X = pX
End Property
Public Property Let X(ByVal RHS As Long)
    pX = RHS
End Property

Public Property Get Y() As Long
    Y = pY
End Property
Public Property Let Y(ByVal RHS As Long)
    pY = RHS
End Property

