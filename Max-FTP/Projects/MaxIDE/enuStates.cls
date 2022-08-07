VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum States
'    Closed = 1
'    Binded = 2
'    Request = 4
'    Opening = 8
'    Opened = 16
'    Sending = 32
'    Reading = 64
'    Closing = 128
'End Enum

Public Property Get Closed() As Variant
    Closed = 1
End Property

Public Property Get Binded() As Variant
    Binded = 2
End Property

Public Property Get Request() As Variant
    Request = 4
End Property

Public Property Get Opening() As Variant
    Opening = 8
End Property

Public Property Get Opened() As Variant
    Opened = 16
End Property

Public Property Get Sending() As Variant
    Sending = 32
End Property

Public Property Get Reading() As Variant
    Reading = 64
End Property

Public Property Get Closing() As Variant
    Closing = 128
End Property

Attribute 