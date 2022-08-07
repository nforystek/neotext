VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuMessageTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary


'Public Enum MessageTypes
'    Outgoing = 1
'    Incomming = 2
'    Incorrect = 3
'End Enum

Public Property Get Outgoing() As Variant
    Outgoing = 1
End Property

Public Property Get Incomming() As Variant
    Incomming = 2
End Property

Public Property Get Incorrect() As Variant
    Incorrect = 3
End Property

Attribute 