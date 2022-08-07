VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuConnectionModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum ConnectionModes
'    Passive = 0
'    Active = 1
'End Enum


Public Property Get Passive() As Variant
    Passive = 0
End Property
Public Property Get Active() As Variant
    Active = 1
End Property


Attribute 