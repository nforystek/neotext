VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuErrorReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum ErrorReturns
'    Void = -1
'    None = 0
'    Last = 1
'End Enum

Public Property Get Void() As Variant
    Void = -1
End Property

Public Property Get None() As Variant
    None = 0
End Property
Public Property Get Last() As Variant
    Last = 1
End Property


Attribute 