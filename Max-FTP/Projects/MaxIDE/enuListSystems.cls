VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuListSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary


'Public Enum ListSystems
'    None = 0
'    Wins = -1
'    Unix = -2
'End Enum

Public Property Get None() As Variant
    None = 0
End Property
Public Property Get Wins() As Variant
    Wins = -1
End Property
Public Property Get Unix() As Variant
    Unix = -2
End Property


Attribute 