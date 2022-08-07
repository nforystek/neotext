VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuGatherTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

'Private Enum GatherTypes
'    None = 0
'    Banner = 1
'    MOTD = 2
'    Help = 3
'    Stat = 4
'End Enum

Public Property Get None() As Variant
    None = 0
End Property
Public Property Get Banner() As Variant
    Banner = 1
End Property
Public Property Get MOTD() As Variant
    MOTD = 2
End Property
Public Property Get Help() As Variant
    Help = 3
End Property
Public Property Get Stat() As Variant
    Stat = 4
End Property

