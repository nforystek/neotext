VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuRateTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum RateTypes
'    HardDrive = 0
'    Download = 1
'    Upload = 2
'End Enum

Public Property Get HardDrive() As Variant
    HardDrive = 0
End Property

Public Property Get Download() As Variant
    Download = 1
End Property

Public Property Get Upload() As Variant
    Upload = 2
End Property


Attribute 