VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuProgressTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum ProgressTypes
'    FileListing = 0
'    AllocatingFile = 1
'    PositioningFile = 2
'    TransferingFile = 3
'End Enum

Public Property Get FileListing() As Variant
    FileListing = 0
End Property

Public Property Get AllocatingFile() As Variant
    AllocatingFile = 1
End Property

Public Property Get PositioningFile() As Variant
    PositioningFile = 2
End Property

Public Property Get TransferingFile() As Variant
    TransferingFile = 3
End Property


Attribute 