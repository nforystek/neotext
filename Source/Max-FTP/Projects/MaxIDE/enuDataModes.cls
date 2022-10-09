VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuDataModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get ListString() As Variant
    ListString = 0
End Property

Public Property Get ListFile() As Variant
    ListFile = 1
End Property

Public Property Get GetFile() As Variant
    GetFile = 2
End Property

Public Property Get PutFile() As Variant
    PutFile = 3
End Property

Public Property Get PutFileLocal() As Variant
    PutFileLocal = 4
End Property


Attribute 