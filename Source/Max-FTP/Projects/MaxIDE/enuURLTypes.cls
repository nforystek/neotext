VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuURLTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Unknown() As Variant
    Unknown = 0
End Property
Public Property Get File() As Variant
    File = 1
End Property
Public Property Get Remote() As Variant
    Remote = 2
End Property
Public Property Get FTP() As Variant
    FTP = 3
End Property
Public Property Get HTTP() As Variant
    HTTP = 4
End Property
Public Property Get Server() As Variant
    Server = 5
End Property


Attribute 