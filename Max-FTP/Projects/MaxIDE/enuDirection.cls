VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuDirection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Enum Direction
'    NormalStream = 1
'    TextMessages = 2
'    SocketPaused = 6
'    StandByPause = 8
'End Enum

Public Property Get NormalStream() As Variant
    NormalStream = 1
End Property

Public Property Get TextMessages() As Variant
    TextMessages = 2
End Property

Public Property Get SocketPaused() As Variant
    SocketPaused = 6
End Property

Public Property Get StandByPause() As Variant
    StandByPause = 8
End Property

Attribute 