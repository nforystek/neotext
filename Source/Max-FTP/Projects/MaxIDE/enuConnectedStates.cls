VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enuConnectedStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

'Public Enum ConnectedStates
'    NotState = 0
'    AllStates = -1
'    LocalHardDisk = 1
'    SocketNetwork = 2
'    LoggingInto = 4
'    FullyRemoting = 8
'    CommandRequest = 16
'    ListingToFile = 32
'    LocalToLocal = 64
'    LocalToRemote = 128
'    RemoteToLocal = 256
'    DataStreaming = 512
'End Enum
Public Property Get NotState() As Variant
    NotState = 0
End Property
Public Property Get AllStates() As Variant
    AllStates = -1
End Property
Public Property Get LocalHardDisk() As Variant
    LocalHardDisk = 1
End Property
Public Property Get SocketNetwork() As Variant
    SocketNetwork = 2
End Property
Public Property Get LoggingInto() As Variant
    LoggingInto = 4
End Property
Public Property Get FullyRemoting() As Variant
    FullyRemoting = 8
End Property
Public Property Get CommandRequest() As Variant
    CommandRequest = 16
End Property
Public Property Get ListingToFile() As Variant
    ListingToFile = 32
End Property
Public Property Get LocalToLocal() As Variant
    LocalToLocal = 64
End Property
Public Property Get LocalToRemote() As Variant
    LocalToRemote = 128
End Property
Public Property Get RemoteToLocal() As Variant
    RemoteToLocal = 256
End Property
Public Property Get DataStreaming() As Variant
    DataStreaming = 512
End Property




Attribute 