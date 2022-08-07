VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private myFTPClient1
Private myFTPClient2

Public Property Get Count() As Long
    Count = ThreadManager.AllClients.Count
End Property
Public Property Get LeftClient()
    Set LeftClient = myFTPClient1
End Property

Friend Property Set LeftClient(ByVal newVal)
    Set myFTPClient1 = newVal
End Property

Public Property Get RightClient()
    Set RightClient = myFTPClient2
End Property

Friend Property Set RightClient(ByVal newVal)
    Set myFTPClient2 = newVal
End Property

Friend Sub InitializeClients()

    Set LeftClient = New NTAdvFTP61.Client
    Set RightClient = New NTAdvFTP61.Client

End Sub

Friend Sub TerminateClients()

    Set LeftClient = Nothing
    Set RightClient = Nothing
    
End Sub


