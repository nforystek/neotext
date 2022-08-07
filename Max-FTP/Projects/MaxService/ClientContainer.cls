VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private myFTPClient1 As NTAdvFTP61.Client
Private myFTPClient2 As NTAdvFTP61.Client

Public Property Get FTPClient1() As NTAdvFTP61.Client
    Set FTPClient1 = myFTPClient1
End Property

Public Property Set FTPClient1(ByVal newVal As NTAdvFTP61.Client)
    Set myFTPClient1 = newVal
End Property

Public Property Get FTPClient2() As NTAdvFTP61.Client
    Set FTPClient2 = myFTPClient2
End Property

Public Property Set FTPClient2(ByVal newVal As NTAdvFTP61.Client)
    Set myFTPClient2 = newVal
End Property

Public Sub InitializeClients()

    Set FTPClient1 = New NTAdvFTP61.Client
    Set FTPClient2 = New NTAdvFTP61.Client
    
End Sub

Public Sub TerminateClients()

    Set FTPClient1 = Nothing
    Set FTPClient2 = Nothing
    
End Sub


