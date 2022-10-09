VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary


Public Function AddClient(ByVal InRequest As String) As String
    On Error GoTo catch
    
    Dim IP As String
    Dim Port As Long
    Dim index As Long
    
    Port = 28732
    IP = RemoveNextArg(InRequest, ":")
    If IsNumeric(InRequest) Then Port = CLng(InRequest)
    
    frmBlkLServer.AddClient IP, Port
    
    Exit Function
catch:
    AddClient = Err.Description
    Err.Clear
End Function

