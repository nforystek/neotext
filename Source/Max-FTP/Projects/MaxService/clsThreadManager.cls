VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public AllClients As New ntnodes10.Collection

Public Property Get GetClients(ByVal Index As Long) As MaxService.Application
    Set GetClients = AllClients(Index)
End Property
Public Property Set GetClients(ByVal Index As Long, ByVal newval As MaxService.Application)
    Set AllClients(Index) = newval
End Property

Public Function AddClients() As Long
    Dim obj As MaxService.Application
    Set obj = New MaxService.Application
    obj.InitializeClients
    AllClients.Add obj
    Set obj = Nothing
    AddClients = AllClients.Count
End Function

Public Sub RemoveClients(ByVal Index As Long)
    Dim cnt As Long
    Dim obj As MaxService.Application
    If AllClients.Count > 0 Then
        For cnt = 1 To AllClients.Count
            If cnt = Index Then
                Set obj = AllClients(cnt)
                obj.TerminateClients
                Set obj = Nothing
            End If
        Next
        AllClients.Remove Index
    End If
End Sub

Private Sub Class_Initialize()
    Set AllClients = New ntnodes10.Collection
End Sub

Private Sub Class_Terminate()
    If AllClients.Count > 0 Then
        Dim cnt As Long
        Dim obj As Application
        For cnt = 1 To AllClients.Count
            Set obj = AllClients(cnt)
            obj.TerminateClients
            Set obj = Nothing
        Next
        Do Until AllClients.Count = 0
            AllClients.Remove 1
        Loop
    End If
    Set AllClients = Nothing
End Sub
