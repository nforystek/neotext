


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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private AllClients() As MaxFTP.Application
Private ClientsCount As Long

Public Property Get GetClients(ByVal Index As Long) As MaxFTP.Application
    Set GetClients = AllClients(Index)
End Property

Public Function AddClients() As Long
    ClientsCount = ClientsCount + 1
    
    ReDim Preserve AllClients(1 To ClientsCount) As MaxFTP.Application
    Set AllClients(ClientsCount) = CreateObject("MaxFTP.Application") 'New MaxFTP.Application
    AllClients(ClientsCount).InitializeClients

    AddClients = ClientsCount
End Function

Public Sub RemoveClients(ByVal Index As Long)
    Dim cnt As Long
    If ClientsCount > 1 Then
        For cnt = 1 To ClientsCount
            If cnt = Index Then
                If Not AllClients(cnt) Is Nothing Then
                    AllClients(cnt).TerminateClients
                    Set AllClients(cnt) = Nothing
                End If
            ElseIf cnt > Index Then
            
                If Not cnt = ClientsCount Then
                    Set AllClients(cnt) = AllClients(cnt + 1)
                End If
            End If
        Next
        ClientsCount = ClientsCount - 1
        ReDim Preserve AllClients(1 To ClientsCount) As MaxFTP.Application
    Else
        ClientsCount = ClientsCount - 1
        Set AllClients(1) = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    If ClientsCount > 1 Then
        Dim cnt As Long
        For cnt = 1 To ClientsCount
            If Not TypeName(AllClients(cnt)) = "Nothing" Then
                AllClients(cnt).TerminateClients
            End If
            Set AllClients(cnt) = Nothing
        Next
    ElseIf ClientsCount = 1 Then
        ClientsCount = ClientsCount - 1
        If Not TypeName(AllClients(1)) = "Nothing" Then
            AllClients(1).TerminateClients
        End If
        Set AllClients(1) = Nothing
    End If
End Sub