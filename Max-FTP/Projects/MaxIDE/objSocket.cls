VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents myObject As NTAdvFTP61.Socket
Attribute myObject.VB_VarHelpID = -1

Private myName As String
Private myFrm As Form

Public HasError As String
Public HasConnected As String
Public HasDataArriving As String
Public HasConnection As String
Public HasSendComplete As String
Public HasDisconnected As String

Public Function InitObject(ByVal nName As String, ByRef nForm)
    myName = nName
    Set myFrm = nForm
    Set myObject = New NTAdvFTP61.Socket
End Function
Public Property Get GetObject() As Object
    Set GetObject = myObject
End Property
Public Function TermObject()
    myObject.Disconnect
    Set myObject = Nothing
    Set myFrm = Nothing
    myName = ""
End Function

Private Sub myObject_Connected()
    If Not HasConnected = "" Then myFrm.RaiseCallBack myName, HasConnected
End Sub

Private Sub myObject_Connection(ByRef handle As Long)
    If Not HasConnection = "" Then myFrm.RaiseCallBack myName, HasConnection, handle
End Sub

Private Sub myObject_DataArriving()
    If Not HasDataArriving = "" Then myFrm.RaiseCallBack myName, HasDataArriving
End Sub

Private Sub myObject_Disconnected()
    If Not HasDisconnected = "" Then myFrm.RaiseCallBack myName, HasDisconnected
End Sub

Private Sub myObject_SendComplete()
    If Not HasSendComplete = "" Then myFrm.RaiseCallBack myName, HasSendComplete
End Sub

Private Sub myObject_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    If Not HasError = "" Then myFrm.RaiseCallBack myName, HasError, Number & "," & myFrm.SafeStr(Source) & "," & myFrm.SafeStr(Description) & ""
End Sub

Attribute 