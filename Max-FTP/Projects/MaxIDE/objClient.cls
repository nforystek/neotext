VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents myObject As NTAdvFTP61.Client
Attribute myObject.VB_VarHelpID = -1

Private myName As String
Private myFrm As Form

Public HasError As String
Public HasLogMessage As String
Public HasItemListing As String
Public HasDataProgress As String
Public HasDataComplete As String

Public Function InitObject(ByVal nName As String, ByRef nForm)
    myName = nName
    Set myFrm = nForm
    Set myObject = New NTAdvFTP61.Client
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

Private Sub myObject_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    If Not HasError = "" Then myFrm.RaiseCallBack myName, HasError, Number & "," & myFrm.SafeStr(Source) & "," & myFrm.SafeStr(Description) & ""
End Sub

Private Sub myObject_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    If Not HasDataProgress = "" Then myFrm.RaiseCallBack myName, HasDataProgress, ProgressType & "," & Replace(ReceivedBytes, ",", "")
End Sub

Private Sub myObject_ItemListing(ByVal FileName As String, ByVal FileSize As String, ByVal FileDate As String, ByVal FileAccess As String)
    If Not HasItemListing = "" Then myFrm.RaiseCallBack myName, HasItemListing, myFrm.SafeStr(FileName) & "," & myFrm.SafeStr(FileSize) & "," & myFrm.SafeStr(FileDate) & "," & myFrm.SafeStr(FileAccess)
End Sub

Private Sub myObject_LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)
    If Not HasLogMessage = "" Then myFrm.RaiseCallBack myName, HasLogMessage, MessageType & "," & myFrm.SafeStr(AddedText)
End Sub

Private Sub myObject_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
    If Not HasDataComplete = "" Then myFrm.RaiseCallBack myName, HasDataComplete, ProgressType
End Sub


Attribute 