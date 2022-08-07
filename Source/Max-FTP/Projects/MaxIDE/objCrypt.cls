VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Private WithEvents myObject As ntcypher10.Crypt
Attribute myObject.VB_VarHelpID = -1

Private myName As String
Private myFrm As Form

Public HasProgress As String

Public Function InitObject(ByVal nName As String, ByRef nForm)
    myName = nName
    Set myFrm = nForm
    Set myObject = New ntcypher10.Crypt
End Function
Public Function GetObject()
    Set GetObject = myObject
End Function
Public Function TermObject()
    Set myObject = Nothing
    Set myFrm = Nothing
    myName = ""
End Function

Private Sub myObject_Progress(Percent As Integer)
    If Not HasProgress = "" Then myFrm.RaiseCallBack myName, HasProgress, Percent
End Sub
