VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents myObject As NTSchedule20.Schedule
Attribute myObject.VB_VarHelpID = -1

Private myName As String
Private myFrm As Form

Public HasScheduledEvent As String

Public Function InitObject(ByVal nName As String, ByRef nForm)
    myName = nName
    Set myFrm = nForm
    Set myObject = New NTSchedule20.Schedule
End Function
Public Property Get GetObject() As Object
    Set GetObject = myObject
End Property
Public Function TermObject()
    myObject.Enabled = False
    Set myObject = Nothing
    Set myFrm = Nothing
    myName = ""
End Function

Private Sub myObject_ScheduledEvent()
    If Not HasScheduledEvent = "" Then myFrm.RaiseCallBack myName, HasScheduledEvent
End Sub


Attribute 