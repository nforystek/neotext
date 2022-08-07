VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents myObject As ntsound20.Player
Attribute myObject.VB_VarHelpID = -1

Private myName As String
Private myFrm As Form

Public HasSoundNotify As String

Public Function InitObject(ByVal nName As String, ByRef nForm)
    myName = nName
    Set myFrm = nForm
    Set myObject = New ntsound20.Player
End Function
Public Property Get GetObject() As Object
    Set GetObject = myObject
End Property
Public Function TermObject()
    myObject.StopSound
    Set myObject = Nothing
    Set myFrm = Nothing
    myName = ""
End Function

Private Sub myObject_SoundNotify()
    If Not HasSoundNotify = "" Then myFrm.RaiseCallBack myName, HasSoundNotify
End Sub

Attribute 