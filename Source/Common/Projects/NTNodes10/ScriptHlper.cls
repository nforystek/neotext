VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptHlper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NewObject As IUnknown

Public Sub CreateObjectPrivate(ByVal Classname As String)
    On Error Resume Next
    Set NewObject = modFactory.CreateObjectPrivate(Classname)
    If Err Then
        Err.Clear
        Set NewObject = CreateObject(Classname)
    End If
End Sub
