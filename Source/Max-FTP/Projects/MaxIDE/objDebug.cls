VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Text() As String
    Text = frmDebug.Text1.Text
End Property
Public Property Let Text(ByVal newVal As String)
    frmDebug.Text1.Text = newVal
End Property

Public Sub PrintText(ByVal Text As Variant)
    On Error Resume Next
    frmDebug.PrintText CStr(Text)
    If Not (Err.Number = 0) Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub PrintLine(ByVal Text As String)
    On Error Resume Next
    frmDebug.PrintLine Text
    If Not (Err.Number = 0) Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub Clear()
    frmDebug.Text1.Text = ""
End Sub


Attribute 