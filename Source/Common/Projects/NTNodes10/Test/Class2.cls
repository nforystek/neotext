VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Var1 As Long

Public Var2 As String

Public var8 As Object

Private pvar10 As Class1

Public Property Get var9() As Object
    Set var9 = New Class2
End Property
Public Property Set var9(ByRef RHS As Object)

End Property

'Public Property Get Serialize() As String
'    Serialize = "sending"
'End Property
'Public Property Let Serialize(ByVal RHS As String)
'    Debug.Print RHS
'End Property


Public Property Get var10() As Class1
    Set var10 = pvar10
End Property

Public Property Set var10(ByRef RHS As Class1)
    Set pvar10 = RHS
End Property


Private Sub Class_Initialize()
    Set var8 = New Class2
    Set pvar10 = Me
End Sub
