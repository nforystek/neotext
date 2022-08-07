VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'a numerical value, most basic
'building block of all datatype
Public n As Single
Attribute n.VB_VarMemberFlags = "200"
Attribute n.VB_VarUserMemId = 0

Friend Function Clone() As Atom
    Set Clone = New Atom
    Clone.n = n
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (n = SameAs.n)
End Function

Public Property Get Serialize() As String
    Serialize = "[" & n & "]"
End Property
Public Property Let Serialize(ByVal RHS As String)
    RHS = RemoveQuotedArg(RHS, "[", "]")
    If IsNumeric(RHS) Then
        n = CSng(RHS)
    End If
End Property
