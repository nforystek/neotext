VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pX As Single
Private pY As Single
Private pZ As Single


Public Event Changed(ByRef X As Single, ByRef Y As Single, ByRef Z As Single)
Public Property Get Clone() As Vector
    Set Clone = New Vector
    Clone.X = pX
    Clone.Y = pY
    Clone.Z = pZ
End Property
Public Property Get X() As Single
    X = pX
End Property
Public Property Let X(ByVal RHS As Single)
    pX = RHS
    RaiseEvent Changed(pX, pY, pZ)
End Property

Public Property Get Y() As Single
    Y = pY
End Property
Public Property Let Y(ByVal RHS As Single)
    pY = RHS
    RaiseEvent Changed(pX, pY, pZ)
End Property

Public Property Get Z() As Single
    Z = pZ
End Property
Public Property Let Z(ByVal RHS As Single)
    pZ = RHS
    RaiseEvent Changed(pX, pY, pZ)
End Property

Public Sub Invert()
    pX = -pX
    pY = -pY
    pZ = -pZ
End Sub

Public Property Let ToString(ByVal RHS As String)
    If InStr(RHS, "[") > 0 Then RHS = RemoveQuotedArg(RHS, "[", "]")
    
    Dim nxt As String
    nxt = RemoveNextArg(RHS, ",")
    If IsNumeric(nxt) Then pX = CSng(nxt)
    nxt = RemoveNextArg(RHS, ",")
    If IsNumeric(nxt) Then pY = CSng(nxt)
    nxt = RemoveNextArg(RHS, ",")
    If IsNumeric(nxt) Then pZ = CSng(nxt)

    RaiseEvent Changed(pX, pY, pZ)
End Property
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "200"
    ToString = "[ " & pX & " , " & pY & ", " & pZ & "]"
End Property
