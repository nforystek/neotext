VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'a simple (x,y,z) 3D coordinate
'or any other equivelent memory
'such as a Morph or normalize
Private pInterest As Square
Private pWaypoint As Single
Friend Property Get Interest() As Square
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Square)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Single
    Waypoint = pWaypoint
End Property
Friend Property Let Waypoint(ByRef RHS As Single)
    pWaypoint = RHS
End Property


Public Property Get X() As Single
    X = pInterest.X
End Property

Public Property Let X(ByVal RHS As Single)
    pInterest.X = RHS
End Property

Public Property Get Y() As Single
    Y = pInterest.Y
End Property

Public Property Let Y(ByVal RHS As Single)
    pInterest.Y = RHS
End Property

Public Property Get Z() As Single
    Z = pWaypoint
End Property

Public Property Let Z(ByVal RHS As Single)
    pWaypoint = RHS
End Property

Friend Function Clone() As Point
    Set Clone = New Point
    Set Clone.Interest = pInterest.Clone
    Clone.Waypoint = pWaypoint
End Function

Friend Function Equals(ByRef SameAs As Point) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And (pWaypoint = SameAs.Waypoint))
End Function
Public Property Get DirectX() As D3DVECTOR
    With DirectX
        .X = X
        .Y = Y
        .Z = Z
    End With
End Property
Public Property Let DirectX(ByRef RHS As D3DVECTOR)
    With RHS
        X = .X
        Y = .Y
        Z = .Z
    End With
End Property

Public Property Get Serialize() As String
Attribute Serialize.VB_UserMemId = 0
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = "[" & X & ", " & Y & ", " & Z & "]"
End Property
Public Property Let Serialize(ByVal RHS As String)
    RHS = RemoveQuotedArg(RHS, "[", "]")
    If IsNumeric(NextArg(RHS, ",")) Then
        pInterest.X = CSng(NextArg(RHS, ","))
    End If
    If IsNumeric(NextArg(RemoveArg(RHS, ","), ",")) Then
        pInterest.Y = CSng(NextArg(RemoveArg(RHS, ","), ","))
    End If
    If IsNumeric(NextArg(RemoveArg(RemoveArg(RHS, ","), ","), ",")) Then
        pWaypoint = CSng(NextArg(RemoveArg(RemoveArg(RHS, ","), ","), ","))
    End If
End Property

Private Sub Class_Initialize()
    Set pInterest = New Square
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
End Sub
