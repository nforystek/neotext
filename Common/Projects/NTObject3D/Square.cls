VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Square"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'simplest form is a (x,y)
'or a 2d coordinate, which
'is exponential of squared

Private pInterest As Single
Private pWaypoint As Single
Friend Property Get Interest() As Single
    Interest = pInterest
End Property
Friend Property Let Interest(ByRef RHS As Single)
    pInterest = RHS
End Property
Friend Property Get Waypoint() As Single
    Waypoint = pWaypoint
End Property
Friend Property Let Waypoint(ByRef RHS As Single)
    pWaypoint = RHS
End Property


Public Property Get X() As Single
    X = pInterest
End Property

Public Property Let X(ByVal RHS As Single)
     pInterest = RHS
End Property

Public Property Get Y() As Single
    Y = pWaypoint
End Property

Public Property Let Y(ByVal RHS As Single)
   pWaypoint = RHS
End Property

Friend Function Clone() As Square
    Set Clone = New Square
    Clone.Interest = pInterest
    Clone.Waypoint = pWaypoint
End Function

Friend Function Equals(ByRef SameAs As Square) As Boolean
    Equals = ((pInterest = SameAs.Interest) And (pWaypoint = SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = "[" & X & ", " & Y & "]"
End Property
Public Property Let Serialize(ByVal RHS As String)
    RHS = RemoveQuotedArg(RHS, "[", "]")
    If IsNumeric(NextArg(RHS, ",")) Then
        pInterest = CSng(NextArg(RHS, ","))
    End If
    If IsNumeric(RemoveArg(RHS, ",")) Then
        pWaypoint = CSng(RemoveArg(RHS, ","))
    End If
End Property
