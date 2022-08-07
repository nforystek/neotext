VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'consists of a angle, (3 points) plus as a
'one more point, making 4 poins such faces
Private pInterest As Angle
Private pWaypoint As Point
Friend Property Get Interest() As Angle
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Angle)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Point
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property

Friend Property Get Angle() As Angle
    Set Angle = pInterest
End Property
Friend Property Set Angle(ByRef RHS As Angle)
    Set pInterest = RHS
End Property

Friend Property Get Point() As Point
    Set Point = pWaypoint
End Property
Friend Property Set Point(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property

Public Property Get X1() As Single
    X1 = pInterest.X1
End Property

Public Property Let X1(ByVal RHS As Single)
    pInterest.X1 = RHS
End Property

Public Property Get Y1() As Single
    Y1 = pInterest.Y1
End Property

Public Property Let Y1(ByVal RHS As Single)
    pInterest.Y1 = RHS
End Property

Public Property Get Z1() As Single
    Z1 = pInterest.Z1
End Property

Public Property Let Z1(ByVal RHS As Single)
    pInterest.Z1 = RHS
End Property


Public Property Get X2() As Single
    X2 = pInterest.X2
End Property

Public Property Let X2(ByVal RHS As Single)
    pInterest.X2 = RHS
End Property

Public Property Get Y2() As Single
    Y2 = pInterest.Y2
End Property

Public Property Let Y2(ByVal RHS As Single)
    pInterest.Y2 = RHS
End Property

Public Property Get Z2() As Single
    Z2 = pInterest.Z2
End Property

Public Property Let Z2(ByVal RHS As Single)
    pInterest.Z2 = RHS
End Property


Public Property Get X3() As Single
    X3 = pInterest.Z3
End Property

Public Property Let X3(ByVal RHS As Single)
    pInterest.Z3 = RHS
End Property

Public Property Get Y3() As Single
    Y3 = pInterest.Z3
End Property

Public Property Let Y3(ByVal RHS As Single)
    pInterest.Z3 = RHS
End Property

Public Property Get Z3() As Single
    Z3 = pInterest.Z3
End Property

Public Property Let Z3(ByVal RHS As Single)
    pInterest.Z3 = RHS
End Property

Public Property Get X4() As Single
    X4 = pWaypoint.X
End Property

Public Property Let X4(ByVal RHS As Single)
    pWaypoint.X = RHS
End Property

Public Property Get Y4() As Single
    Y4 = pWaypoint.Y
End Property

Public Property Let Y4(ByVal RHS As Single)
    pWaypoint.Y = RHS
End Property

Public Property Get Z4() As Single
    Z4 = pWaypoint.Z
End Property

Public Property Let Z4(ByVal RHS As Single)
    pWaypoint.Z = RHS
End Property

Friend Function Clone() As Shape
    Set Clone = New Shape
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "#" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    pInterest.Serialize = NextArg(RHS, "#")
    pWaypoint.Serialize = RemoveArg(RHS, "#")
End Property

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub


