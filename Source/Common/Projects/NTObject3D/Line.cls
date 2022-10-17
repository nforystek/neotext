VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'two 3d points that makes
'a line in 3d geomentry
Private pInterest As Point
Private pWaypoint As Point
Friend Property Get Interest() As Point
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Point)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Point
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property


Public Property Get X1() As Single
    X1 = pInterest.X
End Property

Public Property Let X1(ByVal RHS As Single)
    pInterest.X = RHS
End Property

Public Property Get Y1() As Single
    Y1 = pInterest.Y
End Property

Public Property Let Y1(ByVal RHS As Single)
    pInterest.Y = RHS
End Property

Public Property Get Z1() As Single
    Z1 = pInterest.Z
End Property

Public Property Let Z1(ByVal RHS As Single)
    pInterest.Z = RHS
End Property


Public Property Get X2() As Single
    X2 = pWaypoint.X
End Property

Public Property Let X2(ByVal RHS As Single)
    pWaypoint.X = RHS
End Property

Public Property Get Y2() As Single
    Y2 = pWaypoint.Y
End Property

Public Property Let Y2(ByVal RHS As Single)
    pWaypoint.Y = RHS
End Property

Public Property Get Z2() As Single
    Z2 = pWaypoint.Z
End Property

Public Property Let Z2(ByVal RHS As Single)
    pWaypoint.Z = RHS
End Property

Friend Function Clone() As Line
    Set Clone = New Line
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "|" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    pInterest.Serialize = NextArg(RHS, "|")
    pWaypoint.Serialize = RemoveArg(RHS, "|")
End Property


Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub
