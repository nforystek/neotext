VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'no different then a (a,b,c,d) in size
'but this uses r for radius, therefore
'(x,y,z )likely use as a location then
'then a r defined radius maybe like Z-
'order distance in rendering, gravity,
'windows handle for distance to plane
Private pInterest As Point
Private pWaypoint As Single
Friend Property Get Interest() As Point
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Point)
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
    Z = pInterest.Z
End Property
Public Property Let Z(ByVal RHS As Single)
    pInterest.Z = RHS
End Property

Public Property Get R() As Single
    R = pWaypoint
End Property
Public Property Let R(ByVal RHS As Single)
    pWaypoint = RHS
End Property

Friend Function Clone() As Range
    Set Clone = New Range
    Set Clone.Interest = pInterest.Clone
    Clone.Waypoint = pWaypoint
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And (pWaypoint = SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "?" & pWaypoint
End Property
Public Property Let Serialize(ByVal RHS As String)
    If (pInterest Is Nothing) Then Set pInterest = New Point
    pInterest.Serialize = NextArg(RHS, "?")
    pWaypoint = CSng(RemoveArg(RHS, "?"))
End Property

Private Sub Class_Initialize()
    Set pInterest = New Point
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
End Sub


