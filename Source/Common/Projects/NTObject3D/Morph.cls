VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Morph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'defining fector that the transit will affect
'when applied to real space and axis invovled
'giving a tranit, a potential energy, a actual
'kenetic energy in accordance to 3d Vector
Private pInterest As Range
Private pWaypoint As Point
Friend Property Get Interest() As Range
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Range)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Point
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property

Public Property Get Reactive() As Single
    Reactive = pInterest.X
End Property

Public Property Let Reactive(ByVal RHS As Single)
     pInterest.X = RHS
End Property

Public Property Get Recount() As Single
    Recount = pInterest.Y
End Property

Public Property Let Recount(ByVal RHS As Single)
     pInterest.Y = RHS
End Property

Public Property Get Emphasis() As Single
    Emphasis = pInterest.Z
End Property

Public Property Let Emphasis(ByVal RHS As Single)
     pInterest.Z = RHS
End Property

Public Property Get Friction() As Single
    Friction = pInterest.R
End Property

Public Property Let Friction(ByVal RHS As Single)
     pInterest.R = RHS
End Property

Public Property Get X() As Single
    X = pWaypoint.X
End Property

Public Property Let X(ByVal RHS As Single)
    pWaypoint.X = RHS
End Property

Public Property Get Y() As Single
    Y = pWaypoint.Y
End Property

Public Property Let Y(ByVal RHS As Single)
    pWaypoint.Y = RHS
End Property

Public Property Get Z() As Single
    Z = pWaypoint.Z
End Property

Public Property Let Z(ByVal RHS As Single)
    pWaypoint.Z = RHS
End Property

Friend Function Clone() As Morph
    Set Clone = New Morph
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "~" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    pInterest.Serialize = NextArg(RHS, "~")
    pWaypoint.Serialize = RemoveArg(RHS, "~")
End Property

Private Sub Class_Initialize()
    Set pInterest = New Range
    Set pWaypoint = New Point
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

