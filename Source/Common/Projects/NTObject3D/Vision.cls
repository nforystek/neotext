VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'one angle, plus a single decimal
'that is used as the render raerget

Private pInterest As Angle
Private pWaypoint As Single
Friend Property Get Interest() As Angle
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Angle)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Single
    Waypoint = pWaypoint
End Property
Friend Property Let Waypoint(ByRef RHS As Single)
    pWaypoint = RHS
End Property

Public Property Get Eye() As Point
    Set Eye = pInterest.Interest.Interest
End Property
Public Property Set Eye(ByRef RHS As Point)
    Set pInterest.Interest.Interest = RHS
End Property

Public Property Get Up() As Point
    Set Up = pInterest.Interest.Waypoint
End Property
Public Property Set Up(ByRef RHS As Point)
    Set pInterest.Interest.Waypoint = RHS
End Property

Public Property Get At() As Point
    Set At = pInterest.Waypoint
End Property
Public Property Set At(ByRef RHS As Point)
    Set pInterest.Waypoint = RHS
End Property

Public Property Get Dtp() As Single
    Dtp = pWaypoint
End Property

Public Property Let Dtp(ByVal RHS As Single)
    pWaypoint = RHS
End Property

Friend Function Clone() As Vision
    Set Clone = New Vision
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
    pInterest.Serialize = NextArg(RHS, "?")
    pWaypoint = CSng(RemoveArg(RHS, "?"))
End Property

Private Sub Class_Initialize()
    Set pInterest = New Angle
    Set pInterest.Interest = New Line
    Set pInterest.Interest.Interest = New Point
    Set pInterest.Interest.Waypoint = New Point
    
    Set pInterest.Waypoint = New Point
   ' pInterest.Waypoint.Z = -4
    
    pInterest.Interest.Interest.Z = 1 ' FAR \ 2
    pInterest.Interest.Waypoint.Y = 1
End Sub

