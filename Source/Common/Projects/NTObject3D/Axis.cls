VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Axis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'joins a orbit to a Plane giving
'it a location within a 3D space
Private pInterest As Range
Private pWaypoint As Orbit
Friend Property Get Interest() As Range
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Range)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Orbit
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Orbit)
    Set pWaypoint = RHS
End Property

Public Property Get Range() As Range
    Set Range = pInterest
End Property
Public Property Set Range(ByRef RHS As Range)
    Set pInterest = RHS
End Property

Public Property Get Orbit() As Orbit
    Set Orbit = pWaypoint
End Property
Public Property Set Orbit(ByRef RHS As Orbit)
    Set pWaypoint = RHS
End Property

Friend Function Clone() As Axis
    Set Clone = New Axis
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And (pWaypoint = SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "!" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    pInterest.Serialize = NextArg(RHS, "!")
    pWaypoint.Serialize = RemoveArg(RHS, "!")
End Property

Private Sub Class_Initialize()
    Set pInterest = New Range
    'pInterest.R = -1
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

