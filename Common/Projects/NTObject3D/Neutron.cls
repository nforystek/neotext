VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Neutron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'links a Morph to a orbit
'such that the whole orbit
'moves defined in conjuntion
'with every Morph relevant
Private pInterest As Morph
Private pWaypoint As Orbit
Friend Property Get Interest() As Morph
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Morph)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Orbit
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Orbit)
    Set pWaypoint = RHS
End Property

Public Property Get Morph() As Morph
    Set Morph = pInterest
End Property
Public Property Set Morph(ByRef RHS As Morph)
    Set pInterest = RHS
End Property
Public Property Get Orbit() As Orbit
    Set Orbit = pWaypoint
End Property
Public Property Set Orbit(ByRef RHS As Orbit)
    Set pWaypoint = RHS
End Property

Friend Function Clone() As Neutron
    Set Clone = New Neutron
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And (pWaypoint = SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "'" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    If Not (pInterest Is Nothing) Then Set pInterest = New Morph
    pInterest.Serialize = NextArg(RHS, "'")
    If Not (pWaypoint Is Nothing) Then Set pWaypoint = New Orbit
    pWaypoint.Serialize = RemoveArg(RHS, "'")
End Property

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub