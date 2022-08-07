VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spirit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'mart of the stars collection these
'are used for anything of a point
'information that maybe applied to
'an axis, such as rotate/location
Private pInterest As Axis
Private pWaypoint As Point
Friend Property Get Interest() As Axis
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Axis)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Point
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property

Public Property Get Axis() As Axis
    Set Axis = pInterest
End Property
Public Property Set Axis(ByRef RHS As Axis)
    Set pInterest = RHS
End Property

Public Property Get Point() As Point
    Set Point = pWaypoint
End Property
Public Property Set Point(ByRef RHS As Point)
    Set pWaypoint = RHS
End Property

Friend Function Clone() As Spirit
    Set Clone = New Spirit
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    If pInterest Is Nothing Then
        Serialize = ";" & pWaypoint.Serialize()
    Else
        Serialize = pInterest.Serialize() & ";" & pWaypoint.Serialize()
    End If
End Property
Public Property Let Serialize(ByVal RHS As String)
    If Not (pInterest Is Nothing) Then pInterest.Serialize = NextArg(RHS, ";")
    pWaypoint.Serialize = RemoveArg(RHS, ";")
End Property

Private Sub Class_Initialize()
    Set pWaypoint = New Point
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

