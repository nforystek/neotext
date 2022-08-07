VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Matter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'the main building block for
'triangles that make up mesh
'holds a 3 point triangle, a
'index for each, texture UV
'data for each and normal
Private pInterest As Shape
Private pWaypoint As Shape

Friend Property Get Interest() As Shape
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Shape)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Shape
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Shape)
    Set pWaypoint = RHS
End Property

Friend Property Get Shape1() As Shape
    Set Shape1 = pInterest
End Property
Friend Property Set Shape1(ByRef RHS As Shape)
    Set pInterest = RHS
End Property

Friend Property Get Shape2() As Shape
    Set Shape2 = pWaypoint
End Property
Friend Property Set Shape2(ByRef RHS As Shape)
    Set pWaypoint = RHS
End Property

Public Property Get Point1() As Point
    Set Point1 = pInterest.Interest.Interest.Interest
End Property
Friend Property Set Point1(ByRef RHS As Point)
    Set pInterest.Interest.Interest.Interest = RHS
End Property

Public Property Get Point2() As Point
    Set Point2 = pInterest.Interest.Interest.Waypoint
End Property
Friend Property Set Point2(ByRef RHS As Point)
    Set pInterest.Interest.Interest.Waypoint = RHS
End Property

Public Property Get Point3() As Point
    Set Point3 = pInterest.Interest.Waypoint
End Property
Friend Property Set Point3(ByRef RHS As Point)
    Set pInterest.Interest.Waypoint = RHS
End Property

Public Property Get Normal() As Point
    Set Normal = pInterest.Waypoint
End Property
Friend Property Set Normal(ByRef RHS As Point)
    Set pInterest.Waypoint = RHS
End Property

Public Property Get Index1() As Long
    Index1 = pWaypoint.X1
End Property
Friend Property Let Index1(ByRef RHS As Long)
    pWaypoint.X1 = RHS
End Property

Public Property Get Index2() As Long
    Index2 = pWaypoint.Y1
End Property
Friend Property Let Index2(ByRef RHS As Long)
    pWaypoint.Y1 = RHS
End Property

Public Property Get Index3() As Long
    Index3 = pWaypoint.Z1
End Property
Friend Property Let Index3(ByRef RHS As Long)
    pWaypoint.Z1 = RHS
End Property

Public Property Get U1() As Long
    U1 = pWaypoint.X2
End Property
Friend Property Let U1(ByRef RHS As Long)
    pWaypoint.X2 = RHS
End Property

Public Property Get U2() As Long
    U2 = pWaypoint.X3
End Property
Friend Property Let U2(ByRef RHS As Long)
    pWaypoint.X3 = RHS
End Property

Public Property Get U3() As Long
    U3 = pWaypoint.X4
End Property
Friend Property Let U3(ByRef RHS As Long)
    pWaypoint.X4 = RHS
End Property

Public Property Get v1() As Long
    v1 = pWaypoint.Y2
End Property
Friend Property Let v1(ByRef RHS As Long)
    pWaypoint.Y2 = RHS
End Property

Public Property Get v2() As Long
    v2 = pWaypoint.Y3
End Property
Friend Property Let v2(ByRef RHS As Long)
    pWaypoint.Y3 = RHS
End Property

Public Property Get V3() As Long
    V3 = pWaypoint.Y4
End Property
Friend Property Let V3(ByRef RHS As Long)
    pWaypoint.Y4 = RHS
End Property

Public Property Get NormalIndex() As Long
    NormalIndex = pWaypoint.Z1
End Property
Public Property Let NormalIndex(ByRef RHS As Long)
    pWaypoint.Z1 = RHS
End Property

Public Property Get TriangleIndex() As Long
    TriangleIndex = pWaypoint.Z2
End Property
Public Property Let TriangleIndex(ByRef RHS As Long)
    pWaypoint.Z2 = RHS
End Property

Public Property Get ObjectIndex() As Long
    ObjectIndex = pWaypoint.Z3
End Property
Public Property Let ObjectIndex(ByRef RHS As Long)
    pWaypoint.Z3 = RHS
End Property

Public Property Get FaceIndex() As Long
    FaceIndex = pWaypoint.Z4
End Property
Public Property Let FaceIndex(ByRef RHS As Long)
    pWaypoint.Z4 = RHS
End Property

Friend Function Clone() As Matter
    Set Clone = New Matter
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "=" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    pInterest.Serialize = NextArg(RHS, "=")
    pWaypoint.Serialize = RemoveArg(RHS, "=")
End Property

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

