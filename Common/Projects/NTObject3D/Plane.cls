VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'no different then a (a,b,c,d) plane
'nor a (x,y,z,r) where as r is radius
Private pInterest As Square
Private pWaypoint As Square
Friend Property Get Interest() As Square
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Square)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Square
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Square)
    Set pWaypoint = RHS
End Property

Public Property Get A() As Single
    A = pInterest.Interest.n
End Property

Public Property Let A(ByVal RHS As Single)
     pInterest.Interest.n = RHS
End Property

Public Property Get B() As Single
    B = pInterest.Waypoint.n
End Property

Public Property Let B(ByVal RHS As Single)
     pInterest.Waypoint.n = RHS
End Property

Public Property Get C() As Single
    C = pWaypoint.Interest.n
End Property

Public Property Let C(ByVal RHS As Single)
     pWaypoint.Interest.n = RHS
End Property

Public Property Get D() As Single
    D = pWaypoint.Waypoint.n
End Property

Public Property Let D(ByVal RHS As Single)
     pWaypoint.Waypoint.n = RHS
End Property

Friend Function Clone() As Plane
    Set Clone = New Plane
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Friend Property Get Serial() As String
    Serial = pInterest.Serial() & "?" & pWaypoint.Serial()
End Property
Friend Property Let Serial(ByVal RHS As String)
    pInterest.Serial = NextArg(RHS, "?")
    pWaypoint.Serial = RemoveArg(RHS, "?")
End Property


Private Sub Class_Initialize()
    Set pInterest = New Point
    Set pWaypoint = New Atom
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

