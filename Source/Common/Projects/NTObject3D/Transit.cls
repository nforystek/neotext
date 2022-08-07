VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'same logical as a 2D line, (x,y)-(x,y)
'except this is (reactive,recount)-(emphasis,friction)
'where as reactive is the amount in milliseconds that
'it will take for recount, the numeric number of times
'that the transit may commence, -1 for every frame and
'-1 repeat indefinitly.  friction and emphasis are as
'sounds and impressions reactive and recount or vice
'versa, there for having a transferal effect, transit
'fiction is how much hinderance will be applied every
'recount to emphasis which is power duration applied
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

Public Property Get Reactive() As Single
    Reactive = pInterest.Interest.n
End Property

Public Property Let Reactive(ByVal RHS As Single)
     pInterest.Interest.n = RHS
End Property

Public Property Get Recount() As Single
    Recount = pInterest.Waypoint.n
End Property

Public Property Let Recount(ByVal RHS As Single)
     pInterest.Waypoint.n = RHS
End Property

Public Property Get Emphasis() As Single
    Emphasis = pWaypoint.Interest.n
End Property

Public Property Let Emphasis(ByVal RHS As Single)
     pWaypoint.Interest.n = RHS
End Property

Public Property Get Friction() As Single
    Friction = pWaypoint.Waypoint.n
End Property

Public Property Let Friction(ByVal RHS As Single)
     pWaypoint.Waypoint.n = RHS
End Property


Friend Function Clone() As Transit
    Set Clone = New Transit
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = (pInterest.Equals(SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Friend Property Get Serial() As String
    Serial = pInterest.Serial() & "-" & pWaypoint.Serial()
End Property
Friend Property Let Serial(ByVal RHS As String)
    pInterest.Serial = NextArg(RHS, "-")
    pWaypoint.Serial = RemoveArg(RHS, "-")
End Property

Private Sub Class_Initialize()
    Set pInterest = New Square
    Set pWaypoint = New Square
End Sub

Private Sub Class_Terminate()
    Set pInterest = Nothing
    Set pWaypoint = Nothing
End Sub

