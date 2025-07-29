VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Actions
    None = 0
    Directing = 1
    Rotating = 2
    Scaling = 4
    Method = 8
End Enum

Private pIdentity As String
Private pAction As Long
Private pData As Point
Private pScript As String
Private pReactive As Single
Private pLatency As Single
Private pRecount As Single
Private pEmphasis As Single
Private pFriction As Single
Private pInitials As Single
Public Property Get Key() As String
    Key = pIdentity
End Property
    
Public Property Let Key(ByVal RHS As String)
    pIdentity = RHS
End Property

Public Property Get Identity() As String
    Identity = pIdentity
End Property
Public Property Let Identity(ByRef RHS As String)
    pIdentity = RHS
End Property

Public Property Get Action() As Actions
    Action = pAction
End Property
Public Property Let Action(ByVal RHS As Actions)
    pAction = RHS
End Property

Public Property Get Data() As Point
    Set Data = pData
End Property
Public Property Set Data(ByRef RHS As Point)
    Set pData = RHS
End Property
Public Property Let Data(ByVal RHS)
    Set pData = New Point
    pData = RHS
End Property

Public Property Get Script() As String
Attribute Script.VB_MemberFlags = "40"
    Script = pScript
End Property
Public Property Let Script(ByVal RHS As String)
    pScript = RHS
End Property

Public Property Get Reactive() As Single
    Reactive = pReactive
End Property
Public Property Let Reactive(ByRef RHS As Single)
    pReactive = RHS
End Property

Friend Property Get latency() As Single
    latency = pLatency
End Property
Friend Property Let latency(ByRef RHS As Single)
    pLatency = RHS
End Property

Public Property Get Recount() As Single
    Recount = pRecount
End Property
Public Property Let Recount(ByRef RHS As Single)
    pRecount = RHS
End Property

Public Property Get Emphasis() As Single
    Emphasis = pEmphasis
End Property
Public Property Let Emphasis(ByRef RHS As Single)
    pEmphasis = RHS
    If pInitials = 0 Then pInitials = RHS
End Property

Public Property Get Friction() As Single
    Friction = pFriction
End Property
Public Property Let Friction(ByRef RHS As Single)
    pFriction = RHS
End Property

Friend Property Get Initials() As Single
    Initials = pInitials
End Property
Friend Property Let Initials(ByRef RHS As Single)
    pInitials = RHS
End Property

Private Sub Class_Initialize()
    Stats_Motion_Count = Stats_Motion_Count + 1
    pIdentity = "K" & ObjPtr(Me)
    pReactive = -1
    pRecount = -1
End Sub

Private Sub Class_Terminate()
    Stats_Motion_Count = Stats_Motion_Count - 1
End Sub
