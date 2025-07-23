VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pIdentity As String
Private pAction As Long
Private pData As Point
Private pOnEvent As String
Private pReactive As Single
Private pLatency As Single
Private pRecount As Single
Private pEmphasis As Single
Private pFriction As Single
Private pInitials As Single

Public Property Get Identity() As String
    Identity = pIdentity
End Property
Public Property Let Identity(ByRef RHS As String)
    pIdentity = RHS
End Property

Public Property Get Action() As Long
    Action = pAction
End Property
Public Property Let Action(ByRef RHS As Long)
    pAction = RHS
End Property

Public Property Get Data() As Point
    Set Data = pData
End Property
Public Property Set Data(ByRef RHS As Point)
    Set pData = RHS
End Property

Public Property Get OnEvent() As String
    OnEvent = pOnEvent
End Property
Public Property Let OnEvent(ByRef RHS As String)
    pOnEvent = RHS
End Property

Public Property Get Reactive() As Single
    Reactive = pReactive
End Property
Public Property Let Reactive(ByRef RHS As Single)
    pReactive = RHS
End Property

Public Property Get Latency() As Single
    Latency = pLatency
End Property
Public Property Let Latency(ByRef RHS As Single)
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
End Property

Public Property Get Friction() As Single
    Friction = pFriction
End Property
Public Property Let Friction(ByRef RHS As Single)
    pFriction = RHS
End Property

Public Property Get Initials() As Single
    Initials = pInitials
End Property
Public Property Let Initials(ByRef RHS As Single)
    pInitials = RHS
End Property

Private Sub Class_Initialize()
    Stats_Activity_Count = Stats_Activity_Count + 1
End Sub

Private Sub Class_Terminate()
    Stats_Activity_Count = Stats_Activity_Count - 1
End Sub
