VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private pEnabled As Boolean
Private pOnInRange As OnEvent
Private pOnOutRange As OnEvent

Private pRange As Single
Private pLocation As Point
Private pTeleport As Point

Public Motions As NTNodes10.Collection

Private pDropsMotions As Boolean

Private pKey As String
Private pSerialize As Boolean

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal RHS As Boolean)
    pEnabled = RHS
End Property

Public Property Get Range() As Single
    Range = pRange
End Property
Public Property Let Range(ByVal RHS As Single)
    pRange = RHS
End Property

Friend Property Get OnInRange() As OnEvent
    Set OnInRange = pOnInRange
End Property
Friend Property Set OnInRange(ByRef RHS As OnEvent)
    Set pOnInRange = RHS
End Property


Friend Property Get OnOutRange() As OnEvent
    Set OnOutRange = pOnOutRange
End Property
Friend Property Set OnOutRange(ByRef RHS As OnEvent)
    Set pOnOutRange = RHS
End Property


Public Property Get Location() As Point
    Set Location = pLocation
End Property
Public Property Set Location(ByRef RHS As Point)
    Set pLocation = RHS
End Property


Public Property Get Teleport() As Point
    Set Teleport = pTeleport
End Property
Public Property Set Teleport(ByRef RHS As Point)
    Set pTeleport = RHS
End Property


Public Function AddMotion(ByRef Action As Long, ByVal aGUID As String, ByRef Data As Point, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String
    Dim act As Motion
    Set act = New Motion
    With act
        .Identity = IIf(aGUID = "", Replace(modGuid.GUID, "-", "K"), aGUID)
        .Action = Action
        Set .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .latency = Timer
        .Recount = Recount
        .Script = Script
        AddMotion = .Identity
    End With
    If Motions Is Nothing Then
        Set Motions = New NTNodes10.Collection
    End If
    Motions.Add act
End Function

Public Function DeleteMotion(ByVal MGUID As String) As Boolean
    Dim A As Long
    Dim act As Motion
    If Not Motions Is Nothing Then
        A = 1
        Do While A <= Motions.Count
            If Motions(A).Identity = MGUID Or (Motions(A).Identity = "") Then
                Set act = Motions(A)
                Motions.Remove A
                Set act = Nothing
                DeleteMotion = True
            Else
                A = A + 1
            End If
        Loop
    End If
End Function

Public Sub ClearMotions()
    If Not Motions Is Nothing Then
        Dim act As Motion
        Do While Motions.Count > 0
            Set act = Motions(1)
            Motions.Remove 1
            Set act = Nothing
        Loop
    End If
End Sub

Public Property Get DropsMotions() As Boolean
    DropsMotions = pDropsMotions
End Property
Public Property Let DropsMotions(ByVal RHS As Boolean)
    pDropsMotions = RHS
End Property


Private Sub Class_Initialize()
    Stats_Portal_Count = Stats_Portal_Count + 1
    Set pLocation = New Point
    Set pTeleport = New Point
    pEnabled = True
    
    
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    Stats_Portal_Count = Stats_Portal_Count - 1
    Set pLocation = Nothing
    Set pTeleport = Nothing

    ClearMotions
    Set Motions = Nothing
End Sub
