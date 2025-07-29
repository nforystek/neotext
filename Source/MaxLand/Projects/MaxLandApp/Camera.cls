VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pOrigin As Point

Private pAngle As Single
Private pPitch As Single
Private pRoll As Single
Private pZoom As Single
Private pMinZoom As Single
Private pMaxZoom As Single

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

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Angle() As Single
    Angle = pAngle
End Property
Public Property Let Angle(ByVal RHS As Single)
    pAngle = RHS
End Property

Public Property Get Pitch() As Single
    Pitch = pPitch
End Property
Public Property Let Pitch(ByVal RHS As Single)
    pPitch = RHS
End Property

Public Property Get Roll() As Single
    Roll = pRoll
End Property
Public Property Let Roll(ByVal RHS As Single)
    pRoll = RHS
End Property

Public Property Get Zoom() As Single
    Zoom = pZoom
End Property
Public Property Let Zoom(ByVal RHS As Single)
    pZoom = RHS
End Property

Public Property Get ToString() As String
    ToString = "[" & pOrigin.X & ", " & pOrigin.Y & ", " & pOrigin.Z & ", " & pAngle & ", " & pPitch & ", " & pRoll & ", " & pZoom & ", " & pMinZoom & ", " & pMaxZoom & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        pOrigin.X = ParseNumerical(RHS)
        pOrigin.Y = ParseNumerical(RHS)
        pOrigin.Z = ParseNumerical(RHS)

        pAngle = ParseNumerical(RHS)
        pPitch = ParseNumerical(RHS)
        pRoll = ParseNumerical(RHS)
   
        pZoom = ParseNumerical(RHS)
        pMinZoom = ParseNumerical(RHS)
        pMaxZoom = ParseNumerical(RHS)
    End If
End Property

Private Sub Class_Initialize()
    Stats_Camera_Count = Stats_Camera_Count + 1
    Set pOrigin = New Point
    Key = "K" & ObjPtr(Me)
    pMaxZoom = 10
    pMinZoom = 1
    pZoom = 5
End Sub

Private Sub Class_Terminate()
    Set pOrigin = Nothing
    Stats_Camera_Count = Stats_Camera_Count - 1
End Sub

