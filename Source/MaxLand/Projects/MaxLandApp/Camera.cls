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

Implements Orient

Private pAngle As Single
Private pPitch As Single
Private pRoll As Single
Private pZoom As Single
Private pMinZoom As Single
Private pMaxZoom As Single

Private pKey As String
Private pSerialize As Boolean

Private pPlayer As Player
Private pElement As Element

Public Property Get Origin() As Point
    Set Origin = Orient.Origin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set Orient.Origin = RHS
End Property

Public Property Get Offset() As Point
    Set Offset = Orient.Offset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set Orient.Offset = RHS
End Property

Public Property Get Scaled() As Point
    Set Scaled = Orient.Scaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set Orient.Scaled = RHS
End Property

Public Property Get Rotate() As Point
    Set Rotate = Orient.Rotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set Orient.Rotate = RHS
End Property


Friend Property Get Direct() As Point
    Set Direct = Orient.Direct
End Property
Friend Property Set Direct(ByRef RHS As Point)
    Set Orient.Direct = RHS
End Property

Friend Property Get Twists() As Point
    Set Twists = Orient.Twists
End Property
Friend Property Set Twists(ByRef RHS As Point)
    Set Orient.Twists = RHS
End Property

Friend Property Get Scalar() As Point
    Set Scalar = Orient.Scalar
End Property
Friend Property Set Scalar(ByRef RHS As Point)
    Set Orient.Scalar = RHS
End Property

Public Property Get Attachments() As NTNodes10.Collection
    Set Attachments = Orient.Attachments
End Property
Public Property Set Attachments(ByRef RHS As NTNodes10.Collection)
    Set Orient.Attachments = RHS
End Property

Public Property Get AttachedTo() As String
    AttachedTo = Orient.AttachedTo
End Property
Public Property Let AttachedTo(ByVal RHS As String)
    Orient.AttachedTo = RHS
End Property

Public Property Let Key(ByVal RHS As String)
    Orient.Key = RHS
End Property

Public Property Get Key() As String
    Key = Orient.Key
End Property

Public Property Get Element() As Element
    Set Element = pElement
End Property
Friend Property Set Element(ByRef RHS As Element)
    Set pElement = RHS
    If Not RHS.Camera.Value = Me.Value Then Set RHS.Camera = Me
End Property

Public Property Get Player() As Player
    Set Player = pPlayer
End Property
Friend Property Set Player(ByRef RHS As Player)
    Set pPlayer = RHS
    If Not RHS.Camera.Value = Me.Value Then Set RHS.Camera = Me
End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
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

Public Property Get MinZoom() As Single
    MinZoom = pMinZoom
End Property
Public Property Let MinZoom(ByVal RHS As Single)
    pMinZoom = RHS
End Property

Public Property Get MaxZoom() As Single
    MaxZoom = pMaxZoom
End Property
Public Property Let MaxZoom(ByVal RHS As Single)
    pMaxZoom = RHS
End Property



Public Property Set Value(ByVal RHS)
    Dim obj As Object
    Set obj = ObjectByPtr(RHS)
    If TypeName(obj) = "Element" Then
        Set pElement = obj
    ElseIf TypeName(obj) = "Player" Then
        Set pPlayer = obj
    End If
    If obj.Camera Is Nothing Then
        Set RHS.Camera = Me
    ElseIf Not obj.Camera.Value = Me.Value Then
        Set RHS.Camera = Me
    End If
End Property

Public Property Let Value(ByVal RHS As Long)
    Value = RHS
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "40"
    Value = ObjPtr(Me)
End Property

Public Property Get ToString() As String
    ToString = "[" & Orient.Origin.X & ", " & Orient.Origin.Y & ", " & Orient.Origin.Z & ", " & pAngle & ", " & pPitch & ", " & pRoll & ", " & pZoom & ", " & pMinZoom & ", " & pMaxZoom & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        Orient.Origin.X = ParseNumerical(RHS)
        Orient.Origin.Y = ParseNumerical(RHS)
        Orient.Origin.Z = ParseNumerical(RHS)

        pAngle = ParseNumerical(RHS)
        pPitch = ParseNumerical(RHS)
        pRoll = ParseNumerical(RHS)
   
        pZoom = ParseNumerical(RHS)
        pMinZoom = ParseNumerical(RHS)
        pMaxZoom = ParseNumerical(RHS)
    End If
End Property

Private Sub Class_Initialize()
    vbaObjSetAddref Orient, ObjPtr(New Orient)
    
    Stats_Camera_Count = Stats_Camera_Count + 1

    Key = "K" & ObjPtr(Me)

    pMaxZoom = 10
    pMinZoom = 1
    pZoom = 5
    pAngle = 0
    pPitch = 0
    
End Sub

Private Sub Class_Terminate()

    Set pPlayer = Nothing
    
    Set pElement = Nothing

    vbaObjSet Orient, ObjPtr(Nothing)
    Stats_Camera_Count = Stats_Camera_Count - 1
End Sub


Private Property Let Orient_AttachedTo(ByVal RHS As String)
    Orient.AttachedTo = RHS
End Property

Private Property Get Orient_AttachedTo() As String
    Orient_AttachedTo = Orient.AttachedTo
End Property

Private Property Set Orient_Attachments(RHS As NTNodes10.Collection)
    Set Orient.Attachments = RHS
End Property

Private Property Get Orient_Attachments() As NTNodes10.Collection
    Set Orient_Attachments = Orient.Attachments
End Property

Private Property Get Orient_Clone() As Orient
    Set Orient_Clone = Orient.Clone
End Property

Private Property Let Orient_Key(ByVal RHS As String)
    Orient.Key = RHS
End Property

Private Property Get Orient_Key() As String
    Orient_Key = Orient.Key
End Property

Private Property Set Orient_Offset(RHS As Point)
    Set Orient.Offset = RHS
End Property

Private Property Get Orient_Offset() As Point
    Set Orient_Offset = Orient.Offset
End Property

Private Property Set Orient_Origin(RHS As Point)
    Set Orient.Origin = RHS
End Property

Private Property Get Orient_Origin() As Point
    Set Orient_Origin = Orient.Origin
End Property

Private Property Set Orient_Rotate(RHS As Point)
    Set Orient_Rotate = RHS
End Property

Private Property Get Orient_Rotate() As Point
    Set Orient_Rotate = Orient.Rotate
End Property

Private Property Set Orient_Scaled(RHS As Point)
    Set Orient.Scaled = RHS
End Property

Private Property Get Orient_Scaled() As Point
    Set Orient_Scaled = Orient.Scaled
End Property

Private Property Get Orient_Serialize() As Boolean
    Orient_Serialize = Orient.Serialize
End Property
