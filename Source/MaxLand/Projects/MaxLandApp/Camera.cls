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

'Private pAngle As Single
'Private pPitch As Single
'Private pRoll As Single
'Private pZoom As Single
Private pMinZoom As Single
Private pMaxZoom As Single

Private pKey As String
Private pSerialize As Boolean


Private pElement As Element

Friend Sub LateEval()

End Sub

Public Property Get Origin() As Point
    If pElement Is Nothing Then
        Set Origin = Orient.Origin
    Else
        Set Origin = pElement.Origin
    End If
End Property
Public Property Set Origin(ByRef RHS As Point)
    If pElement Is Nothing Then
        Set Orient.Origin = RHS
    Else
        Set pElement.Origin = RHS
    End If
End Property

Public Property Get Offset() As Point
    'If pElement Is Nothing Then
        Set Offset = Orient.Offset
    'Else
    '    Set Offset = pElement.Offset
    'End If
End Property
Public Property Set Offset(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Offset = RHS
    'Else
    '    Set pElement.Offset = RHS
    'End If
End Property

Public Property Get Scaled() As Point
    'If pElement Is Nothing Then
        Set Scaled = Orient.Scaled
    'Else
    '    Set Scaled = pElement.Scaled
    'End If
End Property
Public Property Set Scaled(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Scaled = RHS
    'Else
    '    Set pElement.Scaled = RHS
    'End If
End Property

Public Property Get Rotate() As Point
    'If pElement Is Nothing Then
        Set Rotate = Orient.Rotate
    'Else
    '    Set Rotate = pElement.Rotate
    'End If
End Property
Public Property Set Rotate(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Rotate = RHS
    'Else
    '    Set pElement.Rotate = RHS
    'End If
End Property


Public Property Get Direct() As Point
    'If pElement Is Nothing Then
        Set Direct = Orient.Direct
    'Else
    '    Set Direct = pElement.Direct
    'End If
End Property
Public Property Set Direct(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Direct = RHS
    'Else
    '    Set pElement.Direct = RHS
    'End If
End Property

Public Property Get Twists() As Point
    'If pElement Is Nothing Then
        Set Twists = Orient.Twists
    'Else
    '    Set Twists = pElement.Twists
    'End If
End Property
Public Property Set Twists(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Twists = RHS
    'Else
    '    Set pElement.Twists = RHS
    'End If
End Property

Public Property Get Scalar() As Point
    'If pElement Is Nothing Then
        Set Scalar = Orient.Scalar
    'Else
    '    Set Scalar = pElement.Scalar
    'End If
End Property
Public Property Set Scalar(ByRef RHS As Point)
    'If pElement Is Nothing Then
        Set Orient.Scalar = RHS
    'Else
    '    Set pElement.Scalar = RHS
    'End If
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
Public Property Set Element(ByRef RHS As Element)
    SetObject RHS
End Property
Public Property Let Element(ByVal RHS As Element)
    SetObject RHS
End Property

Private Sub SetObject(ByRef Obj)
    If TypeName(Obj) = "Element" Then
        Set pElement = Obj
       ' vbaObjSetAddref Orient, ObjPtr(pElement)
    End If
End Sub
Public Property Set Value(ByVal RHS)
    If IsNumeric(RHS) Then
        Value = RHS
    ElseIf IsObject(RHS) Then
        SetObject RHS
    End If
End Property
Public Property Let Value(ByVal RHS As Long)
    SetObject ObjectByPtr(RHS)
End Property
Public Property Get Value() As Long
    Value = ObjPtr(Me)
End Property

'Public Property Get Player() As Player
'    Set Player = pPlayer
'End Property
'Friend Property Set Player(ByRef RHS As Player)
'    Set pPlayer = RHS
'    If Not RHS Is Nothing Then
'        If Not RHS.Camera Is Nothing Then
'            If Not RHS.Camera.Value = Me.Value Then Set RHS.Camera = Me
'        Else
'            Set RHS.Camera = Me
'        End If
'    End If
'End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

'Public Property Get Angle() As Single
'    Angle = Orient.Rotate.Y
'End Property
'Public Property Let Angle(ByVal RHS As Single)
'    Orient.Rotate.Y = RHS
'End Property
'
'Public Property Get Pitch() As Single
'    Pitch = Orient.Rotate.X
'End Property
'Public Property Let Pitch(ByVal RHS As Single)
'    Orient.Rotate.X = RHS
'End Property
'
'Public Property Get Roll() As Single
'    Roll = Orient.Rotate.Z
'End Property
'Public Property Let Roll(ByVal RHS As Single)
'    Orient.Rotate.Z = RHS
'End Property
'
'Public Property Get Zoom() As Single
'    Zoom = Orient.Offset.Z
'End Property
'Public Property Let Zoom(ByVal RHS As Single)
'    Orient.Offset.Z = RHS
'End Property

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


'Public Sub ZoomIn(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    If Perspective = ThirdPerson Then
'        Orient.Offset.Z = Orient.Offset.Z - Measure
'        If Orient.Offset.Z < Orient.MinZoom Then Orient.Offset.Z = Orient.MinZoom
'    End If
'
'End Sub

'Public Sub ZoomOut(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    If Perspective = ThirdPerson Then
'        Orient.Offset.Z = Orient.Offset.Z + Measure
'        If Orient.Offset.Z > Orient.MaxZoom Then Orient.Offset.Z = Orient.MaxZoom
'    End If
'
'End Sub
'
'Public Sub ZoomIn(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    If Perspective = ThirdPerson Then
'        Orient.Offset.Z = Orient.Offset.Z - Measure
'        If Orient.Offset.Z < Orient.MinZoom Then Orient.Offset.Z = Orient.MinZoom
'    End If
'
'End Sub


Public Sub Zoom(Optional ByVal Measure As Single = 0.015)
    ResetIdle
   'Debug.Print "Zoom " & Measure
    
    If Perspective = ThirdPerson Then
        Orient.Offset.Z = Orient.Offset.Z - Measure
        If Orient.Offset.Z < pMinZoom Then Orient.Offset.Z = pMinZoom
        If Orient.Offset.Z > pMaxZoom Then Orient.Offset.Z = pMaxZoom
        

    End If
    
End Sub

'Public Sub LookLeft(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    If Not Perspective = CameraMode Then Orient.Rotate.Y = (Orient.Rotate.Y + Measure)
'    If Orient.Rotate.Y < -(PI * 2) Then Orient.Rotate.Y = (Orient.Rotate.Y + (PI * 2))
'
'End Sub
'Public Sub LookRight(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    If Not Perspective = CameraMode Then Orient.Rotate.Y = (Orient.Rotate.Y - Measure)
'    If Orient.Rotate.Y > (PI * 2) Then Orient.Rotate.Y = (Orient.Rotate.Y - (PI * 2))
'
'End Sub

Public Sub Angle(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    'Debug.Print "Angle " & Measure
    If Not Perspective = CameraMode Then Orient.Rotate.Y = (Orient.Rotate.Y - Measure)
    Orient.Rotate.Y = AngleRestrict(Orient.Rotate.Y)
    'If Orient.Rotate.Y > (PI * 2) Then Orient.Rotate.Y = (Orient.Rotate.Y - (PI * 2))
   
End Sub

'Public Sub LookUp(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    Orient.Rotate.X = Orient.Rotate.X - Measure
'    If Orient.Rotate.X < -1.5 Then Orient.Rotate.X = -1.5
'
'End Sub
'
'Public Sub LookDown(Optional ByVal Measure As Single = 0.015)
'    ResetIdle
'
'    Orient.Rotate.X = Orient.Rotate.X + Measure
'    If Orient.Rotate.X > 1.5 Then Orient.Rotate.X = 1.5
'
'End Sub

Public Sub Pitch(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    'Debug.Print "Pitch " & Measure
    Orient.Rotate.X = Orient.Rotate.X - Measure
    If Orient.Rotate.X > 1.5 Then Orient.Rotate.X = 1.5
    If Orient.Rotate.X < -1.5 Then Orient.Rotate.X = -1.5

End Sub



Public Property Get ToString() As String
    ToString = "[" & Orient.Origin.X & ", " & Orient.Origin.Y & ", " & Orient.Origin.Z & ", " & Orient.Rotate.Y & ", " & Orient.Rotate.X & ", " & Orient.Rotate.Z & ", " & Orient.Offset.Z & ", " & pMinZoom & ", " & pMaxZoom & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        Orient.Origin.X = ParseNumerical(RHS)
        Orient.Origin.Y = ParseNumerical(RHS)
        Orient.Origin.Z = ParseNumerical(RHS)

        Orient.Rotate.Y = ParseNumerical(RHS)
        Orient.Rotate.X = ParseNumerical(RHS)
        Orient.Rotate.Z = ParseNumerical(RHS)
   
        Orient.Offset.Z = ParseNumerical(RHS)
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
    Orient.Offset.Z = 5
    Orient.Rotate.Y = 0
    Orient.Rotate.X = 0
    
End Sub

Private Sub Class_Terminate()

'    Set pPlayer = Nothing
    
   ' Set pElement = Nothing

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
    Set Orient.Rotate = RHS
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

Private Property Set Orient_Direct(RHS As Point)
    Set Orient.Direct = RHS
End Property

Private Property Get Orient_Direct() As Point
    Set Orient_Direct = Orient.Direct
End Property

Private Property Set Orient_Scalar(RHS As Point)
    Set Orient.Scalar = RHS
End Property

Private Property Get Orient_Scalar() As Point
    Set Orient_Scalar = Orient.Scalar
End Property

Private Property Set Orient_Twists(RHS As Point)
    Set Orient.Twists = RHS
End Property

Private Property Get Orient_Twists() As Point
    Set Orient_Twists = Orient.Twists
End Property
