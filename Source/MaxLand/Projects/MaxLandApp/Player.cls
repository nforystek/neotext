VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Orient

'Implements Element



'Private pAngle As Single
'Private pPitch As Single
'Private pRoll As Single
'
'Private pZoom As Single
'Private pMinZoom As Single
'Private pMaxZoom As Single


Private pCameraIndex As Long

Private pSpeed As Single

Private pFriction As Single

Private pCamera As Camera
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

Public Property Let Key(ByVal RHS As String)
    Orient.Key = RHS
End Property

Public Property Get Key() As String
    Key = Orient.Key
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

Public Property Get Element() As Element
    Set Element = pElement
End Property
Public Property Set Element(ByRef RHS As Element)
    Set pElement = RHS
    If Not RHS.Player.Value = Me.Value Then Set RHS.Player = Me
End Property

Public Property Get Camera() As Camera
    Set Camera = pCamera
End Property
Public Property Set Camera(ByRef RHS As Camera)
    Set pCamera = RHS
    If Not RHS.Player.Value = Me.Value Then Set RHS.Player = Me
End Property

'Public Property Get Direct() As Point
'    Set Direct = pElement.Direct
'End Property
'Public Property Set Direct(ByRef RHS As Point)
'    Set pElement.Direct = RHS
'End Property
'
'Public Property Get Scalar() As Point
'    Set Scalar = pElement.Scalar
'End Property
'Public Property Set Scalar(ByRef RHS As Point)
'    Set pElement.Scalar = RHS
'End Property
'
'Public Property Get Twists() As Point
'    Set Twists = pElement.Twists
'End Property
'Public Property Set Twists(ByRef RHS As Point)
'    Set pElement.Twists = RHS
'End Property



'Friend Property Set Element_Direct(RHS As Point)
'    pElement.Direct = RHS
'End Property
'
'Friend Property Get Element_Direct() As Point
'    Set Element_Direct = pElement.Direct
'End Property

'Private Property Set Element_Direct(RHS As Point)
'    Set pElement.Direct = RHS
'End Property
'
'Private Property Get Element_Direct() As Point
'    Set Element_Direct = pElement.Direct
'End Property



'Friend Property Set Element_Scalar(RHS As Point)
'    pElement.Scalar = RHS
'End Property
'
'Friend Property Get Element_Scalar() As Point
'    Set Element_Scalar = pElement.Scalar
'End Property

'Private Property Set Element_Scalar(RHS As Point)
'    Set pElement.Scalar = RHS
'End Property
'
'Private Property Get Element_Scalar() As Point
'    Set Element_Scalar = pElement.Scalar
'End Property



'Friend Property Set Element_Twists(RHS As Point)
'    pElement.Twists = RHS
'End Property
'
'Friend Property Get Element_Twists() As Point
'    Set Element_Twists = pElement.Twists
'End Property

'Private Property Set Element_Twists(RHS As Point)
'    Set pElement.Twists = RHS
'End Property
'
'Private Property Get Element_Twists() As Point
'    Set Element_Twists = pElement.Twists
'End Property



Public Property Get CameraClip() As Boolean
    CameraClip = modMain.CameraClip
End Property
Public Property Let CameraClip(ByVal RHS As Boolean)
    modMain.CameraClip = RHS
End Property

'Public Property Get Object() As Element
'    Set Object = pElement
'End Property
'Public Property Set Object(ByRef RHS As Element)
'    Set pElement = RHS
'End Property
'Public Property Let Object(ByRef RHS)
'    Set pElement = ObjectByPtr(RHS)
'End Property
Public Property Set Value(ByVal RHS)
    Dim obj As Object
    Set obj = ObjectByPtr(RHS)
    If TypeName(obj) = "Element" Then
        Set pElement = obj
    ElseIf TypeName(obj) = "Camera" Then
        Set pCamera = obj
    End If
    If obj.Player Is Nothing Then
        Set RHS.Player = Me
    ElseIf Not obj.Player.Value = Me.Value Then
        Set RHS.Player = Me
    End If
End Property

Public Property Let Value(ByVal RHS As Long)
    Value = RHS
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "240"
    Value = ObjPtr(Me)
End Property

'Public Property Get MaxZoom() As Single
'    MaxZoom = pMaxZoom
'End Property
'Public Property Let MaxZoom(ByVal RHS As Single)
'    pMaxZoom = RHS
'End Property
'
'Public Property Get MinZoom() As Single
'    MinZoom = pMinZoom
'End Property
'Public Property Let MinZoom(ByVal RHS As Single)
'    pMinZoom = RHS
'End Property
'
'Public Property Get Angle() As Single
'    Angle = pAngle
'End Property
'Public Property Let Angle(ByVal RHS As Single)
'    pAngle = RHS
''    modParse.Player.pElement.Twists.Y = (modParse.Player.pElement.Twists.Y + (pAngle * DEGREE)) - modParse.Player.pElement.Rotate.Y
'End Property
'
'Public Property Get Pitch() As Single
'    Pitch = pPitch
'End Property
'Public Property Let Pitch(ByVal RHS As Single)
'    pPitch = RHS
'End Property
'
'Public Property Get Zoom() As Single
'    Zoom = pZoom
'End Property
'Public Property Let Zoom(ByVal RHS As Single)
'    pZoom = RHS
'End Property

Public Property Get CameraIndex() As Long
    CameraIndex = pCameraIndex
End Property
Public Property Let CameraIndex(ByVal RHS As Long)
    pCameraIndex = RHS
End Property

Public Property Get Speed() As Single
    Speed = pSpeed
End Property
Public Property Let Speed(ByVal RHS As Single)
    pSpeed = RHS
End Property

Public Property Get Friction() As Single
    Friction = pFriction
End Property
Public Property Let Friction(ByVal RHS As Single)
    pFriction = RHS
End Property

Public Sub ZoomIn(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    If Perspective = ThirdPerson Then
        modParse.Player.Camera.Zoom = modParse.Player.Camera.Zoom - Measure
        If modParse.Player.Camera.Zoom < modParse.Player.Camera.MinZoom Then modParse.Player.Camera.Zoom = modParse.Player.Camera.MinZoom
    End If
End Sub

Public Sub ZoomOut(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    If Perspective = ThirdPerson Then
        modParse.Player.Camera.Zoom = modParse.Player.Camera.Zoom + Measure
        If modParse.Player.Camera.Zoom > modParse.Player.Camera.MaxZoom Then modParse.Player.Camera.Zoom = modParse.Player.Camera.MaxZoom
    End If
End Sub

Public Sub LookLeft(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    If Not Perspective = CameraMode Then modParse.Player.Camera.Angle = (modParse.Player.Camera.Angle + Measure)
    If modParse.Player.Camera.Angle < -(PI * 2) Then modParse.Player.Camera.Angle = (modParse.Player.Camera.Angle + (PI * 2))
    'modParse.Player.Twists.Y = 1
    
End Sub
Public Sub LookRight(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    If Not Perspective = CameraMode Then modParse.Player.Camera.Angle = (modParse.Player.Camera.Angle - Measure)
    If modParse.Player.Camera.Angle > (PI * 2) Then modParse.Player.Camera.Angle = (modParse.Player.Camera.Angle - (PI * 2))
    'modParse.Player.Twists.Y = -1
    
End Sub
Public Sub LookUp(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    modParse.Player.Camera.Pitch = modParse.Player.Camera.Pitch - Measure
    If modParse.Player.Camera.Pitch < -1.5 Then modParse.Player.Camera.Pitch = -1.5

End Sub

Public Sub LookDown(Optional ByVal Measure As Single = 0.015)
    ResetIdle
    modParse.Player.Camera.Pitch = modParse.Player.Camera.Pitch + Measure
    If modParse.Player.Camera.Pitch > 1.5 Then modParse.Player.Camera.Pitch = 1.5

End Sub

Public Sub MoveForward()
    ResetIdle

    If Not modParse.Player.Element Is Nothing Then
        '############################################
        '################ FOREWARDS #################
        Dim vecDirect As D3DVECTOR
        If modParse.Player.Element.Direct.Y = 0 Then
            vecDirect.X = Sin(D720 - modParse.Player.Camera.Angle)
            vecDirect.Z = Cos(D720 - modParse.Player.Camera.Angle)
            If ((Perspective = Spectator) Or DebugMode) Or modParse.Player.Element.InLiquid Then
                vecDirect.Y = -(Tan(D720 - modParse.Player.Camera.Pitch))
            End If
            D3DXVec3Normalize vecDirect, vecDirect
            If modParse.Player.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (modParse.Player.Speed / 2), Friction
            ElseIf (Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), modParse.Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
    End If
End Sub
Public Sub MoveBackwards()
    ResetIdle
    If Not modParse.Player.Element Is Nothing Then
        '############################################
        '################ BACKWARDS #################
        Dim vecDirect As D3DVECTOR
        If modParse.Player.Element.Direct.Y = 0 Then
            vecDirect.X = -Sin(D720 - modParse.Player.Camera.Angle)
            vecDirect.Z = -Cos(D720 - modParse.Player.Camera.Angle)
            If (Perspective = Spectator) Or DebugMode Then
                vecDirect.Y = Tan(D720 - modParse.Player.Camera.Pitch)
            End If
            D3DXVec3Normalize vecDirect, vecDirect
            If modParse.Player.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (modParse.Player.Speed / 2), Friction
            ElseIf (Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), modParse.Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
    End If
End Sub
Public Sub SlideLeft()
    ResetIdle

    If Not modParse.Player.Element Is Nothing Then
        '############################################
        '############### STRAFE LEFT ################
        Dim vecDirect As D3DVECTOR
        If modParse.Player.Element.Direct.Y = 0 Then
            vecDirect.X = Sin((D720 - modParse.Player.Camera.Angle) - D180)
            vecDirect.Z = Cos((D720 - modParse.Player.Camera.Angle) - D180)
            D3DXVec3Normalize vecDirect, vecDirect
            If modParse.Player.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (modParse.Player.Speed / 2), Friction
            ElseIf (Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), modParse.Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
    End If
    
End Sub
Public Sub SlideRight()
    ResetIdle
    
    If Not modParse.Player.Element Is Nothing Then
        '############################################
        '############### STRAFE RIGHT ###############
        Dim vecDirect As D3DVECTOR
        If modParse.Player.Element.Direct.Y = 0 Then
            vecDirect.X = Sin((D720 - modParse.Player.Camera.Angle) + D180)
            vecDirect.Z = Cos((D720 - modParse.Player.Camera.Angle) + D180)
            D3DXVec3Normalize vecDirect, vecDirect
            If modParse.Player.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (modParse.Player.Speed / 2), Friction
            ElseIf (Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), modParse.Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
    End If
End Sub
Public Sub Jump()
    
    ResetIdle
    
    If Not modParse.Player.Element Is Nothing Then
        If ((Perspective = Spectator) Or DebugMode) Or (((Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying))) And _
                                            (Not ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling))) Then
            '############################################
            '################## JUMP ####################
            Dim vecDirect As D3DVECTOR
            If (Perspective = Spectator) Or DebugMode Then
                vecDirect.Y = vecDirect.Y + IIf(modParse.Player.Speed < 1, 1, modParse.Player.Speed)
                modParse.Player.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), modParse.Player.Speed, Friction
            Else
    '            If modParse.Player.Element.MotionExists(JumpGUID) Then
                    If (Not ((modParse.Player.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((modParse.Player.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                        Do Until Not modParse.Player.Element.MotionExists(JumpGUID)
                            modParse.Player.Element.DeleteMotion JumpGUID
                        Loop
                    End If
    '            End If
    '            If Not modParse.Player.Element.MotionExists(JumpGUID) Then
                    vecDirect.Y = IIf(modParse.Player.Element.InLiquid, 5, 9)
                    JumpGUID = modParse.Player.Element.AddMotion(Actions.Directing, JumpGUID, ToPoint(vecDirect), modParse.Player.Speed * 4, Friction)
    '            End If
            End If
            '############################################
            '############################################
        
        End If
    End If
End Sub

Public Sub ChangePerspective()
    '############################################
    '############ SWITCH PERSPECTIVE ############
    ResetIdle
    If Perspective = Playmode.ThirdPerson Then
        Perspective = Playmode.FirstPerson
    ElseIf Perspective = Playmode.FirstPerson Then
        Perspective = IIf((Cameras.Count > 0), Playmode.CameraMode, Playmode.ThirdPerson)
    ElseIf Perspective = Playmode.CameraMode Then
        Perspective = Playmode.ThirdPerson
    End If
    '############################################
    '############################################
End Sub

Private Sub Class_Initialize()
    vbaObjSetAddref Orient, ObjPtr(New Orient)
    
    Stats_Player_Count = Stats_Player_Count + 1
    vbaObjSetAddref Element, ObjPtr(New Element)

    Set pElement = New Element
'    pElement.Key = "Player"
'    Elements.Add pElement, "Player"
    
    Set pCamera = New Camera

    pSpeed = 0.01
    pFriction = 0.05
    modMain.CameraClip = True
    
End Sub

Private Sub Class_Terminate()

    Set pElement = Nothing

    Set pCamera = Nothing

    vbaObjSet Orient, ObjPtr(Nothing)
    Stats_Player_Count = Stats_Player_Count - 1
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

'Public Property Let AttachedTo(ByVal RHS As String)
'    Element.AttachedTo = RHS
'End Property
'
'Public Property Get AttachedTo() As String
'    AttachedTo = Element.AttachedTo
'End Property
'
'Public Property Set Attachments(RHS As NTNodes10.Collection)
'    Set Element.Attachments = RHS
'End Property
'
'Public Property Get Attachments() As NTNodes10.Collection
'    Set Attachments = Element.Attachments
'End Property
'
'Public Sub AutoFulcrums(Optional ByVal Smart As Boolean = True)
'    Element.AutoFulcrums Smart
'End Sub
'
'Public Property Set Centoid(RHS As Point)
'    Set Element.Centoid = RHS
'End Property
'
'Public Property Get Centoid() As Point
'    Set Centoid = Element.Centoid
'End Property
'
'Public Property Set Displace(RHS As Point)
'    Set Element.Displace = RHS
'End Property
'
'Public Property Get Displace() As Point
'    Set Displace = Element.Displace
'End Property
'
'Public Property Set Fulcrum(Optional ByVal Index As Integer = 0, RHS As Point)
'    Set Element.Fulcrums(Index) = RHS
'End Property
'
'Public Property Let Fulcrum(Optional ByVal Index As Integer = 0, ByVal RHS As Variant)
'    Element.Fulcrums(Index) = RHS
'End Property
'
'Public Property Get Fulcrum(Optional ByVal Index As Integer = 0) As Point
'    Set Fulcrum = Element.Fulcrum(Index)
'End Property
'
'Public Property Get Fulcrums() As Collection
'    Set Fulcrums = Element.Fulcrums
'End Property
'
'Public Function MotionExists(ByVal Key As String) As Boolean
'    MotionExists = Element.MotionExists(Key)
'End Function
'
'Public Property Get Motions() As NTNodes10.Collection
'    Set Motions = Element.Motions
'End Property
'
'Public Property Set Motions(ByRef RHS As NTNodes10.Collection)
'    Set Element.Motions = RHS
'End Property
'
'Public Function AddMotion(Action As Long, ByVal aGUID As String, Data As Point, Optional Emphasis As Single = 0#, Optional ByVal Friction As Single = 0#, Optional ByVal Reactive As Single = -1#, Optional ByVal Recount As Single = -1#, Optional Script As String = "") As String
'    AddMotion = Element.AddMotion(Action, aGUID, Data, Emphasis, Friction, Reactive, Recount, Script)
'End Function
'
'Public Sub ApplyMatrix()
'    Element.ApplyMatrix
'End Sub
'
'Public Property Let Alphablend(ByVal RHS As Boolean)
'    Element.Alphablend = RHS
'End Property
'
'Public Property Get Alphablend() As Boolean
'    Alphablend = Element.Alphablend
'End Property
'
'Public Property Let BoundsIndex(ByVal RHS As Long)
'    Element.BoundsIndex = RHS
'End Property
'
'Public Property Get BoundsIndex() As Long
'    BoundsIndex = Element.BoundsIndex
'End Property
'
'Public Property Let BoundsModel(ByVal RHS As String)
'    Element.BoundsModel = RHS
'End Property
'
'Public Property Get BoundsModel() As String
'    BoundsModel = Element.BoundsModel
'End Property
'
'Public Sub ClearMotions()
'    Element.ClearMotions
'End Sub
'
'Public Property Let Collision(ByVal RHS As Boolean)
'    Element.Collision = RHS
'End Property
'
'Public Property Get Collision() As Boolean
'    Collision = Element.Collision
'End Property
'
'Public Property Let CollideFaces(ByVal RHS As Long)
'      Element.CollideFaces = RHS
'End Property
'
'Public Property Get CollideFaces() As Long
'    CollideFaces = Element.CollideFaces
'End Property
'
'Public Property Let CollideIndex(ByVal RHS As Long)
'    Element.CollideIndex = RHS
'End Property
'
'Public Property Get CollideIndex() As Long
'    CollideIndex = Element.CollideIndex
'End Property
'
'Public Property Let CollideObject(ByVal RHS As Long)
'    Element.CollideObject = RHS
'End Property
'
'Public Property Get CollideObject() As Long
'    CollideObject = Element.CollideObject
'End Property
'
'Public Property Let CulledFaces(ByVal RHS As Long)
'    Element.CulledFaces = RHS
'End Property
'
'Public Property Get CulledFaces() As Long
'    CulledFaces = Element.CulledFaces
'End Property
'
'Public Sub DefaultMatrix()
'    Element.DefaultMatrix
'End Sub
'
'Public Function DeleteMotion(ByVal MGUID As String) As Boolean
'    DeleteMotion = Element.DeleteMotion(MGUID)
'End Function
'
'Public Property Let Effect(ByVal RHS As Long)
'    Element.Effect = RHS
'End Property
'
'Public Property Get Effect() As Long
'    Effect = Element.Effect
'End Property
'
'Public Property Let FileName(ByVal RHS As String)
'    Element.FileName = RHS
'End Property
'
'Public Property Get FileName() As String
'    FileName = Element.FileName
'End Property
'
'Public Property Let Gravitational(ByVal RHS As Boolean)
'    Element.Gravitational = RHS
'End Property
'
'Public Property Get Gravitational() As Boolean
'    Gravitational = Element.Gravitational
'End Property
'
'Public Property Let InLiquid(ByVal RHS As Boolean)
'     Element.InLiquid = RHS
'End Property
'
'Public Property Get InLiquid() As Boolean
'    InLiquid = Element.InLiquid
'End Property
'
'Public Property Let IsMoving(ByVal RHS As Long)
'     Element.IsMoving = RHS
'End Property
'
'Public Property Get IsMoving() As Long
'    IsMoving = Element.IsMoving
'End Property
'
'Public Property Let Key(ByVal RHS As String)
'    Element.Key = RHS
'End Property
'
'Public Property Get Key() As String
'    Key = Element.Key
'End Property
'
'Public Property Set Offset(RHS As Point)
'    Set Element.Offset = RHS
'End Property
'
'Public Property Get Offset() As Point
'    Set Offset = Element.Offset
'End Property
'
'Public Property Let OnLadder(ByVal RHS As Boolean)
'     Element.OnLadder = RHS
'End Property
'
'Public Property Get OnLadder() As Boolean
'    OnLadder = Element.OnLadder
'End Property
'
'Public Property Set Origin(RHS As Point)
'    Element.Origin = RHS
'End Property
'
'Public Property Get Origin() As Point
'    Set Origin = Element.Origin
'End Property
'


'Public Sub PlayerMatrix()
'    Element.PlayerMatrix
'End Sub
'
'Public Function PointMatrix(p As Point) As Point
'    Set PointMatrix = Element.PointMatrix(p)
'End Function
'
'Public Sub PrepairMatrix()
'    Element.PrepairMatrix
'End Sub



'Public Sub Replacer(ByVal TextFilename As String, ByVal BoardKey As String)
'    Element.Replacer TextFilename, BoardKey
'End Sub
'
'Public Property Set ReplacerKeys(ByRef RHS As NTNodes10.Collection)
'    Set Element.ReplacerKeys = RHS
'End Property
'
'Public Property Get ReplacerKeys() As NTNodes10.Collection
'    Set ReplacerKeys = Element.ReplacerKeys
'End Property
'
'Public Property Set ReplacerVals(ByRef RHS As NTNodes10.Collection)
'    Set Element.ReplacerVals = RHS
'End Property
'
'Public Property Get ReplacerVals() As NTNodes10.Collection
'    Set ReplacerVals = Element.ReplacerVals
'End Property
'
'Public Property Set Rotate(RHS As Point)
'    Set Element.Rotate = RHS
'End Property
'
'Public Property Get Rotate() As Point
'    Set Rotate = Element.Rotate
'End Property
'
'Public Property Set Scaled(RHS As Point)
'    Set Element.Scaled = RHS
'End Property
'
'Public Property Get Scaled() As Point
'    Set Scaled = Element.Scaled
'End Property

Public Property Let Serialize(ByVal RHS As Boolean)
    Element.Serialize = RHS
End Property

Public Property Get Serialize() As Boolean
    Serialize = Serialize
End Property

'Public Sub SetWorldMatrix()
'    Element.SetWorldMatrix
'End Sub

'Public Property Let Visible(ByVal RHS As Boolean)
'    Element.Visible = RHS
'End Property
'
'Public Property Get Visible() As Boolean
'    Visible = Element.Visible
'End Property
'
'Public Property Let VisualIndex(ByVal RHS As Long)
'    Element.VisualIndex = RHS
'End Property
'
'Public Property Get VisualIndex() As Long
'    VisualIndex = Element.VisualIndex
'End Property
'
'Public Property Let VisualModel(ByVal RHS As String)
'    Element.VisualModel = RHS
'End Property
'
'Public Property Get VisualModel() As String
'    VisualModel = Element.VisualModel
'End Property
'
'Public Property Let WireFrame(ByVal RHS As Boolean)
'    Element.WireFrame = RHS
'End Property
'
'Public Property Get WireFrame() As Boolean
'    WireFrame = Element.WireFrame
'End Property


'Private Property Let Element_AttachedTo(ByVal RHS As String)
'    Element.AttachedTo = RHS
'End Property
'
'Private Property Get Element_AttachedTo() As String
'    Element_AttachedTo = Element.AttachedTo
'End Property
'
'Private Property Set Element_Attachments(RHS As NTNodes10.Collection)
'    Set Element.Attachments = RHS
'End Property
'
'Private Property Get Element_Attachments() As NTNodes10.Collection
'    Set Element_Attachments = Element.Attachments
'End Property
'
'Private Sub Element_AutoFulcrums(Optional ByVal Smart As Boolean = True)
'    Element.AutoFulcrums Smart
'End Sub
'
'Private Property Set Element_Centoid(RHS As Point)
'    Set Element.Centoid = RHS
'End Property
'
'Private Property Get Element_Centoid() As Point
'    Set Element_Centoid = Element.Centoid
'End Property
'
'Private Property Set Element_Displace(RHS As Point)
'    Set Element.Displace = RHS
'End Property
'
'Private Property Get Element_Displace() As Point
'    Set Element_Displace = Element.Displace
'End Property
'
'Private Property Set Element_Fulcrum(Optional ByVal Index As Integer = 0, RHS As Point)
'    Set Element.Fulcrums(Index) = RHS
'End Property
'
'Private Property Let Element_Fulcrum(Optional ByVal Index As Integer = 0, ByVal RHS As Variant)
'    Element.Fulcrums(Index) = RHS
'End Property
'
'Private Property Get Element_Fulcrum(Optional ByVal Index As Integer = 0) As Point
'    Set Element_Fulcrum = Element.Fulcrum(Index)
'End Property
'
'Private Property Get Element_Fulcrums() As Collection
'    Set Element_Fulcrums = Element.Fulcrums
'End Property
'
'Private Function Element_MotionExists(ByVal Key As String) As Boolean
'    Element_MotionExists = Element.MotionExists(Key)
'End Function
'
'Private Property Get Element_Motions() As NTNodes10.Collection
'    Set Element_Motions = Element.Motions
'End Property
'
'Private Property Set Element_Motions(ByRef RHS As NTNodes10.Collection)
'    Set Element.Motions = RHS
'End Property
'
'Private Function Element_AddMotion(Action As Long, ByVal aGUID As String, Data As Point, Optional Emphasis As Single = 0#, Optional ByVal Friction As Single = 0#, Optional ByVal Reactive As Single = -1#, Optional ByVal Recount As Single = -1#, Optional Script As String = "") As String
'    Element_AddMotion = Element.AddMotion(Action, aGUID, Data, Emphasis, Friction, Reactive, Recount, Script)
'End Function
'
'Private Sub Element_ApplyMatrix()
'    Element.ApplyMatrix
'End Sub
'
'Private Property Let Element_Alphablend(ByVal RHS As Boolean)
'    Element.Alphablend = RHS
'End Property
'
'Private Property Get Element_Alphablend() As Boolean
'    Element_Alphablend = Element.Alphablend
'End Property
'
'Private Property Let Element_BoundsIndex(ByVal RHS As Long)
'    Element.BoundsIndex = RHS
'End Property
'
'Private Property Get Element_BoundsIndex() As Long
'    Element_BoundsIndex = Element.BoundsIndex
'End Property
'
'Private Property Let Element_BoundsModel(ByVal RHS As String)
'    Element.BoundsModel = RHS
'End Property
'
'Private Property Get Element_BoundsModel() As String
'    Element_BoundsModel = Element.BoundsModel
'End Property
'
'Private Sub Element_ClearMotions()
'    Element.ClearMotions
'End Sub
'
'Private Property Let Element_Collision(ByVal RHS As Boolean)
'    Element.Collision = RHS
'End Property
'
'Private Property Get Element_Collision() As Boolean
'    Element_Collision = Element.Collision
'End Property
'
'Private Property Let Element_CollideFaces(ByVal RHS As Long)
'    Element.CollideFaces = RHS
'End Property
'
'Private Property Get Element_CollideFaces() As Long
'    Element_CollideFaces = Element.CollideFaces
'End Property
'
'Private Property Let Element_CollideIndex(ByVal RHS As Long)
'    Element.CollideIndex = RHS
'End Property
'
'Private Property Get Element_CollideIndex() As Long
'    Element_CollideIndex = Element.CollideIndex
'End Property
'
'Private Property Let Element_CollideObject(ByVal RHS As Long)
'    Element.CollideObject = RHS
'End Property
'
'Private Property Get Element_CollideObject() As Long
'    Element_CollideObject = Element.CollideObject
'End Property
'
'Private Property Let Element_CulledFaces(ByVal RHS As Long)
'    Element.CulledFaces = RHS
'End Property
'
'Private Property Get Element_CulledFaces() As Long
'    Element_CulledFaces = Element.CulledFaces
'End Property
'
'Private Sub Element_DefaultMatrix()
'    Element.DefaultMatrix
'End Sub
'
'Private Function Element_DeleteMotion(ByVal MGUID As String) As Boolean
'    Element_DeleteMotion = Element.DeleteMotion(MGUID)
'End Function
'
'Private Property Let Element_Effect(ByVal RHS As Long)
'    Element.Effect = RHS
'End Property
'
'Private Property Get Element_Effect() As Long
'    Element_Effect = Element.Effect
'End Property
'
'Private Property Let Element_FileName(ByVal RHS As String)
'    Element.FileName = RHS
'End Property
'
'Private Property Get Element_FileName() As String
'    Element_FileName = Element.FileName
'End Property
'
'Private Property Let Element_Gravitational(ByVal RHS As Boolean)
'    Element.Gravitational = RHS
'End Property
'
'Private Property Get Element_Gravitational() As Boolean
'    Element_Gravitational = Element.Gravitational
'End Property
'
'Private Property Let Element_InLiquid(ByVal RHS As Boolean)
'    Element.InLiquid = RHS
'End Property
'
'Private Property Get Element_InLiquid() As Boolean
'    Element_InLiquid = Element.InLiquid
'End Property
'
'Private Property Let Element_IsMoving(ByVal RHS As Long)
'    Element.IsMoving = RHS
'End Property
'
'Private Property Get Element_IsMoving() As Long
'    Element_IsMoving = Element.IsMoving
'End Property
'
'Private Property Let Element_Key(ByVal RHS As String)
'    Element.Key = RHS
'End Property
'
'Private Property Get Element_Key() As String
'    Element_Key = Element.Key
'End Property
'
'Private Property Set Element_Offset(RHS As Point)
'    Set Element.Offset = RHS
'End Property
'
'Private Property Get Element_Offset() As Point
'    Set Element_Offset = Element.Offset
'End Property
'
'Private Property Let Element_OnLadder(ByVal RHS As Boolean)
'    Element.OnLadder = RHS
'End Property
'
'Private Property Get Element_OnLadder() As Boolean
'    Element_OnLadder = Element.OnLadder
'End Property
'
'Private Property Set Element_Origin(RHS As Point)
'    Element.Origin = RHS
'End Property
'
'Private Property Get Element_Origin() As Point
'    Set Element_Origin = Element.Origin
'End Property
'
'Private Sub Element_PlayerMatrix()
'    Element.PlayerMatrix
'End Sub
'
'Private Function Element_PointMatrix(p As Point) As Point
'    Set Element_PointMatrix = Element.PointMatrix(p)
'End Function
'
'Private Sub Element_PrepairMatrix()
'    Element.PrepairMatrix
'End Sub
'
'Private Sub Element_Replacer(ByVal TextFilename As String, ByVal BoardKey As String)
'    Element.Replacer TextFilename, BoardKey
'End Sub
'
'Private Property Set Element_ReplacerKeys(ByRef RHS As NTNodes10.Collection)
'    Set Element.ReplacerKeys = RHS
'End Property
'
'Private Property Get Element_ReplacerKeys() As NTNodes10.Collection
'    Set Element_ReplacerKeys = Element.ReplacerKeys
'End Property
'
'Private Property Set Element_ReplacerVals(ByRef RHS As NTNodes10.Collection)
'    Set Element.ReplacerVals = RHS
'End Property
'
'Private Property Get Element_ReplacerVals() As NTNodes10.Collection
'    Set Element_ReplacerVals = Element.ReplacerVals
'End Property
'
'Private Property Set Element_Rotate(RHS As Point)
'    Set Element.Rotate = RHS
'End Property
'
'Private Property Get Element_Rotate() As Point
'    Set Element_Rotate = Element.Rotate
'End Property
'
'
'Private Property Set Element_Scaled(RHS As Point)
'    Set Element.Scaled = RHS
'End Property
'
'Private Property Get Element_Scaled() As Point
'    Set Element_Scaled = Element.Scaled
'End Property
'
'Private Property Let Element_Serialize(ByVal RHS As Boolean)
'    Element.Serialize = RHS
'End Property
'
'Private Property Get Element_Serialize() As Boolean
'    Element_Serialize = Element_Serialize
'End Property
'
'Private Sub Element_SetWorldMatrix()
'    Element.SetWorldMatrix
'End Sub
'
'
'Private Property Get Element_Value() As Long
'    Element_Value = ObjPtr(Element)
'End Property
'
'Private Property Let Element_Visible(ByVal RHS As Boolean)
'    Element.Visible = RHS
'End Property
'
'Private Property Get Element_Visible() As Boolean
'    Element_Visible = Element.Visible
'End Property
'
'Private Property Let Element_VisualIndex(ByVal RHS As Long)
'    Element.VisualIndex = RHS
'End Property
'
'Private Property Get Element_VisualIndex() As Long
'    Element_VisualIndex = Element.VisualIndex
'End Property
'
'Private Property Let Element_VisualModel(ByVal RHS As String)
'    Element.VisualModel = RHS
'End Property
'
'Private Property Get Element_VisualModel() As String
'    Element_VisualModel = Element.VisualModel
'End Property
'
'Private Property Let Element_WireFrame(ByVal RHS As Boolean)
'    Element.WireFrame = RHS
'End Property
'
'Private Property Get Element_WireFrame() As Boolean
'    Element_WireFrame = Element.WireFrame
'End Property


