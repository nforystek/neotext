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

Private pCameraIndex As Long

Private pSpeed As Single

Private pFriction As Single

Private pCamera As Camera


Friend Sub LateEval()
'    pCamera.LateEval
    
End Sub

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
    If Not pCamera Is Nothing Then
        Set Element = pCamera.Element
    End If
End Property
Public Property Set Element(ByRef RHS As Element)
    SetObject RHS
End Property
Public Property Let Element(ByRef RHS As Element)
    SetObject RHS
End Property
Public Property Get Camera() As Camera
    Set Camera = pCamera
End Property
Public Property Set Camera(ByRef RHS As Camera)
    SetObject RHS
End Property
Public Property Let Camera(ByRef RHS As Camera)
    SetObject RHS
End Property

Private Sub SetObject(ByRef Obj)
    If TypeName(Obj) = "Camera" Then
        Set pCamera = Obj
        Set modParse.Camera = Obj
        If Not frmMain.StartUp Then frmMain.ExecuteStatement "Set Camera = All(""" & Obj.Key & """)"
        vbaObjSetAddref Orient, ObjPtr(pCamera)
    ElseIf TypeName(Obj) = "Element" Then
        If pCamera Is Nothing Then
            SetObject New Camera
        End If
        Set pCamera.Element = Obj
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
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "600"
    Value = ObjPtr(Me)
End Property


Public Property Get CameraClip() As Boolean
    CameraClip = modMain.CameraClip
End Property
Public Property Let CameraClip(ByVal RHS As Boolean)
    modMain.CameraClip = RHS
End Property


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


Public Sub MoveForward()
    ResetIdle

    If Not Camera.Element Is Nothing Then
    
    
        '############################################
        '################ FOREWARDS #################
        Dim vecDirect As D3DVECTOR
        If Camera.Direct.Y = 0 Then
            vecDirect.X = Sin(D720 - Camera.Rotate.Y)
            vecDirect.Z = Cos(D720 - Camera.Rotate.Y)
            If ((Perspective = Spectator) Or DebugMode) Or Camera.Element.InLiquid Then
                vecDirect.Y = -(Tan(D720 - Camera.Rotate.X))
            End If
            D3DXVec3Normalize vecDirect, vecDirect
            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            ElseIf (Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
        
        
    End If
End Sub
Public Sub MoveBackward()
    ResetIdle
    If Not Camera.Element Is Nothing Then
    
    
        '############################################
        '################ BACKWARDS #################
        Dim vecDirect As D3DVECTOR
        If Camera.Direct.Y = 0 Then
            vecDirect.X = -Sin(D720 - Camera.Rotate.Y)
            vecDirect.Z = -Cos(D720 - Camera.Rotate.Y)
            If (Perspective = Spectator) Or DebugMode Then
                vecDirect.Y = Tan(D720 - Camera.Rotate.X)
            End If
            D3DXVec3Normalize vecDirect, vecDirect
            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            ElseIf (Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
        
        
    End If
End Sub
Public Sub SlideLeft()
    ResetIdle

    If Not Camera.Element Is Nothing Then
    
    
        '############################################
        '############### STRAFE LEFT ################
        Dim vecDirect As D3DVECTOR
        If Camera.Element.Direct.Y = 0 Then
            vecDirect.X = Sin((D720 - Camera.Rotate.Y) - D180)
            vecDirect.Z = Cos((D720 - Camera.Rotate.Y) - D180)
            D3DXVec3Normalize vecDirect, vecDirect
            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            ElseIf (Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
        
        
    End If
    
End Sub
Public Sub SlideRight()
    ResetIdle
    
    If Not Camera.Element Is Nothing Then
    
    
        '############################################
        '############### STRAFE RIGHT ###############
        Dim vecDirect As D3DVECTOR
        If Camera.Element.Direct.Y = 0 Then
            vecDirect.X = Sin((D720 - Camera.Rotate.Y) + D180)
            vecDirect.Z = Cos((D720 - Camera.Rotate.Y) + D180)
            D3DXVec3Normalize vecDirect, vecDirect
            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            ElseIf (Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                            ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
        End If
        '############################################
        '############################################
        
        
    End If
End Sub
Public Sub SlideUp()
    
    ResetIdle
    
    If Not Camera.Element Is Nothing Then
    
    
'        If ((Perspective = Spectator) Or DebugMode) Or (((Not ((Element.IsMoving And Moving.Flying) = Moving.Flying))) And _
'                                            (Not ((Element.IsMoving And Moving.Falling) = Moving.Falling))) Then
                                            
                                            
            '############################################
            '################## JUMP ####################
            Dim vecDirect As D3DVECTOR

                vecDirect.Y = vecDirect.Y + IIf(Player.Speed < 1, 1, Player.Speed)
                
            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            Else
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
            
                

            '############################################
            '############################################
            
        
'        End If
    End If
End Sub
Public Sub SlideDown()
    
    ResetIdle
    
    If Not Camera.Element Is Nothing Then
    
    
'        If ((Perspective = Spectator) Or DebugMode) Or (((Not ((Element.IsMoving And Moving.Flying) = Moving.Flying))) And _
'                                            (Not ((Element.IsMoving And Moving.Falling) = Moving.Falling))) Then
                                            
                                            
            '############################################
            '################## JUMP ####################
            Dim vecDirect As D3DVECTOR

                vecDirect.Y = vecDirect.Y - IIf(Player.Speed < 1, 1, Player.Speed)

            If Camera.Element.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), (Player.Speed / 2), Friction
            Else
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            End If
            
            '############################################
            '############################################
            
        
'        End If
    End If
End Sub
Public Sub Jump()
    
    ResetIdle
    
    If Not Camera.Element Is Nothing Then
    
    
        If ((Perspective = Spectator) Or DebugMode) Or (((Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying))) And _
                                            (Not ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling))) Then
                                            
                                            
            '############################################
            '################## JUMP ####################
            Dim vecDirect As D3DVECTOR
            If (Perspective = Spectator) Or DebugMode Then
                vecDirect.Y = vecDirect.Y + IIf(Player.Speed < 1, 1, Player.Speed)
                Camera.Element.AddMotion Actions.Directing, Replace(modGuid.GUID, "-", "K"), ToPoint(vecDirect), Player.Speed, Friction
            Else
                If (Not ((Camera.Element.IsMoving And Moving.Flying) = Moving.Flying) Or _
                        ((Camera.Element.IsMoving And Moving.Falling) = Moving.Falling)) Then
                    Do Until Not Camera.Element.MotionExists(JumpGUID)
                        Camera.Element.DeleteMotion JumpGUID
                    Loop
                End If
                vecDirect.Y = IIf(Camera.Element.InLiquid, 5, 9)
                JumpGUID = Camera.Element.AddMotion(Actions.Directing, JumpGUID, ToPoint(vecDirect), Player.Speed * 4, Friction)
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
    Set pCamera = New Camera

    
    vbaObjSetAddref Orient, ObjPtr(pCamera)
    
    Stats_Player_Count = Stats_Player_Count + 1
'    vbaObjSetAddref Element, ObjPtr(New Element)
'
'    Set pElement = New Element
'    Set pElement.Player = Me
    
    

'    Set pCamera.Player = Me

'    Set pCamera.Element = pElement
'    Set pElement.Camera = pCamera

    pSpeed = 0.01
    pFriction = 0.05
    modMain.CameraClip = True
    
End Sub

Private Sub Class_Terminate()

'    Set pElement = Nothing

 '  Set pCamera = Nothing

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
