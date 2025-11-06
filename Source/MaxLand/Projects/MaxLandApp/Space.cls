VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Space"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'
'       Scenario:                           Qpaque:                     Gradient:
'
'       Range>0 And Boundary>Range          0-Range=100%-100%           Range-Boundary=100%-0%
'       Range=-1                            *100%
'       Range=0                             *0%                         0-Boundary=100%-0%
'       Range>0 And Boundary=-1                                         0-Range=100%-0%
'       Range>0 And Boundary=0                                          *100%
'       Range<Boundary And Boundary>0       Boundary-Range=100%-100%    0-Boundry=0%-100%
'
'
'
'
'Boundary is used to limit playing field to a fixed invisible fence, at Boundary (radius) away from origin.
'Range is used to considerably determine the emersiveness of whether one is considered in the space or not.
'Whereever boundaries overlap from one space to another, technically being in 2 the limitation doesn't exist
'they both together act differently than their intended uses, with the following explained value scenarios:
'
'Where
'   Boundary = *+# 'any positive number limits players to a distance from origin like being fenced
'   Range = *+# 'any positive number to the origin is where you are as fully emmersed in the space
'
'Where
'   Boundary = 0 'zero causes the origin of the space to remain centered on the expereince, /follow
'   Range = 0 'zero disables any with-in range attributes of the space, the default values are zero
'
'Where
'   Boundary = -1 'removes limitation on stopping the player no matter how far they are from origin
'   Range = -1 'means you are considered with-in range no matter where you are, you're in the space
'
'Where
'   only one of Boundary or Range are set, then that radius to the origin is part of rendering
'   where, if a sky is set with an alphablend opacity it is opaque on orogin (or color is 100%
'   intensity on origin) and the opaque gets gradually more see through until invisible at the
'   Boundary or Range that is set (while color intensity becomes less, merging with black when
'   however with multiple spaces this black could be another spaces color, making a transition)
'
'   both Boundary and Range are set as any positive number, and Range is smaller then boundary,
'   then any alphablending of the sky opacity or any color intensity is at fully opaque when in
'   the range, and while with-in the Boundary but not with-in range, the opacity and intensity
'   will go from full opauqe or intense, and fade on to go fully invisisble at Boundaries end.
'
'   when both Boundary and Range are set as any positive number, and Range is greater then the
'   Boundary, then the opposed opacity and intensity rules apply in inverse nature invisibily.
'   therefore walking away from the origin would begin to turn up the opaque passing Boundary,
'   until full opqaue is everything beyond the Range. Boundary wont limit, smaller then Range.
'
'   For instance in example:  Boundary and Range set to true, which are their defaults, yeilds
'   a most "life-like" natrual sky rendering when provided with textures to build a cube sky of.
'
'   For instance in Example: Boundary set to zero, and Range set to -1, would cause the sky map
'   to hone it's rotation to your own, making just a forward facing view that you never change.
'
    
Private pSkyTop As String
Private pSkyBottom As String
Private pSkyFront As String
Private pSkyBack As String
Private pSkyRight As String
Private pSkyLeft As String
Private pSkyRotate As Single
Private pSkyRotated As Single

Private pOrigin As Point

Private pRange As Single

Private pGravity As Single

Private pFogEnabled As Boolean
Private pFogDistance As Single
Private pFogColor As New Color

Private pColor As New Color

Private pBoundary As Single
Private pVisible As Boolean

Private pSerialize As Boolean
Private pKey As String

Private pTranslucent As Boolean
Private pAlphablend As Boolean

Private SkyPlaq() As MyVertex
Private SkySkin() As Direct3DTexture8
Private SkyVBuf As Direct3DVertexBuffer8

Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
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



Friend Function InSpace(ByRef WayPoint As Point, Optional ByVal RestrictTo As Double = -2) As Boolean
    If RestrictTo = -2 Then
        If Range > 0 And Boundary < Range Then
            RestrictTo = Range
        ElseIf Boundary > 0 And Boundary > Range Then
            RestrictTo = Range + Boundary
        End If
    End If
    If (RestrictTo <> 0) Then
        If (RestrictTo = -1) Then
            InSpace = True
        Else
            InSpace = (DistanceEx(WayPoint, Origin) < RestrictTo)
        End If
    End If
End Function



Friend Sub RenderSpace()
    If InSpace(Player.Element.Origin) Then

        
        If Not SkyRotated = 0 Then
        
            If (FPSRate > 0) And SkyRotate > 0 Then

                SkyRotated = SkyRotated + (360 / (HoursInOneDay * SkyRotate)) * _
                    (PI / (HoursInOneDay * FPSRate)) * (FPSRate / HoursInOneDay)
            End If
            
            D3DXMatrixRotationY matWorld, SkyRotated * RADIAN
        End If
        
        If Visible Then
                        
            
            DDevice.SetTransform D3DTS_WORLD, matWorld
            DDevice.SetStreamSource 0, SkyVBuf, Len(SkyPlaq(0))
            
            
            If Translucent Or Alphablend Then
            

                
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(0)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(0)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 30, 2
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(1)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(1)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(2)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(2)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(3)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(3)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(4)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(4)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 24, 2
                DDevice.SetMaterial LucentMaterial
                DDevice.SetTexture 0, SkySkin(5)
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 1, SkySkin(5)
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 18, 2
                
                
            Else
               
                
            
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(0)
                DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 30, 2
                
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(1)
                'DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(2)
                'DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
                
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(3)
                'DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
                
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(4)
                'DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 24, 2
                
                DDevice.SetMaterial GenericMaterial
                DDevice.SetTexture 0, SkySkin(5)
                'DDevice.SetTexture 1, Nothing
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 18, 2
            
            End If
            
            
        End If
        
    End If

End Sub

Friend Property Get SkyRotated() As Single
    SkyRotated = pSkyRotated
End Property
Friend Property Let SkyRotated(ByVal RHS As Single)
    pSkyRotated = RHS
End Property
Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Property Get Range() As Single
    Range = pRange
End Property
Public Property Let Range(ByVal RHS As Single)
    pRange = RHS
End Property


Public Property Get Serialize() As Boolean
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Function ToString() As String

    ToString = ToString & Include.URLEncode("Space.SkyTop = """ & pSkyTop & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyBottom = """ & pSkyBottom & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyFront = """ & pSkyFront & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyBack = """ & pSkyBack & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyRight = """ & pSkyRight & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyLeft = """ & pSkyLeft & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyRotate = " & pSkyRotate) & vbCrLf
    
    ToString = ToString & Include.URLEncode("Space.FogColor = """ & pFogColor & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.FogEnabled = " & pFogEnabled) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.FogDistance = " & pFogDistance) & vbCrLf
    
    ToString = ToString & Include.URLEncode("Space.Gravity = " & pGravity) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.Range = " & pRange) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.Color = " & pRange) & vbCrLf
    
    ToString = ToString & Include.URLEncode("Space.Origin = """ & pOrigin & """") & vbCrLf

    ToString = ToString & Include.URLEncode("Space.Boundary = " & pBoundary) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.Visible = " & pVisible) & vbCrLf


    ToString = "  <Space>" & vbCrLf & ToString & vbCrLf & "</Space>" & vbCrLf
End Function

Public Property Get FogColor() As Color
    Set FogColor = pFogColor
End Property
Public Property Set FogColor(ByRef RHS)
    Set pFogColor = RHS
    DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(RHS.Alpha, RHS.Red, RHS.Green, RHS.Blue)
End Property
Public Property Let FogColor(ByVal RHS)
    If pFogColor Is Nothing Then Set pFogColor = New Color
    pFogColor.ToString = RHS
    
    DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(pFogColor.Alpha, pFogColor.Red, pFogColor.Green, pFogColor.Blue)
End Property

Public Property Get FogEnabled() As Boolean
    FogEnabled = pFogEnabled
End Property
Public Property Let FogEnabled(ByVal RHS As Boolean)
    pFogEnabled = RHS
    DDevice.SetRenderState D3DRS_FOGENABLE, RHS
End Property

Public Property Get FogDistance() As Single
    FogDistance = pFogDistance
End Property
Public Property Let FogDistance(ByVal RHS As Single)
    pFogDistance = RHS
    DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(RHS / 4)
    DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(RHS)
End Property

Public Property Get SkyRotate() As Single
    SkyRotate = pSkyRotate
End Property
Public Property Let SkyRotate(ByVal RHS As Single)
    pSkyRotate = RHS
End Property

Public Property Get Boundary() As Single
    Boundary = pBoundary
End Property
Public Property Let Boundary(ByVal RHS As Single)
    pBoundary = RHS
End Property


Public Property Get Gravitry() As Single
    Gravity = pGravity
End Property
Public Property Let Gravity(ByVal RHS As Single)
    pGravity = RHS
    SetMotion GlobalGravityDirect, Actions.Directing, MakePoint(0, RHS, 0), 1
    SetMotion LiquidGravityDirect, Actions.Directing, MakePoint(0, RHS / 40, 0), 2
End Property

Public Property Get SkyTop() As String
    SkyTop = pSkyTop
End Property
Public Property Let SkyTop(ByVal RHS As String)
    pSkyTop = RHS
    Set SkySkin(0) = LoadTexture(AppPath & "Models\" & RHS)
End Property

Public Property Get SkyBottom() As String
    SkyBottom = pSkyBottom
End Property
Public Property Let SkyBottom(ByVal RHS As String)
    pSkyBottom = RHS
    Set SkySkin(5) = LoadTexture(AppPath & "Models\" & RHS)
End Property

Public Property Get SkyFront() As String
    SkyFront = pSkyFront
End Property
Public Property Let SkyFront(ByVal RHS As String)
    pSkyFront = RHS
    Set SkySkin(3) = LoadTexture(AppPath & "Models\" & RHS)
End Property

Public Property Get SkyBack() As String
    SkyBack = pSkyBack
End Property
Public Property Let SkyBack(ByVal RHS As String)
    pSkyBack = RHS
    Set SkySkin(1) = LoadTexture(AppPath & "Models\" & RHS)
End Property

Public Property Get SkyLeft() As String
    SkyLeft = pSkyLeft
End Property
Public Property Let SkyLeft(ByVal RHS As String)
    pSkyLeft = RHS
    Set SkySkin(2) = LoadTexture(AppPath & "Models\" & RHS)
End Property

Public Property Get SkyRight() As String
    SkyRight = pSkyRight
End Property
Public Property Let SkyRight(ByVal RHS As String)
    pSkyRight = RHS
    Set SkySkin(4) = LoadTexture(AppPath & "Models\" & RHS)
End Property


Private Sub Class_Initialize()
    Stats_Space_Count = Stats_Space_Count + 1
    Key = "K" & ObjPtr(Me)
    
    If Spaces.Count = 0 Then
        pVisible = True
    End If
    
    DebugMode = False
    Perspective = ThirdPerson
    
    Set pOrigin = New Point
    
    Range = -1
    
    SkyRotated = 0
    
    Boundary = 0
    
    FogColor = "[1, 1.385, 1.275, 1.133]"
    
    pFogDistance = FadeDistance
    
    Gravity = -0.2
    
    ReDim SkyPlaq(0 To 35) As MyVertex
    ReDim SkySkin(0 To 5) As Direct3DTexture8

    CreateSquare SkyPlaq, 0, MakeVector(-5, -5, 5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(-5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 6, MakeVector(-5, -5, -5), _
                            MakeVector(5, -5, -5), _
                            MakeVector(5, 5, -5), _
                            MakeVector(-5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 12, MakeVector(5, -5, -5), _
                            MakeVector(5, -5, 5), _
                            MakeVector(5, 5, 5), _
                            MakeVector(5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 18, MakeVector(5, -5, -5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(5, -5, 5), 1, 1
    CreateSquare SkyPlaq, 24, MakeVector(5, -5, 5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 30, MakeVector(5, 5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(5, 5, -5), 1, 1

    Set SkyVBuf = DDevice.CreateVertexBuffer(Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SkyVBuf, 0, Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, SkyPlaq(0)
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next

    
    Set pOrigin = Nothing
    
    Set SkyVBuf = Nothing
    
    Set SkySkin(0) = Nothing
    Set SkySkin(1) = Nothing
    Set SkySkin(2) = Nothing
    Set SkySkin(3) = Nothing
    Set SkySkin(4) = Nothing
    Set SkySkin(5) = Nothing
    
    Erase SkyPlaq
    Erase SkySkin
    Stats_Space_Count = Stats_Space_Count - 1
End Sub
