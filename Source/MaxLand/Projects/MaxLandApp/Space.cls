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

    
Private pSkyTop As String
Private pSkyBottom As String
Private pSkyFront As String
Private pSkyBack As String
Private pSkyRight As String
Private pSkyLeft As String

Private pOrigin As Point

Private pRange As Single

Private pGravity As Single

Private pFogEnabled As Boolean
Private pFogDistance As Single
Private pFogColor As New Color

Private pSkyRotate As Single
Private pSkyRotated As Single

Private pBoundary As Single
Private pVisible As Boolean

Private pSerialize As Boolean
Private pKey As String

Private SkyPlaq() As MyVertex
Private SkySkin() As Direct3DTexture8
Private SkyVBuf As Direct3DVertexBuffer8

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

Friend Function InSpace(ByRef WayPoint As Point) As Boolean
    InSpace = (Visible And (Range = -1) Or (DistanceEx(WayPoint, Origin) < Range))
End Function

Friend Sub RenderSpace()
    If InSpace(Player.Element.Origin) Then
    
        Dim matProj As D3DMATRIX
        Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
        DDevice.GetTransform D3DTS_VIEW, matViewSave
        matView = matViewSave
        matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
    
        DDevice.SetTransform D3DTS_VIEW, matView
    
        D3DXMatrixPerspectiveFovLH matProj, PI / 2.5, AspectRatio, 0.05, 50
        DDevice.SetTransform D3DTS_PROJECTION, matProj
    
        If Not SkyRotated = 0 Then
        
            If (FPSRate > 0) And SkyRotate > 0 Then

                SkyRotated = SkyRotated + (360 / (HoursInOneDay * SkyRotate)) * _
                    (PI / (HoursInOneDay * FPSRate)) * (FPSRate / HoursInOneDay)
            End If
            
            D3DXMatrixRotationY matWorld, SkyRotated * RADIAN
        End If
        DDevice.SetTransform D3DTS_WORLD, matWorld
                        
        DDevice.SetStreamSource 0, SkyVBuf, Len(SkyPlaq(0))
    
        DDevice.SetTexture 1, Nothing
        
        DDevice.SetTexture 0, SkySkin(0)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 30, 2
        DDevice.SetTexture 0, SkySkin(1)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        DDevice.SetTexture 0, SkySkin(2)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
        DDevice.SetTexture 0, SkySkin(3)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
        DDevice.SetTexture 0, SkySkin(4)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 24, 2
        DDevice.SetTexture 0, SkySkin(5)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 18, 2
        
        D3DXMatrixPerspectiveFovLH matProj, FOVY, AspectRatio, 0.05, FadeDistance
        DDevice.SetTransform D3DTS_PROJECTION, matProj
    
        DDevice.SetTransform D3DTS_VIEW, matViewSave
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
    ToString = ToString & Include.URLEncode("Space.FogColor = """ & pFogColor & """") & vbCrLf
    ToString = ToString & Include.URLEncode("Space.FogEnabled = " & pFogEnabled) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.FogDistance = " & pFogDistance) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.SkyRotate = " & pSkyRotate) & vbCrLf
    ToString = ToString & Include.URLEncode("Space.Gravity = " & pGravity) & vbCrLf


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
