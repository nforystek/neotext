VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Light"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Lighting
    Omni = 1
    Spot = 2
    Directed = 3
End Enum


Private pEnabled As Boolean
Private pSunLight As Boolean
'
Private pLightBlink As Single
Private pLightTimer As Single
Private pLightIsOn As Boolean
'
Private pDiffuseRoll As Single
Private pDiffuseTimer As Single
Private pDIffuseMax As Single
Private pDiffuseNow As Single
Public LightIndex As Long
Attribute LightIndex.VB_VarMemberFlags = "40"

Private pKey As String
Private pSerialize As Boolean

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_MemberFlags = "200"
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

Public Property Get SunLight() As Boolean
    SunLight = pSunLight
End Property
Public Property Let SunLight(ByVal RHS As Boolean)
    pSunLight = RHS
End Property

Public Property Get LightBlink() As Single
    LightBlink = pLightBlink
End Property
Public Property Let LightBlink(ByVal RHS As Single)
    pLightBlink = RHS
End Property

Friend Property Get LightTimer() As Single
    LightTimer = pLightTimer
End Property
Friend Property Let LightTimer(ByVal RHS As Single)
    pLightTimer = RHS
End Property

Public Property Get LightIsOn() As Boolean
Attribute LightIsOn.VB_MemberFlags = "40"
    LightIsOn = pLightIsOn
End Property
Friend Property Let LightIsOn(ByVal RHS As Boolean)
    pLightIsOn = RHS
End Property

Public Property Get DiffuseRoll() As Single
    DiffuseRoll = pDiffuseRoll
End Property
Public Property Let DiffuseRoll(ByVal RHS As Single)
    pDiffuseRoll = RHS
End Property

Public Property Get DiffuseTimer() As Single
Attribute DiffuseTimer.VB_MemberFlags = "40"
    DiffuseTimer = pDiffuseTimer
End Property
Public Property Let DiffuseTimer(ByVal RHS As Single)
    pDiffuseTimer = RHS
End Property

Public Property Get DIffuseMax() As Single
    DIffuseMax = pDIffuseMax
End Property
Public Property Let DIffuseMax(ByVal RHS As Single)
    pDIffuseMax = RHS
End Property

Public Property Get DiffuseNow() As Single
Attribute DiffuseNow.VB_MemberFlags = "40"
    DiffuseNow = pDiffuseNow
End Property
Public Property Let DiffuseNow(ByVal RHS As Single)
    pDiffuseNow = RHS
End Property

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal RHS As Boolean)
    pEnabled = RHS
End Property

Public Property Get Origin() As Point
    Set Origin = New Point
    Origin.X = DXLights(LightIndex).Position.X
    Origin.Y = DXLights(LightIndex).Position.Y
    Origin.Z = DXLights(LightIndex).Position.Z
End Property
Public Property Set Origin(ByRef RHS As Point)
    DXLights(LightIndex).Position.X = RHS.X
    DXLights(LightIndex).Position.Y = RHS.Y
    DXLights(LightIndex).Position.Z = RHS.Z
End Property
Public Property Let Origin(ByVal RHS)
    Dim nc As Point
    Set nc = New Point
    nc = RHS
    
    DXLights(LightIndex).Position.X = nc.X
    DXLights(LightIndex).Position.Y = nc.Y
    DXLights(LightIndex).Position.Z = nc.Z
    
    Set nc = Nothing
End Property

Public Property Get Diffuse() As Color
    Set Diffuse = New Color
    Diffuse.Alpha = DXLights(LightIndex).Diffuse.A
    Diffuse.Red = DXLights(LightIndex).Diffuse.r
    Diffuse.Green = DXLights(LightIndex).Diffuse.g
    Diffuse.Blue = DXLights(LightIndex).Diffuse.B
End Property
Public Property Set Diffuse(ByRef RHS As Color)
    DXLights(LightIndex).Diffuse.A = RHS.Alpha
    DXLights(LightIndex).Diffuse.r = RHS.Red
    DXLights(LightIndex).Diffuse.g = RHS.Green
    DXLights(LightIndex).Diffuse.B = RHS.Blue
End Property
Public Property Let Diffuse(ByVal RHS)
    Dim nc As Color
    Set nc = New Color
    nc = RHS
    
    DXLights(LightIndex).Diffuse.A = nc.Alpha
    DXLights(LightIndex).Diffuse.r = nc.Red
    DXLights(LightIndex).Diffuse.g = nc.Green
    DXLights(LightIndex).Diffuse.B = nc.Blue
    
    Set nc = Nothing
End Property

Public Property Get Specular() As Color
    Set Specular = New Color
    Specular.Alpha = DXLights(LightIndex).Specular.A
    Specular.Red = DXLights(LightIndex).Specular.r
    Specular.Green = DXLights(LightIndex).Specular.g
    Specular.Blue = DXLights(LightIndex).Specular.B
End Property
Public Property Set Specular(ByRef RHS As Color)
   
    DXLights(LightIndex).Specular.A = RHS.Alpha
    DXLights(LightIndex).Specular.r = RHS.Red
    DXLights(LightIndex).Specular.g = RHS.Green
    DXLights(LightIndex).Specular.B = RHS.Blue
End Property
Public Property Let Specular(ByVal RHS)
    Dim nc As Color
    Set nc = New Color
    nc = RHS
    
    DXLights(LightIndex).Specular.A = nc.Alpha
    DXLights(LightIndex).Specular.r = nc.Red
    DXLights(LightIndex).Specular.g = nc.Green
    DXLights(LightIndex).Specular.B = nc.Blue
    
    Set nc = Nothing
End Property

Public Property Get Ambient() As Color
    Set Ambient = New Color
    Ambient.Alpha = DXLights(LightIndex).Ambient.A
    Ambient.Red = DXLights(LightIndex).Ambient.r
    Ambient.Green = DXLights(LightIndex).Ambient.g
    Ambient.Blue = DXLights(LightIndex).Ambient.B
End Property
Public Property Set Ambient(ByRef RHS As Color)
    DXLights(LightIndex).Ambient.A = RHS.Alpha
    DXLights(LightIndex).Ambient.r = RHS.Red
    DXLights(LightIndex).Ambient.g = RHS.Green
    DXLights(LightIndex).Ambient.B = RHS.Blue
End Property
Public Property Let Ambient(ByVal RHS)
    Dim nc As Color
    Set nc = New Color
    nc = RHS
    
    DXLights(LightIndex).Ambient.A = nc.Alpha
    DXLights(LightIndex).Ambient.r = nc.Red
    DXLights(LightIndex).Ambient.g = nc.Green
    DXLights(LightIndex).Ambient.B = nc.Blue
    
    Set nc = Nothing
End Property

Public Property Get Attenuation() As Point
    Set Attenuation = New Point
    Attenuation.X = Attenuation0
    Attenuation.Y = Attenuation1
    Attenuation.Z = Attenuation2
End Property
Public Property Set Attenuation(ByRef RHS As Point)
    Attenuation0 = RHS.X
    Attenuation1 = RHS.Y
    Attenuation2 = RHS.Z
End Property
Public Property Let Attenuation(ByVal RHS)
    Dim nc As Point
    Set nc = New Point
    nc = RHS
    
    Attenuation0 = nc.X
    Attenuation1 = nc.Y
    Attenuation2 = nc.Z
    
    Set nc = Nothing
End Property

Public Property Get Attenuation0() As Single
    Attenuation0 = DXLights(LightIndex).Attenuation0
End Property
Public Property Let Attenuation0(ByVal RHS As Single)
    DXLights(LightIndex).Attenuation0 = RHS
End Property

Public Property Get Attenuation1() As Single
    Attenuation1 = DXLights(LightIndex).Attenuation1
End Property
Public Property Let Attenuation1(ByVal RHS As Single)
    DXLights(LightIndex).Attenuation1 = RHS
End Property

Public Property Get Attenuation2() As Single
    Attenuation2 = DXLights(LightIndex).Attenuation2
End Property
Public Property Let Attenuation2(ByVal RHS As Single)
    DXLights(LightIndex).Attenuation2 = RHS
End Property

Public Property Get Phi() As Single
    Phi = DXLights(LightIndex).Phi
End Property
Public Property Let Phi(ByVal RHS As Single)
    DXLights(LightIndex).Phi = RHS
End Property

Public Property Get Theta() As Single
    Theta = DXLights(LightIndex).Theta
End Property
Public Property Let Theta(ByVal RHS As Single)
    DXLights(LightIndex).Theta = RHS
End Property

Public Property Get Range() As Single
    Range = DXLights(LightIndex).Range
End Property
Public Property Let Range(ByVal RHS As Single)
    DXLights(LightIndex).Range = RHS
End Property


Public Property Get FallOff() As Single
    FallOff = DXLights(LightIndex).FallOff
End Property
Public Property Let FallOff(ByVal RHS As Single)
    DXLights(LightIndex).FallOff = RHS
End Property


Public Property Get Direction() As Point
    Set Direction = New Point
    Direction.X = DXLights(LightIndex).Direction.X
    Direction.Y = DXLights(LightIndex).Direction.Y
    Direction.Z = DXLights(LightIndex).Direction.Z
End Property
Public Property Set Direction(ByRef RHS As Point)
    DXLights(LightIndex).Direction.X = RHS.X
    DXLights(LightIndex).Direction.Y = RHS.Y
    DXLights(LightIndex).Direction.Z = RHS.Z
End Property

Public Property Let Direction(ByVal RHS)
    Dim nc As Point
    Set nc = New Point
    nc = RHS
    
    DXLights(LightIndex).Direction.X = nc.X
    DXLights(LightIndex).Direction.Y = nc.Y
    DXLights(LightIndex).Direction.Z = nc.Z
    
    Set nc = Nothing
End Property

Public Property Get LightType() As Lighting
    Select Case DXLights(LightIndex).Type
        Case D3DLIGHT_DIRECTIONAL
            LightType = Lighting.Directed
        Case D3DLIGHT_POINT
            LightType = Lighting.Omni
        Case D3DLIGHT_SPOT
            LightType = Lighting.Spot
    End Select
End Property
Public Property Let LightType(ByVal RHS As Lighting)
    Select Case RHS
        Case Lighting.Directed
            DXLights(LightIndex).Type = D3DLIGHT_DIRECTIONAL
        Case Lighting.Omni
            DXLights(LightIndex).Type = D3DLIGHT_POINT
        Case Lighting.Spot
            DXLights(LightIndex).Type = D3DLIGHT_SPOT
    End Select
End Property

Private Sub Class_Initialize()
    Stats_Light_Count = Stats_Light_Count + 1

        
    LightIndex = Lights.Count + 1
    ReDim Preserve DXLights(1 To LightIndex) As D3DLIGHT8

    
    Key = "K" & ObjPtr(Me)
    
End Sub

Private Sub Class_Terminate()
    Stats_Light_Count = Stats_Light_Count - 1
    
'    If Stats_Light_Count = 0 Then
'        Erase DXLights
'
'        Erase LightDatas
'        LightDataCount = 0
'
'    End If
End Sub
