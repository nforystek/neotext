Attribute VB_Name = "modLand"
#Const modLand = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Enum Actions
    None = 0
    Directing = 1
    Rotating = 2
    Scaling = 4
    Script = 8
End Enum

Public Enum Collides
    None = 0
    InDoor = 1
    Ground = 2
    Liquid = 3
    Ladder = 4
End Enum

Public Enum Moving
    None = 0
    Level = 1
    Flying = 2
    Falling = 4
    Stepping = 8
End Enum

Public Type MyStates
    IsMoving As Moving
    InLiquid As Boolean
    OnLadder As Boolean
End Type

Public Type MyActivity
    Identity As String
    Action As Actions

    Data As D3DVECTOR
    OnEvent As String
    
    Reactive As Single
    Latency As Single
    Recount As Single

    Emphasis As Single
    Friction As Single
    Initials As Single
End Type

Public Type MyBoard
    Skin() As Direct3DTexture8
    SkinCount As Long
    
    Plaq() As MyVertex
    VBuf As Direct3DVertexBuffer8
    
    Center As D3DVECTOR
    
    Point1 As MyVertex
    Point2 As MyVertex
    Point3 As MyVertex
    Point4 As MyVertex
        
    ScaleX As Single
    ScaleY As Single
    
    Translucent As Boolean
    BlackAlpha As Boolean
    
    AnimateMSecs As Single
    AnimateTimer As Double
    AnimatePoint As Long
    
    Visible As Boolean
    Identity As String
End Type

Public Type MyMesh
    Mesh As D3DXMesh

    Materials() As D3DMATERIAL8
    Textures() As Direct3DTexture8
    
    Verticies() As D3DVERTEX
    Indicies() As Integer
    VBuffer As Direct3DVertexBuffer8
    
    MaterialBuffer As D3DXBuffer
    MaterialCount As Long
    FileName As String
End Type

Public Type MyObject

    Identity As String
    Visible As Boolean
    
    MeshIndex As Long
    VisualIndex As Long
    
    CollideIndex As Long
    CollideObject As Long
    CollideFaces As Long

    CulledFaces As Long
    BlackAlpha As Boolean
    WireFrame As Boolean
    
    Origin As D3DVECTOR
    Direct As D3DVECTOR
    Offset As D3DVECTOR
    
    Folcrum() As D3DVECTOR
    FolcrumCount As Long
    
    Rotate As D3DVECTOR
    Twists As D3DVECTOR
    
    Scaled As D3DVECTOR
    Scalar As D3DVECTOR

    Matrix As D3DMATRIX
    States As MyStates
    Effect As Collides
    
    Activities() As MyActivity
    ActivityCount As Long
    Gravitational As Boolean

    ReplacerVals As New Collection
    ReplacerKeys As New Collection
End Type

Public Type MyLight
    
    Identity As String
    Enabled As Boolean
    SunLight As Boolean
    
    Origin As D3DVECTOR
    
    LightBlink As Single
    LightTimer As Single
    LightIsOn As Boolean
    
    DiffuseRoll As Single
    DiffuseTimer As Single
    DIffuseMax As Single
    DiffuseNow As Single
    
    LightIndex As Long
End Type

Public Type MyImage
    Identity As String
    Visible As Boolean
    
    Image As Direct3DTexture8
    Verticies(0 To 4) As MyScreen
    
    Translucent As Boolean
    BlackAlpha As Boolean
    Dimension As ImgDimType
    
    Padding As Long
End Type

Public Type MyBeacon
    Identity As String
    Visible As Boolean
    
    Origins() As D3DVECTOR
    OriginCount As Long
    
    BeaconSkin() As Direct3DTexture8
    BeaconSkinCount As Long

    BeaconPlaq(0 To 5) As MyVertex
    BeaconVBuf As Direct3DVertexBuffer8

    Dimension As ImgDimType
    PercentXY As ImgDimType
    
    HorizontalLock As Boolean
    VerticalLock As Boolean
    VerticalSkew As Boolean
    RoundingCut As Integer
    
    Translucent As Boolean
    BlackAlpha As Boolean
    
    Consumable As Boolean
    Randomize As Boolean
    Allowance As Long
    
    BeaconAnim As Double
    BeaconText As Long
    BeaconLight As Long
End Type

Public Type MyPlayer
    CameraAngle As Single
    CameraPitch As Single
    CameraZoom As Single
    CameraIndex As Long
    MoveSpeed As Single
    AutoMove As Boolean
    Object As MyObject
    Boundary As Single
End Type

Public Type MyPortal
    Enable As Boolean
    Identity As String
    OnInRange As String
    OnOutRange As String
    
    Location As D3DVECTOR
    Teleport As D3DVECTOR
    Range As Single
    
    Activities() As MyActivity
    ActivityCount As Long
    ClearActivities As Boolean
End Type

Public Type MyCamera
    Identity As String
    
    Location As D3DVECTOR
    Angle As Single
    Pitch As Single
    ModAngle As Single
End Type

Public Type MyVariable
    Identity As String
    Value As Variant
    OnEdit As String
End Type

Public Type MyMethod
    Identity As String
    Script As String
End Type

Public Type MySound
    Identity As String
    Enable As Boolean
    Repeat As Boolean
    Origin As D3DVECTOR
    Range As Single
    Index As Long
End Type

Public Tracks() As clsAmbient
Public TrackCount As Long

Public Sounds() As MySound
Public SoundCount As Long

Public ScreenImages() As MyImage
Public ScreenImageCount As Long

Public Lights() As D3DLIGHT8
Public LightCount As Long

Public LightDatas() As MyLight
Public LightDataCount As Long

Public Meshes() As MyMesh
Public MeshCount As Long

Public Objects() As MyObject
Public ObjectCount As Long

Public BillBoards() As MyBoard
Public BillBoardCount As Long

Public Beacons() As MyBeacon
Public BeaconCount As Long

Public Portals() As MyPortal
Public PortalCount As Long

Public Cameras() As MyCamera
Public CameraCount As Long

Public Variables() As MyVariable
Public VariableCount As Long

Public Methods() As MyMethod
Public MethodCount As Long

Private SkyPlaq(0 To 35) As MyVertex
Private SkySkin(0 To 5) As Direct3DTexture8
Private SkyVBuf As Direct3DVertexBuffer8
Private SkyCRot As Single

Private Serialize As String
Private Deserialize As String

Public GlobalGravityDirect As MyActivity
Public GlobalGravityRotate As MyActivity
Public GlobalGravityScaled As MyActivity

Public LiquidGravityDirect As MyActivity
Public LiquidGravityRotate As MyActivity
Public LiquidGravityScaled As MyActivity

Private CloudRotated As Single

Public matWorld As D3DMATRIX

Public Function GetBoardIndex(ByRef Obj As MyObject, ByVal TextName As String) As Long
    If Obj.ReplacerKeys.Count > 0 Then
        Dim i As Long
        For i = 1 To Obj.ReplacerKeys.Count
            If Obj.ReplacerKeys(i) = Obj.Identity & "_" & Replace(TextName, ".", "") Then
                GetBoardIndex = Obj.ReplacerVals(Obj.Identity & "_" & Replace(TextName, ".", ""))
                Exit Function
            End If
        Next
    End If
End Function

Public Sub RenderPlayer()

    If ((Perspective = Playmode.ThirdPerson) Or (Perspective = Playmode.CameraMode)) And (Not DebugMode) Then
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
        DDevice.SetVertexShader FVF_RENDER
        
        D3DXMatrixIdentity matWorld
        D3DXMatrixIdentity Player.Object.Matrix
        
        D3DXMatrixScaling Player.Object.Matrix, Player.Object.Scaled.X, Player.Object.Scaled.Y, Player.Object.Scaled.z
        D3DXMatrixMultiply matWorld, matWorld, Player.Object.Matrix
        
        D3DXMatrixRotationY Player.Object.Matrix, -Player.CameraAngle
        D3DXMatrixMultiply matWorld, matWorld, Player.Object.Matrix
        
        D3DXMatrixTranslation Player.Object.Matrix, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z
        D3DXMatrixMultiply matWorld, matWorld, Player.Object.Matrix

        DDevice.SetTransform D3DTS_WORLD, matWorld
        
        If Player.Object.Visible Then
            DDevice.SetRenderState D3DRS_FILLMODE, IIf(Player.Object.WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
            
            Dim i As Long
            If Meshes(Player.Object.VisualIndex).MaterialCount > 0 Then
                For i = 0 To Meshes(Player.Object.VisualIndex).MaterialCount - 1
    
                    If Meshes(Player.Object.VisualIndex).Textures(i) Is Nothing Then
                        DDevice.SetPixelShader PixelShaderDefault
                        DDevice.SetMaterial Meshes(Player.Object.VisualIndex).Materials(i)
                        DDevice.SetTexture 0, Nothing
                        DDevice.SetMaterial GenericMaterial
                        DDevice.SetTexture 1, Nothing
                        Meshes(Player.Object.VisualIndex).Mesh.DrawSubset i
                    Else
    
                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
                        DDevice.SetPixelShader PixelShaderDefault
                        DDevice.SetMaterial Meshes(Player.Object.VisualIndex).Materials(i)
                        DDevice.SetTexture 0, Meshes(Player.Object.VisualIndex).Textures(i)
                        DDevice.SetMaterial GenericMaterial
                        DDevice.SetTexture 1, Meshes(Player.Object.VisualIndex).Textures(i)
                        Meshes(Player.Object.VisualIndex).Mesh.DrawSubset i
    
                    End If
    
                Next
            End If
        
            DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        End If
    End If
    
End Sub

Public Sub RenderWorld()
    Dim Dist As Long
    Dim fogVal As Long
    Dim t As Boolean
    Dim o As Long
    Dim i As Long
    Dim l As Long
    Dim v As Long
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim vv(0 To 2) As D3DVECTOR
    Dim r As Single

    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
                        
    If LightDataCount > 0 Then
        For l = 1 To LightDataCount
            If Lights(LightDatas(l).LightIndex).Type = D3DLIGHT_DIRECTIONAL Or (LightDatas(l).Enabled And _
                Distance(Player.Object.Origin, LightDatas(l).Origin) <= (FadeDistance - Lights(LightDatas(l).LightIndex).Range)) Then
                
                If (LightDatas(l).LightBlink > 0) Or (LightDatas(l).DiffuseRoll <> 0) Then
                    If (LightDatas(l).LightBlink > 0) Then
                        If (LightDatas(l).LightTimer = 0) Or ((Timer - LightDatas(l).LightTimer) >= LightDatas(l).LightBlink And (LightDatas(l).LightBlink > 0)) Then
                            LightDatas(l).LightTimer = Timer
                            LightDatas(l).LightIsOn = Not LightDatas(l).LightIsOn
                        End If
                        DDevice.LightEnable (l - 1), LightDatas(l).LightIsOn
                    End If
                    If (LightDatas(l).DiffuseRoll <> 0) Then
                        If (LightDatas(l).DiffuseTimer = 0) Or ((Timer - LightDatas(l).DiffuseTimer) >= Abs(LightDatas(l).DiffuseRoll) And (LightDatas(l).DiffuseTimer > 0)) Then
                            LightDatas(l).DiffuseTimer = Timer
                            If (LightDatas(l).DiffuseRoll > 0) Then
                                If (LightDatas(l).DIffuseMax > 0 And LightDatas(l).DiffuseNow >= LightDatas(l).DIffuseMax) Or (LightDatas(l).DIffuseMax < 0 And LightDatas(l).DiffuseNow >= -0.01) Then
                                    LightDatas(l).DiffuseRoll = -LightDatas(l).DiffuseRoll
                                Else
                                    LightDatas(l).DiffuseNow = LightDatas(l).DiffuseNow + 1
                                    Lights(LightDatas(l).LightIndex).diffuse.r = Lights(LightDatas(l).LightIndex).diffuse.r + 0.01
                                    Lights(LightDatas(l).LightIndex).diffuse.g = Lights(LightDatas(l).LightIndex).diffuse.g + 0.01
                                    Lights(LightDatas(l).LightIndex).diffuse.b = Lights(LightDatas(l).LightIndex).diffuse.b + 0.01
                                End If
                            Else
                                If (LightDatas(l).DIffuseMax > 0 And LightDatas(l).DiffuseNow <= 0.01) Or (LightDatas(l).DIffuseMax < 0 And LightDatas(l).DiffuseNow <= LightDatas(l).DIffuseMax) Then
                                    LightDatas(l).DiffuseRoll = -LightDatas(l).DiffuseRoll
                                Else
                                    LightDatas(l).DiffuseNow = LightDatas(l).DiffuseNow - 1
                                    Lights(LightDatas(l).LightIndex).diffuse.r = Lights(LightDatas(l).LightIndex).diffuse.r - 0.01
                                    Lights(LightDatas(l).LightIndex).diffuse.g = Lights(LightDatas(l).LightIndex).diffuse.g - 0.01
                                    Lights(LightDatas(l).LightIndex).diffuse.b = Lights(LightDatas(l).LightIndex).diffuse.b - 0.01
                                End If
                            End If
                            
                            
                        End If

                        DDevice.SetLight (l - 1), Lights(LightDatas(l).LightIndex)
                        DDevice.LightEnable (l - 1), 1
                    End If
                Else
                    DDevice.LightEnable l - 1, 1
                End If
            Else
                DDevice.LightEnable l - 1, False
            End If
        Next
    End If

    If SoundCount > 0 Then
        For l = 1 To SoundCount
            If Sounds(l).Range > 0 Then
                r = Distance(Player.Object.Origin, Sounds(l).Origin)
                If r < Sounds(l).Range Then
                    
                    r = Round(CSng(Sounds(Index).Range - Dist), 3)
                    r = Abs(-Sounds(Index).Range + r)
                
                    VolumeWave l, r
                    PlayWave l, Sounds(l).Repeat
                    
                Else
                    StopWave l
                    
                End If
            End If
        Next
    End If
    
    If TrackCount > 0 Then
        For l = 1 To TrackCount
            If Tracks(l).Range > 0 Then
                r = Distance(Player.Object.Origin, Tracks(l).Origin)
                If r < Tracks(l).Range Then
                
                    r = Round(CSng(Tracks(l).Range - r), 3)
                    'r = Abs(-Tracks(l).Range + r)
        
                    Tracks(l).TrackVolume = (r * 10)
                    
                Else
                    Tracks(l).TrackVolume = 0
                End If
            End If
        Next
    End If
    
    If ObjectCount > 0 Then
    
        
        For o = 1 To ObjectCount
            
            If Objects(o).Visible And (Not (Objects(o).Effect = Ladder Or Objects(o).Effect = Liquid)) Then
            
                If ((Objects(o).MeshIndex >= 0) And Distance(Player.Object.Origin, Objects(o).Origin) <= FadeDistance) Then
                    
                    If Objects(o).VisualIndex > 0 Then
                        v = Objects(o).VisualIndex
                    Else
                        v = Objects(o).MeshIndex
                    End If

                    If Objects(o).MeshIndex > 0 Then
                    
                    If DebugMode Or Meshes(Objects(o).MeshIndex).MaterialCount > 0 Then

                        Dim matMesh As D3DMATRIX
                        

                        D3DXMatrixIdentity matMesh
                        D3DXMatrixRotationX Objects(o).Matrix, Objects(o).Rotate.X * (PI / 180)
                        D3DXMatrixMultiply Objects(o).Matrix, matMesh, Objects(o).Matrix
                        D3DXMatrixMultiply matMesh, Objects(o).Matrix, matMesh
                        D3DXMatrixRotationY Objects(o).Matrix, Objects(o).Rotate.Y * (PI / 180)
                        D3DXMatrixMultiply Objects(o).Matrix, matMesh, Objects(o).Matrix
                        D3DXMatrixMultiply matMesh, Objects(o).Matrix, matMesh
                        D3DXMatrixRotationZ Objects(o).Matrix, Objects(o).Rotate.z * (PI / 180)
                        D3DXMatrixMultiply Objects(o).Matrix, matMesh, Objects(o).Matrix
                        D3DXMatrixMultiply matMesh, Objects(o).Matrix, matMesh
                        D3DXMatrixTranslation Objects(o).Matrix, Objects(o).Origin.X, Objects(o).Origin.Y, Objects(o).Origin.z
                        D3DXMatrixMultiply Objects(o).Matrix, matMesh, Objects(o).Matrix
                        DDevice.SetTransform D3DTS_WORLD, Objects(o).Matrix
                        D3DXMatrixScaling matMesh, Objects(o).Scaled.X, Objects(o).Scaled.Y, Objects(o).Scaled.z
                        D3DXMatrixMultiply matMesh, matMesh, Objects(o).Matrix
                        DDevice.SetTransform D3DTS_WORLD, matMesh

                    End If
                End If
                        
'                    If DebugMode Then
'                        Index = 0
'                        For Face = 0 To Meshes(Objects(o).MeshIndex).Mesh.GetNumFaces - 1
'                            ReDim Preserve DebugFace(0 To Index + 2) As MyVertex
'
'                            vv(0) = MakeVector(sngVertexX(0, Face), sngVertexY(0, Face), sngVertexZ(0, Face))
'                            D3DXVec3TransformCoord vv(0), vv(0), Objects(o).Matrix
'
'                            DebugFace(Index + 0).X = vv(0).X
'                            DebugFace(Index + 0).Y = vv(0).Y
'                            DebugFace(Index + 0).z = vv(0).z
'                            DebugFace(Index + 0).nx = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 0)).nx
'                            DebugFace(Index + 0).ny = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 0)).ny
'                            DebugFace(Index + 0).nz = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 0)).nz
'                            DebugFace(Index + 0).tu = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 0)).tu
'                            DebugFace(Index + 0).tv = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 0)).tv
'
'                            vv(1) = MakeVector(sngVertexX(1, Face), sngVertexY(1, Face), sngVertexZ(1, Face))
'                            D3DXVec3TransformCoord vv(1), vv(1), Objects(o).Matrix
'
'                            DebugFace(Index + 1).X = vv(1).X
'                            DebugFace(Index + 1).Y = vv(1).Y
'                            DebugFace(Index + 1).z = vv(1).z
'                            DebugFace(Index + 1).nx = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 1)).nx
'                            DebugFace(Index + 1).ny = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 1)).ny
'                            DebugFace(Index + 1).nz = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 1)).nz
'                            DebugFace(Index + 1).tu = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 1)).tu
'                            DebugFace(Index + 1).tv = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 1)).tv
'
'                            vv(2) = MakeVector(sngVertexX(2, Face), sngVertexY(2, Face), sngVertexZ(2, Face))
'                            D3DXVec3TransformCoord vv(2), vv(2), Objects(o).Matrix
'
'                            DebugFace(Index + 2).X = vv(2).X
'                            DebugFace(Index + 2).Y = vv(2).Y
'                            DebugFace(Index + 2).z = vv(2).z
'                            DebugFace(Index + 2).nx = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 2)).nx
'                            DebugFace(Index + 2).ny = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 2)).ny
'                            DebugFace(Index + 2).nz = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 2)).nz
'                            DebugFace(Index + 2).tu = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 2)).tu
'                            DebugFace(Index + 2).tv = Meshes(Objects(o).MeshIndex).Verticies(Meshes(Objects(o).MeshIndex).Indicies(Index + 2)).tv
'
'                            Index = Index + 3
'                        Next
'
'                        Set DebugVBuf = DDevice.CreateVertexBuffer(Len(DebugFace(0)) * (UBound(DebugFace) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
'                        D3DVertexBuffer8SetData DebugVBuf, 0, Len(DebugFace(0)) * (UBound(DebugFace) + 1), 0, DebugFace(0)
'
'                        DDevice.SetStreamSource 0, DebugVBuf, Len(DebugFace(0))
'
'                        Index = 0
'                        For Face = 0 To lngFaceCount - 1
'
'                            If sngFaceVis(3, Face) = 0 Then
'
'                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'                            Else
'                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
'                            End If
'
'                            DDevice.SetMaterial GenericMaterial
'                            Select Case sngFaceVis(3, Face)
'                                Case CULL0
'                                    DDevice.SetTexture 0, DebugSkin(0)
'                                Case CULL1
'                                    DDevice.SetTexture 0, DebugSkin(1)
'                                Case CULL2
'                                    DDevice.SetTexture 0, DebugSkin(2)
'                                Case CULL3
'                                    DDevice.SetTexture 0, DebugSkin(3)
'                                Case CULL4
'                                    DDevice.SetTexture 0, DebugSkin(4)
'                                Case CULL5
'                                    DDevice.SetTexture 0, DebugSkin(3)
'                                Case CULL6
'                                    DDevice.SetTexture 0, DebugSkin(4)
'                            End Select
'
'                            DDevice.SetMaterial GenericMaterial
'                            Select Case sngFaceVis(3, Face)
'                                Case CULL0
'                                    DDevice.SetTexture 1, DebugSkin(0)
'                                Case CULL1
'                                    DDevice.SetTexture 1, DebugSkin(1)
'                                Case CULL2
'                                    DDevice.SetTexture 1, DebugSkin(2)
'                                Case CULL3
'                                    DDevice.SetTexture 1, DebugSkin(3)
'                                Case CULL4
'                                    DDevice.SetTexture 1, DebugSkin(4)
'                                Case CULL5
'                                    DDevice.SetTexture 1, DebugSkin(3)
'                                Case CULL6
'                                    DDevice.SetTexture 1, DebugSkin(4)
'                            End Select
'                            DDevice.SetTexture 1, DebugSkin(sngFaceVis(3, Face))
'
'                            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, Index, 1
'                            Index = Index + 3
'                        Next
'                    Else
                    
                        If Meshes(v).MaterialCount > 0 Then
                            
                            DDevice.SetRenderState D3DRS_FILLMODE, IIf(Objects(o).WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
                            
                            For i = 0 To Meshes(v).MaterialCount - 1
                                     
                                If Objects(o).BlackAlpha Then
                                            
                                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
                                    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        
                                    DDevice.SetMaterial GenericMaterial
                                     
                                    DDevice.SetTexture 0, Meshes(v).Textures(i)
                                    DDevice.SetMaterial GenericMaterial
                                    DDevice.SetTexture 1, Meshes(v).Textures(i)
                                 
                                    Meshes(v).Mesh.DrawSubset i
                                                 
                                 Else
                                     If Not (Meshes(v).Textures(i) Is Nothing) Then
                                         
                                         l = GetBoardIndex(Objects(o), D3DX.BufferGetTextureName(Meshes(v).MaterialBuffer, i))
                                         
                                         If (l > 0) Then
                                    
                                             If Not BillBoards(l).Translucent And Not BillBoards(l).BlackAlpha Then
        
                                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
                                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
                                                
                                                If (BillBoards(l).AnimateMSecs > 0) Then
                                                    If (BillBoards(l).AnimateTimer = 0) Or (CDbl(Timer - BillBoards(l).AnimateTimer) >= BillBoards(l).AnimateMSecs) Then
                                                        BillBoards(l).AnimateTimer = GetTimer
                                                         
                                                        BillBoards(l).AnimatePoint = BillBoards(l).AnimatePoint + 1
                                                        If BillBoards(l).AnimatePoint > BillBoards(l).SkinCount Then
                                                            BillBoards(l).AnimatePoint = 1
                                                        End If
                                                        
                                                    End If
                                
                                                    DDevice.SetMaterial LucentMaterial
                                                    DDevice.SetTexture 0, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                                                    DDevice.SetMaterial GenericMaterial
                                                    DDevice.SetTexture 1, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                                
                                                Else
                                                    DDevice.SetMaterial LucentMaterial
                                                    DDevice.SetTexture 0, BillBoards(l).Skin(1)
                                                    DDevice.SetMaterial GenericMaterial
                                                    DDevice.SetTexture 1, BillBoards(l).Skin(1)
                                                End If
        
                                                Meshes(v).Mesh.DrawSubset i
                                            End If
                                         Else
                                            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                                            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                             
                                            DDevice.SetMaterial Meshes(v).Materials(i)
                                            DDevice.SetTexture 0, Meshes(v).Textures(i)
                                            DDevice.SetMaterial GenericMaterial
                                            DDevice.SetTexture 1, Meshes(v).Textures(i)
                                            Meshes(v).Mesh.DrawSubset i
                                             
                                        End If
                                         
                                    End If
                                End If
                            Next
                        End If
'                    End If
                End If
            End If
        Next
    End If
End Sub

Public Sub RenderBoards()

    Dim o As Long
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

    Dim matScale As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matBoards As D3DMATRIX

    Dim start As Single
    Dim chop As Single
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld

    If BillBoardCount > 0 Then
        For o = 1 To BillBoardCount
            If BillBoards(o).Visible Then
                If Not (BillBoards(o).VBuf Is Nothing) Then
                    If Distance(Player.Object.Origin, BillBoards(o).Center) <= FadeDistance Then
                        
                        If Not BillBoards(o).Translucent Then
                        
                            If (BillBoards(o).AnimateMSecs > 0) Then
                                If CDbl(Timer - BillBoards(o).AnimateTimer) >= BillBoards(o).AnimateMSecs Then
                                    BillBoards(o).AnimateTimer = GetTimer
                                    
                                    BillBoards(o).AnimatePoint = BillBoards(o).AnimatePoint + 1
                                    If BillBoards(o).AnimatePoint > BillBoards(o).SkinCount Then
                                        BillBoards(o).AnimatePoint = 1
                                    End If
                                    
                                End If
                                DDevice.SetMaterial GenericMaterial
                                DDevice.SetTexture 0, BillBoards(o).Skin(BillBoards(o).AnimatePoint)
                                DDevice.SetTexture 1, Nothing
                            Else
                                DDevice.SetMaterial GenericMaterial
                                DDevice.SetTexture 0, BillBoards(o).Skin(1)
                                DDevice.SetTexture 1, Nothing
                            End If
                            
                            DDevice.SetStreamSource 0, BillBoards(o).VBuf, Len(BillBoards(o).Plaq(0))
                            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                        End If

                        
                    End If
                End If
            End If
        Next
    End If
    
End Sub

Public Sub RenderLucent()
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDiffuse
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                       
    Dim p As D3DVECTOR
    Dim o As Long
    Dim fogVal As Long
    Dim t As Boolean
    Dim i As Long
    Dim l As Long
    
    D3DXMatrixIdentity matWorld
        
    If ObjectCount > 0 Then

        For o = 1 To ObjectCount
            If Objects(o).Visible And (Not DebugMode) Then
            
                If (Objects(o).MeshIndex > 0) And Distance(Player.Object.Origin, Objects(o).Origin) <= FadeDistance Then
    
                    For i = 0 To Meshes(Objects(o).MeshIndex).MaterialCount - 1
    
                        If Not (Meshes(Objects(o).MeshIndex).Textures(i) Is Nothing) Then
                            
                            l = GetBoardIndex(Objects(o), D3DX.BufferGetTextureName(Meshes(Objects(o).MeshIndex).MaterialBuffer, i))
                            
                            If (l > 0) Then
                            
                                DDevice.SetTransform D3DTS_WORLD, Objects(o).Matrix
                                    
                                If BillBoards(l).Translucent Then
    
                                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
                                    If (BillBoards(l).AnimateMSecs > 0) Then
                                        If (BillBoards(l).AnimateTimer = 0) Or (CDbl(Timer - BillBoards(l).AnimateTimer) >= BillBoards(l).AnimateMSecs) Then
                                            BillBoards(l).AnimateTimer = GetTimer
                                            
                                            BillBoards(l).AnimatePoint = BillBoards(l).AnimatePoint + 1
                                            If BillBoards(l).AnimatePoint > BillBoards(l).SkinCount Then
                                                BillBoards(l).AnimatePoint = 1
                                            End If
                                            
                                        End If
                    
                                        DDevice.SetMaterial LucentMaterial
                                        DDevice.SetTexture 0, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                    
                                    Else
                                        DDevice.SetMaterial LucentMaterial
                                        DDevice.SetTexture 0, BillBoards(l).Skin(1)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, BillBoards(l).Skin(1)
                                    End If
    
                                    Meshes(Objects(o).MeshIndex).Mesh.DrawSubset i
                                ElseIf BillBoards(l).BlackAlpha Then
                                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
                                    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                    
                                    If (BillBoards(l).AnimateMSecs > 0) Then
                                        If (BillBoards(l).AnimateTimer = 0) Or (CDbl(Timer - BillBoards(l).AnimateTimer) >= BillBoards(l).AnimateMSecs) Then
                                            BillBoards(l).AnimateTimer = GetTimer
                                            
                                            BillBoards(l).AnimatePoint = BillBoards(l).AnimatePoint + 1
                                            If BillBoards(l).AnimatePoint > BillBoards(l).SkinCount Then
                                                BillBoards(l).AnimatePoint = 1
                                            End If
                                            
                                        End If
                    
                                        DDevice.SetMaterial LucentMaterial
                                        DDevice.SetTexture 0, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, BillBoards(l).Skin(BillBoards(l).AnimatePoint)
                    
                                    Else
                                        DDevice.SetMaterial LucentMaterial
                                        DDevice.SetTexture 0, BillBoards(l).Skin(1)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, BillBoards(l).Skin(1)
                                    End If
    
                                    Meshes(Objects(o).MeshIndex).Mesh.DrawSubset i
                                    
                                End If
    
                            End If
    
                        End If
    
                    Next
    
                End If
            End If
        Next
    End If
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                                    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    If BillBoardCount > 0 Then
        For o = 1 To BillBoardCount
            If BillBoards(o).Visible Then
                If Not (BillBoards(o).VBuf Is Nothing) Then
                    If Distance(Player.Object.Origin, BillBoards(o).Center) <= FadeDistance Then
                        If BillBoards(o).Translucent Then
            
                            If (BillBoards(o).AnimateMSecs > 0) Then
                                If (BillBoards(o).AnimateTimer = 0) Or (CDbl(Timer - BillBoards(o).AnimateTimer) >= BillBoards(o).AnimateMSecs) Then
                                    BillBoards(o).AnimateTimer = GetTimer
                                    
                                    BillBoards(o).AnimatePoint = BillBoards(o).AnimatePoint + 1
                                    If BillBoards(o).AnimatePoint > BillBoards(o).SkinCount Then
                                        BillBoards(o).AnimatePoint = 1
                                    End If
                                    
                                End If
            
                                DDevice.SetMaterial LucentMaterial
                                DDevice.SetTexture 0, BillBoards(o).Skin(BillBoards(o).AnimatePoint)
                                DDevice.SetTexture 1, Nothing
            
                            Else
                                DDevice.SetMaterial LucentMaterial
                                DDevice.SetTexture 0, BillBoards(o).Skin(1)
                                DDevice.SetTexture 1, Nothing
                            End If
                            DDevice.SetStreamSource 0, BillBoards(o).VBuf, Len(BillBoards(o).Plaq(0))
                            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
            
                        End If
                    End If
                End If
            End If
        Next
    End If

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub
Public Sub RenderBeacons()

    Dim l As Long
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1

    Dim start As Single
    Dim chop As Single
    Dim a As Single
    
    Dim X As Single
    Dim z As Single
    Dim ok As Boolean
    
    Dim o As Long
    Dim d As Single
    
    Dim matScale As D3DMATRIX
    Dim matPos As D3DMATRIX
    
    D3DXMatrixIdentity matWorld
        
    Dim matBeacon As D3DMATRIX
        
    If BeaconCount > 0 Then
        For o = 1 To BeaconCount
            If Beacons(o).Visible Then
            
                If Beacons(o).BeaconLight > -1 Then
                    If LightDataCount > 0 Then
                        For l = 1 To LightDataCount
                            DDevice.SetLight l - 1, Lights(LightDatas(l).LightIndex)
                            If Lights(LightDatas(l).LightIndex).Type = D3DLIGHT_DIRECTIONAL Then
                                DDevice.LightEnable l - 1, 1
                            Else
                                DDevice.LightEnable l - 1, False
                            End If
                        Next
                    End If
                End If
                
                If Beacons(o).Randomize Then
                    X = IIf((Rnd < 0.5), -RandomPositive(BeaconSpacing, BeaconRange), RandomPositive(BeaconSpacing, BeaconRange))
                    z = IIf((Rnd < 0.5), -RandomPositive(BeaconSpacing, BeaconRange), RandomPositive(BeaconSpacing, BeaconRange))
                    ok = True
                Else
                    ok = False
                End If
                
                If Beacons(o).OriginCount > 0 Then
                    l = 1
    
                    Do While l <= Beacons(o).OriginCount
                        d = Distance(Beacons(o).Origins(l), Player.Object.Origin)
                        If ok Then ok = ok And (Distance(Beacons(o).Origins(l), MakeVector(X, 0, z)) > BeaconSpacing)

                        If d <= FadeDistance Then
                            If Beacons(o).Consumable And (d <= 3) Then
                                If l + 1 < Beacons(o).OriginCount Then
                                    For d = l To (Beacons(o).OriginCount - 1) Step 1
                                        Beacons(o).Origins(l) = Beacons(o).Origins(l + 1)
                                    Next
                                ElseIf Not (l = Beacons(o).OriginCount) Then
                                    Beacons(o).Origins(l) = Beacons(o).Origins(l + 1)
                                End If
                                Beacons(o).OriginCount = Beacons(o).OriginCount - 1
                                If Beacons(o).OriginCount > 0 Then
                                    ReDim Preserve Beacons(o).Origins(1 To Beacons(o).OriginCount) As D3DVECTOR
                                End If
    
                            ElseIf l <= Beacons(o).OriginCount Then
    
                                If (Beacons(o).RoundingCut = 0) Then
                                
                                    D3DXMatrixIdentity matBeacon
                                
                                    If Beacons(o).VerticalLock Then
                                        
                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
                                            a = Cameras(Player.CameraIndex).Angle
                                            D3DXMatrixRotationYawPitchRoll matBeacon, -a, -Cameras(Player.CameraIndex).Pitch, 0
        
                                            D3DXMatrixScaling matScale, 1, 1, 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                            D3DXMatrixTranslation matPos, (Beacons(o).Origins(l).X - (Sin(D720 - a) * (Beacons(o).Dimension.height / (PI * 2.5)))), Beacons(o).Origins(l).Y, (Beacons(o).Origins(l).z - (Cos(D720 - a) * (Beacons(o).Dimension.height / (PI * 2.5))))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                        Else
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle), -Player.CameraPitch, 0
        
                                            D3DXMatrixScaling matScale, 1, 1, 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                            D3DXMatrixTranslation matPos, (Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 2.5)))), Beacons(o).Origins(l).Y, (Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 2.5))))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                                
                                        End If
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                    ElseIf Beacons(o).VerticalSkew Then
    
                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Cameras(Player.CameraIndex).Angle), -(Cameras(Player.CameraIndex).Pitch * 0.25), 0
                                            
                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Cameras(Player.CameraIndex).Pitch > 0, -Cameras(Player.CameraIndex).Pitch, Cameras(Player.CameraIndex).Pitch) * 0.25), 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
        
                                            D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (Beacons(o).Dimension.height / (PI * 6))), Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (Beacons(o).Dimension.height / (PI * 6)))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                        Else
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle), -(Player.CameraPitch * 0.25), 0
                                            
                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Player.CameraPitch > 0, -Player.CameraPitch, Player.CameraPitch) * 0.25), 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
        
                                            D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 6))), Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 6)))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                                
                                        End If
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                    Else
                                    
                                        D3DXMatrixRotationY matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle)
        
                                        D3DXMatrixScaling matScale, 1, 1, 1
                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                        D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z
                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                            
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                    
                                    End If
                             
                                    If Beacons(o).BeaconLight > -1 Then
                                        Lights(Beacons(o).BeaconLight).Position.X = Beacons(o).Origins(l).X - ((Beacons(o).Origins(l).X - Player.Object.Origin.X) / 80)
                                        Lights(Beacons(o).BeaconLight).Position.Y = Beacons(o).Origins(l).Y - ((Beacons(o).Origins(l).Y - Player.Object.Origin.Y) / 80)
                                        Lights(Beacons(o).BeaconLight).Position.z = Beacons(o).Origins(l).z - ((Beacons(o).Origins(l).z - Player.Object.Origin.z) / 80)
                                        
                                        DDevice.SetLight Beacons(o).BeaconLight - 1, Lights(Beacons(o).BeaconLight)
                                        DDevice.LightEnable Beacons(o).BeaconLight - 1, 1
                                    End If
                                    
                                    If (Beacons(o).BeaconAnim = 0) Or (CDbl(Timer - Beacons(o).BeaconAnim) >= 0.05) Then
                                        Beacons(o).BeaconAnim = GetTimer
                                        
                                        Beacons(o).BeaconText = Beacons(o).BeaconText + 1
                                        If Beacons(o).BeaconText > Beacons(o).BeaconSkinCount Then
                                            Beacons(o).BeaconText = 1
                                        End If
                                        
                                    End If
                                    
                                    If Beacons(o).Translucent Then
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        DDevice.SetMaterial LucentMaterial
        
                                    ElseIf Beacons(o).BlackAlpha Then
                                    
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
        
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        DDevice.SetMaterial GenericMaterial
                                    Else
        
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        DDevice.SetMaterial GenericMaterial
                                    End If
                                    DDevice.SetPixelShader PixelShaderDefault
                                    DDevice.SetTexture 0, Beacons(o).BeaconSkin(Beacons(o).BeaconText)
                                    DDevice.SetMaterial GenericMaterial
                                    DDevice.SetTexture 1, Beacons(o).BeaconSkin(Beacons(o).BeaconText)
                        
                                    DDevice.SetStreamSource 0, Beacons(o).BeaconVBuf, Len(Beacons(o).BeaconPlaq(0))
                                    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                                ElseIf Beacons(o).RoundingCut > 0 Then
    
                                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
                                    chop = Beacons(o).RoundingCut
                                    start = 0
                                    
                                    Do Until start >= 360
                          
                                        D3DXMatrixIdentity matBeacon

                                        D3DXMatrixRotationY matBeacon, (start / (PI * 2)) + IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle)
        
                                        D3DXMatrixScaling matScale, 1, 1, 1
                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                        D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z
                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                            
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                        

'                                    D3DXMatrixIdentity matBeacon
'
'                                    If Beacons(o).VerticalLock Then
'
'                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
'                                            a = Cameras(Player.CameraIndex).Angle
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, -a, -Cameras(Player.CameraIndex).Pitch, 0
'
'                                            D3DXMatrixScaling matScale, 1, 1, 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, (Beacons(o).Origins(l).X - (Sin(D720 - a) * (Beacons(o).Dimension.height / (PI * 2.5)))), Beacons(o).Origins(l).Y, (Beacons(o).Origins(l).z - (Cos(D720 - a) * (Beacons(o).Dimension.height / (PI * 2.5))))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'                                        Else
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle), -Player.CameraPitch, 0
'
'                                            D3DXMatrixScaling matScale, 1, 1, 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, (Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 2.5)))), Beacons(o).Origins(l).Y, (Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 2.5))))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        End If
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'                                    ElseIf Beacons(o).VerticalSkew Then
'
'                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Cameras(Player.CameraIndex).Angle), -(Cameras(Player.CameraIndex).Pitch * 0.25), 0
'
'                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Cameras(Player.CameraIndex).Pitch > 0, -Cameras(Player.CameraIndex).Pitch, Cameras(Player.CameraIndex).Pitch) * 0.25), 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (Beacons(o).Dimension.height / (PI * 6))), Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (Beacons(o).Dimension.height / (PI * 6)))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'                                        Else
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle), -(Player.CameraPitch * 0.25), 0
'
'                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Player.CameraPitch > 0, -Player.CameraPitch, Player.CameraPitch) * 0.25), 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 6))), Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.height / (PI * 6)))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        End If
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'                                    Else
'
'                                        D3DXMatrixRotationY matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle)
'
'                                        D3DXMatrixScaling matScale, 1, 1, 1
'                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                        D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z
'                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'
'                                    End If
                                        
                                        
                                        
                                        

                                        
                                        If Beacons(o).Translucent Then
                                            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                            DDevice.SetMaterial LucentMaterial
    
                                        ElseIf Beacons(o).BlackAlpha Then
    
                                            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
    
                                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                            DDevice.SetMaterial GenericMaterial
                                        Else
    
                                            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
                                            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                            DDevice.SetMaterial GenericMaterial
                                        End If

                                                                    
                                        If Beacons(o).BeaconLight > -1 Then
                                            Lights(Beacons(o).BeaconLight).Position.X = Beacons(o).Origins(l).X - ((Beacons(o).Origins(l).X - Player.Object.Origin.X) / 80)
                                            Lights(Beacons(o).BeaconLight).Position.Y = Beacons(o).Origins(l).Y - ((Beacons(o).Origins(l).Y - Player.Object.Origin.Y) / 80)
                                            Lights(Beacons(o).BeaconLight).Position.z = Beacons(o).Origins(l).z - ((Beacons(o).Origins(l).z - Player.Object.Origin.z) / 80)
                                            
                                            DDevice.SetLight Beacons(o).BeaconLight - 1, Lights(Beacons(o).BeaconLight)
                                            DDevice.LightEnable Beacons(o).BeaconLight - 1, 1
                                        End If
                                        
                                        If (Beacons(o).BeaconAnim = 0) Or (CDbl(Timer - Beacons(o).BeaconAnim) >= 0.05) Then
                                            Beacons(o).BeaconAnim = GetTimer
                                            
                                            Beacons(o).BeaconText = Beacons(o).BeaconText + 1
                                            If Beacons(o).BeaconText > Beacons(o).BeaconSkinCount Then
                                                Beacons(o).BeaconText = 1
                                            End If
                                            
                                        End If
                                        
                                        
                                        DDevice.SetPixelShader PixelShaderDefault
                                        DDevice.SetTexture 0, Beacons(o).BeaconSkin(Beacons(o).BeaconText)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, Beacons(o).BeaconSkin(Beacons(o).BeaconText)
                            
                                        DDevice.SetStreamSource 0, Beacons(o).BeaconVBuf, Len(Beacons(o).BeaconPlaq(0))
                                        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                                    
                                        start = (start + chop)
                                    Loop
                                    
                                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                                End If
                            End If
    
                        End If
                        
                        l = l + 1
                        
                    Loop
                    
                End If
                
                If Beacons(o).BeaconLight > -1 Then
                    If LightDataCount > 0 Then
                        For l = 1 To LightDataCount
                            DDevice.SetLight l - 1, Lights(LightDatas(l).LightIndex)
                            If Lights(LightDatas(l).LightIndex).Type = D3DLIGHT_DIRECTIONAL Or (l = Beacons(o).BeaconLight) Then
                                DDevice.LightEnable l - 1, False
                            Else
                                DDevice.LightEnable l - 1, 1
                            End If
                        Next
                    End If
                End If
            End If
            
            If ok And Beacons(o).Randomize And (Beacons(o).OriginCount < Beacons(o).Allowance) Then
                Beacons(o).OriginCount = Beacons(o).OriginCount + 1
                ReDim Preserve Beacons(o).Origins(1 To Beacons(o).OriginCount) As D3DVECTOR
                Beacons(o).Origins(Beacons(o).OriginCount).X = X
                Beacons(o).Origins(Beacons(o).OriginCount).z = z
            End If
            
        Next

    End If
                    
End Sub

Public Sub RenderPlanes()

    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetVertexShader FVF_RENDER

    
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Dim fogSTate As Boolean
    fogSTate = DDevice.GetRenderState(D3DRS_FOGENABLE)
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_ZENABLE, False

    Dim l As Long
    If LightCount > 0 Then
        For l = 1 To LightDataCount
            If LightDatas(l).SunLight Then
                DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(0, 164 + Lights(LightDatas(l).LightIndex).diffuse.r * 255, _
                    164 + Lights(LightDatas(l).LightIndex).diffuse.g * 255, 164 + Lights(LightDatas(l).LightIndex).diffuse.b * 255)
                    
                If (FPSRate > 0) And LightDatas(l).SunLight And CloudRotated > 0 Then

                    CloudRotated = CloudRotated + (360 / (HoursInOneDay * SkyCRot)) * (PI / (HoursInOneDay * FPSRate)) * (FPSRate / HoursInOneDay)
                End If
                
            End If
        Next
    End If
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT



    Dim matProj As D3DMATRIX
    Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
    DDevice.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0

    DDevice.SetTransform D3DTS_VIEW, matView

    D3DXMatrixPerspectiveFovLH matProj, PI / 2.5, AspectRatio, 0.05, 50
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    If Not CloudRotated = 0 Then
        D3DXMatrixRotationY matWorld, CloudRotated * (PI / 180)
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
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, 1

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetRenderState D3DRS_FOGCOLOR, DDevice.GetRenderState(D3DRS_AMBIENT)
    
    DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(0, 0, 0, 0)
    
End Sub

Public Sub CreateLand()
    DebugMode = False
    Perspective = ThirdPerson
    
    MaxCameraZoom = 10
    MinCameraZoom = 1
    CloudRotated = 0
    Serialize = ""
    Deserialize = ""
    
    ParseLand 0, Replace(ReadFile(AppPath & "Levels\" & CurrentLoadedLevel & ".px"), vbTab, "")
    
    ComputeNormals

    db.rsQuery rs, "SELECT * FROM Serials WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND PXFile='" & Replace(CurrentLoadedLevel, "'", "''") & "';"
    If Not db.rsEnd(rs) Then
        ParseLand 0, Replace(rs("Script"), "\n", vbCrLf)
        If Not (Deserialize = "") Then
            ParseLand NextArg(Deserialize, ":"), RemoveArg(Deserialize, ":")
        End If
    End If
    
End Sub
Public Sub ParseLine(ByRef inItem As String, ByRef inText As String, ByRef LineNumber As Long)
    If InStr(inText, vbCrLf) > 0 Then
        If InStr(inText, "[") < InStr(inText, vbCrLf) Then
            inItem = Left(inText, InStr(inText, "[") - 1)
            inText = Mid(inText, InStr(inText, "["))
        Else
        
        End If
    ElseIf InStr(inText, "[") > 0 Then
        inItem = Left(inText, InStr(inText, "[") - 1)
        inText = Mid(inText, InStr(inText, "["))
    Else
        inItem = inText
        inText = ""
    End If
End Sub
Public Function ParseLand(ByVal inLine As Long, ByVal inText As String) As String
On Error GoTo parseerror
    
    Dim r As Single
    Dim o As Long
    Dim i As Long
    Dim cnt As Long
    Dim cnt2 As Long

    Dim vn As D3DVECTOR

    Dim NumLines As Long
    NumLines = inLine + CountWord(inText, vbCrLf)
    
    Dim inArg() As String
    Dim inItem As String
    Dim inName As String
    Dim inData As String
    Dim inTrig As String
    
    Do Until inText = ""
    
        If (Left(Replace(Replace(inText, " ", ""), vbTab, ""), 1) = ";") Then
            RemoveNextArg inText, vbCrLf
        Else
            inItem = RemoveLineArg(inText, "{")
            If (Not (Trim(inItem) = "")) Then
                inLine = (NumLines - CountWord(inText, vbCrLf))
                Select Case Trim(inItem)
                    Case "parse"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 1)
                                    Select Case inArg(0)
                                        Case "filename"
                                            If PathExists(AppPath & "Levels\" & inArg(1), True) Then
                                                ParseLand 0, Replace(ReadFile(AppPath & "Levels\" & inArg(1)), vbTab, "")
                                            Else
                                                AddMessage "Invalid object file [" & AppPath & "Levels\" & inArg(1) & "]"
                                            End If
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                End If
                            End If
                        Loop
    
                    Case "plane"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 4)
                                    Select Case inArg(0)
                                        Case "maxzoom"
                                            MaxCameraZoom = CSng(inArg(1))
                                        Case "minzoom"
                                            MinCameraZoom = CSng(inArg(1))
                                        Case "gravity"
                                            SetActivity GlobalGravityDirect, Actions.Directing, MakeVector(0, CSng(inArg(1)), 0), 1
                                            SetActivity LiquidGravityDirect, Actions.Directing, MakeVector(0, CSng(inArg(1)) / 40, 0), 2
                                        Case "cloudrotate"
                                            SkyCRot = CSng(inArg(1))
                                        Case "fogcolor"
                                            If CSng(inArg(1)) > 0 Then inArg(1) = 255 / CSng(inArg(1))
                                            If CSng(inArg(2)) > 0 Then inArg(2) = 255 / CSng(inArg(2))
                                            If CSng(inArg(3)) > 0 Then inArg(3) = 255 / CSng(inArg(3))
                                            If CSng(inArg(4)) > 0 Then inArg(4) = 255 / CSng(inArg(4))
                                            DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)), CSng(inArg(4)))
                                        Case "fogdisable"
                                            DDevice.SetRenderState D3DRS_FOGENABLE, False
                                        Case "fogdistance"
                                            DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(CSng(inArg(1)) / 4)
                                            DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(CSng(inArg(1)))
                                        Case "skytop"
                                            Set SkySkin(0) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case "skyback"
                                            Set SkySkin(1) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case "skyleft"
                                            Set SkySkin(2) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case "skyfront"
                                            Set SkySkin(3) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case "skyright"
                                            Set SkySkin(4) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case "skybottom"
                                            Set SkySkin(5) = LoadTexture(AppPath & "Models\" & inArg(1))
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                   End Select
                                ElseIf Left(Trim(inData), 1) = "[" Then
                                    AddMessage "Warning, Itemless Brackets at Line " & inLine
                                    GoTo throwerror
                                End If
                            End If
                        Loop
                        
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
                        
                    Case "light"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        LightCount = LightCount + 1
                        ReDim Preserve Lights(1 To LightCount) As D3DLIGHT8
                        With Lights(LightCount)
                            Dim lEnable As Boolean
                            Dim lIdentity As String
                            lEnable = True
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 4)
                                        Select Case inArg(0)
                                            Case "identity"
                                                lIdentity = inArg(1)
                                            Case "position"
                                                LightDataCount = LightDataCount + 1
                                                ReDim Preserve LightDatas(1 To LightDataCount) As MyLight
                                                LightDatas(LightDataCount).Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                                .Position = LightDatas(LightDataCount).Origin
                                                LightDatas(LightDataCount).LightIndex = LightCount
                                                LightDatas(LightDataCount).Enabled = lEnable
                                                DDevice.SetLight LightDataCount - 1, Lights(LightCount)
                                                DDevice.LightEnable LightDataCount - 1, lEnable
                                                LightDatas(LightDataCount).Identity = lIdentity
                                            Case "diffuseroll"
                                                LightDatas(LightDataCount).DiffuseRoll = CSng(inArg(1))
                                                LightDatas(LightDataCount).DIffuseMax = CSng(inArg(2))
                                                LightDatas(LightDataCount).DiffuseNow = 0
                                            Case "sunlight"
                                                LightDatas(LightDataCount).SunLight = True
                                            Case "blink"
                                                LightDatas(LightDataCount).LightBlink = inArg(1)
                                            Case "enabled"
                                                If Not (inArg(1) = "") Then
                                                    lEnable = CBool(inArg(1))
                                                Else
                                                    lEnable = True
                                                End If
                                            Case "disable"
                                                lEnable = False
                                            Case "enable"
                                                lEnable = True
                                            Case "direction"
                                                .Direction = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Case "diffuse"
                                                .diffuse.a = CSng(inArg(1))
                                                .diffuse.r = CSng(inArg(2))
                                                .diffuse.g = CSng(inArg(3))
                                                .diffuse.b = CSng(inArg(4))
                                            Case "ambience"
                                                .Ambient.a = CSng(inArg(1))
                                                .Ambient.r = CSng(inArg(2))
                                                .Ambient.g = CSng(inArg(3))
                                                .Ambient.b = CSng(inArg(4))
                                            Case "specular"
                                                .specular.a = CSng(inArg(1))
                                                .specular.r = CSng(inArg(2))
                                                .specular.g = CSng(inArg(3))
                                                .specular.b = CSng(inArg(4))
                                            Case "attenuation"
                                                .Attenuation0 = CSng(inArg(1))
                                                .Attenuation1 = CSng(inArg(2))
                                                .Attenuation2 = CSng(inArg(3))
                                            Case "phi"
                                                .Phi = CSng(inArg(1))
                                            Case "theta"
                                                .Theta = CSng(inArg(1))
                                            Case "falloff"
                                                .Falloff = CSng(inArg(1))
                                            Case "range"
                                                .Range = CSng(inArg(1))
                                            Case "type"
                                                .Type = CLng(inArg(1))
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                       End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                    Case "ground", "player", "object"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Dim NewObj As MyObject
                        NewObj.VisualIndex = 0
                        NewObj.MeshIndex = 0
                        NewObj.Effect = Collides.None
                        NewObj.Gravitational = False
                        NewObj.BlackAlpha = False
                        NewObj.CollideIndex = 0
                        NewObj.CollideObject = -1
                        NewObj.Identity = ""
                        NewObj.WireFrame = False
                        
                        Do Until NewObj.ActivityCount = 0
                            DeleteActivity NewObj, NewObj.Activities(1).Identity
                        Loop
                        NewObj.Origin = MakeVector(0, 0, 0)
                        
                        If (inItem = "player") Then
                            Player.CameraAngle = 0
                            Player.CameraPitch = 0
                            Player.MoveSpeed = 0.3
                            Player.Boundary = 90
                        End If
                        
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 9)
                                    Select Case inArg(0)
                                        Case "visible"
                                            If inArg(1) = "" Then
                                                NewObj.Visible = True
                                            Else
                                                NewObj.Visible = CBool(inArg(1))
                                            End If
                                        Case "hidden"
                                            NewObj.Visible = False
                                        Case "wireframe"
                                            NewObj.WireFrame = True
                                        Case "indoorcollide"
                                            NewObj.Effect = Collides.InDoor
                                        Case "groundcollide"
                                            NewObj.Effect = Collides.Ground
                                        Case "liquidcollide"
                                            NewObj.Effect = Collides.Liquid
                                        Case "laddercollide"
                                            NewObj.Effect = Collides.Ladder
                                        Case "gravitational"
                                            NewObj.Gravitational = True
                                        Case "identity"
                                            NewObj.Identity = inArg(1)
                                        Case "blackalpha"
                                            NewObj.BlackAlpha = True
                                        Case "nocollision"
                                            NewObj.CollideIndex = -1
                                        Case "origin", "location", "position"
                                            NewObj.Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                        Case "scale"
                                            NewObj.Scaled = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                        Case "rotate"
                                            NewObj.Rotate = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                        Case "offset"
                                            NewObj.Offset = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                        Case "folcrum"
                                            NewObj.FolcrumCount = NewObj.FolcrumCount + 1
                                            ReDim Preserve NewObj.Folcrum(1 To NewObj.FolcrumCount) As D3DVECTOR
                                            NewObj.Folcrum(NewObj.FolcrumCount) = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                        Case "visualobj"
                                            If MeshCount = 0 Then
                                                MeshCount = MeshCount + 1
                                                ReDim Meshes(1 To MeshCount) As MyMesh
                                                NewObj.VisualIndex = MeshCount
                                                Meshes(NewObj.VisualIndex).FileName = LCase(inArg(1))
                                                If PathExists(AppPath & "Models\" & Meshes(NewObj.VisualIndex).FileName, True) Then
                                                    CreateMesh AppPath & "Models\" & Meshes(NewObj.VisualIndex).FileName, Meshes(NewObj.VisualIndex).Mesh, Meshes(NewObj.VisualIndex).MaterialBuffer, _
                                                            NewObj.Origin, NewObj.Scaled, Meshes(NewObj.VisualIndex).Materials, Meshes(NewObj.VisualIndex).Textures, _
                                                            Meshes(NewObj.VisualIndex).Verticies, Meshes(NewObj.VisualIndex).Indicies, Meshes(NewObj.VisualIndex).MaterialCount
                                                Else
                                                    ReDim Meshes(NewObj.VisualIndex).Textures(0 To 0) As Direct3DTexture8
                                                    ReDim Meshes(NewObj.VisualIndex).Materials(0 To 0) As D3DMATERIAL8
                                                    NewObj.VisualIndex = 0
                                                End If
                                            Else
                                                For i = LBound(Meshes) To UBound(Meshes)
                                                    If Meshes(i).FileName = LCase(inArg(1)) Then
                                                        NewObj.VisualIndex = i
                                                        Exit For
                                                    End If
                                                Next
                                                If NewObj.VisualIndex = 0 Then
                                                    MeshCount = MeshCount + 1
                                                    ReDim Preserve Meshes(1 To MeshCount) As MyMesh
                                                    NewObj.VisualIndex = MeshCount
                                                    Meshes(NewObj.VisualIndex).FileName = LCase(inArg(1))
                                                    If PathExists(AppPath & "Models\" & Meshes(NewObj.VisualIndex).FileName, True) Then
                                                        CreateMesh AppPath & "Models\" & Meshes(NewObj.VisualIndex).FileName, Meshes(NewObj.VisualIndex).Mesh, Meshes(NewObj.VisualIndex).MaterialBuffer, _
                                                                NewObj.Origin, NewObj.Scaled, Meshes(NewObj.VisualIndex).Materials, Meshes(NewObj.VisualIndex).Textures, _
                                                                Meshes(NewObj.VisualIndex).Verticies, Meshes(NewObj.VisualIndex).Indicies, Meshes(NewObj.VisualIndex).MaterialCount
                                                    Else
                                                        ReDim Meshes(NewObj.VisualIndex).Textures(0 To 0) As Direct3DTexture8
                                                        ReDim Meshes(NewObj.VisualIndex).Materials(0 To 0) As D3DMATERIAL8
                                                        NewObj.VisualIndex = 0
                                                    End If
                                                End If
                                            End If
                                            
                                            
                                        Case "filename", "boundsobj"
                                            
                                            If MeshCount = 0 Then
                                                MeshCount = MeshCount + 1
                                                ReDim Meshes(1 To MeshCount) As MyMesh
                                                NewObj.MeshIndex = MeshCount
                                                Meshes(NewObj.MeshIndex).FileName = LCase(inArg(1))
                                                If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).FileName, True) Then
                                                    CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).FileName, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
                                                            NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
                                                            Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
                                                Else
                                                    ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
                                                    ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
                                                    NewObj.MeshIndex = 0
                                                End If
                                            Else
                                                For i = LBound(Meshes) To UBound(Meshes)
                                                    If Meshes(i).FileName = LCase(inArg(1)) Then
                                                        NewObj.MeshIndex = i
                                                        Exit For
                                                    End If
                                                Next
                                                
                                                If NewObj.MeshIndex = 0 Then
                                                    MeshCount = MeshCount + 1
                                                    ReDim Preserve Meshes(1 To MeshCount) As MyMesh
                                                    NewObj.MeshIndex = MeshCount
                                                    Meshes(NewObj.MeshIndex).FileName = LCase(inArg(1))
                                                    If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).FileName, True) Then
                                                        CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).FileName, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
                                                                NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
                                                                Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
                                                    Else
                                                        ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
                                                        ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
                                                        NewObj.MeshIndex = 0
                                                    End If
                                            
                                                End If
                                                
                                            End If
                                        Case "replacer"
                                            If BillBoardCount > 0 Then
                                                For o = 1 To BillBoardCount
                                                    If BillBoards(o).Identity = inArg(2) Then
                                                        NewObj.ReplacerVals.Add o, NewObj.Identity & "_" & Replace(inArg(1), ".", "")
                                                        NewObj.ReplacerKeys.Add NewObj.Identity & "_" & Replace(inArg(1), ".", "")
                                                    End If
                                                Next
                                            End If
                                        Case "activity"

                                            Select Case LCase(CStr(inArg(1)))
                                                Case "direct"
                                                    AddActivity NewObj, Actions.Directing, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "rotate"
                                                    
                                                    AddActivity NewObj, Actions.Rotating, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "scale"
                                                    AddActivity NewObj, Actions.Scaling, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "script"
                                                    If InStr(inData, "[") > 0 Then
                                                        inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                        inTrig = RemoveQuotedArg(inData, "[", "]", True)
                                                        inTrig = inLine & ":" & inTrig
                                                        
                                                        AddActivity NewObj, Actions.Script, inArg(2), MakeVector(0, 0, 0), _
                                                                                , , IIf(IsNumeric(inArg(3)), inArg(3), -1), _
                                                                                IIf(IsNumeric(inArg(4)), inArg(4), -1), inTrig
                                                    Else
                                                        AddMessage "Warning, Brackets Required at Line " & inLine
                                                    End If
        
                                            End Select
                                        Case "boundary"
                                            If inItem = "player" Then Player.Boundary = CSng(inArg(1))
                                        Case "movespeed"
                                            If inItem = "player" Then Player.MoveSpeed = CSng(inArg(1))
                                        Case "camerapitch"
                                            If inItem = "player" Then Player.CameraPitch = CSng(inArg(1))
                                        Case "cameraangle"
                                            If inItem = "player" Then Player.CameraAngle = CSng(inArg(1))
        
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                ElseIf Left(Trim(inData), 1) = "[" Then
                                    AddMessage "Warning, Itemless Brackets at Line " & inLine
                                    GoTo throwerror
                                End If
                            End If
                        Loop
                        
                        If (NewObj.MeshIndex > 0) Then
                            
                            D3DXMatrixIdentity NewObj.Matrix
                            D3DXMatrixTranslation NewObj.Matrix, NewObj.Offset.X, NewObj.Offset.Y, NewObj.Offset.z
                            D3DXMatrixRotationX NewObj.Matrix, NewObj.Rotate.X * (PI / 180)
                            D3DXMatrixRotationY NewObj.Matrix, NewObj.Rotate.Y * (PI / 180)
                            D3DXMatrixRotationZ NewObj.Matrix, NewObj.Rotate.z * (PI / 180)
                            D3DXMatrixScaling NewObj.Matrix, NewObj.Scaled.X, NewObj.Scaled.Y, NewObj.Scaled.z
                            D3DXMatrixIdentity matWorld

                        End If
                        
                        If (NewObj.CollideIndex > -1) And (NewObj.MeshIndex > 0) Then
                            AddCollision NewObj
                        Else
                            NewObj.CollideIndex = -1
                        End If
                        
                        If (inItem = "player") Then
                            Player.Object = NewObj
                        Else
                            ObjectCount = ObjectCount + 1
                            ReDim Preserve Objects(1 To ObjectCount) As MyObject
                            Objects(ObjectCount) = NewObj
                        End If
                    Case "sound"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        SoundCount = SoundCount + 1
                        ReDim Preserve Sounds(1 To SoundCount) As MySound
                        With Sounds(SoundCount)
                            .Index = SoundCount
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 3)
                                        Select Case inArg(0)
                                            Case "filename"
                                                LoadWave .Index, AppPath & "Sounds\" & inArg(1)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "enabled"
                                                If Not (inArg(1) = "") Then
                                                    .Enable = CBool(inArg(1))
                                                Else
                                                    .Enable = True
                                                End If
                                            Case "disable"
                                                .Enable = False
                                            Case "enable"
                                                .Enable = True
                                            Case "repeat"
                                                .Repeat = True
                                            Case "origin"
                                                .Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Case "range"
                                                .Range = CSng(inArg(1))
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                    Case "ambient"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        TrackCount = TrackCount + 1
                        ReDim Preserve Tracks(1 To TrackCount) As clsAmbient
                        Set Tracks(TrackCount) = New clsAmbient
                        With Tracks(TrackCount)
                        
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 3)
                                        Select Case inArg(0)
                                            Case "filename"
                                                .FileName = AppPath & "Sounds\" & inArg(1)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "volume"
                                                .TrackVolume = CSng(inArg(1))
                                            Case "origin"
                                                .Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Case "range"
                                                .Range = CSng(inArg(1))
                                            Case "loops"
                                                If IsNumeric(inArg(1)) Then
                                                    .LoopEnabled = True
                                                    .LoopTimes = CLng(inArg(1))
                                                Else
                                                    .LoopEnabled = True
                                                    .LoopTimes = 0
                                                End If
                                            Case "enabled"
                                                If Not (inArg(1) = "") Then
                                                    .Enable = CBool(inArg(1))
                                                Else
                                                    .Enable = True
                                                End If
                                            Case "disable"
                                                .Enable = False
                                            Case "enable"
                                                .Enable = True
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                    Case "portal"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                      
'                        Dim txtobj As String
'                        txtobj = "beacon" & vbCrLf & "{" & vbCrLf
'                        txtobj = txtobj & "identity beaconPortal" & (PortalCount + 1) & vbCrLf
'                        txtobj = txtobj & "visible true" & vbCrLf
                        
                        
                        PortalCount = PortalCount + 1
                        ReDim Preserve Portals(1 To PortalCount) As MyPortal
                        With Portals(PortalCount)
                            .Enable = True

                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 9)
                                        Select Case inArg(0)
                                            Case "enable"
                                                .Enable = CBool(inArg(1))
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "clearactivities"
                                                .ClearActivities = True
                                            Case "range"
                                                .Range = CSng(inArg(1))
'                                                txtobj = txtobj & "percentxy " & (CSng(inArg(1)) * 100) & " " & (CSng(inArg(1)) * 100) & vbCrLf
                        
                                            Case "activity"
                                                Select Case LCase(CStr(inArg(1)))
                                                    Case "direct"
                                                        AddActivityEx Portals(PortalCount), Actions.Directing, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                                IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                                IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                    Case "rotate"
                                                        AddActivityEx Portals(PortalCount), Actions.Rotating, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                                IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                                IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                    Case "scale"
                                                        AddActivityEx Portals(PortalCount), Actions.Scaling, inArg(2), MakeVector(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
                                                                                IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                                IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                    Case "script"
                                                        If InStr(inData, "[") > 0 Then
                                                            inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                            inTrig = RemoveQuotedArg(inData, "[", "]", True)
                                                            inTrig = inLine & ":" & inTrig
                                                            
                                                            AddActivityEx Portals(PortalCount), Actions.Script, inArg(2), MakeVector(0, 0, 0), _
                                                                                    , , IIf(IsNumeric(inArg(3)), inArg(3), -1), _
                                                                                    IIf(IsNumeric(inArg(4)), inArg(4), -1), inTrig
                                                        Else
                                                            AddMessage "Warning, Brackets Required at Line " & inLine
                                                        End If
        
                                                    Case Else
                                                        AddMessage "Warning, Unknown Action at Line " & inLine
                                                End Select
                                            Case "location"
                                                .Location = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                                txtobj = txtobj & "origin " & CSng(inArg(1)) & " " & CSng(inArg(2)) & " " & CSng(inArg(3)) & vbCrLf
                        
                                            Case "teleport"
                                                .Teleport = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Case "oninrange"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .OnInRange = RemoveQuotedArg(inData, "[", "]", True)
                                                    .OnInRange = inLine & ":" & .OnInRange
                                                Else
                                                    AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
        
                                            Case "onoutrange"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .OnOutRange = RemoveQuotedArg(inData, "[", "]", True)
                                                    .OnOutRange = inLine & ":" & .OnOutRange
                                                Else
                                                    AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
                                                
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With

'                        txtobj = txtobj & "blackalpha" & vbCrLf
'                        txtobj = txtobj & "filename bubble.bmp" & vbCrLf
'                        txtobj = txtobj & "beaconlight 1" & vbCrLf
'                        txtobj = txtobj & "verticallock" & vbCrLf
'                        txtobj = txtobj & "}" & vbCrLf
'                        ParseLand 0, txtobj

                    Case "billboard"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        o = 0
                        BillBoardCount = BillBoardCount + 1
                        ReDim Preserve BillBoards(1 To BillBoardCount) As MyBoard
                        With BillBoards(BillBoardCount)
                            ReDim .Plaq(0 To 5) As MyVertex
                        
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 5)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "visible"
                                                ReDim Preserve inArg(0 To 1)
                                                If inArg(1) = "" Then
                                                    .Visible = True
                                                Else
                                                    .Visible = CBool(inArg(1))
                                                End If
                                            Case "hidden"
                                                .Visible = False
                                            Case "filename"
                                                .SkinCount = .SkinCount + 1
                                                ReDim Preserve .Skin(1 To .SkinCount) As Direct3DTexture8
                                                Set .Skin(.SkinCount) = LoadTexture(AppPath & "Models\" & inArg(1))
                                            Case "point1"
                                                .Point1.X = CSng(inArg(1))
                                                .Point1.Y = CSng(inArg(2))
                                                .Point1.z = CSng(inArg(3))
                                                o = 1
                                                If IsNumeric(inArg) = 5 Then
                                                    o = 2
                                                    .Point1.tu = CSng(inArg(4))
                                                    .Point1.tv = CSng(inArg(5))
                                                End If
                                            Case "point2"
                                                .Point2.X = CSng(inArg(1))
                                                .Point2.Y = CSng(inArg(2))
                                                .Point2.z = CSng(inArg(3))
                                                o = 1
                                                If IsNumeric(inArg) = 5 Then
                                                    o = 2
                                                    .Point2.tu = CSng(inArg(4))
                                                    .Point2.tv = CSng(inArg(5))
                                                End If
                                            Case "point3"
                                                .Point3.X = CSng(inArg(1))
                                                .Point3.Y = CSng(inArg(2))
                                                .Point3.z = CSng(inArg(3))
                                                o = 1
                                                If IsNumeric(inArg) = 5 Then
                                                    o = 2
                                                    .Point3.tu = CSng(inArg(4))
                                                    .Point3.tv = CSng(inArg(5))
                                                End If
                                            Case "point4"
                                                .Point4.X = CSng(inArg(1))
                                                .Point4.Y = CSng(inArg(2))
                                                .Point4.z = CSng(inArg(3))
                                                o = 1
                                                If IsNumeric(inArg) = 5 Then
                                                    o = 2
                                                    .Point4.tu = CSng(inArg(4))
                                                    .Point4.tv = CSng(inArg(5))
                                                End If
                                            Case "scalex"
                                                .ScaleX = CSng(inArg(1))
                                                o = 1
                                            Case "scaley"
                                                .ScaleY = CSng(inArg(1))
                                                o = 1
                                            Case "animated"
                                                .AnimateMSecs = CSng(inArg(1))
                                                .AnimateTimer = GetTimer
                                                .AnimatePoint = 1
                                            Case "translucent"
                                                .Translucent = True
                                            Case "blackalpha"
                                                .BlackAlpha = True
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                            
                            If o = 1 Then
                                CreateSquare .Plaq, 0, MakeVector(.Point1.X, .Point1.Y, .Point1.z), _
                                                                MakeVector(.Point2.X, .Point2.Y, .Point2.z), _
                                                                MakeVector(.Point3.X, .Point3.Y, .Point3.z), _
                                                                MakeVector(.Point4.X, .Point4.Y, .Point4.z), _
                                                                .ScaleX, .ScaleY
                            ElseIf o = 2 Then
                                CreateSquareEx .Plaq, 0, .Point1, .Point2, .Point3, .Point4
                                
                            End If
        
                            If o = 1 Or o = 2 Then
                                .Center = SquareCenter(MakeVector(.Point1.X, .Point1.Y, .Point1.z), _
                                                                MakeVector(.Point2.X, .Point2.Y, .Point2.z), _
                                                                MakeVector(.Point3.X, .Point3.Y, .Point3.z), _
                                                                MakeVector(.Point4.X, .Point4.Y, .Point4.z))
        
                                Set .VBuf = DDevice.CreateVertexBuffer(Len(.Plaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
                                D3DVertexBuffer8SetData .VBuf, 0, Len(.Plaq(0)) * 6, 0, .Plaq(0)
                            End If
                        End With
    
                    Case "beacon"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        BeaconCount = BeaconCount + 1
                        ReDim Preserve Beacons(1 To BeaconCount) As MyBeacon
                        With Beacons(BeaconCount)
                        
                            .BeaconLight = -1
                            .Dimension.width = 1
                            .Dimension.height = 1
                            .PercentXY.width = 100
                            .PercentXY.height = 100
                            .Allowance = 1
                            
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 3)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "visible"
                                                ReDim Preserve inArg(0 To 1)
                                                If inArg(1) = "" Then
                                                    .Visible = True
                                                Else
                                                    .Visible = CBool(inArg(1))
                                                End If
                                            Case "hidden"
                                                .Visible = False
                                            Case "verticalskew"
                                                .VerticalSkew = True
                                            Case "roundingcut"
                                                .RoundingCut = CLng(inArg(1))
                                            Case "consumable"
                                                .Consumable = True
                                            Case "randomize"
                                                .Randomize = True
                                            Case "allowance"
                                                .Allowance = CLng(inArg(1))
                                            Case "origin"
                                                .OriginCount = .OriginCount + 1
                                                ReDim Preserve .Origins(1 To .OriginCount) As D3DVECTOR
                                                .Origins(.OriginCount) = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Case "translucent"
                                                .Translucent = True
                                                .BlackAlpha = False
                                            Case "blackalpha"
                                                .BlackAlpha = True
                                                .Translucent = False
                                            Case "dimension"
                                                .Dimension.width = CSng(inArg(1))
                                                .Dimension.height = CSng(inArg(2))
                                            Case "percentxy"
                                                .PercentXY.width = CSng(inArg(1))
                                                .PercentXY.height = CSng(inArg(2))
                                            Case "verticallock"
                                                .VerticalLock = True
                                            Case "horizontallock"
                                                .HorizontalLock = True
                                            Case "filename"
                                                .BeaconSkinCount = .BeaconSkinCount + 1
                                                ReDim Preserve .BeaconSkin(1 To .BeaconSkinCount) As Direct3DTexture8
                                                Set .BeaconSkin(.BeaconSkinCount) = LoadTexture(AppPath & "Models\" & inArg(1))
                                            Case "beaconlight"
                                                .BeaconLight = CLng(inArg(1))
                                                
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                            
                            CreateSquare .BeaconPlaq, 0, _
                                MakeVector(((.Dimension.width * (.PercentXY.width / 100)) / 2), 0, 0), _
                                MakeVector(-((.Dimension.width * (.PercentXY.width / 100)) / 2), 0, 0), _
                                MakeVector(-((.Dimension.width * (.PercentXY.width / 100)) / 2), (.Dimension.height * (.PercentXY.height / 100)), 0), _
                                MakeVector(((.Dimension.width * (.PercentXY.width / 100)) / 2), (.Dimension.height * (.PercentXY.height / 100)), 0)
                                
                            Set .BeaconVBuf = DDevice.CreateVertexBuffer(Len(.BeaconPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
                            D3DVertexBuffer8SetData .BeaconVBuf, 0, Len(.BeaconPlaq(0)) * 6, 0, .BeaconPlaq(0)
                        End With
                    Case "image"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        ScreenImageCount = ScreenImageCount + 1
                        ReDim Preserve ScreenImages(1 To ScreenImageCount) As MyImage
                        With ScreenImages(ScreenImageCount)
                        
                            .Verticies(0) = MakeScreen(0, 0, -1, 0, 0)
                            .Verticies(1) = MakeScreen(0, 0, -1, 1, 0)
                            .Verticies(2) = MakeScreen(0, 0, -1, 0, 1)
                            .Verticies(3) = MakeScreen(0, 0, -1, 1, 1)
                            
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 2)
                                        Select Case inArg(0)
                                            Case "visible"
                                                ReDim Preserve inArg(0 To 1)
                                                If inArg(1) = "" Then
                                                    .Visible = True
                                                Else
                                                    .Visible = CBool(inArg(1))
                                                End If
                                            Case "blackalpha"
                                                .BlackAlpha = True
                                            Case "hidden"
                                                .Visible = False
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "filename"
                                                 If PathExists(AppPath & "Models\" & inArg(1), True) Then
                                                    Set .Image = LoadTexture(AppPath & "Models\" & inArg(1))
                                                    If Not ImageDimensions(AppPath & "Models\" & inArg(1), .Dimension) Then
                                                        Debug.Print "Image Dimension Error"
                                                    End If
                                                Else
                                                    Debug.Print "Image Not Found"
                                                End If
                                            Case "translucent"
                                                .Translucent = True
                                            Case "padding"
                                                .Padding = CLng(inArg(1))
                                                
                                            Case "coordinate"
                                                .Verticies(0).X = .Padding + CLng(inArg(1))
                                                .Verticies(2).X = .Padding + CLng(inArg(1))
                                                .Verticies(1).X = .Verticies(0).X + .Dimension.width
                                                .Verticies(3).X = .Verticies(2).X + .Dimension.width
                                                .Verticies(0).Y = .Padding + CLng(inArg(2))
                                                .Verticies(1).Y = .Padding + CLng(inArg(2))
                                                .Verticies(2).Y = .Verticies(0).Y + .Dimension.height
                                                .Verticies(3).Y = .Verticies(1).Y + .Dimension.height
                                            Case "alignxleft" 'left center right
                                                        .Verticies(0).X = .Padding
                                                        .Verticies(2).X = .Padding
                                                .Verticies(1).X = .Verticies(0).X + .Dimension.width
                                                .Verticies(3).X = .Verticies(2).X + .Dimension.width
                                            Case "alignxcenter"
                                                        .Verticies(0).X = ((frmMain.width / Screen.TwipsPerPixelX) / 2) - (.Dimension.width / 2)
                                                        .Verticies(2).X = ((frmMain.width / Screen.TwipsPerPixelX) / 2) - (.Dimension.width / 2)
                                                .Verticies(1).X = .Verticies(0).X + .Dimension.width
                                                .Verticies(3).X = .Verticies(2).X + .Dimension.width
                                            Case "alignxright"
                                                        .Verticies(0).X = (frmMain.width / Screen.TwipsPerPixelX) - .Padding - .Dimension.width
                                                        .Verticies(2).X = (frmMain.width / Screen.TwipsPerPixelX) - .Padding - .Dimension.width
                                                .Verticies(1).X = .Verticies(0).X + .Dimension.width
                                                .Verticies(3).X = .Verticies(2).X + .Dimension.width
                                            Case "alignytop"
                                                        .Verticies(0).Y = .Padding
                                                        .Verticies(1).Y = .Padding
                                                .Verticies(2).Y = .Verticies(0).Y + .Dimension.height
                                                .Verticies(3).Y = .Verticies(1).Y + .Dimension.height
                                            Case "alignymiddle"
                                                        .Verticies(0).Y = ((frmMain.height / Screen.TwipsPerPixelY) / 2) - (.Dimension.height / 2)
                                                        .Verticies(1).Y = ((frmMain.height / Screen.TwipsPerPixelY) / 2) - (.Dimension.height / 2)
                                                .Verticies(2).Y = .Verticies(0).Y + .Dimension.height
                                                .Verticies(3).Y = .Verticies(1).Y + .Dimension.height
                                            Case "alignybottom"
                                                        .Verticies(0).Y = (frmMain.height / Screen.TwipsPerPixelY) - .Padding - .Dimension.height
                                                        .Verticies(1).Y = (frmMain.height / Screen.TwipsPerPixelY) - .Padding - .Dimension.height
                                                .Verticies(2).Y = .Verticies(0).Y + .Dimension.height
                                                .Verticies(3).Y = .Verticies(1).Y + .Dimension.height
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                    Case "camera"
                        inData = RemoveQuotedArg(inText, "{", "}", True)

                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 5)
                                    Select Case inArg(0)
                                        Case "location"
                                            CameraCount = CameraCount + 1
                                            ReDim Preserve Cameras(1 To CameraCount) As MyCamera
                                            Cameras(CameraCount).Location = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                            Cameras(CameraCount).Angle = CSng(inArg(4))
                                            Cameras(CameraCount).Pitch = CSng(inArg(5))
                            
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                ElseIf Left(Trim(inData), 1) = "[" Then
                                    AddMessage "Warning, Itemless Brackets at Line " & inLine
                                    GoTo throwerror
                                End If
                            End If
                        Loop
    
                    Case "method"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        MethodCount = MethodCount + 1
                        ReDim Preserve Methods(1 To MethodCount) As MyMethod
                        With Methods(MethodCount)
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 1)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "script"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .Script = RemoveQuotedArg(inData, "[", "]", True)
                                                    .Script = inLine & ":" & .Script
                                                Else
                                                    AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                        
                    Case "variable"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        VariableCount = VariableCount + 1
                        ReDim Preserve Variables(1 To VariableCount) As MyVariable
                        With Variables(VariableCount)
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 1)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "value"
                                                If InStr(inArg(1), """") > 0 Then
                                                    .Value = CVar(RemoveQuotedArg(inName, """", """"))
                                                Else
                                                     .Value = CVar(inArg(1))
                                                End If
                                            Case "onedit"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .OnEdit = RemoveQuotedArg(inData, "[", "]", True)
                                                    .OnEdit = inLine & ":" & .OnEdit
                                                Else
                                                    AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                    Case "database"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 0)
                                    Select Case inArg(0)
                                        Case "bindings"
                                            If InStr(inData, "[") > 0 Then
                                                inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                ParseBind inLine, RemoveQuotedArg(inData, "[", "]", True)
                                            Else
                                                AddMessage "Warning, Brackets Required at Line " & inLine
                                            End If
                                        Case "deserialize"
                                            If InStr(inData, "[") > 0 Then
                                                inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                Deserialize = RemoveQuotedArg(inData, "[", "]", True)
                                                Deserialize = inLine & ":" & Deserialize
                                            Else
                                                AddMessage "Warning, Brackets Required at Line " & inLine
                                            End If
                                        Case "serialize"
                                            If InStr(inData, "[") > 0 Then
                                                inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                Serialize = RemoveQuotedArg(inData, "[", "]", True)
                                                Serialize = inLine & ":" & Serialize
                                            Else
                                                AddMessage "Warning, Brackets Required at Line " & inLine
                                            End If
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                ElseIf Left(Trim(inData), 1) = "[" Then
                                    AddMessage "Warning, Itemless Brackets at Line " & inLine
                                    GoTo throwerror
                                End If
                            End If
                        Loop
                    Case Else
                        
                        Dim inVar As String
                        Dim inVal As Variant
                        inVal = RemoveArg(Trim(inItem), " ")
                        inVar = NextArg(Trim(inItem), " ")
                        If Left(inVar, 1) = ";" Then
                            inName = RemoveNextArg(inText, vbCrLf)
                            inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                        ElseIf Left(inVar, 1) = "!" Then
                            Process Mid(inVar, 2) & " " & NextArg(RemoveArg(inItem, " "), vbCrLf)
                        ElseIf Left(inVar, 1) = "=" Then
                            If inVal = "" Then inVal = Mid(inVar, 2)
                            ParseLand = ParseLand & Mid(inVar, 2) & " " & ParseSetGet(inLine, inVal) & vbCrLf
                        ElseIf Left(inVar, 1) = "$" Then
                            ParseSetGet inLine, inVar, ParseExpr(inLine, inVal)
                        ElseIf Left(inVar, 1) = "&" Then
                            If (MethodCount > 0) Then
                                If InStr(Mid(inVar, 2), ".") > 0 Then
                                    inVar = ParseSetGet(inLine, inVar, inVal)
                                    'inVal = RemoveArg(inVar, ".")
                                    'inVar = NextArg(inVar, ".")
                                Else
                                    For cnt = 1 To MethodCount
                                        If LCase(Methods(cnt).Identity) = LCase(Mid(inVar, 2)) Then
                                            inVal = Methods(cnt).Script
                                            If Not (inVal = "") Then
                                                ParseLand NextArg(inVal, ":"), RemoveArg(inVal, ":")
                                            End If
                                            inVar = ""
                                        End If
                                    Next
                                End If
                            End If
                            
                            If Not inVar = "" Then
                                AddMessage "Warning, Unknown Method at Line " & inLine
                            End If
                            
                        ElseIf inVar = "if" Then
                            inText = inVar & " " & inVal & inText
                            inLine = (NumLines - CountWord(inText, vbCrLf))
    
                            Dim ifCode As String
                            Dim inExp As String
                            Dim inIs As Variant
                            Dim elseCode As String
                            Dim useexp As Boolean
                            Dim lnum As Long
                            Dim calls As String
                            inName = RemoveLineArg(inText, "[")
                            inName = Trim(Mid(Trim(inName), 3))
                            
                            If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                inExp = RemoveQuotedArg(inName, "(", ")", True)
                                If LCase(Left(Trim(inName), 2)) = "is" Then
                                    inName = Trim(Mid(Trim(inName), 3))
                                    If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                        inIs = RemoveQuotedArg(inName, "(", ")", True)
                                    ElseIf Not (Left(Replace(Trim(inVar), vbTab, ""), 1) = ";") Then
                                        AddMessage "Error, Is Expression at Line " & inLine
                                    End If
                                Else
                                    inIs = True
                                End If
                                inData = RemoveQuotedArg(inText, "[", "]", True)
                                
                                If ParseExpr(inLine, inExp) = ParseSetGet(inLine, inIs) Then
                                
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Mid(Trim(inText), 3)
                                    Loop
                                    inText = Trim(inText)
                                    inLine = (NumLines - CountWord(inText, vbCrLf))
                                    
                                    useexp = True
                                    calls = inLine & ":" & inData
    
                                Else
    
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Trim(Mid(Trim(inText), 3))
                                    Loop
                                    inText = Trim(inText)
                                    inLine = (NumLines - CountWord(inText, vbCrLf))
                                        
                                    Do While (NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif") And (Not useexp)
                                        inName = RemoveLineArg(inText, "[")
                                        inName = Trim(Mid(Trim(inName), 7))
    
                                        If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                            inExp = RemoveQuotedArg(inName, "(", ")", True)
                                            If LCase(Left(Trim(inName), 2)) = "is" Then
                                                inName = Trim(Mid(Trim(inName), 3))
                                                If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                                    inIs = RemoveQuotedArg(inName, "(", ")", True)
                                                Else
                                                    AddMessage "Error, Is Expression at Line " & inLine
                                                End If
                                            Else
                                                inIs = True
                                            End If
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                        
                                            If ParseExpr(inLine, inExp) = ParseSetGet(inLine, inIs) Then
                                                useexp = True
                                                calls = inLine & ":" & inData
                                            End If
                                            
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                            Loop
                                            inText = Trim(inText)
                                            inLine = (NumLines - CountWord(inText, vbCrLf))
                                        Else
                                            AddMessage "Error, If Expression at Line " & inLine
                                        End If
                                    Loop
                                    If Not useexp Then
    
                                        Do While Left(Trim(inText), 2) = vbCrLf
                                            inText = Trim(Mid(Trim(inText), 3))
                                        Loop
                                        inText = Trim(inText)
                                        inLine = (NumLines - CountWord(inText, vbCrLf))
                                            
                                        If NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else" Then
                                            inName = RemoveLineArg(inText, "[")
    
                                            inName = Trim(Mid(Trim(inName), 5))
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                                inLine = inLine + 1
                                            Loop
                                            inText = Trim(inText)
                                            
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                            calls = inLine & ":" & inData
    
                                        End If
                                    End If
                                    
                                End If
                                
                                If calls <> "" Then
                                    ParseLand NextArg(calls, ":"), RemoveArg(calls, ":")
                                End If
    
                                Do While Left(Trim(inText), 2) = vbCrLf
                                    inText = Trim(Mid(Trim(inText), 3))
                                Loop
                                inText = Trim(inText)
                                    
                                If useexp And ((NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif") Or (NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else")) Then
    
                                    Do While (NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif")
                                        inName = RemoveLineArg(inText, "[")
                                        inName = Trim(Mid(Trim(inName), 7))
                                        If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                            inExp = RemoveQuotedArg(inName, "(", ")", True)
                                            If LCase(Left(Trim(inName), 2)) = "is" Then
                                                inName = Trim(Mid(Trim(inName), 3))
                                                If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                                    inIs = RemoveQuotedArg(inName, "(", ")", True)
                                                Else
                                                    AddMessage "Error, Is Expression at Line " & inLine
                                                End If
                                            Else
                                                inIs = True
                                            End If
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                            
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                            Loop
                                            inText = Trim(inText)
                                            
                                        Else
                                            AddMessage "Error, If Expression at Line " & inLine
                                        End If
                                    Loop
                                    
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Trim(Mid(Trim(inText), 3))
                                    Loop
                                    inText = Trim(inText)
                                            
                                    If (NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else") Then
                                        inName = RemoveLineArg(inText, "[")
    
                                        inName = Trim(Mid(Trim(inName), 5))
                                        inData = RemoveQuotedArg(inText, "[", "]", True)
    
                                    End If
        
                                End If
                            Else
                                AddMessage "Error, If Expression at Line " & inLine
                            End If
                            
                        Else
                            AddMessage "Warning, Unknown Object at Line " & inLine
                        End If
                End Select
            End If
        End If
    Loop
    
    Exit Function
parseerror:
    AddMessage "Script Error at Line " & inLine
throwerror:
    If Not ConsoleVisible Then ConsoleToggle
    Err.Clear
End Function
Public Function ParseBind(ByVal inLine As Long, ByVal inExp As String) As Variant
    Dim bind As String
    Dim idx As Integer
    
    Do Until inExp = ""
        bind = RemoveNextArg(inExp, vbCrLf)
        If bind <> "" Then
            idx = GetBindingIndex(NextArg(bind, "="))
            If idx > -1 Then
                Bindings(idx) = RemoveArg(bind, "=")
            Else
                AddMessage "Warning, Invalid Bind at Line " & inLine
            End If
        End If
        inLine = inLine + 1
    Loop
End Function

Public Function ParseExpr(ByVal inLine As Long, ByVal inExp As String) As Variant
    Dim exp As Variant
    Dim opr As String
    Dim val As Variant
    
    If InStr(inExp, "(") > 0 Then
        opr = Left(inExp, InStr(inExp, "(") - 1)
        opr = opr & " " & RemoveQuotedArg(inExp, "(", ")", True)
        inLine = inLine + CountWord(opr, vbCrLf)
        inExp = opr & " " & inExp
        opr = ""
    End If
    
    Do
        
        If opr = "" Then opr = RemoveNextArg(inExp, " ")
        If opr = "not" Then
            val = RemoveNextArg(inExp, " ")
            exp = Not ParseSetGet(inLine, val)
        Else
            Select Case LCase(opr)
                Case "or"
                    val = RemoveNextArg(inExp, " ")
                    exp = exp Or ParseSetGet(inLine, val)
                Case "and"
                    val = RemoveNextArg(inExp, " ")
                    exp = exp And ParseSetGet(inLine, val)
                Case Else
                    exp = ParseSetGet(inLine, opr)
            End Select
        End If
    
        opr = RemoveNextArg(inExp, " ")
        
    Loop Until inExp = ""
        
    ParseExpr = exp
End Function
Public Function ParseValues(ByVal inText As String) As String
    Dim outText As String
    
    Do Until inText = ""
        If InStr(inText, "$") > 0 Then
            outText = outText & Left(inText, InStr(inText, "$") - 1)
            inText = Mid(inText, InStr(inText, "$"))
            If InStr(inText, " ") > 0 Then
                outText = outText & ParseSetGet(0, Left(inText, InStr(inText, " ") - 1))
                inText = Mid(inText, InStr(inText, " ") + 1)
            Else
                outText = outText & ParseSetGet(0, inText)
                inText = ""
            End If
        Else
            outText = outText & inText
            inText = ""
        End If
    Loop
    ParseValues = outText
End Function
Public Function ParseSetGet(ByVal inLine As Long, ByVal inItem As Variant, Optional ByVal SetValue As Variant = Empty) As Variant
    
    If ((Left(Trim(inItem), 1) = """") And (Right(Trim(inItem), 1) = """")) Then
        
        ParseSetGet = RemoveQuotedArg(CStr(inItem), """", """")
    
    ElseIf ((Left(Trim(inItem), 1) = "$") Or (Left(Trim(inItem), 1) = "&")) And InStr(inItem, ".") > 0 Then
        
        ParseSetGet = SetValue
        
        Dim inProp As String
        Dim cnt As Long
        Dim cnt2 As Long
    
        inProp = Trim(Mid(inItem, InStr(inItem, ".") + 1))
        inItem = Mid(Left(inItem, InStr(inItem, ".") - 1), 2)
      
        If (SoundCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To SoundCount
                If LCase(Sounds(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "enable"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Sounds(cnt).Enable)
                            Else
                                Sounds(cnt).Enable = CBool(SetValue)
                            End If
                        Case "play"
                            PlayWave cnt, Sounds(cnt).Repeat
                        Case "stop"
                            StopWave cnt
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If

        If (TrackCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To TrackCount
                If LCase(Tracks(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "enable"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Tracks(cnt).Enable)
                            Else
                                Tracks(cnt).Enable = CBool(SetValue)
                            End If
                        Case "play"
                            Tracks(cnt).PlaySound
                        Case "stop"
                            Tracks(cnt).StopSound
                        Case "fadein"
                            Tracks(cnt).FadeIn
                        Case "fadeout"
                            Tracks(cnt).FadeOut
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If
        
        If (VariableCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To VariableCount
                If LCase(Variables(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "value"
                            If SetValue = Empty Then
                                ParseSetGet = CVar(Variables(cnt).Value)
                            Else
                                Variables(cnt).Value = CVar(SetValue)
                                If Not Variables(cnt).OnEdit = "" Then
                                    ParseLand NextArg(Variables(cnt).OnEdit, ":"), RemoveArg(Variables(cnt).OnEdit, ":")
                                End If
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If

        If (LightDataCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To LightDataCount
                If LCase(LightDatas(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "x"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(LightDatas(cnt).Origin.X)
                            Else
                                LightDatas(cnt).Origin.X = CSng(SetValue)
                            End If
                        Case "y"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(LightDatas(cnt).Origin.Y)
                            Else
                                LightDatas(cnt).Origin.Y = CSng(SetValue)
                            End If
                        Case "z"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(LightDatas(cnt).Origin.z)
                            Else
                                LightDatas(cnt).Origin.z = CSng(SetValue)
                            End If
                        Case "enabled"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(LightDatas(cnt).Enabled)
                            Else
                                LightDatas(cnt).Enabled = CBool(SetValue)
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If

        If (BillBoardCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To BillBoardCount
                If LCase(BillBoards(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "visible"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(BillBoards(cnt).Visible)
                            Else
                                BillBoards(cnt).Visible = CBool(SetValue)
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If

        If (ScreenImageCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To ScreenImageCount
                If LCase(ScreenImages(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "visible"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(ScreenImages(cnt).Visible)
                            Else
                                ScreenImages(cnt).Visible = CBool(SetValue)
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If
      
        If (ObjectCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To ObjectCount
                If LCase(Objects(cnt).Identity) = LCase(inItem) Then
                    Select Case LCase(inProp)
                        Case "rotate.x"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Rotate.X)
                            Else
                                Objects(cnt).Rotate.X = CSng(SetValue)
                            End If
                        Case "rotate.y"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Rotate.Y)
                            Else
                                Objects(cnt).Rotate.Y = CSng(SetValue)
                            End If
                        Case "rotate.z"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Rotate.z)
                            Else
                                Objects(cnt).Rotate.z = CSng(SetValue)
                            End If
                            
                        Case "x", "origin.x"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Origin.X)
                            Else
                                Objects(cnt).Origin.X = CSng(SetValue)
                            End If
                        Case "y", "origin.y"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Origin.Y)
                            Else
                                Objects(cnt).Origin.Y = CSng(SetValue)
                            End If
                        Case "z", "origin.z"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Objects(cnt).Origin.z)
                            Else
                                Objects(cnt).Origin.z = CSng(SetValue)
                            End If
                        Case "gravitational"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Objects(cnt).Gravitational)
                            Else
                                Objects(cnt).Gravitational = CBool(SetValue)
                            End If
                        Case "wireframe"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Objects(cnt).WireFrame)
                            Else
                                Objects(cnt).WireFrame = CBool(SetValue)
                            End If
                        Case "visible"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Objects(cnt).Visible)
                            Else
                                Objects(cnt).Visible = CBool(SetValue)
                            End If
                        Case Else
                            If Objects(cnt).ActivityCount And (Not inItem = "$") Then
                                For cnt2 = 1 To Objects(cnt).ActivityCount
                                    If LCase(Objects(cnt).Activities(cnt2).Identity) = LCase(NextArg(inProp, ".")) Then
                                        Select Case RemoveArg(inProp, ".")
                                            Case "power"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Objects(cnt).Activities(cnt2).Emphasis)
                                                Else
                                                    Objects(cnt).Activities(cnt2).Emphasis = CSng(SetValue)
                                                End If
                                            Case "drag"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Objects(cnt).Activities(cnt2).Friction)
                                                Else
                                                    Objects(cnt).Activities(cnt2).Friction = CSng(SetValue)
                                                End If
                                            Case "reactive"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Objects(cnt).Activities(cnt2).Reactive)
                                                Else
                                                    Objects(cnt).Activities(cnt2).Reactive = CSng(SetValue)
                                                End If
                                            Case "recount"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Objects(cnt).Activities(cnt2).Recount)
                                                Else
                                                    Objects(cnt).Activities(cnt2).Recount = CSng(SetValue)
                                                End If
                                            Case Else
                                                AddMessage "Warning, Unknown Sub Entity " & RemoveArg(inProp, ".")
                                        End Select
                                        inItem = "$"
                                    End If
                                Next
                            Else
                                AddMessage "Warning, Unknown Sub Entity " & inProp
                            End If
                            
                    End Select
                    inItem = "$"
                End If

            Next
        End If
        
        If (BeaconCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To BeaconCount
                If LCase(Beacons(cnt).Identity) = LCase(inItem) Then
                    Select Case NextArg(inProp, " ")
                        Case "visible"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Beacons(cnt).Visible)
                            Else
                                Beacons(cnt).Visible = CBool(SetValue)
                            End If
                        Case "x"
                            If Beacons(cnt).OriginCount > 0 Then
                                If SetValue = Empty Then
                                    ParseSetGet = CSng(Beacons(cnt).Origins(1).X)
                                Else
                                    Beacons(cnt).Origins(1).X = CSng(SetValue)
                                End If
                            End If
                        Case "y"
                            If Beacons(cnt).OriginCount > 0 Then
                                If SetValue = Empty Then
                                    ParseSetGet = CSng(Beacons(cnt).Origins(1).Y)
                                Else
                                    Beacons(cnt).Origins(1).Y = CSng(SetValue)
                                End If
                            End If
                        Case "z"
                            If Beacons(cnt).OriginCount > 0 Then
                                If SetValue = Empty Then
                                    ParseSetGet = CSng(Beacons(cnt).Origins(1).z)
                                Else
                                    Beacons(cnt).Origins(1).z = CSng(SetValue)
                                End If
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If
        
        If (PortalCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To PortalCount
                If LCase(Portals(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "x"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Portals(cnt).Location.X)
                            Else
                                Portals(cnt).Location.X = CSng(SetValue)
                            End If
                        Case "y"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Portals(cnt).Location.Y)
                            Else
                                Portals(cnt).Location.Y = CSng(SetValue)
                            End If
                        Case "z"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Portals(cnt).Location.z)
                            Else
                                Portals(cnt).Location.z = CSng(SetValue)
                            End If
                        Case "enable"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Portals(cnt).Enable)
                            Else
                                Portals(cnt).Enable = CBool(SetValue)
                            End If
                        Case Else
                            If Portals(cnt).ActivityCount And (Not inItem = "$") Then
                                For cnt2 = 1 To Portals(cnt).ActivityCount
                                    If LCase(Portals(cnt).Activities(cnt2).Identity) = LCase(NextArg(inProp, ".")) Then
                                        Select Case RemoveArg(inProp, ".")
                                            Case "power"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Portals(cnt).Activities(cnt2).Emphasis)
                                                Else
                                                    Portals(cnt).Activities(cnt2).Emphasis = CSng(SetValue)
                                                End If
                                            Case "drag"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Portals(cnt).Activities(cnt2).Friction)
                                                Else
                                                    Portals(cnt).Activities(cnt2).Friction = CSng(SetValue)
                                                End If
                                            Case "reactive"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Portals(cnt).Activities(cnt2).Reactive)
                                                Else
                                                    Portals(cnt).Activities(cnt2).Reactive = CSng(SetValue)
                                                End If
                                            Case "recount"
                                                If SetValue = Empty Then
                                                    ParseSetGet = CSng(Portals(cnt).Activities(cnt2).Recount)
                                                Else
                                                    Portals(cnt).Activities(cnt2).Recount = CSng(SetValue)
                                                End If
                                            Case Else
                                                AddMessage "Warning, Unknown Sub Entity " & RemoveArg(inProp, ".")
                                        End Select
                                        inItem = "$"
                                    End If
                                Next
                            Else
                                AddMessage "Warning, Unknown Sub Entity " & inItem
                            End If
                            
                    End Select
                    inItem = "$"
                End If
'                If Portals(cnt).ActivityCount And (Not inItem = "$") Then
'                    For cnt2 = 1 To Portals(cnt).ActivityCount
'                        If LCase(Portals(cnt).Activities(cnt2).Identity) = LCase(inItem) Then
'                            Select Case inProp
'                                Case Else
'                                    AddMessage "Warning, Unknown Sub Entity " & inItem
'                            End Select
'                            inItem = "$"
'                        End If
'                    Next
'                End If
            Next
        End If
        
        If (CameraCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To CameraCount
                If LCase(Cameras(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "x"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Cameras(cnt).Location.X)
                            Else
                                Cameras(cnt).Location.X = CSng(SetValue)
                            End If
                        Case "y"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Cameras(cnt).Location.Y)
                            Else
                                Cameras(cnt).Location.Y = CSng(SetValue)
                            End If
                        Case "z"
                            If SetValue = Empty Then
                                ParseSetGet = CSng(Cameras(cnt).Location.z)
                            Else
                                Cameras(cnt).Location.z = CSng(SetValue)
                            End If
                        Case "angle"
                             If SetValue = Empty Then
                                ParseSetGet = CSng(Cameras(cnt).Angle)
                            Else
                                Cameras(cnt).Angle = CSng(SetValue)
                            End If
                        Case "pitch"
                             If SetValue = Empty Then
                                ParseSetGet = CSng(Cameras(cnt).Pitch)
                            Else
                                Cameras(cnt).Pitch = CSng(SetValue)
                            End If
                        Case Else
                            AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If

        If LCase(Player.Object.Identity) = LCase(inItem) Then
            Select Case inProp
                Case "visible"
                    If SetValue = Empty Then
                        ParseSetGet = CBool(Player.Object.Visible)
                    Else
                        Player.Object.Visible = CBool(SetValue)
                    End If
                Case "x"
                    If SetValue = Empty Then
                        ParseSetGet = CSng(Player.Object.Origin.X)
                    Else
                        Player.Object.Origin.X = CSng(SetValue)
                    End If
                Case "y"
                    If SetValue = Empty Then
                        ParseSetGet = CSng(Player.Object.Origin.Y)
                    Else
                        Player.Object.Origin.Y = CSng(SetValue)
                    End If
                Case "z"
                    If SetValue = Empty Then
                        ParseSetGet = CSng(Player.Object.Origin.z)
                    Else
                        Player.Object.Origin.z = CSng(SetValue)
                    End If
                Case "angle"
                     If SetValue = Empty Then
                        ParseSetGet = CSng(Player.CameraAngle)
                    Else
                        Player.CameraAngle = CSng(SetValue)
                    End If
                Case "pitch"
                     If SetValue = Empty Then
                        ParseSetGet = CSng(Player.CameraPitch)
                    Else
                        Player.CameraPitch = CSng(SetValue)
                    End If
                Case "view"
                     If SetValue = Empty Then
                        ParseSetGet = CSng(Perspective)
                    Else
                        Perspective = CSng(SetValue)
                    End If
                Case "zoom"
                     If SetValue = Empty Then
                        ParseSetGet = CSng(Player.CameraZoom)
                    Else
                        Player.CameraZoom = CSng(SetValue)
                    End If
                Case Else
                    AddMessage "Warning, Unknown Sub Entity " & inItem
            End Select
            inItem = "$"
        End If
                
        If Not inItem = "$" Then
            AddMessage "Warning, Unkown Identity " & inItem
        End If
    Else
        ParseSetGet = inItem
    End If
End Function

Public Sub CleanupLand(Optional ByVal NoSerialize As Boolean = False)

    ShowCredits = False
    
    Dim uid As Long
    Dim ser As String
    
    If Serialize <> "" Then
        ser = IIf(Not NoSerialize, ParseLand(NextArg(Serialize, ":"), RemoveArg(Serialize, ":")), "")
        Do While InStr(ser, ";") > 0
            If InStr(InStr(ser, ";"), ser, vbCrLf) > InStr(ser, ";") Then
                ser = Left(ser, InStr(ser, ";") - 1)
            Else
                ser = Left(ser, InStr(ser, ";") - 1) & Mid(ser, InStr(InStr(ser, ";"), ser, vbCrLf) + 2)
                
            End If
        
        Loop

        db.rsQuery rs, "SELECT * FROM Serials WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND PXFile='" & Replace(CurrentLoadedLevel, "'", "''") & "';"
        If Not db.rsEnd(rs) Then
            db.dbQuery "UPDATE Serials SET Script = '" & Replace(ser, "'", "''") & "' WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND PXFile='" & Replace(CurrentLoadedLevel, "'", "''") & "';"
        Else
            db.dbQuery "INSERT INTO Serials (Username, PXFile, Script) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "', '" & Replace(CurrentLoadedLevel, "'", "''") & "', '" & Replace(ser, "'", "''") & "');"
        End If

    End If
    
    Dim q As Integer
    Dim o As Integer
    
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            If Objects(o).FolcrumCount > 0 Then
                Erase Objects(o).Folcrum
                Objects(o).FolcrumCount = 0
            End If
            If Objects(o).ActivityCount > 0 Then
                Erase Objects(o).Activities
                Objects(o).ActivityCount = 0
            End If
            If Not Objects(o).ReplacerVals Is Nothing Then
                Do Until Objects(o).ReplacerVals.Count = 0
                    Objects(o).ReplacerVals.Remove 1
                Loop
            End If
            If Not Objects(o).ReplacerKeys Is Nothing Then
                Do Until Objects(o).ReplacerKeys.Count = 0
                    Objects(o).ReplacerKeys.Remove 1
                Loop
            End If
        Next
        Erase Objects
        ObjectCount = 0
    End If
    
    If MeshCount > 0 Then
        For o = 1 To MeshCount
        
            If Meshes(o).MaterialCount > 0 Then
                For q = LBound(Meshes(o).Textures) To UBound(Meshes(o).Textures)
                    Set Meshes(o).Textures(q) = Nothing
                Next q
    
                Erase Meshes(o).Materials
                Erase Meshes(o).Textures
                Meshes(o).MaterialCount = 0
            End If
            Set Meshes(o).Mesh = Nothing
            Set Meshes(o).MaterialBuffer = Nothing
        Next
        Erase Meshes
        MeshCount = 0
    End If
    
    If BillBoardCount > 0 Then
        For o = 1 To BillBoardCount
            If UBound(BillBoards(o).Skin) > 0 Then
                For q = 1 To UBound(BillBoards(o).Skin)
                    Set BillBoards(o).Skin(q) = Nothing
                Next q
            End If
            Set BillBoards(o).VBuf = Nothing
            Erase BillBoards(o).Skin
            BillBoards(o).SkinCount = 0
            Erase BillBoards(o).Plaq
        Next
        Erase BillBoards
        BillBoardCount = 0
    End If
    
    If LightDataCount > 0 Then
        For o = 1 To LightDataCount
            DDevice.LightEnable o, False
        Next
        Erase LightDatas
        LightDataCount = 0
    End If
    
    If LightCount > 0 Then
        Erase Lights
        LightCount = 0
    End If

    
    If BeaconCount > 0 Then
        For q = 1 To BeaconCount
            If Beacons(q).OriginCount > 0 Then
                Erase Beacons(q).Origins
                Beacons(q).OriginCount = 0
            End If
            If Beacons(q).BeaconSkinCount > 0 Then
                For o = 1 To Beacons(q).BeaconSkinCount
                    Set Beacons(q).BeaconSkin(o) = Nothing
                Next
                Erase Beacons(q).BeaconSkin
                Beacons(q).BeaconSkinCount = 0
            End If
            Set Beacons(q).BeaconVBuf = Nothing
        Next
        Erase Beacons
        BeaconCount = 0
    End If
    
    Set SkySkin(0) = Nothing
    Set SkySkin(1) = Nothing
    Set SkySkin(2) = Nothing
    Set SkySkin(3) = Nothing
    Set SkySkin(4) = Nothing
    Set SkySkin(5) = Nothing
    
    Set SkyVBuf = Nothing

    If ScreenImageCount > 0 Then
        For o = 1 To ScreenImageCount
            Set ScreenImages(o).Image = Nothing
        Next

        Erase ScreenImages
        ScreenImageCount = 0
    End If
    
    If PortalCount > 0 Then
        For o = 1 To PortalCount
            If Portals(o).ActivityCount > 0 Then
                Erase Portals(o).Activities
                Portals(o).ActivityCount = 0
            End If
        Next
        
        Erase Portals
        PortalCount = 0
    End If

    If CameraCount > 0 Then
        Erase Cameras
        CameraCount = 0
    End If

    If VariableCount > 0 Then
        Erase Variables
        VariableCount = 0
    End If
    
    If MethodCount > 0 Then
        Erase Methods
        MethodCount = 0
    End If
    
    If TrackCount > 0 Then
        For o = 1 To TrackCount
            Tracks(o).StopSound
            Set Tracks(o) = Nothing
        Next
        
        Erase Tracks
        TrackCount = 0
    End If
    
    If SoundCount > 0 Then
    
        For o = 1 To SoundCount
            StopWave Sounds(o).Index
            Set Waves(o) = Nothing
        Next
        Erase Waves
        Erase Sounds
        SoundCount = 0
    End If
End Sub



