Attribute VB_Name = "modLand"
#Const modLand = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

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

'###########################################################################
'###################### BEGIN UNIQUE NON GLOBALS ###########################
'###########################################################################


Public Meshes() As MyMesh
Public MeshCount As Long


Private Serialize As String
Private Deserialize As String

Public DXLights() As D3DLIGHT8


Public BackgroundColor As Long

Public GlobalGravityDirect As New Motion
Public GlobalGravityRotate As New Motion
Public GlobalGravityScaled As New Motion

Public LiquidGravityDirect As New Motion
Public LiquidGravityRotate As New Motion
Public LiquidGravityScaled As New Motion



Public matWorld As D3DMATRIX

Private matProj As D3DMATRIX
Private matView As D3DMATRIX
Private matViewSave As D3DMATRIX


Private Sub RenderSpacesSetup()
    DDevice.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
    DDevice.SetTransform D3DTS_VIEW, matView
    D3DXMatrixPerspectiveFovLH matProj, PI / 2.5, AspectRatio, 0.05, 50
    DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

Private Sub RenderSpacesClose()
        
    D3DXMatrixPerspectiveFovLH matProj, FOVY, AspectRatio, 0.05, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
    DDevice.SetTransform D3DTS_VIEW, matViewSave
End Sub

Public Function GetBoardKey(ByRef Obj As Element, ByVal TextName As String) As String
    If Not Obj.ReplacerKeys Is Nothing Then
        If Obj.ReplacerKeys.Count > 0 Then
            Dim i As Long
            For i = 1 To Obj.ReplacerKeys.Count
                If Obj.ReplacerKeys(i) = Obj.Key & "_" & Replace(TextName, ".", "") Then
                    GetBoardKey = Obj.ReplacerVals(Obj.Key & "_" & Replace(TextName, ".", ""))
                    Exit Function
                End If
            Next
        End If
    End If
End Function

Public Sub RenderPlayer()

    If ((Perspective = Playmode.ThirdPerson) Or (Perspective = Playmode.CameraMode)) And (Not DebugMode) Then
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
        DDevice.SetVertexShader FVF_RENDER
        
        Player.Element.PlayerMatrix
        
        If Player.Element.Visible Then
            DDevice.SetRenderState D3DRS_FILLMODE, IIf(Player.Element.WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
            
            Dim i As Long
            If Player.Element.VisualIndex > 0 Then
                With Meshes(Player.Element.VisualIndex)
                    If .MaterialCount > 0 Then
                        For i = 0 To .MaterialCount - 1
            
                            If .Textures(i) Is Nothing Then
                                DDevice.SetPixelShader PixelShaderDefault
                                DDevice.SetMaterial .Materials(i)
                                DDevice.SetTexture 0, Nothing
                                DDevice.SetMaterial GenericMaterial
                                DDevice.SetTexture 1, Nothing
                                .Mesh.DrawSubset i
                            Else
            
                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
            
                                DDevice.SetPixelShader PixelShaderDefault
                                DDevice.SetMaterial .Materials(i)
                                DDevice.SetTexture 0, .Textures(i)
                                DDevice.SetMaterial GenericMaterial
                                DDevice.SetTexture 1, .Textures(i)
                                .Mesh.DrawSubset i
            
                            End If
            
                        Next
                    End If
                End With
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
    Dim bkey As String

    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
                        
    If Lights.Count > 0 Then
        Dim l1 As Light
        l = 1
        For Each l1 In Lights


            If l1.LightType = Lighting.Directed Or (l1.Enabled And _
                DistanceEx(Player.Element.Origin, l1.Origin) <= (FadeDistance - l1.Range)) Then
                
                If (l1.LightBlink > 0) Or (l1.DiffuseRoll <> 0) Then
                    If (l1.LightBlink > 0) Then
                        If (l1.LightTimer = 0) Or ((Timer - l1.LightTimer) >= l1.LightBlink And (l1.LightBlink > 0)) Then
                            l1.LightTimer = Timer
                            l1.LightIsOn = Not l1.LightIsOn
                        End If
                        DDevice.LightEnable (l - 1), l1.LightIsOn
                    End If
                    If (l1.DiffuseRoll <> 0) Then
                        If (l1.DiffuseTimer = 0) Or ((Timer - l1.DiffuseTimer) >= Abs(l1.DiffuseRoll) And (l1.DiffuseTimer > 0)) Then
                            l1.DiffuseTimer = Timer
                            If (l1.DiffuseRoll > 0) Then
                                If (l1.DIffuseMax > 0 And l1.DiffuseNow >= l1.DIffuseMax) Or (l1.DIffuseMax < 0 And l1.DiffuseNow >= -0.01) Then
                                    l1.DiffuseRoll = -l1.DiffuseRoll
                                Else
                                    l1.DiffuseNow = l1.DiffuseNow + 1
                                    l1.Diffuse.Red = l1.Diffuse.Red + 0.01
                                    l1.Diffuse.Green = l1.Diffuse.Green + 0.01
                                    l1.Diffuse.Blue = l1.Diffuse.Blue + 0.01
                                End If
                            Else
                                If (l1.DIffuseMax > 0 And l1.DiffuseNow <= 0.01) Or (l1.DIffuseMax < 0 And l1.DiffuseNow <= l1.DIffuseMax) Then
                                    l1.DiffuseRoll = -l1.DiffuseRoll
                                Else
                                    l1.DiffuseNow = l1.DiffuseNow - 1
                                    l1.Diffuse.Red = l1.Diffuse.Red - 0.01
                                    l1.Diffuse.Green = l1.Diffuse.Green - 0.01
                                    l1.Diffuse.Blue = l1.Diffuse.Blue - 0.01
                                End If
                            End If
                            
                            
                        End If


                        DDevice.SetLight (l - 1), DXLights(l)
                        DDevice.LightEnable (l - 1), 1
                    End If
                Else
                    DDevice.LightEnable l - 1, 1
                End If
            Else
                DDevice.LightEnable l - 1, False
            End If
            l = l + 1
        Next
    End If

    If Sounds.Count > 0 Then
        Dim s1 As Sound
        For Each s1 In Sounds
            If s1.Range > 0 And s1.Enabled Then
                r = DistanceEx(Player.Element.Origin, s1.Origin)
                If r < s1.Range Then
                    
                    r = Round(CSng(s1.Range - Dist), 3)
                    r = Abs(-s1.Range + r)
                
                    VolumeWave l, r
                    PlayWave l, s1.Repeat
                    
                Else
                    StopWave l
                    
                End If
            End If
        Next
    End If
    
    If Tracks.Count > 0 Then
        Dim t1 As Track
        For Each t1 In Tracks
            If t1.Range > 0 And t1.Enabled Then
                r = DistanceEx(Player.Element.Origin, t1.Origin)
                If r < t1.Range Then
                
                    r = Round(CSng(t1.Range - r), 3)
        
                    t1.Volume = (r * 10)
                    
                Else
                    t1.Volume = 0
                End If
            End If
        Next
    End If
    
    If Elements.Count > 0 Then
        Dim e1 As Element
        Dim b1 As Board
        For cnt = 1 To Elements.Count
            Set e1 = Elements(cnt)
            
        'For Each e1 In Elements

            If e1.Visible And (Not (e1.Effect = Collides.Ladder Or e1.Effect = Collides.Liquid)) Then
            
                If ((e1.BoundsIndex >= 0) And DistanceEx(Player.Element.Origin, e1.Origin) <= FadeDistance) Then
                    
                    If e1.VisualIndex > 0 Then
                        v = e1.VisualIndex
                    Else
                        v = e1.BoundsIndex
                    End If

                    If v > 0 Then
                    
                        'If DebugMode Or Meshes(e1.BoundsIndex).MaterialCount > 0 Then

                             e1.ApplyMatrix
    
                        'End If
                    End If
                        
                    If MeshCount > 0 And Not v = 0 Then
                    
                        With Meshes(v)
                            
                            If .MaterialCount > 0 Then
                                                            
                                DDevice.SetRenderState D3DRS_FILLMODE, IIf(e1.WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
                                
                                For i = 0 To .MaterialCount - 1
                                         
                                    If e1.Alphablend Then
                                                
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                            
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 0, .Textures(i)
                                        DDevice.SetMaterial GenericMaterial
                                        DDevice.SetTexture 1, .Textures(i)
                                     
                                        .Mesh.DrawSubset i
                                                     
                                     Else
                                         If Not (.Textures(i) Is Nothing) Then
                                             
                                             bkey = GetBoardKey(e1, D3DX.BufferGetTextureName(.MaterialBuffer, i))
                                             
                                             If (bkey <> "") Then
                                                
                                                Set b1 = Boards(bkey)
                                                
                                                 If Not b1.Translucent And Not b1.Alphablend Then
            
                                                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                                                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                                    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
                                                    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
                                                    
                                                    If (b1.Animated > 0) Then
                                                        If (b1.AnimateTimer = 0) Or (CDbl(Timer - b1.AnimateTimer) >= b1.Animated) Then
                                                            b1.AnimateTimer = GetTimer
                                                             
                                                            b1.AnimatePoint = b1.AnimatePoint + 1
                                                            If b1.AnimatePoint > b1.FileNames.Count Then
                                                                b1.AnimatePoint = 1
                                                            End If
                                                            
                                                        End If
                                    
                                                        DDevice.SetMaterial LucentMaterial
                                                        b1.SetTexture 0, b1.AnimatePoint
                                                        DDevice.SetMaterial GenericMaterial
                                                        b1.SetTexture 1, b1.AnimatePoint
                                    
                                                    Else
                                                        DDevice.SetMaterial LucentMaterial
                                                        b1.SetTexture 0, 1
                                                        DDevice.SetMaterial GenericMaterial
                                                        b1.SetTexture 1, 1
                                                    End If
            
                                                    .Mesh.DrawSubset i
                                                End If
                                                Set b1 = Nothing
                                             Else
                                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                                 
                                                DDevice.SetMaterial .Materials(i)
                                                DDevice.SetTexture 0, .Textures(i)
                                                DDevice.SetMaterial GenericMaterial
                                                DDevice.SetTexture 1, .Textures(i)
                                                .Mesh.DrawSubset i
                                                 
                                            End If
                                             
                                        End If
                                    End If
                                Next
                            End If
                        End With
                    End If
                End If
            End If
            Set e1 = Nothing
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

    If Boards.Count > 0 Then
        Dim b1 As Board
        For Each b1 In Boards
        
            If b1.Visible And Not b1.Mirror Then
                If DistanceEx(Player.Element.Origin, b1.Origin) <= FadeDistance Then
                    
                    If Not b1.Translucent Then
                    
                        If (b1.Animated > 0) Then
                            If CDbl(Timer - b1.AnimateTimer) >= b1.Animated Then
                                b1.AnimateTimer = GetTimer
                                
                                b1.AnimatePoint = b1.AnimatePoint + 1
                                If b1.AnimatePoint > b1.FileNames.Count Then
                                    b1.AnimatePoint = 1
                                End If
                            End If
                            DDevice.SetMaterial GenericMaterial
                            b1.SetTexture 0, b1.AnimatePoint
                            b1.SetTexture 1, -1
                        Else
                            DDevice.SetMaterial GenericMaterial
                            b1.SetTexture 0, 1
                            b1.SetTexture 1, -1
                        End If
                        b1.Render

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
    Dim bkey As String
    Dim cnt As Long
    
    'Dim matLucent As D3DMATRIX
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
        
    Dim b1 As Board
    If Elements.Count > 0 Then
        Dim e1 As Element
        For cnt = 1 To Elements.Count
            Set e1 = Elements(cnt)
        'For Each e1 In Elements
        
        'For o = 1 To Elements.count
            If e1.Visible And (Not DebugMode) Then
            
                If (e1.BoundsIndex > 0) And DistanceEx(Player.Element.Origin, e1.Origin) <= FadeDistance Then
                    With Meshes(e1.BoundsIndex)
                    
                        For i = 0 To .MaterialCount - 1
        
                            If Not (.Textures(i) Is Nothing) Then
                                
                                bkey = GetBoardKey(e1, D3DX.BufferGetTextureName(.MaterialBuffer, i))
                                
                                If (bkey <> "") Then
                                
                                    Set b1 = Boards(bkey)
                                    
                                    
                                    e1.SetWorldMatrix
                                    
                                    'D3DXMatrixIdentity matLucent
                                    'DDevice.SetTransform D3DTS_WORLD, matLucent
                                    'e1.PrepairMatrix
                                    
                                        
                                    If b1.Translucent Then
        
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
        
                                        If (b1.Animated > 0) Then
                                            If (b1.AnimateTimer = 0) Or (CDbl(Timer - b1.AnimateTimer) >= b1.Animated) Then
                                                b1.AnimateTimer = GetTimer
                                                
                                                b1.AnimatePoint = b1.AnimatePoint + 1
                                                If b1.AnimatePoint > b1.FileNames.Count Then
                                                    b1.AnimatePoint = 1
                                                End If
                                                
                                            End If
                               
                                                                
                                            DDevice.SetMaterial LucentMaterial
                                            b1.SetTexture 0, b1.AnimatePoint
                                            DDevice.SetMaterial GenericMaterial
                                            b1.SetTexture 1, b1.AnimatePoint
                        
                                        Else
                                            DDevice.SetMaterial LucentMaterial
                                            b1.SetTexture 0, 1
                                            DDevice.SetMaterial GenericMaterial
                                            b1.SetTexture 1, 1
                                        End If
        
                                        .Mesh.DrawSubset i
                                    ElseIf b1.Alphablend Then
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        
                                        If (b1.Animated > 0) Then
                                            If (b1.AnimateTimer = 0) Or (CDbl(Timer - b1.AnimateTimer) >= b1.Animated) Then
                                                b1.AnimateTimer = GetTimer
                                                
                                                b1.AnimatePoint = b1.AnimatePoint + 1
                                                If b1.AnimatePoint > b1.FileNames.Count Then
                                                    b1.AnimatePoint = 1
                                                End If
                                                
                                            End If
                        
                                            DDevice.SetMaterial LucentMaterial
                                            b1.SetTexture 0, b1.AnimatePoint
                                            DDevice.SetMaterial GenericMaterial
                                            b1.SetTexture 1, b1.AnimatePoint
                        
                                        Else
                                            DDevice.SetMaterial LucentMaterial
                                            b1.SetTexture 0, 1
                                            DDevice.SetMaterial GenericMaterial
                                            b1.SetTexture 1, 1
                                        End If
        
                                        .Mesh.DrawSubset i
                                        
                                    End If
                                    Set b1 = Nothing
                                End If
        
                            End If
        
                        Next
                    End With
                End If
            End If
            Set e1 = Nothing
        Next
    End If
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                                    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    If Boards.Count > 0 Then
        For Each b1 In Boards
        
        'For o = 1 To Boards.count
            If b1.Visible Then

                    If DistanceEx(Player.Element.Origin, b1.Origin) <= FadeDistance Then
                        If b1.Translucent Then
            
                            If (b1.Animated > 0) Then
                                If (b1.AnimateTimer = 0) Or (CDbl(Timer - b1.AnimateTimer) >= b1.Animated) Then
                                    b1.AnimateTimer = GetTimer
                                    
                                    b1.AnimatePoint = b1.AnimatePoint + 1
                                    If b1.AnimatePoint > b1.FileNames.Count Then
                                        b1.AnimatePoint = 1
                                    End If
                                    
                                End If
            
                                DDevice.SetMaterial LucentMaterial
                                b1.SetTexture 0, b1.AnimatePoint
                                b1.SetTexture 1, -1
            
                            Else
                                DDevice.SetMaterial LucentMaterial
                                b1.SetTexture 0, 1
                                b1.SetTexture 1, -1
                            End If
                            b1.Render

            
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
    Dim A As Single
    
    Dim X As Single
    Dim Z As Single
    Dim ok As Boolean
    
    Dim o As Long
    Dim d As Single
    Dim p As Point
    
    Dim matScale As D3DMATRIX
    Dim matPos As D3DMATRIX
    
    D3DXMatrixIdentity matWorld
   ' DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Dim matBeacon As D3DMATRIX
    
    Dim a1 As Beacon
    
    If Beacons.Count > 0 Then
        For Each a1 In Beacons
        
        'For o = 1 To Beacons.count
            If a1.Visible Then
            
                If a1.BeaconLight > -1 Then
                    If Lights.Count > 0 Then
                        For l = 1 To Lights.Count

                            DDevice.SetLight l - 1, DXLights(l)
                            If Lights(l).LightType = Lighting.Directed Then
                                DDevice.LightEnable l - 1, 1
                            Else
                                DDevice.LightEnable l - 1, False
                            End If
                        Next
                    End If
                End If
                
                If a1.Randomize Then
                    X = IIf((Rnd < 0.5), -RandomPositive(BeaconSpacing, BeaconRange), RandomPositive(BeaconSpacing, BeaconRange))
                    Z = IIf((Rnd < 0.5), -RandomPositive(BeaconSpacing, BeaconRange), RandomPositive(BeaconSpacing, BeaconRange))
                    ok = True
                Else
                    ok = False
                End If
                
                If a1.Origins.Count > 0 Then
                    l = 1
    
                    Do While l <= a1.Origins.Count
                        d = DistanceEx(a1.Origin(l), Player.Element.Origin)
                        If ok Then ok = ok And (DistanceEx(a1.Origin(l), MakePoint(X, 0, Z)) > BeaconSpacing)

                        If d <= FadeDistance Then
                            If a1.Consumable And (d <= 3) Then

                                a1.Origins.Remove l

    
                            ElseIf l <= a1.Origins.Count Then
    
    
                                D3DXMatrixIdentity matBeacon
                                
                                If (a1.RoundingCut = 0) Then

                                        
                                    If a1.VerticalLock Then
                                        
                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
                                            A = Cameras(Player.CameraIndex).Angle
                                            D3DXMatrixRotationYawPitchRoll matBeacon, -A, -Cameras(Player.CameraIndex).Pitch, 0
        
                                            D3DXMatrixScaling matScale, 1, 1, 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                            D3DXMatrixTranslation matPos, (a1.Origin(l).X - (Sin(D720 - A) * (1 / (PI * 2.5)))), a1.Origin(l).Y, (a1.Origin(l).Z - (Cos(D720 - A) * (1 / (PI * 2.5))))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                        Else


                                            
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle), -Player.Camera.Pitch, 0
                                            D3DXMatrixScaling matScale, 1, 1, 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            If Player.Camera.Pitch >= -1.5 And Player.Camera.Pitch < -0 Then
                                               D3DXMatrixTranslation matPos, (a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 2.5)))), a1.Origin(l).Y, (a1.Origin(l).Z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 2.5))))
                                            Else
                                               D3DXMatrixTranslation matPos, (a1.Origin(l).X - (Sin(D360 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 4.5)))), a1.Origin(l).Y, (a1.Origin(l).Z - (Cos(D360 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 4.5))))
                                            End If
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                                
                                        End If
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                        
                                    ElseIf a1.VerticalSkew Then
    
                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Cameras(Player.CameraIndex).Angle), -(Cameras(Player.CameraIndex).Pitch * 0.25), 0
                                            
                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Cameras(Player.CameraIndex).Pitch > 0, -Cameras(Player.CameraIndex).Pitch, Cameras(Player.CameraIndex).Pitch) * 0.25), 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
        
                                            D3DXMatrixTranslation matPos, a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (1 / (PI * 6))), a1.Origin(l).Y, a1.Origin(l).Z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (1 / (PI * 6)))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                        Else
                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle), -(Player.Camera.Pitch * 0.25), 0
                                            
                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Player.Camera.Pitch > 0, -Player.Camera.Pitch, Player.Camera.Pitch) * 0.25), 1
                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
        
                                            D3DXMatrixTranslation matPos, a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 6))), a1.Origin(l).Y, a1.Origin(l).Z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (1 / (PI * 6)))
                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                                
                                        End If
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                    Else
                                    
                                        D3DXMatrixRotationY matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle)
        
                                        D3DXMatrixScaling matScale, 1, 1, 1
                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                        D3DXMatrixTranslation matPos, a1.Origin(l).X, a1.Origin(l).Y, a1.Origin(l).Z
                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                            
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                    
                                    End If
                             
                                    If a1.BeaconLight > -1 Then
                                        DXLights(a1.BeaconLight).Position.X = a1.Origin(l).X - ((a1.Origin(l).X - Player.Element.Origin.X) / 80)
                                        DXLights(a1.BeaconLight).Position.Y = a1.Origin(l).Y - ((a1.Origin(l).Y - Player.Element.Origin.Y) / 80)
                                        DXLights(a1.BeaconLight).Position.Z = a1.Origin(l).Z - ((a1.Origin(l).Z - Player.Element.Origin.Z) / 80)

                                        DDevice.SetLight a1.BeaconLight - 1, DXLights(a1.BeaconLight)
                                        DDevice.LightEnable a1.BeaconLight - 1, 1
                                    End If
                                    
                                    If (a1.BeaconAnim = 0) Or (CDbl(Timer - a1.BeaconAnim) >= 0.05) Then
                                        a1.BeaconAnim = GetTimer
                                        
                                        a1.BeaconText = a1.BeaconText + 1
                                        If a1.BeaconText > a1.FileNames.Count Then
                                            a1.BeaconText = 1
                                        End If
                                        
                                    End If
                                    
                                    If a1.Translucent Then
                                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                        DDevice.SetMaterial LucentMaterial
        
                                    ElseIf a1.Alphablend Then
                                    
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
                                    'DDevice.SetTexture 0, a1.BeaconSkin(a1.BeaconText)
                                    a1.SetTexture 0, a1.BeaconText
                                    DDevice.SetMaterial GenericMaterial
                                    'DDevice.SetTexture 1, a1.BeaconSkin(a1.BeaconText)
                                    a1.SetTexture 1, a1.BeaconText
                        
                                    a1.Render
                                   'DDevice.SetStreamSource 0, a1.BeaconVBuf, Len(a1.BeaconPlaq(0))
                                   ' DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                                   
'                                    If a1.BeaconLight > -1 Then
'                                        DDevice.LightEnable a1.BeaconLight - 1, False
'                                    End If
                                    
                                ElseIf a1.RoundingCut > 0 Then
    
                                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
                                    chop = a1.RoundingCut
                                    start = 0
                                    
                                    Do Until start >= 360
                          
                                        D3DXMatrixIdentity matBeacon

                                        D3DXMatrixRotationY matBeacon, (start / (PI * 2)) + IIf(a1.HorizontalLock, 0, -Player.Camera.Angle)
        
                                        D3DXMatrixScaling matScale, 1, 1, 1
                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                            
                                        D3DXMatrixTranslation matPos, a1.Origin(l).X, a1.Origin(l).Y, a1.Origin(l).Z
                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
                                            
                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
                                        

'                                    D3DXMatrixIdentity matBeacon
'
'                                    If a1.VerticalLock Then
'
'                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
'                                            a = Cameras(Player.CameraIndex).Angle
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, -a, -Cameras(Player.CameraIndex).Pitch, 0
'
'                                            D3DXMatrixScaling matScale, 1, 1, 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, (a1.Origin(l).X - (Sin(D720 - a) * (a1.Dimension.height / (PI * 2.5)))), a1.Origin(l).Y, (a1.Origin(l).z - (Cos(D720 - a) * (a1.Dimension.height / (PI * 2.5))))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'                                        Else
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle), -Player.Camera.Pitch, 0
'
'                                            D3DXMatrixScaling matScale, 1, 1, 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, (a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (a1.Dimension.height / (PI * 2.5)))), a1.Origin(l).Y, (a1.Origin(l).z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (a1.Dimension.height / (PI * 2.5))))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        End If
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'                                    ElseIf a1.VerticalSkew Then
'
'                                        If (Perspective = CameraMode) And (Player.CameraIndex > 0) Then
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Cameras(Player.CameraIndex).Angle), -(Cameras(Player.CameraIndex).Pitch * 0.25), 0
'
'                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Cameras(Player.CameraIndex).Pitch > 0, -Cameras(Player.CameraIndex).Pitch, Cameras(Player.CameraIndex).Pitch) * 0.25), 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (a1.Dimension.height / (PI * 6))), a1.Origin(l).Y, a1.Origin(l).z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Cameras(Player.CameraIndex).Angle)) * (a1.Dimension.height / (PI * 6)))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'                                        Else
'                                            D3DXMatrixRotationYawPitchRoll matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle), -(Player.Camera.Pitch * 0.25), 0
'
'                                            D3DXMatrixScaling matScale, 1, 1 - (IIf(Player.Camera.Pitch > 0, -Player.Camera.Pitch, Player.Camera.Pitch) * 0.25), 1
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                            D3DXMatrixTranslation matPos, a1.Origin(l).X - (Sin(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (a1.Dimension.height / (PI * 6))), a1.Origin(l).Y, a1.Origin(l).z - (Cos(D720 - IIf(a1.HorizontalLock, 0, Player.Camera.Angle)) * (a1.Dimension.height / (PI * 6)))
'                                            D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        End If
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'                                    Else
'
'                                        D3DXMatrixRotationY matBeacon, IIf(a1.HorizontalLock, 0, -Player.Camera.Angle)
'
'                                        D3DXMatrixScaling matScale, 1, 1, 1
'                                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                                        D3DXMatrixTranslation matPos, a1.Origin(l).X, a1.Origin(l).Y, a1.Origin(l).z
'                                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
'
'                                        DDevice.SetTransform D3DTS_WORLD, matBeacon
'
'                                    End If
                                        
                                        
                                        
                                        

                                        
                                        If a1.Translucent Then
                                            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                                            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                                            DDevice.SetMaterial LucentMaterial
    
                                        ElseIf a1.Alphablend Then
    
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

                                                                    
                                        If a1.BeaconLight > -1 Then
                                            DXLights(a1.BeaconLight).Position.X = a1.Origin(l).X - ((a1.Origin(l).X - Player.Element.Origin.X) / 80)
                                            DXLights(a1.BeaconLight).Position.Y = a1.Origin(l).Y - ((a1.Origin(l).Y - Player.Element.Origin.Y) / 80)
                                            DXLights(a1.BeaconLight).Position.Z = a1.Origin(l).Z - ((a1.Origin(l).Z - Player.Element.Origin.Z) / 80)

                                            DDevice.SetLight a1.BeaconLight - 1, DXLights(a1.BeaconLight)
                                            DDevice.LightEnable a1.BeaconLight - 1, 1
                                        End If
                                        
                                        If (a1.BeaconAnim = 0) Or (CDbl(Timer - a1.BeaconAnim) >= 0.05) Then
                                            a1.BeaconAnim = GetTimer
                                            
                                            a1.BeaconText = a1.BeaconText + 1
                                            If a1.BeaconText > a1.FileNames.Count Then
                                                a1.BeaconText = 1
                                            End If
                                            
                                        End If
                                        
                                        
                                        DDevice.SetPixelShader PixelShaderDefault
                                        'DDevice.SetTexture 0, a1.BeaconSkin(a1.BeaconText)
                                        a1.SetTexture 0, a1.BeaconText
                                        DDevice.SetMaterial GenericMaterial
                                        'DDevice.SetTexture 1, a1.BeaconSkin(a1.BeaconText)
                                        a1.SetTexture 1, a1.BeaconText
                            
                                        a1.Render
                                        'DDevice.SetStreamSource 0, a1.BeaconVBuf, Len(a1.BeaconPlaq(0))
                                        'DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                                    
                                        start = (start + chop)
'                                        If a1.BeaconLight > -1 Then
'
'
'                                            DDevice.LightEnable a1.BeaconLight - 1, False
'                                        End If
                                    Loop
                                    
                                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                                End If
                            End If
    
                        End If
                        
                        l = l + 1
                        
                    Loop
                    
                End If
                
                If a1.BeaconLight > 0 Then
                    If Lights.Count > 0 Then
                        For l = 1 To Lights.Count
                            
                            DDevice.SetLight l - 1, DXLights(l)
                            If Lights(l).LightType = Lighting.Directed Or (l = a1.BeaconLight) Then
                                DDevice.LightEnable l - 1, False
                            Else
                                DDevice.LightEnable l - 1, 1
                            End If
                        Next
                    End If
                End If
            End If
            
            If ok And a1.Randomize And (a1.Origins.Count < a1.Allowance) Then
                Dim nb As New Beacon
                nb.Origin.X = X
                nb.Origin.Z = Z
                a1.Origins.Add nb
                Set nb = Nothing
'                a1.Origins.Count = a1.Origins.Count + 1
'                ReDim Preserve a1.Origin(1 To a1.Origins.Count) As D3DVECTOR
'                a1.Origin(a1.Origins.Count).X = X
'                a1.Origin(a1.Origins.Count).Z = Z
            End If
            
        Next

    End If
         
    
End Sub

Public Sub RenderSpaces()

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
    If Lights.Count > 0 Then
        For l = 1 To Lights.Count
            If Lights(l).SunLight Then
                DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(0, 164 + Lights(l).Diffuse.Red * 255, _
                    164 + Lights(l).Diffuse.Green * 255, 164 + Lights(l).Diffuse.Blue * 255)
'
'                If (FPSRate > 0) And Lights(l).SunLight And Spaces(1).SkyRotate > 0 Then
'
'                    Spaces(1).SkyRotated = Spaces(1).SkyRotated + (360 / (HoursInOneDay * Spaces(1).SkyRotate)) * _
'                        (PI / (HoursInOneDay * FPSRate)) * (FPSRate / HoursInOneDay)
'                End If
                
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




    RenderSpacesSetup
    
    Dim S As Space
    For Each S In Spaces
    
    
        If S.Translucent Then
            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
            
        ElseIf S.Alphablend Then

            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

        Else
            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                    
        End If
        
        S.RenderSpace

    Next
    
        
    RenderSpacesClose
    
    
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, 1

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetRenderState D3DRS_FOGCOLOR, DDevice.GetRenderState(D3DRS_AMBIENT)
    DDevice.SetRenderState D3DRS_AMBIENT, D3DColorRGBA(0, 0, 0, 0)
    
End Sub

Public Sub CreateLand(Optional ByVal NoDeserialize As Boolean = False)
 
    Set All = New NTNodes10.Collection
    Set Beacons = New NTNodes10.Collection
    Set Boards = New NTNodes10.Collection
    Set Cameras = New NTNodes10.Collection
    Set Elements = New NTNodes10.Collection
    Set Lights = New NTNodes10.Collection
    Set Portals = New NTNodes10.Collection
    Set Screens = New NTNodes10.Collection
    Set Sounds = New NTNodes10.Collection
    Set Tracks = New NTNodes10.Collection
    Set Spaces = New NTNodes10.Collection

    Set Player = New Player
    
    'Set Space = New Space

    
    Bindings.MouseInput = Trap
    Perspective = ThirdPerson
    CameraClip = True
        
    frmMain.Startup
    
    If ScriptRoot = "" Then
        If PathExists(AppPath & "Levels\" & CurrentLoadedLevel & ".vbx") Then
            ScriptRoot = AppPath
        ElseIf PathExists(CurDir & "\" & CurrentLoadedLevel & ".vbx") Then
            ScriptRoot = CurDir
        ElseIf PathExists(AppPath & CurrentLoadedLevel & ".vbx") Then
            ScriptRoot = Left(AppPath, Len(AppPath) - 1)
        End If
        If ScriptRoot = "" Then
            ScriptRoot = modFolders.SearchPath(CurrentLoadedLevel & ".vbx", True, CurDir, FirstOnly)
            If ScriptRoot <> "" Then ScriptRoot = GetFilePath(ScriptRoot)
        End If
    End If

    If PathExists(ScriptRoot & "Levels\" & CurrentLoadedLevel & ".vbx") Then
        ParseScript ScriptRoot & "Levels\" & CurrentLoadedLevel & ".vbx", , , NoDeserialize
    End If
    
    ComputeNormals
    
End Sub



Public Sub CleanupLand(Optional ByVal NoSerialize As Boolean = False)
    Dim uid As Long
    Dim ser As String
    Dim A As Long
        
    If Not NoSerialize Then
        On Error GoTo serialerror:
        
        ser = frmMain.Evaluate("Serialize")
        If ser <> "" Then WriteFile ScriptRoot & "Levels\" & CurrentLoadedLevel & ".xml", ser
serialerror:
        If Err.Number <> 0 Then
            MsgBox "Unable to save information due to an error." & vbCrLf & Err.source & " " & Err.Description, vbCritical, "Serialize Error"
            Err.Clear
            
        End If

    End If
    
    frmMain.Reset
    

    BackgroundColor = 0

    ShowCredits = False

'    If Not Players Is Nothing Then
'        Players.Clear
'        Set Players = Nothing
'    End If
    Set Player = Nothing
        
    
    If Not Spaces Is Nothing Then
        Spaces.Clear
        Set Spaces = Nothing
    End If
    
    Dim q As Integer
    Dim o As Integer
    

        
        
    If Not Elements Is Nothing Then
        Dim e As Element
        o = 1
        Do While o <= Elements.Count
            Elements(o).AttachedTo = ""
            Elements(o).ClearMotions
            o = o + 1
        Loop
        
        CleanUpAllCollection Elements
    End If
    

    If Not Portals Is Nothing Then
        If Portals.Count > 0 Then
            For o = 1 To Portals.Count
                Portals(o).ClearMotions
            Next
        End If
        
        
        CleanUpAllCollection Portals
    End If
    
    
    If Not Boards Is Nothing Then
        CleanUpAllCollection Boards
    End If

    
    If Not Beacons Is Nothing Then
        Beacons.Clear
        Set Beacons = Nothing
    End If

    If Not Screens Is Nothing Then
        CleanUpAllCollection Screens
    End If

    If Not Cameras Is Nothing Then
        CleanUpAllCollection Cameras
    End If
    
    If Not Tracks Is Nothing Then
        For o = 1 To Tracks.Count
            Tracks(o).StopTrack
            Set Tracks(o) = Nothing
        Next
        CleanUpAllCollection Tracks
    End If
    
    If Not Sounds Is Nothing Then
        For o = 1 To Sounds.Count
            StopWave Sounds(o).Index
            Set Waves(o) = Nothing
        Next
        CleanUpAllCollection Sounds
        Erase Waves
    End If

    If Not All Is Nothing Then
        All.Clear
        Set All = Nothing
    End If

    If Not Lights Is Nothing Then
        For o = 1 To Lights.Count
            DDevice.LightEnable o, False
        Next
        CleanUpAllCollection Lights
    End If

    Erase Waves
    Erase DXLights
        
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
    
    CurrentLoadedLevel = ""
    
End Sub

Private Sub CleanUpAllCollection(ByRef SubCollection As NTNodes10.Collection)
    Do While SubCollection.Count > 0
        If Not All Is Nothing Then
            If All.Exists(SubCollection(1).Key) Then
                All.Remove SubCollection(1).Key
            End If
        End If
        SubCollection.Remove 1
    Loop
    Set SubCollection = Nothing
End Sub


Public Sub BeginMirrors()

'    Dim e As Board
'    Dim i As Long
'    Dim l As Single
'
'    Dim dm As D3DDISPLAYMODE
'    Dim pal As PALETTEENTRY
'    Dim rct As RECT
''
''    If Not Mirrors Is Nothing Then Mirrors.Clear
'    If Boards.Count > 0 Then
'        For i = 1 To Boards.Count
'            Set e = Boards(i)
'
'            If e.Visible And e.Mirror Then
'
'                l = Distance(Player.Origin.X, Player.Origin.Y, Player.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
'                If l <= FAR Then
'
'                    DViewPort.Width = 128
'                    DViewPort.Height = 128
'
'                    DSurface.BeginScene DefaultRenderTarget, DViewPort
'
''                    BeginWorld UserControl, e.Transposing
''
''                    RenderPlanets UserControl, e.Transposing
''                    RenderObject UserControl, e.Transposing
'
'                    DSurface.EndScene
'
'                    DDevice.GetDisplayMode dm
'
'                    rct.Top = 0
'                    rct.Left = 0
'
'                    rct.Right = DViewPort.Width
'                    rct.Bottom = DViewPort.Height
'
'                    D3DX.SaveSurfaceToFile GetTemporaryFolder & "\" & e.Key & ".bmp", D3DXIFF_BMP, DefaultRenderTarget, pal, rct
'                    Mirrors.Add D3DX.CreateTextureFromFileEx(DDevice, GetTemporaryFolder & "\" & e.Key & ".bmp", _
'                        DViewPort.Width, DViewPort.Height, D3DX_FILTER_NONE, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
'                        D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0), e.Key
'                    Kill GetTemporaryFolder & "\" & e.Key & ".bmp"
'
'                End If
'
'            End If
'            Set e = Nothing
'        Next
'''    End If
End Sub


Public Sub RenderMirror()

'    Dim e As Billboard
'    Dim i As Long
'    Dim l As Single
'
'    If Boards.Count > 0 Then
'        For i = 1 To Boards.Count
'            Set e = Boards(i)
'
'            If e.Visible And ((e.Form And ThreeDimensions) = ThreeDimensions) Then
'
'                l = Distance(Player.Origin.X, Player.Origin.Y, Player.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
'                If l <= FAR Then
'
''                    If Mirrors.Exists(Billboards.Key(i)) Then
''                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
''                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
''                        DDevice.SetMaterial GenericMaterial
''                        DDevice.SetTexture 0, Mirrors.Item(Boards.Key(i))
''                        DDevice.SetTexture 1, Nothing
''
''                        DDevice.SetStreamSource 0, Faces(e.FaceIndex).VBuffer, Len(Faces(e.FaceIndex).Verticies(0))
''                        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
''
''                    End If
'
'                End If
'
'            End If
'            Set e = Nothing
'        Next
'    End If
End Sub



