Attribute VB_Name = "modLawn"

#Const modLawn = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Public Type MyMotion
    Direction As D3DVECTOR
    BurstRate As Single
    OffsetLoc As D3DVECTOR
End Type

Public Type MyActivity
    Identity As String
    Direction As D3DVECTOR
    OffsetLoc As D3DVECTOR
    BurstRate As Single
    Friction As Single
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
    Multiplayer As Boolean
    
    AnimateMSecs As Single
    AnimateTimer As Double
    AnimatePoint As Long
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

    IsIdol As Boolean
    MeshIndex As Integer
    
    Origin As D3DVECTOR
    Scaled As D3DVECTOR
    Rotate As D3DVECTOR
    Matrix As D3DMATRIX
    XModel As D3DMATRIX
    
    Motion As MyMotion
    
    Activities() As MyActivity
    ActivityCount As Long
End Type

Public Type MyBeacon
    Origins() As D3DVECTOR
    OriginCount As Long
    
    BeaconSkin() As Direct3DTexture8
    BeaconSkinCount As Long

    BeaconPlaq(0 To 5) As MyVertex
    BeaconVBuf As Direct3DVertexBuffer8

    Dimension As ImgDimType
    PercentXY As ImgDimType
    Translucent As Boolean
    
    SinglePlayer As Boolean
    
    Disposable As Boolean
    Consumable As Boolean
    Randomize As Boolean
    Allowance As Long
    
    BeaconAnim As Double
    BeaconText As Long
    BeaconLight As Long
    
    Identity As String
End Type

Public Type MyLight
    
    Origin As D3DVECTOR
    
    LightBlink As Single
    LightTimer As Single
    LightIsOn As Boolean
    
    LightIndex As Long
End Type

Public Type MyPlayer
    Object As MyObject
    Rotation As Single
    CameraAngle As Single
    CameraPitch As Single
    CameraZoom As Single
    MoveSpeed As Single
    LeftFlap As Boolean
    RightFlap As Boolean
    FlapLock As Boolean
    AutoMove As Boolean
    Gravity As Single
    Stalled As Boolean
    Texture As Byte
    Model As Byte
    name As String
    Trails As Boolean
    Spots() As D3DVECTOR
    Flag As Boolean
End Type

Public SundialAim As Single
Public Sunrotated As Double

Public Lights() As D3DLIGHT8
Public LightCount As Long

Public LightDatas() As MyLight
Public LightDataCount As Long

Public Meshes() As MyMesh
Public MeshCount As Long

Public Objects() As MyObject
Public ObjectCount As Long

Private BillBoards() As MyBoard
Private BillBoardCount As Long

Public Beacons() As MyBeacon
Public BeaconCount As Long

Private PlaneSkin() As Direct3DTexture8
Private PlaneHole() As D3DVECTOR

Private PlanePlaq() As MyVertex
Private PlaneVBuf As Direct3DVertexBuffer8

Public matWorld As D3DMATRIX

Public Function NearSymbol(ByRef bot As MyPlayer) As Boolean
    Dim o As Long
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            If Objects(o).MeshIndex > 0 Then
                If Objects(o).IsIdol And IsNumeric(Left(Meshes(Objects(o).MeshIndex).FileName, 1)) Then
                    If Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Objects(o).Origin.X, Objects(o).Origin.Y, Objects(o).Origin.z) <= ZoneDistance Then
                        NearSymbol = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
End Function

Public Sub RenderLights()
    
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
    
    Dim l As Long

    For l = 1 To LightDataCount
        
        If Lights(LightDatas(l).LightIndex).Type = D3DLIGHT_POINT Or Lights(LightDatas(l).LightIndex).Type = D3DLIGHT_SPOT Then
            If Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, LightDatas(l).Origin.X, LightDatas(l).Origin.Y, LightDatas(l).Origin.z) <= (FadeDistance * 2) Then
            
                If ((LightDatas(l).LightTimer = 0) And (LightDatas(l).LightBlink > 0)) Or ((Timer - LightDatas(l).LightTimer) >= LightDatas(l).LightBlink And (LightDatas(l).LightBlink > 0)) Then
                    LightDatas(l).LightTimer = Timer
                    LightDatas(l).LightIsOn = Not LightDatas(l).LightIsOn
                    DDevice.LightEnable (l - 1), LightDatas(l).LightIsOn
                    
                ElseIf (LightDatas(l).LightBlink = 0) Then
                    DDevice.LightEnable (l - 1), 1
                End If
            Else
                DDevice.LightEnable (l - 1), False
            End If
        Else
            DDevice.LightEnable (l - 1), 1
        End If
    Next
    
End Sub

Public Sub RenderGalaxy()

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    SetAmbientRGB 255, 255, 255, 0
    
    DDevice.SetMaterial GenericMaterial
   
    Dim matProj As D3DMATRIX
    Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
    DDevice.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
    
    DDevice.SetTransform D3DTS_VIEW, matView
    DDevice.SetTransform D3DTS_WORLD, matWorld
    DDevice.SetRenderState D3DRS_ZENABLE, 0

    If (Sunrotated = 0) Or (((Timer - Sunrotated) * 0.6) >= 360) Then Sunrotated = Timer

    D3DXMatrixPerspectiveFovLH matProj, PI / 3.5, AspectRatio, 1, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
        
    D3DXMatrixRotationY matWorld, ((Timer - Sunrotated) * 0.6) * (PI / 180)
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    DDevice.SetTexture 0, PlaneSkin(4)
    DDevice.SetStreamSource 0, PlaneVBuf, Len(PlanePlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 102, 2
    
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, AspectRatio, 30, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
  
    DDevice.SetTransform D3DTS_VIEW, matViewSave
    DDevice.SetTransform D3DTS_WORLD, matWorld

    SetAmbientRGB 0, 0, 0, 1
    
End Sub
Public Sub RenderPlane()
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    
    If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
    If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    End If

    Dim dist As Single
    Dim rgbval As Single
    
    dist = Distance(0, 0, 0, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z)

    If (dist <= (FadeDistance + WithInCityLimits)) Then

        If dist < ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits Then
    
            SetAmbientRGB Ambient_HI, Ambient_HI, Ambient_HI, 1
    
        ElseIf dist >= ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits And dist < (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then

            rgbval = (((((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits) - dist)
            rgbval = Abs(-100 + ((((((FadeDistance - WithInCityLimits) / 10) * 7) - WithInCityLimits) / (((rgbval - Ambient_LO) / Ambient_HI) + Ambient_LO)) - 100))
            rgbval = Round(rgbval, 0)
            rgbval = ((Ambient_HI - Ambient_LO) * (rgbval / 100)) + Ambient_LO
            
            SetAmbientRGB rgbval, rgbval, rgbval, 1
            
        ElseIf dist >= (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then
    
            SetAmbientRGB Ambient_LO, Ambient_LO, Ambient_LO, 1
    
        End If
                
    End If
   
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
       
    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 0, PlaneSkin(1)
    DDevice.SetStreamSource 0, PlaneVBuf, Len(PlanePlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 16

    If Player.Object.Origin.Y < 0 Then
        DDevice.SetTexture 0, PlaneSkin(2)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 48, 16

        DDevice.SetTexture 0, PlaneSkin(3)
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 96, 2
    End If

End Sub

Public Sub RenderLawn()

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    Dim dist As Single
    Dim rgbval As Single
    
    Dim o As Long
    Dim i As Long
    
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            
            dist = Distance(Objects(o).Origin.X, Objects(o).Origin.Y, Objects(o).Origin.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z)
    
            If (dist <= (FadeDistance + WithInCityLimits)) Then
                    
                If Objects(o).IsIdol Then
                    
'                    Static SymbolTimer As Single
'
'                    If IsNumeric(Left(Meshes(Objects(o).MeshIndex).FileName, 1)) And _
'                        (Not (Left(Meshes(Objects(o).MeshIndex).FileName, 1) = PortalHold)) And _
'                        (InStr(ScoreIdol, Left(Meshes(Objects(o).MeshIndex).FileName, 1)) = 0) Then
'
'                        If dist <= ZoneDistance Then
'
'                            If (SymbolTimer = 0) Then SymbolTimer = Timer
'                            If (SymbolTimer = 0) Or (Timer - SymbolTimer > 7) Then
'                                SymbolTimer = -1
'
'                                PortalHold = Left(Meshes(Objects(o).MeshIndex).FileName, 1)
'
'                            End If
'
'                        End If
'                    Else
'                        SymbolTimer = 0
'                    End If
'
'                    If SymbolTimer >= 0 Then
                    
                        If dist < ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits Then
                    
                            SetAmbientRGB Ambient_HI, Ambient_HI, Ambient_HI, 1
                    
                        ElseIf dist >= ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits And dist < (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then
    
                            rgbval = (((((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits) - dist)
                            rgbval = Abs(-100 + ((((((FadeDistance - WithInCityLimits) / 10) * 7) - WithInCityLimits) / (((rgbval - Ambient_LO) / Ambient_HI) + Ambient_LO)) - 100))
                            rgbval = Round(rgbval, 0)
                            rgbval = ((Ambient_HI - Ambient_LO) * (rgbval / 100)) + Ambient_LO
                            
                            SetAmbientRGB rgbval, rgbval, rgbval, 1
                            
                        ElseIf dist >= (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then
                    
                            SetAmbientRGB Ambient_LO, Ambient_LO, Ambient_LO, 1
                    
                        End If
                    'End If
                End If
            
            End If
            
            If (dist <= (FadeDistance * 2)) Then

                If Objects(o).MeshIndex > 0 Then
                
                    Dim faceNum As Long
                    
                    DDevice.SetTransform D3DTS_WORLD, Objects(o).Matrix
                    For i = 0 To Meshes(Objects(o).MeshIndex).MaterialCount - 1
                        DDevice.SetMaterial Meshes(Objects(o).MeshIndex).Materials(i)

                        If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
                            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                        End If

                        If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
                            DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
                            DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
                        End If

                        DDevice.SetTexture 0, Meshes(Objects(o).MeshIndex).Textures(i)
                        Meshes(Objects(o).MeshIndex).Mesh.DrawSubset i
                    Next

                    If frmMain.Multiplayer And (Not Objects(o).IsIdol) Then
                        DDevice.SetTransform D3DTS_WORLD, Objects(o).XModel
                        For i = 0 To Meshes(Objects(o).MeshIndex).MaterialCount - 1
                            DDevice.SetMaterial Meshes(Objects(o).MeshIndex).Materials(i)

                            If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
                                DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                                DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
                            End If

                            If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
                                DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
                                DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
                            End If

                            DDevice.SetTexture 0, Meshes(Objects(o).MeshIndex).Textures(i)
                            Meshes(Objects(o).MeshIndex).Mesh.DrawSubset i
                        Next
                        
                    End If
                    
                End If
                
            End If
  
        Next
    End If
    
    PointSundial
    
End Sub

Public Sub RenderBeacons()

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1
         
    If Not (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
    End If
    If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    End If
                     
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                                
    SetAmbientRGB 255, 255, 255, 0
    
    DDevice.SetMaterial GenericMaterial
        
    Dim o As Long

    Dim matScale As D3DMATRIX
    Dim matPos As D3DMATRIX

    D3DXMatrixIdentity matWorld
                         
    Dim l As Long
    Dim i As Long
    Dim X As Single
    Dim z As Single
    Dim ok As Boolean

    Dim d As Single
        
    Dim matBeacon As D3DMATRIX
        
    If (BeaconCount > 0) Then
        For o = 1 To BeaconCount

            If (Not Beacons(o).SinglePlayer) Or (Beacons(o).SinglePlayer And (Not frmMain.Multiplayer)) Then
            
                If (Beacons(o).Randomize And (Beacons(o).OriginCount < Beacons(o).Allowance)) And ((frmMain.Recording And (Not frmMain.IsPlayback)) Or (Not frmMain.Recording)) Then
                    X = IIf((Rnd < 0.5), -RandomPositive(WithInCityLimits, BlackBoundary), RandomPositive(WithInCityLimits, BlackBoundary))
                    z = IIf((Rnd < 0.5), -RandomPositive(WithInCityLimits, BlackBoundary), RandomPositive(WithInCityLimits, BlackBoundary))
                    ok = True
                Else
                    ok = False
                End If
                        
                If (Beacons(o).OriginCount > 0) Then
                    l = 1
    
                    Do While l <= Beacons(o).OriginCount
    
                        d = Distance(Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z)
                        If ok Then ok = ok And (Distance(Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z, X, 0, z) > WithInCityLimits)
                        
                        If d <= FadeDistance Then
                            If (frmMain.Recording And (Not frmMain.IsPlayback)) Then
                                If Beacons(o).Disposable Or Beacons(o).Consumable Then
                                    Dim addItem As String
                                    addItem = "B" & Beacons(o).Identity & "," & Trim(CStr(Beacons(o).Origins(l).X)) & "," & Trim(CStr(Beacons(o).Origins(l).z)) & "|"
                                    If InStr(ViewData, addItem) = 0 Then
                                        ViewData = ViewData & addItem
                                        Put #FilmFileNum, , addItem
                                    End If
                                End If
                            End If
                            
                            If (Beacons(o).Consumable Or Beacons(o).Disposable) And (d <= 100) Then
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
                                
                                If Beacons(o).Consumable Or Beacons(o).Disposable Then
                                    
                                    If Beacons(o).Consumable Then
                                        ScoreNth = IIf(ScoreNth = -1, -ScoreNth, ScoreNth + 1)
                                    Else
                                        ScoreNth = -1
                                    End If
                                    
                                    Select Case Beacons(o).Identity
                                        Case "n0"
                                            If (ScoreNth >= 2000) And (InStr(Bodykits, "2") = 0) Then
                                                Bodykits = Bodykits & "2"
                                                ScoreNth = -1
                                                FadeMessage "You collected 2k nickels!  You receive the first garage body kit."
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            ElseIf (ScoreNth >= 5000) And (InStr(Bodykits, "1") = 0) Then
                                                Bodykits = Bodykits & "1"
                                                ScoreNth = -1
                                                FadeMessage "You collected 5k nickels!  You receive the second garage body kit."
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            ElseIf (ScoreNth >= 10000) And (InStr(Bodykits, "6") = 0) Then
                                                Bodykits = Bodykits & "6"
                                                FadeMessage "You collected 10k nickels!  You receive the elite pin striped N body kit."
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            End If
                                            
                                        Case "n1"
                                            If (InStr(Bodykits, "3") = 0) Then
                                                Bodykits = Bodykits & "3"
                                                FadeMessage "You found a custom N!  You receive a new special unique body kit!"
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            End If
                                        Case "n2"
                                            If (InStr(Bodykits, "4") = 0) Then
                                                Bodykits = Bodykits & "4"
                                                FadeMessage "You found a custom N!  You receive a new special unique body kit!"
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            End If
                                        Case "n3"
                                            If (InStr(Bodykits, "5") = 0) Then
                                                Bodykits = Bodykits & "5"
                                                FadeMessage "You found a custom N!  You receive a new special unique body kit!"
                                                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Bodykits='" & Bodykits & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                                            End If
                                        Case "n4"
                                            FadeMessage "You lost all of your nickels for picking up a trick N!"
                                    End Select
                                    
                                    If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreNth = " & ScoreNth & " WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"

                                End If
                                
                                PlayWave SOUND_NICKEL
    
                            ElseIf l <= Beacons(o).OriginCount Then
    
                                D3DXMatrixIdentity matBeacon
                                
                                D3DXMatrixRotationY matBeacon, -Player.CameraAngle
                                
                                D3DXMatrixScaling matScale, 1, 1, 1
                                D3DXMatrixMultiply matBeacon, matBeacon, matScale
                                
                                D3DXMatrixTranslation matPos, Beacons(o).Origins(l).X, Beacons(o).Origins(l).Y, Beacons(o).Origins(l).z
                                D3DXMatrixMultiply matBeacon, matBeacon, matPos
                    
                                DDevice.SetTransform D3DTS_WORLD, matBeacon
                         
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
    
                                DDevice.SetTexture 0, Beacons(o).BeaconSkin(Beacons(o).BeaconText)
                    
                                DDevice.SetStreamSource 0, Beacons(o).BeaconVBuf, Len(Beacons(o).BeaconPlaq(0))
                                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    
                            End If
    
                        End If
                        
                        l = l + 1
                        
                    Loop
                    
                End If
                
                If (frmMain.Recording And (Not frmMain.IsPlayback)) Or (Not frmMain.Recording) Then
                
                    ok = ok And (Beacons(o).OriginCount < Beacons(o).Allowance)
                    If ok And (ObjectCount > 0) And Beacons(o).Randomize Then
                        For i = 1 To ObjectCount
                            If ok Then ok = ok And (Distance(Objects(i).Origin.X, Objects(i).Origin.Y, Objects(i).Origin.z, X, 0, z) > WithInCityLimits)
                        Next
                    End If
        
                    If (ok And Beacons(o).Randomize And (Beacons(o).OriginCount < Beacons(o).Allowance)) Then
                        Beacons(o).OriginCount = Beacons(o).OriginCount + 1
                        ReDim Preserve Beacons(o).Origins(1 To Beacons(o).OriginCount) As D3DVECTOR
                        Beacons(o).Origins(Beacons(o).OriginCount).X = X
                        Beacons(o).Origins(Beacons(o).OriginCount).z = z
                        If Beacons(o).OriginCount > 1 Then NthTime = Beacons(o).OriginCount
                    End If
                
                End If
                
            End If
            
        Next

    End If
End Sub

Public Sub AddBeacon(ByVal id As String, ByVal X As Single, ByVal z As Single)
    Dim o As Long
    If (BeaconCount > 0) Then
        For o = 1 To BeaconCount
            If id = Beacons(o).Identity Then
                Beacons(o).OriginCount = Beacons(o).OriginCount + 1
                ReDim Preserve Beacons(o).Origins(1 To Beacons(o).OriginCount) As D3DVECTOR
                Beacons(o).Origins(Beacons(o).OriginCount).X = X
                Beacons(o).Origins(Beacons(o).OriginCount).z = z
                If Beacons(o).OriginCount > 1 Then NthTime = Beacons(o).OriginCount
            End If
        Next
    End If
End Sub

Public Sub RenderBoards()
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
    If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    End If
    
    DDevice.SetMaterial GenericMaterial
    
    Dim o As Long
    Dim dist As Single
    Dim rgbval As Single
    
    If BillBoardCount > 0 Then
        For o = 1 To BillBoardCount
        
            dist = Distance(BillBoards(o).Center.X, BillBoards(o).Center.Y, BillBoards(o).Center.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z)
    
            If (dist <= (FadeDistance + WithInCityLimits)) Then

                If dist < ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits Then
            
                    SetAmbientRGB Ambient_HI, Ambient_HI, Ambient_HI, 1
            
                ElseIf dist >= ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits And dist < (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then

                    rgbval = (((((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits) - dist)
                    rgbval = Abs(-100 + ((((((FadeDistance - WithInCityLimits) / 10) * 7) - WithInCityLimits) / (((rgbval - Ambient_LO) / Ambient_HI) + Ambient_LO)) - 100))
                    rgbval = Round(rgbval, 0)
                    rgbval = ((Ambient_HI - Ambient_LO) * (rgbval / 100)) + Ambient_LO
                    
                    SetAmbientRGB rgbval, rgbval, rgbval, 1
                    
                ElseIf dist >= (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then
            
                    SetAmbientRGB Ambient_LO, Ambient_LO, Ambient_LO, 1
            
                End If
                        
            End If
            
            If Distance(BillBoards(o).Center.X, BillBoards(o).Center.Y, BillBoards(o).Center.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z) <= FadeDistance Then
                If (Not BillBoards(o).Multiplayer) Or (BillBoards(o).Multiplayer And frmMain.Multiplayer) Then
                    If Not BillBoards(o).Translucent Then
                    
                        If (BillBoards(o).AnimateMSecs > 0) Then
                            If CDbl(Timer - BillBoards(o).AnimateTimer) >= BillBoards(o).AnimateMSecs Then
                                BillBoards(o).AnimateTimer = GetTimer
                                
                                BillBoards(o).AnimatePoint = BillBoards(o).AnimatePoint + 1
                                If BillBoards(o).AnimatePoint > BillBoards(o).SkinCount Then
                                    BillBoards(o).AnimatePoint = 1
                                End If
                                
                            End If
                       
                            DDevice.SetTexture 0, BillBoards(o).Skin(BillBoards(o).AnimatePoint)
                        Else

                            DDevice.SetTexture 0, BillBoards(o).Skin(1)
                        End If
                        DDevice.SetStreamSource 0, BillBoards(o).VBuf, Len(BillBoards(o).Plaq(0))
                        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                    End If
                End If
            End If
        Next
    End If

End Sub

Public Sub RenderGlass()
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR

    If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
    If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    End If
    
    DDevice.SetMaterial GenericMaterial
                             
    Dim p As D3DVECTOR
    Dim o As Long
    
    If BillBoardCount > 0 Then
        For o = 1 To BillBoardCount
            If Distance(BillBoards(o).Center.X, BillBoards(o).Center.Y, BillBoards(o).Center.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z) <= FadeDistance Then
                If (Not BillBoards(o).Multiplayer) Or (BillBoards(o).Multiplayer And frmMain.Multiplayer) Then
                    If BillBoards(o).Translucent Then

                        If (BillBoards(o).AnimateMSecs > 0) Then
                            If CDbl(Timer - BillBoards(o).AnimateTimer) >= BillBoards(o).AnimateMSecs Then
                                BillBoards(o).AnimateTimer = GetTimer
                                
                                BillBoards(o).AnimatePoint = BillBoards(o).AnimatePoint + 1
                                If BillBoards(o).AnimatePoint > BillBoards(o).SkinCount Then
                                    BillBoards(o).AnimatePoint = 1
                                End If
                                
                            End If

                            DDevice.SetTexture 0, BillBoards(o).Skin(BillBoards(o).AnimatePoint)
                        Else
                            DDevice.SetTexture 0, BillBoards(o).Skin(1)
                        End If
                        DDevice.SetStreamSource 0, BillBoards(o).VBuf, Len(BillBoards(o).Plaq(0))
                        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

                    End If
                End If
            End If
        Next
    End If
                                
End Sub

Public Sub PointSundial()
    Dim r As Single
    Dim cnt As Long
    
    If (Hour(Now) > 0) And (Hour(Now) <= 6) Then
        r = 0
    ElseIf (Hour(Now) > 6) And (Hour(Now) <= 12) Then
        r = 90
    ElseIf (Hour(Now) > 12) And (Hour(Now) <= 18) Then
        r = 180
    ElseIf (Hour(Now) > 18) And (Hour(Now) <= 24) Then
        r = 270
    End If
    
    If Not (SundialAim = r) Then
        SundialAim = r
    
        If ObjectCount > 0 Then
            
            For cnt = 1 To ObjectCount
                If Objects(cnt).MeshIndex > 0 Then
                    If Meshes(Objects(cnt).MeshIndex).FileName = "sundial.x" Then
        
                        D3DXMatrixIdentity matWorld
                        D3DXMatrixIdentity Objects(cnt).Matrix
        
                        D3DXMatrixScaling matWorld, Objects(cnt).Scaled.X, Objects(cnt).Scaled.Y, Objects(cnt).Scaled.z
                        D3DXMatrixMultiply Objects(cnt).Matrix, Objects(cnt).Matrix, matWorld
        
                        D3DXMatrixRotationY matWorld, r * (PI / 180)
                        D3DXMatrixMultiply Objects(cnt).Matrix, Objects(cnt).Matrix, matWorld
        
                        D3DXMatrixTranslation matWorld, Objects(cnt).Origin.X, Objects(cnt).Origin.Y, Objects(cnt).Origin.z
                        D3DXMatrixMultiply Objects(cnt).Matrix, Objects(cnt).Matrix, matWorld
        
                    End If
                End If
            Next
        End If
    
    End If
End Sub

Public Sub CreateLawn()
    
    ReDim PlaneSkin(1 To 4) As Direct3DTexture8
    ReDim PlaneHole(1 To 4) As D3DVECTOR

    ReDim PlanePlaq(0 To 107) As MyVertex
    
    ParseLawn AppPath & "Base\Lawn\lawn.px"

End Sub

Private Sub ParseLawn(ByVal inFile As String)
    On Error GoTo parseerror
    
    Dim inText As String
    inText = Replace(ReadFile(inFile), vbTab, "")
    
    Dim r As Single
    Dim o As Long
    Dim i As Long
    Dim vn As D3DVECTOR
    Dim LineNumber As Long
    Dim inArg() As String
    Dim inItem As String
    Dim inName As String
    Dim inData As String
    
    Do Until inText = ""
    
        inItem = Replace(RemoveNextArg(inText, vbCrLf), vbTab, "")
        LineNumber = LineNumber + 1
        If (Not (inItem = "")) And (Not (Left(inItem, 1) = ";")) Then
            inData = RemoveQuotedArg(inText, "{", "}")
            Select Case inItem
                Case "parse"
                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "filename"
                                     ParseLawn AppPath & "Base\Lawn\" & inArg(1)
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If
                            End Select
                        End If
                    Loop
                    
                Case "light"
                    LightCount = LightCount + 1
                    ReDim Preserve Lights(1 To LightCount) As D3DLIGHT8
                    
                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        LineNumber = LineNumber + 1
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "position"
                                    LightDataCount = LightDataCount + 1
                                    ReDim Preserve LightDatas(1 To LightDataCount) As MyLight
                                    LightDatas(LightDataCount).Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                    Lights(LightCount).Position = LightDatas(LightDataCount).Origin
                                    LightDatas(LightDataCount).LightIndex = LightCount
                                    DDevice.SetLight (LightDataCount - 1), Lights(LightCount)
                                    DDevice.LightEnable (LightDataCount - 1), 1
                                Case "direction"
                                    Lights(LightCount).Direction = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "diffuse"
                                    Lights(LightCount).diffuse.a = CSng(inArg(1))
                                    Lights(LightCount).diffuse.r = CSng(inArg(2))
                                    Lights(LightCount).diffuse.G = CSng(inArg(3))
                                    Lights(LightCount).diffuse.b = CSng(inArg(4))
                                Case "ambient"
                                    Lights(LightCount).Ambient.a = CSng(inArg(1))
                                    Lights(LightCount).Ambient.r = CSng(inArg(2))
                                    Lights(LightCount).Ambient.G = CSng(inArg(3))
                                    Lights(LightCount).Ambient.b = CSng(inArg(4))
                                Case "specular"
                                    Lights(LightCount).specular.a = CSng(inArg(1))
                                    Lights(LightCount).specular.r = CSng(inArg(2))
                                    Lights(LightCount).specular.G = CSng(inArg(3))
                                    Lights(LightCount).specular.b = CSng(inArg(4))
                                Case "attenuation"
                                    Lights(LightCount).Attenuation0 = CSng(inArg(1))
                                    Lights(LightCount).Attenuation1 = CSng(inArg(2))
                                    Lights(LightCount).Attenuation2 = CSng(inArg(3))
                                Case "phi"
                                    Lights(LightCount).Phi = CSng(inArg(1))
                                Case "theta"
                                    Lights(LightCount).Theta = CSng(inArg(1))
                                Case "falloff"
                                    Lights(LightCount).Falloff = CSng(inArg(1))
                                Case "range"
                                    Lights(LightCount).Range = inArg(1)
                                Case "type"
                                    Lights(LightCount).Type = inArg(1)
                                Case "blink"
                                    LightDatas(LightDataCount).LightBlink = inArg(1)
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If
                            End Select
                        End If
                    Loop
                    
                Case "object", "idol"
                    Dim NewObj As MyObject
                    NewObj.MeshIndex = 0
                    
                    NewObj.IsIdol = (inItem = "idol")

                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        LineNumber = LineNumber + 1
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "filename"
                                    
                                    If MeshCount = 0 Then
                                        MeshCount = MeshCount + 1
                                        ReDim Meshes(1 To MeshCount) As MyMesh
                                        NewObj.MeshIndex = MeshCount
                                        Meshes(NewObj.MeshIndex).FileName = LCase(inArg(1))
                                        If PathExists(AppPath & "Base\Lawn\" & Meshes(NewObj.MeshIndex).FileName, True) Then
                                            CreateMesh AppPath & "Base\Lawn\" & Meshes(NewObj.MeshIndex).FileName, Meshes(NewObj.MeshIndex)
                                        Else
                                            ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
                                            ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
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
                                            If PathExists(AppPath & "Base\Lawn\" & Meshes(NewObj.MeshIndex).FileName, True) Then
                                                CreateMesh AppPath & "Base\Lawn\" & Meshes(NewObj.MeshIndex).FileName, Meshes(NewObj.MeshIndex)
                                            Else
                                                ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
                                                ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
                                            End If
                                    
                                        End If
                                        
                                    End If

                                Case "origin"
                                    NewObj.Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "scale"
                                    NewObj.Scaled = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "rotate"
                                    NewObj.Rotate = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If
                            End Select
                        End If
                    Loop

                    D3DXMatrixIdentity matWorld
                    D3DXMatrixIdentity NewObj.Matrix
                    
                    D3DXMatrixScaling matWorld, NewObj.Scaled.X, NewObj.Scaled.Y, NewObj.Scaled.z
                    D3DXMatrixMultiply NewObj.Matrix, NewObj.Matrix, matWorld
                    
                    D3DXMatrixRotationX matWorld, NewObj.Rotate.X * (PI / 180)
                    D3DXMatrixMultiply NewObj.Matrix, NewObj.Matrix, matWorld
                    D3DXMatrixRotationY matWorld, NewObj.Rotate.Y * (PI / 180)
                    D3DXMatrixMultiply NewObj.Matrix, NewObj.Matrix, matWorld
                    D3DXMatrixRotationZ matWorld, NewObj.Rotate.z * (PI / 180)
                    D3DXMatrixMultiply NewObj.Matrix, NewObj.Matrix, matWorld
                    
                    D3DXMatrixTranslation matWorld, NewObj.Origin.X, NewObj.Origin.Y, NewObj.Origin.z
                    D3DXMatrixMultiply NewObj.Matrix, NewObj.Matrix, matWorld
                
                    If Not NewObj.IsIdol Then
                        D3DXMatrixIdentity matWorld
                        D3DXMatrixScaling matWorld, ScaleModelSize, ScaleModelSize, ScaleModelSize
                        D3DXMatrixMultiply NewObj.XModel, NewObj.Matrix, matWorld
                        
                        D3DXMatrixTranslation matWorld, ScaleModelLocX, ScaleModelLocY, ScaleModelLocZ
                        D3DXMatrixMultiply NewObj.XModel, NewObj.XModel, matWorld
                    End If
                    
                    ObjectCount = ObjectCount + 1
                    ReDim Preserve Objects(1 To ObjectCount) As MyObject
                    Objects(ObjectCount) = NewObj
                    
                Case "beacon"
                    BeaconCount = BeaconCount + 1
                    ReDim Preserve Beacons(1 To BeaconCount) As MyBeacon
                    Beacons(BeaconCount).BeaconLight = -1
                    Beacons(BeaconCount).Dimension.width = 1
                    Beacons(BeaconCount).Dimension.height = 1
                    Beacons(BeaconCount).PercentXY.width = 100
                    Beacons(BeaconCount).PercentXY.height = 100
                    Beacons(BeaconCount).Allowance = 1
                    
                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        LineNumber = LineNumber + 1
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "identity"
                                    Beacons(BeaconCount).Identity = inArg(1)
                                Case "disposable"
                                    Beacons(BeaconCount).Disposable = True
                                    Beacons(BeaconCount).Consumable = False
                                Case "consumable"
                                    Beacons(BeaconCount).Consumable = True
                                    Beacons(BeaconCount).Disposable = False
                                Case "randomize"
                                    Beacons(BeaconCount).Randomize = True
                                Case "singleplayer"
                                    Beacons(BeaconCount).SinglePlayer = True
                                Case "allowance"
                                    Beacons(BeaconCount).Allowance = CLng(inArg(1))
                                Case "origin"
                                    Beacons(BeaconCount).OriginCount = Beacons(BeaconCount).OriginCount + 1
                                    ReDim Preserve Beacons(BeaconCount).Origins(1 To Beacons(BeaconCount).OriginCount) As D3DVECTOR
                                    Beacons(BeaconCount).Origins(Beacons(BeaconCount).OriginCount) = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "translucent"
                                    Beacons(BeaconCount).Translucent = True
                                Case "dimension"
                                    Beacons(BeaconCount).Dimension.width = CSng(inArg(1))
                                    Beacons(BeaconCount).Dimension.height = CSng(inArg(2))
                                Case "percentxy"
                                    Beacons(BeaconCount).PercentXY.width = CSng(inArg(1)) * 50
                                    Beacons(BeaconCount).PercentXY.height = CSng(inArg(2)) * 50
                                
                                Case "filename"
                                    Beacons(BeaconCount).BeaconSkinCount = Beacons(BeaconCount).BeaconSkinCount + 1
                                    ReDim Preserve Beacons(BeaconCount).BeaconSkin(1 To Beacons(BeaconCount).BeaconSkinCount) As Direct3DTexture8
                                    Set Beacons(BeaconCount).BeaconSkin(Beacons(BeaconCount).BeaconSkinCount) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "beaconlight"
                                    Beacons(BeaconCount).BeaconLight = CLng(inArg(1))
                                    
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If
                            End Select
                        End If
                    Loop
                    
                    CreateSquare Beacons(BeaconCount).BeaconPlaq, 0, _
                        MakeVector(((Beacons(BeaconCount).Dimension.width * (Beacons(BeaconCount).PercentXY.width / 100)) / 2), 0, 0), _
                        MakeVector(-((Beacons(BeaconCount).Dimension.width * (Beacons(BeaconCount).PercentXY.width / 100)) / 2), 0, 0), _
                        MakeVector(-((Beacons(BeaconCount).Dimension.width * (Beacons(BeaconCount).PercentXY.width / 100)) / 2), (Beacons(BeaconCount).Dimension.height * (Beacons(BeaconCount).PercentXY.height / 100)), 0), _
                        MakeVector(((Beacons(BeaconCount).Dimension.width * (Beacons(BeaconCount).PercentXY.width / 100)) / 2), (Beacons(BeaconCount).Dimension.height * (Beacons(BeaconCount).PercentXY.height / 100)), 0)
                        
                    Set Beacons(BeaconCount).BeaconVBuf = DDevice.CreateVertexBuffer(Len(Beacons(BeaconCount).BeaconPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
                    D3DVertexBuffer8SetData Beacons(BeaconCount).BeaconVBuf, 0, Len(Beacons(BeaconCount).BeaconPlaq(0)) * 6, 0, Beacons(BeaconCount).BeaconPlaq(0)

                Case "billboard"
                    o = 0
                    BillBoardCount = BillBoardCount + 1
                    ReDim Preserve BillBoards(1 To BillBoardCount) As MyBoard
                    ReDim BillBoards(BillBoardCount).Plaq(0 To 5) As MyVertex

                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        LineNumber = LineNumber + 1
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "filename"
                                    BillBoards(BillBoardCount).SkinCount = BillBoards(BillBoardCount).SkinCount + 1
                                    ReDim Preserve BillBoards(BillBoardCount).Skin(1 To BillBoards(BillBoardCount).SkinCount) As Direct3DTexture8
                                    Set BillBoards(BillBoardCount).Skin(BillBoards(BillBoardCount).SkinCount) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "point1"
                                    BillBoards(BillBoardCount).Point1.X = CSng(inArg(1))
                                    BillBoards(BillBoardCount).Point1.Y = CSng(inArg(2))
                                    BillBoards(BillBoardCount).Point1.z = CSng(inArg(3))
                                    If UBound(inArg) = 5 Then
                                        o = 2
                                        BillBoards(BillBoardCount).Point1.tu = CSng(inArg(4))
                                        BillBoards(BillBoardCount).Point1.tv = CSng(inArg(5))
                                    End If
                                Case "point2"
                                    BillBoards(BillBoardCount).Point2.X = CSng(inArg(1))
                                    BillBoards(BillBoardCount).Point2.Y = CSng(inArg(2))
                                    BillBoards(BillBoardCount).Point2.z = CSng(inArg(3))
                                    If UBound(inArg) = 5 Then
                                        o = 2
                                        BillBoards(BillBoardCount).Point2.tu = CSng(inArg(4))
                                        BillBoards(BillBoardCount).Point2.tv = CSng(inArg(5))
                                    End If
                                Case "point3"
                                    BillBoards(BillBoardCount).Point3.X = CSng(inArg(1))
                                    BillBoards(BillBoardCount).Point3.Y = CSng(inArg(2))
                                    BillBoards(BillBoardCount).Point3.z = CSng(inArg(3))
                                    If UBound(inArg) = 5 Then
                                        o = 2
                                        BillBoards(BillBoardCount).Point3.tu = CSng(inArg(4))
                                        BillBoards(BillBoardCount).Point3.tv = CSng(inArg(5))
                                    End If
                                Case "point4"
                                    BillBoards(BillBoardCount).Point4.X = CSng(inArg(1))
                                    BillBoards(BillBoardCount).Point4.Y = CSng(inArg(2))
                                    BillBoards(BillBoardCount).Point4.z = CSng(inArg(3))
                                    If UBound(inArg) = 5 Then
                                        o = 2
                                        BillBoards(BillBoardCount).Point4.tu = CSng(inArg(4))
                                        BillBoards(BillBoardCount).Point4.tv = CSng(inArg(5))
                                    End If
                                Case "scalex"
                                    BillBoards(BillBoardCount).ScaleX = CSng(inArg(1))
                                    o = 1
                                Case "scaley"
                                    BillBoards(BillBoardCount).ScaleY = CSng(inArg(1))
                                    o = 1
                                Case "animated"
                                    BillBoards(BillBoardCount).AnimateMSecs = CSng(inArg(1))
                                    BillBoards(BillBoardCount).AnimateTimer = GetTimer
                                    BillBoards(BillBoardCount).AnimatePoint = 1
                                Case "translucent"
                                    BillBoards(BillBoardCount).Translucent = True
                                Case "multiplayer"
                                    BillBoards(BillBoardCount).Multiplayer = True
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If
                            End Select
                        End If
                    Loop
                    
                    If o = 1 Then
                        CreateSquare BillBoards(BillBoardCount).Plaq, 0, MakeVector(BillBoards(BillBoardCount).Point1.X, BillBoards(BillBoardCount).Point1.Y, BillBoards(BillBoardCount).Point1.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point2.X, BillBoards(BillBoardCount).Point2.Y, BillBoards(BillBoardCount).Point2.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point3.X, BillBoards(BillBoardCount).Point3.Y, BillBoards(BillBoardCount).Point3.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point4.X, BillBoards(BillBoardCount).Point4.Y, BillBoards(BillBoardCount).Point4.z), _
                                                        BillBoards(BillBoardCount).ScaleX, BillBoards(BillBoardCount).ScaleY
                    ElseIf o = 2 Then
                        CreateSquareEx BillBoards(BillBoardCount).Plaq, 0, BillBoards(BillBoardCount).Point1, BillBoards(BillBoardCount).Point2, BillBoards(BillBoardCount).Point3, BillBoards(BillBoardCount).Point4
                        
                    End If
                    BillBoards(BillBoardCount).Center = SquareCenter(MakeVector(BillBoards(BillBoardCount).Point1.X, BillBoards(BillBoardCount).Point1.Y, BillBoards(BillBoardCount).Point1.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point2.X, BillBoards(BillBoardCount).Point2.Y, BillBoards(BillBoardCount).Point2.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point3.X, BillBoards(BillBoardCount).Point3.Y, BillBoards(BillBoardCount).Point3.z), _
                                                        MakeVector(BillBoards(BillBoardCount).Point4.X, BillBoards(BillBoardCount).Point4.Y, BillBoards(BillBoardCount).Point4.z))
                    If o = 1 Or o = 2 Then
                        Set BillBoards(BillBoardCount).VBuf = DDevice.CreateVertexBuffer(Len(BillBoards(BillBoardCount).Plaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
                        D3DVertexBuffer8SetData BillBoards(BillBoardCount).VBuf, 0, Len(BillBoards(BillBoardCount).Plaq(0)) * 6, 0, BillBoards(BillBoardCount).Plaq(0)
                    End If

                Case "plane"

                    Do Until inData = ""
                        inName = Replace(RemoveNextArg(inData, vbCrLf), vbTab, "")
                        LineNumber = LineNumber + 1
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "highground"
                                    Set PlaneSkin(1) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "lowground"
                                    Set PlaneSkin(2) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "lowsky"
                                    Set PlaneSkin(3) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "highsky"
                                    Set PlaneSkin(4) = LoadTexture(AppPath & "Base\Lawn\" & inArg(1))
                                Case "void1"
                                    PlaneHole(1) = MakeVector(CSng(inArg(1)), 0, CSng(inArg(2)))
                                Case "void2"
                                    PlaneHole(2) = MakeVector(CSng(inArg(1)), 0, CSng(inArg(2)))
                                Case "void3"
                                    PlaneHole(3) = MakeVector(CSng(inArg(1)), 0, CSng(inArg(2)))
                                Case "void4"
                                    PlaneHole(4) = MakeVector(CSng(inArg(1)), 0, CSng(inArg(2)))
                                Case Else
                                    If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                                    End If

                            End Select
                        End If
                    Loop

                    Dim width As Single
                    Dim height As Single
                    width = Abs(PlaneHole(1).X) + PlaneHole(4).X
                    height = PlaneHole(2).z + Abs(PlaneHole(1).z)
                    
                    'left
                    CreateSquare PlanePlaq, 0, MakeVector(-(BlackBoundary + FadeDistance), 0, PlaneHole(1).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), 0, PlaneHole(2).z), _
                                                MakeVector(PlaneHole(2).X, 0, PlaneHole(2).z), _
                                                MakeVector(PlaneHole(1).X, 0, PlaneHole(1).z), _
                                                height / 128, _
                                                ((BlackBoundary + FadeDistance) - (width / 2)) / 128
                    
                    'front
                    CreateSquare PlanePlaq, 6, MakeVector(PlaneHole(2).X, 0, PlaneHole(2).z), _
                                                MakeVector(PlaneHole(2).X, 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(3).X, 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(3).X, 0, PlaneHole(3).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                width / 128
                                                
                    'right
                    CreateSquare PlanePlaq, 12, MakeVector(PlaneHole(3).X, 0, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(3).X, 0, PlaneHole(3).z), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, PlaneHole(3).z), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, PlaneHole(4).z), _
                                                height / 128, _
                                                ((BlackBoundary + FadeDistance) / 128)
                                                
                    'back
                    CreateSquare PlanePlaq, 18, MakeVector(PlaneHole(1).X, 0, -(BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(1).X, 0, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(4).X, 0, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(4).X, 0, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                (width / 128)
                                                
                    CreateSquare PlanePlaq, 24, MakeVector(-(BlackBoundary + FadeDistance), 0, -(BlackBoundary + FadeDistance)), _
                                                MakeVector(-(BlackBoundary + FadeDistance), 0, PlaneHole(1).z), _
                                                MakeVector(PlaneHole(1).X, 0, PlaneHole(1).z), _
                                                MakeVector(PlaneHole(1).X, 0, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                (((BlackBoundary + FadeDistance) - (width / 2)) / 128)

                    CreateSquare PlanePlaq, 30, MakeVector(-(BlackBoundary + FadeDistance), 0, PlaneHole(2).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(2).X, 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(2).X, 0, PlaneHole(2).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                ((BlackBoundary + FadeDistance) - (width / 2)) / 128

                    CreateSquare PlanePlaq, 36, MakeVector(PlaneHole(3).X, 0, PlaneHole(3).z), _
                                                MakeVector(PlaneHole(3).X, 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, (BlackBoundary + FadeDistance)), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, PlaneHole(3).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                (BlackBoundary + FadeDistance) / 128
                                                
                    CreateSquare PlanePlaq, 42, MakeVector(PlaneHole(4).X, 0, -(BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(4).X, 0, PlaneHole(4).z), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, PlaneHole(4).z), _
                                                MakeVector((BlackBoundary + FadeDistance), 0, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                ((BlackBoundary + FadeDistance) / 128)
                       
                    'left
                    CreateSquare PlanePlaq, 48, MakeVector(PlaneHole(1).X, -40, PlaneHole(1).z), _
                                                MakeVector(PlaneHole(2).X, -40, PlaneHole(2).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, PlaneHole(2).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, PlaneHole(1).z), _
                                                height / 128, _
                                                ((BlackBoundary + FadeDistance) - (width / 2)) / 128

                    'front
                    CreateSquare PlanePlaq, 54, MakeVector(PlaneHole(3).X, -40, PlaneHole(3).z), _
                                                MakeVector(PlaneHole(3).X, -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(2).X, -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(2).X, -40, PlaneHole(2).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                width / 128
                                                
                    'right
                    CreateSquare PlanePlaq, 60, MakeVector((BlackBoundary + FadeDistance), -40, PlaneHole(4).z), _
                                                MakeVector((BlackBoundary + FadeDistance), -40, PlaneHole(3).z), _
                                                MakeVector(PlaneHole(3).X, -40, PlaneHole(3).z), _
                                                MakeVector(PlaneHole(3).X, -40, PlaneHole(4).z), _
                                                height / 128, _
                                                ((BlackBoundary + FadeDistance) / 128)
                                                
                    'back
                    CreateSquare PlanePlaq, 66, MakeVector(PlaneHole(4).X, -40, -(BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(4).X, -40, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(1).X, -40, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(1).X, -40, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                (width / 128)
                                                
                    CreateSquare PlanePlaq, 72, MakeVector(PlaneHole(1).X, -40, -(BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(1).X, -40, PlaneHole(1).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, PlaneHole(1).z), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                (((BlackBoundary + FadeDistance) - (width / 2)) / 128)

                    CreateSquare PlanePlaq, 78, MakeVector(PlaneHole(2).X, -40, PlaneHole(2).z), _
                                                MakeVector(PlaneHole(2).X, -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(-(BlackBoundary + FadeDistance), -40, PlaneHole(2).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                ((BlackBoundary + FadeDistance) - (width / 2)) / 128

                    CreateSquare PlanePlaq, 84, MakeVector((BlackBoundary + FadeDistance), -40, PlaneHole(3).z), _
                                                MakeVector((BlackBoundary + FadeDistance), -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(3).X, -40, (BlackBoundary + FadeDistance)), _
                                                MakeVector(PlaneHole(3).X, -40, PlaneHole(3).z), _
                                                ((BlackBoundary + FadeDistance) - (height / 2)) / 128, _
                                                (BlackBoundary + FadeDistance) / 128
                                                
                    CreateSquare PlanePlaq, 90, MakeVector((BlackBoundary + FadeDistance), -40, -(BlackBoundary + FadeDistance)), _
                                                MakeVector((BlackBoundary + FadeDistance), -40, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(4).X, -40, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(4).X, -40, -(BlackBoundary + FadeDistance)), _
                                                (((BlackBoundary + FadeDistance) + height) / 128), _
                                                ((BlackBoundary + FadeDistance) / 128)
                          
                    CreateSquare PlanePlaq, 96, MakeVector(PlaneHole(4).X, -40, PlaneHole(4).z), _
                                                MakeVector(PlaneHole(3).X, -40, PlaneHole(3).z), _
                                                MakeVector(PlaneHole(2).X, -40, PlaneHole(2).z), _
                                                MakeVector(PlaneHole(1).X, -40, PlaneHole(1).z), _
                                                1, 1
                                                
                    CreateSquare PlanePlaq, 102, MakeVector(15, 10, -15), _
                                                MakeVector(15, 10, 15), _
                                                MakeVector(-15, 10, 15), _
                                                MakeVector(-15, 10, -15), _
                                                1, 1
                                                
                    Set PlaneVBuf = DDevice.CreateVertexBuffer(Len(PlanePlaq(0)) * 108, 0, FVF_RENDER, D3DPOOL_DEFAULT)
                    D3DVertexBuffer8SetData PlaneVBuf, 0, Len(PlanePlaq(0)) * 108, 0, PlanePlaq(0)
                Case Else
                    If (Not (Left(Replace(Replace(inItem, " ", ""), vbTab, ""), 1) = ";")) Then
                        AddMessage "Warning, Unknown Identifier in " & GetFileName(inFile) & " at Line " & LineNumber
                    End If
            End Select
        End If
    Loop
    Exit Sub
parseerror:
    AddMessage "Script Error at Line " & LineNumber
    If Not ConsoleVisible Then ConsoleToggle
    Err.Clear
End Sub

Public Function CreateMesh(ByVal FileName As String, Mesh As MyMesh)
    Dim TextureName As String

    Set Mesh.Mesh = D3DX.LoadMeshFromX(FileName, D3DXMESH_VB_MANAGED, DDevice, Nothing, Mesh.MaterialBuffer, Mesh.MaterialCount)

    If Mesh.MaterialCount > 0 Then
    
        ReDim Mesh.Materials(0 To Mesh.MaterialCount - 1) As D3DMATERIAL8
        ReDim Mesh.Textures(0 To Mesh.MaterialCount - 1) As Direct3DTexture8
    
        Dim d As ImgDimType
        
        Dim q As Integer
        For q = 0 To Mesh.MaterialCount - 1
            
            D3DX.BufferGetMaterial Mesh.MaterialBuffer, q, Mesh.Materials(q)
            Mesh.Materials(q).Ambient = Mesh.Materials(q).diffuse
       
            TextureName = D3DX.BufferGetTextureName(Mesh.MaterialBuffer, q)
            If (TextureName <> "") Then
                If ImageDimensions(AppPath & "Base\Lawn\" & TextureName, d) Then
                    Set Mesh.Textures(q) = D3DX.CreateTextureFromFileEx(DDevice, AppPath & "Base\Lawn\" & TextureName, d.width, d.height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Transparent, ByVal 0, ByVal 0)
                Else
                    Debug.Print "IMAGE ERROR: ImageDimensions - " & AppPath & "Base\Lawn\" & TextureName
                End If
            End If
            
        Next
    Else
        ReDim Mesh.Textures(0 To 0) As Direct3DTexture8
        ReDim Mesh.Materials(0 To 0) As D3DMATERIAL8
    End If
    
    D3DX.ComputeNormals Mesh.Mesh

    Dim vd As D3DVERTEXBUFFER_DESC
    Mesh.Mesh.GetVertexBuffer.GetDesc vd

    ReDim Mesh.Verticies(0 To ((vd.Size \ FVF_VERTEX_SIZE) - 1)) As D3DVERTEX
    D3DVertexBuffer8GetData Mesh.Mesh.GetVertexBuffer, 0, vd.Size, 0, Mesh.Verticies(0)

    Dim id As D3DINDEXBUFFER_DESC
    Mesh.Mesh.GetIndexBuffer.GetDesc id

    ReDim Mesh.Indicies(0 To ((id.Size \ 2) - 1)) As Integer
    D3DIndexBuffer8GetData Mesh.Mesh.GetIndexBuffer, 0, id.Size, 0, Mesh.Indicies(0)
End Function


Public Sub CleanupLawn()
    Dim q As Integer
    Dim o As Integer
    
    If ObjectCount > 0 Then
        Erase Objects
        ObjectCount = 0
    End If
    
    If UBound(Meshes) - LBound(Meshes) > 0 Then
        For o = 1 To UBound(Meshes)
            For q = LBound(Meshes(o).Textures) To UBound(Meshes(o).Textures)
                Set Meshes(o).Textures(q) = Nothing
            Next q

            Erase Meshes(o).Materials
            Erase Meshes(o).Textures
            Set Meshes(o).Mesh = Nothing
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
    
    If LightCount > 0 Then
        Erase Lights
        LightCount = 0
    End If

    If LightDataCount > 0 Then
        Erase LightDatas
        LightDataCount = 0
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
    
    Erase PlanePlaq
    Erase PlaneHole
    Set PlaneSkin(1) = Nothing
    Set PlaneSkin(2) = Nothing
    Set PlaneSkin(3) = Nothing
    Set PlaneSkin(4) = Nothing
    Erase PlaneSkin
    Set PlaneVBuf = Nothing
    
End Sub
