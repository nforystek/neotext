#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modShip"
#Const modShip = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Public Type ShipModel
    Count As Long
    Colors() As Long
    Wing1() As MyVertex
    Wing2() As MyVertex
    VBuffer1 As Direct3DVertexBuffer8
    VBuffer2 As Direct3DVertexBuffer8
End Type

Public Type PartnerMind
    TimeRotation As Single
    PlayRotation As Long
    
    AwayFromHome As Single
    
    NooneIsDancing As Single
    WantsToDance As Long
    
    FollowingMode As Boolean
End Type

Public Partner As MyPlayer
Public Clocker As PartnerMind

Public ShipN As Long
Public ShipP As Long

Private Models() As MyPlayer

Private Ships() As ShipModel
Private Skins() As Direct3DTexture8

Private Tails() As MyVertex
Private VBuffer3 As Direct3DVertexBuffer8

Private OrnamentSkin As Direct3DTexture8
Private OrnamentPlaq() As MyVertex
Private OrnamentVBuf As Direct3DVertexBuffer8

Private PowerupSkin As Direct3DTexture8
Private PowerupPlaq() As MyVertex
Private PowerupVBuf As Direct3DVertexBuffer8

Public Players() As MyPlayer
Public Partners() As MyPlayer
Public PlayerCount As Long

Public Sub RenderShips()
    RenderShip Player
    RenderPartner
    RenderPlayers
    RenderModels
End Sub

Public Function FlagPlayers()
    Dim cnt As Long
    If PlayerCount > 0 Then
        For cnt = 1 To PlayerCount
            Players(cnt).Flag = True
        Next
    End If
End Function

Public Function PlayersDone()
    Dim cnt As Long
    
    Dim cnt2 As Long
    Dim idx As Long
    If PlayerCount > 0 Then
        For cnt = 1 To PlayerCount
            If Not Players(cnt).Flag Then
                idx = idx + 1
            Else
                If Not (cnt = PlayerCount) Then
                    For cnt2 = cnt To (PlayerCount - 1)
                        Players(cnt2) = Players(cnt2 + 1)
                        Partners(cnt2) = Partners(cnt2 + 1)
                    Next
                End If
            End If
        Next
        If idx >= 1 Then
            ReDim Preserve Players(1 To idx) As MyPlayer
            ReDim Preserve Partners(1 To idx) As MyPlayer
            PlayerCount = idx
        Else
            Erase Players
            Erase Partners
            PlayerCount = 0
        End If
    End If
End Function
Public Sub LoadPlayer(ByRef newPlayer As MyPlayer)

    PlayerCount = PlayerCount + 1
    ReDim Preserve Players(1 To PlayerCount) As MyPlayer
    ReDim Players(PlayerCount).Spots(0 To 50) As D3DVECTOR
    Players(PlayerCount) = newPlayer
    
    ReDim Preserve Partners(1 To PlayerCount) As MyPlayer
    ReDim Partners(PlayerCount).Spots(0 To 0) As D3DVECTOR
    Partners(PlayerCount).Model = ShipP

End Sub

Public Function PlayerExists(ByVal PlayerName As String) As Integer
    Dim cnt As Long
    If PlayerCount > 0 Then
        For cnt = 1 To PlayerCount
            If Players(cnt).name = PlayerName Then
                PlayerExists = cnt
                Exit Function
            End If
        Next
    End If
End Function

Public Function RenderPlayers()
On Error GoTo multiplayererror

    Dim p As Integer
    Dim pTemp As String
    Dim allPlayers As String
    Dim playerData As String
    
    allPlayers = frmMain.PlayerMoveStates
    Do Until (allPlayers = "")
        playerData = RemoveNextArg(allPlayers, "|")
        If (Not (playerData = "")) Then
        
            pTemp = RemoveNextArg(playerData, ",")
       
            p = PlayerExists(pTemp)
            If Not p = 0 Then
                
                RenderShip Partners(p)
                RenderShip Players(p)
                
            End If
        End If
    Loop
    PlayersDone
        
    Exit Function
multiplayererror:
    Err.Clear
End Function

Public Sub PartnersDancing()
    If Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z) <= ZoneDistance Then
        Clocker.WantsToDance = Clocker.WantsToDance + 1
    End If
    Clocker.NooneIsDancing = Timer
End Sub

Private Function GetRotation(ByRef X As Single, ByRef z As Single) As Long
    Dim tmp As D3DVECTOR
    Dim rot As Long
    Dim dist As Single
    
    dist = Abs(Distance(Player.Object.Origin.X, 0, Player.Object.Origin.z, X, 0, z))
    For rot = 1 To 360
    
        tmp.X = Partner.Object.Origin.X + (Sin(-rot) * 1)
        tmp.z = Partner.Object.Origin.z + (Cos(-rot) * 1)
        
        If Abs(Distance(Player.Object.Origin.X, 0, Player.Object.Origin.z, tmp.X, 0, tmp.z)) <= dist Then
            dist = Abs(Distance(Player.Object.Origin.X, 0, Player.Object.Origin.z, tmp.X, 0, tmp.z))
            X = tmp.X
            z = tmp.z
            GetRotation = rot
        End If
            
    Next
End Function

Public Sub RenderPartner()

    If (TrapMouse And (Not ConsoleVisible)) Then

        Dim vecDirect As D3DVECTOR
        Dim dist As Single
        dist = Abs(Distance(Player.Object.Origin.X, 0, Player.Object.Origin.z, Partner.Object.Origin.X, 0, Partner.Object.Origin.z))

        If Clocker.FollowingMode Then
                
            If Partner.Object.Origin.Y < Player.Object.Origin.Y Then
                If Partner.Object.Origin.Y + IIf(Partner.MoveSpeed <= MoveSpeedMax, MoveSpeedMax, Partner.MoveSpeed) > Player.Object.Origin.Y Then
                    Partner.Object.Origin.Y = Player.Object.Origin.Y
                Else
                    Partner.Object.Origin.Y = Partner.Object.Origin.Y + IIf(Partner.MoveSpeed <= MoveSpeedMax, MoveSpeedMax, Partner.MoveSpeed)
                End If
            ElseIf Partner.Object.Origin.Y > Player.Object.Origin.Y Then
                If Partner.Object.Origin.Y - IIf(Partner.MoveSpeed <= MoveSpeedMax, MoveSpeedMax, Partner.MoveSpeed) < Player.Object.Origin.Y Then
                    Partner.Object.Origin.Y = Player.Object.Origin.Y
                Else
                    Partner.Object.Origin.Y = Partner.Object.Origin.Y - IIf(Partner.MoveSpeed <= MoveSpeedMax, MoveSpeedMax, Partner.MoveSpeed)
                End If
            End If
            
            If Not (dist <= (ZoneDistance / 3)) Then
                Dim dir As D3DVECTOR
                
                dir.X = Player.Object.Origin.X - Partner.Object.Origin.X
                dir.z = Player.Object.Origin.z - Partner.Object.Origin.z
    
                D3DXVec3Normalize dir, dir
                
                dir.X = dir.X + Partner.Object.Origin.X
                dir.z = dir.z + Partner.Object.Origin.z

                dist = GetRotation(dir.X, dir.z)
                If dist > 0 Then
                    Partner.Rotation = dist
                    Partner.MoveSpeed = Partner.MoveSpeed + 1
                    If Partner.MoveSpeed > Player.MoveSpeed Then Partner.MoveSpeed = Player.MoveSpeed
                End If
                
            Else
                Partner.MoveSpeed = 0
            End If
        
            If Partner.MoveSpeed > Player.MoveSpeed Then Partner.MoveSpeed = Player.MoveSpeed
        End If
            
        If (Not Clocker.FollowingMode) Then
        
            If (Clocker.TimeRotation = 0) Or (Timer - Clocker.TimeRotation >= 2) Then
                Clocker.TimeRotation = Timer
                Clocker.PlayRotation = Round(RandomPositive(1, 3), 0)
            End If
            If Clocker.PlayRotation > 1 Then
                If Clocker.PlayRotation = 2 Then
                    Partner.Rotation = Partner.Rotation - 0.01
                ElseIf Clocker.PlayRotation = 3 Then
                    Partner.Rotation = Partner.Rotation + 0.01
                End If
            End If

            If Partner.Rotation > (PI * 2) Then Partner.Rotation = Partner.Rotation - (PI * 2)
            If Partner.Rotation < -(PI * 2) Then Partner.Rotation = Partner.Rotation + (PI * 2)
            
        End If
        
        If (Not Clocker.FollowingMode) Then
            
            Dim PlaySpeed As Single
            
            PlaySpeed = CSng(IIf(RandomPositive(1, 2) <= 2, "-", "") & "0.05")

            If (Partner.MoveSpeed + PlaySpeed < PartnerMinSpeed) Or (Partner.MoveSpeed + PlaySpeed > PartnerMaxSpeed) Then PlaySpeed = -Abs(PlaySpeed)
            
            Partner.MoveSpeed = Partner.MoveSpeed + PlaySpeed
        
            If Partner.MoveSpeed < PartnerMinSpeed Then Partner.MoveSpeed = PartnerMinSpeed
            If Partner.MoveSpeed > PartnerMaxSpeed Then Partner.MoveSpeed = PartnerMaxSpeed
        
        End If
        
        vecDirect.X = Sin(D720 - Partner.Rotation)
        vecDirect.z = Cos(D720 - Partner.Rotation)
        AddActivity Partner.Object, vecDirect, Partner.MoveSpeed, GroundFriction
       
        If (Not Clocker.FollowingMode) Then

            If (Clocker.AwayFromHome = 0) Then
                Clocker.AwayFromHome = Timer
            ElseIf (Distance(Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z, 0, 0, 0) >= FadeDistance) And (Timer - Clocker.AwayFromHome >= (60 * 30)) Then
                If Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z) > FadeDistance Then
                    Clocker.AwayFromHome = Timer
                    Partner.Object.Origin.X = 0
                    Partner.Object.Origin.Y = 0
                    Partner.Object.Origin.z = 0
                End If
            ElseIf (Distance(Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z, 0, 0, 0) < FadeDistance) And (Timer - Clocker.AwayFromHome >= (60 * 30)) Then
                If Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z) > FadeDistance Then
                    Clocker.AwayFromHome = Timer
                    Partner.Object.Origin.X = IIf((Rnd < 0.5), -RandomPositive(FadeDistance, BlackBoundary), RandomPositive(FadeDistance, BlackBoundary))
                    Partner.Object.Origin.z = IIf((Rnd < 0.5), -RandomPositive(FadeDistance, BlackBoundary), RandomPositive(FadeDistance, BlackBoundary))
                End If
            End If
            
            If Clocker.NooneIsDancing = 0 Then
                Clocker.NooneIsDancing = Timer
            ElseIf (Timer - Clocker.NooneIsDancing) >= 20 Then
                Clocker.NooneIsDancing = Timer
                Clocker.WantsToDance = Clocker.WantsToDance - 1
            End If
        
            If (Clocker.WantsToDance >= 50) And _
                Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Partner.Object.Origin.X, Partner.Object.Origin.Y, Partner.Object.Origin.z) < ZoneDistance Then
                Clocker.FollowingMode = Not Clocker.FollowingMode
            End If
    
        End If
                
    End If

    RenderShip Partner

End Sub

Public Sub RenderShip(ByRef bot As MyPlayer)
    
    Dim o As Long
    Dim dist As Single
    Dim rgbval As Single

    If UBound(Objects) > 0 Then
        For o = LBound(Objects) To UBound(Objects)

            dist = Distance(Objects(o).Origin.X, Objects(o).Origin.Y, Objects(o).Origin.z, bot.Object.Origin.X, bot.Object.Origin.Y, bot.Object.Origin.z)
    
            If (dist <= (FadeDistance + WithInCityLimits)) Then

                If dist < ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits Then
            
                    SetAmbientRGB Ambient_HI, Ambient_HI, Ambient_HI, 2
            
                ElseIf dist >= ((FadeDistance - WithInCityLimits) / 10) + WithInCityLimits And dist < (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then

                    rgbval = (((((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits) - dist)
                    rgbval = Abs(-100 + ((((((FadeDistance - WithInCityLimits) / 10) * 7) - WithInCityLimits) / (((rgbval - Ambient_LO) / Ambient_HI) + Ambient_LO)) - 100))
                    rgbval = Round(rgbval, 0)
                    rgbval = ((Ambient_HI - Ambient_LO) * (rgbval / 100)) + Ambient_LO
                    
                    SetAmbientRGB rgbval, rgbval, rgbval, 2
                    
                ElseIf dist >= (((FadeDistance - WithInCityLimits) / 10) * 7) + WithInCityLimits Then
            
                    SetAmbientRGB Ambient_LO, Ambient_LO, Ambient_LO, 2
            
                End If

            End If
        Next
    End If
    
    Dim cnt As Long
    Dim dis As Single
    dis = Distance(bot.Object.Origin.X, bot.Object.Origin.Y, bot.Object.Origin.z, 0, 0, 0)

    Dim matModelScale As D3DMATRIX
    Dim matModelPos As D3DMATRIX
    
    Dim matModel1 As D3DMATRIX
    Dim matModel2 As D3DMATRIX
    Dim matModel3 As D3DMATRIX

    Dim matAdjust As D3DMATRIX
    
    Dim matAngle As D3DMATRIX
    Dim matShip As D3DMATRIX
    Dim matPos As D3DMATRIX
    
    Dim matLeft As D3DMATRIX
    Dim matRight As D3DMATRIX

    Dim matLeft1 As D3DMATRIX
    Dim matRight2 As D3DMATRIX
    Dim matTail As D3DMATRIX
    
    D3DXMatrixIdentity matAngle
    D3DXMatrixRotationY matAngle, -(bot.CameraAngle + bot.Rotation)
    
    D3DXMatrixTranslation matPos, bot.Object.Origin.X, bot.Object.Origin.Y, bot.Object.Origin.z
    If frmMain.Multiplayer Then D3DXMatrixScaling matModelScale, ScaleModelSize, ScaleModelSize, ScaleModelSize

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
    If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    End If
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 0, Skins(bot.Texture)
    
    Static DotRate As Double
    If bot.Trails Then
        If (Not ((bot.Spots(0).X = bot.Object.Origin.X) And _
            (bot.Spots(0).Y = bot.Object.Origin.Y) And _
            (bot.Spots(0).z = bot.Object.Origin.z))) Then
        
            bot.Spots(0) = bot.Object.Origin
            
            DotRate = Timer
            ReDim Preserve bot.Spots(0 To UBound(bot.Spots) + 1) As D3DVECTOR

        End If
            
        If (((DotRate = 0) Or ((Timer - DotRate) >= 1))) Or _
            ((UBound(bot.Spots) - LBound(bot.Spots)) >= 50) Then
            
            DotRate = Timer
            If (UBound(bot.Spots) - LBound(bot.Spots) > 0) Then
                ReDim Preserve bot.Spots(0 To UBound(bot.Spots) - 1) As D3DVECTOR
            End If

        End If
                    
        If Not (UBound(bot.Spots) = LBound(bot.Spots)) Then
            If (UBound(bot.Spots) - LBound(bot.Spots) > 0) Then
                For cnt = UBound(bot.Spots) To 1 Step -1
                    bot.Spots(cnt) = bot.Spots(cnt - 1)
                Next
            End If
        End If

    End If

    If bot.Trails Then
        If UBound(bot.Spots) - LBound(bot.Spots) > 0 Then
            For cnt = 1 To UBound(bot.Spots)
                D3DXMatrixTranslation matTail, bot.Spots(cnt).X, bot.Spots(cnt).Y, bot.Spots(cnt).z
                DDevice.SetTransform D3DTS_WORLD, matTail
        
                DDevice.SetStreamSource 0, VBuffer3, Len(Tails(0))
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    
                If (frmMain.Multiplayer And (dis <= 5000)) Then
                    D3DXMatrixIdentity matAdjust
                    D3DXMatrixTranslation matModel1, ScaleModelLocX, ScaleModelLocY, ScaleModelLocZ
                    DDevice.SetTransform D3DTS_WORLD, matModelScale
                    D3DXMatrixMultiply matAdjust, matModelScale, matModel1
                    D3DXMatrixTranslation matModel1, bot.Spots(cnt).X, bot.Spots(cnt).Y, bot.Spots(cnt).z
                    D3DXMatrixMultiply matModel1, matModel1, matAdjust
                    DDevice.SetTransform D3DTS_WORLD, matModel1
                    DDevice.SetStreamSource 0, VBuffer3, Len(Tails(0))
                    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, ((UBound(Tails) + 1) / 6)
                End If
                
            Next
        End If
    End If
    
    If bot.Model = ShipN Then
        If Not (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
            DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
            DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
        End If
        If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
        End If
    
        D3DXMatrixIdentity matShip
        D3DXMatrixMultiply matShip, matShip, matAngle
        
        D3DXMatrixIdentity matTail
        D3DXMatrixTranslation matTail, bot.Object.Origin.X, bot.Object.Origin.Y + 27, bot.Object.Origin.z
        
        D3DXMatrixMultiply matShip, matShip, matTail
        
        D3DXMatrixScaling matTail, 0.14, 0.14, 0.14
        D3DXMatrixMultiply matShip, matTail, matShip
        DDevice.SetTransform D3DTS_WORLD, matShip
                        
        DDevice.SetTexture 0, OrnamentSkin
        DDevice.SetStreamSource 0, OrnamentVBuf, Len(OrnamentPlaq(0))
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

        If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
            DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
            DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
        End If
        If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
        End If
    
    End If

    DDevice.SetTexture 0, Skins(bot.Texture)
    
    D3DXMatrixIdentity matShip
    If (Not bot.LeftFlap) Or bot.FlapLock Then
        D3DXMatrixRotationZ matShip, -20 * (PI / 180)
    End If
    D3DXMatrixMultiply matShip, matShip, matAngle
    If bot.FlapLock Then
        D3DXMatrixRotationY matLeft, 50 * (PI / 180)
        D3DXMatrixMultiply matShip, matShip, matLeft
    End If
    D3DXMatrixMultiply matShip, matShip, matPos
    DDevice.SetTransform D3DTS_WORLD, matShip
    DDevice.SetStreamSource 0, Ships(bot.Model).VBuffer1, Len(Ships(bot.Model).Wing1(0))
    For cnt = 1 To UBound(Ships(bot.Model).Colors)
        DDevice.SetTexture 0, IIf(Ships(bot.Model).Colors(cnt) = 0, Skins(bot.Texture), Skins(Ships(bot.Model).Colors(cnt)))
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, (cnt - 1) * 6, 2
    Next
    
    If (frmMain.Multiplayer And (dis <= 5000)) Then
        D3DXMatrixIdentity matAdjust
        D3DXMatrixTranslation matModel2, ScaleModelLocX, ScaleModelLocY, ScaleModelLocZ
        DDevice.SetTransform D3DTS_WORLD, matModelScale
        D3DXMatrixMultiply matAdjust, matModelScale, matModel2
        D3DXMatrixTranslation matModel2, bot.Object.Origin.X, bot.Object.Origin.Y, bot.Object.Origin.z
        D3DXMatrixMultiply matModel2, matModel2, matAdjust
        DDevice.SetTransform D3DTS_WORLD, matModel2
        
        D3DXMatrixIdentity matShip
        If (Not bot.LeftFlap) Or bot.FlapLock Then
            D3DXMatrixRotationZ matShip, -20 * (PI / 180)
        End If
        D3DXMatrixMultiply matShip, matShip, matAngle
        If bot.FlapLock Then
            D3DXMatrixRotationY matLeft, 50 * (PI / 180)
            D3DXMatrixMultiply matShip, matShip, matLeft
        End If
        D3DXMatrixMultiply matShip, matShip, matModel2
        DDevice.SetTransform D3DTS_WORLD, matShip
        DDevice.SetStreamSource 0, Ships(bot.Model).VBuffer1, Len(Ships(bot.Model).Wing1(0))
        For cnt = 1 To UBound(Ships(bot.Model).Colors)
            DDevice.SetTexture 0, IIf(Ships(bot.Model).Colors(cnt) = 0, Skins(bot.Texture), Skins(Ships(bot.Model).Colors(cnt)))
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, (cnt - 1) * 6, 2
        Next
    End If

    D3DXMatrixIdentity matShip
    If (Not bot.RightFlap) Or bot.FlapLock Then
        D3DXMatrixRotationZ matShip, 20 * (PI / 180)
    End If
    D3DXMatrixMultiply matShip, matShip, matAngle
    If bot.FlapLock Then
        D3DXMatrixRotationY matRight, -50 * (PI / 180)
        D3DXMatrixMultiply matShip, matShip, matRight
    End If
    D3DXMatrixMultiply matShip, matShip, matPos
    DDevice.SetTransform D3DTS_WORLD, matShip
    DDevice.SetStreamSource 0, Ships(bot.Model).VBuffer2, Len(Ships(bot.Model).Wing2(0))
    For cnt = 1 To UBound(Ships(bot.Model).Colors)
        DDevice.SetTexture 0, IIf(Ships(bot.Model).Colors(cnt) = 0, Skins(bot.Texture), Skins(Ships(bot.Model).Colors(cnt)))
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, (cnt - 1) * 6, 2
    Next
    
    If (frmMain.Multiplayer And (dis <= 5000)) Then
        D3DXMatrixIdentity matAdjust
        D3DXMatrixTranslation matModel3, ScaleModelLocX, ScaleModelLocY, ScaleModelLocZ
        DDevice.SetTransform D3DTS_WORLD, matModelScale
        D3DXMatrixMultiply matAdjust, matModelScale, matModel3
        D3DXMatrixTranslation matModel3, bot.Object.Origin.X, bot.Object.Origin.Y, bot.Object.Origin.z
        D3DXMatrixMultiply matModel3, matModel3, matAdjust
        DDevice.SetTransform D3DTS_WORLD, matModel3
        
        D3DXMatrixIdentity matShip
        If (Not bot.RightFlap) Or bot.FlapLock Then
            D3DXMatrixRotationZ matShip, 20 * (PI / 180)
        End If
        D3DXMatrixMultiply matShip, matShip, matAngle
        If bot.FlapLock Then
            D3DXMatrixRotationY matRight, -50 * (PI / 180)
            D3DXMatrixMultiply matShip, matShip, matRight
        End If
        D3DXMatrixMultiply matShip, matShip, matModel3
        DDevice.SetTransform D3DTS_WORLD, matShip
        DDevice.SetStreamSource 0, Ships(bot.Model).VBuffer2, Len(Ships(bot.Model).Wing2(0))
        For cnt = 1 To UBound(Ships(bot.Model).Colors)
            DDevice.SetTexture 0, IIf(Ships(bot.Model).Colors(cnt) = 0, Skins(bot.Texture), Skins(Ships(bot.Model).Colors(cnt)))
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, (cnt - 1) * 6, 2
        Next
    End If
    
End Sub

Public Function RenderModels()
    
    Dim cnt As Long
    Dim dist As Single
    
    For cnt = 0 To UBound(Models)
        dist = Distance(Models(cnt).Object.Origin.X, Models(cnt).Object.Origin.Y, Models(cnt).Object.Origin.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z)
        If (dist <= FadeDistance) Then
        
            Models(cnt).Rotation = Models(cnt).Rotation + 0.003
        
            If Models(cnt).Rotation > (PI * 2) Then Models(cnt).Rotation = Models(cnt).Rotation - (PI * 2)
            If Models(cnt).Rotation < -(PI * 2) Then Models(cnt).Rotation = Models(cnt).Rotation + (PI * 2)
            
            If frmMain.Multiplayer Then RenderShip Models(cnt)

            If (Not (Player.Model = 0)) And (cnt = 0) Then
                If (dist <= 100) Then Player.Model = cnt
                If Not frmMain.Multiplayer Then RenderShip Models(cnt)
            ElseIf ((cnt = 1) Or (cnt = 2)) Then
                If (InStr(Bodykits, cnt) > 0) And (dist <= 100) Then Player.Model = cnt
                If Not frmMain.Multiplayer Then RenderShip Models(cnt)
            ElseIf (InStr(Bodykits, cnt) > 0) And ((cnt >= 3) And (cnt <= 6)) Then
                If (dist <= 100) Then Player.Model = cnt
                If Not frmMain.Multiplayer Then RenderShip Models(cnt)
            End If

        End If
    Next
    
End Function

Public Function RenderPowerup()

    Static SputterMut As Boolean
    Static FrameCards As Long
     
    Dim HasEmblem As Boolean
    
    HasEmblem = NearSymbol(Player)
    
    If (SputterMut = 0) Then
        If HasEmblem Then SputterMut = Timer
    End If
    
    If Not (SputterMut = 0) Then

        If HasEmblem Then
            FrameCards = Abs(FrameCards)
        Else
            FrameCards = -Abs(FrameCards)
        End If
        DrawPowerup Player, FrameCards
        
        If Not HasEmblem And FrameCards = 0 Then
            SputterMut = 0
        End If

    End If

End Function

Private Function DrawPowerup(ByRef bot As MyPlayer, ByRef Frame As Long)
    Dim matTail As D3DMATRIX
    Dim matAngle As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matShip As D3DMATRIX

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    If Not (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
    End If
    If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    End If
    
    DDevice.SetMaterial GenericMaterial
        
    D3DXMatrixIdentity matAngle
    D3DXMatrixIdentity matPitch

    D3DXMatrixIdentity matShip
    D3DXMatrixIdentity matTail

    D3DXMatrixRotationY matAngle, -bot.CameraAngle

    D3DXMatrixIdentity matShip
    D3DXMatrixMultiply matShip, matShip, matAngle
    
    D3DXMatrixIdentity matTail
    D3DXMatrixTranslation matTail, bot.Object.Origin.X, bot.Object.Origin.Y + IIf(bot.Model = ShipN, 43, 27), bot.Object.Origin.z
    
    D3DXMatrixMultiply matShip, matShip, matTail

    If Frame >= 0 Then
        If 0.001 + (Abs(Frame) * 0.006) >= 0.14 Then
            D3DXMatrixScaling matTail, 0.14, 0.14, 0.14
        Else
            Frame = Frame + 1
            D3DXMatrixScaling matTail, 0.001 + (Abs(Frame) * 0.006), 0.001 + (Abs(Frame) * 0.006), 0.001 + (Abs(Frame) * 0.006)
        End If
    ElseIf Frame < 0 Then
        If Abs(Frame) >= 0 Then
            Frame = Frame + 1
            D3DXMatrixScaling matTail, 0.001 + (Abs(Frame) * 0.006), 0.001 + (Abs(Frame) * 0.006), 0.001 + (Abs(Frame) * 0.006)
        End If
    End If
   
    If Frame <> 0 Then
    
        D3DXMatrixMultiply matShip, matTail, matShip
        DDevice.SetTransform D3DTS_WORLD, matShip

        DDevice.SetTexture 0, PowerupSkin
        DDevice.SetStreamSource 0, PowerupVBuf, Len(PowerupPlaq(0))
        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        DDevice.SetRenderState D3DRS_ZENABLE, 1
    End If
    
    If (DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER) = D3DTEXF_NONE) Then
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
    If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    End If
End Function

Public Sub CreateShips()

    ReDim Models(0 To 6) As MyPlayer

    Models(0).Model = 0
    Models(0).Object.Origin.X = 795
    Models(0).Object.Origin.z = -2063
    
    Models(1).Model = 1
    Models(1).Object.Origin.X = -167
    Models(1).Object.Origin.z = -1852
    
    Models(2).Model = 2
    Models(2).Object.Origin.X = -167
    Models(2).Object.Origin.z = -2105

    Models(3).Model = 3
    Models(3).Object.Origin.X = 775
    Models(3).Object.Origin.z = -1404
    
    Models(4).Model = 4
    Models(4).Object.Origin.X = 1038
    Models(4).Object.Origin.z = -1404
   
    Models(5).Model = 5
    Models(5).Object.Origin.X = 1293
    Models(5).Object.Origin.z = -1404
   
    Models(6).Model = 6
    Models(6).Object.Origin.X = 795
    Models(6).Object.Origin.Y = 227
    Models(6).Object.Origin.z = -2063
    
    ReDim Player.Spots(0 To 0) As D3DVECTOR
    
    ReDim Skins(0 To 9) As Direct3DTexture8
    Set Skins(0) = LoadTexture(AppPath & "Base\Player\01.bmp")
    Set Skins(1) = LoadTexture(AppPath & "Base\Player\01.bmp")
    Set Skins(2) = LoadTexture(AppPath & "Base\Player\02.bmp")
    Set Skins(3) = LoadTexture(AppPath & "Base\Player\03.bmp")
    Set Skins(4) = LoadTexture(AppPath & "Base\Player\04.bmp")
    Set Skins(5) = LoadTexture(AppPath & "Base\Player\05.bmp")
    Set Skins(6) = LoadTexture(AppPath & "Base\Player\06.bmp")
    Set Skins(7) = LoadTexture(AppPath & "Base\Player\07.bmp")
    Set Skins(8) = LoadTexture(AppPath & "Base\Player\08.bmp")
    Set Skins(9) = LoadTexture(AppPath & "Base\Player\09.bmp")
    
    ReDim Preserve Tails(0 To 5) As MyVertex
    CreateSquare Tails, 0, MakeVector(-2, 1, -2), MakeVector(-2, 1, 2), MakeVector(2, 1, 2), MakeVector(2, 1, -2)
    Set VBuffer3 = DDevice.CreateVertexBuffer(Len(Tails(0)) * (UBound(Tails) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData VBuffer3, 0, Len(Tails(0)) * (UBound(Tails) + 1), 0, Tails(0)

    Dim cnt As Long
    Do Until Not PathExists(AppPath & "Base\Player\ship" & Trim(cnt) & ".sx", True)
        LoadShip cnt
        cnt = cnt + 1
    Loop
    
    LoadShip "N"

    LoadShip "P"
    
    Set OrnamentSkin = LoadTexture(AppPath & "Base\Player\NN.bmp")
    ReDim OrnamentPlaq(0 To 5) As MyVertex
    CreateSquare OrnamentPlaq, 0, MakeVector(50, 0, 0), MakeVector(-50, 0, 0), MakeVector(-50, 100, 0), MakeVector(50, 100, 0)
    Set OrnamentVBuf = DDevice.CreateVertexBuffer(Len(OrnamentPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData OrnamentVBuf, 0, Len(OrnamentPlaq(0)) * 6, 0, OrnamentPlaq(0)
    
    Set PowerupSkin = LoadTexture(AppPath & "Base\Player\UP.bmp")
    ReDim PowerupPlaq(0 To 5) As MyVertex
    CreateSquare PowerupPlaq, 0, MakeVector(50, 0, 0), MakeVector(-50, 0, 0), MakeVector(-50, 100, 0), MakeVector(50, 100, 0)
    Set PowerupVBuf = DDevice.CreateVertexBuffer(Len(PowerupPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData PowerupVBuf, 0, Len(PowerupPlaq(0)) * 6, 0, PowerupPlaq(0)
    
    If Not ((Player.Model >= LBound(Ships)) And (Player.Model <= UBound(Ships))) Then Player.Model = 0
    If Player.Model = UBound(Ships) Then Player.Model = 0
    
    ShipN = (UBound(Ships) - 1)
    ShipP = UBound(Ships)
    
    ReDim Partner.Spots(0 To 0) As D3DVECTOR
    Partner.Model = ShipP
End Sub

Private Sub LoadShip(ByVal Index As String)
    Dim FileName As String
    FileName = AppPath & "Base\Player\ship" & Trim(Index) & ".sx"
    
    Dim X1 As Single
    Dim X2 As Single
    Dim Y1 As Single
    Dim Y2 As Single
    Dim Color As Long
    
    Dim inData As String
    Dim inLine As String
    inData = ReadFile(FileName)
    
    If Not IsNumeric(Index) Then Index = CInt(UBound(Ships) + 1)
    
    ReDim Preserve Ships(0 To Index) As ShipModel
    ReDim Ships(Index).Colors(0 To 0) As Long
    Ships(Index).Count = 0
    
    Do Until inData = ""
    
        inLine = RemoveNextArg(inData, vbCrLf)
        If Not (inLine = "") Then
        
            X1 = RemoveNextArg(inLine, ":")
            Y1 = RemoveNextArg(inLine, ":")
            X2 = RemoveNextArg(inLine, ":")
            Y2 = RemoveNextArg(inLine, ":")
            Color = RemoveNextArg(inLine, ":")
        
            If X1 < X2 Then Swap X1, X2
            If Y1 > Y2 Then Swap Y1, Y2
            
            ReDim Preserve Ships(Index).Colors(0 To UBound(Ships(Index).Colors) + 1) As Long
            Ships(Index).Colors(UBound(Ships(Index).Colors)) = Color
            
            ReDim Preserve Ships(Index).Wing1(0 To Ships(Index).Count + 6) As MyVertex
            ReDim Preserve Ships(Index).Wing2(0 To Ships(Index).Count + 6) As MyVertex
    
            CreateSquare Ships(Index).Wing1, Ships(Index).Count, MakeVector(X1, Y1, 0), MakeVector(X2, Y1, 0), MakeVector(X2, Y2, 0), MakeVector(X1, Y2, 0)
            CreateSquare Ships(Index).Wing2, Ships(Index).Count, MakeVector(-X2, Y1, 0), MakeVector(-X1, Y1, 0), MakeVector(-X1, Y2, 0), MakeVector(-X2, Y2, 0)
            
            Ships(Index).Count = Ships(Index).Count + 6
        End If
        
    Loop

    Set Ships(Index).VBuffer1 = DDevice.CreateVertexBuffer(Len(Ships(Index).Wing1(0)) * Ships(Index).Count, 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Ships(Index).VBuffer1, 0, Len(Ships(Index).Wing1(0)) * Ships(Index).Count, 0, Ships(Index).Wing1(0)

    Set Ships(Index).VBuffer2 = DDevice.CreateVertexBuffer(Len(Ships(Index).Wing2(0)) * Ships(Index).Count, 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Ships(Index).VBuffer2, 0, Len(Ships(Index).Wing2(0)) * Ships(Index).Count, 0, Ships(Index).Wing2(0)
    
End Sub

Public Sub CleanupShips()

    Dim cnt As Long
    If UBound(Ships) - LBound(Ships) > 0 Then
        For cnt = LBound(Ships) To UBound(Ships)
            Set Ships(cnt).VBuffer1 = Nothing
            Set Ships(cnt).VBuffer2 = Nothing
            Erase Ships(cnt).Wing1
            Erase Ships(cnt).Wing2
            Erase Ships(cnt).Colors
        Next
    End If
    
    Erase Models
    
    Erase Player.Spots
    Erase Skins
    
    Set VBuffer3 = Nothing
    
    Set OrnamentSkin = Nothing
    Set OrnamentVBuf = Nothing
    Erase OrnamentPlaq
    
    Set PowerupSkin = Nothing
    Set PowerupVBuf = Nothing
    Erase PowerupPlaq
 
    Erase Partner.Spots
    
    If PlayerCount > 0 Then
        For cnt = 1 To PlayerCount
            Erase Players(cnt).Spots
            Erase Partners(cnt).Spots
        Next
        Erase Players
        Erase Partners
        PlayerCount = 0
    End If
End Sub