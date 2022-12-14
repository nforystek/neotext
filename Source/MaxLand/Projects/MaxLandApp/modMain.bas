Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module


Public Enum Playmode
    Spectator = 0
    ThirdPerson = 1
    FirstPerson = 2
    CameraMode = 3
End Enum

Public Perspective As Playmode
Public Resolution As String
Public FullScreen As Boolean
Public SoundFX As Boolean
Public Ambient As Boolean

Public DebugMode As Boolean
Public CameraClip As Boolean
Public CameraZoom As Single

Public Surface As Boolean
Public AspectRatio As Single
Public TrapMouse As Boolean
Public PauseGame As Boolean
Public StopGame As Boolean
Public ShowHelp As Boolean
Public ShowStat As Boolean
Public ShowCredits As Boolean

Public Player As MyPlayer

Public FPSTimer As Double
Public FPSCount As Long
Public FPSRate As Long

Public db As clsDatabase
Public rs As ADODB.Recordset

Public dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public DDevice As Direct3DDevice8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public Display As D3DDISPLAYMODE
Public DSound As DirectSound8

Public CurrentLoadedLevel As String
Public PixelShaderDefault As Long
Public PixelShaderDiffuse As Long

Public Sub Main()

    On Error GoTo fault:
    
    Dim inCmd As String
    inCmd = Command
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)

    Select Case inCmd
        Case "setupreset"
            ResetDB
            CompactDB
            End
        Case "backupdb"
            BackupDB
            End
        Case "restoredb"
            RestoreDB
            End
    End Select
    
    Set db = New clsDatabase
    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    If db.rsEnd(rs) Then inCmd = "setup"
    
    If Trim(LCase(inCmd)) = "setup" Then
        frmSetup.Show
        Do While frmSetup.Visible
           DoTasks
        Loop
        If frmSetup.Play Then inCmd = ""
        Unload frmSetup
    End If
    
    If inCmd = "" Then
        db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
        
        If Not db.rsEnd(rs) Then
            Resolution = CStr(rs("Resolution"))
            FullScreen = CBool(Not rs("Windowed"))
            Perspective = CLng(rs("Perspective"))
            Surface = CBool(rs("Surface"))
        End If
        CameraClip = True

        SetActivity GlobalGravityDirect, Actions.Directing, MakeVector(0, -0.2, 0), 1
        SetActivity GlobalGravityRotate, Actions.Rotating, MakeVector(0, 0, 0), 0
        SetActivity GlobalGravityScaled, Actions.Scaling, MakeVector(0, 0, 0), 0
        
        SetActivity LiquidGravityDirect, Actions.Directing, MakeVector(0, -0.005, 0), 2
        SetActivity LiquidGravityRotate, Actions.Rotating, MakeVector(0, 0, 0), 0
        SetActivity LiquidGravityScaled, Actions.Scaling, MakeVector(0, 0, 0), 0
        
        db.rsClose rs
                
        FPSCount = 36
        TrapMouse = True
        Player.CameraZoom = 5

        Load frmMain
                
        frmMain.width = CSng(NextArg(Resolution, "x")) * Screen.TwipsPerPixelY
        frmMain.height = CSng(RemoveArg(Resolution, "x")) * Screen.TwipsPerPixelX
        AspectRatio = CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))
    
        On Error GoTo fault
        InitDirectX
        InitGameData
        On Error GoTo 0
        
        frmMain.Show

        Do While Not StopGame
            
            If PauseGame Then
                
                If Not (frmMain.WindowState = 1) Then
                    If TestDirectX Then
                    
                        On Error Resume Next
                        DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
                        DDevice.BeginScene
                        DDevice.EndScene
                        DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                        If (Err.Number = 0) Then
                            PauseGame = False
                        Else
                            Err.Clear
                        End If
                        On Error GoTo 0
                            
                    End If
                    
                    If (Not PauseGame) And (Err.Number = 0) Then
                        InitGameData
                    Else
                        TermDirectX
                    End If
                Else
                    DoTasks
                End If
                
            Else

                On Error GoTo Render
                DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
                                                   
                DDevice.BeginScene
                
                SetupWorld
                
                RenderActive
                RenderPlanes
                RenderWorld
                RenderPlayer
                RenderBoards
                RenderLucent
                RenderBeacons
                RenderPortals
                RenderCameras

                On Error GoTo 0
        
                InputMove
                ResetMotion
                InputScene

                If Not PauseGame Then
                    
                    RenderInfo
                    RenderCmds
   
                    DDevice.EndScene

                    On Error Resume Next
                    DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                    
                    FPSCount = FPSCount + 1
                    If (FPSTimer = 0) Or ((Timer - FPSTimer) >= 1) Then
                        FPSTimer = Timer
                        FPSRate = FPSCount
                        FPSCount = 0
                    End If
                    If Err.Number Then
                        Err.Clear
                        DoPauseGame
                    End If
                    On Error GoTo 0
                Else
                    DoTasks
                End If
            End If
            
            If D3DWindow.Windowed Then DoTasks
        
        Loop
        
        TermGameData
        TermDirectX
        
        Unload frmMain
    End If

    Set db = Nothing
    End
    
Exit Sub
fault:
    
    MsgBox "There was an error initializing the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.Number & ", " & Err.Description, vbOKOnly + vbInformation, App.Title
    TermDirectX
    Err.Clear

    End
Exit Sub
Render:

    TermGameData
    TermDirectX

    Unload frmMain
    
    MsgBox "There was an error trying to run the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.Number & ", " & Err.Description, vbOKOnly + vbInformation, App.Title

    Err.Clear
    End
End Sub

Public Sub SetupWorld()
On Error GoTo WorldError

    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX

    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matWorld As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matTemp As D3DMATRIX

    DDevice.SetTransform D3DTS_WORLD, matWorld
    D3DXMatrixIdentity matWorld

    DDevice.SetTransform D3DTS_WORLD1, matWorld
    
    D3DXMatrixMultiply matTemp, matWorld, matWorld
    D3DXMatrixRotationY matRotation, 0.5
    D3DXMatrixRotationX matPitch, 0.5
    
    
    D3DXMatrixIdentity matWorld
    D3DXMatrixMultiply matLook, matRotation, matPitch
    
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    
    If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex > 0 And Player.CameraIndex <= CameraCount)) Or (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0)) Then
        
        D3DXMatrixRotationY matRotation, Cameras(Player.CameraIndex).Angle
        D3DXMatrixRotationX matPitch, Cameras(Player.CameraIndex).Pitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        D3DXMatrixTranslation matPos, -Cameras(Player.CameraIndex).Location.X, -Cameras(Player.CameraIndex).Location.Y, -Cameras(Player.CameraIndex).Location.z
        D3DXMatrixMultiply matLook, matPos, matLook
        D3DXMatrixTranslation matPos, -Player.Object.Origin.X, -Player.Object.Origin.Y + 0.2, -Player.Object.Origin.z
        
    Else
    
        D3DXMatrixRotationY matRotation, Player.CameraAngle
        D3DXMatrixRotationX matPitch, Player.CameraPitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        If Player.CameraPitch > 0 Then
        
            D3DXMatrixTranslation matPos, -Player.Object.Origin.X, -Player.Object.Origin.Y, -Player.Object.Origin.z
            D3DXMatrixMultiply matLook, matPos, matLook
        Else
            D3DXMatrixTranslation matPos, -Player.Object.Origin.X, -Player.Object.Origin.Y + 0.2, -Player.Object.Origin.z
            D3DXMatrixMultiply matLook, matPos, matLook
        
        End If
        
    End If
    
    lCulledFaces = 0
    lCullCalls = 0
    
    If ((Perspective = Playmode.ThirdPerson) Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0))) Then
    
        If (CameraClip Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not ((Perspective = Spectator) Or DebugMode)) Then

            If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0)) Then

                Player.CameraAngle = 3
            
            End If
        
            Dim cnt As Long
            Dim cnt2 As Long

            Dim Face As Long
            Dim zoom As Single
            Dim factor As Single

            Dim verts(0 To 2) As D3DVECTOR
            Dim touched As Boolean

            For cnt = 1 To lngFaceCount - 1
                            On Error GoTo isdivcheck0
                                sngFaceVis(3, cnt) = 0
                                GoTo notdivcheck0
isdivcheck0:
                                If Err.Number = 11 Then Resume
notdivcheck0:
                                If Err Then Err.Clear
                                On Error GoTo 0
                
            Next

            zoom = 0.2
            factor = 0.5

            Do

                verts(0) = MakeVector(Player.Object.Origin.X, _
                                            Player.Object.Origin.Y - 0.2, _
                                            Player.Object.Origin.z)

                verts(1) = MakeVector(Player.Object.Origin.X - (Sin(D720 - Player.CameraAngle) * (zoom + factor)), _
                                            Player.Object.Origin.Y - 0.2 + (Tan(D720 - Player.CameraPitch) * (zoom + factor)), _
                                            Player.Object.Origin.z - (Cos(D720 - Player.CameraAngle) * (zoom + factor)))

                verts(2) = MakeVector(Player.Object.Origin.X - (Sin(D720 - Player.CameraAngle)), _
                                      Player.Object.Origin.Y - 0.1 + (Tan(D720 - Player.CameraPitch) * zoom), _
                                      Player.Object.Origin.z - (Cos(D720 - Player.CameraAngle)))

                sngCamera(0, 0) = Player.Object.Origin.X
                sngCamera(0, 1) = Player.Object.Origin.Y
                sngCamera(0, 2) = Player.Object.Origin.z

                sngCamera(1, 0) = 1
                sngCamera(1, 1) = -1
                sngCamera(1, 2) = -1

                sngCamera(2, 0) = -1
                sngCamera(2, 1) = 1
                sngCamera(2, 2) = -1

                If lngFaceCount > 0 Then
                    lCulledFaces = lCulledFaces + Culling(2, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
                    lCullCalls = lCullCalls + 1
                End If

                If (ObjectCount > 0) Then
                    For cnt = 1 To ObjectCount
                        If ((Not (Objects(cnt).Effect = Collides.Ground)) And (Not (Objects(cnt).Effect = Collides.InDoor))) And (Objects(cnt).CollideIndex > -1) Then
                            For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                            On Error GoTo isdivcheck1
                                sngFaceVis(3, cnt2) = 0
                                GoTo notdivcheck1
isdivcheck1:
                                If Err.Number = 11 Then Resume
notdivcheck1:
                                If Err Then Err.Clear
                                On Error GoTo 0
                                
                            Next
                        ElseIf (Objects(cnt).Effect = Collides.Ground) And (Objects(cnt).CollideIndex > -1) And Objects(cnt).Visible Then
                            For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                                If Not (((sngFaceVis(0, cnt2) = 0) Or (sngFaceVis(0, cnt2) = 1) Or (sngFaceVis(0, cnt2) = -1)) And _
                                    ((sngFaceVis(1, cnt2) = 0) Or (sngFaceVis(1, cnt2) = 1) Or (sngFaceVis(1, cnt2) = -1)) And _
                                    ((sngFaceVis(2, cnt2) = 0) Or (sngFaceVis(2, cnt2) = 1) Or (sngFaceVis(2, cnt2) = -1))) Then
                                    sngFaceVis(3, cnt2) = 2
                                End If
                            Next
                        End If
                    Next
                    If (Player.Object.CollideIndex > -1) Then
                        For cnt2 = Player.Object.CollideIndex To (Player.Object.CollideIndex + Meshes(Player.Object.MeshIndex).Mesh.GetNumFaces) - 1
                            On Error GoTo isdivcheck2
                                sngFaceVis(3, cnt2) = 0
                                GoTo notdivcheck2
isdivcheck2:
                                If Err.Number = 11 Then Resume
notdivcheck2:
                                If Err Then Err.Clear
                                On Error GoTo 0

                        Next
                    End If
                End If

                Face = AddCollisionEx(verts, 1)
                touched = TestCollisionEx(Face, 2)
                DelCollisionEx Face, 1

                If ((Not touched) And (zoom < Player.CameraZoom)) Then zoom = zoom + factor

            Loop Until ((touched) Or (zoom >= Player.CameraZoom))

            If (touched And (zoom > 0.2)) Then zoom = zoom + -factor

            D3DXMatrixTranslation matTemp, 0, 0.2, zoom
            D3DXMatrixMultiply matView, matLook, matTemp

            Player.Object.WireFrame = (zoom < 0.8)

        Else
            D3DXMatrixTranslation matTemp, 0, 0, IIf(Not ((Perspective = Spectator) Or DebugMode), Player.CameraZoom, 0)
            D3DXMatrixMultiply matView, matLook, matTemp
        End If
        DDevice.SetTransform D3DTS_VIEW, matView
    Else
        DDevice.SetTransform D3DTS_VIEW, matLook
    End If

    D3DXMatrixPerspectiveFovLH matProj, FOVY, AspectRatio, 0.01, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
    
    Exit Sub
WorldError:
    If Err.Number = 6 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub InitDirectX()

    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    Set D3DX = New D3DX8

    InitialDevice frmMain.hwnd
    
    Set DSound = dx.DirectSoundCreate("")
    DSound.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
        
End Sub

Private Sub InitialDevice(ByVal hwnd As Long)

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display

    D3DWindow.BackBufferCount = 1

    D3DWindow.BackBufferWidth = CDbl(NextArg(Resolution, "x"))
    D3DWindow.BackBufferHeight = CDbl(RemoveArg(Resolution, "x"))

'    D3DWindow.BackBufferWidth = Screen.width / Screen.TwipsPerPixelX
'    D3DWindow.BackBufferHeight = Screen.height / Screen.TwipsPerPixelY
    
    D3DWindow.BackBufferFormat = Display.Format
    D3DWindow.MultiSampleType = D3DMULTISAMPLE_NONE
    
    If Not FullScreen Then
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        D3DWindow.FullScreen_PresentationInterval = 0
    Else
        D3DWindow.Windowed = 0
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    End If
    D3DWindow.hDeviceWindow = hwnd
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    D3DWindow.EnableAutoDepthStencil = True
    frmMain.Show
    
    On Error Resume Next
    Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
    If Err.Number <> 0 Then
        Err.Clear
        Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
        If Err.Number <> 0 Then
            Err.Clear
            Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
        End If
    End If
    On Error GoTo 0
    
    If Not DDevice Is Nothing Then
            
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_LIGHTING, 1
        DDevice.SetRenderState D3DRS_DITHERENABLE, False
        DDevice.SetRenderState D3DRS_EDGEANTIALIAS, False
    
        DDevice.SetRenderState D3DRS_INDEXVERTEXBLENDENABLE, False
        DDevice.SetRenderState D3DRS_VERTEXBLEND, False
    
        DDevice.SetRenderState D3DRS_CLIPPING, 1
    
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    
        DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        
        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        
        DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    
        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
        DDevice.SetRenderState D3DRS_ALPHAREF, Transparent
        DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
        DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESSEQUAL
    
        DDevice.SetRenderState D3DRS_FOGENABLE, 0
        DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
        DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
        DDevice.SetRenderState D3DRS_RANGEFOGENABLE, False
        DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(FadeDistance / 4)
        DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(FadeDistance)
        DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(255, 184, 200, 225)
    
        If frmMain.WindowState = vbMinimized Then frmMain.WindowState = IIf(FullScreen, vbMaximized, vbNormal)

'        On Error Resume Next
'        Set DSound = dx.DirectSoundCreate("")
'        DSound.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY
'        If Err.Number <> 0 Then Err.Clear
'        On Error GoTo 0

        Dim shArray() As Long
        Dim shLength As Long
        Dim shCode As D3DXBuffer

        Set shCode = D3DX.AssembleShader("ps.1.0" & vbCrLf & _
                                            "tex t0" & vbCrLf & _
                                            "mul r0, t0,v0" & vbCrLf, 0, Nothing)
        shLength = shCode.GetBufferSize() / 4
        ReDim shArray(shLength - 1) As Long
        D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
        PixelShaderDefault = DDevice.CreatePixelShader(shArray(0))
        Set shCode = Nothing

        Set shCode = D3DX.AssembleShader("ps.1.1" & vbCrLf & _
                                            "tex t0" & vbCrLf & _
                                            "mov r0,t0" & vbCrLf, 0, Nothing)
        shLength = shCode.GetBufferSize() / 4
        ReDim shArray(shLength - 1) As Long
        D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
        PixelShaderDiffuse = DDevice.CreatePixelShader(shArray(0))
        Set shCode = Nothing
    
    End If
    
End Sub


'Public Sub InitDirectX()
'
'    Set dx = New DirectX8
'    Set D3D = dx.Direct3DCreate
'    Set D3DX = New D3DX8
'
'    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display
'
'    D3DWindow.BackBufferCount = 1
'    D3DWindow.BackBufferWidth = CDbl(NextArg(Resolution, "x"))
'    D3DWindow.BackBufferHeight = CDbl(RemoveArg(Resolution, "x"))
'    D3DWindow.BackBufferFormat = Display.Format
'
'    If Not FullScreen Then
'        D3DWindow.Windowed = 1
'        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
'    Else
'        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
'        D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
'       ' D3DWindow.FullScreen_RefreshRateInHz = 0
'    End If
'
'    D3DWindow.MultiSampleType = D3DMULTISAMPLE_NONE
'
''    If Not FullScreen Then
''        D3DWindow.Windowed = 1
''        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
''    Else
''        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
''        D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
''    End If
'
'    D3DWindow.hDeviceWindow = frmMain.hwnd
'
'    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
'    D3DWindow.EnableAutoDepthStencil = True
'
'    On Error Resume Next
'    Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
'    If Err.Number <> 0 Then
'        Err.Clear
'        Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
'        If Err.Number <> 0 Then
'            Err.Clear
'            Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
'        End If
'    End If
'    On Error GoTo 0
'
'    If Not DDevice Is Nothing Then
'
'        DDevice.SetRenderState D3DRS_ZENABLE, 1
'        DDevice.SetRenderState D3DRS_LIGHTING, 1
'        DDevice.SetRenderState D3DRS_DITHERENABLE, False
'        DDevice.SetRenderState D3DRS_EDGEANTIALIAS, False
'
'        DDevice.SetRenderState D3DRS_INDEXVERTEXBLENDENABLE, False
'        DDevice.SetRenderState D3DRS_VERTEXBLEND, False
'
'        DDevice.SetRenderState D3DRS_CLIPPING, 1
'
'        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'
'        DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'        DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'        DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
'        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'        DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'        DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'        DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
'        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'
'        DDevice.SetRenderState D3DRS_ALPHAREF, Transparent
'        DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
'        DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESSEQUAL
'
'        DDevice.SetRenderState D3DRS_FOGENABLE, 1
'        DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
'        DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
'        DDevice.SetRenderState D3DRS_RANGEFOGENABLE, False
'        DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(FadeDistance / 4)
'        DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(FadeDistance)
'        DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(255, 184, 200, 225)
'
'        If frmMain.WindowState = vbMinimized Then frmMain.WindowState = IIf(FullScreen, vbMaximized, vbNormal)
'
'        On Error Resume Next
'
'        Set DSound = dx.DirectSoundCreate("")
'        DSound.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
'        If Err.Number <> 0 Then
'            Err.Clear
'            DisableSound = True
'        End If
'        On Error GoTo 0
'
'        Dim shArray() As Long
'        Dim shLength As Long
'        Dim shCode As D3DXBuffer
'
'        Set shCode = D3DX.AssembleShader("ps.1.0" & vbCrLf & _
'                                            "tex t0" & vbCrLf & _
'                                            "mul r0, t0,v0" & vbCrLf, 0, Nothing)
'        shLength = shCode.GetBufferSize() / 4
'        ReDim shArray(shLength - 1) As Long
'        D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
'        PixelShaderDefault = DDevice.CreatePixelShader(shArray(0))
'        Set shCode = Nothing
'
'        Set shCode = D3DX.AssembleShader("ps.1.1" & vbCrLf & _
'                                            "tex t0" & vbCrLf & _
'                                            "mov r0,t0" & vbCrLf, 0, Nothing)
'        shLength = shCode.GetBufferSize() / 4
'        ReDim shArray(shLength - 1) As Long
'        D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
'        PixelShaderDiffuse = DDevice.CreatePixelShader(shArray(0))
'        Set shCode = Nothing
'    End If
'End Sub

Public Sub DoPauseGame()
    PauseGame = True
    TermGameData
    TermDirectX
End Sub


Public Sub InitGameData()

    CreateInfo
    CreateCmds
    CreateText
    CreateMove
    CreateLand
    
End Sub

Public Sub TermGameData()

    CleanupLand
    CleanupMove
    CleanupText
    CleanupCmds
    CleanupInfo
    
End Sub

Public Sub TermDirectX()
    
    If Not DDevice Is Nothing Then
        DDevice.DeletePixelShader PixelShaderDiffuse
        DDevice.DeletePixelShader PixelShaderDefault
    End If

    Set DSound = Nothing
    Set DDevice = Nothing
    Set D3DX = Nothing
    Set D3D = Nothing
    Set dx = Nothing
    
End Sub

Private Function TestDirectX() As Boolean

    On Error Resume Next
    InitDirectX
    TestDirectX = (Err.Number = 0)
    If Err.Number Then Err.Clear
    On Error GoTo 0

End Function

