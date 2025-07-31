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
Public AmbientFX As Boolean

Public DebugMode As Boolean

Public Surface As Boolean
Public AspectRatio As Single
Public TrapMouse As Boolean
Public PauseGame As Boolean
Public StopGame As Boolean
Public ShowHelp As Boolean
Public ShowStat As Boolean
Public ShowCredits As Boolean
Public CameraClip As Boolean

Public FPSTimer As Double
Public FPSCount As Long
Public FPSRate As Long

Public db As Database
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

Private elapsed As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


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
    
    Set db = New Database
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

        
        db.rsClose rs
                
        FPSCount = 36
        TrapMouse = True
        ShowStat = False
        CameraClip = True
        
        JumpGUID = Replace(modGuid.GUID, "-", "")
        
        SetMotion GlobalGravityDirect, Actions.Directing, MakePoint(0, -0.2, 0), 1
        SetMotion GlobalGravityRotate, Actions.Rotating, MakePoint(0, 0, 0), 0
        SetMotion GlobalGravityScaled, Actions.Scaling, MakePoint(0, 0, 0), 0
        
        SetMotion LiquidGravityDirect, Actions.Directing, MakePoint(0, -0.005, 0), 2
        SetMotion LiquidGravityRotate, Actions.Rotating, MakePoint(0, 0, 0), 0
        SetMotion LiquidGravityScaled, Actions.Scaling, MakePoint(0, 0, 0), 0

        Load frmMain
                
        frmMain.BackColor = &H323232

        
        frmMain.Width = CSng(NextArg(Resolution, "x")) * Screen.TwipsPerPixelY
        frmMain.Height = CSng(RemoveArg(Resolution, "x")) * Screen.TwipsPerPixelX
        AspectRatio = CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))
        


        WorkingScreen "Loading..."
    
    
        On Error GoTo fault
        InitDirectX
        InitGameData
        On Error GoTo 0
        frmMain.AutoRedraw = False
        


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
                        WorkingScreen "Resuming..."
                        InitGameData
                        frmMain.AutoRedraw = False
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
    
                'elapsed = GetTickCount
                SetupWorld
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
                
                'elapsed = GetTickCount
                RenderMotion
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderMotion: " & elapsed
                
                'elapsed = GetTickCount
                RenderPlanes
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderPlanes: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderWorld
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderPlayer
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderBoards
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderLucent
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderBeacons
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderPortals
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderPortals: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderCameras
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderCameras: " & elapsed
                
                
                'elapsed = GetTickCount
                RenderRoutine
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderRoutine: " & elapsed
                
                On Error GoTo 0
        
                'elapsed = GetTickCount
                InputMove
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "InputMove: " & elapsed
                
                
                'elapsed = GetTickCount
                ResetMotion
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "ResetMotion: " & elapsed
                
                
                'elapsed = GetTickCount
                InputScene
                'elapsed = (GetTickCount - elapsed)
                'If elapsed > 0 Then Debug.Print "InputScene: " & elapsed
                
                

                If Not PauseGame Then
                    
                    'elapsed = GetTickCount
                    RenderInfo
                    'elapsed = (GetTickCount - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderInfo: " & elapsed
                    
                    
                    'elapsed = GetTickCount
                    RenderCmds
                    'elapsed = (GetTickCount - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderCmds: " & elapsed
   
   
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
            
'            If D3DWindow.Windowed Then
'                Static skipframes As Integer
'                skipframes = skipframes + 1
'                If skipframes >= 5 Then
'                    DoTasks
'                    skiframes = 0
'                End If
'            End If
            If D3DWindow.Windowed And FPSCount >= FPSRate / 3 Then DoTasks
        
        Loop
        WorkingScreen "Exiting..."
       
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

Public Sub WorkingScreen(ByVal Text As String)
    frmMain.AutoRedraw = True
    frmMain.Cls
    frmMain.CurrentY = Screen.TwipsPerPixelY * 5
    frmMain.CurrentX = Screen.TwipsPerPixelX * 5
    frmMain.Print Text
    frmMain.Show
    DoEvents
End Sub
Public Sub RenderRoutine()

    
    If Millis <> 0 Then
        If Timer - Millis >= 0.1 Then
            Millis = Timer
            frmMain.Run "Millis"
        End If
    End If

    If Second <> 0 Then
        If Timer - Second >= 1 Then
            Second = Timer
            frmMain.Run "Second"
        End If
    End If

    If CheckIdle(60) Then
        ResetIdle
        frmMain.Run "OnIdle"
    End If

    If Frame Then
        
        frmMain.Run "Frame"

    End If
    
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

    
    D3DXMatrixIdentity matWorld

    DDevice.SetTransform D3DTS_WORLD, matWorld
    'DDevice.SetTransform D3DTS_WORLD1, matWorld
    
    D3DXMatrixMultiply matTemp, matWorld, matWorld
    D3DXMatrixRotationY matRotation, 0.5
    D3DXMatrixRotationX matPitch, 0.5
    
    
    D3DXMatrixIdentity matWorld
    D3DXMatrixMultiply matLook, matRotation, matPitch
    
    DDevice.SetTransform D3DTS_WORLD, matWorld
  
    If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex > 0 And Player.CameraIndex <= Cameras.Count)) Or (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0)) Then
        
        D3DXMatrixRotationY matRotation, Cameras(Player.CameraIndex).Angle
        D3DXMatrixRotationX matPitch, Cameras(Player.CameraIndex).Pitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        D3DXMatrixTranslation matPos, -Cameras(Player.CameraIndex).Origin.X, -Cameras(Player.CameraIndex).Origin.Y, -Cameras(Player.CameraIndex).Origin.Z
        D3DXMatrixMultiply matLook, matPos, matLook
        D3DXMatrixTranslation matPos, -Player.Origin.X, -Player.Origin.Y + 0.2, -Player.Origin.Z
        
    Else
    
        D3DXMatrixRotationY matRotation, Player.Angle
        D3DXMatrixRotationX matPitch, Player.Pitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        If Player.Pitch > 0 Then
        
            D3DXMatrixTranslation matPos, -Player.Origin.X, -Player.Origin.Y, -Player.Origin.Z
            D3DXMatrixMultiply matLook, matPos, matLook
        Else
            D3DXMatrixTranslation matPos, -Player.Origin.X, -Player.Origin.Y + 0.2, -Player.Origin.Z
            D3DXMatrixMultiply matLook, matPos, matLook
        
        End If
        
    End If
    
    lCulledFaces = 0
    lCullCalls = 0
    
    If ((Perspective = Playmode.ThirdPerson) Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0))) Then
    
        If (CameraClip Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not ((Perspective = Spectator) Or DebugMode)) Then

            If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0)) Then

                Player.Twists.Y = 3
            
            End If
        
            Dim cnt As Long
            Dim cnt2 As Long

            Dim Face As Long
            Dim Zoom As Single
            Dim factor As Single
            Dim e1 As Element

            Dim verts(0 To 2) As D3DVECTOR
            Dim touched As Boolean
            
            'initialie sngFaceVis for camera collision checking
            For cnt = 1 To lngFaceCount - 1
                           ' On Error GoTo isdivcheck0
                                sngFaceVis(3, cnt) = 0
                             '   GoTo notdivcheck0
'isdivcheck0:
                                'If Err.Number = 11 Then Resume
'notdivcheck0:
                               ' If Err Then Err.Clear
                              '  On Error GoTo 0
                
            Next


            'commence the camera clip collision checking, this is what keeps
            'the camera from being inside of the level seeing out backfaces
            
            Zoom = 0.2
            factor = 0.5

            Do

                verts(0) = MakeVector(Player.Origin.X, _
                                            Player.Origin.Y - 0.2, _
                                            Player.Origin.Z)

                verts(1) = MakeVector(Player.Origin.X - (Sin(D720 - Player.Twists.Y) * (Zoom + factor)), _
                                            Player.Origin.Y - 0.2 + (Tan(D720 - Player.Pitch) * (Zoom + factor)), _
                                            Player.Origin.Z - (Cos(D720 - Player.Twists.Y) * (Zoom + factor)))

                verts(2) = MakeVector(Player.Origin.X - (Sin(D720 - Player.Twists.Y)), _
                                      Player.Origin.Y - 0.1 + (Tan(D720 - Player.Pitch) * Zoom), _
                                      Player.Origin.Z - (Cos(D720 - Player.Twists.Y)))

                sngCamera(0, 0) = Player.Origin.X
                sngCamera(0, 1) = Player.Origin.Y
                sngCamera(0, 2) = Player.Origin.Z

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

                If (Elements.Count > 0) Then
                    For Each e1 In Elements
                    
                    
                    'For cnt = 1 To Elements.Count
                        If ((Not (e1.Effect = Collides.Ground)) And (Not (e1.Effect = Collides.InDoor))) And (e1.CollideIndex > -1) And (e1.BoundsIndex > 0) Then
                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                      '      On Error GoTo isdivcheck1
                                sngFaceVis(3, cnt2) = 0
                           '     GoTo notdivcheck1
'isdivcheck1:
                             '   If Err.Number = 11 Then Resume
'notdivcheck1:
                            '    If Err Then Err.Clear
                           '     On Error GoTo 0
                                
                            Next
                        ElseIf (e1.Effect = Collides.Ground) And (e1.CollideIndex > -1) And e1.Visible And (e1.BoundsIndex > 0) Then
                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                                If Not (((sngFaceVis(0, cnt2) = 0) Or (sngFaceVis(0, cnt2) = 1) Or (sngFaceVis(0, cnt2) = -1)) And _
                                    ((sngFaceVis(1, cnt2) = 0) Or (sngFaceVis(1, cnt2) = 1) Or (sngFaceVis(1, cnt2) = -1)) And _
                                    ((sngFaceVis(2, cnt2) = 0) Or (sngFaceVis(2, cnt2) = 1) Or (sngFaceVis(2, cnt2) = -1))) Then
                                    sngFaceVis(3, cnt2) = 2
                                End If
                            Next
                        End If
                    Next
                    If (Player.CollideIndex > -1) And (Player.BoundsIndex > 0) And (Player.BoundsIndex > 0) Then
                        For cnt2 = Player.CollideIndex To (Player.CollideIndex + Meshes(Player.BoundsIndex).Mesh.GetNumFaces) - 1
                          '  On Error GoTo isdivcheck2
                                sngFaceVis(3, cnt2) = 0
                                'GoTo notdivcheck2
'isdivcheck2:
   '                             If Err.Number = 11 Then Resume
'notdivcheck2:
                          '      If Err Then Err.Clear
                          '      On Error GoTo 0

                        Next
                    End If
                End If

                Face = AddCollisionEx(verts, 1)
                touched = TestCollisionEx(Face, 2)
                DelCollisionEx Face, 1

                If ((Not touched) And (Zoom < Player.Zoom)) Then Zoom = Zoom + factor

            Loop Until ((touched) Or (Zoom >= Player.Zoom))

            If (touched And (Zoom > 0.2)) Then Zoom = Zoom + -factor

            D3DXMatrixTranslation matTemp, 0, 0.2, Zoom
            D3DXMatrixMultiply matView, matLook, matTemp


            'all said and done, if the zoom is under a certian val the
            'toon is in the way, so change it to wireframe see through
            Player.WireFrame = (Zoom < 0.8)

        Else
        
            D3DXMatrixTranslation matTemp, 0, 0, IIf(Not ((Perspective = Spectator) Or DebugMode), Player.Zoom, 0)
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
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Resume
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
'        DSound.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY
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

