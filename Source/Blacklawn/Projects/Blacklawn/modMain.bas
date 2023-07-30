Attribute VB_Name = "modMain"

#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public ViewIntro As Boolean
Public Resolution As String
Public PlaySound As Boolean
Public PlayMusic As Boolean
Public FullScreen As Boolean
Public Version2D As Boolean
Public WireFrame As Boolean
Public ClearScore As Boolean
Public AspectRatio As Single

Public TrapMouse As Boolean
Public PauseGame As Boolean
Public StopGame As Boolean

Public GodMode As Boolean
Public Player As MyPlayer

Public UserData As String
Public WarpData As String
Public ViewData As String

Public DownTime As String
Public StartNow As String

Public FPSTimer As Single
Public FPSCount As Long
Public FPSRate As Long

Public DISTMoved As D3DVECTOR
Public DISTLimit As Long
Public DISTSkips As Long

Public db As clsDatabase
Public rs As ADODB.Recordset

Public dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public DDevice As Direct3DDevice8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public Display As D3DDISPLAYMODE
Public DSound As DirectSound8

Public Sub Main()
    On Error GoTo fault:

    Dim inCmd As String
    inCmd = Command
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    If Trim(LCase(inCmd)) = "setupreset" Then
        ResetDB
        CompactDB
    ElseIf Trim(LCase(inCmd)) = "setupbackup" Then
        BackupDB
    ElseIf Trim(LCase(inCmd)) = "setuprestore" Then
        RestoreDB
    ElseIf Trim(LCase(inCmd)) = "shipeditor" Then
        frmEdit.Show
    Else
        Load frmSplash
        
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
        
        If (Trim(LCase(inCmd)) = "") And (Not Version2D) Then
            frmSplash.Show
            DoTasks
            frmSplash.Picture1.SetFocus
        End If
        
        db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"

        If Not db.rsEnd(rs) Then
            ViewIntro = CBool(rs("ViewIntro"))
            Resolution = CStr(rs("Resolution"))
            FullScreen = CBool(Not rs("Windowed"))
            Version2D = CBool(rs("Version2D"))
            PlaySound = CBool(rs("SoundEnabled"))
            PlayMusic = CBool(rs("MusicEnabled"))
            WireFrame = CBool(rs("WireFrame"))
            ClearScore = CBool(rs("ClearScore"))
        End If
        
        db.rsClose rs
    
        TrapMouse = True
        
        If (Trim(LCase(inCmd)) = "") And (Not Version2D) Then
            Randomize
            
            LoadUserData
            Load frmMain
            
            frmMain.width = CSng(NextArg(Resolution, "x")) * Screen.TwipsPerPixelY
            frmMain.height = CSng(RemoveArg(Resolution, "x")) * Screen.TwipsPerPixelX
            AspectRatio = CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))
            
            If ViewIntro And (Not frmSplash.Skip) Then
                If PlayMusic Then
                    CreateTracks
                    Track1.TrackVolume = 0
                    Track2.TrackVolume = 1000
                    Track1.PlaySound
                    Track2.PlaySound
                End If
                frmSplash.PlayIntro
            Else
                If PlayMusic Then
                    CreateTracks
                    Track2.TrackVolume = 0
                    Track1.TrackVolume = 1000
                    Track1.PlaySound
                    Track2.PlaySound
                End If
            End If
                
            On Error GoTo fault
            InitDirectX
            InitGameData
            On Error GoTo 0
            
            frmMain.Show
            
            If PlaySound Then PlayWave SOUND_ROLLOUT
            
            InitialCommands
            
            Unload frmSplash
            
            Do While Not StopGame
                
                If PauseGame Then
                    
                    If Not (frmMain.WindowState = 1) Then
                        If TestDirectX Then
                        
                            On Error Resume Next
                            DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
                            DDevice.BeginScene
                            DDevice.EndScene
                            DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                            If Not Err.Number Then
                                PauseGame = False
                            Else
                                Err.Clear
                            End If
                            On Error GoTo 0
                                
                        End If
                        
                        If Not PauseGame Then
                            InitGameData
                        Else
                            TermDirectX
                        End If
                    Else
                        DoTasks
                    End If
                    
                Else
                        
                    On Error GoTo Render
                    DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 2, 0
                    DDevice.BeginScene

                    SetupWorld
                    
                    SetupActivity
                    RenderLights
                    RenderGalaxy
                    
                    RenderLawn
                    
                    RenderPlane
                    RenderBoards
                    
                    RenderShips
                    
                    RenderBeacons
                    RenderPowerup
                    RenderGlass
                    RenderAudio
                    
                    On Error GoTo 0

                    If FPSCount > 0 And DISTSkips > 0 Then

                        If DISTSkips <= 0 Then
                            RenderActivity
                            InputScene
                            ScoreUser
                        End If

                    Else
                        RenderActivity
                        InputScene
                        ScoreUser
                    End If

                    DISTSkips = (Abs(Distance(Player.Object.Origin.X, Player.Object.Origin.Y, _
                            Player.Object.Origin.z, DISTMoved.X, DISTMoved.Y, DISTMoved.z)) / Player.MoveSpeed)
                    If DISTSkips > 250 Then
                        DISTLimit = DISTLimit + 1
                    ElseIf DISTSkips > 0 Then
                        DISTLimit = DISTLimit - 1
                    End If
                    If DISTLimit < 0 Then DISTLimit = 0
                    DISTMoved = Player.Object.Origin
                    
                    If (Not PauseGame) Then
                        
                        RenderInfo
                        RenderCmds
                        
                        DDevice.EndScene
                        On Error Resume Next
                        DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                        
                        FPSCount = FPSCount + 1
                        If (FPSTimer = 0) Or ((Timer - FPSTimer) >= 1) Then
                            FPSRate = FPSCount
                            FPSTimer = Timer
                            FPSCount = 0
                            DISTSkips = Round(Abs(DISTLimit / FPSRate), 3)
                        End If
                        
                        If Err.Number Then
                            Err.Clear
                            DoPauseGame
                        End If
                        On Error GoTo 0
                        
                    End If
                    
                End If
                
                If D3DWindow.Windowed Then DoEvents
            
            Loop
            
            TermGameData
            TermDirectX
            
            SaveUserData
            Unload frmMain
            
            db.rsClose rs
            Set db = Nothing
            End
        
        ElseIf (Trim(LCase(inCmd)) = "") And Version2D Then
            frmLawn.Show
        Else
            db.rsClose rs
            Set db = Nothing
            End
        End If
        
    End If
    
Exit Sub
fault:
    TermDirectX
    
    MsgBox "There was an error initializing the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.Number & ", " & Err.Description, vbOKOnly + vbInformation, App.Title
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

    Dim matWorld As D3DMATRIX
    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matOffset As D3DMATRIX

    DDevice.SetTransform D3DTS_WORLD, matWorld

    D3DXMatrixRotationY matRotation, Player.CameraAngle
    D3DXMatrixRotationX matPitch, Player.CameraPitch
    D3DXMatrixMultiply matLook, matRotation, matPitch

    D3DXMatrixTranslation matPos, -Player.Object.Origin.X, -Player.Object.Origin.Y, -Player.Object.Origin.z
    D3DXMatrixMultiply matLook, matPos, matLook
        
    D3DXMatrixTranslation matOffset, 0, -20, Player.CameraZoom
    D3DXMatrixMultiply matView, matLook, matOffset
    DDevice.SetTransform D3DTS_VIEW, matView

    D3DXMatrixPerspectiveFovLH matProj, PI / 4, AspectRatio, 30, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub

Private Sub InitDirectX()

    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    Set D3DX = New D3DX8

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display

    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferWidth = CDbl(NextArg(Resolution, "x"))
    D3DWindow.BackBufferHeight = CDbl(RemoveArg(Resolution, "x"))
    D3DWindow.BackBufferFormat = Display.Format
            
    If Not FullScreen Then
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    Else
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
        D3DWindow.FullScreen_RefreshRateInHz = 0
    End If
    
    D3DWindow.MultiSampleType = D3DMULTISAMPLE_NONE
    
    D3DWindow.hDeviceWindow = frmMain.hwnd
        
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    D3DWindow.EnableAutoDepthStencil = True

    On Error Resume Next
    Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
    If Err.Number <> 0 Then
        Err.Clear
        Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
        If Err.Number <> 0 Then
            Err.Clear
            Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
        End If
    End If
    On Error GoTo 0
        
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_DITHERENABLE, False
    DDevice.SetRenderState D3DRS_EDGEANTIALIAS, False
  
    DDevice.SetRenderState D3DRS_INDEXVERTEXBLENDENABLE, False
    DDevice.SetRenderState D3DRS_VERTEXBLEND, False
    
    DDevice.SetRenderState D3DRS_CLIPPING, 1
    
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
  
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    
    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    DDevice.SetRenderState D3DRS_ALPHAREF, Transparent
    DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
                            
    DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESSEQUAL
        
    DDevice.SetRenderState D3DRS_FOGENABLE, 1
    DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
    DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
    DDevice.SetRenderState D3DRS_RANGEFOGENABLE, False
    DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(FadeDistance / 2)
    DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(FadeDistance)

    If frmMain.WindowState = vbMinimized Then
        frmMain.WindowState = IIf(FullScreen, vbMaximized, vbNormal)
    End If

    On Error Resume Next
    Set DSound = dx.DirectSoundCreate("")
    DSound.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
    If Err.Number <> 0 Then
        Err.Clear
        DisableSound = True
    End If
    On Error GoTo 0
    
End Sub

Public Sub DoPauseGame()
    PauseGame = True
    TermGameData
    TermDirectX
End Sub

Public Sub InitGameData()

    CreateCmds
    CreateInfo
    CreateText
    CreateLawn
    CreateShips
    CreateTracks
    CreateSounds
    
End Sub

Public Sub TermGameData()

    CleanupSounds
    CleanupTracks
    CleanupShips
    CleanupLawn
    CleanupText
    CleanupInfo
    CleanupCmds

End Sub

Private Sub TermDirectX()
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
