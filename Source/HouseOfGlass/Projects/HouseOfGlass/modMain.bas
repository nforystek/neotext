Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type TVERTEX0
    X As Single
    Y As Single
    Z As Single
    tu As Single
    tv As Single
End Type

Public Type TVERTEX1
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public Type TVERTEX2
    X    As Single
    Y    As Single
    Z    As Single
    nx   As Single
    ny   As Single
    nz   As Single
    tu   As Single
    tv   As Single
End Type

Public Type UserType
    Location As D3DVECTOR
    CameraAngle As Single
    CameraPitch As Single
    CameraZoom As Single
    MoveSpeed As Single
    AutoMove As Boolean
End Type

Public Resolution As String
Public FullScreen As Boolean
Public WireFrame As Boolean
Public PlayMusic As Boolean

Public MenuMode As Long
Public TrapMouse As Boolean
Public PauseGame As Boolean
Public StopGame As Boolean

Public Player As UserType

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

Public Const FVF_VTEXT0 = D3DFVF_NORMAL Or D3DFVF_TEX1 Or D3DFVF_XYZ
Public Const FVF_VTEXT1 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_VTEXT2 = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Const PI As Single = 3.14159265359
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2

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
    ElseIf Trim(LCase(inCmd)) = " " Then
        frmEdit.Show
    Else
        
        Set db = New clsDatabase
        db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
        If db.rsEnd(rs) Then inCmd = "setup"
        
        If Trim(LCase(inCmd)) = "setup" Then
            frmSetup.Show
            Do While frmSetup.Visible
               DoEvents
            Loop
            If frmSetup.Play Then inCmd = ""
            Unload frmSetup
        End If
        
        
        If inCmd = "" Then
        
            db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
            
            If Not db.rsEnd(rs) Then
                Resolution = CStr(rs("Resolution"))
                FullScreen = CBool(Not rs("Windowed"))
                WireFrame = CBool(rs("WireFrame"))
            End If
            
            db.rsClose rs
            
            FPSCount = 36
            MenuMode = -1
            TrapMouse = True
            Player.MoveSpeed = 3
        
            Load frmMain
            
            frmMain.width = CSng(NextArg(Resolution, "x")) * Screen.TwipsPerPixelY
            frmMain.height = CSng(RemoveArg(Resolution, "x")) * Screen.TwipsPerPixelX
            
            On Error GoTo fault
            InitDirectX
            InitGameData
            On Error GoTo 0
            
            If Not DisableSound Then
                If PlayMusic Then
                    Track1.TrackVolume = 1000
                    Track2.TrackVolume = 0
                    Track1.PlaySound
                    Track2.PlaySound
                End If
            End If
            
            frmMain.Show
        
            InitialCommands
            AtLvl = AtLvl - 1
            MenuMode = 1
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
                    
                    DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
                    DDevice.BeginScene
        
                    SetupWorld
                    RenderLand
                    RenderLawn
                    RenderAudio
                    
                    On Error GoTo 0
                    
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
    Resume
    Err.Clear
    End
End Sub

Public Sub SetupWorld()

    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matTemp As D3DMATRIX

    D3DXMatrixRotationY matRotation, Player.CameraAngle
    D3DXMatrixRotationX matPitch, Player.CameraPitch
    D3DXMatrixMultiply matLook, matRotation, matPitch

    D3DXMatrixTranslation matPos, -Player.Location.X, -Player.Location.Y - (Player.Location.Y / 2), -Player.Location.Z
    D3DXMatrixMultiply matLook, matPos, matLook

    DDevice.SetTransform D3DTS_VIEW, matLook

    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 0.75, 5, 50000
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
    End If
        
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
    
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
    
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP

    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
  
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    DDevice.SetRenderState D3DRS_ALPHAREF, Transparent
    DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
                            
    DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESSEQUAL
        
    DDevice.SetRenderState D3DRS_FOGENABLE, 1
    DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
    DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
    DDevice.SetRenderState D3DRS_RANGEFOGENABLE, False
    DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(FadeDistance / 8)
    DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(FadeDistance)
    DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(255, 184, 200, 225)
    
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = IIf(FullScreen, vbMaximized, vbNormal)

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

Private Sub InitGameData()

    CreateCmds
    CreateInfo
    CreateText
    CreateLand
    CreateLawn
    CreateSounds

End Sub

Private Sub TermGameData()

    CleanupSounds
    CleanupLawn
    CleanupLand
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

