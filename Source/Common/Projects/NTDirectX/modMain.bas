Attribute VB_Name = "modMain"
#Const DxVBLibA = -1

#Const modMain = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public Resolution As String
Public WireFrame As Boolean
Public FullScreen As Boolean
Public SilentMode As Boolean
Public PauseGame As Boolean
Public ScreenSaver As Boolean
Public TrapMouse As Boolean
Public StopGame As Boolean
Public ShowSetup As Boolean
Public ResetGame As Boolean
Public ShowStats As Boolean

Public FPSTimer As Double
Public FPSCount As Long
Public FPSRate As Long
Public FPSElapse As Single
Public FPSLatency As Single

Public BackColor As Long

Public matWorld As D3DMATRIX
Public matView As D3DMATRIX
Public matProj As D3DMATRIX
    
Public dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public DDevice As Direct3DDevice8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public Display As D3DDISPLAYMODE
Public DSound As DirectSound8

Public DViewPort As D3DVIEWPORT8
Public DSurface As D3DXRenderToSurface

Public PixelShaderDefault As Long
Public PixelShaderDiffuse As Long

Public Sub ShowSetupForm(ByRef UserControl As Macroscopic)

    UserControl.PauseRendering
    frmSetup.Left = UserControl.Parent.Left + ((UserControl.Parent.Width / 2) - (frmSetup.Width / 2))
    frmSetup.Top = UserControl.Parent.Top + ((UserControl.Parent.Height / 2) - (frmSetup.Height / 2))
    
    frmSetup.Show
        
    Do While frmSetup.Visible
       DoTasks
    Loop

    Unload frmSetup
    
    UserControl.ResumeRendering
End Sub

Public Sub RenderFrame(ByRef UserControl As Macroscopic)

    Do While Not StopGame

        If PauseGame Then

            If Not ((frmMain.WindowState = 1) Or (UserControl.Parent.WindowState = 1)) Then
                If TestDirectX(UserControl) Then
    
                    On Error Resume Next
                    DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET + D3DCLEAR_ZBUFFER, Camera.Color.RGBA, 1, 0
                    DDevice.BeginScene
                    DDevice.EndScene

                    DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                    
                    If (Err.number = 0) Then
                        PauseGame = False
                    Else
                        Err.Clear
                    End If
                    On Error GoTo 0
    
                End If
                
                If (Not PauseGame) And (Err.number = 0) Then
                    InitGameData UserControl
                Else
                    TermDirectX UserControl
                End If
            ElseIf Not D3DWindow.Windowed = 1 Then
                DoTasks
            End If
            
           ' If GetKeyState(VK_F1) = -128 Then ShowSetup UserControl

        Else

            On Error GoTo nofocus

            'BeginMirrors UserControl, Camera.Player
            
            DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET + D3DCLEAR_ZBUFFER, Camera.Color.ARGB, 1, 0    ' D3DCLEAR_ZBUFFER, Camera.Color.ARGB, 1, 0

            On Error GoTo 0 'temporary
     

            RenderCamera UserControl, Camera
            
            DDevice.BeginScene

            
            RenderEvents UserControl, Camera
         '   RenderMotions UserControl, Camera
            
            RenderBrilliants UserControl, Camera


                        
            RenderPlanets UserControl, Camera

            
            RenderMolecules UserControl, Camera


            InputScene UserControl


            
            If Not PauseGame Then
                
                RenderCmds UserControl
            
                On Error GoTo nofocus
                
                DDevice.EndScene

                PresentScene UserControl
                
                On Error GoTo 0 'temporary

                FPSCount = FPSCount + 1
                If (FPSTimer = 0) Or ((Timer - FPSTimer) >= 1) Then
                    FPSTimer = Timer
                    FPSRate = FPSCount
                    FPSCount = 0
                End If
                FPSElapse = Round(((Timer - FPSLatency) + FPSElapse) / 2, 4)
                FPSLatency = Timer
                    
            End If

            If ShowSetup Then
                ShowSetupForm UserControl
                ShowSetup = False
            End If
            If ResetGame Then
                UserControl.PauseRendering
                If PathExists(ScriptRoot & "\Serial.xml", True) Then
                    Kill ScriptRoot & "\Serial.xml"
                End If
                UserControl.ResumeRendering
                If ConsoleVisible Then ConsoleToggle
                ResetGame = False
            End If

        End If

        If D3DWindow.Windowed = 1 Then DoTasks
    Loop

Exit Sub
nofocus:

    Err.Clear
   ' DoPauseGame UserControl
    UserControl.PauseRendering

'    If Err.number <> 0 Then
'        Err.Clear
'        If IsScreenSaverActive And (Not ScreenSaver) Then
'            ScreenSaver = True
'            UserControl.PauseRendering
'            TrapMouse = False
'        End If
'        DoTasks
'    End If
        
    'UserControl.PauseRendering
    'DoPauseGame UserControl
    
'Exit Sub
'Render:
'    TermGameData UserControl
'    TermDirectX UserControl
'    StopGame = True
'
'    Unload frmMain
'
'    MsgBox "There was an error trying to run the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.number & ", " & Err.description, vbOKOnly + vbInformation, App.Title
'    Err.Clear
    
End Sub

Public Sub DoPauseGame(ByRef UserControl As Macroscopic)

    On Error Resume Next
    PauseGame = True
    TermGameData UserControl
    TermDirectX UserControl
End Sub

Public Sub RenderEvents(ByRef UserControl As Macroscopic, ByRef Camera As Camera)


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

    If Frame Then
        
        frmMain.Run "Frame"

    End If
    
    
    
    Dim ev As OnEvent
    
    For Each ev In OnEvents
        Select Case ev.EventType
            Case EventTypes.Ranged
                        
            
            Case EventTypes.Contact
                        
            
        End Select
    Next
        
    
End Sub


Public Sub PresentScene(ByRef UserControl As Macroscopic)

    On Error Resume Next

    Do
        If Err.number <> 0 Then
            Err.Clear
            If IsScreenSaverActive And (Not ScreenSaver) Then
                ScreenSaver = True
                UserControl.PauseRendering
                TrapMouse = False
            ElseIf (Not IsScreenSaverActive) And ScreenSaver Then
                UserControl.ResumeRendering
                ScreenSaver = False
            End If
            DoTasks
        End If

        DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    Loop Until Err.number = 0

    On Error GoTo 0
End Sub

Public Sub RenderCamera(ByRef UserControl As Macroscopic, ByRef Camera As Camera)

    Dim matYaw As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matRoll As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matRot As D3DMATRIX

    DDevice.SetTransform D3DTS_WORLD, matWorld
    D3DXMatrixIdentity matWorld

    D3DXMatrixIdentity matView
'    DDevice.SetTransform D3DTS_VIEW, matView

        
    If Not Camera.Player Is Nothing Then

        If Not Camera.Planet Is Nothing Then

            D3DXMatrixRotationX matPitch, AngleRestrict(-Camera.Player.Absolute.Rotate.X)
            D3DXMatrixMultiply matView, matPitch, matView

            D3DXMatrixRotationY matYaw, AngleRestrict(-Camera.Player.Absolute.Rotate.Y)
            D3DXMatrixMultiply matView, matYaw, matView

            D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3D(-Camera.Player.Absolute.Rotate.Z)
            D3DXMatrixMultiply matView, matRoll, matView

            DDevice.SetTransform D3DTS_VIEW, matView
            
            D3DXMatrixTranslation matPos, -Camera.Player.Absolute.Origin.X + Camera.Planet.Absolute.Origin.X, -Camera.Player.Absolute.Origin.Y + Camera.Planet.Absolute.Origin.Y, -Camera.Player.Absolute.Origin.Z + Camera.Planet.Absolute.Origin.Z
            D3DXMatrixMultiply matView, matPos, matView

            DDevice.SetTransform D3DTS_VIEW, matView

            D3DXMatrixRotationX matPitch, AngleRestrict(-Camera.Planet.Rotate.X)
            D3DXMatrixMultiply matView, matPitch, matView

            D3DXMatrixRotationY matYaw, AngleRestrict(-Camera.Planet.Rotate.Y)
            D3DXMatrixMultiply matView, matYaw, matView

            D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3D(-Camera.Planet.Rotate.Z)
            D3DXMatrixMultiply matView, matRoll, matView

            DDevice.SetTransform D3DTS_VIEW, matView

       Else

            D3DXMatrixRotationX matPitch, AngleRestrict(-Camera.Player.Rotate.X)
            D3DXMatrixMultiply matView, matPitch, matView

            D3DXMatrixRotationY matYaw, AngleRestrict(-Camera.Player.Rotate.Y)
            D3DXMatrixMultiply matView, matYaw, matView

            D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3D(-Camera.Player.Rotate.Z)
            D3DXMatrixMultiply matView, matRoll, matView

            DDevice.SetTransform D3DTS_VIEW, matView

            D3DXMatrixTranslation matPos, -Camera.Player.Origin.X, -Camera.Player.Origin.Y, -Camera.Player.Origin.Z
            D3DXMatrixMultiply matView, matPos, matView

            DDevice.SetTransform D3DTS_VIEW, matView

        End If

    Else
        D3DXMatrixRotationX matPitch, 0
        D3DXMatrixMultiply matView, matPitch, matView
    
        D3DXMatrixRotationY matYaw, 0
        D3DXMatrixMultiply matView, matYaw, matView
        
        D3DXMatrixRotationZ matRoll, 0
        D3DXMatrixMultiply matView, matRoll, matView
        
        D3DXMatrixTranslation matPos, 0, 0, 0
        D3DXMatrixMultiply matView, matPos, matView
        
        DDevice.SetTransform D3DTS_VIEW, matView
    End If
    
    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), Near, Far
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    
End Sub

Public Sub InitDirectX(ByRef UserControl As Macroscopic)

    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    Set D3DX = New D3DX8
        
    If FullScreen Then
        InitialDevice UserControl, frmMain.hwnd
    Else
        InitialDevice UserControl, frmMain.Picture1.hwnd
    End If
        
End Sub

Private Sub InitialDevice(ByRef UserControl As Macroscopic, ByVal hwnd As Long)
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display

    D3DWindow.BackBufferCount = 1
    
    D3DWindow.BackBufferWidth = (VB.Screen.Width / VB.Screen.TwipsPerPixelX)
    D3DWindow.BackBufferHeight = (VB.Screen.Height / VB.Screen.TwipsPerPixelY)
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
    D3DWindow.EnableAutoDepthStencil = 1
    
    DViewPort.MaxZ = Far
    DViewPort.MinZ = Near
    DViewPort.Width = (VB.Screen.Width / VB.Screen.TwipsPerPixelX)
    DViewPort.Height = (VB.Screen.Height / VB.Screen.TwipsPerPixelY)
    
    On Error Resume Next
'    D3DFMT_X8R8G8B8 , D3DUSAGE_DEPTHSTENCIL, _
'        D3DRTYPE_SURFACE, D3DFMT_D16
    Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
    If Err.number <> 0 Then
        Err.Clear
        Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
        If Err.number <> 0 Then
            Err.Clear
            Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
            If Err.number <> 0 Then
               ' MsgBox "Error: " & Err.description
            Err.Clear
            End If
        End If
    End If
    On Error GoTo 0
    
    If Not DDevice Is Nothing Then

         DDevice.GetViewport DViewPort
         
         'DViewPort.X = (((Screen.Width / VB.Screen.TwipsPerPixelX) / 2) - 256)
         'DViewPort.Width = DViewPort.Width - (DViewPort.X * 2)

         'DViewPort.Y = (((Screen.Height / VB.Screen.TwipsPerPixelY) / 2) - 256)
         'DViewPort.Height = DViewPort.Height - (DViewPort.Y * 2)

        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_LIGHTING, 1
        DDevice.SetRenderState D3DRS_DITHERENABLE, 0
        DDevice.SetRenderState D3DRS_EDGEANTIALIAS, 0

        DDevice.SetRenderState D3DRS_INDEXVERTEXBLENDENABLE, 0
        DDevice.SetRenderState D3DRS_VERTEXBLEND, 0

        DDevice.SetRenderState D3DRS_CLIPPING, 1
       ' DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, 1

        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID

        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

        DDevice.SetRenderState D3DRS_ALPHAREF, Transparent
        DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
        DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESSEQUAL
 
'        DDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD

'        DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
'        DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
'        DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
'        DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
 
 
        DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

'        DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_DISABLE
'        DDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
'        DDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
'        DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_DIFFUSE

        DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
        DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

        DDevice.SetRenderState D3DRS_FOGENABLE, 0
        DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
        DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
        DDevice.SetRenderState D3DRS_RANGEFOGENABLE, 0
        DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(Far / 4)
        DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(Far)
        DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(255, 184, 200, 225)

        If frmMain.WindowState = vbMinimized Then frmMain.WindowState = IIf(FullScreen, vbMaximized, vbNormal)

        On Error Resume Next
        Set DSound = dx.DirectSoundCreate("")
        DSound.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
        If Err.number <> 0 Then Err.Clear
        On Error GoTo 0


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


        GravityVector.Y = -0.05
        
        LiquidVector.Y = -0.005
        
        Set DSurface = D3DX.CreateRenderToSurface(DDevice, VB.Screen.Width / VB.Screen.TwipsPerPixelX, VB.Screen.Height / VB.Screen.TwipsPerPixelY, Display.Format, 1, D3DFMT_D16)
        
    End If
    
End Sub

Public Sub InitGameData(ByRef UserControl As Macroscopic)

    CreateCmds
    CreateText
    CreateProj
    CreateObjs


End Sub

Public Sub TermGameData(ByRef UserControl As Macroscopic)
    
    CleanUpObjs
    CleanUpProj
    CleanupText
    CleanupCmds

End Sub

Public Sub TermDirectX(ByRef UserControl As Macroscopic)
    PauseGame = True
    
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

Public Function TestDirectX(ByRef UserControl As Macroscopic) As Boolean

    On Error Resume Next
    InitDirectX UserControl
    TestDirectX = (Err.number = 0)
    If Err.number <> 0 Then Err.Clear
    On Error GoTo 0

End Function







