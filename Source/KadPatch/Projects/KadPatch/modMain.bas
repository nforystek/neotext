Attribute VB_Name = "modMain"

#Const modMain = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public SettingsID As Long
Public Resolution As String
Public WireFrame As Boolean
Public NotFocused As Boolean
Public TrapMouse As Integer

Public Player As UserType

Public FPSTimer As Double
Public FPSCount As Long
Public FPSRate As Long

Public BackColor As Long

Public matWorld As D3DMATRIX
Public matView As D3DMATRIX

Public matRight As D3DMATRIX
Public matLeft As D3DMATRIX

Public matRot As D3DMATRIX
Public matScale As D3DMATRIX

Public db As clsDatabase

Public dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public DDevice As Direct3DDevice8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public Display As D3DDISPLAYMODE
Public DSound As DirectSound8

Public GenericMaterial As D3DMATERIAL8
Public LucentMaterial As D3DMATERIAL8

Public PixelShaderDefault As Long
Public PixelShaderDiffuse As Long


Public Sub Main()

    Dim inCmd As String
    inCmd = Command
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)

    Set db = New clsDatabase
    Dim rs As ADODB.Recordset
    
    If LCase(inCmd) = "setupreset" Then
        db.dbQuery "DELETE * FROM Materials;"
        db.dbQuery "DELETE * FROM Settings;"
        
        Dim files As String
        Dim file As String
        files = SearchPath("*.bmp", -1, AppPath & "Base\Stitchings\FlossThreads\", FindAll)
        Do Until files = ""
            file = RemoveNextArg(files, vbCrLf)
            If PathExists(file, True) Then Kill file
        Loop
        files = SearchPath("*.bmp", -1, AppPath & "Base\Stitchings\LegendKeys\", FindAll)
        Do Until files = ""
            file = RemoveNextArg(files, vbCrLf)
            If PathExists(file, True) Then Kill file
        Loop
        
        End
    ElseIf LCase(inCmd) = "setupbackup" Then
        db.rsQuery rs, "SELECT * FROM Materials;"
        Dim out As String
        If Not db.rsEnd(rs) Then
            Do
                out = out & rs("Color") & "," & rs("Symbol") & vbCrLf
                rs.MoveNext
            Loop Until db.rsEnd(rs)
        End If
        WriteFile AppPath & "KadPatch.bak", out
        
        End
        
    ElseIf LCase(inCmd) = "setuprestore" Then
        If PathExists(AppPath & "KadPatch.bak", True) Then
            Dim bak As String
            Dim line As String
            
            bak = ReadFile(AppPath & "KadPatch.bak")
            Do Until bak = ""
                line = RemoveNextArg(bak, vbCrLf)
                db.rsQuery rs, "SELECT * FROM Materials WHERE Color=" & NextArg(line, ",") & ";"
                
                If db.rsEnd(rs) Then
                    db.dbQuery "INSERT INTO Materials (Color, Symbol) VALUES (" & NextArg(line, ",") & ",'" & RemoveArg(line, ",") & "');"
                Else
                    db.dbQuery "UPDATE Materials SET Symbol='" & RemoveArg(line, ",") & "' WHERE Color=" & NextArg(line, ",") & ";"
                End If
    
            Loop
        
            Kill AppPath & "KadPatch.bak"
        End If
        
        End
    End If
    
    db.rsClose rs
    
    If inCmd = "" Or (InStr(inCmd, "demo") > 0) Then

        FPSCount = 36
        NotFocused = False
        Player.MoveSpeed = 6
        
        BackColor = ConvertColor(SystemColorConstants.vbButtonFace)

        Load frmSplash
        frmSplash.Show 1

        Load frmStudio
        frmStudio.Show

        Dim chk As Boolean
        
        Do Until Forms.count = 0
            RenderFrame
            DoTasks

        Loop
        
    End If

End Sub

Public Sub CleanupProjFiles()
    Dim rs As New ADODB.Recordset

    Dim id As Long

    Dim files As String
    Dim file As String
    files = SearchPath("*.bmp", -1, AppPath & "Base\Stitchings\LegendKeys", FindAll)
    Do Until files = ""
        file = RemoveNextArg(files, vbCrLf)
        id = Val("&" & GetFileTitle(file))
        db.rsQuery rs, "SELECT * FROM Materials WHERE ID = " & id & ";"
        If db.rsEnd(rs) Then Kill file
    Loop


    files = SearchPath("*.bmp", -1, AppPath & "Base\Stitchings\FlossThreads", FindAll)
    Do Until files = ""
        file = RemoveNextArg(files, vbCrLf)
        id = Val("&" & GetFileTitle(file))
        db.rsQuery rs, "SELECT * FROM Materials WHERE Color = " & id & ";"
        If db.rsEnd(rs) Then Kill file
    Loop


    db.rsClose rs
End Sub
Public Sub RenderFrame()
            
    If NotFocused Then
        
        If Not (frmMain.WindowState = 1) Then
            If TestDirectX Then

                On Error Resume Next
                DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1, 0
                DDevice.BeginScene
                DDevice.EndScene
                DDevice.Present ByVal 0, ByVal 0, frmMain.Picture1.hwnd, ByVal 0
                If (Err.number = 0) And (GetActiveWindow = frmStudio.hwnd Or GetActiveWindow = frmMain.hwnd) Then
                    NotFocused = False

                Else
                    Err.Clear
                End If
                On Error GoTo 0


            End If
            
            If (Not NotFocused) And (Err.number = 0) Then
                On Error Resume Next
                InitGameData
                If Err.number <> 0 Then
                    TermGameData
                End If
            Else
                TermDirectX
            End If
        Else
            DoTasks
        End If
                    
    Else
        Dim x1 As Long
        Dim y1 As Long
        
        Dim expCap As Boolean
        expCap = (frmStudio.ExportCapture <> 0)

        On Error GoTo Render

        Dim dViewPort As D3DVIEWPORT8
        DDevice.GetViewport dViewPort
        
      '  dViewPort.X = ((frmStudio.Left + frmStudio.Designer.Left) / Screen.TwipsPerPixelX)
      '  dViewPort.Y = ((frmStudio.Top + frmStudio.Designer.Top) / Screen.TwipsPerPixelX)

        dViewPort.Width = (frmStudio.Designer.Width / Screen.TwipsPerPixelX) '- ((frmStudio.Left + frmStudio.Designer.Left) / Screen.TwipsPerPixelX)
        dViewPort.Height = (frmStudio.Designer.Height / Screen.TwipsPerPixelY) '- ((frmStudio.Top + frmStudio.Designer.Top) / Screen.TwipsPerPixelX)
        DDevice.SetViewport dViewPort
    
        DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1, 0
        DDevice.BeginScene

        On Error GoTo nofocus

        SetupWorld
        
        If Not expCap Then InputScene

        RenderView expCap

        RenderInfo

        DDevice.EndScene
        
        On Error Resume Next
        DDevice.Present ByVal 0, ByVal 0, frmMain.Picture1.hwnd, ByVal 0
   
        If expCap Then
            frmStudio.ExportCapture = frmStudio.ExportCapture + 1

            If frmStudio.ExportCapture Mod 3 = 0 Then
            
                frmStudio.FinishCapture
            End If

            
        End If
   
        FPSCount = FPSCount + 1
        If (FPSTimer = 0) Or ((Timer - FPSTimer) >= 1) Then
            FPSTimer = Timer
            FPSRate = FPSCount
            FPSCount = 0
        End If
        If Err.number Then
            Err.Clear
            DoNotFocused
        End If
        On Error GoTo 0
        
    End If

Exit Sub
nofocus:
    Debug.Print Err.Description
    Err.Clear
    DoNotFocused

Exit Sub
Render:
    TermGameData
    TermDirectX

    Unload frmMain
        
    MsgBox "There was an error trying to run the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.number & ", " & Err.Description, vbOKOnly + vbInformation, App.Title
    Err.Clear
    End
End Sub

Public Sub SetupWorld()

    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX

    D3DXMatrixIdentity matRot
    D3DXMatrixIdentity matScale

    
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matWorld As D3DMATRIX
    Dim matPos As D3DMATRIX
    
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
     
    D3DXMatrixRotationY matRotation, Player.CameraAngle
    D3DXMatrixRotationX matPitch, Player.CameraPitch
    D3DXMatrixMultiply matLook, matRotation, matPitch
    
    D3DXMatrixTranslation matPos, -Player.Location.X, -Player.Location.Y, Player.CameraZoom
    D3DXMatrixMultiply matView, matPos, matLook
    DDevice.SetTransform D3DTS_VIEW, matView
        
    

    D3DXMatrixPerspectiveFovLH matProj, FOVY, ASPECT, NEAR, FAR
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    GenericMaterial.Ambient.a = 1
    GenericMaterial.Ambient.R = 1
    GenericMaterial.Ambient.G = 1
    GenericMaterial.Ambient.B = 1
    GenericMaterial.diffuse.a = 1
    GenericMaterial.diffuse.R = 1
    GenericMaterial.diffuse.G = 1
    GenericMaterial.diffuse.B = 1
    GenericMaterial.power = 1

    LucentMaterial.Ambient.a = 1
    LucentMaterial.Ambient.R = 1
    LucentMaterial.Ambient.G = 1
    LucentMaterial.Ambient.B = 1
    LucentMaterial.diffuse.a = 1
    LucentMaterial.diffuse.R = 0
    LucentMaterial.diffuse.G = 0
    LucentMaterial.diffuse.B = 0
    LucentMaterial.power = 1
    
End Sub

Public Sub InitDirectX()

    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    Set D3DX = New D3DX8
        
    InitialDevice frmMain.Picture1.hwnd
    
'    Set DSound = dx.DirectSoundCreate("")
'    DSound.SetCooperativeLevel frmMain.Picture1.hwnd, DSSCL_PRIORITY
        
End Sub

Private Sub InitialDevice(ByVal hwnd As Long)
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display

    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferWidth = Screen.Width / Screen.TwipsPerPixelX
    D3DWindow.BackBufferHeight = Screen.Height / Screen.TwipsPerPixelY
    D3DWindow.BackBufferFormat = Display.Format
    D3DWindow.Windowed = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    D3DWindow.hDeviceWindow = hwnd
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    D3DWindow.EnableAutoDepthStencil = True
    
    On Error Resume Next
    Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
    If Err.number <> 0 Then
        Err.Clear
        Set DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
        If Err.number <> 0 Then
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
        DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 32
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
    
    End If
    
End Sub

Public Sub DoNotFocused()
    On Error Resume Next
    NotFocused = True
    
'    TermGameData
'    TermDirectX
End Sub

Public Sub InitGameData()

    CreateCmds
    CreateInfo
    CreateText
    CreateProj

    
End Sub

Public Sub TermGameData()
    
    CleanUpProj
    CleanupText
    CleanupInfo
    CleanupCmds


End Sub

Public Sub TermDirectX()
    NotFocused = True
    
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

Public Function TestDirectX() As Boolean

    On Error Resume Next
    InitDirectX
    TestDirectX = (Err.number = 0)
    If Err.number Then Err.Clear
    On Error GoTo 0

End Function





