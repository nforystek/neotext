Attribute VB_Name = "modText"
#Const modText = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public DPI As ImageDimensions

Public SpecialMaterial As D3DMATERIAL8
Public GenericMaterial As D3DMATERIAL8
Public LucentMaterial As D3DMATERIAL8

'###########################################################################
'###################### BEGIN UNIQUE NON GLOBALS ###########################
'###########################################################################

Public TextColor As Long

Public Fnt As StdFont
Public MainFont As D3DXFont
Public MainFontDesc As IFont


Public DefaultRenderTarget As Direct3DSurface8
Public DefaultStencilDepth As Direct3DSurface8

Public DSurface As D3DXRenderToSurface
Public ReflectRenderTarget As Direct3DSurface8

'Public ReflectFrontBuffer As Direct3DSurface8
Public ReflectStencilDepth As Direct3DSurface8
'Public BufferedTexture As Direct3DTexture8

Public ColumnCount As Long
Public RowCount As Long

Public Function GetDynamic(ByVal FileName As String) As String
    If PathExists(ScriptRoot & "Models\" & FileName, True) Then
        GetDynamic = "Size:" & FileLen(ScriptRoot & "Models\" & FileName) & "Date:" & FileDateTime(ScriptRoot & "Models\" & FileName)
    Else
        GetDynamic = "Size:Date:"
    End If
End Function

Public Sub CreateText()

    TextColor = D3DColorARGB(255, 0, 0, 0)

    frmMain.Font.Name = "Lucida Console"
    frmMain.Font.Bold = False
    frmMain.Font.Italic = False
    frmMain.Font.CharSet = 0

    DPI = GetMonitorDPI

    Dim TwipsRatioPerCharInch As Double
    Dim DotsRatioPerCharInch As Double
    Dim PixelCubicCharInch As Double

    Dim PixelPerDotCharHeight As Double
    Dim PixelPerDotCharWidth As Double

    TwipsRatioPerCharInch = Sqr(((LetterPerInch * Screen.TwipsPerPixelX) * TextHeight) / Screen.TwipsPerPixelY)
    DotsRatioPerCharInch = Sqr(LetterPerInch * (TextHeight / Screen.TwipsPerPixelY))
    PixelPerDotCharHeight = TwipsRatioPerCharInch / DotsRatioPerCharInch
    PixelCubicCharInch = Sqr((DPI.Height * Screen.TwipsPerPixelY) * (DPI.Width * Screen.TwipsPerPixelX)) / (LetterPerInch ^ 2)
    PixelPerDotCharWidth = Sqr((TwipsRatioPerCharInch + DotsRatioPerCharInch + PixelCubicCharInch) * 4) / PixelCubicCharInch

    ColumnCount = ((Screen.Width / Screen.TwipsPerPixelX) / Round(DPI.Width * (LetterPerInch / 100), 0)) * ((frmMain.Width / Screen.TwipsPerPixelX) / (Screen.Width / Screen.TwipsPerPixelX))
    frmMain.Font.Size = (frmMain.Width / (Screen.TwipsPerPixelX * PixelPerDotCharWidth)) / ColumnCount

    RowCount = 1
    Do Until ((((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((frmMain.ScaleHeight - TextHeight) / Screen.TwipsPerPixelY)
        RowCount = RowCount + 1
    Loop


    Set Fnt = frmMain.Font
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(DDevice, MainFontDesc.hFont)

    GenericMaterial.Ambient.A = 1
    GenericMaterial.Ambient.r = 1
    GenericMaterial.Ambient.g = 1
    GenericMaterial.Ambient.B = 1
    GenericMaterial.Diffuse.A = 1
    GenericMaterial.Diffuse.r = 1
    GenericMaterial.Diffuse.g = 1
    GenericMaterial.Diffuse.B = 1
    GenericMaterial.Power = 1

    LucentMaterial.Ambient.A = 1
    LucentMaterial.Ambient.r = 1
    LucentMaterial.Ambient.g = 1
    LucentMaterial.Ambient.B = 1
    LucentMaterial.Diffuse.A = 1
    LucentMaterial.Diffuse.r = 0
    LucentMaterial.Diffuse.g = 0
    LucentMaterial.Diffuse.B = 0
    LucentMaterial.Power = 1

    SpecialMaterial.Ambient.A = 0
    SpecialMaterial.Ambient.r = 0.89
    SpecialMaterial.Ambient.g = 0.89
    SpecialMaterial.Ambient.B = 0.89
    SpecialMaterial.Diffuse.A = 0.4
    SpecialMaterial.Diffuse.r = 0.01
    SpecialMaterial.Diffuse.g = 0.01
    SpecialMaterial.Diffuse.B = 0.01
    SpecialMaterial.Specular.A = 0
    SpecialMaterial.Specular.r = 0.5
    SpecialMaterial.Specular.g = 0.5
    SpecialMaterial.Specular.B = 0.5
    SpecialMaterial.emissive.A = 0.3
    SpecialMaterial.emissive.r = 0.21
    SpecialMaterial.emissive.g = 0.3
    SpecialMaterial.emissive.B = 0.3
    SpecialMaterial.Power = 0

    Set DefaultRenderTarget = DDevice.GetRenderTarget
    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface
    

    Dim Width As Single
    Dim Height As Single
    
    If Not FullScreen Then
        Width = (frmMain.Width / Screen.TwipsPerPixelX)
        Height = (frmMain.Height / Screen.TwipsPerPixelY)
    Else
        Width = (Screen.Width / Screen.TwipsPerPixelX)
        Height = (Screen.Height / Screen.TwipsPerPixelY)
    End If


    '#######################################################################################################################
    '######## Other testing and/or debuggin attempts of able/figuring out multiple rendering surfaces ######################
    '#######################################################################################################################
    
    Set DSurface = D3DX.CreateRenderToSurface(DDevice, Width, Height, Display.Format, False, D3DFMT_D16)
    Set ReflectRenderTarget = DDevice.CreateRenderTarget(Width, Height, Display.Format, D3DMULTISAMPLE_NONE, True)
    
    
'    Set DSurface = D3DX.CreateRenderToSurface(DDevice, Width, Height, Display.Format, 1, D3DFMT_D16)
'    Set ReflectRenderTarget = DDevice.CreateRenderTarget(Width, Height, Display.Format, D3DMULTISAMPLE_NONE, True)
'    Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface(Width, Height, D3DFMT_D16, D3DMULTISAMPLE_NONE)


    '#######################################################################################################################
    '######## Other testing and/or debuggin attempts of able/figuring out multiple rendering surfaces ######################
    '#######################################################################################################################


    'Set BufferedTexture = DDevice.CreateTexture(Width, Height, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    'Set ReflectRenderTarget = BufferedTexture.GetSurfaceLevel(0)
    'Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface(Width, Height, D3DFMT_D24S8, D3DMULTISAMPLE_NONE)



'    Set ReflectRenderTarget = DDevice.CreateImageSurface(Width, Height, Display.Format)
'    Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface(Width, Height, D3DFMT_D16, D3DMULTISAMPLE_NONE)


   'Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface(Width, Height, D3DFMT_D24S8, D3DMULTISAMPLE_NONE)


'    Set ReflectRenderTarget = DDevice.CreateRenderTarget(Width, Height, D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
'    Set ReflectFrontBuffer = DDevice.CreateImageSurface(Width, Height, D3DFMT_A8R8G8B8)
'    DDevice.GetFrontBuffer ReflectFrontBuffer





End Sub

Public Sub CleanupText()
    Set DefaultRenderTarget = Nothing
    Set DefaultStencilDepth = Nothing

    Set DSurface = Nothing
    Set ReflectRenderTarget = Nothing
    Set ReflectStencilDepth = Nothing
'    Set ReflectFrontBuffer = Nothing
'    Set BufferedTexture = Nothing
    
    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set Fnt = Nothing
End Sub

'Public Function DrawText(Text As String, X As Single, Y As Single)
'
'    Dim TextRect As RECT
'    Dim Allignment As CONST_DTFLAGS
'    Allignment = DT_TOP Or DT_LEFT
'
'    TextRect.Top = Y
'    TextRect.Left = X
'    TextRect.Bottom = Y + (frmMain.TextHeight(Text) / Screen.TwipsPerPixelY)
'    TextRect.Right = X + (frmMain.TextWidth(Text) / Screen.TwipsPerPixelX)
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'    DDevice.SetPixelShader PixelShaderDefault
'    DDevice.SetVertexShader FVF_RENDER
'
'    D3DX.DrawText MainFont, TextColor, Text, TextRect, Allignment
'End Function


'Public Function ReadFile(ByVal Path As String) As String
'    Dim Num As Long
'    Dim Text As String
'    Dim timeout As Single
'
'    Num = FreeFile
'    On Error Resume Next
'    On Local Error Resume Next
'    If PathExists(Path, True) Then
'        Open Path For Append Shared As #Num Len = 1 ' LenB(Chr(CByte(0)))
'        Close #Num
'        Select Case Err.Number
'            Case 54, 70, 75
'                Err.Clear
'                On Error GoTo tryagain
'                On Local Error GoTo tryagain
'
'                Open Path For Binary Access Read Lock Write As Num Len = 1
'                If timeout <> 0 Then
'                    Open Path For Binary Shared As #Num Len = 1
'                End If
'                Text = String(LOF(Num), " ")
'                Get #Num, 1, Text
'                Close #Num
'            Case Else
'                On Error GoTo tryagain
'                On Local Error GoTo tryagain
'
'                Open Path For Binary Access Read As Num Len = 1
'                If timeout <> 0 Then
'                    Open Path For Binary Shared As Num Len = 1
'                End If
'                Text = String(LOF(Num), " ")
'                Get #Num, 1, Text
'                Close #Num
'        End Select
'        If Err Then GoTo failit
'        On Error GoTo 0
'        On Local Error GoTo 0
'    End If
'    ReadFile = Text

'End Function


Public Function LoadTexture(ByVal FileName As String) As Direct3DTexture8
    Dim e As String
    Dim t As Direct3DTexture8
    Dim Dimensions As ImageDimensions
    Dim timeout As Single
    Dim Num As Long

    On Error Resume Next
    On Local Error Resume Next
                
    If ImageDimensions(FileName, Dimensions, e) Then
        Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
        Set LoadTexture = t
        Set t = Nothing
    End If
    If Err.Number <> 0 Then
    
        On Error GoTo tryagain
        On Local Error GoTo tryagain

        If ImageDimensions(FileName, Dimensions, e) Then
            Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
                D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
            Set LoadTexture = t
            Set t = Nothing
        End If
        
    End If

    If Err Then GoTo failit
    On Error GoTo 0
    On Local Error GoTo 0
    
    Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        On Error GoTo failit
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "LoadTexture"
End Function


Public Function LoadTextureEx(ByVal FileName As String, ByRef Dimensions As ImageDimensions) As Direct3DTexture8
    Dim e As String
    Dim t As Direct3DTexture8
    Dim timeout As Single
    Dim Num As Long

    On Error Resume Next
    On Local Error Resume Next
                
    If ImageDimensions(FileName, Dimensions, e) Then
        Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
        Set LoadTextureEx = t
        Set t = Nothing
    End If
    
    If Err.Number <> 0 Then
    
        On Error GoTo tryagain
        On Local Error GoTo tryagain

        If ImageDimensions(FileName, Dimensions, e) Then
            Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
                D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
            Set LoadTextureEx = t
            Set t = Nothing
        End If
        
    End If
    
    If Err Then GoTo failit
    On Error GoTo 0
    On Local Error GoTo 0
    
        Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        On Error GoTo failit
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "LoadTextureEx"
End Function


