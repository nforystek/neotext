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



Public ReflectRenderTarget As Direct3DSurface8

Public ReflectFrontBuffer As Direct3DSurface8

Public ReflectStencilDepth As Direct3DSurface8
Public BufferedTexture As Direct3DTexture8

Public ColumnCount As Long
Public RowCount As Long


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
    GenericMaterial.Ambient.b = 1
    GenericMaterial.Diffuse.A = 1
    GenericMaterial.Diffuse.r = 1
    GenericMaterial.Diffuse.g = 1
    GenericMaterial.Diffuse.b = 1
    GenericMaterial.Power = 1

    LucentMaterial.Ambient.A = 1
    LucentMaterial.Ambient.r = 1
    LucentMaterial.Ambient.g = 1
    LucentMaterial.Ambient.b = 1
    LucentMaterial.Diffuse.A = 1
    LucentMaterial.Diffuse.r = 0
    LucentMaterial.Diffuse.g = 0
    LucentMaterial.Diffuse.b = 0
    LucentMaterial.Power = 1

    SpecialMaterial.Ambient.A = 0
    SpecialMaterial.Ambient.r = 0.89
    SpecialMaterial.Ambient.g = 0.89
    SpecialMaterial.Ambient.b = 0.89
    SpecialMaterial.Diffuse.A = 0.4
    SpecialMaterial.Diffuse.r = 0.01
    SpecialMaterial.Diffuse.g = 0.01
    SpecialMaterial.Diffuse.b = 0.01
    SpecialMaterial.Specular.A = 0
    SpecialMaterial.Specular.r = 0.5
    SpecialMaterial.Specular.g = 0.5
    SpecialMaterial.Specular.b = 0.5
    SpecialMaterial.emissive.A = 0.3
    SpecialMaterial.emissive.r = 0.21
    SpecialMaterial.emissive.g = 0.3
    SpecialMaterial.emissive.b = 0.3
    SpecialMaterial.Power = 0

    Set DefaultRenderTarget = DDevice.GetRenderTarget
    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface
    
    If Not FullScreen Then
        Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        Set ReflectRenderTarget = BufferedTexture.GetSurfaceLevel(0)
        Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), D3DFMT_D24S8, D3DMULTISAMPLE_NONE)
    Else
        Set BufferedTexture = DDevice.CreateTexture((Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY), 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        Set ReflectRenderTarget = BufferedTexture.GetSurfaceLevel(0)
        Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY), D3DFMT_D24S8, D3DMULTISAMPLE_NONE)
    End If
''
''    If Not FullScreen Then
''        Set DSurface = D3DX.CreateRenderToSurface(DDevice, (frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), Display.Format, D3DMULTISAMPLE_NONE, D3DFMT_D16)
'
''        Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
''        Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
''        Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
'
'
''    Else
''        Set DSurface = D3DX.CreateRenderToSurface(DDevice, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, Display.Format, D3DMULTISAMPLE_NONE, D3DFMT_D16)
''    End If
'
''    Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
''    Set BufferedTexture = DSurface.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
''    Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
'
'
'
'        If Not FullScreen Then
'            'Set DSurface = D3DX.CreateRenderToSurface(DDevice, (frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, 1, D3DFMT_D24S8)
'            'Set DSurface = D3DX.CreateRenderToSurface(DDevice, (frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), Display.Format, 1, D3DFMT_D16)
'            Set ReflectRenderTarget = DDevice.CreateImageSurface((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), Display.Format)
'        Else
'            'Set DSurface = D3DX.CreateRenderToSurface(DDevice, (VB.Screen.Width / VB.Screen.TwipsPerPixelX), (VB.Screen.Height / VB.Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, 1, D3DFMT_D24S8)
'            'Set DSurface = D3DX.CreateRenderToSurface(DDevice, (VB.Screen.Width / VB.Screen.TwipsPerPixelX), (VB.Screen.Height / VB.Screen.TwipsPerPixelY), Display.Format, 1, D3DFMT_D16)
'            Set ReflectRenderTarget = DDevice.CreateImageSurface((VB.Screen.Width / VB.Screen.TwipsPerPixelX), (VB.Screen.Height / VB.Screen.TwipsPerPixelY), Display.Format)
'
'        End If
'
'
'
'          '  Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
'          '  Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
'          '  Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
'
''        Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
''        Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
''        Set ReflectRenderTarget = DDevice.CreateImageSurface((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), D3DFMT_A8R8G8B8)
''        Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), D3DFMT_D24S8, D3DMULTISAMPLE_NONE)
'
'
''    DDevice.GetFrontBuffer ReflectFrontBuffer
''
''
''
''
'' '   DDevice.SetClipPlane
'
'
'
''    Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
''    Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
''    Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_D16, D3DMULTISAMPLE_NONE) ' CONST_D3DFORMAT.D3DFMT_D24S8, D3DMULTISAMPLE_NONE)


End Sub

Public Sub CleanupText()
    Set DefaultRenderTarget = Nothing
    Set DefaultStencilDepth = Nothing

    Set ReflectRenderTarget = Nothing
    Set ReflectStencilDepth = Nothing
    Set ReflectFrontBuffer = Nothing
    Set BufferedTexture = Nothing
    
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


Public Function LoadTexture(ByVal FileName As String) As Direct3DTexture8
    Dim e As String
    Dim t As Direct3DTexture8
    Dim Dimensions As ImageDimensions
    
    If ImageDimensions(FileName, Dimensions, e) Then
        Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
        Set LoadTexture = t

    End If
    Set t = Nothing
End Function


Public Function LoadTextureEx(ByVal FileName As String, ByRef Dimensions As ImageDimensions) As Direct3DTexture8
    Dim e As String
    Dim t As Direct3DTexture8
    
    If ImageDimensions(FileName, Dimensions, e) Then
        Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
        Set LoadTextureEx = t

    End If
    Set t = Nothing
End Function


