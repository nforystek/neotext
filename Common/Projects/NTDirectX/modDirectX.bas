Attribute VB_Name = "modDirectX"
Option Explicit
'
'Public PixelShaderDefault As Long
'Public PixelShaderDiffuse As Long
'
'Public SpecialMaterial As D3DMATERIAL8
'Public GenericMaterial As D3DMATERIAL8
'Public LucentMaterial As D3DMATERIAL8
'
'Public DefaultRenderTarget As Direct3DSurface8
'Public DefaultStencilDepth As Direct3DSurface8
'Public BufferedTexture As Direct3DTexture8
'Public ReflectRenderTarget As Direct3DSurface8
'Public ReflectFrontBuffer As Direct3DSurface8
'
'Public Sub CreateText()
'
'    GenericMaterial.Ambient.a = 1
'    GenericMaterial.Ambient.R = 1
'    GenericMaterial.Ambient.g = 1
'    GenericMaterial.Ambient.B = 1
'    GenericMaterial.diffuse.a = 1
'    GenericMaterial.diffuse.R = 1
'    GenericMaterial.diffuse.g = 1
'    GenericMaterial.diffuse.B = 1
'    GenericMaterial.power = 1
'
'    LucentMaterial.Ambient.a = 1
'    LucentMaterial.Ambient.R = 1
'    LucentMaterial.Ambient.g = 1
'    LucentMaterial.Ambient.B = 1
'    LucentMaterial.diffuse.a = 1
'    LucentMaterial.diffuse.R = 0
'    LucentMaterial.diffuse.g = 0
'    LucentMaterial.diffuse.B = 0
'    LucentMaterial.power = 1
'
'    SpecialMaterial.Ambient.a = 0
'    SpecialMaterial.Ambient.R = 0.89
'    SpecialMaterial.Ambient.g = 0.89
'    SpecialMaterial.Ambient.B = 0.89
'    SpecialMaterial.diffuse.a = 0.4
'    SpecialMaterial.diffuse.R = 0.01
'    SpecialMaterial.diffuse.g = 0.01
'    SpecialMaterial.diffuse.B = 0.01
'    SpecialMaterial.specular.a = 0
'    SpecialMaterial.specular.R = 0.5
'    SpecialMaterial.specular.g = 0.5
'    SpecialMaterial.specular.B = 0.5
'    SpecialMaterial.emissive.a = 0.3
'    SpecialMaterial.emissive.R = 0.21
'    SpecialMaterial.emissive.g = 0.3
'    SpecialMaterial.emissive.B = 0.3
'    SpecialMaterial.power = 0
'
'    Set DefaultRenderTarget = DDevice.GetRenderTarget
'    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface
'    Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
'    Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
'    Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
'
'End Sub
'
'Public Sub CleanupText()
'
'    Set DefaultRenderTarget = Nothing
'    Set DefaultStencilDepth = Nothing
'
'End Sub
