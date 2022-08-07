Attribute VB_Name = "modText"
#Const modText = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module


Public ColumnCount As Long
Public RowCount As Long

Public Fnt As StdFont
Public MainFont As D3DXFont
Public MainFontDesc As IFont

Public Type ImgDimType
  height As Long
  width As Long
End Type

Public DPI As ImgDimType

Public TextColor As Long

Public SpecialMaterial As D3DMATERIAL8
Public GenericMaterial As D3DMATERIAL8
Public LucentMaterial As D3DMATERIAL8

Public DefaultRenderTarget As Direct3DSurface8
Public DefaultStencilDepth As Direct3DSurface8

Public BufferedTexture As Direct3DTexture8
Public ReflectRenderTarget As Direct3DSurface8
Public ReflectStencilDepth As Direct3DSurface8

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long

Private Const LOGPIXELSX = 88 ' Logical pixels/inch in X
Private Const LOGPIXELSY = 90 ' Logical pixels/inch in Y

Public Function GetMonitorDPI() As ImgDimType
    Dim hDc As Long
    Dim lngRetVal As Long

    hDc = GetDC(0)

    GetMonitorDPI.width = GetDeviceCaps(hDc, LOGPIXELSX)
    GetMonitorDPI.height = GetDeviceCaps(hDc, LOGPIXELSY)

    lngRetVal = ReleaseDC(0, hDc)

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
    PixelCubicCharInch = Sqr((DPI.height * Screen.TwipsPerPixelY) * (DPI.width * Screen.TwipsPerPixelX)) / (LetterPerInch ^ 2)
    PixelPerDotCharWidth = Sqr((TwipsRatioPerCharInch + DotsRatioPerCharInch + PixelCubicCharInch) * 4) / PixelCubicCharInch

    ColumnCount = ((Screen.width / Screen.TwipsPerPixelX) / Round(DPI.width * (LetterPerInch / 100), 0)) * ((frmMain.width / Screen.TwipsPerPixelX) / (Screen.width / Screen.TwipsPerPixelX))
    frmMain.Font.Size = (frmMain.width / (Screen.TwipsPerPixelX * PixelPerDotCharWidth)) / ColumnCount

    RowCount = 1
    Do Until ((((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((frmMain.ScaleHeight - TextHeight) / Screen.TwipsPerPixelY)
        RowCount = RowCount + 1
    Loop


    Set Fnt = frmMain.Font
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(DDevice, MainFontDesc.hFont)

    GenericMaterial.Ambient.a = 1
    GenericMaterial.Ambient.r = 1
    GenericMaterial.Ambient.g = 1
    GenericMaterial.Ambient.b = 1
    GenericMaterial.diffuse.a = 1
    GenericMaterial.diffuse.r = 1
    GenericMaterial.diffuse.g = 1
    GenericMaterial.diffuse.b = 1
    GenericMaterial.power = 1

    LucentMaterial.Ambient.a = 1
    LucentMaterial.Ambient.r = 1
    LucentMaterial.Ambient.g = 1
    LucentMaterial.Ambient.b = 1
    LucentMaterial.diffuse.a = 1
    LucentMaterial.diffuse.r = 0
    LucentMaterial.diffuse.g = 0
    LucentMaterial.diffuse.b = 0
    LucentMaterial.power = 1

    SpecialMaterial.Ambient.a = 0
    SpecialMaterial.Ambient.r = 0.89
    SpecialMaterial.Ambient.g = 0.89
    SpecialMaterial.Ambient.b = 0.89
    SpecialMaterial.diffuse.a = 0.4
    SpecialMaterial.diffuse.r = 0.01
    SpecialMaterial.diffuse.g = 0.01
    SpecialMaterial.diffuse.b = 0.01
    SpecialMaterial.specular.a = 0
    SpecialMaterial.specular.r = 0.5
    SpecialMaterial.specular.g = 0.5
    SpecialMaterial.specular.b = 0.5
    SpecialMaterial.emissive.a = 0.3
    SpecialMaterial.emissive.r = 0.21
    SpecialMaterial.emissive.g = 0.3
    SpecialMaterial.emissive.b = 0.3
    SpecialMaterial.power = 0

    Set DefaultRenderTarget = DDevice.GetRenderTarget
    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface

'    Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
'    Set ReflectRenderTarget = BufferedTexture.GetSurfaceLevel(0)
'    Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), D3DFMT_D24S8, D3DMULTISAMPLE_NONE)

End Sub

Public Sub CleanupText()
    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set Fnt = Nothing
End Sub

Public Function DrawText(Text As String, X As Single, Y As Single)

    Dim TextRect As RECT
    Dim Allignment As CONST_DTFLAGS
    Allignment = DT_TOP Or DT_LEFT

    TextRect.Top = Y
    TextRect.Left = X
    TextRect.Bottom = Y + (frmMain.TextHeight(Text) / Screen.TwipsPerPixelY)
    TextRect.Right = X + (frmMain.TextWidth(Text) / Screen.TwipsPerPixelX)

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    DDevice.SetPixelShader PixelShaderDefault
    DDevice.SetVertexShader FVF_RENDER

    D3DX.DrawText MainFont, TextColor, Text, TextRect, Allignment
End Function

Public Function ImageDimensions(ByVal FileName As String, ByRef ImgDim As ImgDimType, Optional ByRef Ext As String = "") As Boolean

    If PathExists(FileName, True) Then
    
        'Inputs:
        '
        'fileName is a string containing the path name of the image file.
        '
        'ImgDim is passed as an empty type var and contains the height
        'and width that's passed back.
        '
        'Ext is passed as an empty string and contains the image type
        'as a 3 letter description that's passed back.
        '
        'Returns:
        '
        'True if the function was successful.
        
        'declare vars
        Dim handle As Integer, isValidImage As Boolean
        Dim byteArr(255) As Byte, i As Integer
        
        'init vars
        isValidImage = False
        ImgDim.height = 0
        ImgDim.width = 0
        
        'open file and get 256 byte chunk
        handle = FreeFile
        On Error GoTo endFunction
        Open FileName For Binary Access Read As #handle
            
            Get handle, , byteArr
        Close #handle
        
        'check for jpg header (SOI): &HFF and &HD8
        ' contained in first 2 bytes
        If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
            isValidImage = True
        Else
            GoTo checkGIF
        End If
        
        'check for SOF marker: &HFF and &HC0 TO &HCF
        For i = 0 To 255
            If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 And byteArr(i + 1) <= &HCF Then
                ImgDim.height = byteArr(i + 5) * 256 + byteArr(i + 6)
                ImgDim.width = byteArr(i + 7) * 256 + byteArr(i + 8)
                Exit For
            End If
        Next i
        
        'get image type and exit
        Ext = "jpg"
        GoTo endFunction
        
checkGIF:
        
        'check for GIF header
        If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 And byteArr(3) = &H38 Then
            ImgDim.width = byteArr(7) * 256 + byteArr(6)
            ImgDim.height = byteArr(9) * 256 + byteArr(8)
            isValidImage = True
        Else
            GoTo checkBMP
        End If
        
        'get image type and exit
        Ext = "gif"
        GoTo endFunction
        
checkBMP:
        
        'check for BMP header
        If byteArr(0) = 66 And byteArr(1) = 77 Then
            isValidImage = True
        Else
            GoTo checkPNG
        End If
        
        'get record type info
        If byteArr(14) = 40 Then
        
            'get width and height of BMP
            ImgDim.width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
            + byteArr(19) * 256 + byteArr(18)
            
            ImgDim.height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
            + byteArr(23) * 256 + byteArr(22)
        
        'another kind of BMP
        ElseIf byteArr(17) = 12 Then
        
            'get width and height of BMP
            ImgDim.width = byteArr(19) * 256 + byteArr(18)
            ImgDim.height = byteArr(21) * 256 + byteArr(20)
            
        End If
        
        'get image type and exit
        Ext = "bmp"
        GoTo endFunction
        
checkPNG:
        
        'check for PNG header
        If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E And byteArr(3) = &H47 Then
            ImgDim.width = byteArr(18) * 256 + byteArr(19)
            ImgDim.height = byteArr(22) * 256 + byteArr(23)
            isValidImage = True
        Else
            GoTo endFunction
        End If
        
        Ext = "png"
    
    Else
        AddMessage "Invalid picture file [" & FileName & "]"
    End If
endFunction:
    
    'return function's success status
    ImageDimensions = isValidImage

End Function

Public Function LoadTexture(ByVal FileName As String) As Direct3DTexture8
    Dim d As ImgDimType
    Dim e As String
    Dim t As Direct3DTexture8
    
    If ImageDimensions(FileName, d, e) Then
        Set t = D3DX.CreateTextureFromFileEx(DDevice, FileName, d.width, d.height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
        Set LoadTexture = t
    End If
End Function

