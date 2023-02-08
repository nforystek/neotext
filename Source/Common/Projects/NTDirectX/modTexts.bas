Attribute VB_Name = "modTexts"


#Const modText = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public ColumnCount As Long
Public RowCount As Long

Public Fnt As StdFont
Public MainFont As D3DXFont
Public MainFontDesc As IFont

Public Type ImgDimType
  Height As Long
  Width As Long
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
Public ReflectFrontBuffer As Direct3DSurface8

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Const LOGPIXELSX = 88 ' Logical pixels/inch in X
Private Const LOGPIXELSY = 90 ' Logical pixels/inch in Y

Public Function GetMonitorDPI() As ImgDimType
    Dim hdc As Long
    Dim lngRetVal As Long

    hdc = GetDC(0)

    GetMonitorDPI.Width = GetDeviceCaps(hdc, LOGPIXELSX)
    GetMonitorDPI.Height = GetDeviceCaps(hdc, LOGPIXELSY)

    lngRetVal = ReleaseDC(0, hdc)

End Function

Public Sub CreateText()

    TextColor = D3DColorARGB(255, 0, 0, 0)

    frmMain.Font.Name = "Lucida Console"
    frmMain.Font.Bold = False
    frmMain.Font.Italic = False
    frmMain.Font.Charset = 0

    DPI = GetMonitorDPI

    Dim TwipsRatioPerCharInch As Double
    Dim DotsRatioPerCharInch As Double
    Dim PixelCubicCharInch As Double

    Dim PixelPerDotCharHeight As Double
    Dim PixelPerDotCharWidth As Double

    TwipsRatioPerCharInch = Sqr(((LetterPerInch * VB.Screen.TwipsPerPixelX) * TextHeight) / VB.Screen.TwipsPerPixelY)
    DotsRatioPerCharInch = Sqr(LetterPerInch * (TextHeight / VB.Screen.TwipsPerPixelY))
    PixelPerDotCharHeight = TwipsRatioPerCharInch / DotsRatioPerCharInch
    PixelCubicCharInch = Sqr((DPI.Height * VB.Screen.TwipsPerPixelY) * (DPI.Width * VB.Screen.TwipsPerPixelX)) / (LetterPerInch ^ 2)
    PixelPerDotCharWidth = Sqr((TwipsRatioPerCharInch + DotsRatioPerCharInch + PixelCubicCharInch) * 4) / PixelCubicCharInch

    'ColumnCount = ((Screen.Width / VB.Screen.TwipsPerPixelX) / Round(DPI.Width * (LetterPerInch / 100), 0)) * ((frmMain.Width / VB.Screen.TwipsPerPixelX) / (Screen.Width / VB.Screen.TwipsPerPixelX))
    ColumnCount = 129

'    frmMain.Font.Size = (frmMain.Width / (VB.Screen.TwipsPerPixelX * PixelPerDotCharWidth)) / ColumnCount
    Dim Size As Long
    Size = 30
    frmMain.Font.Size = Size
    Do Until frmMain.TextWidth(String(ColumnCount, "A")) < (VB.Screen.Width - (TextSpace * 2))
        Size = Size - 1
        frmMain.Font.Size = Size
    Loop


'    RowCount = 1
'    Do Until ((((TextHeight / VB.Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((frmMain.ScaleHeight - TextHeight) / VB.Screen.TwipsPerPixelY)
'        RowCount = RowCount + 1
'    Loop
    RowCount = 1
    Do Until ((((TextHeight / VB.Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((VB.Screen.Height - TextHeight) / VB.Screen.TwipsPerPixelY)
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
    GenericMaterial.power = 1

    LucentMaterial.Ambient.A = 1
    LucentMaterial.Ambient.r = 1
    LucentMaterial.Ambient.g = 1
    LucentMaterial.Ambient.b = 1
    LucentMaterial.Diffuse.A = 1
    LucentMaterial.Diffuse.r = 0
    LucentMaterial.Diffuse.g = 0
    LucentMaterial.Diffuse.b = 0
    LucentMaterial.power = 1

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
    SpecialMaterial.power = 0

    Set DefaultRenderTarget = DDevice.GetRenderTarget
    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface

    Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
    

 '   Set ReflectFrontBuffer = DDevice.CreateImageSurface((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), D3DFMT_A8R8G8B8)
'
'    DDevice.GetFrontBuffer ReflectFrontBuffer
                                
    
    
    
 '   DDevice.SetClipPlane
    
    Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    
    Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)
    
 '   Set ReflectStencilDepth = DDevice.CreateDepthStencilSurface((frmMain.Width / VB.Screen.TwipsPerPixelX), (frmMain.Height / VB.Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_D16, D3DMULTISAMPLE_NONE) ' CONST_D3DFORMAT.D3DFMT_D24S8, D3DMULTISAMPLE_NONE)

End Sub



Public Sub CleanupText()

    Set DefaultRenderTarget = Nothing
    Set DefaultStencilDepth = Nothing

    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set Fnt = Nothing
End Sub

Public Function DrawText(Text As String, X As Single, Y As Single)

    Dim TextRect As DxVBLibA.RECT
    Dim Allignment As CONST_DTFLAGS
    Allignment = DT_TOP Or DT_LEFT

    TextRect.Top = Y
    TextRect.Left = X
    TextRect.Bottom = Y + (frmMain.TextHeight(Text) / VB.Screen.TwipsPerPixelY)
    TextRect.Right = X + (frmMain.TextWidth(Text) / VB.Screen.TwipsPerPixelX)

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    DDevice.SetPixelShader PixelShaderDefault
    DDevice.SetVertexShader FVF_RENDER

    D3DX.DrawText MainFont, TextColor, Text, TextRect, Allignment
End Function
'Public Function DrawTextByRowCol(Text As String, X As Long, Y As Long)
'
'    Dim rec As RECT
'
'    GetWindowRect frmMain.hWnd, rec
'    Dim screenX As Single
'    Dim screenY As Single
'    screenX = (rec.Left - rec.Right) / Screen.Width
'    screenY = (rec.Bottom - rec.Top) / Screen.Height
''        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)
'
'    Dim TextRect As RECT
'    Dim Color As Long
'    Dim Allignment As CONST_DTFLAGS
'    Color = &HFFFFFFFF
'    Allignment = DT_TOP Or DT_LEFT
'
'    TextRect.Top = Y
'    TextRect.Left = X
'    TextRect.Bottom = (Y + (frmMain.TextHeight(Text) / VB.Screen.TwipsPerPixelY))
'    TextRect.Right = X + (frmMain.TextWidth(Text) / VB.Screen.TwipsPerPixelX)
'
'    MainFont.Begin
'
'    D3DX.DrawText MainFont, &HFFFFFFFF, Text, TextRect, Allignment
'    MainFont.End
'
'End Function
'
'Public Function DrawTextByCoord(Text As String, X As Single, Y As Single)
'
''    Dim rec As RECT
''
''    GetWindowRect frmMain.hWnd, rec
''    Dim screenX As Single
''    Dim screenY As Single
''    screenX = (rec.Left - rec.Right) / Screen.Width
''    screenY = (rec.Bottom - rec.Top) / Screen.Height
''        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)
'
'
'    Dim TextRect As RECT
'    Dim Color As Long
'    Dim Allignment As CONST_DTFLAGS
'    Color = &HFFFFFFFF
'    Allignment = DT_TOP Or DT_LEFT
'
'    TextRect.Top = Y
'    TextRect.Left = X
'    TextRect.Bottom = (Y + (frmMain.TextHeight(Text) / VB.Screen.TwipsPerPixelY))
'    TextRect.Right = (X + (frmMain.TextWidth(Text) / VB.Screen.TwipsPerPixelX))
'
'
'    D3DX.DrawText MainFont, Color, Text, TextRect, Allignment
'
'
'End Function

'Public Function ImageDimensions(ByVal FileName As String, ByRef ImgDim As ImgDimType, Optional ByRef Ext As String = "") As Boolean
'
'    If PathExists(FileName, True) Then
'
'        'Inputs:
'        '
'        'fileName is a string containing the path name of the image file.
'        '
'        'ImgDim is passed as an empty type var and contains the height
'        'and width that's passed back.
'        '
'        'Ext is passed as an empty string and contains the image type
'        'as a 3 letter description that's passed back.
'        '
'        'Returns:
'        '
'        'True if the function was successful.
'
'        'declare vars
'        Dim handle As Integer, isValidImage As Boolean
'        Dim byteArr(255) As Byte, i As Integer
'
'        'init vars
'        isValidImage = False
'        ImgDim.height = 0
'        ImgDim.width = 0
'
'        'open file and get 256 byte chunk
'        handle = FreeFile
'        On Error GoTo endFunction
'        Open FileName For Binary Access Read As #handle
'            Get handle, , byteArr
'        Close #handle
'
'        'check for jpg header (SOI): &HFF and &HD8
'        ' contained in first 2 bytes
'        If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
'            isValidImage = True
'        Else
'            GoTo checkGIF
'        End If
'
'        'check for SOF marker: &HFF and &HC0 TO &HCF
'        For i = 0 To 255
'            If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 And byteArr(i + 1)<= &HCF Then
'                ImgDim.height = byteArr(i + 5) * 256 + byteArr(i + 6)
'                ImgDim.width = byteArr(i + 7) * 256 + byteArr(i + 8)
'                Exit For
'            End If
'        Next i
'
'        'get image type and exit
'        Ext = "jpg"
'        GoTo endFunction
'
'checkGIF:
'
'        'check for GIF header
'        If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 And byteArr(3) = &H38 Then
'            ImgDim.width = byteArr(7) * 256 + byteArr(6)
'            ImgDim.height = byteArr(9) * 256 + byteArr(8)
'            isValidImage = True
'        Else
'            GoTo checkBMP
'        End If
'
'        'get image type and exit
'        Ext = "gif"
'        GoTo endFunction
'
'checkBMP:
'
'        'check for BMP header
'        If byteArr(0) = 66 And byteArr(1) = 77 Then
'            isValidImage = True
'        Else
'            GoTo checkPNG
'        End If
'
'        'get record type info
'        If byteArr(14) = 40 Then
'
'            'get width and height of BMP
'            ImgDim.width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
'            + byteArr(19) * 256 + byteArr(18)
'
'            ImgDim.height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
'            + byteArr(23) * 256 + byteArr(22)
'
'        'another kind of BMP
'        ElseIf byteArr(17) = 12 Then
'
'            'get width and height of BMP
'            ImgDim.width = byteArr(19) * 256 + byteArr(18)
'            ImgDim.height = byteArr(21) * 256 + byteArr(20)
'
'        End If
'
'        'get image type and exit
'        Ext = "bmp"
'        GoTo endFunction
'
'checkPNG:
'
'        'check for PNG header
'        If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E And byteArr(3) = &H47 Then
'            ImgDim.width = byteArr(18) * 256 + byteArr(19)
'            ImgDim.height = byteArr(22) * 256 + byteArr(23)
'            isValidImage = True
'        Else
'            GoTo endFunction
'        End If
'
'        Ext = "png"
'
'    Else
'        AddMessage "Invalid picture file [" & FileName & "]"
'    End If
'endFunction:
'
'    'return function's success status
'    ImageDimensions = isValidImage
'
'End Function
'
'Public ColumnCount As Long
'Public RowCount As Long
'
'Public Fnt As StdFont
'Public MainFont As D3DXFont
'Public MainFontDesc As IFont
'
'Public Type ImgDimType
'  Height As Long
'  Width As Long
'End Type
'
'Public DefaultRenderTarget As Direct3DSurface8
'Public DefaultStencilDepth As Direct3DSurface8
'
'Public GenericMat As D3DMATERIAL8
'Public LucentMat As D3DMATERIAL8
'Public SpecialMat As D3DMATERIAL8
'

'
'Public Sub CleanupText()
'    Set DefaultRenderTarget = Nothing
'    Set DefaultStencilDepth = Nothing
'
'    Set MainFont = Nothing
'    Set MainFontDesc = Nothing
'    Set Fnt = Nothing
'End Sub
'


Public Function BitmapDimensions(ByVal FileName As String, imgdim As ImgDimType, ext As String) As Boolean

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
'
'Returns:
'
'True if the function was successful.

  'declare vars
  Dim handle As Integer, isValidImage As Boolean
  Dim byteArr(255) As Byte, i As Integer

  'init vars
  isValidImage = False
  imgdim.Height = 0
  imgdim.Width = 0
  
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
    If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 _
                         And byteArr(i + 1) <= &HCF Then
      imgdim.Height = byteArr(i + 5) * 256 + byteArr(i + 6)
      imgdim.Width = byteArr(i + 7) * 256 + byteArr(i + 8)
      Exit For
    End If
  Next i
  
  'get image type and exit
  ext = "jpg"
  GoTo endFunction

checkGIF:
  
  'check for GIF header
  If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 _
  And byteArr(3) = &H38 Then
    imgdim.Width = byteArr(7) * 256 + byteArr(6)
    imgdim.Height = byteArr(9) * 256 + byteArr(8)
    isValidImage = True
  Else
    GoTo checkBMP
  End If
  
  'get image type and exit
  ext = "gif"
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
    imgdim.Width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
                 + byteArr(19) * 256 + byteArr(18)
    
    imgdim.Height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
                  + byteArr(23) * 256 + byteArr(22)
  
  'another kind of BMP
  ElseIf byteArr(17) = 12 Then
  
    'get width and height of BMP
    imgdim.Width = byteArr(19) * 256 + byteArr(18)
    imgdim.Height = byteArr(21) * 256 + byteArr(20)
    
  End If
  
  'get image type and exit
  ext = "bmp"
  GoTo endFunction
  
checkPNG:

  'check for PNG header
  If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E _
  And byteArr(3) = &H47 Then
    imgdim.Width = byteArr(18) * 256 + byteArr(19)
    imgdim.Height = byteArr(22) * 256 + byteArr(23)
    isValidImage = True
  Else
    GoTo endFunction
  End If
  
  ext = "png"

endFunction:

  'return function's success status
  BitmapDimensions = isValidImage

End Function

Public Function LoadTexture(ByVal FileName As String) As Direct3DTexture8
    
    Dim t As String
    Dim Dimensions As ImgDimType

    If BitmapDimensions(FileName, Dimensions, t) Then
        Set LoadTexture = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, modDecs.Transparent, ByVal 0, ByVal 0)
    End If

End Function
Public Function LoadTextureEx(ByVal FileName As String, ByRef Dimensions As ImgDimType) As Direct3DTexture8

    Dim t As String

    If Dimensions.Width = 0 Or Dimensions.Height = 0 Then
        If BitmapDimensions(FileName, Dimensions, t) Then
            Set LoadTextureEx = D3DX.CreateTextureFromFileEx(DDevice, FileName, Dimensions.Width, Dimensions.Height, D3DX_FILTER_NONE, 0, _
                D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, modDecs.Transparent, ByVal 0, ByVal 0)
        End If
    Else
        Dim tmp As ImgDimType
        If BitmapDimensions(FileName, tmp, t) Then
            Set LoadTextureEx = D3DX.CreateTextureFromFileEx(DDevice, FileName, tmp.Width, tmp.Height, D3DX_FILTER_NONE, 0, _
                D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, modDecs.Transparent, ByVal 0, ByVal 0)
        End If
    End If

End Function

'Public Sub DelImageIndex(ByVal Index As Long)
'    If (ImageCount > 0) And (Index > (ImageCount - 1)) Then
'        Dim cnt As Long
'        For cnt = Index To ImageCount - 1
'            Images(cnt) = Images(cnt + 1)
'        Next
'        ImageCount = ImageCount - 1
'        ReDim Preserve Images(1 To ImageCount) As MyImage
'    End If
'End Sub

Public Function GetIndexFile(ByVal ID As Long) As String
    If FileCount > 0 And ID < FileCount Then
        GetIndexFile = LCase(Trim(Files(ID).path))
    End If
End Function

Public Function GetFileIndex(Optional ByVal ID As String) As Long
    If ID = "" Then Exit Function
    Dim cnt As Long
    Dim idx As Long
    If FileCount > 0 Then
        For cnt = 1 To FileCount
            If LCase(Trim(Files(cnt).path)) = LCase(Trim(ID)) Then
                GetFileIndex = cnt
                Exit Function
            ElseIf Files(cnt).path = "" And idx = 0 Then
                idx = cnt
            End If
        Next
        If idx > 0 Then
            GetFileIndex = idx
            Files(idx).path = ID
            Exit Function
        End If
    End If
    FileCount = FileCount + 1
    ReDim Preserve Files(1 To FileCount) As MyFile
    Files(FileCount).path = ID
    GetFileIndex = FileCount
End Function

'Public Function GetBillboardByFile(Optional ByVal ID As String) As Object
'    Dim cnt As Long
'    Dim e As Billboard
'    If FileCount > 0 And Billboards.Count > 0 Then
'        For cnt = 1 To FileCount
'            If LCase(Trim(Files(cnt).path)) = LCase(Trim(ID)) Then
'                For Each e In Billboards
'                    If e.FaceIndex = cnt Then
'                        Set GetBillboardByFile = e
'                        Exit Function
'                    End If
'                Next
'
'            End If
'        Next
'    End If
'End Function


Public Function LoadTextureRes(ByRef byteArr() As Byte) As Direct3DTexture8
    Dim imgdim As ImgDimType

        'check for BMP header
    If byteArr(0) = 66 And byteArr(1) = 77 Then

    
        'get record type info
        If byteArr(14) = 40 Then
        
            'get width and height of BMP
            imgdim.Width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
            + byteArr(19) * 256 + byteArr(18)
            
            imgdim.Height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
            + byteArr(23) * 256 + byteArr(22)
        
        'another kind of BMP
        ElseIf byteArr(17) = 12 Then
        
            'get width and height of BMP
            imgdim.Width = byteArr(19) * 256 + byteArr(18)
            imgdim.Height = byteArr(21) * 256 + byteArr(20)
            
        End If

        Set LoadTextureRes = D3DX.CreateTextureFromFileInMemoryEx(DDevice, byteArr(0), UBound(byteArr) + 1, imgdim.Width, imgdim.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, modDecs.Transparent, ByVal 0, ByVal 0)
    
    End If
End Function







