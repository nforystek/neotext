Attribute VB_Name = "modText"


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

Public DefaultRenderTarget As Direct3DSurface8
Public DefaultStencilDepth As Direct3DSurface8

Public GenericMat As D3DMATERIAL8
Public LucentMat As D3DMATERIAL8
Public SpecialMat As D3DMATERIAL8

Public Sub CreateText()
    ColumnCount = 129
    
    frmMain.Font.name = "Lucida Console"
    frmMain.Font.Bold = False
    frmMain.Font.Italic = False
    frmMain.Font.CharSet = 0
    
    Dim Size As Long
    
    Size = 30
    frmMain.Font.Size = Size
    Do Until frmMain.TextWidth(String(ColumnCount, "A")) < (Screen.Width - (TextSpace * 2))
        Size = Size - 1
        frmMain.Font.Size = Size
    Loop

    RowCount = 1
    Do Until ((((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((Screen.Height - TextHeight) / Screen.TwipsPerPixelY)
        RowCount = RowCount + 1
    Loop
    
    Set Fnt = frmMain.Font
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(DDevice, MainFontDesc.hFont)
    
    GenericMat.Ambient.a = 1
    GenericMat.Ambient.R = 1
    GenericMat.Ambient.G = 1
    GenericMat.Ambient.B = 1
    GenericMat.diffuse.a = 1
    GenericMat.diffuse.R = 1
    GenericMat.diffuse.G = 1
    GenericMat.diffuse.B = 1
    GenericMat.power = 1
  
    LucentMat.Ambient.a = 1
    LucentMat.Ambient.R = 1
    LucentMat.Ambient.G = 1
    LucentMat.Ambient.B = 1
    LucentMat.diffuse.a = 1
    LucentMat.diffuse.R = 0
    LucentMat.diffuse.G = 0
    LucentMat.diffuse.B = 0
    LucentMat.power = 1
    
    SpecialMat.Ambient.a = 0
    SpecialMat.Ambient.R = 0.89
    SpecialMat.Ambient.G = 0.89
    SpecialMat.Ambient.B = 0.89
    SpecialMat.diffuse.a = 0.4
    SpecialMat.diffuse.R = 0.01
    SpecialMat.diffuse.G = 0.01
    SpecialMat.diffuse.B = 0.01
    SpecialMat.specular.a = 0
    SpecialMat.specular.R = 0.5
    SpecialMat.specular.G = 0.5
    SpecialMat.specular.B = 0.5
    SpecialMat.emissive.a = 0.3
    SpecialMat.emissive.R = 0.21
    SpecialMat.emissive.G = 0.3
    SpecialMat.emissive.B = 0.3
    SpecialMat.power = 0
      
'    Set DefaultRenderTarget = DDevice.GetRenderTarget
'    Set DefaultStencilDepth = DDevice.GetDepthStencilSurface
'
End Sub

Public Sub CleanupText()
    Set DefaultRenderTarget = Nothing
    Set DefaultStencilDepth = Nothing
    
    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set Fnt = Nothing
End Sub

Public Function DrawTextByRowCol(Text As String, X As Long, Y As Long)

    Dim rec As RECT

    GetWindowRect frmMain.hwnd, rec
    Dim screenX As Single
    Dim screenY As Single
    screenX = (rec.Left - rec.Right) / Screen.Width
    screenY = (rec.Bottom - rec.Top) / Screen.Height
'        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)

        
    Dim TextRect As dxvbliba.RECT
    Dim Color As Long
    Dim Allignment As CONST_DTFLAGS
    Color = &HFFFFFFFF
    Allignment = DT_TOP Or DT_LEFT
    
    TextRect.Top = Y
    TextRect.Left = X
    TextRect.Bottom = (Y + (frmMain.TextHeight(Text) / Screen.TwipsPerPixelY))
    TextRect.Right = X + (frmMain.TextWidth(Text) / Screen.TwipsPerPixelX)

    MainFont.Begin
    
    D3DX.DrawText MainFont, &HFFFFFFFF, Text, TextRect, Allignment
    MainFont.End
    
End Function

Public Function DrawTextByCoord(Text As String, X As Single, Y As Single)

'    Dim rec As RECT
'
'    GetWindowRect frmMain.hWnd, rec
'    Dim screenX As Single
'    Dim screenY As Single
'    screenX = (rec.Left - rec.Right) / Screen.Width
'    screenY = (rec.Bottom - rec.Top) / Screen.Height
'        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)

        
    Dim TextRect As dxvbliba.RECT
    Dim Color As Long
    Dim Allignment As CONST_DTFLAGS
    Color = &HFFFFFFFF
    Allignment = DT_TOP Or DT_LEFT
    
    TextRect.Top = Y
    TextRect.Left = X
    TextRect.Bottom = (Y + (frmMain.TextHeight(Text) / Screen.TwipsPerPixelY))
    TextRect.Right = (X + (frmMain.TextWidth(Text) / Screen.TwipsPerPixelX))

    
    D3DX.DrawText MainFont, Color, Text, TextRect, Allignment

    
End Function

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
  Dim Handle As Integer, isValidImage As Boolean
  Dim byteArr(255) As Byte, i As Integer

  'init vars
  isValidImage = False
  imgdim.Height = 0
  imgdim.Width = 0
  
  'open file and get 256 byte chunk
  Handle = FreeFile
  On Error GoTo endFunction
  Open FileName For Binary Access Read As #Handle
  Get Handle, , byteArr
  Close #Handle

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

Public Function LoadTexture(ByVal FileName As String, Optional ByRef Height As Single, Optional ByRef Width As Single) As Direct3DTexture8
    Dim d As ImgDimType
    Dim t As String

    If BitmapDimensions(FileName, d, t) Then
        Set LoadTexture = D3DX.CreateTextureFromFileEx(DDevice, FileName, d.Width, d.Height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
    End If
End Function

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
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
    
    End If
End Function





