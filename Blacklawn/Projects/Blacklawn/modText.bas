#Const [True] = -1
#Const [False] = 0

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

Public GenericMaterial As D3DMATERIAL8

Public Const LOGPIXELSX = 88
Private Const POINTS_PER_INCH As Long = 72

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Function PixelPerPoint() As Double
    Dim hDC As Long
    Dim lPixelPerInch As Long
    hDC = GetDC(ByVal 0&)
    lPixelPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelPerPoint = POINTS_PER_INCH / lPixelPerInch
    ReleaseDC ByVal 0&, hDC
    PixelPerPoint = 1 + PixelPerPoint
End Function

Public Sub CreateText()
    ColumnCount = 129
    
    frmMain.Font.name = "Lucida Console"
    frmMain.Font.Bold = False
    frmMain.Font.Italic = False
    frmMain.Font.CharSet = 0
    
    Dim Size As Long
    
    Size = 30
    frmMain.Font.Size = Size
    Do Until frmMain.TextWidth(String(ColumnCount, "A")) < (frmMain.ScaleWidth - (TextSpace * 2))
        Size = Size - 1
        frmMain.Font.Size = Size
    Loop

    RowCount = 1
    Do Until ((((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * RowCount) + 2) >= ((frmMain.ScaleHeight - TextHeight) / Screen.TwipsPerPixelY)
        RowCount = RowCount + 1
    Loop
    
    Set Fnt = frmMain.Font
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(DDevice, MainFontDesc.hFont)
    
    GenericMaterial.Ambient.a = 0.1
    GenericMaterial.Ambient.r = 1
    GenericMaterial.Ambient.G = 1
    GenericMaterial.Ambient.b = 1
    GenericMaterial.diffuse.a = 1
    GenericMaterial.diffuse.r = 1
    GenericMaterial.diffuse.G = 1
    GenericMaterial.diffuse.b = 1

    GenericMaterial.power = 1
    
End Sub

Public Sub CleanupText()
    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set Fnt = Nothing
End Sub

Public Function DrawText(Text As String, X As Long, Y As Long)
    
    Dim TextRect As RECT
    Dim Color As Long
    Dim Allignment As CONST_DTFLAGS
    Color = &HFFFFFFFF
    Allignment = DT_TOP Or DT_LEFT
    
    TextRect.Top = Y
    TextRect.Left = X
    TextRect.Bottom = Y + (frmMain.TextHeight(Text) / Screen.TwipsPerPixelY)
    TextRect.Right = X + (frmMain.TextWidth(Text) / Screen.TwipsPerPixelX)
    
    D3DX.DrawText MainFont, &HFFFFFFFF, Text, TextRect, Allignment
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
    Dim t As String

    If ImageDimensions(FileName, d, t) Then
        Set LoadTexture = D3DX.CreateTextureFromFileEx(DDevice, FileName, d.width, d.height, D3DX_FILTER_NONE, 0, _
            D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0)
    End If
End Function
