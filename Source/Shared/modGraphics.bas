Attribute VB_Name = "modGraphics"
#Const modGraphics = -1
Option Explicit
'TOP DOWN

Option Private Module


Public Type POINTAPI
        X As Long
        Y As Long
End Type
#If Not modBitBlt = -1 Then
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
#End If
Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type


Public Type POLYTEXT
        X As Long
        Y As Long
        n As Long
        lpStr As String
        uiFlags As Long
        rcl As RECT
        pdx As Long
End Type
Public Enum PEN_STYLE
    PS_ALTERNATE = 8
    PS_COSMETIC = &H0
    PS_DASH = 1
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_DOT = 2
    PS_ENDCAP_FLAT = &H200
    PS_ENDCAP_MASK = &HF00
    PS_ENDCAP_ROUND = &H0
    PS_ENDCAP_SQUARE = &H100
    PS_GEOMETRIC = &H10000
    PS_INSIDEFRAME = 6
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MASK = &HF000
    PS_JOIN_MITER = &H2000
    PS_JOIN_ROUND = &H0
    PS_NULL = 5
    PS_SOLID = 0
    PS_STYLE_MASK = &HF
    PS_TYPE_MASK = &HF0000
    PS_USERSTYLE = 7
End Enum
Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type
' *** BRUSHES ***
Public Enum BRUSH_STYLE
    BS_DIBPATTERN = 5
    BS_DIBPATTERN8X8 = 8
    BS_DIBPATTERNPT = 6
    BS_NULL = 1
    BS_HOLLOW = BS_NULL
    BS_HATCHED = 2
    BS_INDEXED = 4
    BS_PATTERN = 3
    BS_PATTERN8X8 = 7
    BS_SOLID = 0
End Enum
Public Enum HATCH_STYLE
    HS_BDIAGONAL = 3
    HS_BDIAGONAL1 = 7
    HS_CROSS = 4
    HS_DENSE1 = 9
    HS_DENSE2 = 10
    HS_DENSE3 = 11
    HS_DENSE4 = 12
    HS_DENSE5 = 13
    HS_DENSE6 = 14
    HS_DENSE7 = 15
    HS_DENSE8 = 16
    HS_DIAGCROSS = 5
    HS_DITHEREDBKCLR = 24
    HS_DITHEREDCLR = 20
    HS_DITHEREDTEXTCLR = 22
    HS_FDIAGONAL = 2
    HS_FDIAGONAL1 = 6
    HS_HALFTONE = 18
    HS_HORIZONTAL = 0
    HS_NOSHADE = 17
    HS_SOLID = 8
    HS_SOLIDBKCLR = 23
    HS_SOLIDCLR = 19
    HS_SOLIDTEXTCLR = 21
    HS_VERTICAL = 1
End Enum
' brush properties
Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To 32) As Byte
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Const FW_NORMAL = 400
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const Transparent = 1
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal Op As Long, ByVal cP As Long, ByVal q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_REGULAR = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_DEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_ULTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = 900
Public Const ANSI_CHARSET = 0
Public Const ARABIC_CHARSET = 178
Public Const BALTIC_CHARSET = 186
Public Const CHINESEBIG5_CHARSET = 136
Public Const DEFAULT_CHARSET = 1
Public Const EASTEUROPE_CHARSET = 238
Public Const GB2312_CHARSET = 134
Public Const GREEK_CHARSET = 161
Public Const HANGEUL_CHARSET = 129
Public Const HEBREW_CHARSET = 177
Public Const MAC_CHARSET = 77
Public Const OEM_CHARSET = 255
Public Const SHIFTJIS_CHARSET = 128
Public Const SYMBOL_CHARSET = 2
Public Const TURKISH_CHARSET = 162
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_STROKE_PRECIS = 2
Public Const ANTIALIASED_QUALITY = 4
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const NONANTIALIASED_QUALITY = 3
Public Const PROOF_QUALITY = 2
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2
Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_MODERN = 48
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32

'Public Const DFC_CAPTION = 1            'Title bar
'Public Const DFC_MENU = 2               'Menu
'Public Const DFC_SCROLL = 3             'Scroll bar
'Public Const DFC_BUTTON = 4             'Standard button
'
'Public Const DFCS_CAPTIONCLOSE = &H0    'Close button
'Public Const DFCS_CAPTIONMIN = &H1      'Minimize button
'Public Const DFCS_CAPTIONMAX = &H2      'Maximize button
'Public Const DFCS_CAPTIONRESTORE = &H3  'Restore button
'Public Const DFCS_CAPTIONHELP = &H4     'Windows 95 only:
'                                        'Help button
'
'Public Const DFCS_MENUARROW = &H0       'Submenu arrow
'Public Const DFCS_MENUCHECK = &H1       'Check mark
'Public Const DFCS_MENUBULLET = &H2      'Bullet
'Public Const DFCS_MENUARROWRIGHT = &H4
'
'Public Const DFCS_SCROLLUP = &H0               'Up arrow of scroll
'                                               'bar
'Public Const DFCS_SCROLLDOWN = &H1             'Down arrow of
'                                               'scroll bar
'Public Const DFCS_SCROLLLEFT = &H2             'Left arrow of
'                                               'scroll bar
'Public Const DFCS_SCROLLRIGHT = &H3            'Right arrow of
'                                               'scroll bar
'Public Const DFCS_SCROLLCOMBOBOX = &H5         'Combo box scroll
'                                               'bar
'Public Const DFCS_SCROLLSIZEGRIP = &H8         'Size grip
'Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10   'Size grip in
'                                               'bottom-right
'                                               'corner of window
'
'Public Const DFCS_BUTTONCHECK = &H0      'Check box
'
'Public Const DFCS_BUTTONRADIO = &H4     'Radio button
'Public Const DFCS_BUTTON3STATE = &H8    'Three-state button
'Public Const DFCS_BUTTONPUSH = &H10     'Push button
'
'Public Const DFCS_INACTIVE = &H100      'Button is inactive
'                                        '(grayed)
'Public Const DFCS_PUSHED = &H200        'Button is pushed
'Public Const DFCS_CHECKED = &H400       'Button is checked
'
'Public Const DFCS_ADJUSTRECT = &H2000   'Bounding rectangle is
'                                        'adjusted to exclude the
'                                        'surrounding edge of the
'                                        'push button
'
'Public Const DFCS_FLAT = &H4000         'Button has a flat border
'Public Const DFCS_MONO = &H8000         'Button has a monochrome
'                                        'border

Public Declare Function DrawFrameControl Lib "user32" (ByVal _
   hdc&, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) _
   As Boolean
   
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dW As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

#If Not modBitBlt Then
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
#End If
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function lineto Lib "gdi32" Alias "LineTo" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbrush As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function TextOutPtr Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
'Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
'Public Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
'Public Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
'Public Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
'Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Public Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
'Public Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
'Public Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
'Public Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As Long, pptxt As POLYTEXT, cStrings As Long) As Long
'Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal Blend As Long) As Long

'Public Declare Function BitBlt Lib "gdi32" ( _
'   ByVal hDCDest As Long, ByVal XDest As Long, _
'   ByVal YDest As Long, ByVal nWidth As Long, _
'   ByVal nHeight As Long, ByVal hDCSrc As Long, _
'   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
'   As Long

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE


Public Type ImageDimensions
  Height As Long
  Width As Long
End Type

Public Const PixelsPerInchX As Single = 90 ' Logical pixels/inch in X
Public Const PixelsPerInchY As Single = 90 ' Logical pixels/inch in Y

Public Const PixelPerPointX As Single = 1.76630434782609
Public Const PixelPerPointY As Single = 1.7578125

Public Const PointsPerInchX As Single = 52.0861538461537 ' 1.625
Public Const PointsPerInchY As Single = 54.6133333333333 ' 1.6875

Public Const InchesPerPointX As Single = 0.017663043478261
Public Const InchesPerPointY As Single = 0.017578125

Public Const PointPerPixelX As Single = 0.566153846153845
Public Const PointPerPixelY As Single = 0.568888888888889

Public Const ColorCount8bitDepth As Long = 256
Public Const ColorCount16bitDepth As Long = 65536
Public Const ColorCount24bitDepth As Long = 16777216

'  Device Parameters for GetDeviceCaps()
Private Const DRIVERVERSION = 0      '  Device driver version
Private Const TECHNOLOGY = 2         '  Device classification
Private Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const VERTSIZE = 6           '  Vertical size in millimeters
Private Const HORZRES = 8            '  Horizontal width in pixels
Private Const VERTRES = 10           '  Vertical width in pixels
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const Plane = 14            '  Number of Plane
Private Const NUMBRUSHES = 16        '  Number of brushes the device has
Private Const NUMPENS = 18           '  Number of pens the device has
Private Const NUMMARKERS = 20        '  Number of markers the device has
Private Const NUMFONTS = 22          '  Number of fonts the device has
Private Const NUMCOLORS = 24         '  Number of colors the device supports
Private Const PDEVICESIZE = 26       '  Size required for device descriptor
Private Const CURVECAPS = 28         '  Curve capabilities
Private Const LINECAPS = 30          '  Line capabilities
Private Const POLYGONALCAPS = 32     '  Polygonal capabilities
Private Const TEXTCAPS = 34          '  Text capabilities
Private Const CLIPCAPS = 36          '  Clipping capabilities
Private Const RASTERCAPS = 38        '  Bitblt capabilities
Private Const ASPECTX = 40           '  Length of the X leg
Private Const ASPECTY = 42           '  Length of the Y leg
Private Const ASPECTXY = 44          '  Length of the hypotenuse

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SIZEPALETTE = 104      '  Number of entries in physical palette
Private Const NUMRESERVED = 106      '  Number of reserved entries in palette
Private Const COLORRES = 108         '  Actual color resolution

'  Printing related DeviceCaps. These replace the appropriate Escapes
Private Const PHYSICALWIDTH = 110 '  Physical Width in device units
Private Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Private Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Private Const SCALINGFACTORX = 114 '  Scaling factor x
Private Const SCALINGFACTORY = 115 '  Scaling factor y



Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
'Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type Bitmap
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlane As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
   GUID As GUID
   NumberOfValues As Long
   Type As Long
   Value As Long
End Type

Private Type EncoderParameters
   Count As Long
   Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
   
Public Const COLOR_SCROLLBAR = 0           ' Scroll Bar
Private Const COLOR_BACKGROUND = 1          ' Windows desktop
Private Const COLOR_ACTIVECAPTION = 2       ' Caption of active window
Private Const COLOR_INACTIVECAPTION = 3     ' Caption of inactive window
Private Const COLOR_MENU = 4                ' Menu
Public Const COLOR_WINDOW = 5              ' Window background
Private Const COLOR_WINDOWFRAME = 6         ' Window frame
Private Const COLOR_MENUTEXT = 7            ' Menu text
Public Const COLOR_WINDOWTEXT = 8          ' Window text
Private Const COLOR_CAPTIONTEXT = 9         ' Text in window caption
Private Const COLOR_ACTIVEBORDER = 10       ' Border of active window
Private Const COLOR_INACTIVEBORDER = 11     ' Border of inactive window
Private Const COLOR_APPWORKSPACE = 12       ' Background of MDI desktop
Public Const COLOR_HIGHLIGHT = 13          ' Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14      ' Selected item text
Private Const COLOR_BTNFACE = 15            ' Button
Private Const COLOR_BTNSHADOW = 16          ' 3D shading of button
Public Const COLOR_GRAYTEXT = 17           ' Gray text, or zero if dithering is used
Private Const COLOR_BTNTEXT = 18            ' Button text
Private Const COLOR_INACTIVECAPTIONTEXT = 19    ' Text of inactive window
Public Const COLOR_BTNHIGHLIGHT = 20       ' 3D highlight of button

Private Const COLOR_3DDKSHADOW = 21         ' 3D dark shadow (Win95)
Private Const COLOR_3DLIGHT = 22            ' Light color for 3D shaded objects (Win95)
Private Const COLOR_INFOTEXT = 23           ' Tooltip text color (Win95)
Private Const COLOR_INFOBK = 24             ' Tooltip background color (Win95)

Private Const COLOR_HOTLIGHT = 26           ' Color for a hot-tracked item
Private Const COLOR_GRADIENTACTIVECAPTION = 27    ' Right side color in the color gradient of an active window's title bar.
Private Const COLOR_GRADIENTINACTIVECAPTION = 28    ' Right side color in the color gradient of an inactive window's title bar.

Private Const COLOR_DESKTOP = COLOR_BACKGROUND
Private Const COLOR_3DFACE = COLOR_BTNFACE  ' Face color for 3D shaded objects (Win95)
Private Const COLOR_3DSHADOW = COLOR_BTNSHADOW
Private Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT  ' 3D Highlight color (Win95)
Private Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT


Private Declare Function SetSysColors _
 Lib "user32" _
 (ByVal nChanges As Long, lpSysColor As Long, _
 lpColorValues As Long) As Long
 
Public Declare Function GetSysColor _
 Lib "user32" (ByVal nIndex As Long) As Long

Public Function RECT(Left, Top, Right, Bottom) As RECT
    With RECT
        .Left = CLng(Left)
        .Top = CLng(Top)
        .Right = CLng(Right)
        .Bottom = CLng(Bottom)
    End With
End Function

Public Function ConvertColor(ByVal color As Variant, Optional ByRef red As Long, Optional ByRef green As Long, Optional ByRef blue As Long) As Long
On Error GoTo catch
    Dim lngColor As Long
    If InStr(CStr(color), "#") > 0 Then
        GoTo HTMLorHexColor
    ElseIf InStr(CStr(color), "&H") > 0 Then
        GoTo SysOrLongColor
    ElseIf IsAlphaNumeric(color) Then
        If (Not (Len(color) = 6)) And (Not Left(color, 1) = "0") Then
            GoTo SysOrLongColor
        Else
            GoTo HTMLorHexColor
        End If
    End If
SysOrLongColor:
    lngColor = CLng(color)
    If Not (lngColor >= 0 And lngColor <= 16777215) Then 'if system colour
        Select Case lngColor
            Case SystemColorConstants.vbScrollBars
                lngColor = COLOR_SCROLLBAR          ' Scroll Bar
            Case SystemColorConstants.vbDesktop
                lngColor = COLOR_BACKGROUND        ' Windows desktop
            Case SystemColorConstants.vbActiveTitleBar
                lngColor = COLOR_ACTIVECAPTION       ' Caption of active window
            Case SystemColorConstants.vbInactiveTitleBar
                lngColor = COLOR_INACTIVECAPTION     ' Caption of inactive window
            Case SystemColorConstants.vbMenuBar
                lngColor = COLOR_MENU                 ' Menu
            Case SystemColorConstants.vbWindowBackground
                lngColor = COLOR_WINDOW               ' Window background
            Case SystemColorConstants.vbWindowFrame
                lngColor = COLOR_WINDOWFRAME         ' Window frame
            Case SystemColorConstants.vbMenuText
                lngColor = COLOR_MENUTEXT            ' Menu text
            Case SystemColorConstants.vbWindowText
                lngColor = COLOR_WINDOWTEXT           ' Window text
            Case SystemColorConstants.vbTitleBarText
                lngColor = COLOR_CAPTIONTEXT         ' Text in window caption
            Case SystemColorConstants.vbActiveBorder
                lngColor = COLOR_ACTIVEBORDER       ' Border of active window
            Case SystemColorConstants.vbInactiveBorder
                lngColor = COLOR_INACTIVEBORDER     ' Border of inactive window
            Case SystemColorConstants.vbApplicationWorkspace
                lngColor = COLOR_APPWORKSPACE       ' Background of MDI desktop
            Case SystemColorConstants.vbHighlight
                lngColor = COLOR_HIGHLIGHT          ' Selected item background
            Case SystemColorConstants.vbHighlightText
                lngColor = COLOR_HIGHLIGHTTEXT      ' Selected item text
            Case SystemColorConstants.vbButtonFace
                lngColor = COLOR_BTNFACE             ' Button
            Case SystemColorConstants.vbButtonShadow
                lngColor = COLOR_BTNSHADOW          ' 3D shading of button
            Case SystemColorConstants.vbGrayText
                lngColor = COLOR_GRAYTEXT           ' Gray text, or zero if dithering is used
            Case SystemColorConstants.vbButtonText
                lngColor = COLOR_BTNTEXT             ' Button text
            Case SystemColorConstants.vbInactiveCaptionText
                lngColor = COLOR_INACTIVECAPTIONTEXT     ' Text of inactive window
            Case SystemColorConstants.vb3DHighlight
                lngColor = COLOR_BTNHIGHLIGHT        ' 3D highlight of button
        End Select

'        lngColor = lngColor And Not &H80000000
        lngColor = GetSysColor(lngColor)
'
'    Else
'
    End If
    color = Right("000000" & Hex(lngColor), 6)
HTMLorHexColor:
    red = CByte("&h" & Mid(color, 5, 2))
    green = CByte("&h" & Mid(color, 3, 2))
    blue = CByte("&h" & Mid(color, 1, 2))
    
    ConvertColor = RGB(red, green, blue)
    If ConvertColor <> lngColor Then
        Err.Raise 8, "Exception."
    End If
    Exit Function

'    green = Val("&H" & Right(color, 2))
'    red = Val("&H" & Mid(color, 2, 2))
'    blue = Val("&H" & Mid(color, 4, 2))
'    ConvertColor = RGB(red, green, blue)
'    Exit Function
catch:
    Err.Clear
    ConvertColor = 0
End Function



'Public Function PixelPerPoint() As Double
'    Dim hdc As Long
'    Dim lPixelPerInch As Long
'    hdc = GetDC(ByVal 0&)
'    lPixelPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
'    PixelPerPoint = POINTS_PER_INCH / lPixelPerInch
'    ReleaseDC ByVal 0&, hdc
'    PixelPerPoint = 1 + PixelPerPoint
'End Function

Public Function PixelPerPoint() As Double
    Dim hdc As Long
    Dim lPixelPerInch As Long
    hdc = GetDC(ByVal 0&)
    lPixelPerInch = GetDeviceCaps(hdc, ASPECTXY)
   ' PixelPerPoint = 72 / lPixelPerInch
    ReleaseDC ByVal 0&, hdc
    PixelPerPoint = lPixelPerInch
End Function


' ----==== SaveJPG ====----

Public Sub SaveJPG(ByVal pict As StdPicture, ByVal FileName As String, Optional ByVal Quality As Byte = 80)

Dim tSI As GdiplusStartupInput
Dim lRes As Long
Dim lGDIP As Long
Dim lBitmap As Long

   ' Initialize GDI+
   tSI.GdiplusVersion = 1
   lRes = GdiplusStartup(lGDIP, tSI)

   If lRes = 0 Then

      ' Create the GDI+ bitmap
      ' from the image handle
      lRes = GdipCreateBitmapFromHBITMAP(pict.handle, pict.hPal, lBitmap)

      If lRes = 0 Then
         Dim tJpgEncoder As GUID
         Dim tParams As EncoderParameters

         ' Initialize the encoder GUID
         CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder

         ' Initialize the encoder parameters
         tParams.Count = 1
         With tParams.Parameter ' Quality
            ' Set the Quality GUID
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
            .NumberOfValues = 1
            .Type = 4
            .Value = VarPtr(Quality)
         End With

         ' Save the image
         lRes = GdipSaveImageToFile( _
                  lBitmap, _
                  StrPtr(FileName), _
                  tJpgEncoder, _
                  tParams)

         ' Destroy the bitmap
         GdipDisposeImage lBitmap

      End If

      ' Shutdown GDI+
      GdiplusShutdown lGDIP

   End If

   If lRes Then
      Err.Raise 5, , "Cannot save the image. GDI+ Error:" & lRes
   End If

End Sub
Public Sub WriteBytes(ByVal FileName As String, ByRef C() As Byte)
    Dim FileNo As Integer
    On Error GoTo Err_Init
    If PathExists(FileName, True) Then Kill FileName
    FileNo = FreeFile
    Open FileName For Output As #FileNo
    Close #FileNo
    FileNo = FreeFile
    Open FileName For Binary Access Write As #FileNo

    Put #FileNo, , C
    Close #FileNo
    
    Exit Sub
Err_Init:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Function LoadFile(ByVal FileName As String) As Byte()
    Dim FileNo As Integer, b() As Byte
    On Error GoTo Err_Init
    If Dir(FileName, vbNormal Or vbArchive) = "" Then
        Exit Function
    End If
    FileNo = FreeFile
    Open FileName For Binary Access Read As #FileNo
    ReDim b(0 To LOF(FileNo) - 1)
    Get #FileNo, , b
    Close #FileNo
    LoadFile = b
    Exit Function
Err_Init:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function PictureFromByteStream(b() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then

                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "You must pass a non-empty byte array to this function!"
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function

Public Function GetMonitorDPI(Optional ByVal LogicalPixelsX As Long = 88, Optional ByVal LogicalPixelsY As Long = 90) As ImageDimensions
    
    Dim hdc As Long
    Dim lngRetVal As Long
    
    hdc = GetDC(0)
    
    GetMonitorDPI.Width = GetDeviceCaps(hdc, LogicalPixelsX)
    GetMonitorDPI.Height = GetDeviceCaps(hdc, LogicalPixelsY)
    
    lngRetVal = ReleaseDC(0, hdc)

End Function

Public Function ImageDimensions(ByVal FileName As String, ByRef imgdim As ImageDimensions, Optional ByRef ext As String = "") As Boolean

    If PathExists(FileName, True) Then

        'declare vars
        Dim handle As Integer
        Dim byteArr(255) As Byte
        
        'open file and get 256 byte chunk
        handle = FreeFile
        On Error GoTo endFunction
        Open FileName For Binary Access Read As #handle
            Get handle, , byteArr
        Close #handle
        
        ImageDimensions = ImageDimensionsFromBytes(byteArr, imgdim, ext)

    
    Else
        Debug.Print "Invalid picture file [" & FileName & "]"
    End If
endFunction:


End Function

Public Function ImageDimensionsFromBytes(ByRef byteArr() As Byte, ByRef imgdim As ImageDimensions, Optional ByRef ext As String = "") As Boolean

    Dim isValidImage As Boolean
    Dim i As Integer
    
    'init vars
    isValidImage = False
    imgdim.Height = 0
    imgdim.Width = 0

    
    'check for jpg header (SOI): &HFF and &HD8
    ' contained in first 2 bytes
    If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
        isValidImage = True
    Else
        GoTo checkGIF
    End If
    
    'check for SOF marker: &HFF and &HC0 TO &HCF
    For i = 0 To UBound(byteArr) - 1
        If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 And byteArr(i + 1) <= &HCF Then
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
    If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 And byteArr(3) = &H38 Then
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
    If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E And byteArr(3) = &H47 Then
        imgdim.Width = byteArr(18) * 256 + byteArr(19)
        imgdim.Height = byteArr(22) * 256 + byteArr(23)
        isValidImage = True
    Else
        GoTo endFunction
    End If
    
    ext = "png"


endFunction:
    
    'return function's success status
    ImageDimensionsFromBytes = isValidImage

End Function


#If Not modCommon Then

Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim C2 As Integer
    Dim retval As Boolean
    retval = True
    If Not IsNumeric(Text) Then
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) = 46) Then
                C2 = C2 + 1
            ElseIf (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                retval = False
                Exit For
            End If
        Next
    Else
        retval = False
    End If
    Else
        retval = True
    End If
    IsAlphaNumeric = retval And (C2 <= 1)
End Function

#End If






