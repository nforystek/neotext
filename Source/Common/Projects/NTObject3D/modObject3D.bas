Attribute VB_Name = "modObject3D"
Option Explicit


'Angle      ^
'Axis       !
'Bound      <></>
'Field      $
'Line       |
'Matter     =
'Neutron    '
'Orbit      @
'Protron    `
'Point      [,,]
'Range      ?
'Shape      #
'Space      %
'Spirit     ;
'Square     [,]
'Stars      *
'Vector     ~
'Vision     ?
'Volume     &

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

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

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dW As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function ArcTo Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hbrush As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public OrbitCount As Long
Public SpaceCount As Long
Public StarsCount As Long
Public VolumeCount As Long

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

Public Static Function HookObj(ByRef Obj)
    
    Static hc As Collection
    Static ha As Collection
    If IsNumeric(Obj) Then
        If hc Is Nothing And Obj > 0 Then
            DestroyWindow Obj
        Else
            If Obj < 0 Then
                HookObj = ha("k" & -Obj)
            Else
                Set HookObj = hc("k" & Obj)
            End If
        End If
    Else
        If hc Is Nothing Then
            Set hc = New Collection
            Set ha = New Collection
        End If
        Dim cnt As Long
        If hc.Count > 0 Then
            For cnt = 1 To hc.Count
                If hc(cnt).hwnd = Obj.hwnd Then
                    SetWindowLong Obj.hwnd, _
                    GWL_WNDPROC, ha("k" & Obj.hwnd)
                    hc.Remove "k" & Obj.hwnd
                    ha.Remove "k" & Obj.hwnd
                    GoTo hookok
                End If
            Next
        End If
        hc.Add Obj, "k" & Obj.hwnd
        ha.Add SetWindowLong(Obj.hwnd, GWL_WNDPROC, _
        AddressOf ControlWndProc), "k" & Obj.hwnd
    End If
hookok:
    If hc.Count = 0 Then
        Set hc = Nothing
        Set ha = Nothing
    End If
End Function

Private Function ControlWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (HookObj(-hwnd) <> 0) Then
        Dim R As Field
        Set R = HookObj(hwnd)
        ControlWndProc = R.ControlWndProc(hwnd, uMsg, wParam, lParam)
        Set R = Nothing
        If ControlWndProc = 0 Then
            If CallWindowProc(HookObj(-hwnd), hwnd, uMsg, wParam, lParam) = 0 Then
                ControlWndProc = 1
            Else
                ControlWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
            End If
        End If
    End If
End Function
