VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Backbuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FrameTypes

    DFC_CAPTION = 1            'Title bar
    DFC_MENU = 2               'Menu
    DFC_SCROLL = 3             'Scroll bar
    DFC_BUTTON = 4             'Standard button

End Enum

Public Enum FrameStyles
    DFCS_CAPTIONCLOSE = &H0    'Close button
    DFCS_CAPTIONMIN = &H1      'Minimize button
    DFCS_CAPTIONMAX = &H2      'Maximize button
    DFCS_CAPTIONRESTORE = &H3  'Restore button
    DFCS_CAPTIONHELP = &H4     'Windows 95 only:
                                           'Help button
    
    DFCS_MENUARROW = &H0       'Submenu arrow
    DFCS_MENUCHECK = &H1       'Check mark
    DFCS_MENUBULLET = &H2      'Bullet
    DFCS_MENUARROWRIGHT = &H4
    
    DFCS_SCROLLUP = &H0               'Up arrow of scroll
                                                  'bar
    DFCS_SCROLLDOWN = &H1             'Down arrow of
                                                  'scroll bar
    DFCS_SCROLLLEFT = &H2             'Left arrow of
                                                  'scroll bar
    DFCS_SCROLLRIGHT = &H3            'Right arrow of
                                                  'scroll bar
    DFCS_SCROLLCOMBOBOX = &H5         'Combo box scroll
                                                  'bar
    DFCS_SCROLLSIZEGRIP = &H8         'Size grip
    DFCS_SCROLLSIZEGRIPRIGHT = &H10   'Size grip in
                                                  'bottom-right
                                                  'corner of window
    
    DFCS_BUTTONCHECK = &H0      'Check box
    
    DFCS_BUTTONRADIO = &H4     'Radio button
    DFCS_BUTTON3STATE = &H8    'Three-state button
    DFCS_BUTTONPUSH = &H10     'Push button
    
    DFCS_INACTIVE = &H100      'Button is inactive
                                           '(grayed)
    DFCS_PUSHED = &H200        'Button is pushed
    DFCS_CHECKED = &H400       'Button is checked
    
    DFCS_ADJUSTRECT = &H2000   'Bounding rectangle is
                                           'adjusted to exclude the
                                           'surrounding edge of the
                                           'push button
    
    DFCS_FLAT = &H4000         'Button has a flat border
    DFCS_MONO = &H8000         'Button has a monochrome
                                           'border
End Enum

Public Enum LineStyles
    Default = 0
    [B] = 1
    [bf] = 2
    Box = 1
    boxfilled = 2
End Enum

Private myBackBuffer As Long
Private myBufferBMP1 As Long

Private hpic()  As Long
Private picold() As Long
Private picdim() As ImageDimensions
Private piccnt As Long

Private pOldBMP1 As Long

Private pOldBrush As Long
Private pOldPen As Long
Private pOldFont As Long

Private pHWnd As Long
Private pHDC As Long
Private hpen As Long
Private hbrush As Long
Private Brush As LOGBRUSH
Private Pen As LOGPEN

Private pForecolor As OLE_COLOR
Private pBackcolor As OLE_COLOR

Private pFont As StdFont
Private hFont As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Property Get Font() As StdFont
    Set Font = pFont
End Property
Public Property Set Font(ByRef RHS As StdFont)
    Set pFont = RHS
    UpdateFont
End Property

Private Function UpdateFont()
    If hFont <> 0 Then
        SelectObject myBackBuffer, pOldFont
        DeleteObject hFont
    End If
    Dim plf As LOGFONT, I As Long
    Dim FontName As String
    
    FontName = pFont.name + String(32 - Len(pFont.name), 0)
    For I = 1 To 32
        plf.lfFaceName(I) = Asc(Mid$(FontName, I, 1))
    Next
    'Height
    plf.lfHeight = (pFont.Size + 4)
    'Width
    plf.lfWidth = 0 ' Screen.TwipsPerPixelX / PixelPerPoint

    'Bold ,Underline ,Strikeout ,Italic
    plf.lfWeight = pFont.Weight
    plf.lfUnderline = -CInt(pFont.Underline)
    plf.lfStrikeOut = -CInt(pFont.Strikethrough)
    plf.lfItalic = -CInt(pFont.Italic)
    
    'Anti Aliasing
    plf.lfQuality = ANTIALIASED_QUALITY
    hFont = CreateFontIndirect(plf)
    If myBackBuffer <> 0 Then
        SetTextColor myBackBuffer, pForecolor
    End If
    pOldFont = SelectObject(myBackBuffer, hFont)
    
    SetBkMode myBackBuffer, -1
    
End Function

Public Property Get hdc() As Long
    hdc = pHDC
End Property
Public Property Let hdc(ByRef RHS As Long)
        
    If (pHWnd <> 0) Or (pHDC <> 0) Then
        
        Class_Terminate
        
    End If
    pHDC = RHS
    If (RHS <> 0) Then

        myBackBuffer = CreateCompatibleDC(GetDC(0))
        myBufferBMP1 = CreateCompatibleBitmap(GetDC(0), (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY))

        hpen = CreatePenIndirect(Pen)
        hbrush = CreateBrushIndirect(Brush)
        
        pOldBMP1 = SelectObject(myBackBuffer, myBufferBMP1)

        pOldPen = SelectObject(myBackBuffer, hpen)
        pOldBrush = SelectObject(myBackBuffer, hbrush)
        
        UpdateFont
        
        BitBlt myBackBuffer, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY), myBufferBMP1, 0, 0, vbSrcCopy

    End If

End Property

Public Property Get hWnd() As Long
    hWnd = pHWnd
End Property
Public Property Let hWnd(ByRef RHS As Long)

    pHWnd = RHS
    hdc = GetDC(hWnd)

End Property

Public Function LoadBitmap(ByVal Bitmap As String) As Long
    ' Device Context creation function
    
    If piccnt > 0 Then
        Dim I As Long
        For I = 1 To piccnt
            If hpic(I) = 0 Then
                LoadBitmap = I
                Exit For
            End If
        Next
    End If

    If LoadBitmap = 0 Then
        piccnt = piccnt + 1
        ReDim Preserve hpic(1 To piccnt) As Long
        ReDim Preserve picold(1 To piccnt) As Long
        ReDim Preserve picdim(1 To piccnt) As ImageDimensions
        LoadBitmap = piccnt
    End If
    
    hpic(LoadBitmap) = CreateCompatibleDC(GetDC(0))
    Dim imgdim As ImageDimensions
    
    ImageDimensions Bitmap, imgdim
    picdim(LoadBitmap).Width = imgdim.Width
    picdim(LoadBitmap).Height = imgdim.Height
    
    ' load the graphic file into the DC
    picold(LoadBitmap) = SelectObject(hpic(LoadBitmap), LoadPicture(Bitmap))

End Function

Public Sub UnloadBitmap(ByVal Load As Long)
    If Load >= 1 And Load <= piccnt Then
        If hpic(Load) <> 0 Then
            SelectObject hpic(Load), picold(Load)
            DeleteObject hpic(Load)
            hpic(Load) = 0
        End If
    End If
End Sub

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = pForecolor
End Property
Public Property Let Forecolor(ByRef RHS As OLE_COLOR)
    If Not (pForecolor = RHS) Then
        If pHDC <> 0 Then

            SelectObject myBackBuffer, pOldPen
            SelectObject myBackBuffer, pOldBrush
            
            DeleteObject hpen
            DeleteObject hbrush
        End If
        pForecolor = ConvertColor(RHS)
        If pHDC <> 0 Then
            Pen.lopnColor = pForecolor
            hpen = CreatePenIndirect(Pen)
            
            Brush.lbColor = pForecolor
            hbrush = CreateBrushIndirect(Brush)
            
            pOldPen = SelectObject(myBackBuffer, hpen)
            pOldBrush = SelectObject(myBackBuffer, hbrush)
            
        End If
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = pBackcolor
End Property
Public Property Let BackColor(ByRef RHS As OLE_COLOR)
    If Not (pBackcolor = RHS) Then
        pBackcolor = RHS
    End If
End Property

Public Sub DrawFrame(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal FrameType As FrameTypes, ByVal FrameStyle As FrameStyles)

    DrawFrameControl myBackBuffer, modGraphics.RECT(X1, Y1, X2, Y2), FrameType, FrameStyle
    
'    Select Case FrameStyle
'        Case BoxFills.Edge
'            MoveToEx myBackBuffer, X - 1, Y - 1, pt
'            LineTo myBackBuffer, X - 1, Y + Width - 1
'            LineTo myBackBuffer, X + Height - 1, Y + Width - 1
'            LineTo myBackBuffer, X + Height - 1, Y - 1
'            LineTo myBackBuffer, X - 1, Y - 1
'        Case BoxFills.Filled
'            Rectangle myBackBuffer, X - 1, Y - 1, X + Height - 1, Y + Width - 1
'
'        Case Gripper
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
'
'        Case Button, Button + Pushed
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_BUTTON, _
'                DFCS_BUTTONPUSH Or IIf(FrameStyle = Button + Pushed, DFCS_PUSHED, 0)
'
'        Case ArrowUp, ArrowUp + Disabled, ArrowUp + Disabled + Pushed
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_SCROLL, DFCS_SCROLLUP Or _
'                IIf(FrameStyle = ArrowUp + Disabled, DFCS_INACTIVE, 0) Or IIf(FrameStyle = ArrowUp + Disabled + Pushed, DFCS_PUSHED, 0)
'
'        Case ArrowDown, ArrowDown + Disabled, ArrowDown + Disabled + Pushed
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_SCROLL, DFCS_SCROLLDOWN Or _
'                IIf(FrameStyle = ArrowDown + Disabled, DFCS_INACTIVE, 0) Or IIf(FrameStyle = ArrowDown + Disabled + Pushed, DFCS_PUSHED, 0)
'
'        Case ArrowLeft, ArrowLeft + Disabled, ArrowLeft + Disabled + Pushed
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_SCROLL, DFCS_SCROLLLEFT Or _
'                IIf(FrameStyle = ArrowLeft + Disabled, DFCS_INACTIVE, 0) Or IIf(FrameStyle = ArrowLeft + Disabled + Pushed, DFCS_PUSHED, 0)
'
'        Case ArrowRight, ArrowRight + Disabled, ArrowRight + Disabled + Pushed
'            DrawFrameControl myBackBuffer, RECT(X, Y, X + Width, Y + Height), DFC_SCROLL, DFCS_SCROLLRIGHT Or _
'                IIf(FrameStyle = ArrowRight + Disabled, DFCS_INACTIVE, 0) Or IIf(FrameStyle = ArrowRight + Disabled + Pushed, DFCS_PUSHED, 0)
'
'    End Select

End Sub

Public Sub DrawCls(Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional BackgroundColor As Variant)
    If Width = -1 Then Width = Screen.Width / Screen.TwipsPerPixelX
    If Height = -1 Then Height = Screen.Height / Screen.TwipsPerPixelY
    
    Dim hold As Long

    If Not IsMissing(BackgroundColor) Then
        hold = Forecolor
        Forecolor = ConvertColor(BackgroundColor)
    Else
        hold = Forecolor
        Forecolor = pBackcolor
    End If

    Dim rct As RECT
    rct.Right = Width
    rct.Bottom = Height
    
    FillRect myBackBuffer, rct, hbrush
    
    BitBlt myBackBuffer, 0, 0, Width, Height, myBufferBMP1, 0, 0, vbSrcCopy

    Forecolor = hold

End Sub

Public Sub DrawBitmap(ByVal Load As Long, ByVal X As Long, ByVal Y As Long)
    If Load >= 1 And Load <= piccnt Then
        If hpic(Load) <> 0 Then
            BitBlt myBackBuffer, X, Y, picdim(Load).Width, picdim(Load).Height, hpic(Load), 0, 0, vbSrcCopy
        End If
    End If
End Sub


Public Sub Paint(Optional ByVal X1 As Long = 1, Optional ByVal Y1 As Long = 1, Optional ByVal X2 As Long = -1, Optional ByVal Y2 As Long = -1)
    If X2 = -1 Then X2 = Screen.Width / Screen.TwipsPerPixelX
    If Y1 = -1 Then Y1 = Screen.Height / Screen.TwipsPerPixelY
    
    BitBlt pHDC, X1, Y1, X2, Y2, myBackBuffer, 0, 0, vbSrcCopy
    
    BitBlt myBackBuffer, X1, Y1, X2, Y2, myBufferBMP1, 0, 0, vbSrcAnd
End Sub

Public Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ForegroundColor As Variant, Optional ByVal LineStyle As LineStyles)
    Dim hold As Long
    Dim pt As POINTAPI
    
    If Not IsMissing(ForegroundColor) Then
        hold = Forecolor
        Forecolor = ConvertColor(ForegroundColor)
    Else
        hold = Forecolor
        Forecolor = pBackcolor
    End If

    Select Case LineStyle
        Case LineStyles.[B], LineStyles.Box
            MoveToEx myBackBuffer, X1, Y1, pt
            lineto myBackBuffer, X1, Y2
            lineto myBackBuffer, X2, Y2
            lineto myBackBuffer, X2, Y1
            lineto myBackBuffer, X1, Y1
        Case LineStyles.[bf], LineStyles.boxfilled
            Rectangle myBackBuffer, X1, Y1, X2, Y2
            
        Case Else
            MoveToEx myBackBuffer, X1, Y1, pt
            lineto myBackBuffer, X2, Y2
    End Select
    
    Forecolor = hold
End Sub

Public Sub DrawText(ByVal X As Long, ByVal Y As Long, ByVal Text As String, Optional ForegroundColor As Variant)
    
    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, ConvertColor(ForegroundColor)
    Else
        SetTextColor myBackBuffer, ConvertColor(Forecolor)
    End If

    TextOut myBackBuffer, X, Y, Text, Len(Text)
    

    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, pForecolor
    End If

End Sub

Public Sub DrawByte(ByVal X As Long, ByVal Y As Long, ByVal Text As Byte, Optional ForegroundColor As Variant)
    
    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, ConvertColor(ForegroundColor)
    Else
        SetTextColor myBackBuffer, ConvertColor(Forecolor)
    End If

    TextOut myBackBuffer, X, Y, Chr(Text), 1
    

    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, pForecolor
    End If

End Sub
Public Sub DrawStrands(ByVal X As Long, ByVal Y As Long, ByRef Text As Strands, Optional ForegroundColor As Variant)
    
    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, ConvertColor(ForegroundColor)
    Else
        SetTextColor myBackBuffer, ConvertColor(Forecolor)
    End If

    TextOutPtr myBackBuffer, X, Y, ByVal Text.Address + 4, Text.Length
    

    If Not IsMissing(ForegroundColor) Then
        SetTextColor myBackBuffer, pForecolor
    End If

End Sub

Private Sub Class_Initialize()
    
    With Pen
        .lopnColor = vbBlack
        .lopnStyle = PS_SOLID
        .lopnWidth.X = 1
    End With

    With Brush
        .lbColor = vbWhite
        .lbStyle = PS_SOLID
        .lbHatch = HS_SOLID
    End With
    
    Set pFont = New StdFont

End Sub

Private Sub Class_Terminate()
    Dim I As Long
    If piccnt > 0 Then
        For I = 1 To piccnt
            If hpic(I) <> 0 Then DeleteObject hpic(I)
        Next
        Erase hpic
        Erase picdim
    End If

    SelectObject myBackBuffer, pOldFont
    DeleteObject hFont
    hFont = 0
    
    SelectObject myBackBuffer, pOldBrush
    DeleteObject hbrush
    hbrush = 0
    
    SelectObject myBackBuffer, pOldPen
    DeleteObject hpen
    hpen = 0
    
    If pHWnd <> 0 Then ReleaseDC pHWnd, pHDC
    
    SelectObject myBackBuffer, pOldBMP1
    DeleteObject myBufferBMP1
    myBufferBMP1 = 0

    DeleteDC myBackBuffer

End Sub
