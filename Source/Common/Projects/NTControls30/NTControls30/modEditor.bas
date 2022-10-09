Attribute VB_Name = "modEditor"
Option Explicit



Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RangeType
    StartPos As Long
    StopPos As Long
End Type

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TextRange
    Range As RangeType
    lpStr As Long
End Type

Public Type UndoType
    PriorTextData As String
    AfterTextData As String
    PriorSelRange As RangeType
    AfterSelRange As RangeType
End Type

Public Enum TextModes
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32
End Enum

Public Type SetTextEx
    Flags As Long
    codepage As Long
End Type

Public Type GetTextEx
    cb As Long
    Flags As Long
    codepage As Long
    lpDefaultChar As Long
    lpUsedDefChar As Long
End Type

Public Type GetTextLengthEx
    Flags As Long
    codepage As Long
End Type

Public Const ST_DEFAULT = 0
Public Const ST_KEEPUNDO = 1
Public Const ST_SELECTION = 2
Public Const ST_NEWCHARS = 3
Public Const ST_UNICODE = 4


Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1
Public Const GT_SELECTION = 2

Public Const GTL_DEFAULT = 0
Public Const GTL_USECRLF = 1
Public Const GTL_PRECISE = 2
Public Const GTL_CLOSE = 4
Public Const GTL_NUMCHARS = 8
Public Const GTL_NUMBYTES = 16

Public Const SB_HORZ = 0
Public Const SB_VERT = 1

Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5

Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&

Public Const SB_ENDSCROLL = 8

Public Const WM_COMMAND = &H111
Public Const WM_ERASEBKGND = &H14

Public Const WM_USER = &H400

Public Const WM_MOUSEWHEEL = &H20A ' window message for mouse wheel
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_PAINT = &HF
Public Const WM_GETTEXT = &HD
Public Const EM_GETSEL = &HB0&
Public Const EM_SETSEL = &HB1&
Public Const EM_CANUNDO = &HC6&
Public Const EM_UNDO = &HC7&
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)

Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21

Public Const GWL_WNDPROC = (-4)

Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_SCROLL As Long = &HB5
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_LINESCROLL = &HB6
Public Const EM_GETLINE = &HC4&
Public Const EM_LINEINDEX = &HBB&
Public Const EM_LINELENGTH = &HC1&
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETSCROLLPOS = (WM_USER + 221)
Public Const EM_SETSCROLLPOS = (WM_USER + 222)
Public Const EM_GETFIRSTVISIBLELINE = &HCE&
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_SCROLLCARET = &HB7&
Public Const EM_CHARFROMPOS = &HD7&
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_SETTEXTEX = (WM_USER + 97)
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_LINEFROMCHAR = &HC9&

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_MBUTTON = &H4

Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long


Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'Gets state of one key

Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long

Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageMemory Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any) As Variant

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Declare Function IntersectRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
'Public Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal lLeft As Long, ByVal lTop As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal lLeft As Long, ByVal lTop As Long) As Long


Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As Rect) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPIF_SENDCHANGE = &H2

Public Const DFC_CAPTION = 1            'Title bar
Public Const DFC_MENU = 2               'Menu
Public Const DFC_SCROLL = 3             'Scroll bar
Public Const DFC_BUTTON = 4             'Standard button

Public Const DFCS_CAPTIONCLOSE = &H0    'Close button
Public Const DFCS_CAPTIONMIN = &H1      'Minimize button
Public Const DFCS_CAPTIONMAX = &H2      'Maximize button
Public Const DFCS_CAPTIONRESTORE = &H3  'Restore button
Public Const DFCS_CAPTIONHELP = &H4     'Windows 95 only:
                                        'Help button

Public Const DFCS_MENUARROW = &H0       'Submenu arrow
Public Const DFCS_MENUCHECK = &H1       'Check mark
Public Const DFCS_MENUBULLET = &H2      'Bullet
Public Const DFCS_MENUARROWRIGHT = &H4

Public Const DFCS_SCROLLUP = &H0               'Up arrow of scroll
                                               'bar
Public Const DFCS_SCROLLDOWN = &H1             'Down arrow of
                                               'scroll bar
Public Const DFCS_SCROLLLEFT = &H2             'Left arrow of
                                               'scroll bar
Public Const DFCS_SCROLLRIGHT = &H3            'Right arrow of
                                               'scroll bar
Public Const DFCS_SCROLLCOMBOBOX = &H5         'Combo box scroll
                                               'bar
Public Const DFCS_SCROLLSIZEGRIP = &H8         'Size grip
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10   'Size grip in
                                               'bottom-right
                                               'corner of window

Public Const DFCS_BUTTONCHECK = &H0      'Check box

Public Const DFCS_BUTTONRADIO = &H4     'Radio button
Public Const DFCS_BUTTON3STATE = &H8    'Three-state button
Public Const DFCS_BUTTONPUSH = &H10     'Push button

Public Const DFCS_INACTIVE = &H100      'Button is inactive
                                        '(grayed)
Public Const DFCS_PUSHED = &H200        'Button is pushed
Public Const DFCS_CHECKED = &H400       'Button is checked

Public Const DFCS_ADJUSTRECT = &H2000   'Bounding rectangle is
                                        'adjusted to exclude the
                                        'surrounding edge of the
                                        'push button

Public Const DFCS_FLAT = &H4000         'Button has a flat border
Public Const DFCS_MONO = &H8000         'Button has a monochrome
                                        'border

Public Declare Function DrawFrameControl Lib "user32" (ByVal _
   hdc&, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) _
   As Boolean


Public Const SM_CYSIZE = 31&   'Titlebar height
Public Const SM_CXSIZEFRAME = 32&
Public Const SM_CYSIZEFRAME = 33&

Public Const SM_CXBORDER = 5&  'Borders width
Public Const SM_CYCAPTION = 4& 'caption height

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)


Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As Rect, lpClipRect As Rect) As Long


Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const MOUSE_MOVED = &H1

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public Const LOGPIXELSX = 88
Private Const POINTS_PER_INCH As Long = 100

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Public SubClassed As New VBA.Collection

Public SubClassed2 As New VBA.Collection


Public Function PixelPerPoint() As Double
    Dim hdc As Long
    Dim lPixelPerInch As Long
    hdc = GetDC(ByVal 0&)
    lPixelPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
    PixelPerPoint = POINTS_PER_INCH / lPixelPerInch
    ReleaseDC ByVal 0&, hdc
    PixelPerPoint = 1 + PixelPerPoint
End Function

Public Function Rect(l, t, r, b) As Rect
    With Rect
        .Left = CLng(l)
        .Top = CLng(t)
        .Right = CLng(r)
        .Bottom = CLng(b)
    End With
End Function

Public Function pt(X, Y) As POINTAPI
    pt.X = CLng(X)
    pt.Y = CLng(Y)
End Function

Public Function Range(ByVal StartPos As Long, ByVal StopPos As Long) As RangeType
    Range.StartPos = StartPos
    Range.StopPos = StopPos
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function

Public Function RangeLength(ByRef Range As RangeType) As Long
    RangeLength = IIf((Range.StopPos >= Range.StartPos), (Range.StopPos - Range.StartPos), IIf((Range.StartPos = -1) And (Range.StopPos = 0), -1, 0))
End Function

Public Function GetTextRange(ByVal hWnd As Long, ByVal StartPos As Long, ByVal StopPos As Long) As String
    If (StopPos - StartPos) > 0 Then
        Dim txt As TextRange
        Dim str As String
        str = Space((StopPos - StartPos))
        txt.Range.StartPos = StartPos
        txt.Range.StopPos = StopPos
        RtlMoveMemory txt.lpStr, ByVal VarPtr(str), 4&
        If SendMessageLngPtr(hWnd, EM_GETTEXTRANGE, 0, ByVal VarPtr(txt)) Then
            RtlMoveMemory ByVal VarPtr(str), txt.lpStr, 4&
            GetTextRange = Left(StrConv(str, vbUnicode), StopPos - StartPos)
        End If
    End If
End Function

Public Sub SetScrollbarsByVisibility(ByRef frm As Neotext, Optional ByVal SelStartWithLength As Boolean = True)
    frm.VScroll.AutoRedraw = False
    frm.HScroll.AutoRedraw = False
    
    Dim lLength As Long
    Dim lLine As Long
    Dim textSel As POINTAPI
    Dim lIndex  As Long
    Dim useCursor As Long
    Dim txt As String
    lLine = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
    If lLine > frm.VScroll.Max Then
        If frm.VScroll.Value <> frm.VScroll.Max Then
            frm.VScroll.Value = frm.VScroll.Max
        End If
    ElseIf frm.VScroll.Value <> lLine Then
        frm.VScroll.Value = lLine
    End If
    SendMessageStruct frm.RichTextBox.hWnd, EM_EXGETSEL, 0, textSel
    useCursor = IIf(SelStartWithLength, textSel.Y, textSel.X)
    lLine = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_EXLINEFROMCHAR, 0, useCursor)
    lIndex = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_LINEINDEX, lLine, 0&)
    lLength = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_LINELENGTH, useCursor, 0&)
    txt = GetTextRange(frm.RichTextBox.hWnd, lIndex, lIndex + lLength)
    lLength = frm.TextWidth(Left(txt, useCursor - lIndex))
    If frm.TextWidth(Left(txt, useCursor - lIndex)) - frm.HScroll.Value < 0 Or _
        frm.TextWidth(Left(txt, useCursor - lIndex)) - frm.HScroll.Value > frm.RichTextBox.Width Then
        lLength = lLength - (frm.RichTextBox.Width \ 4)
        If lLength >= frm.HScroll.Min And lLength <= frm.HScroll.Max Then
            If frm.HScroll.Value <> lLength Then
                frm.HScroll.Value = lLength
            End If
        ElseIf lLength > frm.HScroll.Max Then
            If frm.HScroll.Value <> frm.HScroll.Max Then
                frm.HScroll.Value = frm.HScroll.Max
            End If
        ElseIf lLength < frm.HScroll.Min Then
            If frm.HScroll.Value <> frm.HScroll.Min Then
                frm.HScroll.Value = frm.HScroll.Min
            End If
        End If
    End If
    
    frm.VScroll.AutoRedraw = True
    frm.HScroll.AutoRedraw = True
End Sub
Public Sub AdjustScollbarValue(ByRef ScrollBar As ScrollBar, ByVal Offset As Long)
    ScrollBar.AutoRedraw = False
    
    If Offset > 0 Then
        If ScrollBar.Value + Offset <= ScrollBar.Max Then
            If ScrollBar.Value <> ScrollBar.Value + Offset Then
                ScrollBar.Value = ScrollBar.Value + Offset
            End If
        ElseIf ScrollBar.Value <> ScrollBar.Max Then
            ScrollBar.Value = ScrollBar.Max
        End If
    ElseIf Offset < 0 Then
        If ScrollBar.Value + Offset >= ScrollBar.Min Then
            If ScrollBar.Value <> ScrollBar.Value + Offset Then
                ScrollBar.Value = ScrollBar.Value + Offset
            End If
        ElseIf ScrollBar.Value <> ScrollBar.Min Then
            ScrollBar.Value = ScrollBar.Min
        End If
    End If
    
    ScrollBar.AutoRedraw = True
End Sub

Public Function WinProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim lDispatch As Boolean
    lDispatch = True
    
    If SubClassed.Count > 0 Then
        Dim frm As Neotext
        Dim cnt As Long
        For cnt = 1 To SubClassed.Count
            Set frm = SubClassed(cnt)
            If frm.RichTextBox.hWnd = hWnd Then
            
                Select Case wMsg
                    Case WM_MOUSEWHEEL
                        If GetAsyncKeyState(VK_LSHIFT) Or GetAsyncKeyState(VK_RSHIFT) Then
                            AdjustScollbarValue frm.HScroll, ((frm.HScroll.LargeChange - frm.HScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                        Else
                            AdjustScollbarValue frm.VScroll, ((frm.VScroll.LargeChange - frm.VScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                        End If
                        If frm.LineNumbers Then frm.SetScrollBarsMax True
                    Case WM_VSCROLL
                        SetScrollbarsByVisibility frm
                        frm.SetScrollBarsMax
                    Case WM_HSCROLL
                        SetScrollbarsByVisibility frm
                        frm.SetScrollBarsMax
                    Case EM_SETSCROLLPOS
                       ' Debug.Print "Scroll"
                       

'                    Case WM_COPY
'
'                    Case WM_PASTE
'
'                    Case WM_CUT
'
'                    Case WM_CLEAR
'
'                    Case EM_GETSEL
'                    Case EM_SETSEL
'                        If frm.Cancel Then lDispatch = False
'                        WinProc = 1
'                    Case EM_CANUNDO
'                        lDispatch = False
'                    Case EM_CANREDO
'                        lDispatch = False
'                    Case EM_UNDO
'                        lDispatch = False
'                    Case EM_REDO
'                        lDispatch = False
'                    Case WM_PAINT
'                        lDispatch = Not frm.Cancel
                End Select
    
               ' If (lDispatch And (WinProc = 0)) Then
                    WinProc = CallWindowProc(frm.OldProc, hWnd, wMsg, wParam, lParam)

                'End If

            End If
            Set frm = Nothing
        Next
    End If
End Function



Public Function TextBoxProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim lDispatch As Boolean
    lDispatch = False
    
    If SubClassed2.Count > 0 Then
        Dim frm As TextBox
        Dim cnt As Long
        For cnt = 1 To SubClassed2.Count
            Set frm = SubClassed2(cnt)
            If frm.hWnd = hWnd Then
            
                Select Case wMsg
                    Case WM_MOUSEWHEEL
                        If GetAsyncKeyState(VK_LSHIFT) Or GetAsyncKeyState(VK_RSHIFT) Then
                            AdjustScollbarValue frm.HScroll, ((frm.HScroll.LargeChange - frm.HScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                        Else
                        
                            AdjustScollbarValue frm.VScroll, ((frm.VScroll.LargeChange - frm.VScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                        End If
                    Case WM_VSCROLL
                    Case WM_HSCROLL
                    Case WM_COPY
                    
                    Case WM_PASTE
                    Case WM_CUT
                    Case WM_CLEAR
                    Case WM_PAINT
                        lDispatch = True

                    Case EM_GETSEL
                    Case EM_SETSEL
                    Case EM_EXGETSEL
                    Case EM_EXSETSEL

                    Case EM_CANUNDO
                    Case EM_CANREDO
                    Case EM_UNDO
                    Case EM_REDO
                                      
                    Case EM_GETSCROLLPOS
                    Case EM_SETSCROLLPOS
                    
                    Case EM_GETLINECOUNT

                    Case EM_CHARFROMPOS
                    Case EM_EXLINEFROMCHAR
                    Case EM_GETFIRSTVISIBLELINE

                    Case EM_LINEINDEX
                    Case EM_LINELENGTH
                    Case EM_GETTEXTLENGTHEX
                    Case EM_SETTEXTEX
                    Case EM_GETTEXTRANGE

                    Case EM_SETUNDOLIMIT
                    Case EM_SETTEXTMODE
                    Case Else
                        lDispatch = True
                End Select
    
                If (lDispatch And (TextBoxProc = 0)) Then
                    TextBoxProc = CallWindowProc(frm.OldProc, hWnd, wMsg, wParam, lParam)
                End If

            End If
            Set frm = Nothing
        Next
    End If
End Function


