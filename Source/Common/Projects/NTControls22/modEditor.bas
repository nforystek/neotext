Attribute VB_Name = "modEditor"
#Const [True] = -1
#Const [False] = 0
#Const modEditor = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Type RectType
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RangeType
    StartPos As Long
    StopPos As Long
End Type

Public Type TextRange
    Range As RangeType
    lpStr As Long
End Type


Public Type PageType
    Flags As Long
    CodePage As Long
End Type

Public Type FindType
    Range As RangeType
    Buffer As Long
    Result As RangeType
End Type

Public Type PointType
    x As Long
    y As Long
End Type

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Public Const GWL_WNDPROC = -4

Public Const GW_OWNER = 4

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3

Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7

Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&

Public Const SB_BOTTOM = 7
Public Const SB_ENDSCROLL = 8
Public Const SB_LEFT = 6
Public Const SB_LINEDOWN = 1
Public Const SB_LINELEFT = 0
Public Const SB_LINERIGHT = 1
Public Const SB_LINEUP = 0
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGELEFT = 2
Public Const SB_PAGERIGHT = 3
Public Const SB_PAGEUP = 2
Public Const SB_RIGHT = 7
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6

Public Const WM_COMMAND = &H111
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_USER = &H400
Public Const WM_DRAWITEM = &H2B
Public Const WM_SETCURSOR = &H20
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303

Public Const WM_PAINT = &HF
Public Const WM_SETREDRAW = &HB

Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_ERASEBKGND = &H14
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCPAINT = &H85
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_PAINTICON = &H26
Public Const WM_HOTKEY = &H312
Public Const WM_ICONERASEBKGND = &H27

Public Const WM_CLOSE = &H10
Public Const WM_SYSCOMMAND = &H112
Public Const WM_ACTIVATE = &H6
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_KILLFOCUS = &H8
Public Const WM_NCACTIVATE = &H86
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const VK_HOME = &H24
Public Const VK_END = &H23
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_RIGHT = &H27
Public Const VK_LEFT = &H25

Public Const VK_CAPITAL = &H14
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91

Public Const VK_MENU = &H12
Public Const VK_RMENU = &HA5
Public Const VK_LMENU = &HA4

Public Const VK_CONTROL = &H11
Public Const VK_RCONTROL = &HA3
Public Const VK_LCONTROL = &HA2

Public Const VK_SHIFT = &H10
Public Const VK_RSHIFT = &HA1
Public Const VK_LSHIFT = &HA0

Public Const VK_ESCAPE = &H1B
Public Const VK_TAB = &H9
Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD

Public Const VK_PAUSE = &H13
Public Const VK_PRINT = &H2A
Public Const VK_SELECT = &H29

Public Const VK_BACK = &H8
Public Const VK_NEXT = &H22

Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const ST_DEFAULT = 0
Public Const ST_KEEPUNDO = 1
Public Const ST_SELECTION = 2

Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1
Public Const GT_SELECTION = 2

Public Const GTL_DEFAULT = 0
Public Const GTL_USECRLF = 1
Public Const GTL_PRECISE = 2
Public Const GTL_CLOSE = 4
Public Const GTL_NUMCHARS = 8
Public Const GTL_NUMBYTES = 16

Public Const CP_ACP = 0
Public Const CP_UNI = 1200

Public Const FR_DOWN = &H1
Public Const FR_WHOLEWORD = &H2
Public Const FR_MATCHCASE = &H4&
Public Const FR_NUMERIC = &H8&

Public Const EM_GETSEL = &HB0&
Public Const EM_SETSEL = &HB1&
Public Const EM_GETRECT = &HB2&
Public Const EM_SETRECT = &HB3&
Public Const EM_SETRECTNP = &HB4&
Public Const EM_SCROLL = &HB5&
Public Const EM_LINESCROLL = &HB6&
Public Const EM_SCROLLCARET = &HB7&
Public Const EM_GETMODIFY = &HB8&
Public Const EM_SETMODIFY = &HB9&
Public Const EM_GETLINECOUNT = &HBA&
Public Const EM_LINEINDEX = &HBB&
Public Const EM_SETHANDLE = &HBC&
Public Const EM_GETHANDLE = &HBD&
Public Const EM_GETTHUMB = &HBE&
Public Const EM_LINELENGTH = &HC1&
Public Const EM_REPLACESEL = &HC2&
Public Const EM_GETLINE = &HC4&
Public Const EM_LIMITTEXT = &HC5&
Public Const EM_CANUNDO = &HC6&
Public Const EM_UNDO = &HC7&
Public Const EM_FMTLINES = &HC8&
Public Const EM_LINEFROMCHAR = &HC9&
Public Const EM_SETTABSTOPS = &HCB&
Public Const EM_SETPASSWORDCHAR = &HCC&
Public Const EM_EMPTYUNDOBUFFER = &HCD&
Public Const EM_GETFIRSTVISIBLELINE = &HCE&
Public Const EM_SETREADONLY = &HCF&
Public Const EM_SETWORDBREAKPROC = &HD0&
Public Const EM_GETWORDBREAKPROC = &HD1&
Public Const EM_GETPASSWORDCHAR = &HD2&
Public Const EM_SETMARGINS = &HD3&
Public Const EM_GETMARGINS = &HD4&
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT
Public Const EM_GETLIMITTEXT = &HD5&
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_CHARFROMPOS = &HD7&

Public Const EM_SETTEXTEX = (WM_USER + 97)
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_DISPLAYBAND = (WM_USER + 51)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXLIMITTEXT = (WM_USER + 53)
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_FORMATRANGE = (WM_USER + 57)
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_GETEVENTMASK = (WM_USER + 59)
Public Const EM_GETOLEINTERFACE = (WM_USER + 60)
Public Const EM_GETPARAFORMAT = (WM_USER + 61)
Public Const EM_GETSELTEXT = (WM_USER + 62)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_PASTESPECIAL = (WM_USER + 64)
Public Const EM_REQUESTRESIZE = (WM_USER + 65)
Public Const EM_SELECTIONTYPE = (WM_USER + 66)
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const EM_SETOLECALLBACK = (WM_USER + 70)
Public Const EM_SETPARAFORMAT = (WM_USER + 71)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_STREAMOUT = (WM_USER + 74)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_FINDWORDBREAK = (WM_USER + 76)
Public Const EM_SETOPTIONS = (WM_USER + 77)
Public Const EM_GETOPTIONS = (WM_USER + 78)
Public Const EM_FINDTEXTEX = (WM_USER + 79)
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80)
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81)

Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)
Public Const EM_FINDTEXTW = (WM_USER + 123)
Public Const EM_FINDTEXTEXW = (WM_USER + 124)

Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)
Public Const EM_SETPALETTE = (WM_USER + 93)
Public Const EM_GETTEXTEX = (WM_USER + 94)
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)

Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const EN_CHANGE = &H300
Public Const EN_UPDATE = &H400
Public Const EN_ERRSPACE = &H500
Public Const EN_MAXTEXT = &H501
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602

Public Const CFM_BOLD = &H1
Public Const CFM_ITALIC = &H2
Public Const CFM_UNDERLINE = &H4
Public Const CFM_STRIKEOUT = &H8
Public Const CFM_PROTECTED = &H10
Public Const CFM_LINK = &H20&
Public Const CFM_SIZE = &H80000000
Public Const CFM_COLOR = &H40000000
Public Const CFM_FACE = &H20000000
Public Const CFM_OFFSET = &H10000000
Public Const CFM_CHARSET = &H8000000

Public Const CFE_BOLD = &H1&
Public Const CFE_ITALIC = &H2&
Public Const CFE_UNDERLINE = &H4&
Public Const CFE_STRIKEOUT = &H8&
Public Const CFE_PROTECTED = &H10&
Public Const CFE_LINK = &H20&
Public Const CFE_AUTOCOLOR = &H40000000

Public Const SCF_SELECTION = &H1&
Public Const SCF_WORD = &H2&
Public Const SCF_DEFAULT = &H0&
Public Const SCF_ALL = &H4&
Public Const SCF_USEUIRULES = &H8&

Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
Public Const ENM_CORRECTTEXT = &H400000
Public Const ENM_SCROLLEVENTS = &H8
Public Const ENM_DRAGDROPDONE = &H10

Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Function SendMessageInteger Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As PointType) As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal x As Integer, ByVal y As Integer) As Boolean

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointType) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long
Public Declare Function GetTextCharacterExtra Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RectType, ByVal bErase As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RectType) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RectType) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal N As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RectType, lpClipRect As RectType) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const LOGPIXELSX = 88
Private Const POINTS_PER_INCH As Long = 72

Public xHooked As Boolean

Public Sub Main()

'%LICENSE%
    
    xHooked = (Not IsRunningMode) Or (Not IsDebugger)
End Sub

Public Function PixelPerPoint() As Double
    Dim hDC As Long
    Dim lPixelPerInch As Long
    hDC = GetDC(ByVal 0&)
    lPixelPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelPerPoint = POINTS_PER_INCH / lPixelPerInch
    ReleaseDC ByVal 0&, hDC
    PixelPerPoint = 1 + PixelPerPoint
End Function


Public Sub FileProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error GoTo notloaded
    Static stacking As Long
    stacking = stacking + 1
    If stacking > 1 Then GoTo notloaded
    Dim obj As CodeEdit
    Set obj = PtrObj(idEvent)
        
    If (Not (obj Is Nothing)) Then

        Static startByte As Long
        Static stopByte As Long
        
        If obj.FillFilePass(startByte, stopByte, idEvent) Then
            startByte = 0
            stopByte = 0
        End If

        Set obj = Nothing
    Else
        KillTimer hwnd, idEvent
    End If

notloaded:

    If Err Then Err.Clear
    On Error GoTo 0
    stacking = stacking - 1
End Sub

Public Static Function HookObj(ByRef obj)
On Error GoTo hookerror
    Static hc As Collection
    Static ha As Collection
    If xHooked Then
        If IsNumeric(obj) Then
            If obj < 0 Then
                HookObj = ha("k" & -obj)
            Else
                Set HookObj = hc("k" & obj)
            End If
        Else
            If ha Is Nothing Then Set ha = New Collection
            If hc Is Nothing Then Set hc = New Collection
            Dim cnt As Long
            If hc.count > 0 Then
                For cnt = 1 To hc.count
                    If hc(cnt).hwnd = obj.hwnd Then
                        If obj.hwnd > 0 Then
                        SetWindowLong obj.hwnd, _
                        GWL_WNDPROC, ha("k" & obj.hwnd)
                        hc.Remove "k" & obj.hwnd
                        ha.Remove "k" & obj.hwnd
                        End If
                        GoTo hookok
                    End If
                Next
            End If
            hc.Add obj, "k" & obj.hwnd
            ha.Add GetWindowLong(obj.hwnd, GWL_WNDPROC), "k" & obj.hwnd
            SetWindowLong obj.hwnd, GWL_WNDPROC, AddressOf ControlWndProc
        End If
hookok:
        If ha.count = 0 Then Set ha = Nothing
        If hc.count = 0 Then Set hc = Nothing
    End If
hookerror:
    If Err Then Err.Clear
End Function
Private Function ControlWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Debug.Print TypeName(HookObj(hWnd)) & ", " & hWnd & ", " & uMsg & ", " & wParam & ", " & lParam

    Dim lDispatch As Boolean
    lDispatch = True

    If (HookObj(-hwnd) <> 0) Then
        Dim obj As RichTextBox
        Dim par As CodeEdit
        Set obj = HookObj(hwnd)
        If Not obj Is Nothing Then
            Set par = obj.Parent
        
            Select Case uMsg
                Case EM_CANUNDO
                    lDispatch = False
                Case EM_CANREDO
                    lDispatch = False
                Case EM_UNDO
                    lDispatch = False
                Case EM_REDO
                    lDispatch = False
                Case WM_CUT
                    lDispatch = True
                Case WM_COPY  '= &H301
                    lDispatch = True
                Case WM_PASTE  '= &H302
                    lDispatch = True
                Case WM_CLEAR  '= &H303
                    lDispatch = True
                Case WM_PAINT
                    lDispatch = Not par.Cancel
                    'If Not lDispatch Then ControlWndProc = 1
                Case Else
                    If Not par.Cancel Then
                        Select Case uMsg
                            Case WM_MOUSEWHEEL
                                SendMessageLngPtr hwnd, EM_LINESCROLL, 0, IIf((wParam > 0), -4, 4)
                                par.RefreshView
                                lDispatch = False
                            Case WM_VSCROLL, WM_HSCROLL

                                If Not ((wParam And SB_THUMBPOSITION) = SB_THUMBPOSITION) Then
                                    AdjustScrollBar hwnd, uMsg, SendMessageLngPtr(hwnd, EM_GETLINECOUNT, 0, 0)
                                End If
                                
                                par.RefreshView
                        End Select
                    End If
            End Select
    
            If (lDispatch And (ControlWndProc = 0)) Then
                ControlWndProc = CallWindowProc(HookObj(-hwnd), hwnd, uMsg, wParam, lParam)
            End If
    
            Set par = Nothing
        End If
        
        Set obj = Nothing

    End If

End Function

Public Sub AdjustScrollBar(ByVal hwnd As Long, ByVal uMsg As Long, ByVal sCount As Long)
On Error GoTo walkoff
            
    If sCount > 0 Then
        Dim sInfo As SCROLLINFO
        Dim sBar As Long
        
        sInfo.cbSize = Len(sInfo)
        sInfo.fMask = SIF_POS Or SIF_RANGE
        sBar = IIf(uMsg = WM_VSCROLL, SB_VERT, SB_HORZ)
        
        If GetScrollInfo(hwnd, sBar, sInfo) Then

            sCount = CDbl(CDbl(sInfo.nMax) \ CDbl(sCount))
            sInfo.nPos = Round(CDbl(sInfo.nPos) / CDbl(sCount), 0) * sCount
            
            SendMessageLngPtr hwnd, uMsg, SB_THUMBPOSITION Or &H10000 * sInfo.nPos, 0
            
        End If
    Else
        SendMessageLngPtr hwnd, uMsg, SB_THUMBPOSITION Or &H10000 * 0, 0
    End If
    
    On Error GoTo 0
Exit Sub
walkoff:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub


Public Function RangeLength(ByRef Range As RangeType) As Long
    RangeLength = IIf((Range.StopPos >= Range.StartPos), (Range.StopPos - Range.StartPos), IIf((Range.StartPos = -1) And (Range.StopPos = 0), -1, 0))
End Function

Public Function GetLineNumber(ByVal hwnd As Long, ByVal nIndex As Long) As Long
    GetLineNumber = SendMessageLngPtr(hwnd, EM_EXLINEFROMCHAR, 0, nIndex)
End Function

Public Property Get PtrObj(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    CopyMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    CopyMemory NewObj, lZero, 4&
End Property

Public Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
    If (DWord And &H8000&) = 0 Then
        LoWord = DWord And &HFFFF&
    Else
        LoWord = DWord Or &HFFFF0000
    End If
End Function







