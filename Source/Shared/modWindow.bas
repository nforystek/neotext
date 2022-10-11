Attribute VB_Name = "modWindow"

#Const modWindow = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Public Const WM_NULL = &H0
Public Const WM_USER = &H400
Public Const WM_APP = &H8000

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_ENABLE = &HA
Public Const WM_CLOSE = &H10
Public Const WM_QUIT = &H12

Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_SIZING = &H214
Public Const WM_MOVING = &H216
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232

Public Const WM_MOUSEHOVER = &H2A1
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSEHWHEEL = &H20E

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_KEYLAST = &H108

Public Const WM_SHOWWINDOW = &H18
Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_NOTIFY = &H4E
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8

Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const WM_PAINT = &HF
Public Const WM_SETREDRAW = &HB
Public Const WM_ERASEBKGND = &H14
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306

Public Const WM_QUEUESYNC = &H23
Public Const WM_QUERYOPEN = &H13
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_ENDSESSION = &H16

Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSTEMERROR = &H17
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_SYSCOMMAND = &H112

Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F

Public Const WM_SETCURSOR = &H20
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_HOTKEY = &H312

Public Const WM_POWER = &H48
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NOTIFYFORMAT = &H55
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41
Public Const WM_TCARD = &H52
Public Const WM_HELP = &H53

Public Const WM_CAPTURECHANGED = &H215
Public Const WM_DEVICECHANGE = &H219
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F

Public Const WM_INPUTLANGCHANGE = &H51
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_INPUTLANGCHANGEREQUEST = &H50
Public Const WM_USERCHANGED = &H54
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_STYLECHANGING = &H7C
Public Const WM_STYLECHANGED = &H7D

Public Const WM_GETDLGCODE = &H87
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_INITDIALOG = &H110

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMOUSELEAVE = &H2A2

Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_TIMER = &H113
Public Const WM_COMMAND = &H111
Public Const WM_PRINT = &H317
Public Const WM_PRINTCLIENT = &H318
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_DROPFILES = &H233
Public Const WM_POWERBROADCAST = &H218

Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F

Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_NEXTMENU = &H213

Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CTLCOLOR = &H19

Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_MDIREFRESHMENU = &H234

Public Const WM_HANDHELDFIRST = &H358
Public Const WM_HANDHELDLAST = &H35F
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_COALESCE_FIRST = &H390
Public Const WM_COALESCE_LAST = &H39F

Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291

Public Const WM_DDE_FIRST = &H3E0
Public Const WM_DDE_INITIATE = &H3E0
Public Const WM_DDE_TERMINATE = &H3E1
Public Const WM_DDE_ADVISE = &H3E2
Public Const WM_DDE_UNADVISE = &H3E3
Public Const WM_DDE_ACK = &H3E4
Public Const WM_DDE_DATA = &H3E5
Public Const WM_DDE_REQUEST = &H3E6
Public Const WM_DDE_POKE = &H3E7
Public Const WM_DDE_EXECUTE = &H3E8
Public Const WM_DDE_LAST = &H3E8

Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_HSCROLLCLIPBOARD = &H30E


Public Const WS_CUSTOM = 0
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_DISABLED = &H8000000


Public Const GWL_WNDPROC = (-4)


Public Const SC_SIZE = &HF000&
Public Const SC_MOVE = &HF010&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_NEXTWINDOW = &HF040&
Public Const SC_PREVWINDOW = &HF050&
Public Const SC_CLOSE = &HF060&
Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&
Public Const SC_MOUSEMENU = &HF090&
Public Const SC_KEYMENU = &HF100&
Public Const SC_ARRANGE = &HF110&
Public Const SC_RESTORE = &HF120&
Public Const SC_TASKLIST = &HF130&
Public Const SC_SCREENSAVE = &HF140&
Public Const SC_HOTKEY = &HF150&

Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2

Public Const IDC_ARROW = 32512&
Public Const ES_MULTILINE = &H4&
Public Const CW_USEDEFAULT = &H80000000
Public Const COLOR_WINDOW = 5
Public Const IDI_APPLICATION = 32512&

Public Const SW_SHOWNORMAL = 1

Public Const MB_OK = &H0&
Public Const MB_ICONEXCLAMATION = &H30&

Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)

Public Enum TextModes
  TM_PLAINTEXT = 1
  TM_RICHTEXT = 2
  TM_SINGLELEVELUNDO = 4
  TM_MULTILEVELUNDO = 8
  TM_SINGLECODEPAGE = 16
  TM_MULTICODEPAGE = 32
End Enum

Public Const CP_UNICODE As Long = 1200&
    
Private Const HWND_MESSAGE As Long = -3

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_RESTORE = 9
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10

#If Not modWndProc = -1 Then
Public Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Public Const PBT_APMSUSPEND As Long = &H4
#End If
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
    
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function GetMessageTime Lib "user32" () As Long
Public Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As Msg, ByVal nCode As Long) As Long

Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function CreateWindow Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const DecaultClassName = "Message"

#If Not modCommon Then
    Private Const GMEM_FIXED = &H0
    Private Const GMEM_MOVEABLE = &H2
    Private Const GPTR = &H40
    Private Const GHND = &H42
    
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Sub RtlMoveMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Long, ByRef Source As Long, ByVal Length As Long)
    Private Declare Sub RtlMoveMemory3 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Public Function WindowClassName(ByVal hwnd As Long) As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetClassName(hwnd, sBuffer, lSize)
    If lSize > 0 Then
        WindowClassName = Replace(Left$(sBuffer, lSize), Chr(0), "")
    End If
End Function

Public Function WindowText(ByVal hwnd As Long) As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetWindowText(hwnd, sBuffer, lSize)
    If lSize > 0 Then
        WindowText = Trim(Replace(Left$(sBuffer, lSize), Chr(0), ""))
    End If
End Function

Private Function GetAddress(ByVal lngAddr As Long) As Long
   
    GetAddress = lngAddr
    
End Function

Public Function WindowHook(ByRef lHwnd As Long, ByVal lWndProc As Long) As Long

    If (Not (lHwnd = 0)) Then
        WindowHook = SetWindowLong(lHwnd, GWL_WNDPROC, lWndProc)
    End If

End Function

Public Sub WindowUnhook(ByRef lHwnd As Long, ByVal lWndProc As Long)
    
    If (Not (lHwnd = 0)) And (Not (lWndProc = 0)) Then
        SetWindowLong lHwnd, GWL_WNDPROC, lWndProc
    End If
    
End Sub

Public Function WindowInitialize(Optional ByVal lpWndProc As Long = -1, Optional ByVal ClassName As String = "", Optional ByVal WindowName As String = "") As Long

'    Dim Class As WNDCLASS
'
'    Class.style = WS_DISABLED
'    If lpWndProc = -1 Then
'        Class.lpfnwndproc = Val(AddressOf WindowDefaultProc)
'    Else
'        Class.lpfnwndproc = lpWndProc
'    End If
    If ClassName = "" Then
        ClassName = DecaultClassName
    End If
'    If RegisterClass(Class) <> 0 Then
    
       '  RegisterClass
        Dim hwnd As Long
        hwnd = CreateWindowEx(ByVal 0&, ClassName, WindowName, 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)          ' DecaultClassName, WindowName, WS_DISABLED, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, App.hInstance, ByVal 0&)
        
        If lpWndProc = -1 Then
            SetWindowLong hwnd, GWL_WNDPROC, AddressOf WindowDefaultProc
        ElseIf lpWndProc > 0 Then
            SetWindowLong hwnd, GWL_WNDPROC, lpWndProc
        End If
    
'    End If
'
    WindowInitialize = hwnd
    
End Function

Public Sub WindowTerminate(ByRef hwnd As Long)
    If hwnd <> 0 Then
        DestroyWindow hwnd
        hwnd = 0
    End If

End Sub

Public Function WindowDefaultProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    WindowDefaultProc = DefWindowProc(hwnd, uMsg, wParam, lParam)

End Function





