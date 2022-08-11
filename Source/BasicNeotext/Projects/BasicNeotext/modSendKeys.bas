#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modSendKeys"
Option Explicit
    Public Enum Keys
        VK_None = 0
        VK_LButton = 1
        VK_RButton = 2
        VK_Cancel = 3
        VK_MButton = 4
        VK_XButton1 = 5
        VK_XButton2 = 6
        VK_LButton_XButton2 = 7
        VK_Back = 8
        VK_Tab = 9
        VK_LineFeed = 10
        VK_LButton_LineFeed = 11
        VK_Clear = 12
        VK_Return = 13
        VK_RButton_Clear = 14
        VK_RButton_Return = 15
        VK_ShiftKey = 16
        VK_ControlKey = 17
        VK_Menu = 18
        VK_Pause = 19
        VK_Capital = 20
        VK_KanaMode = 21
        VK_RButton_Capital = 22
        VK_JunjaMode = 23
        VK_FinalMode = 24
        VK_HanjaMode = 25
        VK_RButton_FinalMode = 26
        VK_Escape = 27
        VK_IMEConvert = 28
        VK_IMENonconvert = 29
        VK_IMEAceept = 30
        VK_IMEModeChange = 31
        VK_Space = 32
        VK_PageUp = 33
        VK_Next = 34
        VK_End = 35
        VK_Home = 36
        VK_Left = 37
        VK_Up = 38
        VK_Right = 39
        VK_Down = 40
        VK_Select = 41
        VK_Print = 42
        VK_Execute = 43
        VK_PrintScreen = 44
        VK_Insert = 45
        VK_Delete = 46
        VK_Help = 47
        VK_D0 = 48
        VK_D1 = 49
        VK_D2 = 50
        VK_D3 = 51
        VK_D4 = 52
        VK_D5 = 53
        VK_D6 = 54
        VK_D7 = 55
        VK_D8 = 56
        VK_D9 = 57
        VK_RButton_D8 = 58
        VK_RButton_D9 = 59
        VK_MButton_D8 = 60
        VK_MButton_D9 = 61
        VK_XButton2_D8 = 62
        VK_XButton2_D9 = 63
        VK_64 = 64
        VK_A = 65
        VK_B = 66
        VK_C = 67
        VK_D = 68
        VK_E = 69
        VK_F = 70
        VK_G = 71
        VK_H = 72
        VK_I = 73
        VK_J = 74
        VK_K = 75
        VK_L = 76
        VK_M = 77
        VK_N = 78
        VK_O = 79
        VK_P = 80
        VK_Q = 81
        VK_R = 82
        VK_S = 83
        VK_T = 84
        VK_U = 85
        VK_V = 86
        VK_W = 87
        VK_X = 88
        VK_Y = 89
        VK_Z = 90
        VK_LWin = 91
        VK_RWin = 92
        VK_Apps = 93
        VK_RButton_RWin = 94
        VK_Sleep = 95
        VK_NumPad0 = 96
        VK_NumPad1 = 97
        VK_NumPad2 = 98
        VK_NumPad3 = 99
        VK_NumPad4 = 100
        VK_NumPad5 = 101
        VK_NumPad6 = 102
        VK_NumPad7 = 103
        VK_NumPad8 = 104
        VK_NumPad9 = 105
        VK_Multiply = 106
        VK_Add = 107
        VK_Separator = 108
        VK_Subtract = 109
        VK_Decimal = 110
        VK_Divide = 111
        VK_F1 = 112
        VK_F2 = 113
        VK_F3 = 114
        VK_F4 = 115
        VK_F5 = 116
        VK_F6 = 117
        VK_F7 = 118
        VK_F8 = 119
        VK_F9 = 120
        VK_F10 = 121
        VK_F11 = 122
        VK_F12 = 123
        VK_F13 = 124
        VK_F14 = 125
        VK_F15 = 126
        VK_F16 = 127
        VK_F17 = 128
        VK_F18 = 129
        VK_F19 = 130
        VK_F20 = 131
        VK_F21 = 132
        VK_F22 = 133
        VK_F23 = 134
        VK_F24 = 135
        VK_Back_F17 = 136
        VK_Back_F18 = 137
        VK_Back_F19 = 138
        VK_Back_F20 = 139
        VK_Back_F21 = 140
        VK_Back_F22 = 141
        VK_Back_F23 = 142
        VK_Back_F24 = 143
        VK_NumLock = 144
        VK_Scroll = 145
        VK_RButton_NumLock = 146
        VK_RButton_Scroll = 147
        VK_MButton_NumLock = 148
        VK_MButton_Scroll = 149
        VK_XButton2_NumLock = 150
        VK_XButton2_Scroll = 151
        VK_Back_NumLock = 152
        VK_Back_Scroll = 153
        VK_LineFeed_NumLock = 154
        VK_LineFeed_Scroll = 155
        VK_Clear_NumLock = 156
        VK_Clear_Scroll = 157
        VK_RButton_Clear_NumLock = 158
        VK_RButton_Clear_Scroll = 159
        VK_LShiftKey = 160
        VK_RShiftKey = 161
        VK_LControlKey = 162
        VK_RControlKey = 163
        VK_LMenu = 164
        VK_RMenu = 165
        VK_BrowserBack = 166
        VK_BrowserForward = 167
        VK_BrowserRefresh = 168
        VK_BrowserStop = 169
        VK_BrowserSearch = 170
        VK_BrowserFavorites = 171
        VK_BrowserHome = 172
        VK_VolumeMute = 173
        VK_VolumeDown = 174
        VK_VolumeUp = 175
        VK_MediaNextTrack = 176
        VK_MediaPreviousTrack = 177
        VK_MediaStop = 178
        VK_MediaPlayPause = 179
        VK_LaunchMail = 180
        VK_SelectMedia = 181
        VK_LaunchApplication1 = 182
        VK_LaunchApplication2 = 183
        VK_Back_MediaNextTrack = 184
        VK_Back_MediaPreviousTrack = 185
        VK_Oem1 = 186
        VK_Oemplus = 187
        VK_Oemcomma = 188
        VK_OemMinus = 189
        VK_OemPeriod = 190
        VK_OemQuestion = 191
        VK_Oemtilde = 192
        VK_LButton_Oemtilde = 193
        VK_RButton_Oemtilde = 194
        VK_Cancel_Oemtilde = 195
        VK_MButton_Oemtilde = 196
        VK_XButton1_Oemtilde = 197
        VK_XButton2_Oemtilde = 198
        VK_LButton_XButton2_Oemtilde = 199
        VK_Back_Oemtilde = 200
        VK_Tab_Oemtilde = 201
        VK_LineFeed_Oemtilde = 202
        VK_LButton_LineFeed_Oemtilde = 203
        VK_Clear_Oemtilde = 204
        VK_Return_Oemtilde = 205
        VK_RButton_Clear_Oemtilde = 206
        VK_RButton_Return_Oemtilde = 207
        VK_ShiftKey_Oemtilde = 208
        VK_ControlKey_Oemtilde = 209
        VK_Menu_Oemtilde = 210
        VK_Pause_Oemtilde = 211
        VK_Capital_Oemtilde = 212
        VK_KanaMode_Oemtilde = 213
        VK_RButton_Capital_Oemtilde = 214
        VK_JunjaMode_Oemtilde = 215
        VK_FinalMode_Oemtilde = 216
        VK_HanjaMode_Oemtilde = 217
        VK_RButton_FinalMode_Oemtilde = 218
        VK_OemOpenBrackets = 219
        VK_Oem5 = 220
        VK_Oem6 = 221
        VK_Oem7 = 222
        VK_Oem8 = 223
        VK_Space_Oemtilde = 224
        VK_PageUp_Oemtilde = 225
        VK_OemBackslash = 226
        VK_LButton_OemBackslash = 227
        VK_Home_Oemtilde = 228
        VK_ProcessKey = 229
        VK_MButton_OemBackslash = 230
        VK_Packet = 231
        VK_Down_Oemtilde = 232
        VK_Select_Oemtilde = 233
        VK_Back_OemBackslash = 234
        VK_Tab_OemBackslash = 235
        VK_PrintScreen_Oemtilde = 236
        VK_Back_ProcessKey = 237
        VK_Clear_OemBackslash = 238
        VK_Back_Packet = 239
        VK_D0_Oemtilde = 240
        VK_D1_Oemtilde = 241
        VK_ShiftKey_OemBackslash = 242
        VK_ControlKey_OemBackslash = 243
        VK_D4_Oemtilde = 244
        VK_ShiftKey_ProcessKey = 245
        VK_Attn = 246
        VK_Crsel = 247
        VK_Exsel = 248
        VK_EraseEof = 249
        VK_Play = 250
        VK_Zoom = 251
        VK_NoName = 252
        VK_Pa1 = 253
        VK_OemClear = 254
        VK_LButton_OemClear = 255
    End Enum
    Const ASFW_ANY As Long = -1
    Const GW_HWNDNEXT As Long = 2
    Const GA_ROOT As Long = 2
    Const HC_GETNEXT As Long = 1
    Const KEYEVENTF_KEYUP As Long = 2
    Const LSFW_LOCK As Long = 1
    Const LSFW_UNLOCK As Long = 2
    Const MAX_PATH = 260
    Const NEGATIVE As Long = -1
    Const PROCESS_QUERY_INFORMATION = 1024
    Const PROCESS_VM_READ = 16
    Const SW_HIDE As Long = 0
    Const SW_SHOWNORMAL As Long = 1
    Const SW_SHOWMINIMIZED As Long = 2
    Const WH_KEYBOARD_LL As Long = 13
    Const WM_SETTEXT As Long = 12
    Const WM_KEYDOWN As Long = 256
    Const WM_KEYUP As Long = 257
    Const HWND_DESKTOP As Long = 0
    Const HWND_NOTOPMOST As Long = -2
    Const HWND_TOP As Long = 0
    Const HWND_TOPMOST As Long = -1
    Const GW_HWNDFIRST As Long = 0
    Const MOUSEEVENTF_MOVE As Long = 1
    Const MOUSEEVENTF_LEFTDOWN As Long = 2
    Const MOUSEEVENTF_LEFTUP As Long = 4
    Const MOUSEEVENTF_RIGHTDOWN As Long = 8
    Const MOUSEEVENTF_RIGHTUP As Long = 16
    Const MOUSEEVENTF_MIDDLEDOWN As Long = 32
    Const MOUSEEVENTF_MIDDLEUP As Long = 64
    Const MOUSEEVENTF_XDOWN As Long = 128
    Const MOUSEEVENTF_XUP As Long = 256
    Const MOUSEEVENTF_WHEEL As Long = 2048
    Const MOUSEEVENTF_VIRTUALDESK As Long = 16384
    Const MOUSEEVENTF_ABSOLUTE As Long = 32768
    Const QS_ALLQUEUE As Long = 511
    Const SM_CXSCREEN As Long = 0
    Const SM_CYSCREEN As Long = 1
    Const SM_FULLSCREEN As Long = 65535
    Const SWP_NOSIZE As Long = 1
    Const SWP_NOMOVE As Long = 2
    Const SWP_NOACTIVATE As Long = 16
    Const SWP_SHOWWINDOW As Long = 64
    Const WM_COMMAND As Long = 273
    Const WM_LBUTTONDBLCLK As Long = 515
    Const WM_LBUTTONDOWN As Long = 513
    Const WM_LBUTTONUP As Long = 514
    Const WM_MBUTTONDBLCLK As Long = 521
    Const WM_MBUTTONDOWN As Long = 519
    Const WM_MBUTTONUP As Long = 520
    Const WM_RBUTTONDBLCLK As Long = 518
    Const WM_RBUTTONDOWN As Long = 516
    Const WM_RBUTTONUP As Long = 517
    ' An enumeration of different mouse events.
    ' Use in the first parameter of the Click function.
    Public Enum Buttons
        LeftClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
        LeftDoubleClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
        MiddleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
        MiddleDoubleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP + MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
        Move = MOUSEEVENTF_MOVE
        MoveAbsolute = MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE
        RightClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
        RightDoubleClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
        VirtualDesk = MOUSEEVENTF_VIRTUALDESK
        Wheel = MOUSEEVENTF_WHEEL
        xClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
        xDoubleClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP + MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
    End Enum
    ' A group of integer values containing the handle of the main window, and the handle of the focus window.
    Public Type WINFOCUS
        Foreground As Long
        Focus As Long
    End Type
    Private Type EVENTCLICK
        mUp As Boolean
        mDown As Boolean
        mButtons As Long
        X As Long
        Y As Long
        WFOCUS As WINFOCUS
    End Type
    Private Type RECT
        rLeft As Long
        rTop As Long
        rRight As Long
        rBottom As Long
    End Type
    Private Type GUITHREADINFO
         cbSize As Long
         flags As Long
         hWndActive As Long
         hWndFocus As Long
         hWndCapture As Long
         hWndMenuOwner As Long
         hWndMoveSize As Long
         hWndCaret As Long
         rcCaret As RECT
    End Type
    Private Type KBDLLHOOKSTRUCT
       vkCode As Long
       scanCode As Long
       flags As Long
       time As Long
       dwExtraInfo As Long
    End Type
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    Private Type WINNAME
       lpText As String
       lpClass As String
    End Type
    Private Type ITEMINFO
        Width As Long
        Height As Long
        Right As Long
        Left As Long
        Top As Long
        Bottom As Long
        Center As POINTAPI
    End Type
    Private Type MENUINFO
        hwnd As Long
        hMenu As Long
        hSubMenu As Long
    End Type
    Private Type WINSTATE
         IsIconic As Boolean
         IsHidden As Boolean
         IsDisabled As Boolean
         IsChildHidden As Boolean
         IsChildDisabled As Boolean
    End Type
    Private Declare Function apiAllowSetForegroundWindow Lib "user32" Alias "AllowSetForegroundWindow" (ByVal dwProcessId As Long) As Boolean
    Private Declare Function apiAttachThreadInput Lib "user32" Alias "AttachThreadInput" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
    Private Declare Function apiCallNextKeyHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function apiCharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
    Private Declare Function apiChildWindowFromPointEx Lib "user32" Alias "ChildWindowFromPointEx" (ByVal hWndParent As Long, ByVal ptx As Long, ByVal pty As Long, ByVal uFlags As Long) As Long
    Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal Handle As Long) As Long
    Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As KBDLLHOOKSTRUCT, ByVal pSource As Long, ByVal cb As Long) As Long
    Private Declare Function apiEnableWindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Boolean) As Boolean
    Private Declare Function apiEnumProcessModules Lib "PSAPI" Alias "EnumProcessModules" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Private Declare Function apiEnumProcesses Lib "PSAPI" Alias "EnumProcesses" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Private Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function apiFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
    Private Declare Function apiGetAncestor Lib "user32" Alias "GetAncestor" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
    Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function apiGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
    Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Long
    Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
    Private Declare Function apiGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
    Private Declare Function apiGetGUIThreadInfo Lib "user32" Alias "GetGUIThreadInfo" (ByVal dwThreadId As Long, ByRef lpGUIThreadInfo As GUITHREADINFO) As Long
    Private Declare Function apiGetInputState Lib "user32" Alias "GetInputState" () As Long
    Private Declare Function apiGetKeyState Lib "user32" Alias "GetKeyState" (ByVal vKey As Long) As Long
    Private Declare Function apiGetModuleFileNameExA Lib "PSAPI" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
    Private Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
    Private Declare Function apiGetQueueStatus Lib "user32" Alias "GetQueueStatus" (ByVal fuFlags As Long) As Long
    Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
    Private Declare Function apiGetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
    Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function apiGetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function apiGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
    Private Declare Function apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hwnd As Long) As Boolean
    Private Declare Function apiIsWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Boolean
    Private Declare Function apiIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Boolean
    Private Declare Function apiIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Boolean
    Private Declare Function apikeybd_event Lib "user32" Alias "keybd_event" (ByVal vKey As Long, ByVal bScan As Long, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Boolean
    Private Declare Function apiLockSetForegroundWindow Lib "user32" Alias "LockSetForegroundWindow" (ByVal uLockCode As Long) As Boolean
    Private Declare Function apiLockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hWndLock As Long) As Long
    Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
    Private Declare Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Boolean
    Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Boolean
    Private Declare Function apiSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
    Private Declare Function apiSetWindowsKeyHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
    Private Declare Function apiSwitchToThread Lib "kernel32" Alias "SwitchToThread" () As Long
    Private Declare Function apiUnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
    Private Declare Function apiVkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar2 As Long) As Long
    Private Declare Function apiWaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function apiGetMenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
    Private Declare Function apiGetMenuItemCount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
    Private Declare Function apiGetMenuItemID Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Private Declare Function apiGetMenuItemRect Lib "user32" Alias "GetMenuItemRect" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, ByRef lprcItem As RECT) As Long
    Private Declare Function apiGetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Private Declare Function apiGetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
    Private Declare Function apiGetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Private Declare Function apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, ByRef lpRect As RECT) As Boolean
    Private Declare Function apiIsMenu Lib "user32" Alias "IsMenu" (ByVal hMenu As Long) As Boolean
    Private Declare Function apimouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) As Boolean
    Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Boolean
    Private Declare Function apiSetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal X As Long, ByVal Y As Long) As Boolean
    Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Sub apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Private hKey, fWnd, kSent As Long 'Handle for the keyboard, and the main window with directive focus, and the count of the keys sent

    ' Processes any keyboard or mouse messages currently in the queue.
    ' Returns true is there are no messages in the queue to be processed.
    Public Function Flush() As Boolean
        On Error Resume Next
        If apiGetQueueStatus(QS_ALLQUEUE) <> 0 Then DoEvents 'Process any messages in the queue
        Flush = Not CBool(apiGetQueueStatus(QS_ALLQUEUE)) 'Set return value to current state
    End Function

    ' Gets the handle of the current foreground window and the handle of the child window with keyboard focus.
    ' If keyFocus is False then the handles retrieved are the windows under the cursor with mouse focus.
    ' Returns the WINFOCUS structure containing the current focus handles. Depends on the keyFocus parameter.
    ' "showNames"(Optional) Shows the title and class name of the windows,
    ' to be used with the GetWinHandles function.
    ' "keyFocus"(Optional) Returns the handle of the main window, and child window with keyboard focus.
    ' If false then the function returns the handles of the windows currently under the cursor.</param>
    Public Function GetWinFocus(Optional ByVal showNames As Boolean = False, Optional ByVal keyFocus As Boolean = True) As WINFOCUS
        On Error Resume Next
        If keyFocus = True Then ''''''''''''''''''''''Keyboard focus
            Dim g As GUITHREADINFO '''''''''''''''''''Dimension a thread input structure
            g.cbSize = Len(g) ''''''''''''''''''''''''Initialize structure
            Call apiGetGUIThreadInfo(apiGetWindowThreadProcessId(0, 0), g) 'Retrieve information about the active window
            GetWinFocus.Foreground = g.hWndActive ''''Set handle of active foreground
            GetWinFocus.Focus = g.hWndFocus ''''''''''Set handle of focus window
        Else '''''''''''''''''''''''''''''''''''''''''Mouse focus instead
            Dim p As POINTAPI  '''''''''''''''''''''''Dimension a point for the mouse position
            Call apiGetCursorPos(p)  '''''''''''''''''Get current cursor position
            GetWinFocus.Focus = apiWindowFromPoint(p.X, p.Y) 'Get handle of window under cursor
            GetWinFocus.Foreground = apiGetAncestor(GetWinFocus.Focus, GA_ROOT) 'Try to get it's ancestor
            If GetWinFocus.Foreground = 0 Then GetWinFocus.Foreground = GetWinFocus.Focus 'If no ancestor then set main focus to child focus
        End If
        If showNames = True Then Call GetWinAncestory(GetWinFocus.Focus, True) 'Set ancestory using focus window
    End Function

    ' Gets the handle of the main window and the handle of the child window that is to recieve
    ' keyboard or mouse focus, by using the specified title, class name, or index.
    ' Returns the WINFOCUS structure containing the specified focus handles.
    ' "wName1"The main window title or class name.
    ' "wIndex1"(Optional) The main window index.
    ' "wName2"(Optional) The first child window.
    ' "wIndex2"(Optional) The first child index.
    ' "wName3"(Optional) The second child window.
    ' "wIndex3"(Optional) The second child index.
    ' "wName4"(Optional) The third child window.
    ' "wIndex4"(Optional) The third child index.
    ' "wName5"(Optional) The fourth child window.
    ' "wIndex5"(Optional) The fourth child index.
    ' "wName6"(Optional) The fifth child window.
    ' "wIndex6"(Optional) The fifth child index.
    ' "wName7"(Optional) The sixth child window.
    ' "wIndex7"(Optional) The sixth child index.
    ' "wName8"(Optional) The seventh child window.
    ' "wIndex8"(Optional) The seventh child index.
    ' "wName9"(Optional) The eighth child window.
    ' "wIndex9"(Optional) The eighth child index.
    ' "wName10"(Optional) The ninth child window.
    ' "wIndex10"(Optional) The ninth child index.
    Public Function GetWinHandles(ByVal wName1 As String, Optional ByVal wIndex1 As Long = 1, Optional ByVal wName2 As String = " ", Optional ByVal wIndex2 As Long = 1, Optional ByVal wName3 As String = " ", Optional ByVal wIndex3 As Long = 1, Optional ByVal wName4 As String = " ", Optional ByVal wIndex4 As Long = 1, Optional ByVal wName5 As String = " ", Optional ByVal wIndex5 As Long = 1, Optional ByVal wName6 As String = " ", Optional ByVal wIndex6 As Long = 1, Optional ByVal wName7 As String = " ", Optional ByVal wIndex7 As Long = 1, Optional ByVal wName8 As String = " ", Optional ByVal wIndex8 As Long = 1, Optional ByVal wName9 As String = " ", Optional ByVal wIndex9 As Long = 1, Optional ByVal wName10 As String = " ", Optional ByVal wIndex10 As Long = 1) As WINFOCUS
        On Error Resume Next
        Dim hwnd As Long
        Dim cwnd As Long
        Dim i As Long
        Dim wn As String
        Dim wF As WINFOCUS
        Dim wNames() As String '''''''''''''''''''''''Dimension an array string the size of max index
        If wName1 <> " " Then i = 1: ReDim wNames(1) 'Set i to index of parameters
        If wName2 <> " " Then i = 3: ReDim wNames(3)
        If wName3 <> " " Then i = 5: ReDim wNames(5)
        If wName4 <> " " Then i = 7: ReDim wNames(7)
        If wName5 <> " " Then i = 9: ReDim wNames(9)
        If wName6 <> " " Then i = 11: ReDim wNames(11)
        If wName7 <> " " Then i = 13: ReDim wNames(13)
        If wName8 <> " " Then i = 15: ReDim wNames(15)
        If wName9 <> " " Then i = 17: ReDim wNames(17)
        If wName10 <> " " Then i = 19: ReDim wNames(19)
        If wName1 <> " " Then wNames(0) = wName1: wNames(1) = CStr(wIndex1) 'Set array elements
        If wName2 <> " " Then wNames(2) = wName2: wNames(3) = CStr(wIndex2)
        If wName3 <> " " Then wNames(4) = wName3: wNames(5) = CStr(wIndex3)
        If wName4 <> " " Then wNames(6) = wName4: wNames(7) = CStr(wIndex4)
        If wName5 <> " " Then wNames(8) = wName5: wNames(9) = CStr(wIndex5)
        If wName6 <> " " Then wNames(10) = wName6: wNames(11) = CStr(wIndex6)
        If wName7 <> " " Then wNames(12) = wName7: wNames(13) = CStr(wIndex7)
        If wName8 <> " " Then wNames(14) = wName8: wNames(15) = CStr(wIndex8)
        If wName9 <> " " Then wNames(16) = wName9: wNames(17) = CStr(wIndex9)
        If wName10 <> " " Then wNames(18) = wName10: wNames(19) = CStr(wIndex10)
        hwnd = apiFindWindow(vbNullString, wName1) '''Look for handle from title
        If hwnd = 0 Then hwnd = apiFindWindow(wName1, vbNullString) 'If not found by title, then look for handle from class name
        If hwnd <> 0 And CInt(wNames(1)) > 1 Then  '''If searching for window by index, ie more than 1
            Dim nxtwnd As Long
            nxtwnd = hwnd ''''''''''''''''''''''''''''Initialize handle containing the next window in the top level z-order
            Dim Index As Long
            Index = 1 ''''''''''''''''''''''''''''''''Initialize index to 1, since handle was found.
            Dim n As WINNAME  ''''''''''''''''''''''''Create structure for window names
            Do '''''''''''''''''''''''''''''''''''''''Do loop to look for handles matching paramaters
                nxtwnd = apiGetWindow(hwnd, GW_HWNDNEXT) 'Set next window
                If nxtwnd = 0 Then Exit Do '''''''''''Eject if there are no more top level windows
                If wNames(0) <> "" Then ''''''''''''''If title or class name was given
                    n = GetWinName(nxtwnd, True, True) 'Get title and classname of the window
                    If n.lpText = wNames(0) Or n.lpClass = wNames(0) Then 'If title or class name matches first wNames parameter
                        Index = Index + 1 ''''''''''''Increment Index for the matching window
                        hwnd = nxtwnd ''''''''''''''''Set handle to the handle of the matching window
                        If Index >= CInt(wNames(1)) Then Exit Do 'If the specified index has been reached then exit with last matching handle
                    End If
                Else '''''''''''''''''''''''''''''''''Then "" indicates the search is by index only
                    Index = Index + 1 ''''''''''''''''Increment Index for the matching window
                    hwnd = nxtwnd ''''''''''''''''''''Set handle to the handle of the matching window
                    If Index >= CInt(wNames(1)) Then Exit Do 'If index specified has been reached then exit with last matching handle
                End If
            Loop
        End If
        If hwnd = 0 Then '''''''''''''''''''''''''''''If not found by title or class name, then look for special commands
            Dim pId As Long ''''''''''''''''''''''''''Dimension process identification
            Dim p As POINTAPI  '''''''''''''''''''''''Dimension a point structure
            wn = wName1 ''''''''''''''''''''''''''''''Set string to name for lower case conversion
            Call apiCharLower(wn)  '''''''''''''''''''Does not need to be case sensitive
            If wn = "{focus}" Then '''''''''''''''''''If focus window is specified
                hwnd = GetWinFocus.Focus '''''''''''''Get the handle of the focus window
            ElseIf wn = "{foreground}" Then ''''''''''If foreground window is specified
                hwnd = apiGetForegroundWindow ''''''''Get the handle of the foreground window
            ElseIf wn = "{active}" Then ''''''''''''''If active window is specified
                hwnd = apiGetActiveWindow ''''''''''''Get the handle of the active window
            ElseIf wn = "{desktop}" Then '''''''''''''If desktop window is specified
                hwnd = apiGetDesktopWindow() '''''''''Get the handle of the desktop window
            ElseIf wn = "{top}" Then '''''''''''''''''If top window is specified
                hwnd = apiGetTopWindow(hwnd) '''''''''Get the handle of the top window
            ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
                hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
            ElseIf wn = "{parent}" Then ''''''''''''''If parent window is specified
                hwnd = apiGetParent(wIndex1) '''''''''Get the handle of the parent window
            ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
                hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
            ElseIf wn = "{frompoint}" Then '''''''''''If window under the cursor is specified
                If wIndex1 <> 1 And wIndex2 <> 1 Then 'If point specified
                    p.X = wIndex1 ''''''''''''''''''''If x specified
                    p.Y = wIndex2 ''''''''''''''''''''If y specified
                Else
                   Call apiGetCursorPos(p)  ''''''''''Get cursor position as POINTAPI
                End If
                hwnd = apiWindowFromPoint(p.X, p.Y) ''Get the handle of the window under the cursor
            ElseIf wn = "{childfrompoint}" Then ''''''If child window under the cursor is specified
                If wIndex2 <> 1 And wIndex3 <> 1 Then 'If point specified
                    p.X = wIndex2 ''''''''''''''''''''If x specified
                    p.Y = wIndex3 ''''''''''''''''''''If y specified
                Else
                    Call apiGetCursorPos(p)  '''''''''Get cursor position as POINTAPI
                End If
                hwnd = apiChildWindowFromPointEx(wIndex1, p.X, p.Y, wIndex2) 'Get the handle of the child window under the cursor
            ElseIf InStr(1, wName1, ".exe") <> 0 Then 'If process name specified
                pId = ProcessNameToId(wName1) ''''''''Get process Id and wait for input idle
                If pId <> 0 Then hwnd = ProcessIdToHandle(pId) 'If process id found then set handle
            End If
        End If
        If hwnd = 0 Then '''''''''''''''''''''''''''''If no handle found
            pId = Shell(wName1, wIndex1) '''''''''''''Shell(start) the process
            If pId <> 0 Then '''''''''''''''''''''''''If process id found
                hwnd = ProcessIdToHandle(pId) ''''''''Get handle from process identification
                Call ProcessNameToId("", pId) ''''''''Wait for process id to be idle by specifying it in the second parameter
            End If
        End If
        If hwnd = 0 Then hwnd = -1 '''''''''''''''''''If main window not found, then set failure return
        wF.Foreground = hwnd '''''''''''''''''''''''''Set structure return
        wF.Focus = hwnd ''''''''''''''''''''''''''''''Set focus to main handle for now
        If hwnd = -1 Then GetWinHandles = wF: Exit Function ''''''''''''''''''If no handle found then return now
        If i = 1 Then ''''''''''''''''''''''''''''''''If only the main window was specified
            Dim prvswnd As Long
            prvswnd = apiGetForegroundWindow '''''''''Remember the current foreground window
            ForceForeground (hwnd) '''''''''''''''''''Force foreground onto specified main window
            Sleep (0) ''''''''''''''''''''''''''''''''Sleep a moment
            wF = GetWinFocus() '''''''''''''''''''''''Set structure
            ForceForeground (prvswnd) ''''''''''''''''Force foreground back to where it was
        ElseIf i > 1 Then ''''''''''''''''''''''''''''If there are child windows specified
            cwnd = GetChildWindow(hwnd, wNames(2), CInt(wNames(3))) 'Set first child handle
            wF.Focus = cwnd ''''''''''''''''''''''''''Set focus handle to first child
            If (i - 1) > 1 Then ''''''''''''''''''''''If more than one child specified
                Dim q As Long
                For q = 4 To i - 1 Step 2  '''''''''''Step through array looking for grandchildren.
                    cwnd = GetChildWindow(cwnd, wNames(q), CInt(wNames(q + 1))) 'Set new child window
                    wF.Focus = cwnd ''''''''''''''''''Set focus handle to youngest grandchild so far
                Next '''''''''''''''''''''''''''''''''Next in array
            End If
        End If
        GetWinHandles = wF ''''''''''''''''''''''''''''''''''''Return WINFOCUS structure
    End Function

    ' Sleeps for the specified time, while flushing keyboard messages at the specified interval.
    ' Returns true if there are no more messages in the queue.
    ' "dwMilliseconds"(Optional) The number of milliseconds to sleep in milliseconds, - 5000 = 5 seconds.
    ' If 0 is specified then the function sleeps until there are no more input messages in the queue to process or 5 seconds elapses.
    ' Specify an integer.
    ' "fInterval"(Optional) The number of milliseconds between flushes.
    ' Flushing helps process window messages that are still in the queue.  Specify an integer from 0-4999.</param>
    Public Function Sleep(Optional ByVal dwMilliseconds As Long = 0, Optional ByVal fInterval As Long = 1) As Boolean
        On Error Resume Next
        Dim tick As Long
        tick = apiGetTickCount '''''''''''''''''''''''Get the number of millisecons since last log on
        If fInterval > 4999 Then fInterval = 4999 ''''Make sure interval is not longer than 5 seconds to avoid hanging the main app window
        Do '''''''''''''''''''''''''''''''''''''''''''Loop until the specified timeout
            Call apiSleep(fInterval) '''''''''''''''''Sleep for one interval at a time
            Call apiSwitchToThread
            Call Flush  ''''''''''''''''''''''''''''''Process any messages in the queue ''''''
            If dwMilliseconds = 0 Then '''''''''''''''If short rest specified
                If apiGetQueueStatus(QS_ALLQUEUE) = 0 Or Sleep >= (tick + 5000) Then Sleep = True: Exit Function 'If there are no key/mouse messages in the queue or 5 seconds has elapsed then return
            Else '''''''''''''''''''''''''''''''''''''If a timeout is specified
                If apiGetTickCount >= (tick + dwMilliseconds) Then Sleep = Not CBool(apiGetQueueStatus(QS_ALLQUEUE)): Exit Function 'If time is up then exit with return of state
            End If
        Loop
    End Function

    ' Waits for a window to exist, and waits for it to be ready for keyboard and mouse input.
    ' Returns true if a window exists. Returns false if it does not, or the specified time ends.
    ' wTitle The title or class name of the window.  Specify an application title or class name.
    ' dwMilliseconds The number of milliseconds to wait.  Specify an integer.
    Public Function WaitForWindow(ByVal wTitle As String, Optional ByVal dwMilliseconds As Long = 5000) As Boolean
        On Error Resume Next
        Dim tick As Long
        Dim hwnd As Long
        tick = apiGetTickCount '''''''''''''''''''''''Get the number of millisecons since last log on
        hwnd = 0 '''''''''''''''''''''''''''''''''''''Initialize handle
        Do '''''''''''''''''''''''''''''''''''''''''''Loop this thread until it's time
            hwnd = apiFindWindow(vbNullString, wTitle) 'Find window by title
            If hwnd = 0 Then hwnd = apiFindWindow(wTitle, vbNullString) 'If not found by title then find by class name
            If hwnd <> 0 Then WaitForWindow = WaitForWindowIdle(hwnd): Exit Do  'If found, then wait for input idle and return
            If apiGetTickCount >= (tick + dwMilliseconds) Then WaitForWindow = False: Exit Do  'If time is up then return that value
            apiSleep (1) '''''''''''''''''''''''''''''Sleep thread for one millisecond
            Flush ''''''''''''''''''''''''''''''''''''Process messages in the queue
        Loop
    End Function

    ' Sends or posts key messages directly to a window.  Use GetWinHandles or GetWinFocus to get a structure.
    ' Returns true if successful.
    ' vKey The virtual keycode to send.  Use Keys.VK_ enumeration.
    ' wFocus The structure of the window.  Use GetWinHandles or GetWinFocus to get a structure.
    ' kDown (Optional) Presses the key down.  Specify true or false.
    ' kUp (Optional) Lifts key the up.  Specify true or false.
    ' bPost (Optional) Posts the message into the queue instead of sending it.  Specify true or false.
    Public Function Message(ByVal vKey As Long, ByRef WFOCUS As WINFOCUS, Optional ByVal kDown As Boolean = True, Optional ByVal kUp As Boolean = True, Optional ByVal bPost As Boolean = False) As Boolean
        On Error Resume Next
        Message = True '''''''''''''''''''''''''''''''Set default return value
        If bPost = True Then '''''''''''''''''''''''''If post specified
            If kDown = True Then
                If apiPostMessage(WFOCUS.Focus, WM_KEYDOWN, vKey, vbNullString) = False Then Message = False 'If key is to be pressed down
            End If
            If kUp = True Then
                If apiPostMessage(WFOCUS.Focus, WM_KEYUP, vKey, vbNullString) = False Then Message = False 'If key is to be lifted up
            End If
        Else '''''''''''''''''''''''''''''''''''''''''If Sending message
            If kDown = True Then
                If apiSendMessage(WFOCUS.Focus, WM_KEYDOWN, vKey, vbNullString) = False Then Message = False 'If key is to be pressed down
            End If
            If kUp = True Then
                If apiSendMessage(WFOCUS.Focus, WM_KEYUP, vKey, vbNullString) = False Then Message = False 'If key is to be lifted up
            End If
        End If
    End Function

    ' Sends keyboard messages or events to the specified window.
    ' If asMessage is false, then this function simulates keyboard events,
    ' and directs keyboard focus onto the specified window structure.
    ' Use GetWinHandles or GetWinFocus to get a structure.
    ' Returns the number of keys sent.  Returns 0 only if the keyboard could not be hooked,
    ' or a valid window handle cannot be found, or the foregroundwindow cannot be forced,
    ' or focus could not be set, or the keys cannot be simulated.
    ' "cText"The text or key commands to send.  See the GetCommands sub,
    ' for more details about commands.
    ' "wFocus"The structure of the window to focus on.
    ' Use GetWinHandles or GetWinFocus to get a structure.
    ' "asMessage"(Optional) Send as window message, or as a keyboard event.
    ' It's recommended that you use True in this parameter for writing text. If false then rForeground applies.</param>
    ' "rForeground"(Optional) Returns the foreground window to it's previous focus
    ' before the keys were sent. Applies only if asMessage is false.
    Public Function Send(ByVal cText As String, ByRef WFOCUS As WINFOCUS, Optional ByVal asMessage As Boolean = True, Optional ByVal rForeground As Boolean = False) As Long
        On Error Resume Next
        Dim ws As WINSTATE
        Dim UserCapsLock As Boolean
        Dim sThread As Boolean
        Dim IsLetter As Boolean ''''''''''''''''''''''Declare some switches
        Dim hwnd As Long
        Dim cwnd As Long
        Dim prvWnd As Long
        Dim txtLength As Long
        Dim VK As Long
        Dim repeat As Long '''''''''''''''''''''''''''Dimension some integers
        Dim txtRemain As String
        Dim Letter As String
        Dim PrevLetter As String '''''''''''''''''''''Dimension some strings
        Dim LowerCase As String
        txtRemain = cText ''''''''''''''''''''''''''''Initialize remaining text
        txtLength = Len(cText) '''''''''''''''''''''''Initialize Text
        If WFOCUS.Foreground = -1 Then Send = 0: Exit Function '''''Exit failure if GetWinHandles returned -1
        If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus()
        hwnd = WFOCUS.Foreground '''''''''''''''''''''Read structure handles, and make easier to use
        cwnd = WFOCUS.Focus ''''''''''''''''''''''''''Set child handle
        repeat = 1 '''''''''''''''''''''''''''''''''''Initialize repeat to one, for commands
        Letter = "" ''''''''''''''''''''''''''''''''''Initialize the first letter to nothing
        If cwnd = 0 Then cwnd = hwnd '''''''''''''''''If no child specified, then set it to main window
        If apiIsWindow(hwnd) = False Then Send = 0: Exit Function ''Make sure it's a valid window, if not abort
        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL) 'If window is minimized then show it, and remember
        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True) 'If window is disabled then enable it and remember
        If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True) 'If child window is disabled then enable it and remember
        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL) 'If window is hidden then show it, and remember
        If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL) 'If child window is hidden then show it, and remember
        WaitForWindowIdle (hwnd) '''''''''''''''''''''Wait for input idle
        If apiGetKeyState(VK_Capital) = 1 Then '''''''Get state of the caps lock key.
            UserCapsLock = KeyEvent(VK_Capital) ''''''Remember state if toggled, by return of keyevent
            Call apiSleep(0)  ''''''''''''''''''''''''Wait for other threads
        End If
        Call Lift   ''''''''''''''''''''''''''''''''''Lift shift, control and menu keys
        If asMessage = False Then ''''''''''''''''''''Send as chain of events
            hKey = apiSetWindowsKeyHookEx(WH_KEYBOARD_LL, AddressOf Callback, App.hInstance, 0)
            If hKey = 0 Then Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, False, prvWnd): Exit Function
            If apiGetWindowThreadProcessId(hwnd, 0) = apiGetCurrentThreadId Then sThread = True
            If sThread = False Then ''''''''''''''''''If handle is a different thread than this one
                sThread = Not AttachInput(hwnd, True) 'Attach thread input to this thread
            Else '''''''''''''''''''''''''''''''''''''If the same thread
                rForeground = False ''''''''''''''''''No reason to reset foreground window, because focus never leaves to begin with
            End If
            prvWnd = apiGetForegroundWindow ''''''''''Remember current foreground window
            If ForceForeground(hwnd) = False Then Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, False, prvWnd): Exit Function 'If the foreground cannot be set within 5 seconds then abort
            If cwnd <> hwnd Then '''''''''''''''''''''If child window is different than main window
                If apiSetFocus(cwnd) = 0 Then Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, rForeground, prvWnd): Exit Function 'If focus window cannot be set then abort
            End If
            fWnd = hwnd ''''''''''''''''''''''''''''''Set form class variable for the foregroundwindow handle
        Else '''''''''''''''''''''''''''''''''''''''''If message
            sThread = True '''''''''''''''''''''''''''Set same thread to true, so it doesn't detach for no reason, since it was never set
            rForeground = False ''''''''''''''''''''''Foreground is not set, so no need to return it
            If hwnd = cwnd Then ''''''''''''''''''''''If the main window is the same as child
                If ForceForeground(hwnd) = False Then 'If Force foreground window fails
                    Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, False, prvWnd): Exit Function 'return abort's return
                End If
                Sleep (0) '''''''' '''''''''''''''''''Sleep a moment and flush any messages so that the child window if any, gets focus.
                WFOCUS = GetWinFocus() '''''''''''''''Set by getting the WINFOCUS structure
                rForeground = True '''''''''''''''''''Should return it
            End If
        End If
        Call Lift   ''''''''''''''''''''''''''''''''''Lift shift, control and menu keys
        Dim i As Long
        For i = 1 To txtLength  ''''''''''''''''''''''Loop through each character of the text specified
            Dim cmdKey As String
            cmdKey = "" ''''''''''''''''''''''''''''''Initialize a string for special command keys
            PrevLetter = Letter ''''''''''''''''''''''Set previous letter to the current before setting the current
            Letter = Left(txtRemain, 1) ''''''''''''''Set current letter to the left most
            LowerCase = Letter '''''''''''''''''''''''Initialize lower case letter
            txtRemain = Mid(txtRemain, 2) ''''''''''''Cut string to get remaining text
            VK = apiVkKeyScan(Asc(Letter)) And 255 '''Set keyscan of that letter
            apiCharLower (LowerCase) '''''''''''''''''Convert in place to lowercase
            If Letter <> LowerCase Or Letter = "!" Or Letter = "@" Or Letter = "$" Or Letter = "&" Or Letter = "*" Or Letter = "_" Or Letter = "|" Or Letter = ":" Or Letter = "<" Or Letter = ">" Or Letter = "?" Then
                Call KeyEvent(VK_ShiftKey, True, False) 'Press shift if necessary
            End If
            If Letter = "{" Then '''''''''''''''''''''If letter is a command bracket
                VK = 0 '''''''''''''''''''''''''''''''Set virtual key to zero(none)
                cmdKey = Mid(txtRemain, 1, InStr(1, txtRemain, "}") - 1) 'Get text within command brackets
                If Len(txtRemain) - (Len(cmdKey) + 1) > 0 Then txtRemain = Mid(txtRemain, InStr(1, txtRemain, "}") + 1) ' 'Set remaining text, as the text(if any) to the right of the command bracket
                i = i + Len(cmdKey) + 1 ''''''''''''''Increment i the number of characters within command brackets since they are not going to be processed individually
                If InStr(1, cmdKey, " ") <> 0 And InStr(1, cmdKey, ", ") = 0 Then 'If there is a space and no comma preceeding it then the command is to be repeated
                    If IsNumeric(Mid(cmdKey, InStr(1, cmdKey, " ") + 1)) Then repeat = CInt(Mid(cmdKey, InStr(1, cmdKey, " ") + 1)) 'If a number can be identified, then set it otherwise it stays 1._+{}|:"<>?
                    cmdKey = Left(cmdKey, InStr(1, cmdKey, " ") - 1) 'Strip off left side, which is the actual command
                End If
                If Len(cmdKey) = 1 Then ''''''''''''''If command is a single character
                    If cmdKey = "#" Or cmdKey = "+" Or cmdKey = "^" Or cmdKey = "%" Or cmdKey = "~" Or cmdKey = "(" Or cmdKey = ")" Or cmdKey = "[" Or cmdKey = "]" Then
                        Call KeyEvent(VK_ShiftKey, True, False) 'Press shift for single character commands
                    End If
                    VK = apiVkKeyScan(Asc(cmdKey)) And 255 'Set keyscan for that letter
                Else '''''''''''''''''''''''''''''''''If command is a key name
                    VK = NEGATIVE ''''''''''''''''''''Set to negative in case this is not a valid key command, then it's a simple string
                    apiCharLower (cmdKey)
                    If cmdKey = "none" Then ''''''''''Set virtual key to specification
                        VK = VK_None
                    ElseIf cmdKey = "lbutton" Then
                        VK = VK_LButton
                    ElseIf cmdKey = "rbutton" Then
                        VK = VK_RButton
                    ElseIf cmdKey = "cancel" Then
                        VK = VK_Cancel
                    ElseIf cmdKey = "mbutton" Then
                        VK = VK_MButton
                    ElseIf cmdKey = "xbutton1" Then
                        VK = VK_XButton1
                    ElseIf cmdKey = "xbutton2" Then
                        VK = VK_XButton2
                    ElseIf cmdKey = "lbutton, xbutton2" Then
                        VK = VK_LButton_XButton2
                    ElseIf cmdKey = "back" Or cmdKey = "backspace" Or cmdKey = "bs" Or cmdKey = "bksp" Then
                        VK = VK_Back
                    ElseIf cmdKey = "tab" Then
                        VK = VK_Tab
                    ElseIf cmdKey = "linefeed" Then
                        VK = VK_LineFeed
                    ElseIf cmdKey = "lbutton, linefeed" Then
                        VK = VK_LButton_LineFeed
                    ElseIf cmdKey = "clear" Then
                        VK = VK_Clear
                    ElseIf cmdKey = "return" Or cmdKey = "enter" Then
                        VK = VK_Return
                    ElseIf cmdKey = "rbutton, clear" Then
                        VK = VK_RButton_Clear
                    ElseIf cmdKey = "rbutton, return" Then
                        VK = VK_RButton_Return
                    ElseIf cmdKey = "shiftkey" Or cmdKey = "shift" Then
                        VK = VK_ShiftKey
                    ElseIf cmdKey = "controlkey" Or cmdKey = "control" Then
                        VK = VK_ControlKey
                    ElseIf cmdKey = "menu" Or cmdKey = "alt" Then
                        VK = VK_Menu
                    ElseIf cmdKey = "pause" Or cmdKey = "break" Then
                        VK = VK_Pause
                    ElseIf cmdKey = "capital" Or cmdKey = "capslock" Then
                        VK = VK_Capital
                    ElseIf cmdKey = "kanamode" Then
                        VK = VK_KanaMode
                    ElseIf cmdKey = "rbutton, capital" Then
                        VK = VK_RButton_Capital
                    ElseIf cmdKey = "junjamode" Then
                        VK = VK_JunjaMode
                    ElseIf cmdKey = "finalmode" Then
                        VK = VK_FinalMode
                    ElseIf cmdKey = "hanjamode" Then
                        VK = VK_HanjaMode
                    ElseIf cmdKey = "rbutton, finalmode" Then
                        VK = VK_RButton_FinalMode
                    ElseIf cmdKey = "escape" Or cmdKey = "esc" Then
                        VK = VK_Escape
                    ElseIf cmdKey = "imeconvert" Then
                        VK = VK_IMEConvert
                    ElseIf cmdKey = "imenonconvert" Then
                        VK = VK_IMENonconvert
                    ElseIf cmdKey = "imeaceept" Then
                        VK = VK_IMEAceept
                    ElseIf cmdKey = "imemodechange" Then
                        VK = VK_IMEModeChange
                    ElseIf cmdKey = "space" Then
                        VK = VK_Space
                    ElseIf cmdKey = "pageup" Or cmdKey = "pgup" Then
                        VK = VK_PageUp
                    ElseIf cmdKey = "next" Then
                        VK = VK_Next
                    ElseIf cmdKey = "end" Then
                        VK = VK_End
                    ElseIf cmdKey = "home" Then
                        VK = VK_Home
                    ElseIf cmdKey = "left" Then
                        VK = VK_Left
                    ElseIf cmdKey = "up" Then
                        VK = VK_Up
                    ElseIf cmdKey = "right" Then
                        VK = VK_Right
                    ElseIf cmdKey = "down" Then
                        VK = VK_Down
                    ElseIf cmdKey = "select" Then
                        VK = VK_Select
                    ElseIf cmdKey = "print" Then
                        VK = VK_Print
                    ElseIf cmdKey = "execute" Then
                        VK = VK_Execute
                    ElseIf cmdKey = "printscreen" Or cmdKey = "prtsc" Or cmdKey = "snapshot" Then
                        VK = VK_PrintScreen
                    ElseIf cmdKey = "Insert" Or cmdKey = "ins" Then
                        VK = VK_Insert
                    ElseIf cmdKey = "delete" Or cmdKey = "del" Then
                        VK = VK_Delete
                    ElseIf cmdKey = "help" Then
                        VK = VK_Help
                    ElseIf cmdKey = "d0" Then
                        VK = VK_D0
                    ElseIf cmdKey = "d1" Then
                        VK = VK_D1
                    ElseIf cmdKey = "d2" Then
                        VK = VK_D2
                    ElseIf cmdKey = "d3" Then
                        VK = VK_D3
                    ElseIf cmdKey = "d4" Then
                        VK = VK_D4
                    ElseIf cmdKey = "d5" Then
                        VK = VK_D5
                    ElseIf cmdKey = "d6" Then
                        VK = VK_D6
                    ElseIf cmdKey = "d7" Then
                        VK = VK_D7
                    ElseIf cmdKey = "d8" Then
                        VK = VK_D8
                    ElseIf cmdKey = "d9" Then
                        VK = VK_D9
                    ElseIf cmdKey = "rbutton, d8" Then
                        VK = VK_RButton_D8
                    ElseIf cmdKey = "rbutton, d9" Then
                        VK = VK_RButton_D9
                    ElseIf cmdKey = "mbutton, d8" Then
                        VK = VK_MButton_D8
                    ElseIf cmdKey = "mbutton, d9" Then
                        VK = VK_MButton_D9
                    ElseIf cmdKey = "xbutton2, d8" Then
                        VK = VK_XButton2_D8
                    ElseIf cmdKey = "xbutton2, d9" Then
                        VK = VK_XButton2_D9
                    ElseIf cmdKey = "64" Then
                        VK = VK_64
                    ElseIf cmdKey = "a" Then
                        VK = VK_A
                    ElseIf cmdKey = "b" Then
                        VK = VK_B
                    ElseIf cmdKey = "c" Then
                        VK = VK_C
                    ElseIf cmdKey = "d" Then
                        VK = VK_D
                    ElseIf cmdKey = "e" Then
                        VK = VK_E
                    ElseIf cmdKey = "f" Then
                        VK = VK_F
                    ElseIf cmdKey = "g" Then
                        VK = VK_G
                    ElseIf cmdKey = "h" Then
                        VK = VK_H
                    ElseIf cmdKey = "i" Then
                        VK = VK_I
                    ElseIf cmdKey = "j" Then
                        VK = VK_J
                    ElseIf cmdKey = "k" Then
                        VK = VK_K
                    ElseIf cmdKey = "l" Then
                        VK = VK_L
                    ElseIf cmdKey = "m" Then
                        VK = VK_M
                    ElseIf cmdKey = "n" Then
                        VK = VK_N
                    ElseIf cmdKey = "o" Then
                        VK = VK_O
                    ElseIf cmdKey = "p" Then
                        VK = VK_P
                    ElseIf cmdKey = "q" Then
                        VK = VK_Q
                    ElseIf cmdKey = "r" Then
                        VK = VK_R
                    ElseIf cmdKey = "s" Then
                        VK = VK_S
                    ElseIf cmdKey = "t" Then
                        VK = VK_T
                    ElseIf cmdKey = "u" Then
                        VK = VK_U
                    ElseIf cmdKey = "v" Then
                        VK = VK_V
                    ElseIf cmdKey = "w" Then
                        VK = VK_W
                    ElseIf cmdKey = "X" Then
                        VK = VK_X
                    ElseIf cmdKey = "Y" Then
                        VK = VK_Y
                    ElseIf cmdKey = "z" Then
                        VK = VK_Z
                    ElseIf cmdKey = "lwin" Then
                        VK = VK_LWin
                    ElseIf cmdKey = "rwin" Then
                        VK = VK_RWin
                    ElseIf cmdKey = "apps" Then
                        VK = VK_Apps
                    ElseIf cmdKey = "rbutton, rwin" Then
                        VK = VK_RButton_RWin
                    ElseIf cmdKey = "Sleep" Then
                        VK = VK_Sleep
                    ElseIf cmdKey = "numpad0" Then
                        VK = VK_NumPad0
                    ElseIf cmdKey = "numpad1" Then
                        VK = VK_NumPad1
                    ElseIf cmdKey = "numpad2" Then
                        VK = VK_NumPad2
                    ElseIf cmdKey = "numpad3" Then
                        VK = VK_NumPad3
                    ElseIf cmdKey = "numpad4" Then
                        VK = VK_NumPad4
                    ElseIf cmdKey = "numpad5" Then
                        VK = VK_NumPad5
                    ElseIf cmdKey = "numpad6" Then
                        VK = VK_NumPad6
                    ElseIf cmdKey = "numpad7" Then
                        VK = VK_NumPad7
                    ElseIf cmdKey = "numpad8" Then
                        VK = VK_NumPad8
                    ElseIf cmdKey = "numpad9" Then
                        VK = VK_NumPad9
                    ElseIf cmdKey = "multiply" Then
                        VK = VK_Multiply
                    ElseIf cmdKey = "add" Then
                        VK = VK_Add
                    ElseIf cmdKey = "separator" Then
                        VK = VK_Separator
                    ElseIf cmdKey = "subtract" Then
                        VK = VK_Subtract
                    ElseIf cmdKey = "decimal" Then
                        VK = VK_Decimal
                    ElseIf cmdKey = "divide" Then
                        VK = VK_Divide
                    ElseIf cmdKey = "f1" Then
                        VK = VK_F1
                    ElseIf cmdKey = "f2" Then
                        VK = VK_F2
                    ElseIf cmdKey = "f3" Then
                        VK = VK_F3
                    ElseIf cmdKey = "f4" Then
                        VK = VK_F4
                    ElseIf cmdKey = "f5" Then
                        VK = VK_F5
                    ElseIf cmdKey = "f6" Then
                        VK = VK_F6
                    ElseIf cmdKey = "f7" Then
                        VK = VK_F7
                    ElseIf cmdKey = "f8" Then
                        VK = VK_F8
                    ElseIf cmdKey = "f9" Then
                        VK = VK_F9
                    ElseIf cmdKey = "f10" Then
                        VK = VK_F10
                    ElseIf cmdKey = "f11" Then
                        VK = VK_F11
                    ElseIf cmdKey = "f12" Then
                        VK = VK_F12
                    ElseIf cmdKey = "f13" Then
                        VK = VK_F13
                    ElseIf cmdKey = "f14" Then
                        VK = VK_F14
                    ElseIf cmdKey = "f15" Then
                        VK = VK_F15
                    ElseIf cmdKey = "f16" Then
                        VK = VK_F16
                    ElseIf cmdKey = "f17" Then
                        VK = VK_F17
                    ElseIf cmdKey = "f18" Then
                        VK = VK_F18
                    ElseIf cmdKey = "f19" Then
                        VK = VK_F19
                    ElseIf cmdKey = "f20" Then
                        VK = VK_F20
                    ElseIf cmdKey = "f21" Then
                        VK = VK_F21
                    ElseIf cmdKey = "f22" Then
                        VK = VK_F22
                    ElseIf cmdKey = "f23" Then
                        VK = VK_F23
                    ElseIf cmdKey = "f24" Then
                        VK = VK_F24
                    ElseIf cmdKey = "back, f17" Then
                        VK = VK_Back_F17
                    ElseIf cmdKey = "back, f18" Then
                        VK = VK_Back_F18
                    ElseIf cmdKey = "back, f19" Then
                        VK = VK_Back_F19
                    ElseIf cmdKey = "back, f20" Then
                        VK = VK_Back_F20
                    ElseIf cmdKey = "back, f21" Then
                        VK = VK_Back_F21
                    ElseIf cmdKey = "back, f22" Then
                        VK = VK_Back_F22
                    ElseIf cmdKey = "back, f23" Then
                        VK = VK_Back_F23
                    ElseIf cmdKey = "back, f24" Then
                        VK = VK_Back_F24
                    ElseIf cmdKey = "numlock" Then
                        VK = VK_NumLock
                    ElseIf cmdKey = "scroll" Or cmdKey = "scrolllock" Then
                        VK = VK_Scroll
                    ElseIf cmdKey = "rbutton, numlock" Then
                        VK = VK_RButton_NumLock
                    ElseIf cmdKey = "rbutton, scroll" Then
                        VK = VK_RButton_Scroll
                    ElseIf cmdKey = "mbutton, numlock" Then
                        VK = VK_MButton_NumLock
                    ElseIf cmdKey = "mbutton, scroll" Then
                        VK = VK_MButton_Scroll
                    ElseIf cmdKey = "xbutton2, numlock" Then
                        VK = VK_XButton2_NumLock
                    ElseIf cmdKey = "xbutton2, scroll" Then
                        VK = VK_XButton2_Scroll
                    ElseIf cmdKey = "back, numlock" Then
                        VK = VK_Back_NumLock
                    ElseIf cmdKey = "back, scroll" Then
                        VK = VK_Back_Scroll
                    ElseIf cmdKey = "linefeed, numlock" Then
                        VK = VK_LineFeed_NumLock
                    ElseIf cmdKey = "linefeed, scroll" Then
                        VK = VK_LineFeed_Scroll
                    ElseIf cmdKey = "clear, numlock" Then
                        VK = VK_Clear_NumLock
                    ElseIf cmdKey = "clear, scroll" Then
                        VK = VK_Clear_Scroll
                    ElseIf cmdKey = "rbutton, clear, numlock" Then
                        VK = VK_RButton_Clear_NumLock
                    ElseIf cmdKey = "rbutton, clear, scroll" Then
                        VK = VK_RButton_Clear_Scroll
                    ElseIf cmdKey = "lshiftkey" Or cmdKey = "lshift" Then
                        VK = VK_LShiftKey
                    ElseIf cmdKey = "rshiftkey" Or cmdKey = "rshift" Then
                        VK = VK_RShiftKey
                    ElseIf cmdKey = "lcontrolkey" Or cmdKey = "lcontrol" Then
                        VK = VK_LControlKey
                    ElseIf cmdKey = "rcontrolkey" Or cmdKey = "rcontrol" Then
                        VK = VK_RControlKey
                    ElseIf cmdKey = "lmenu" Then
                        VK = VK_LMenu
                    ElseIf cmdKey = "rmenu" Then
                        VK = VK_RMenu
                    ElseIf cmdKey = "browserback" Then
                        VK = VK_BrowserBack
                    ElseIf cmdKey = "browserforward" Then
                        VK = VK_BrowserForward
                    ElseIf cmdKey = "browserrefresh" Then
                        VK = VK_BrowserRefresh
                    ElseIf cmdKey = "browserstop" Then
                        VK = VK_BrowserStop
                    ElseIf cmdKey = "browsersearch" Then
                        VK = VK_BrowserSearch
                    ElseIf cmdKey = "browserfavorites" Then
                        VK = VK_BrowserFavorites
                    ElseIf cmdKey = "browserhome" Then
                        VK = VK_BrowserHome
                    ElseIf cmdKey = "volumemute" Then
                        VK = VK_VolumeMute
                    ElseIf cmdKey = "volumedown" Then
                        VK = VK_VolumeDown
                    ElseIf cmdKey = "volumeup" Then
                        VK = VK_VolumeUp
                    ElseIf cmdKey = "medianexttrack" Then
                        VK = VK_MediaNextTrack
                    ElseIf cmdKey = "mediaprevioustrack" Then
                        VK = VK_MediaPreviousTrack
                    ElseIf cmdKey = "mediastop" Then
                        VK = VK_MediaStop
                    ElseIf cmdKey = "mediaplaypause" Then
                        VK = VK_MediaPlayPause
                    ElseIf cmdKey = "launchmail" Then
                        VK = VK_LaunchMail
                    ElseIf cmdKey = "selectmedia" Then
                        VK = VK_SelectMedia
                    ElseIf cmdKey = "launchapplication1" Then
                        VK = VK_LaunchApplication1
                    ElseIf cmdKey = "launchapplication2" Then
                        VK = VK_LaunchApplication2
                    ElseIf cmdKey = "back, medianexttrack" Then
                        VK = VK_Back_MediaNextTrack
                    ElseIf cmdKey = "back, mediaprevioustrack" Then
                        VK = VK_Back_MediaPreviousTrack
                    ElseIf cmdKey = "oem1" Then
                        VK = VK_Oem1
                    ElseIf cmdKey = "oemplus" Then
                        VK = VK_Oemplus
                    ElseIf cmdKey = "oemcomma" Then
                        VK = VK_Oemcomma
                    ElseIf cmdKey = "oemminus" Then
                        VK = VK_OemMinus
                    ElseIf cmdKey = "oemperiod" Then
                        VK = VK_OemPeriod
                    ElseIf cmdKey = "oemquestion" Then
                        VK = VK_OemQuestion
                    ElseIf cmdKey = "oemtilde" Then
                        VK = VK_Oemtilde
                    ElseIf cmdKey = "lbutton, oemtilde" Then
                        VK = VK_LButton_Oemtilde
                    ElseIf cmdKey = "rbutton, oemtilde" Then
                        VK = VK_RButton_Oemtilde
                    ElseIf cmdKey = "cancel, oemtilde" Then
                        VK = VK_Cancel_Oemtilde
                    ElseIf cmdKey = "mbutton, oemtilde" Then
                        VK = VK_MButton_Oemtilde
                    ElseIf cmdKey = "xbutton1, oemtilde" Then
                        VK = VK_XButton1_Oemtilde
                    ElseIf cmdKey = "xbutton2, oemtilde" Then
                        VK = VK_XButton2_Oemtilde
                    ElseIf cmdKey = "lbutton, xbutton2, oemtilde" Then
                        VK = VK_LButton_XButton2_Oemtilde
                    ElseIf cmdKey = "back, oemtilde" Then
                        VK = VK_Back_Oemtilde
                    ElseIf cmdKey = "tab, oemtilde" Then
                        VK = VK_Tab_Oemtilde
                    ElseIf cmdKey = "linefeed, oemtilde" Then
                        VK = VK_LineFeed_Oemtilde
                    ElseIf cmdKey = "lbutton, linefeed, oemtilde" Then
                        VK = VK_LButton_LineFeed_Oemtilde
                    ElseIf cmdKey = "clear, oemtilde" Then
                        VK = VK_Clear_Oemtilde
                    ElseIf cmdKey = "return, oemtilde" Then
                        VK = VK_Return_Oemtilde
                    ElseIf cmdKey = "rbutton, clear, oemtilde" Then
                        VK = VK_RButton_Clear_Oemtilde
                    ElseIf cmdKey = "rbutton, return, oemtilde" Then
                        VK = VK_RButton_Return_Oemtilde
                    ElseIf cmdKey = "shiftkey, oemtilde" Then
                        VK = VK_ShiftKey_Oemtilde
                    ElseIf cmdKey = "controlkey, oemtilde" Then
                        VK = VK_ControlKey_Oemtilde
                    ElseIf cmdKey = "menu, oemtilde" Then
                        VK = VK_Menu_Oemtilde
                    ElseIf cmdKey = "pause, oemtilde" Then
                        VK = VK_Pause_Oemtilde
                    ElseIf cmdKey = "capital, oemtilde" Then
                        VK = VK_Capital_Oemtilde
                    ElseIf cmdKey = "kanamode, oemtilde" Then
                        VK = VK_KanaMode_Oemtilde
                    ElseIf cmdKey = "rbutton, capital, oemtilde" Then
                        VK = VK_RButton_Capital_Oemtilde
                    ElseIf cmdKey = "junjamode, oemtilde" Then
                        VK = VK_JunjaMode_Oemtilde
                    ElseIf cmdKey = "finalmode, oemtilde" Then
                        VK = VK_FinalMode_Oemtilde
                    ElseIf cmdKey = "hanjamode, oemtilde" Then
                        VK = VK_HanjaMode_Oemtilde
                    ElseIf cmdKey = "rbutton, finalmode, oemtilde" Then
                        VK = VK_RButton_FinalMode_Oemtilde
                    ElseIf cmdKey = "oemopenbrackets" Then
                        VK = VK_OemOpenBrackets
                    ElseIf cmdKey = "oem5" Then
                        VK = VK_Oem5
                    ElseIf cmdKey = "oem6" Then
                        VK = VK_Oem6
                    ElseIf cmdKey = "oem7" Then
                        VK = VK_Oem7
                    ElseIf cmdKey = "oem8" Then
                        VK = VK_Oem8
                    ElseIf cmdKey = "space, oemtilde" Then
                        VK = VK_Space_Oemtilde
                    ElseIf cmdKey = "pageup, oemtilde" Then
                        VK = VK_PageUp_Oemtilde
                    ElseIf cmdKey = "oembackslash" Then
                        VK = VK_OemBackslash
                    ElseIf cmdKey = "lbutton, oembackslash" Then
                        VK = VK_LButton_OemBackslash
                    ElseIf cmdKey = "home, oemtilde" Then
                        VK = VK_Home_Oemtilde
                    ElseIf cmdKey = "processkey" Then
                        VK = VK_ProcessKey
                    ElseIf cmdKey = "mbutton, oembackslash" Then
                        VK = VK_MButton_OemBackslash
                    ElseIf cmdKey = "packet" Then
                        VK = VK_Packet
                    ElseIf cmdKey = "down, oemtilde" Then
                        VK = VK_Down_Oemtilde
                    ElseIf cmdKey = "select, oemtilde" Then
                        VK = VK_Select_Oemtilde
                    ElseIf cmdKey = "back, oembackslash" Then
                        VK = VK_Back_OemBackslash
                    ElseIf cmdKey = "tab, oembackslash" Then
                        VK = VK_Tab_OemBackslash
                    ElseIf cmdKey = "printscreen, oemtilde" Then
                        VK = VK_PrintScreen_Oemtilde
                    ElseIf cmdKey = "back, processkey" Then
                        VK = VK_Back_ProcessKey
                    ElseIf cmdKey = "clear, oembackslash" Then
                        VK = VK_Clear_OemBackslash
                    ElseIf cmdKey = "back, packet" Then
                        VK = VK_Back_Packet
                    ElseIf cmdKey = "d0, oemtilde" Then
                        VK = VK_D0_Oemtilde
                    ElseIf cmdKey = "d1, oemtilde" Then
                        VK = VK_D1_Oemtilde
                    ElseIf cmdKey = "shiftkey, oembackslash" Then
                        VK = VK_ShiftKey_OemBackslash
                    ElseIf cmdKey = "controlkey, oembackslash" Then
                        VK = VK_ControlKey_OemBackslash
                    ElseIf cmdKey = "d4, oemtilde" Then
                        VK = VK_D4_Oemtilde
                    ElseIf cmdKey = "shiftkey, processkey" Then
                        VK = VK_ShiftKey_ProcessKey
                    ElseIf cmdKey = "attn" Then
                        VK = VK_Attn
                    ElseIf cmdKey = "crsel" Then
                        VK = VK_Crsel
                    ElseIf cmdKey = "exsel" Then
                        VK = VK_Exsel
                    ElseIf cmdKey = "eraseeof" Then
                        VK = VK_EraseEof
                    ElseIf cmdKey = "play" Then
                        VK = VK_Play
                    ElseIf cmdKey = "zoom" Then
                        VK = VK_Zoom
                    ElseIf cmdKey = "noname" Then
                        VK = VK_NoName
                    ElseIf cmdKey = "pa1" Then
                        VK = VK_Pa1
                    ElseIf cmdKey = "oemclear" Then
                        VK = VK_OemClear
                    ElseIf cmdKey = "lbutton, oemclear" Then
                        VK = VK_LButton_OemClear
                    End If
                End If
            ElseIf Letter = "~" Then '''''''''''''''''If letter is tilde
                VK = 13 ''''''''''''''''''''''''''''''Set to return(enter) key
            ElseIf Letter = "+" Then '''''''''''''''''If letter is plus
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
                Call KeyEvent(VK_ShiftKey, True, False) 'Press shift instead
            ElseIf Letter = "^" Then '''''''''''''''''If letter is caret
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
                Call KeyEvent(VK_ControlKey, True, False) 'Press control instead
            ElseIf Letter = "#" Then '''''''''''''''''If letter is number signifier
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
                Call KeyEvent(VK_LWin, True, False) ''Press left window key instead
            ElseIf Letter = "%" Then '''''''''''''''''If letter is percent
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
                Call KeyEvent(VK_Menu, True, False) ''Press menu key instead
            ElseIf Letter = "(" Then '''''''''''''''''If letter is left parenthesis
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
            ElseIf Letter = ")" Then '''''''''''''''''If letter is right parenthesis
                VK = 0 '''''''''''''''''''''''''''''''Do not send a regular key
                Call Lift ''''''''''''''''''''''''''''Lift extented keys
            End If
            If VK > NEGATIVE Then ''''''''''''''''''''If valid key code
                If asMessage = True Then '''''''''''''If sending a message
                    If (VK > 64 And VK < 91) Or (VK > 47 And VK < 58) Or (VK > 105 And VK < 112) Or (VK > 185 And VK < 193) Or (VK > 218 And VK < 224) Then IsLetter = True
                    Call Message(VK, WFOCUS, Not IsLetter, True, True) 'Send message down and up
                Else '''''''''''''''''''''''''''''''''If sending an event
                    If KeyEvent(VK) = False Then Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, rForeground, prvWnd): Exit Function 'If key event fails return abort
                End If
                kSent = kSent + 1 ''''''''''''''''''''Count the sent keys
            Else '''''''''''''''''''''''''''''''''''''It's a string to be repeated
                Dim n As Long
                For n = 1 To repeat  '''''''''''''''''Repeat key press
                    Dim r As String
                    r = cmdKey '''''''''''''''''''''''Initialize remaining text
                    Dim a As String
                    a = "" '''''''''''''''''''''''''''Initialize current letter
                    Dim w As Long
                    For w = 1 To Len(cmdKey) '''''''''Iterate the length of the string
                        a = Left(r, 1) '''''''''''''''Strip off left most character
                        r = Mid(r, 2) ''''''''''''''''Set remaining to right most characters
                        VK = apiVkKeyScan(Asc(a)) And 255 'Get scan code for this key
                        If asMessage = True Then '''''If sending a message
                            If (VK > 64 And VK < 91) Or (VK > 47 And VK < 58) Or (VK > 105 And VK < 112) Or (VK > 185 And VK < 193) Or (VK > 218 And VK < 224) Then IsLetter = True
                            Call Message(VK, WFOCUS, Not IsLetter, True, True) 'Send message down and up
                        Else '''''''''''''''''''''''''If sending an event
                            If KeyEvent(VK) = False Then Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, rForeground, prvWnd): Exit Function 'If key event fails return abort
                        End If
                         kSent = kSent + 1 '''''''''''Count the sent keys
                    Next w '''''''''''''''''''''''''''Next character in string
                    VK = NEGATIVE ''''''''''''''''''''Reset to non valid command for next loop
                Next n '''''''''''''''''''''''''''''''Next repeated
                repeat = 1 '''''''''''''''''''''''''''Reset repeat to one for next loop
            End If
            If Letter <> "(" Then ''''''''''''''''''''If character is not parenthesis
                If PrevLetter = "#" Then '''''''''''''If previous letter was numeric signifier
                    Call KeyEvent(VK_LWin, False, True) 'Lift win key
                ElseIf PrevLetter = "+" Then '''''''''If previous letter was shift
                    Call KeyEvent(VK_ShiftKey, False, True) 'Lift shift key
                ElseIf PrevLetter = "%" Then '''''''''If previous letter was percent
                    Call KeyEvent(VK_Menu, False, True) 'Lift menu key
                ElseIf PrevLetter = "^" Then '''''''''If previous letter was caret
                    Call KeyEvent(VK_ControlKey, False, True) 'Lift control key
                End If
            End If
            If Len(cmdKey) = 1 And cmdKey = "#" Or cmdKey = "+" Or cmdKey = "^" Or cmdKey = "%" Or cmdKey = "~" Or cmdKey = "(" Or cmdKey = ")" Or cmdKey = "[" Or cmdKey = "]" Then
                Call KeyEvent(VK_ShiftKey, False, True) 'Lift shift for command keys
            End If
            If Letter <> LowerCase Or Letter = "!" Or Letter = "@" Or Letter = "$" Or Letter = "&" Or Letter = "*" Or Letter = "_" Or Letter = "|" Or Letter = ":" Or Letter = "<" Or Letter = ">" Or Letter = "?" Then
                Call KeyEvent(VK_ShiftKey, False, True) 'Lift shift if necessary
            End If
            If asMessage = True Then apiSleep (0) ''''Wait for other threads
        Next i '''''''''''''''''''''''''''''''''''''''Next character
        Send = KeyAbort(WFOCUS, ws, UserCapsLock, sThread, rForeground, prvWnd)  'Final abort upon completion
    End Function

    ' Sends text directly to a window.  Use GetWinHandles or GetWinFocus to get a structure.
    ' Returns true if successful.
    ' sText The text to send.
    ' wFocus The structure of the window.  Use GetWinHandles or GetWinFocus to get a structure.
    Public Function Text(ByVal sText As String, ByRef WFOCUS As WINFOCUS) As Boolean
        On Error Resume Next
        Text = apiSendMessage(WFOCUS.Focus, WM_SETTEXT, 0, sText) 'Return the full result of SendMessage
    End Function

    Private Function AttachInput(ByVal hwnd As Long, Optional ByVal bAttach As Boolean = True) As Boolean
        On Error Resume Next
        AttachInput = CBool(apiAttachThreadInput(apiGetWindowThreadProcessId(hwnd, 0), apiGetCurrentThreadId, CInt(bAttach)))
    End Function

    Private Function Callback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        On Error Resume Next
        Static hStruct As KBDLLHOOKSTRUCT
            Call apiCopyMemory(hStruct, lParam, Len(hStruct))
            If fWnd <> 0 Then
                If apiGetForegroundWindow <> fWnd Then 'If foreground window is not where it's supposed to be
                    Sleep (1) ''''''''''''''''''''''''Sleep for one millisecond and flush messages
                    If apiSetForegroundWindow(fWnd) = 0 Then 'If foreground window cannot be set
                        If hStruct.dwExtraInfo = -11 Then kSent = kSent - 1 'Uncount this key, since it's blocked, and has been sent from the Send function
                        Callback = HC_GETNEXT ''''''''If foreground cannot be set then block key no matter if it's a user or internally sent from here
                        Exit Function
                    End If
                End If
            End If
            If hStruct.dwExtraInfo <> -11 Then '''''''If key press is not simulated from this module with -11 attached as extrainfomessage, then block this user key
                Callback = HC_GETNEXT ''''''''''''''''If key action and stroke blocked then get next key in the hook chain
                Exit Function
            End If
        Callback = apiCallNextKeyHookEx(hKey, Code, wParam, lParam) 'Call next key hook if no action
    End Function

    Private Function ForceForeground(ByVal hwnd As Long) As Boolean
        On Error Resume Next
        apiLockSetForegroundWindow (LSFW_UNLOCK) '''''Unlock setforegroundwindow calls
        apiAllowSetForegroundWindow (ASFW_ANY) '''''''Allow setforeground window calls
        Call KeyEvent(VK_Menu, False, True) ''''''''''Lift menu key if pressed, and it also allows the foreground window to be set
        ForceForeground = CBool(apiSetForegroundWindow(hwnd)) 'Set foreground window
        apiLockSetForegroundWindow (LSFW_LOCK) '''''''Lock other apps from using setforegroundwindow
    End Function

    Private Function GetChildWindow(ByVal hwnd As Long, Optional ByVal wName As String = "", Optional ByVal wIndex As Long = 1) As Long
        On Error Resume Next
        Dim cwnd As Long
        Dim TextCount As Long
        Dim ClassCount As Long
        Dim NullCount As Long
        Dim w As WINNAME
        Do '''''''''''''''''''''''''''''''''''''''''''Loop through sibling windows
            cwnd = apiFindWindowEx(hwnd, cwnd, vbNullString, vbNullString) 'Set child handle
            If cwnd = 0 Then GetChildWindow = cwnd: Exit Do  'If no more sibling children then return
            w = GetWinName(cwnd, True, True) '''''''''Get the text, and class from that window
            If w.lpText = wName Then TextCount = TextCount + 1 'If text matches the specified wName then increment the text count
            If TextCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If text count is equal to the specified wIndex then return
            If w.lpClass = wName Then ClassCount = ClassCount + 1 'If class name matches the specified wName then increment the class count
            If ClassCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If class count is equal to the specified wIndex then return
            If w.lpText = "" And w.lpClass = "" And wName = "" Then NullCount = NullCount + 1 'If text and class name not specified, then increment by index only
            If NullCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If null count is equal to wIndex then return
        Loop
    End Function

    Private Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String
        On Error Resume Next
        Dim lngCounter As Long
        strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
        lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn) ' Calculate the offset for the item required based on the number of columns the list strList' has i.e. 'lngNumColumns' and from which row the element is to be  selected i.e. 'lngRow'.
        For lngCounter = 0 To lngColumn - 1 ''''''''''Search for the 'lngColumn' item from the list 'strList'.
            strList = Mid(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))  ' Remove each item from the list.
            If Len(strList) = 0 Then GetElement = "": Exit Function 'If list becomes empty before 'lngColumn' is found then just return an empty string.
        Next lngCounter
        GetElement = Left(strList, InStr(strList, strDelimiter) - 1) 'Return the sought list element.
    End Function

    Private Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer
        On Error Resume Next
        Dim intElementCount As Integer
        If Len(strList) = 0 Then GetNumElements = 0: Exit Function 'If no elements in the list 'strList' then just return 0.
        strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
        While InStr(strList, strDelimiter) > 0 '''''''Count the number of elements in 'strlist'
            intElementCount = intElementCount + 1
            strList = Mid(strList, InStr(strList, strDelimiter) + 1, Len(strList))
        Wend
        GetNumElements = intElementCount '''''''''''''Return the number of elements in 'strList'.
    End Function

    Private Function GetWinAncestory(ByVal cwnd As Long, ByVal gName As Boolean) As Long()
        On Error Resume Next
        Dim i As Long
        Dim z As Long
        Dim cWnd2 As Long
        Dim hwnds() As Long
        Dim w As WINNAME
        cWnd2 = cwnd '''''''''''''''''''''''''''''''''Remember handle for seconds loop
        w = GetWinName(cwnd, True, True) '''''''''''''Get text and class name of the window
        If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Focus window:  Title | Class name") 'Display to developer if specified
        Do '''''''''''''''''''''''''''''''''''''''''''Loop while counting ancestors
            cwnd = apiGetParent(cwnd) ''''''''''''''''Get new parent handle if any found
            If cwnd = 0 Then Exit Do '''''''''''''''''If there are no more parents then abort without setting anymore to the array
            i = i + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
        Loop
        ReDim hwnds(i) '''''''''''''''''''''''''''''''Re-Dimension array to number of ancestors including the handle specified
        Do '''''''''''''''''''''''''''''''''''''''''''Do look for parents of the window specified.
            hwnds(z) = cWnd2 '''''''''''''''''''''''''Set the indexed array element to the specified handle or it's ancestors
            cWnd2 = apiGetParent(cWnd2) ''''''''''''''Get new parent handle if any found
            If cWnd2 = 0 Then Exit Do ''''''''''''''''If there are no more parents then abort
            z = z + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
            w = GetWinName(cWnd2, True, True) ''''''''Get text and class name of the window
            If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Parent window:  Title | Class name")  'Display to developer if specified
        Loop
        GetWinAncestory = hwnds ''''''''''''''''''''''Return array of integer handles, or handle.
    End Function

    Private Function GetWinName(ByVal hwnd As Long, Optional ByVal gText As Boolean = True, Optional ByVal gClass As Boolean = True) As WINNAME
        On Error Resume Next
        Dim tLength As Long
        Dim rValue As Long
        GetWinName.lpText = "" '''''''''''''''''''''''Initialize string for text name
        GetWinName.lpClass = "" ''''''''''''''''''''''Initialize string for class name
        If gText = True Then '''''''''''''''''''''''''If text is to be retrieved
            tLength = apiGetWindowTextLength(hwnd) + 4 'Get length
            GetWinName.lpText = Strings.Space(tLength) 'Pad with buffer
            rValue = apiGetWindowText(hwnd, GetWinName.lpText, tLength) 'Get text
            GetWinName.lpText = Left(GetWinName.lpText, rValue) 'Strip buffer
        End If
        If gClass = True Then ''''''''''''''''''''''''If class name is to be retrieved
            GetWinName.lpClass = Strings.Space(260) ''Pad with buffer
            rValue = apiGetClassName(hwnd, GetWinName.lpClass, 260) 'Get classname
            GetWinName.lpClass = Left(GetWinName.lpClass, rValue) 'Strip buffer
        End If
    End Function

    Private Function HandleToProcessId(ByVal hwnd As Long) As Long
        On Error Resume Next
        Dim wnd As Long
        Dim pId As Long
        If apiGetParent(hwnd) <> 0 Then hwnd = apiGetAncestor(hwnd, GA_ROOT)
        wnd = apiGetTopWindow(0) '''''''''''''''''''''Get the top window in the z-order
        Do
            If wnd = hwnd Then
                Call apiGetWindowThreadProcessId(wnd, pId)  'Get the window's process id
                HandleToProcessId = pId: Exit Do '''''Set process id and return
            End If
            wnd = apiGetWindow(wnd, GW_HWNDNEXT) '''''Retrieve the next window
            If wnd = 0 Then Exit Do
        Loop
    End Function

   Private Function KeyAbort(ByRef WFOCUS As WINFOCUS, ByRef WSTATE As WINSTATE, ByVal uCapsLock As Boolean, ByVal sThread As Boolean, ByVal rForeground As Boolean, ByVal prvWnd As Long) As Long
       On Error Resume Next
        Call Lift    '''''''''''''''''''''''''''''''''''''Lift any extended keys just in case
        If sThread = False Then Call AttachInput(WFOCUS.Foreground, False) 'If thread was attached then detatch it
        If WSTATE.IsIconic = True Then Call apiShowWindow(WFOCUS.Foreground, SW_SHOWMINIMIZED)  'If window was minimized, then re-minimize it
        If WSTATE.IsDisabled = True Then Call apiEnableWindow(WFOCUS.Foreground, False)  'If window was disabled then re-disable it
        If WSTATE.IsChildDisabled = True Then Call apiEnableWindow(WFOCUS.Focus, False)  'If child window was disabled then re-disable it
        If WSTATE.IsHidden = True Then Call apiShowWindow(WFOCUS.Foreground, SW_HIDE)  'If window was hidden, then re-hide it
        If WSTATE.IsChildHidden = True Then Call apiShowWindow(WFOCUS.Focus, SW_HIDE)  'If child window was hidden, then re-hide it
        Sleep (0)  '''''''''''''''''''''''''''''''''''''''Sleep for a moment and flush keys if needed
        If uCapsLock = True Then KeyEvent (VK_Capital) 'If caps lock was on toggle capslock
        If rForeground = True Then ForceForeground (prvWnd)
        If hKey <> 0 Then
            If apiUnhookWindowsHookEx(hKey) = 1 Then hKey = 0 'Unhook keyboard and free keyboard handle if unhooked  'If return foreground was specified then return it
        End If
        KeyAbort = kSent '''''''''''''''''''''''''''''Set return value
        fWnd = 0 '''''''''''''''''''''''''''''''''''''Free foregroundwindow handle
        kSent = 0 ''''''''''''''''''''''''''''''''''''Free number of keys sent
    End Function

    Private Function KeyEvent(Optional ByVal vKey As Long = 0, Optional ByVal kDown As Boolean = True, Optional ByVal kUp As Boolean = True) As Boolean
        On Error Resume Next
        If vKey < 0 Or vKey > 255 Then KeyEvent = False: Exit Function 'If vKey is not valid between 0-255
        KeyEvent = True
        If kDown = True Then '''''''''''''''''''''''''If key down specified
            If apikeybd_event(vKey, 0, 0, -11) = False Then KeyEvent = False 'press key down.  Set return if false
        End If
        If kUp = True Then '''''''''''''''''''''''''''If key up specified
            If apikeybd_event(vKey, 0, KEYEVENTF_KEYUP, -11) = False Then KeyEvent = False 'lift key up. Set return if false
        End If
    End Function

    Private Function Lift() As Boolean
        On Error Resume Next
        Call KeyEvent(VK_ControlKey, False, True)  '''Lift control
        Call KeyEvent(VK_LControlKey, False, True) '''Lift left control
        Call KeyEvent(VK_RControlKey, False, True) '''Lift right control
        Call KeyEvent(VK_ShiftKey, False, True) ''''''Lift shift
        Call KeyEvent(VK_LShiftKey, False, True)  ''''Lift left shift
        Call KeyEvent(VK_RShiftKey, False, True) '''''Lift right shift
        Call KeyEvent(VK_Menu, False, True) ''''''''''Lift menu
        Call KeyEvent(VK_LMenu, False, True) '''''''''Lift left menu
        Call KeyEvent(VK_RMenu, False, True)  ''''''''Lift right menu
        Lift = True
    End Function

    Private Function ProcessIdToHandle(ByVal pId As Long) As Long
        On Error Resume Next
        Dim hwnd As Long
        Dim processId As Long
        hwnd = apiGetTopWindow(0) ''''''''''''''''''''Get the top window in the z-order
        Do
            If apiGetParent(hwnd) = 0 Then '''''''''''If window has no parents
                Call apiGetWindowThreadProcessId(hwnd, processId)  'Get the window's process id
                If processId = pId Then ProcessIdToHandle = hwnd: Exit Do 'If pid matches then return
            End If
            hwnd = apiGetWindow(hwnd, GW_HWNDNEXT) '''Retrieve the next window
            If hwnd = 0 Then Exit Do '''''''''''''''''Exit if there are no more windows
        Loop
    End Function

    Private Function ProcessNameToId(ByVal processName As String, Optional ByVal pId As Long = 0) As Long
        On Error Resume Next
        Dim pLength As Long
        Dim cbSizeReturned As Long
        Dim numElements As Long
        Dim pIDs() As Long
        Dim cbSize As Long
        Dim cbSize2 As Long
        Dim Ret As Long
        Dim pSize As Long
        Dim hProcess As Long
        Dim pLoop As Long
        Dim pName As String
        Dim mName As String
        Dim prName As String
        Dim pModules(1 To 200) As Long
        processName = UCase(Trim(processName))
        pLength = Len(processName)
        cbSize = 8
        cbSizeReturned = 96
            Do While cbSize <= cbSizeReturned
                cbSize = cbSize * 2 ''''''''''''''''''Increment Size
                ReDim pIDs(cbSize / 4) As Long '''''''Allocate Memory for Array
                Ret = apiEnumProcesses(pIDs(1), cbSize, cbSizeReturned) 'Get Process ID's
            Loop
            numElements = cbSizeReturned / 4  ''''''''Count number of processes returned
            For pLoop = 1 To numElements '''''''''''''Loop thru each process
                hProcess = apiOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDs(pLoop))  'Get a handle to the Process and Open
                If hProcess <> 0 Then
                    Ret = apiEnumProcessModules(hProcess, pModules(1), 200, cbSize2) 'Get an array of the module handles for the specified process
                    If Ret <> 0 Then '''''''''''''''''If the Module Array is retrieved, Get the ModuleFileName
                        mName = Space(MAX_PATH) ''''''Buffer with spaces first to allocate memory for byte array
                        pSize = 500  '''''''''''''''''Must be set prior to calling API
                        Ret = apiGetModuleFileNameExA(hProcess, pModules(1), mName, pSize) 'Get Process Name
                        pName = Left(mName, Ret) '''''Remove trailing spaces
                        pName = UCase(Trim(pName)) '''Check for Matching Upper case result
                        prName = GetElement(Trim(Replace(pName, Chr(0), "")), "\", 0, 0, GetNumElements(Trim(Replace(pName, Chr(0), "")), "\") - 1)
                        If pId = 0 Then ''''''''''''''If getting process id
                            If prName = processName Then 'If process name matches specification
                               ProcessNameToId = pIDs(pLoop) 'Set pid return
                               Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
                            End If
                        Else '''''''''''''''''''''''''If getting process handle from id
                            If pId = pIDs(pLoop) Then 'If process id matches specification
                                ProcessNameToId = hProcess 'Set process handle return
                                Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
                            End If
                        End If
                    End If
                End If
                Ret = apiCloseHandle(hProcess) '''''''Close the handle to this process
            Next
    End Function

    Private Function WaitForWindowIdle(ByVal hwnd As Long) As Boolean
        On Error Resume Next
        Dim pId As Long
        pId = HandleToProcessId(hwnd) ''''''''''''''''Get process id from handle
        If pId <> 0 Then WaitForWindowIdle = CBool(ProcessNameToId("", pId)) 'If process id found then set return value of specified pId, which is a process handle
    End Function

    ' This function sends window messages directly to a window.
    ' If the asMessage parameter is false then, this function simulates mouse events directly to a window.
    ' Returns false if the specified window/s cannot be found.
    ' If the asMessage parameter is false then, this function returns false if a handle cannot be established or
    ' the top window could not be set, or a window rectangle could not be found, or the cursor could not be set,
    ' or the mouse event could not be simulated.
    ' "mButtons" The button to click.  Use the Buttons enumeration for this parameter.
    '  "wFocus" The structure of the window to focus on.
    ' Use GetWinHandles or GetWinFocus to get a structure.
    '  "mDown" (Optional) Press mouse button down only.
    '  "mUp" (Optional) Lift mouse button up only.
    '  "asMessage" (Optional) Send as window message, or as a mouse event.
    ' It's recommended that you use True in this parameter.
    ' If false then the x and y parameters may apply.
    '  "x" (Optional) The x coordinate only applies if asMessage is false, and mButtons is
    ' Buttons.Move or Buttons.MoveAbsolute.
    '  "y" (Optional) The y coordinate only applies if asMessage is false, and mButtons is
    ' Buttons.Move or Buttons.MoveAbsolute.
    Public Function Click(ByVal mButtons As Long, ByRef WFOCUS As WINFOCUS, Optional ByVal mDown As Boolean = True, Optional ByVal mUp As Boolean = True, Optional ByVal asMessage As Boolean = True, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1) As Boolean
        On Error Resume Next
        Dim hwnd As Long
        Dim cwnd As Long
        Dim zOrder  As Long
        Dim zOrderChild As Long
        Dim p As POINTAPI
        Dim ws As WINSTATE
        Click = True
        If WFOCUS.Foreground = -1 Then Click = False: Exit Function
        If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False)
        If asMessage = True Then  ''''''''''''''''''''If sending as message
            Dim WM_DOWN As Long
            Dim WM_UP As Long
            If mButtons = Buttons.Move Then ''''''''''If moving from current position
                Dim cP As POINTAPI
               Call apiGetCursorPos(cP)  '''''''''''''Get cursor position
                Call apiSetCursorPos(cP.X + X, cP.Y + Y) 'Add to coordinates, and set cursor there
                Click = True: Exit Function ''''''''''Return success
            ElseIf mButtons = Buttons.MoveAbsolute Then
                Call apiSetCursorPos(X, Y) '''''''''''Set to absolute coordinates
                Click = False: Exit Function '''''''''Return success
            ElseIf mButtons = Buttons.LeftClick Or mButtons = Buttons.LeftDoubleClick Then 'If button specified is left
                WM_DOWN = WM_LBUTTONDOWN: WM_UP = WM_LBUTTONUP 'Then set left messages
            ElseIf mButtons = Buttons.RightClick Or mButtons = Buttons.RightDoubleClick Then 'If button specified is right
                WM_DOWN = WM_RBUTTONDOWN: WM_UP = WM_RBUTTONUP 'Then set right messages
            ElseIf mButtons = Buttons.MiddleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If button specified is middle
                WM_DOWN = WM_MBUTTONDOWN: WM_UP = WM_MBUTTONUP ' Then set middle messages
            End If
            hwnd = WFOCUS.Foreground '''''''''''''''''Set main handle to something smaller
            cwnd = WFOCUS.Focus ''''''''''''''''''''''Set child handle(if any) to something smaller
            If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
            If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'if minimized then show normal
            If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enable
            If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True): Sleep (25) 'If child disabled then enable
            If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show
            If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL): Sleep (25) 'If child hidden then show
            zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
            zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
            If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
            If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
            If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'If button down is specified, then press down
            If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'If button up is specified, then lift up
            If mButtons = Buttons.LeftDoubleClick Or mButtons = Buttons.RightDoubleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If it's a double click
                If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'Down again
                If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'Up again
            End If
            Call MouseAbort(ws, False, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
        Else '''''''''''''''''''''''''''''''''''''''''Else it's an event to send
            Dim repeat As Long
            Dim rCursor As Boolean
            Dim pts As POINTAPI
            If mButtons = Buttons.Move Then  '''''''''Move cursor to click point
                Call MouseEvent(mButtons, X, Y) ''''''Move cursor to click point
                Exit Function '''''''''''''''''''''''''''''''''Exit thread
            ElseIf mButtons = Buttons.MoveAbsolute Then
                 pts = ToScreen(X, Y)   ''''''''''''''Convert to screen coordinates
                If pts.X <> 0 Then X = pts.X  ''''''''If x point found then set
                If pts.Y <> 0 Then Y = pts.Y  ''''''''If y point found then set
                Call MouseEvent(mButtons, X, Y) ''''''Move cursor to click point
                Exit Function ''''''''''''''''''''''''Exit thread
            End If
            If WFOCUS.Foreground = -1 Then Exit Function  'Exit if return from GetWinHandles is negative
            If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False) 'Get current focus
            hwnd = WFOCUS.Foreground  ''''''''''''''''Set main handle to something smaller
            cwnd = WFOCUS.Focus  '''''''''''''''''''''Set child handle(if any) to something smaller
            If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
            If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'if minimized then show normal
            If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enable
            If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True): Sleep (25) 'If child disabled then enable
            If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show
            If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL): Sleep (25) 'If child hidden then show
            zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
            zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
            If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
            If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
            If mButtons = 12 Or mButtons = 48 Or mButtons = 768 Then repeat = 1    'If double click
            If mDown = True And mUp = False Then '''''If mouse down
                If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
                     mButtons = MOUSEEVENTF_LEFTDOWN 'Set as left down
                ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
                     mButtons = MOUSEEVENTF_RIGHTDOWN 'Set as right down
                ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
                     mButtons = MOUSEEVENTF_MIDDLEDOWN 'Set as middle down
                End If
            ElseIf mDown = False And mUp = True Then 'If mouse up
                If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
                     mButtons = MOUSEEVENTF_LEFTUP '''Set as left up
                ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
                     mButtons = MOUSEEVENTF_RIGHTUP ''Set as right up
                ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
                     mButtons = MOUSEEVENTF_MIDDLEUP 'Set as middle up
                End If
            End If
            If mButtons <> Buttons.Wheel And mButtons <> Buttons.VirtualDesk Then   'If it's a click
                Dim r As RECT
                If apiGetWindowRect(cwnd, r) = False Then 'If no rectangle is found
                    Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'If rectangle not found, then exit failure
                    Exit Function ''''''''''''''''''''Exit
                End If
                pts.X = CInt(r.rLeft + ((r.rRight - r.rLeft) / 2)) 'Set to the center of the horizon
                pts.Y = CInt(r.rTop + ((r.rBottom - r.rTop) / 2)) 'Set to the center of the vertical
                pts = ToScreen(pts.X, pts.Y) '''''''''Convert to screen coordinates
                rCursor = True '''''''''''''''''''''''Cursor position changed, remember to return it later
                 X = 0 '''''''''''''''''''''''''''''''null for click
                 Y = 0 '''''''''''''''''''''''''''''''null for click
                Call apiGetCursorPos(p)  '''''''''''''Get the current cursor position, to be returned later
                If MouseEvent(Buttons.MoveAbsolute, pts.X, pts.Y) = False Then 'Move cursor to click point
                    Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if fails to move
                    Exit Function ''''''''''''''''''''Exit thread
                End If
            End If
            Dim i As Long
            For i = 1 To repeat + 1  '''''''''''''''''Loop the number of repeats
                If MouseEvent(mButtons, X, Y) = False Then  'Do mouse event repeated number of times
                    Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if failure
                    Exit Function ''''''''''''''''''''Exit thread
                End If
            Next
            Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
        End If
    End Function

    ' Clicks on the specified menu item, by sending it a command message.
    ' If asMessage is false then this function clicks on the specified menu item,
    ' by simulating an entire chain of events.
    ' Returns false if window or menu cannot be found.
    ' If asMessage is false then this function returns false if a window or menu or rectangle cannot be found,
    ' or an event fails to be simulated.
    '  "asMessage"Send as window message, or as a mouse event.
    ' It's recommended that you use True in this parameter.
    '  "wName"Title or class name of the main window.
    '  "wIndex"Index of the main window.
    '  "mnuName1"Main menu title.
    '  "mnuIndex1"Main menu index.
    '  "mnuName2"(Optional) Sub menu title.
    '  "mnuIndex2"(Optional) Sub menu index.
    '  "mnuName3"(Optional) Sub menu title.
    '  "mnuIndex3"(Optional) Sub menu index.
    '  "mnuName4"(Optional) Sub menu title.
    '  "mnuIndex4"(Optional) Sub menu index.
    '  "mnuName5"(Optional) Sub menu title.
    '  "mnuIndex5"(Optional) Sub menu index.
    '  "mnuName6"(Optional) Sub menu title.
    '  "mnuIndex6"(Optional) Sub menu index.
    '  "mnuName7"(Optional) Sub menu title.
    '  "mnuIndex7"(Optional) Sub menu index.
    '  "mnuName8"(Optional) Sub menu title.
    '  "mnuIndex8"(Optional) Sub menu index.
    '  "mnuName9"(Optional) Sub menu title.
    '  "mnuIndex9"(Optional) Sub menu index.
    '  "mnuName10"Sub menu title.
    '  "mnuIndex10"(Optional) Sub menu index.</param>
    Public Function ClickMenu(ByVal asMessage As Boolean, ByVal wName As String, ByVal wIndex As Long, ByVal mnuName1 As String, Optional ByVal mnuIndex1 As Long = 1, Optional ByVal mnuName2 As String = " ", Optional ByVal mnuIndex2 As Long = 1, Optional ByVal mnuName3 As String = " ", Optional ByVal mnuIndex3 As Long = 1, Optional ByVal mnuName4 As String = " ", Optional ByVal mnuIndex4 As Long = 1, Optional ByVal mnuName5 As String = " ", Optional ByVal mnuIndex5 As Long = 1, Optional ByVal mnuName6 As String = " ", Optional ByVal mnuIndex6 As Long = 1, Optional ByVal mnuName7 As String = " ", Optional ByVal mnuIndex7 As Long = 1, Optional ByVal mnuName8 As String = " ", Optional ByVal mnuIndex8 As Long = 1, Optional ByVal mnuName9 As String = " ", Optional ByVal mnuIndex9 As Long = 1, Optional ByVal mnuName10 As String = " ", Optional ByVal mnuIndex10 As Long = 1) As Boolean
        On Error Resume Next
        Dim n As Long
        Dim mnuNames() As String
        Dim i As Long
        Dim nPos As Long
        Dim ws As WINSTATE
        If asMessage = True Then
            Dim hwnd As Long
            Dim hMenu As Long
            Dim hSubMenu As Long
            Dim hId As Long
            hwnd = GetWinHandles(wName, wIndex).Foreground 'Set handle to specification
            If hwnd = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
            If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If minimized then show it
            If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enabled it
            If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If Hidden then show it
            hMenu = apiGetMenu(hwnd) '''''''''''''''''Set handle of main menu
            If hMenu = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
            hSubMenu = hMenu '''''''''''''''''''''''''Initialize sub menu to main menu
            If mnuName1 <> " " Then n = 1: ReDim mnuNames(1) 'Re-dimension mnuNames array
            If mnuName2 <> " " Then n = 3: ReDim mnuNames(3)
            If mnuName3 <> " " Then n = 5: ReDim mnuNames(5)
            If mnuName4 <> " " Then n = 7: ReDim mnuNames(7)
            If mnuName5 <> " " Then n = 9: ReDim mnuNames(9)
            If mnuName6 <> " " Then n = 11: ReDim mnuNames(11)
            If mnuName7 <> " " Then n = 13: ReDim mnuNames(13)
            If mnuName8 <> " " Then n = 15: ReDim mnuNames(15)
            If mnuName9 <> " " Then n = 17: ReDim mnuNames(17)
            If mnuName10 <> " " Then n = 19: ReDim mnuNames(19)
            If n > 0 Then mnuNames(0) = mnuName1: mnuNames(1) = CStr(mnuIndex1) 'Set elements of array to specification
            If n > 2 Then mnuNames(2) = mnuName2: mnuNames(3) = CStr(mnuIndex2)
            If n > 4 Then mnuNames(4) = mnuName3: mnuNames(5) = CStr(mnuIndex3)
            If n > 6 Then mnuNames(6) = mnuName4: mnuNames(7) = CStr(mnuIndex4)
            If n > 8 Then mnuNames(8) = mnuName5: mnuNames(9) = CStr(mnuIndex5)
            If n > 10 Then mnuNames(10) = mnuName6: mnuNames(11) = CStr(mnuIndex6)
            If n > 12 Then mnuNames(12) = mnuName7: mnuNames(13) = CStr(mnuIndex7)
            If n > 14 Then mnuNames(14) = mnuName8: mnuNames(15) = CStr(mnuIndex8)
            If n > 16 Then mnuNames(16) = mnuName9: mnuNames(17) = CStr(mnuIndex9)
            If n > 18 Then mnuNames(18) = mnuName10: mnuNames(19) = CStr(mnuIndex10)
            For i = 0 To (n - 1) Step 2  '''''''''''''Loop through menu tree
                If mnuNames(i) = "" Then '''''''''''''If menu name not specified then
                    nPos = CInt(mnuNames(i + 1)) - 1 'Set position as an index
                Else '''''''''''''''''''''''''''''''''Then name was set
                    nPos = FindMenuItemPos(hSubMenu, mnuNames(i), CInt(mnuNames(i + 1))) 'Set position to name and index of specified item.
                End If
                If nPos <> -1 Then
                    If apiGetSubMenu(hSubMenu, nPos) <> 0 Then hSubMenu = apiGetSubMenu(hSubMenu, nPos) 'If sub menu exits, and is specified then get the handle
                End If
            Next
            If nPos <> -1 Then '''''''''''''''''''''''If final item has a valid position
                hId = apiGetMenuItemID(hSubMenu, nPos) 'Get menu id
                If hId <> -1 Then ''''''''''''''''''''If item has no sub menus
                    ClickMenu = Not CBool(apiSendMessage(hwnd, WM_COMMAND, hId, vbNullString)) 'Send command message
                Else
                    ClickMenu = False ''''''''''''''''Return failure
                End If
            End If
            If ws.IsIconic = True Then Sleep (25): Call apiShowWindow(hwnd, SW_SHOWMINIMIZED) 'If window was minimized then re-minimize it
            If ws.IsDisabled = True Then Sleep (25): Call apiEnableWindow(hwnd, False) 'If window was disabled then re-disable it
            If ws.IsHidden = True Then Sleep (25): Call apiShowWindow(hwnd, SW_HIDE) 'If window was hidden then re-hide it
        Else
            Dim rOffset As Long
            Dim tOffset As Long
            Dim LeftMost As Long
            Dim TopMost As Long
            Dim ArrayLength As Long
            Dim m As MENUINFO
            Dim mi As ITEMINFO
            Dim p As POINTAPI
            Dim poi As POINTAPI
            Dim r As RECT
            ClickMenu = True '''''''''''''''''''''''''Signal true indicating that the thread was started
            If mnuName1 <> " " Then n = 3: ReDim mnuNames(3)  'Re-dimension mnuNames array
            If mnuName2 <> " " Then n = 5: ReDim mnuNames(5)
            If mnuName3 <> " " Then n = 7: ReDim mnuNames(7)
            If mnuName4 <> " " Then n = 9: ReDim mnuNames(9)
            If mnuName5 <> " " Then n = 11: ReDim mnuNames(11)
            If mnuName6 <> " " Then n = 13: ReDim mnuNames(13)
            If mnuName7 <> " " Then n = 15: ReDim mnuNames(15)
            If mnuName8 <> " " Then n = 17: ReDim mnuNames(17)
            If mnuName9 <> " " Then n = 19: ReDim mnuNames(19)
            If mnuName10 <> " " Then n = 21: ReDim mnuNames(21)
            If n > 0 Then mnuNames(0) = wName: mnuNames(1) = wIndex  'Set elements of array to specification
            If n > 2 Then mnuNames(2) = mnuName1: mnuNames(3) = CStr(mnuIndex1)
            If n > 4 Then mnuNames(4) = mnuName2: mnuNames(5) = CStr(mnuIndex2)
            If n > 6 Then mnuNames(6) = mnuName3: mnuNames(7) = CStr(mnuIndex3)
            If n > 8 Then mnuNames(8) = mnuName4: mnuNames(9) = CStr(mnuIndex4)
            If n > 10 Then mnuNames(10) = mnuName5: mnuNames(11) = CStr(mnuIndex5)
            If n > 12 Then mnuNames(12) = mnuName6: mnuNames(13) = CStr(mnuIndex6)
            If n > 14 Then mnuNames(14) = mnuName7: mnuNames(15) = CStr(mnuIndex7)
            If n > 16 Then mnuNames(16) = mnuName8: mnuNames(17) = CStr(mnuIndex8)
            If n > 18 Then mnuNames(18) = mnuName9: mnuNames(19) = CStr(mnuIndex9)
            If n > 20 Then mnuNames(20) = mnuName10: mnuNames(21) = CStr(mnuIndex10)
            ArrayLength = n ''''''''''''''''''''''''''Set length of array
            m.hwnd = GetWinHandles(wName).Foreground 'Get the handle of the specified window
            If m.hwnd = 0 Then ClickMenu = False: Exit Function 'If window not found, then exit failure
            If apiIsIconic(m.hwnd) = True Then ws.IsIconic = apiShowWindow(m.hwnd, SW_SHOWNORMAL): Sleep (25) 'If minimized then show it
            If apiIsWindowEnabled(m.hwnd) = False Then ws.IsDisabled = apiEnableWindow(m.hwnd, True): Sleep (25) 'If disabled then enable it
            If apiIsWindowVisible(m.hwnd) = False Then ws.IsHidden = Not apiShowWindow(m.hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show it
            Call apiGetWindowRect(m.hwnd, r) '''''''''Set confirmation of window rectangle
            If r.rBottom <> 0 Then
               Call apiMoveWindow(m.hwnd, 0, 0, r.rRight - r.rLeft, r.rBottom - r.rTop, True): Sleep (25) 'If rectangle found then move window with coordinates
            End If
            Call apiGetCursorPos(poi)  '''''''''''''''Get the current position of the user's cursor, so that it can be returned
            m.hMenu = apiGetMenu(m.hwnd) '''''''''''''Set handle of the main menu
            If m.hMenu = 0 Then ClickMenu = False: Exit Function '''''''''''''If no handle found then exit sub with failure
            m.hSubMenu = apiGetSubMenu(m.hMenu, 0) '''Set handle of the first sub menu if any
            mi = MenuItemDim(m.hwnd, m.hMenu, 0) '''''Get the dimensions of the menu item
            If mi.Top = -1 And mi.Bottom = -1 And mi.Left = -1 And mi.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon negative results
            LeftMost = mi.Left '''''''''''''''''''''''Initialize the left most coordinate
            If mnuNames(2) = "" Then '''''''''''''''''If no name specified
                nPos = CInt(mnuNames(3)) - 1 '''''''''Set position by index only
            Else '''''''''''''''''''''''''''''''''''''Otherwise set position by name and index
                nPos = FindMenuItemPos(m.hMenu, mnuNames(2), CInt(mnuNames(3))) 'Find position of the menu item
            End If
            If nPos = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
            mi = MenuItemDim(m.hwnd, m.hMenu, nPos) ''Get item  dimensions
            If mi.Top = -1 And mi.Bottom = -1 And mi.Left = -1 And mi.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
            p = ToScreen(mi.Center.X, mi.Center.Y) '''Convert point to screen coordinates
            If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.X, p.Y) = False Then ClickMenu = False: Exit Function 'Move mouse, and exit if failure
            rOffset = mi.Left - LeftMost '''''''''''''Initialize offset from the left
            TopMost = mi.Bottom ''''''''''''''''''''''Initialize offset from the top
            If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click mouse, and exit upon failure
            m.hSubMenu = apiGetSubMenu(m.hMenu, nPos) 'Set handle of submenu
            If m.hSubMenu <> 0 Then ''''''''''''''''''If handle found
            If mnuNames(4) = "" Then '''''''''''''''''If no name specified
                nPos = CInt(mnuNames(5)) - 1 '''''''''Set by index only
            Else '''''''''''''''''''''''''''''''''''''Name and index specified
                nPos = FindMenuItemPos(m.hSubMenu, mnuNames(4), CInt(mnuNames(5))) 'Find by name and index
            End If
            If nPos = -1 Then ClickMenu = False: Exit Function 'Exit if position is invalid
            mi = MenuItemDim(m.hwnd, m.hSubMenu, nPos) 'Get dimensinos
            If mi.Top = -1 And mi.Bottom = -1 And mi.Left = -1 And mi.Right = -1 Then ClickMenu = False: Exit Function 'Exit if fails
            p = ToScreen(rOffset + mi.Center.X, mi.Center.Y) 'Convert point
            If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.X, p.Y) = False Then ClickMenu = False: Exit Function 'Move and exit if failure
            For i = 6 To ArrayLength Step 2 ''''''''''Step through the array
                If MoveItemToItem(mnuNames(i), mnuNames(i + 1), TopMost, nPos, rOffset, tOffset, m, mi) = False Then ClickMenu = False: Exit Function
            Next
            If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click final menu item in the chain, and exit if failure
            End If
            If r.rBottom <> 0 Then
                Call apiMoveWindow(m.hwnd, r.rLeft, r.rTop, r.rRight - r.rLeft, r.rBottom - r.rTop, True) 'If there was a rectangle move the window back to where it was
            End If
            Call apiSetCursorPos(poi.X, poi.Y) '''''''Return the position of the cursor back to the user
            If ws.IsIconic = True Then Sleep (25): Call apiShowWindow(m.hwnd, SW_SHOWMINIMIZED) 'If was minimized then re-minimize
            If ws.IsDisabled = True Then Sleep (25): Call apiEnableWindow(m.hwnd, False) 'If was disabled then re-disable
            If ws.IsHidden = True Then Sleep (25): Call apiShowWindow(m.hwnd, SW_HIDE) 'If was Hidden then re-Hide
        End If
    End Function

    Private Function FindMenuItemPos(ByVal hMenu As Long, Optional ByVal iName As String = "", Optional ByVal iIndex As Long = 1) As Long
        On Error Resume Next
        Dim i As Long
        Dim itemCount As Long
        Dim indexCount As Long
        Dim retValue As Long
        Dim mnuCaption As String
        Dim woShortcut As String
        If apiIsMenu(hMenu) = False Then FindMenuItemPos = -1: Exit Function 'Return negative result if it's not a menu handle
        FindMenuItemPos = NEGATIVE '''''''''''''''''''Set a default return value
        itemCount = apiGetMenuItemCount(hMenu) '''''''Count the number of menu items
        For i = 0 To itemCount - 1 '''''''''''''''''''Loop through all menu items
            mnuCaption = "" ''''''''''''''''''''''''''Initialize
            mnuCaption = Space(1024) '''''''''''''''''Pad with a buffer
            retValue = apiGetMenuString(hMenu, i, mnuCaption, Len(mnuCaption), 1024) 'Get menu caption
            mnuCaption = Left(mnuCaption, retValue) ''mnuCaption.Substring(0, retValue) 'Strip off buffer
            woShortcut = "" ''''''''''''''''''''''''''Initialize
            If InStr(mnuCaption, "&") = True Then woShortcut = Replace(mnuCaption, "&", "") 'If the & character exists, then remove it, so the developer doesn't have to specify
            apiCharLower (iName) '''''''''''''''''''''Convert to lower case
            apiCharLower (woShortcut)
            apiCharLower (mnuCaption)
            If iName = woShortcut Or iName = mnuCaption Then 'if specified name matches menu name, as non-case sensitive
                FindMenuItemPos = i ''''''''''''''''''Set return value as that position
                indexCount = indexCount + 1 ''''''''''Increment index by one
                If indexCount = iIndex Then Exit For 'If index matches the specification then exit loop
            End If
        Next
    End Function

    Private Function GetSetZOrder(ByVal hwnd As Long, Optional ByVal sPosition As Long = NEGATIVE) As Long
        On Error Resume Next
        Dim z As Long
        Dim swnd As Long
        swnd = apiGetWindow(hwnd, GW_HWNDFIRST) ''''''Get top or topmost window in context
        Do
            If sPosition = NEGATIVE Then '''''''''''''If not setting the z-order
                If swnd = hwnd Then GetSetZOrder = (z + 1): Exit Function 'If handle specified matches sibling window, then return the position in the z-order
            Else '''''''''''''''''''''''''''''''''''''Then this function sets the z-order to the specified position
                If z = sPosition - 1 Then GetSetZOrder = apiSetWindowPos(hwnd, swnd, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE): Exit Function
            End If
            swnd = apiGetWindow(swnd, GW_HWNDNEXT) '''Get the next sibling window handle, in the loop
            If swnd = 0 Then Exit Do '''''''''''''''''If there are no more sibling windows, then exit loop with default return
            z = z + 1 ''''''''''''''''''''''''''''''''Increment i by one
        Loop
        GetSetZOrder = swnd
    End Function

    Private Function MenuItemDim(ByVal hwnd As Long, ByVal hMenu As Long, ByVal nPos As Long) As ITEMINFO
        On Error Resume Next
        Dim m As ITEMINFO
        Dim r As RECT
        If apiGetMenuItemRect(hwnd, hMenu, nPos, r) = 0 Then 'If rectangle not found then set negative returns
            r.rTop = NEGATIVE ''''''''''''''''''''''''Fail
            r.rBottom = NEGATIVE '''''''''''''''''''''Fail
            r.rLeft = NEGATIVE '''''''''''''''''''''''Fail
            r.rRight = NEGATIVE ''''''''''''''''''''''Fail
        Else '''''''''''''''''''''''''''''''''''''''''Or set dimensions of menu item
            m.Width = (r.rRight - r.rLeft) '''''''''''Set width
            m.Height = (r.rBottom - r.rTop) ''''''''''Set height
            m.Center.X = CInt(r.rLeft + (m.Width / 2)) 'Set center point x
            m.Center.Y = CInt(r.rTop + (m.Height / 2)) 'Set center point y
        End If
        m.Left = r.rLeft '''''''''''''''''''''''''''''Set left coordinate
        m.Right = r.rRight '''''''''''''''''''''''''''Set right coordinate
        m.Top = r.rTop '''''''''''''''''''''''''''''''Set top coordinate
        m.Bottom = r.rBottom '''''''''''''''''''''''''Set bottom coordinate
        MenuItemDim = m
    End Function

    Private Function MouseAbort(ByRef WSTATE As WINSTATE, ByVal rCursor As Boolean, ByVal hwnd As Long, ByVal cwnd As Long, ByVal zOrder As Long, ByVal zOrderChild As Long, ByRef p As POINTAPI) As Boolean
        On Error Resume Next
        If rCursor = True Then Call apiSetCursorPos(p.X, p.Y) 'If it was a click then return cursor to user position
        If WSTATE.IsIconic = True Then Sleep (25): Call apiShowWindow(hwnd, SW_SHOWMINIMIZED) 'If main window was minimized before, then re-minimize it
        If WSTATE.IsDisabled = True Then Sleep (25): Call apiEnableWindow(hwnd, False) 'If main window was disabled before, then re-disable it
        If WSTATE.IsChildDisabled = True Then Sleep (25): Call apiEnableWindow(cwnd, False) 'If child window was disabled before, then re-disable it
        If WSTATE.IsHidden = True Then Sleep (25): Call apiShowWindow(hwnd, SW_HIDE) 'If main window was hidden before, then re-hide it
        If WSTATE.IsChildHidden = True Then Sleep (25): Call apiShowWindow(cwnd, SW_HIDE) 'If main window was hidden before, then re-hide it
        Call apiSetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE): Sleep (25) 'Set the window position to not topmost window.  TODO Get topmost status first, so original state can be restored.
        If zOrderChild > 0 Then Sleep (25): Call GetSetZOrder(cwnd, zOrderChild) 'If a z-order for the child window was obtained, reset the z-order of the child window
        If zOrder > 0 Then Sleep (25): Call GetSetZOrder(hwnd, zOrder) 'If a z-order for the main window was obtained, reset the z-order of the main window
        MouseAbort = True ''''''''''''''''''''''''''''Return when finished
    End Function

    Private Function MouseEvent(Optional ByVal mEvents As Long = 0, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0) As Boolean
        On Error Resume Next
        MouseEvent = apimouse_event(mEvents, X, Y, 0, apiGetMessageExtraInfo) 'Return results
    End Function

    Private Function MoveItemToItem(ByVal wName As String, ByVal wName2 As String, ByVal tMost As Long, ByRef nPos As Long, ByRef rOffset As Long, ByRef tOffset As Long, ByRef m As MENUINFO, ByRef mi As ITEMINFO) As Boolean
        On Error Resume Next
        Dim p As POINTAPI
        tOffset = tOffset + mi.Top - tMost '''''''''''Keep offset from top most
        rOffset = rOffset + mi.Width '''''''''''''''''Keep offset from left most
        If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then MoveItemToItem = False: Exit Function 'click sub menu, and return false if failure
        m.hSubMenu = apiGetSubMenu(m.hSubMenu, nPos) 'Get handle of submenu
        If wName = "" Then '''''''''''''''''''''''''''If name not specified
            nPos = CInt(wName2) - 1 ''''''''''''''''''Then the search is by index
        Else '''''''''''''''''''''''''''''''''''''''''Then the search is by name
            nPos = FindMenuItemPos(m.hSubMenu, wName, CInt(wName2)) 'Get position from handle and name
        End If
        If nPos = -1 Then MoveItemToItem = False: Exit Function '''''''''''''''If return is negative then exit and return false
        mi = MenuItemDim(m.hwnd, m.hSubMenu, 0) ''''''Get menu item dimensions
        If mi.Top = -1 And mi.Bottom = -1 And mi.Left = -1 And mi.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit if there is a negative return
        p = ToScreen(rOffset + mi.Center.X, tOffset + mi.Center.Y) 'Covert to screen coordinates
        If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.X, p.Y) = False Then MoveItemToItem = False: Exit Function 'Move to new screen location
        mi = MenuItemDim(m.hwnd, m.hSubMenu, nPos) ''Get menu item dimensions
        If mi.Top = -1 And mi.Bottom = -1 And mi.Left = -1 And mi.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit upon negative result
        p = ToScreen(rOffset + mi.Center.X, tOffset + mi.Center.Y) 'Convert to screen
        If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.X, p.Y) = False Then MoveItemToItem = False: Exit Function 'Move to point
        MoveItemToItem = True
    End Function

    Private Function ToScreen(ByVal X As Long, ByVal Y As Long) As POINTAPI
        On Error Resume Next
        ToScreen.X = CInt(X * SM_FULLSCREEN / apiGetSystemMetrics(SM_CXSCREEN)) 'Set the return value for x.
        ToScreen.Y = CInt(Y * SM_FULLSCREEN / apiGetSystemMetrics(SM_CYSCREEN)) 'Set the return value for y.
    End Function


