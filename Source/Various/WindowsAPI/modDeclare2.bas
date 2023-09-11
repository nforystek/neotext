Attribute VB_Name = "modDeclare2"
Option Explicit
' ------------------------------------------------------------------------
'
'    WIN32API.TXT -- Win32 API Declarations for Visual Basic
'
'              Copyright (C) 1994-98 Microsoft Corporation
'
'  This file is required for the Visual Basic 6.0 version of the APILoader.
'  Older versions of this file will not work correctly with the version
'  6.0 APILoader.  This file is backwards compatible with previous releases
'  of the APILoader with the exception that Constants are no longer declared
'  as Global or Public in this file.
'
'  This file contains only the Const, Type,
'  and Public Declare statements for  Win32 APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Public Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long
Public Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Public Declare Function DestroyCaret Lib "user32" () As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
' Resource Loading Routines

Public Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function CreateCursor Lib "user32" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function CreateIcon Lib "user32" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LookupIconIdFromDirectory Lib "user32" (presbits As Byte, ByVal fIcon As Long) As Long
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
' Dialog Manager Routines

Public Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As Msg) As Long
Public Declare Function MapDialogRect Lib "user32" (ByVal hDlg As Long, lpRect As RECT) As Long
Public Declare Function DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Public Declare Function DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
Public Declare Function DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Public Declare Function DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long
Public Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TranslateMDISysAccel Lib "user32" (ByVal hWndClient As Long, lpMsg As Msg) As Long
Public Declare Function ArrangeIconicWindows Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function DdeSetQualityOfService Lib "user32" (ByVal hWndClient As Long, pqosNew As SECURITY_QUALITY_OF_SERVICE, pqosPrev As SECURITY_QUALITY_OF_SERVICE) As Long
Public Declare Function ImpersonateDdeClientWindow Lib "user32" (ByVal hWndClient As Long, ByVal hWndServer As Long) As Long
Public Declare Function PackDDElParam Lib "user32" (ByVal Msg As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Public Declare Function UnpackDDElParam Lib "user32" (ByVal Msg As Long, ByVal lParam As Long, puiLo As Long, puiHi As Long) As Long
Public Declare Function FreeDDElParam Lib "user32" (ByVal Msg As Long, ByVal lParam As Long) As Long
Public Declare Function ReuseDDElParam Lib "user32" (ByVal lParam As Long, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Public Declare Function DdeUninitialize Lib "user32" (ByVal idInst As Long) As Long
' conversation enumeration functions

Public Declare Function DdeConnectList Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, ByVal hConvList As Long, pCC As CONVCONTEXT) As Long
Public Declare Function DdeQueryNextServer Lib "user32" (ByVal hConvList As Long, ByVal hConvPrev As Long) As Long
Public Declare Function DdeDisconnectList Lib "user32" (ByVal hConvList As Long) As Long
' conversation control functions

Public Declare Function DdeConnect Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, pCC As CONVCONTEXT) As Long
Public Declare Function DdeDisconnect Lib "user32" (ByVal hConv As Long) As Long
Public Declare Function DdeReconnect Lib "user32" (ByVal hConv As Long) As Long
Public Declare Function DdeQueryConvInfo Lib "user32" (ByVal hConv As Long, ByVal idTransaction As Long, pConvInfo As CONVINFO) As Long
Public Declare Function DdeSetUserHandle Lib "user32" (ByVal hConv As Long, ByVal id As Long, ByVal hUser As Long) As Long
Public Declare Function DdeAbandonTransaction Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal idTransaction As Long) As Long
' app server interface functions

Public Declare Function DdePostAdvise Lib "user32" (ByVal idInst As Long, ByVal hszTopic As Long, ByVal hszItem As Long) As Long
Public Declare Function DdeEnableCallback Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal wCmd As Long) As Long
Public Declare Function DdeImpersonateClient Lib "user32" (ByVal hConv As Long) As Long
Public Declare Function DdeNameService Lib "user32" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As Long
' app client interface functions

Public Declare Function DdeClientTransaction Lib "user32" (pData As Byte, ByVal cbData As Long, ByVal hConv As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As Long
' data transfer functions

Public Declare Function DdeCreateDataHandle Lib "user32" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal afCmd As Long) As Long
Public Declare Function DdeAddData Lib "user32" Alias "DdeAddDataA" (ByVal hData As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As Long
Public Declare Function DdeGetData Lib "user32" Alias "DdeGetDataA" (ByVal hData As Long, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
Public Declare Function DdeAccessData Lib "user32" Alias "DdeAccessDataA" (ByVal hData As Long, pcbDataSize As Long) As Long
Public Declare Function DdeUnaccessData Lib "user32" Alias "DdeUnaccessDataA" (ByVal hData As Long) As Long
Public Declare Function DdeFreeDataHandle Lib "user32" (ByVal hData As Long) As Long
Public Declare Function DdeGetLastError Lib "user32" (ByVal idInst As Long) As Long
Public Declare Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As Long
Public Declare Function DdeQueryString Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Long, ByVal hsz As Long, ByVal psz As String, ByVal cchMax As Long, ByVal iCodePage As Long) As Long
Public Declare Function DdeFreeStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Public Declare Function DdeKeepStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Public Declare Function DdeCmpStringHandles Lib "user32" (ByVal hsz1 As Long, ByVal hsz2 As Long) As Long
Public Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Public Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Public Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Public Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Public Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Public Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Public Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Public Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Public Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Public Declare Function midiStreamOpen Lib "winmm.dll" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Public Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long
Public Declare Function midiStreamProperty Lib "winmm.dll" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
Public Declare Function midiStreamPosition Lib "winmm.dll" (ByVal hms As Long, lpmmt As MMTIME, ByVal cbmmt As Long) As Long
Public Declare Function midiStreamOut Lib "winmm.dll" (ByVal hms As Long, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Public Declare Function midiStreamPause Lib "winmm.dll" (ByVal hms As Long) As Long
Public Declare Function midiStreamRestart Lib "winmm.dll" (ByVal hms As Long) As Long
Public Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long
Public Declare Function midiConnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Public Declare Function midiDisconnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Public Declare Function CloseDriver Lib "winmm.dll" (ByVal hDriver As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Public Declare Function OpenDriver Lib "winmm.dll" (ByVal szDriverName As String, ByVal szSectionName As String, ByVal lParam2 As Long) As Long
Public Declare Function SendDriverMessage Lib "winmm.dll" (ByVal hDriver As Long, ByVal message As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Public Declare Function DrvGetModuleHandle Lib "winmm.dll" (ByVal hDriver As Long) As Long
Public Declare Function GetDriverModuleHandle Lib "winmm.dll" (ByVal hDriver As Long) As Long
Public Declare Function DefDriverProc Lib "winmm.dll" (ByVal dwDriverIdentifier As Long, ByVal hdrvr As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Public Declare Function mmsystemGetVersion Lib "winmm.dll" () As Long
Public Declare Sub OutputDebugStr Lib "winmm.dll" (ByVal lpszOutputString As String)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutBreakLoop Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveOutGetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwPitch As Long) As Long
Public Declare Function waveOutSetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
Public Declare Function waveOutGetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwRate As Long) As Long
Public Declare Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
Public Declare Function waveOutGetID Lib "winmm.dll" (ByVal hWaveOut As Long, lpuDeviceID As Long) As Long
Public Declare Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Public Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInGetPosition Lib "winmm.dll" (ByVal hWaveIn As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveInGetID Lib "winmm.dll" (ByVal hWaveIn As Long, lpuDeviceID As Long) As Long
Public Declare Function waveInMessage Lib "winmm.dll" (ByVal hWaveIn As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Public Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutCachePatches Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
Public Declare Function midiOutCacheDrumPatches Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
Public Declare Function midiOutGetID Lib "winmm.dll" (ByVal hMidiOut As Long, lpuDeviceID As Long) As Long
Public Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Public Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Public Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Public Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Public Declare Function midiInGetID Lib "winmm.dll" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long
Public Declare Function midiInMessage Lib "winmm.dll" (ByVal hMidiIn As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Declare Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDev" () As Long
Public Declare Function joyGetThreshold Lib "winmm.dll" (ByVal id As Long, lpuThreshold As Long) As Long
Public Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Public Declare Function joySetCapture Lib "winmm.dll" (ByVal hWnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Public Declare Function joySetThreshold Lib "winmm.dll" (ByVal id As Long, ByVal uThreshold As Long) As Long
Public Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Public Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As MMIOINFO, ByVal dwOpenFlags As Long) As Long
Public Declare Function mmioRename Lib "winmm.dll" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long
Public Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Public Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Public Declare Function mmioWrite Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Public Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Public Declare Function mmioGetInfo Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioSetInfo Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioSetBuffer Lib "winmm.dll" (ByVal hmmio As Long, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Public Declare Function mmioFlush Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Public Declare Function mmioAdvance Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioSendMessage Lib "winmm.dll" (ByVal hmmio As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Public Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
' MCI functions

Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetCreatorTask Lib "winmm.dll" (ByVal wDeviceID As Long) As Long
Public Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Public Declare Function mciGetDeviceIDFromElementID Lib "winmm.dll" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function EnumPrinterPropertySheets Lib "winspool.drv" (hPrinter As Long, hWnd As Long, lpfnAdd As Long, ByVal lParam As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal Command As Long) As Long
Public Declare Function GetJob Lib "winspool.drv" Alias "GetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Public Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Public Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Public Declare Function GetPrinterDriver Lib "winspool.drv" Alias "GetPrinterDriverA" (ByVal hPrinter As Long, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pDriverName As String) As Long
Public Declare Function AddPrintProcessor Lib "winspool.drv" Alias "AddPrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPathName As String, ByVal pPrintProcessorName As String) As Long
Public Declare Function EnumPrintProcessors Lib "winspool.drv" Alias "EnumPrintProcessorsA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pPrintProcessorInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function GetPrintProcessorDirectory Lib "winspool.drv" Alias "GetPrintProcessorDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pPrintProcessorInfo As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function EnumPrintProcessorDatatypes Lib "winspool.drv" Alias "EnumPrintProcessorDatatypesA" (ByVal pName As String, ByVal pPrintProcessorName As String, ByVal Level As Long, pDatatypes As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Public Declare Function DeletePrintProcessor Lib "winspool.drv" Alias "DeletePrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProcessorName As String) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As Byte) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AbortPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function ReadPrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pNoBytesRead As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AddJob Lib "winspool.drv" Alias "AddJobA" (ByVal hPrinter As Long, ByVal Level As Long, pData As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function ScheduleJob Lib "winspool.drv" (ByVal hPrinter As Long, ByVal JobId As Long) As Long
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, ByVal fMode As Long) As Long
Public Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE) As Long
Public Declare Function GetPrinterData Lib "winspool.drv" Alias "GetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, pType As Long, pData As Byte, ByVal nSize As Long, pcbNeeded As Long) As Long
Public Declare Function SetPrinterData Lib "winspool.drv" Alias "SetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, ByVal dwType As Long, pData As Byte, ByVal cbData As Long) As Long
Public Declare Function WaitForPrinterChange Lib "winspool.drv" (ByVal hPrinter As Long, ByVal Flags As Long) As Long
Public Declare Function PrinterMessageBox Lib "winspool.drv" Alias "PrinterMessageBoxA" (ByVal hPrinter As Long, ByVal error As Long, ByVal hWnd As Long, ByVal pText As String, ByVal pCaption As String, ByVal dwType As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function AddMonitor Lib "winspool.drv" Alias "AddMonitorA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte) As Long
Public Declare Function DeleteMonitor Lib "winspool.drv" Alias "DeleteMonitorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pMonitorName As String) As Long
Public Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pMonitorName As String) As Long
Public Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pPortName As String) As Long
Public Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pPortName As String) As Long
Public Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Public Declare Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long
Public Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hWnd As Long, ByVal Flags As Long) As Long
Public Declare Function AddPrintProvidor Lib "winspool.drv" Alias "AddPrintProvidorA" (ByVal pName As String, ByVal Level As Long, pProvidorInfo As Byte) As Long
Public Declare Function DeletePrintProvidor Lib "winspool.drv" Alias "DeletePrintProvidorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProvidorName As String) As Long
Public Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As NETRESOURCE, lphEnum As Long) As Long
Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
'
' Win32 NetAPIs.
'

Public Declare Function NetUserChangePassword Lib "netapi32.dll" (Domain As Any, User As Any, OldPass As Byte, NewPass As Byte) As Long
Public Declare Function NetUserGetInfo Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Public Declare Function NetUserGetGroups Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Public Declare Function NetUserGetLocalGroups Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, ByVal Flags As Long, lpBuffer As Long, ByVal MaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
'Public Declare Function NetUserAdd Lib "netapi32" (lpServer As Any, ByVal Level As Long, lpUser As USER_INFO_3_API, lpError As Long) As Long
Public Declare Function NetWkstaGetInfo Lib "netapi32.dll" (lpServer As Any, ByVal Level As Long, lpBuffer As Any) As Long
Public Declare Function NetWkstaUserGetInfo Lib "netapi32.dll" (ByVal Reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
Public Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal lpBuffer As Long) As Long
Public Declare Function NetRemoteTOD Lib "netapi32.dll" (yServer As Any, pBuffer As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidW" (ByVal lpSystemName As Any, Sid As Any, Name As Any, cbName As Long, ReferencedDomainName As Any, cbReferencedDomainName As Long, peUse As Integer) As Long
Public Declare Function NetLocalGroupDelMembers Lib "netapi32.dll" (ByVal psServer As Long, ByVal psLocalGroup As Long, ByVal lLevel As Long, uMember As LOCALGROUP_MEMBERS_INFO_0, ByVal lMemberCount As Long) As Long
Public Declare Function NetLocalGroupGetMembers Lib "netapi32.dll" (ByVal psServer As Long, ByVal psLocalGroup As Long, ByVal lLevel As Long, pBuffer As Long, ByVal lMaxLength As Long, plEntriesRead As Long, plTotalEntries As Long, phResume As Long) As Long
'Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
'Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
'Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
'Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
' Registry API prototypes

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you Public Declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you Public Declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Public Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
'  Prototype for the Service Control Handler Function
' /////////////////////////////////////////////////////////////////////////
'  API Function Prototypes
' /////////////////////////////////////////////////////////////////////////

Public Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Long
Public Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Public Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Public Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Public Declare Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
Public Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Public Declare Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As Long, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
Public Declare Function GetServiceDisplayName Lib "advapi32.dll" Alias "GetServiceDisplayNameA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
Public Declare Function LockServiceDatabase Lib "advapi32.dll" (ByVal hSCManager As Long) As Long
Public Declare Function NotifyBootConfigStatus Lib "advapi32.dll" (ByVal BootAcceptable As Long) As Long
Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Public Declare Function QueryServiceLockStatus Lib "advapi32.dll" Alias "QueryServiceLockStatusA" (ByVal hSCManager As Long, lpLockStatus As QUERY_SERVICE_LOCK_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Public Declare Function QueryServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Public Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long
Public Declare Function SetServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
Public Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
Public Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Public Declare Function UnlockServiceDatabase Lib "advapi32.dll" (ScLock As Any) As Long
Public Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Public Declare Function LZInit Lib "lz32.dll" (ByVal hfSrc As Long) As Long
Public Declare Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Public Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Public Declare Function LZSeek Lib "lz32.dll" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Public Declare Function LZRead Lib "lz32.dll" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Public Declare Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)
'  prototype of IMM API

Public Declare Function ImmInstallIME Lib "imm32.dll" Alias "ImmInstallIMEA" (ByVal lpszIMEFileName As String, ByVal lpszLayoutText As String) As Long
Public Declare Function ImmGetDefaultIMEWnd Lib "imm32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmGetIMEFileName Lib "imm32.dll" Alias "ImmGetIMEFileNameA" (ByVal hkl As Long, ByVal lpStr As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmGetProperty Lib "imm32.dll" (ByVal hkl As Long, ByVal dw As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function ImmSimulateHotKey Lib "imm32.dll" (ByVal hWnd As Long, ByVal dw As Long) As Long
Public Declare Function ImmCreateContext Lib "imm32.dll" () As Long
Public Declare Function ImmDestroyContext Lib "imm32.dll" (ByVal himc As Long) As Long
Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As Long, ByVal himc As Long) As Long
Public Declare Function ImmAssociateContext Lib "imm32.dll" (ByVal hWnd As Long, ByVal himc As Long) As Long
Public Declare Function ImmGetCompositionString Lib "imm32.dll" Alias "ImmGetCompositionStringA" (ByVal himc As Long, ByVal dw As Long, lpv As Any, ByVal dw2 As Long) As Long
Public Declare Function ImmSetCompositionString Lib "imm32.dll" Alias "ImmSetCompositionStringA" (ByVal himc As Long, ByVal dwIndex As Long, lpComp As Any, ByVal dw As Long, lpRead As Any, ByVal dw2 As Long) As Long
Public Declare Function ImmGetCandidateListCount Lib "imm32.dll" Alias "ImmGetCandidateListCountA" (ByVal himc As Long, lpdwListCount As Long) As Long
Public Declare Function ImmGetCandidateList Lib "imm32.dll" Alias "ImmGetCandidateListA" (ByVal himc As Long, ByVal deIndex As Long, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long) As Long
Public Declare Function ImmGetGuideLine Lib "imm32.dll" Alias " ImmGetGuideLineA" (ByVal himc As Long, ByVal dwIndex As Long, ByVal lpStr As String, ByVal dwBufLen As Long) As Long
Public Declare Function ImmGetConversionStatus Lib "imm32.dll" (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long
Public Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function ImmGetOpenStatus Lib "imm32.dll" (ByVal himc As Long) As Long
Public Declare Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As Long, ByVal b As Long) As Long
Public Declare Function ImmGetCompositionFont Lib "imm32.dll" Alias "ImmGetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Public Declare Function ImmSetCompositionFont Lib "imm32.dll" Alias "ImmSetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Public Declare Function ImmConfigureIME Lib "imm32.dll" (ByVal hkl As Long, ByVal hWnd As Long, ByVal dw As Long) As Long
Public Declare Function ImmEscape Lib "imm32.dll" Alias "ImmEscapeA" (ByVal hkl As Long, ByVal himc As Long, ByVal un As Long, lpv As Any) As Long
Public Declare Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListA" (ByVal hkl As Long, ByVal himc As Long, ByVal lpsz As String, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Public Declare Function ImmNotifyIME Lib "imm32.dll" (ByVal himc As Long, ByVal dwAction As Long, ByVal dwIndex As Long, ByVal dwValue As Long) As Long
Public Declare Function ImmGetStatusWindowPos Lib "imm32.dll" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Public Declare Function ImmSetStatusWindowPos Lib "imm32.dll" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Public Declare Function ImmGetCompositionWindow Lib "imm32.dll" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Public Declare Function ImmSetCompositionWindow Lib "imm32.dll" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Public Declare Function ImmGetCandidateWindow Lib "imm32.dll" (ByVal himc As Long, ByVal dw As Long, lpCandidateForm As CANDIDATEFORM) As Long
Public Declare Function ImmSetCandidateWindow Lib "imm32.dll" (ByVal himc As Long, lpCandidateForm As CANDIDATEFORM) As Long
Public Declare Function ImmIsUIMessage Lib "imm32.dll" Alias "ImmIsUIMessageA" (ByVal hWnd As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ImmGetVirtualKey Lib "imm32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ImmRegisterWord Lib "imm32.dll" Alias "ImmRegisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String) As Long
Public Declare Function ImmUnregisterWord Lib "imm32.dll" Alias "ImmUnregisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszUnregister As String) As Long
Public Declare Function ImmGetRegisterWordStyle Lib "imm32.dll" Alias " ImmGetRegisterWordStyleA" (ByVal hkl As Long, ByVal nItem As Long, lpStyleBuf As STYLEBUF) As Long
Public Declare Function ImmEnumRegisterWord Lib "imm32.dll" Alias "ImmEnumRegisterWordA" (ByVal hkl As Long, ByVal RegisterWordEnumProc As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String, lpv As Any) As Long
' *****************************************************************************                                                                             *
' * shellapi.h -  SHELL.DLL functions, types, and definitions                   *
' *                                                                             *
' * Copyright (c) 1992-1995, Microsoft Corp.  All rights reserved               *
' *                                                                             *
' \*****************************************************************************/

Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As Long, ByVal hIcon As Long) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
' //  EndAppBar

Public Declare Function DoEnvironmentSubst Lib "shell32.dll" Alias "DoEnvironmentSubstA" (ByVal szString As String, ByVal cbString As Long) As Long
Public Declare Function FindEnvironmentString Lib "shell32.dll" Alias "FindEnvironmentStringA" (ByVal szEnvVar As String) As String
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Sub SHFreeNameMappings Lib "shell32.dll" (ByVal hNameMappings As Long)
Public Declare Sub WinExecError Lib "shell32.dll" Alias "WinExecErrorA" (ByVal hWnd As Long, ByVal error As Long, ByVal lpstrFileName As String, ByVal lpstrTitle As String)
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function SHGetNewLinkInfo Lib "shell32.dll" Alias "SHGetNewLinkInfoA" (ByVal pszLinkto As String, ByVal pszDir As String, ByVal pszName As String, pfMustCopy As Long, ByVal uFlags As Long) As Long
'  ----- Function prototypes -----

Public Declare Function VerFindFile Lib "version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Public Declare Function VerInstallFile Lib "version.dll" Alias " VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long
'  Returns size of version info in Bytes

Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'  Read version info into buffer
' /* Length of buffer for info *
' /* Information from GetFileVersionSize *
' /* Filename of version stamped file *

Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long
'  Define API decoration for direct importing of DLL references.

Public Declare Function HeapValidate Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapCompact Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Public Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long
Public Declare Function LogonUser Lib "kernel32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Public Declare Function ImpersonateLoggedOnUser Lib "kernel32" (ByVal hToken As Long) As Long
Public Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Public Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA " (pFindreplace As FINDREPLACE) As Long
Public Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (pFindreplace As FINDREPLACE) As Long
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Public Declare Function DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As Long, ByVal afCmd As Long, ByVal ulRes As Long) As Integer
Public Declare Function SetServiceBits Lib "advapi32" (ByVal hServiceStatus As Long, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Boolean, ByVal bUpdateImmediately As Boolean) As Long
Public Declare Function CopyLZFile Lib "lz32" (ByVal n1 As Long, ByVal n2 As Long) As Long
Public Declare Function LZStart Lib "lz32" () As Long
Public Declare Sub LZDone Lib "lz32" ()
Public Declare Function mciGetYieldProc Lib "winmm" (ByVal mciId As Long, pdwYieldData As Long) As Long
Public Declare Function mciSetYieldProc Lib "winmm" (ByVal mciId As Long, ByVal fpYieldProc As Long, ByVal dwYieldData As Long) As Long
Public Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
'Public Declare Function mmioInstallIOProcA Lib "winmm" Alias "mmioInstallIOProcA" (ByVal fccIOProc As String * 4, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Public Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As String, pNumArgs As Integer) As Long
Public Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Public Declare Function NotifyChangeEventLog Lib "advapi32" (ByVal hEventLog As Long, ByVal hEvent As Long) As Long
Public Declare Function SetThreadToken Lib "advapi32" (Thread As Long, ByVal Token As Long) As Long
Public Declare Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hWnd As Long, lpCC As COMMCONFIG) As Long
Public Declare Function CreateIoCompletionPort Lib "kernel32" (ByVal FileHandle As Long, ByVal ExistingCompletionPort As Long, ByVal CompletionKey As Long, ByVal NumberOfConcurrentThreads As Long) As Long
Public Declare Function DisableThreadLibraryCalls Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long
Public Declare Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
Public Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function GetCommConfig Lib "kernel32" (ByVal hCommDev As Long, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Public Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Public Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Public Declare Function GetHandleInformation Lib "kernel32" (ByVal hObject As Long, lpdwFlags As Long) As Long
Public Declare Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
Public Declare Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Public Declare Function GetQueuedCompletionStatus Lib "kernel32" (ByVal CompletionPort As Long, lpNumberOfBytesTransferred As Long, lpCompletionKey As Long, lpOverlapped As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long) As Long
Public Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Public Declare Sub GlobalFix Lib "kernel32" (ByVal hMem As Long)
Public Declare Sub GlobalUnfix Lib "kernel32" (ByVal hMem As Long)
Public Declare Function GlobalWire Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnWire Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Public Declare Function LocalCompact Lib "kernel32" (ByVal uMinFree As Long) As Long
Public Declare Function LocalShrink Lib "kernel32" (ByVal hMem As Long, ByVal cbNewSize As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Function ReadFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Public Declare Function SetCommConfig Lib "kernel32" (ByVal hCommDev As Long, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Public Declare Function SetDefaultCommConfig Lib "kernel32" Alias "SetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Public Declare Sub SetFileApisToANSI Lib "kernel32" ()
Public Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SetSystemTimeAdjustment Lib "kernel32" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Long
Public Declare Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SystemTime, lpLocalTime As SystemTime) As Long
Public Declare Function WriteFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Public Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumObjects Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal lpGOBJEnumProc As Long, lpVoid As Any) As Long
Public Declare Function FixBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LineDDA Lib "gdi32" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As Long, ByVal lParam As Long) As Long
Public Declare Function SETABORTPROC Lib "gdi32" Alias "SetAbortProc" (ByVal hdc As Long, ByVal lpAbortProc As Long) As Long
Public Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Public Declare Function GetCurrencyFormat Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long
Public Declare Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr As String, ByVal cchNumber As Long) As Long
Public Declare Function GetStringTypeEx Lib "kernel32" Alias "GetStringTypeExA" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Public Declare Function GetStringTypeW Lib "kernel32" (ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function FindClosePrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long) As Long
Public Declare Function FindFirstPrinterChangeNotification Lib "winspool.drv" (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal pPrinterNotifyOptions As String) As Long
Public Declare Function FindNextPrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long, pdwChange As Long, ByVal pvReserved As String, ByVal ppPrinterNotifyInfo As Long) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
Public Declare Function BroadcastSystemMessage Lib "user32" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CascadeWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, ByVal lpRect As RECT, ByVal cKids As Long, lpKids As Long) As Integer
Public Declare Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal Flags As Long) As Long
Public Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
'Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwndParent As Long, ByVal pt As POINTAPI) As Long
Public Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal pt As POINTAPI, ByVal un As Long) As Long
Public Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Public Declare Function CloseWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Public Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As String, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Public Declare Function CreateIconFromResource Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Public Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Public Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal pt As POINTAPI) As Long
Public Declare Function DragObject Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal un As Long, ByVal dw As Long, ByVal hCursor As Long) As Long
Public Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinSta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long
Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Public Declare Function GetWindowContextHelpId Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Public Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LookupIconIdFromDirectoryEx Lib "user32" (presbits As Byte, ByVal fIcon As Boolean, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Public Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Public Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Public Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Public Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Public Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMessageExtraInfo Lib "user32" (ByVal lParam As Long) As Long
Public Declare Function SetMessageQueue Lib "user32" (ByVal cMessagesMax As Long) As Long
Public Declare Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Public Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Long
Public Declare Function SetWindowContextHelpId Lib "user32" (ByVal hWnd As Long, ByVal dw As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Public Declare Function TileWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As Long) As Integer
Public Declare Function ToAsciiEx Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
Public Declare Function UnhookWindowsHook Lib "user32" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As Long) As Integer
Public Declare Function WNetGetUniversalName Lib "mpr" Alias "WNetGetUniversalNameA" (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long

