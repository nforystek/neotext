Attribute VB_Name = "modDoTasks"
#Const modDoTasks = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const PM_REMOVE = &H1
Private Const PM_NOREMOVE = &H0
Private Const PM_NOYIELD = &H2

Private Const HWND_ALL = 0
Private Const HWND_APP = -1

Private Const DO_STACK = 1
Private Const DO_EVENT = 2
Private Const DO_CHILD = 4
Private Const DO_OTHER = 8

Private Const MSG_LEVEL = 1
Private Const MSG_TIER2 = 2
Private Const MSG_EMBED = 4

Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'Private Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Private Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
'Private Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Private Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long) As Long
'Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private pStacks As Long

Public Sub DoTasks()
    Static dMsg As Msg
    Do While PeekMessage(dMsg, -1, 0, 0, PM_REMOVE + PM_NOYIELD)
        TranslateMessage dMsg
        DispatchMessage dMsg
    Loop
    If PeekMessage(dMsg, 0, 0, 0, PM_NOREMOVE + PM_NOYIELD) Then
        pStacks = pStacks + 1
        If pStacks = 1 Then
            EnumWindows AddressOf WinEvents, GetCurrentProcessId
            Do While PeekMessage(dMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD)
                TranslateMessage dMsg
                DispatchMessage dMsg
                EnumWindows AddressOf WinEvents, GetCurrentProcessId
            Loop
            EnumWindows AddressOf WinEvents, -4
        End If
        pStacks = pStacks - 1
        EnumWindows AddressOf WinEvents, -2
    ElseIf pStacks = 1 Then
        EnumWindows AddressOf WinEvents, -1
    Else
        EnumWindows AddressOf WinEvents, -3
    End If
End Sub

Private Function WinEvents(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim pID As Long
    Static wMsg As Msg
    If (lParam <= 0) And (lParam >= -3) Then
        If PeekMessage(wMsg, hwnd, 0, 0, PM_REMOVE + PM_NOYIELD) Then
            Do
                TranslateMessage wMsg
                DispatchMessage wMsg
            Loop While PeekMessage(wMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD)
        ElseIf (lParam >= -1) Then
            Sleep 1
        End If
        If (lParam = -3) Then Sleep 1
        If pStacks = 1 Then DoEvents
    ElseIf (lParam = -4) Then
        If (pStacks > 1) And (pStacks < 3) Then
            If PeekMessage(wMsg, 0, 0, 0, PM_NOREMOVE + PM_NOYIELD) Then
                Do
                    Sleep 0
                Loop While PeekMessage(wMsg, hwnd, 0, 0, PM_NOREMOVE + PM_NOYIELD)
            End If
        End If
    Else
        Dim nMsg As Msg
        GetWindowThreadProcessId hwnd, pID
        If (pID = lParam) And IsWindow(hwnd) Then
            If PeekMessage(nMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD) Then
                Do
                    TranslateMessage nMsg
                    DispatchMessage nMsg
                Loop While PeekMessage(nMsg, hwnd, 0, 0, PM_REMOVE + PM_NOYIELD)
            End If
        End If
        WinEvents = True
    End If
End Function
