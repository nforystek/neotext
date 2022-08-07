Attribute VB_Name = "modTasks"
#Const modTasks = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const MSG_PEEK = 2

Private Const PM_NOREMOVE = &H0
Private Const PM_REMOVE = &H1
Private Const PM_NOYIELD = &H2

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private doStack As Long
Private gMsg As Msg

Public Function WinEvents(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim pID As Long
    Static wMsg As Msg
    If (lParam <= 0) And (lParam >= -3) Then
        If PeekMessage(wMsg, hWnd, 0, 0, PM_REMOVE + PM_NOYIELD) Then
            Do
                TranslateMessage wMsg
                DispatchMessage wMsg
            Loop While PeekMessage(wMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD)
        ElseIf (lParam >= -1) Then
            Sleep 1
        End If
        If (lParam = -3) Then Sleep 1
        If doStack = 1 Then
            DoEvents
        End If
    ElseIf (lParam = -4) Then
        If (doStack > 1) And (doStack < 3) Then
            If PeekMessage(wMsg, 0, 0, 0, PM_NOREMOVE + PM_NOYIELD) Then
                Do
                    Sleep 0
                Loop While PeekMessage(wMsg, hWnd, 0, 0, PM_NOREMOVE + PM_NOYIELD)
            End If
        End If
    Else
        Dim nMsg As Msg
        GetWindowThreadProcessId hWnd, pID
        If (pID = lParam) And IsWindow(hWnd) Then
            If PeekMessage(nMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD) Then
                Do
                    TranslateMessage nMsg
                    DispatchMessage nMsg
                Loop While PeekMessage(nMsg, hWnd, 0, 0, PM_REMOVE + PM_NOYIELD)
            End If
        End If
        WinEvents = True
    End If
End Function

Public Sub DoTasks()
    
    Static dMsg As Msg
    Do While PeekMessage(dMsg, -1, 0, 0, PM_REMOVE + PM_NOYIELD)
        TranslateMessage dMsg
        DispatchMessage dMsg
    Loop
            
    If PeekMessage(dMsg, 0, 0, 0, PM_NOREMOVE + PM_NOYIELD) Then
        doStack = doStack + 1
        If doStack = 1 Then
            EnumWindows AddressOf WinEvents, GetCurrentProcessId
            Do While PeekMessage(dMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD)
                TranslateMessage dMsg
                DispatchMessage dMsg
                EnumWindows AddressOf WinEvents, GetCurrentProcessId
            Loop
            EnumWindows AddressOf WinEvents, -4
        End If
        doStack = doStack - 1
        EnumWindows AddressOf WinEvents, -2
    ElseIf doStack = 1 Then
        EnumWindows AddressOf WinEvents, -1
    Else
        EnumWindows AddressOf WinEvents, -3
    End If
    
End Sub

