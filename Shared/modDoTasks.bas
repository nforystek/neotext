Attribute VB_Name = "modDoTasks"
#Const [True] = -1
#Const [False] = 0



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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If Not modCommon Then

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
    Dim pId As Long
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
        GetWindowThreadProcessId hwnd, pId
        If (pId = lParam) And IsWindow(hwnd) Then
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

#End If

'#####################################################################################
Public Sub DoPower(Optional ByVal Sleeper As Long = 3)
    DoEvents
    Sleep Sleeper
    #If Not modCommon Then
        modDoTasks.DoTasks
    #Else
        modCommon.DoTasks
    #End If
   ' DoTasks
End Sub
'#####################################################################################
Public Static Sub DoLoop()
    'designed for continuous
    'repeating looping calls
    'where the elapse timing
    'would a few millisecond
    Static elapse As Single
    Static latency As Single
    Static lastlat As Single
    Static multity As Long
    If elapse <> 0 Then
        elapse = Timer - elapse
        If elapse > latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then Sleep lastlat
            End Select
            lastlat = elapse - latency
            multity = multity + 16
        ElseIf elapse < latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then Sleep lastlat
            End Select
            lastlat = latency - elapse
            multity = multity + 4
        ElseIf lastlat <> 0 Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then Sleep lastlat
            End Select
            If lastlat > 0 Then
                If Not multity = 0 Then
                    multity = multity \ 2
                Else
                    multity = multity + 2
                End If
            ElseIf lastlat < 0 Then
                If Not multity = 1024 Then
                    multity = multity * 2
                Else
                    multity = multity - 2
                End If
            End If
        ElseIf lastlat = 0 Then
            lastlat = 1
        End If
        latency = elapse
    End If
    elapse = Timer
End Sub
