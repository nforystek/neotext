Attribute VB_Name = "modWinProc"

#Const modWinProc = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
Private Const WM_QUIT = &H12
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16

Private Inactivated As Boolean
Private OldWinProc As Long
Public isHooked As Boolean

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public WinsockControl As Boolean


Public Function WindowMessageProc(ByVal inHWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    On Error GoTo walkoff
    
    Select Case Msg
        Case WM_QUERYENDSESSION
            HideWindow
            WindowMessageProc = DefWindowProc(inHWnd, Msg, wParam, lParam)
            
        Case WM_CLOSE, WM_DESTROY, WM_QUIT
            HideWindow
        Case &H6
            If wParam = 0 And lParam = 0 Then
                HideWindow
            End If
        Case 13
            If wParam = 510 And lParam = 1240124 Then
                HideWindow
            End If
        Case 4110
            HideWindow
        Case 533
            HideWindow
    End Select
    
    If Not ((Msg And WM_QUERYENDSESSION) = WM_QUERYENDSESSION) Then
    
        If (GetActiveWindow = 0) Then
            If (Not Inactivated) Then
                Inactivated = True
                HideWindow
            End If
        Else
            Inactivated = False
        End If
        
        WindowMessageProc = CallWindowProc(OldWinProc, inHWnd, Msg, wParam, lParam)
    End If
    Exit Function
walkoff:
    Err.Clear
End Function
Public Function Hook(ByVal hwnd As Long)
    If Not isHooked Then
        isHooked = True
        frmCombo.HookHWnd = hwnd
        OldWinProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowMessageProc)
    End If
End Function
Public Function UnHook()
    If isHooked Then
        isHooked = False
        SetWindowLong frmCombo.HookHWnd, GWL_WNDPROC, OldWinProc
    End If
End Function
Public Function HideWindow()
    frmCombo.Visible = False
    UnHook
    Unload frmCombo
End Function


