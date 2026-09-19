Attribute VB_Name = "modHook"
Option Explicit

Option Compare Binary
#Const modHook = -1

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEACTIVATE = &H21

Public Const WM_SETCURSOR = &H20

Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Static Function HookObj(ByRef Obj)
    Static hc As Collection
    Static ha As Collection
    If IsNumeric(Obj) Then
        If Not (hc Is Nothing) Then
            If Obj < 0 Then
                HookObj = ha("k" & -Obj)
            Else
                Set HookObj = hc("k" & Obj)
            End If
        End If
    Else
        If hc Is Nothing Then
            Set hc = New Collection
            Set ha = New Collection
        End If
        Dim cnt As Long
        If hc.Count > 0 Then
            For cnt = 1 To hc.Count
                If hc(cnt).hwnd = Obj.hwnd Then
                    SetWindowLong Obj.hwnd, _
                    GWL_WNDPROC, ha("k" & Obj.hwnd)
                    hc.Remove "k" & Obj.hwnd
                    ha.Remove "k" & Obj.hwnd
                    GoTo hookok
                End If
            Next
        End If
        hc.Add Obj, "k" & Obj.hwnd
        ha.Add SetWindowLong(Obj.hwnd, GWL_WNDPROC, _
        AddressOf ControlWndProc), "k" & Obj.hwnd
    End If
hookok:
    If Not (hc Is Nothing) Then
        If hc.Count = 0 Then
            Set hc = Nothing
            Set ha = Nothing
        End If
    End If
End Function

Private Function ControlWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (HookObj(-hwnd) <> 0) Then
        'Debug.Print TypeName(HookObj(hWnd)) & ", " & hWnd & ", " & uMsg & ", " & wParam & ", " & lParam
        Select Case uMsg
            Case WM_MOUSEWHEEL
                Dim ctl1 As WaveView

                Set ctl1 = HookObj(hwnd)
                If wParam > 0 Then
                    ctl1.ScrollUp
                ElseIf wParam < 0 Then
                    ctl1.ScrollDown
                End If
                Set ctl1 = Nothing
            
        End Select
        If CallWindowProc(HookObj(-hwnd), hwnd, uMsg, wParam, lParam) = 0 Then
            ControlWndProc = 1
        Else
            ControlWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
        End If
    End If
    
End Function



