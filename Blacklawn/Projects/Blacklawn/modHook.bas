Attribute VB_Name = "modHook"
#Const modHook = -1
Option Explicit
'TOP DOWN

Option Private Module
Public Enum PROCESS_DPI_AWARENESS

    PROCESS_DPI_UNAWARE = 0
    PROCESS_SYSTEM_DPI_AWARE = 1
    PROCESS_PER_MONITOR_DPI_AWARE = 2
End Enum
Const WM_DPICHANGED = 736

Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Static Function HookObj(ByRef obj)
    Static hc As Collection
    Static ha As Collection
    If IsNumeric(obj) Then
        If hc Is Nothing And obj > 0 Then
            DestroyWindow obj
        Else
            If obj < 0 Then
                HookObj = ha("k" & -obj)
            Else
                Set HookObj = hc("k" & obj)
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
                If hc(cnt).hwnd = obj.hwnd Then
                    SetWindowLong obj.hwnd, _
                    GWL_WNDPROC, ha("k" & obj.hwnd)
                    hc.Remove "k" & obj.hwnd
                    ha.Remove "k" & obj.hwnd
                    GoTo hookok
                End If
            Next
        End If
        hc.Add obj, "k" & obj.hwnd
        ha.Add SetWindowLong(obj.hwnd, GWL_WNDPROC, _
        AddressOf ControlWndProc), "k" & obj.hwnd
    End If
hookok:
    If hc.Count = 0 Then
        Set hc = Nothing
        Set ha = Nothing
    End If
End Function

Private Function ControlWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (HookObj(-hwnd) <> 0) Then
        Debug.Print TypeName(HookObj(hwnd)) & ", " & hwnd & ", " & uMsg & ", " & wParam & ", " & lParam
        If CallWindowProc(HookObj(-hwnd), hwnd, uMsg, wParam, lParam) = 0 Then
            Select Case uMsg
                Case WM_DPICHANGED
                    ControlWndProc = PROCESS_DPI_AWARENESS.PROCESS_SYSTEM_DPI_AWARE
            End Select
            ControlWndProc = 1
        Else
            ControlWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
        End If
    End If
End Function


