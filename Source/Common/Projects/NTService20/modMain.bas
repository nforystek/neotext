#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
Option Compare Binary

Option Private Module

Private Const PM_NOREMOVE = &H0
Private Const PM_REMOVE = &H1
Private Const PM_NOYIELD = &H2

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND = &H112
Private Const WM_SHOWWINDOW = &H18
Private Const WM_PAINT = &HF
Private Const WM_DESTROY = &H2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private ServiceFormHWnd As New Collection
Private ServiceFormOldProc As New Collection
Private ServiceObjectLPtrs As New Collection

Public Sub Main()

End Sub

Private Property Get PtrObj(ByRef lptr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    CopyMemory NewObj, lptr, 4&
    Set PtrObj = NewObj
    CopyMemory NewObj, lZero, 4&
End Property

Private Function EnumAddWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    If IsWindow(hwnd) And Not ServiceFormExists(hwnd) Then
        Dim lpProcessID As Long
        Dim ctrl As Controller
        GetWindowThreadProcessId hwnd, lpProcessID
        Set ctrl = PtrObj(lParam)
        If lpProcessID = ctrl.ProcessID Then
            Set ctrl = Nothing
            AddServiceForm hwnd, lParam
        Else
            Set ctrl = Nothing
        End If
    End If
    EnumAddWindowsProc = True
End Function

Public Function EnumAddServiceForms(ByVal lParam As Long)
    EnumWindows AddressOf EnumAddWindowsProc, lParam
End Function

Public Sub ClearServiceForms()
    Do While ServiceFormHWnd.Count > 0
        RemoveServiceForm ServiceFormHWnd(1).hwnd
    Loop
End Sub

Public Function WindowProcedure(ByVal inHwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo catch
    Dim c As Controller
    Dim proc As Long
    Dim lptr As Long
    Select Case Msg
        Case WM_ENDSESSION
            If ServiceFormExists(inHwnd, proc, lptr) Then
                Set c = PtrObj(lptr)
                If Not (c Is Nothing) Then
                    c.WinServe_LoggedOffService
                    Set c = Nothing
                End If
            End If
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_QUERYENDSESSION
            If ServiceFormExists(inHwnd, proc, lptr) Then
                EnumAddServiceForms ServiceObjectLPtrs("h" & inHwnd)
            End If
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_SYSCOMMAND
            If (wParam = 61536) Then
                If ServiceFormExists(inHwnd, proc, lptr) Then
                    ServiceFormHWnd("h" & inHwnd).Hide
                End If
            End If
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_DESTROY, WM_NCDESTROY, WM_MDIDESTROY
            If ServiceFormExists(inHwnd, proc, lptr) Then
                Set c = PtrObj(lptr)
                If Not (c Is Nothing) Then
                    If c.IsLegacyOS Then
                        c.Win98Svc_StopService
                    End If
                    Set c = Nothing
                End If
                WindowProcedure = CallWindowProc(proc, inHwnd, Msg, wParam, lParam)
            End If
        Case WM_NULL
            If ServiceFormExists(inHwnd, proc, lptr) Then
                Set c = PtrObj(lptr)
                If Not (c Is Nothing) Then
                    If c.IsLegacyOS Then
                        c.Running = c.Win98Svc_StartService
                    End If
                    Set c = Nothing
                End If
                WindowProcedure = CallWindowProc(proc, inHwnd, Msg, wParam, lParam)
            Else
                WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
            End If
        Case WM_CLOSE
            CloseWindow inHwnd
        Case Else
            If ServiceFormExists(inHwnd, proc, lptr) Then
                If CallWindowProc(proc, inHwnd, Msg, wParam, lParam) = 0 Then
                    WindowProcedure = 1
                Else
                    WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
                End If
            End If
    End Select
    Exit Function
catch:
    Err.Clear
    WindowProcedure = 0
End Function

Public Sub PostQuit(ByRef frm)
    Dim nxtHwnd As Long
    Do While ServiceFormHWnd.Count > 0
        nxtHwnd = ServiceFormHWnd(1).hwnd
        RemoveServiceForm ServiceFormHWnd(1)
        CloseWindow nxtHwnd
    Loop
End Sub

Public Sub AddServiceForm(ByRef frm, ByVal ptr As Long)
    If (Not ServiceFormExists(frm)) Then
        Select Case TypeName(frm)
            Case "Long"
                Dim blankForm As New FormHWnd
                blankForm.hwnd = frm
                ServiceFormHWnd.Add blankForm, "h" & blankForm.hwnd
                ServiceObjectLPtrs.Add ptr, "h" & blankForm.hwnd
                ServiceFormOldProc.Add SetWindowLong(blankForm.hwnd, GWL_WNDPROC, AddressOf WindowProcedure), "h" & blankForm.hwnd
            Case Else
                ServiceFormHWnd.Add frm, "h" & frm.hwnd
                ServiceObjectLPtrs.Add ptr, "h" & frm.hwnd
                ServiceFormOldProc.Add SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf WindowProcedure), "h" & frm.hwnd
        End Select
    End If
End Sub

Public Sub RemoveServiceForm(ByRef frm)
    If ServiceFormExists(frm) Then
        Select Case TypeName(frm)
            Case "Long"
                SetWindowLong frm, GWL_WNDPROC, ServiceFormOldProc("h" & frm)
                ServiceFormHWnd.Remove "h" & frm
                ServiceFormOldProc.Remove "h" & frm
                ServiceObjectLPtrs.Remove "h" & frm
            Case Else
                SetWindowLong frm.hwnd, GWL_WNDPROC, ServiceFormOldProc("h" & frm.hwnd)
                ServiceFormHWnd.Remove "h" & frm.hwnd
                ServiceFormOldProc.Remove "h" & frm.hwnd
                ServiceObjectLPtrs.Remove "h" & frm.hwnd
                Unload frm
        End Select
    End If
End Sub

Public Function ServiceFormExists(ByRef frm, Optional ByRef proc As Long = 0, Optional ByRef lptr As Long = 0) As Boolean
    If Not ServiceFormHWnd Is Nothing Then
        If ServiceFormHWnd.Count > 0 Then
            Dim tmp As Object
            Select Case TypeName(frm)
                Case "String"
                    For Each tmp In ServiceFormHWnd
                        If tmp.ProcessName = frm And tmp.ProcessID = GetCurrentProcessId Then
                            proc = ServiceFormOldProc("h" & frm.hwnd)
                            lptr = ServiceObjectLPtrs("h" & frm.hwnd)
                            ServiceFormExists = True
                            Exit For
                        End If
                    Next
                Case "Long"
                    For Each tmp In ServiceFormHWnd
                        If tmp.hwnd = frm Then
                            proc = ServiceFormOldProc("h" & frm)
                            lptr = ServiceObjectLPtrs("h" & frm)
                            ServiceFormExists = True
                            Exit For
                        End If
                    Next
                Case Else
                    For Each tmp In ServiceFormHWnd
                        If tmp.hwnd = frm.hwnd Then
                            proc = ServiceFormOldProc("h" & frm.hwnd)
                            lptr = ServiceObjectLPtrs("h" & frm.hwnd)
                            ServiceFormExists = True
                            Exit For
                        End If
                    Next
            End Select
            Set tmp = Nothing
        End If
    End If
End Function

