Attribute VB_Name = "modService"
#Const modService = -1
Option Explicit
'TOP DOWN
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

Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND = &H112
Private Const WM_SHOWWINDOW = &H18
Private Const WM_PAINT = &HF
Private Const WM_DESTROY = &H2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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

Private Function PtrObj(ByVal lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    CopyMemory NewObj, lPtr, ByVal 4&
    Set PtrObj = NewObj
    CopyMemory NewObj, lZero, ByVal 4&
End Function

Public Function IsWindows98() As Boolean

    Dim i As Long
    IsWindows98 = False
    
    On Error Resume Next

    i = RegisterServiceProcess(GetCurrentProcessId, 0)

    If Not (Err = 453) Then
        IsWindows98 = True
    Else
        Err.Clear
    End If

    On Error GoTo 0

End Function

Private Function EnumAddWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    If IsWindow(hwnd) And Not ServiceFormExists(hwnd) Then
        Dim lpProcessID As Long
        GetWindowThreadProcessId hwnd, lpProcessID
        If lpProcessID = GetCurrentProcessId Then
            AddServiceForm hwnd, lParam
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
    Select Case Msg

        Case WM_ENDSESSION
'            Set c = PtrObj(ServiceObjectLPtrs("h" & inHwnd))
'            If Not (c Is Nothing) Then
'                c.WinServe_UserLoggedOff
'                Set c = Nothing
'            End If
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_QUERYENDSESSION
            'EnumAddServiceForms ServiceObjectLPtrs("h" & inHwnd)
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_SYSCOMMAND
            On Error Resume Next
            If (wParam = 61536) Then
                ServiceFormHWnd("h" & inHwnd).Hide
            End If
            If Err Then Err.Clear
            On Error GoTo catch
            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
        Case WM_DESTROY, WM_NCDESTROY, WM_MDIDESTROY
            On Error Resume Next
            Set c = PtrObj(ServiceObjectLPtrs("h" & inHwnd))
            If Not (c Is Nothing) Then
                If c.IsLegacyOS Then
                    c.Win98Svc_StopService
                End If
                Set c = Nothing
            End If
            If Err Then Err.Clear
            On Error GoTo catch
            WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
        Case WM_NULL
            On Error Resume Next
            On Local Error GoTo 0
            Set c = PtrObj(ServiceObjectLPtrs("h" & inHwnd))
            If Not (c Is Nothing) Then
                If c.IsLegacyOS Then
                    c.Running = c.Win98Svc_StartService
                End If
                Set c = Nothing
            End If
            If Err Then Err.Clear
            On Error GoTo catch
            If ServiceFormExists(inHwnd) Then
                WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
            Else
                WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
            End If
            

        
'        Case WM_MDICREATE, WM_CREATE, WM_NCCREATE, WM_SHOWWINDOW, WM_PAINT, WM_NCPAINT, WM_MOVE, WM_SIZE, WM_ACTIVATE
'            If ServiceFormExists(inHwnd) Then CallWindowProc ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_MDIACTIVATE, WM_NCACTIVATE, WM_MDIREFRESHMENU, WM_ACTIVATEAPP, WM_CHILDACTIVATE, WM_PARENTNOTIFY
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_SIZING, WM_MOVING, WM_STYLECHANGING, WM_STYLECHANGED
'            WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
'            DefWindowProc inHwnd, Msg, wParam, lParam
'        Case WM_ENTERMENULOOP, WM_EXITMENULOOP, WM_NEXTMENU, WM_CUT, WM_COPY, WM_PASTE, WM_CLEAR, WM_UNDO
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_MDIRESTORE, WM_MDINEXT, WM_MDIMAXIMIZE, WM_MDITILE, WM_MDICASCADE, WM_MDIICONARRANGE, WM_MDIGETACTIVE, WM_MDISETMENU, WM_MDIREFRESHMENU
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_COMMAND, WM_KILLFOCUS, WM_SETFOCUS, WM_CLOSE, WM_QUIT, WM_ENABLE, WM_ENTERSIZEMOVE
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_NCCALCSIZE, WM_NCHITTEST, WM_NCMOUSEMOVE, WM_NCLBUTTONDOWN, WM_NCLBUTTONUP, WM_NCLBUTTONDBLCLK, WM_NCRBUTTONDOWN, WM_NCRBUTTONUP, WM_NCRBUTTONDBLCLK, WM_NCMBUTTONDOWN, WM_NCMBUTTONUP, WM_NCMBUTTONDBLCLK
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_MOUSEFIRST, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MBUTTONDBLCLK, WM_MOUSEWHEEL, WM_MOUSEHWHEEL, WM_MOUSEHOVER, WM_NCMOUSELEAVE, WM_MOUSELEAVE
'            WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
'        Case WM_KEYFIRST, WM_KEYDOWN, WM_KEYUP, WM_CHAR, WM_DEADCHAR, WM_SYSKEYDOWN, WM_SYSKEYUP, WM_SYSCHAR, WM_SYSDEADCHAR, WM_KEYLAST
'            WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
'        Case WM_QUERYOPEN, WM_ERASEBKGND, WM_SYSCOLORCHANGE
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_SYSTEMERROR, WM_CTLCOLOR, WM_WININICHANGE, WM_SETTINGCHANGE, WM_DEVMODECHANGE, WM_FONTCHANGE, WM_TIMECHANGE, WM_CANCELMODE, WM_SETCURSOR
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_MOUSEACTIVATE, WM_CHILDACTIVATE, WM_QUEUESYNC, WM_GETMINMAXINFO, WM_PAINTICON, WM_ICONERASEBKGND, WM_SPOOLERSTATUS, WM_DRAWITEM, WM_MEASUREITEM, WM_DELETEITEM, WM_VKEYTOITEM, WM_CHARTOITEM
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_SETFONT, WM_GETFONT, WM_SETHOTKEY, WM_GETHOTKEY, WM_QUERYDRAGICON, WM_COMPAREITEM, WM_COMPACTING
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_POWER, WM_COPYDATA, WM_CANCELJOURNAL, WM_NOTIFY, WM_INPUTLANGCHANGEREQUEST, WM_INPUTLANGCHANGE, WM_TCARD, WM_HELP, WM_NOTIFYFORMAT, WM_CONTEXTMENU, WM_DISPLAYCHANGE, WM_GETICON, WM_SETICON
'             WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_NULL, WM_SETREDRAW, WM_SETTEXT, WM_GETTEXT, WM_GETTEXTLENGTH
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_INITDIALOG, WM_TIMER, WM_HSCROLL, WM_VSCROLL, WM_INITMENU, WM_INITMENUPOPUP, WM_MENUSELECT, WM_MENUCHAR, WM_ENTERIDLE
'            WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
'        Case WM_CAPTURECHANGED, WM_POWERBROADCAST, WM_DEVICECHANGE
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_CTLCOLORMSGBOX, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORBTN, WM_CTLCOLORDLG, WM_CTLCOLORSCROLLBAR, WM_CTLCOLORSTATIC, WM_GETDLGCODE, WM_NEXTDLGCTL
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_IME_SETCONTEXT, WM_IME_NOTIFY, WM_IME_CONTROL, WM_IME_COMPOSITIONFULL, WM_IME_SELECT, WM_IME_CHAR, WM_IME_KEYDOWN, WM_IME_KEYUP, WM_IME_STARTCOMPOSITION, WM_IME_ENDCOMPOSITION, WM_IME_COMPOSITION, WM_IME_KEYLAST
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_RENDERFORMAT, WM_RENDERALLFORMATS, WM_DESTROYCLIPBOARD, WM_DRAWCLIPBOARD, WM_PAINTCLIPBOARD, WM_VSCROLLCLIPBOARD, WM_SIZECLIPBOARD, WM_ASKCBFORMATNAME, WM_CHANGECBCHAIN, WM_HSCROLLCLIPBOARD, WM_QUERYNEWPALETTE, WM_PALETTEISCHANGING, WM_PALETTECHANGED
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_HOTKEY, WM_PRINTCLIENT, WM_EXITSIZEMOVE, WM_DROPFILES
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_HANDHELDFIRST, WM_HANDHELDLAST, WM_PENWINFIRST, WM_PENWINLAST, WM_COALESCE_FIRST, WM_COALESCE_LAST, WM_DDE_FIRST, WM_DDE_INITIATE, WM_DDE_TERMINATE, WM_DDE_ADVISE, WM_DDE_UNADVISE, WM_DDE_ACK, WM_DDE_DATA, WM_DDE_REQUEST, WM_DDE_POKE, WM_DDE_EXECUTE, WM_DDE_LAST
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'        Case WM_USER, WM_APP
'            WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
            
            
        Case Else

'            If ServiceFormExists(inHwnd) Then
'                 WindowProcedure = CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam)
'            Else
'                WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
'            End If
            
            If ServiceFormExists(inHwnd) Then
                If CallWindowProc(ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam) = 0 Then
                    WindowProcedure = 1
                Else
                    WindowProcedure = DefWindowProc(inHwnd, Msg, wParam, lParam)
                End If
            End If
    End Select


'        Case WM_SYSCOMMAND
'            On Error Resume Next
'            If (wParam = 61536) Then
'                ServiceFormHWnd("h" & inHwnd).Hide
'            End If
'            CallWindowProc ServiceFormOldProc("h" & inHwnd), inHwnd, Msg, wParam, lParam
'            If Err Then Err.Clear
'            On Error GoTo 0
'            WindowProcedure = 1

'        Case WM_NULL
'            On Error Resume Next
'            Set c = Object(ServiceObjectLPtrs("h" & inHwnd))
'            If Not (c Is Nothing) Then
'                If c.IsLegacyOS Then
'                    c.Running = c.Win98Svc_StartService
'                End If
'                Set c = Nothing
'            End If
'            If Err Then Resume
'            On Error GoTo 0
'            WindowProcedure = 1
'

'    End Select
    Exit Function
catch:
    Err.Clear
    WindowProcedure = 1
End Function

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

Public Function ServiceFormExists(ByRef frm) As Boolean
    Dim tmp As Object
    On Error Resume Next
    Select Case TypeName(frm)
        Case "Long"
            Set tmp = ServiceFormHWnd("h" & frm)
        Case Else
            Set tmp = ServiceFormHWnd("h" & frm.hwnd)
    End Select
    ServiceFormExists = (Err.Number = 0)
    Set tmp = Nothing
    If Err Then Err.Clear
    On Error GoTo 0
End Function




