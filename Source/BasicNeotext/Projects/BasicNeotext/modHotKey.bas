#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modHotKey"

#Const modMWheel = -1
Option Explicit
'TOP DOWN
Option Compare Text

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId _
    As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, _
    ByVal cch As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
    Long, ByVal lParam As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal VK As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Global Const WM_HOTKEY = &H312      'message sent when hotkey registered with RegisterHotKey is pressed
Public Const MOD_SHIFT = &H4        'check the WM_HOTKEY message to see if the shift key was held down
Global Const MOD_WIN = &H8          'check the WM_HOTKEY message to see if the windows key was held down
Global Const MIN_HOTKEY = &H5F      'user defined, this is my unique (to this process) id for the minimize hotkey
Global Const RST_HOTKEY = &H6F      'user defined, this is my unique (to this process) id for the restore hotkey

Private Const WM_USER = &H400

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SETTEXT = &HC

Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B

Private Const GWL_WNDPROC = -4

Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Sub UpdateGUI()
    On Error GoTo exitupdate
    If Hooks.Count > 0 Then
        Dim cnt As Long
        Dim ptr As Long
        Dim hwnd As Long
        
        Dim conn As Connect
        Dim winTitle As String
        For cnt = 1 To Hooks.Count
            FromSci Hooks(cnt), hwnd, , ptr

            Set conn = PtrObj(ptr)
            conn.SetUIState WindowText(hwnd)
            Set conn = Nothing

        Next
    End If
    Exit Sub
exitupdate:
    Err.Clear
End Sub

Private Sub ToSci(ByRef sci As String, ByVal lhwnd As Long, ByVal oldProc As Long, ByVal connPtr As Long)
    sci = lhwnd & ":" & oldProc & ":" & connPtr
End Sub
Private Sub FromSci(ByVal sci As String, Optional ByRef lhwnd As Long, Optional ByRef oldProc As Long, Optional ByRef connPtr As Long)
    lhwnd = CLng(RemoveNextArg(sci, ":"))
    oldProc = CLng(RemoveNextArg(sci, ":"))
    connPtr = CLng(RemoveNextArg(sci, ":"))
End Sub

Public Function WindowClassName(ByVal hwnd As Long) As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetClassName(hwnd, sBuffer, lSize)
    If lSize > 0 Then
        WindowClassName = Replace(Left$(sBuffer, lSize), Chr(0), "")
    End If
End Function

Public Function WindowText(ByVal hwnd As Long) As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetWindowText(hwnd, sBuffer, lSize)
    If lSize > 0 Then
        WindowText = Trim(Replace(Left$(sBuffer, lSize), Chr(0), ""))
    End If
End Function

Private Function PtrObj(ByRef lPtr As Long) As Connect
    Dim lZero As Long
    Dim NewObj As Connect
    RtlMoveMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    RtlMoveMemory NewObj, lZero, 4&
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

        
    If uMsg = WM_HOTKEY Then
            
        If Hooks.Count > 0 And wParam <= Hooks.Count Then
    
            Dim oldProc As Long
            Dim connPtr As Long
            Dim conn As Connect
            Dim handled As Boolean
            Dim CancelDefault As Boolean
                   
            FromSci Hooks(wParam), , oldProc, connPtr
            If connPtr <> 0 And oldProc <> 0 Then
 
                Set conn = PtrObj(connPtr)
                
                
                
                conn.StartEvent Nothing, handled, CancelDefault
                
                If (Not CancelDefault) Then
                
                    Dim w As WINFOCUS
                    w.Focus = hwnd
                    
                    UnregisterHotKey hwnd, wParam
                    modSendKeys.Send "{f5}", w, False, True
                    RegisterHotKey hwnd, wParam, 0, VK_F5
                
                End If
                
            End If
            
            Set conn = Nothing
            
'            Dim winTitle As String
'            winTitle = WindowText(hwnd)
'
'            If InStr(winTitle, "Microsoft Visual Basic [design]") > 0 Then
'                conn.SetUIState Design
'            ElseIf InStr(winTitle, "Microsoft Visual Basic [running]") > 0 Then
'                conn.SetUIState Running
'            ElseIf InStr(winTitle, "Microsoft Visual Basic [run]") > 0 Then
'                conn.SetUIState Run
'            ElseIf InStr(winTitle, "Microsoft Visual Basic [break]") > 0 Then
'                conn.SetUIState Break
'            End If
            
            If (Not handled) Then
                If uMsg = &H2 Or uMsg = &H82 Or uMsg = &H210 Then
                    SetWindowLong hwnd, GWL_WNDPROC, oldProc
                    Exit Function
                End If
                WindowProc = CallWindowProc(oldProc, hwnd, uMsg, wParam, lParam)
            End If
        End If
    End If
End Function

Public Sub UnHook()
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    'Ensures that you don't try to unsubclass the window when
    'it is not subclassed.
    If Hooks.Count = 0 Then Exit Sub
    
    Dim cnt As Long
    Dim lhwnd As Long
    Dim oldProc As Long
    Dim connPtr As Long
    
    For cnt = Hooks.Count To 1 Step -1
        FromSci Hooks(cnt), lhwnd, oldProc, connPtr
        UnregisterHotKey lhwnd, cnt
        SetWindowLong lhwnd, GWL_WNDPROC, oldProc
    Next
    
    Do Until Hooks.Count = 0
        Hooks.Remove 1
    Loop
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub


Function EnumThreadProc(ByVal lhwnd As Long, ByVal lParam As Long) As Long

    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    EnumThreadProc = True
    
    Dim WinClass As String
    Dim winTitle As String
     
    winTitle = WindowText(lhwnd)
    WinClass = WindowClassName(lhwnd)
    
    If InStr(winTitle, "Microsoft Visual Basic") > 0 And InStr(WinClass, "wndclass_desked_gsk") > 0 Then

        If Hooks.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To Hooks.Count
                If Trim(NextArg(Hooks(cnt), ":")) = Trim(CStr(lhwnd)) Then Exit Function
            Next
        End If
        
        Dim sci As String
        ToSci sci, lhwnd, GetWindowLong(lhwnd, GWL_WNDPROC), lParam
        Hooks.Add sci, "H" & lhwnd
        RegisterHotKey lhwnd, Hooks.Count, 0, VK_F5
        SetWindowLong lhwnd, GWL_WNDPROC, AddressOf WindowProc

    End If
   
    Exit Function
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Function








