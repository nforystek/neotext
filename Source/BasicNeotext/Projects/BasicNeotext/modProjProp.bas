Attribute VB_Name = "modProjProp"
Option Explicit

Private Type RangeType
    StartPos As Long
    StopPos As Long
End Type

Public Type SetTextEx
    flags As Long
    codepage As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Public Const ST_DEFAULT = 0
Public Const ST_KEEPUNDO = 1
Public Const ST_SELECTION = 2
Public Const ST_NEWCHARS = 3
Public Const ST_UNICODE = 4

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5


Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_FRAME = &H400
Private Const RDW_INVALIDATE = &H1

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Private Const WM_USER = &H400
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_EXGETSEL = (WM_USER + 52)
Private Const EM_GETSEL = &HB0&
Private Const EM_SETSEL = &HB1&
Private Const EM_HIDESELECTION = (WM_USER + 63)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Private Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lparam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETREDRAW = &HB


Private Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

Private Const WA_INACTIVE = 0
Private Const WA_ACTIVE = 1
Private Const WA_CLICKACTIVE = 2

Private Const WM_SETCURSOR = &H20
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_CHILDACTIVATE = &H22



Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lparam As Long
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
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long



Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lparam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lparam As Long) As Boolean
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


Private hWndProp As Long
Private hWndCust As Long
Private hWndProc As Long
Private hWndMSVB As Long
Private hWndCode As String


Public Sub MSVBRedraw(ByVal IsEnabled As Boolean)
    Static DisHwnd As Long
    
    If (hWndMSVB <> 0) And ((Not IsEnabled) And (DisHwnd = 0)) Then
        DisHwnd = hWndMSVB

        SendMessage hWndMSVB, WM_SETREDRAW, 0, ByVal 0&
        
    End If
    
    If (hWndMSVB <> 0) And (IsEnabled And (DisHwnd <> 0)) Then

        SendMessage hWndMSVB, WM_SETREDRAW, 1, ByVal 0&
        
        RedrawWindow hWndMSVB, ByVal 0&, ByVal 0&, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_FRAME Or RDW_INVALIDATE

        DisHwnd = 0
        
    End If

End Sub


Private Sub CleanHooks()
    
    If Hooks.Count > 0 Then
        Dim cnt As Long
        cnt = 1
        Do While cnt <= Hooks.Count

            If IsWindowVisible(Hooks(cnt).hwnd) = 0 Then
                hWndCode = Replace(hWndCode, "h" & Hooks(cnt).hwnd, "")
                Hooks.Remove cnt
            Else
                cnt = cnt + 1
            End If
        Loop

    End If

End Sub

Public Sub ProcWindowSets(Optional ByVal hwnd As Long = 0)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        If (Not SubCheckChildHwnds(-1)) Then
            EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents3, 0
        End If
    End If
End Sub
Public Sub ItterateDialogs(ByRef VBInstance As VBE)
            
    Dim flagCust As Boolean
    Dim flagProc As Boolean
    
    If hWndProc <> 0 Then
        flagProc = True
        ProcWindowSets hWndProc
        hWndProc = 0
    End If

    If hWndCust <> 0 Then
        flagCust = True
        hWndCust = 0
    End If
    
    Dim VBI As New VBI
    VBI.VBPID = VBPID
    Set VBI.VBInstance = VBInstance
    
    EnumWindows AddressOf ItterateDialogsWinEvents, ObjPtr(VBI)
    
    Set VBI = Nothing

    If hWndCust = 0 And flagCust Then
        SetBNSettings
    End If
    
    If hWndProc = 0 And flagProc Then
        UpdateAttributeToCommentDescriptions VBInstance
        SubCheckChildHwnds 0
        CleanHooks
    End If

    If (hWndProp <> 0) Then
        FixConditionalCompile hWndProp
        hWndProp = 0
    End If
    
    If Hooks.Count > 0 Then

        Dim frm As FormHWnd
        For Each frm In Hooks

            frm.SaveVisibility

            If frm.CodeModule Is Nothing Then

                Set frm.CodeModule = GetCodeModuleByCaption(VBInstance, GetCaption(frm.hwnd))

            End If

        Next
    End If
    
End Sub

Private Function ItterateDialogsWinEvents(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
    ItterateDialogsWinEvents = Not SubCheckHwnds(hwnd, lparam)
    If ItterateDialogsWinEvents Then EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents1, lparam
End Function

Private Function ItterateDialogsWinChildEvents1(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
    ItterateDialogsWinChildEvents1 = Not SubCheckHwnds(hwnd, lparam)
    If ItterateDialogsWinChildEvents1 Then EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents2, lparam
End Function

Private Function ItterateDialogsWinChildEvents2(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
    ItterateDialogsWinChildEvents2 = Not SubCheckHwnds(hwnd, lparam)
    If ItterateDialogsWinChildEvents2 Then EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents1, lparam
End Function

Private Function ItterateDialogsWinChildEvents3(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
    ItterateDialogsWinChildEvents3 = Not SubCheckChildHwnds(hwnd)
End Function

Private Function SubCheckChildHwnds(ByVal hwnd As Long) As Boolean
    Static Check1 As Boolean
    Static Check2 As Boolean

    If hwnd = 0 Then
        Check1 = False
        Check2 = False
    ElseIf hwnd > 0 Then
        Dim pId As Long
        Dim txt As String
        txt = GetCaption(hwnd)
        
        Dim cls As String
        cls = GetClass(hwnd)
    
        If GetWindowThreadProcessId(hwnd, pId) Then
            If pId = VBPID Then
            
                If ((cls = "Button") And (txt = "Ad&vanced >>")) And (Not Check1) Then

                    Dim at As POINTAPI
                    Dim p As POINTAPI
                    
                    GetCursorPos at
                    
                    p.x = 0
                    p.y = 0
                    
                    ClientToScreen hwnd, p
                    SetCursorPos p.x, p.y
                    mouse_event MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0
                    
                    SetCursorPos at.x, at.y

                    Check1 = True
                End If
            
                If ((cls = "Edit") And (txt = "")) And (Not Check2) Then
                    EnableWindow hwnd, False
                    Check2 = True
                End If
            
            End If
        End If
    End If
    SubCheckChildHwnds = Check1 And Check2

End Function


Private Function PtrObj(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    CopyMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    CopyMemory NewObj, lZero, 4&
End Function

Private Function SubCheckHwnds(ByVal hwnd As Long, ByVal lparam As Long) As Boolean

    Dim pId As Long
    Dim txt As String
    txt = GetCaption(hwnd)
    
    Dim cls As String
    cls = GetClass(hwnd)
    
    If GetWindowThreadProcessId(hwnd, pId) Then
        If pId = VBPID Then
        
            If (txt = "Con&ditional Compilation Arguments:") And hWndProp = 0 Then
                hWndProp = GetWindow(hwnd, GW_HWNDNEXT)
            End If
        
            If (txt = "Customize") And hWndCust = 0 Then
                hWndCust = hwnd
            End If
        
            If (txt = "Procedure Attributes") And hWndProc = 0 Then
                hWndProc = hwnd
            End If
            
            If (InStr(txt, "Microsoft Visual Basic") > 0) And hWndMSVB = 0 Then
                hWndMSVB = hwnd
            End If
            
            If (cls = "VbaWindow") And InStr(hWndCode, "h" & hwnd) = 0 Then
                hWndCode = hWndCode & "h" & hwnd
                
                Dim VBI As VBI
                Set VBI = PtrObj(lparam)
                
                Dim frm As FormHWnd
                Set frm = New FormHWnd
                
                frm.hwnd = hwnd
                
                frm.SaveVisibility
                
                MSVBRedraw False
                
                Hooks.Add frm, "h" & hwnd

                Set frm = Nothing
                Set VBI = Nothing

                MSVBRedraw True
            End If

            If (cls = "#32770") Then
                MSVBRedraw True
            End If
        End If
    End If
    SubCheckHwnds = ((hWndProp <> 0) Or (hWndCust <> 0) Or (hWndProc <> 0)) And (hWndMSVB <> 0)
End Function

Private Sub FixConditionalCompile(ByVal hwnd As Long)

    Dim invar As String
    Dim inval As String
    Dim outCC As String
    Dim inCC As String
    Dim atCC As String

    atCC = GetText(hwnd)
    inCC = atCC
    Do Until inCC = ""
        If (InStr(inCC, "=") > 0) And ((InStr(inCC, "=") < InStr(inCC, ":")) Or (InStr(inCC, ":") = 0)) Then

            invar = NextArg(NextArg(inCC, ":"), "=")
            inval = RemoveArg(NextArg(inCC, ":"), "=")
            If inval = "" Then Exit Sub
            RemoveNextArg inCC, ":"
            If (Not (Trim(invar) = "")) Then
                outCC = outCC & Trim(invar) & "=" & Trim(inval) & ":"
            End If
        Else
            outCC = outCC & RemoveNextArg(inCC, ":") & ":"
        End If
    Loop

    If Right(outCC, 1) = ":" And Not Right(atCC, 1) = ":" Then outCC = Left(outCC, Len(outCC) - 1)

    inCC = GetText(hwnd)
    If ((Not (inCC = outCC)) And (Not (inCC = ""))) Then
        Dim offsetat As Long
        Do
            offsetat = offsetat + 1
        Loop While Mid(outCC, offsetat, 1) = Mid(inCC, offsetat, 1) And offsetat < Len(inCC)
            
        SetText hwnd, outCC, offsetat
    End If

End Sub
Private Function GetClass(ByVal hwnd As Long) As String
    Dim cls As String
    Dim lSize As Long
    cls = String(255, Chr(0))
    lSize = Len(cls)
    Call GetClassName(hwnd, cls, lSize)
    GetClass = Trim(Replace(cls, Chr(0), ""))
End Function

Public Function GetCaption(ByVal hwnd As Long) As String
    Dim txt As String
    Dim lSize As Long
    txt = String(255, Chr(0))
    lSize = Len(txt)
    Call GetWindowText(hwnd, txt, lSize)
    GetCaption = Trim(Replace(txt, Chr(0), ""))
End Function
Private Function GetText(ByVal hwnd As Long) As String
    Dim Text As String
    Dim tlen As Long
    tlen = SendMessageStruct(hwnd, WM_GETTEXTLENGTH, 0&, 0&) + 1
    Text = String(tlen, Chr(0)) 'Space(tlen)
    Call SendMessageString(hwnd, WM_GETTEXT, tlen, Text)
    'GetText = Left(Text, tlen)
    GetText = Replace(Text, Chr(0), "")
End Function

Private Sub SetText(ByVal hwnd As Long, ByVal Text As String, ByVal OffsetsAt As Long)

    Dim start As Long
    Dim endpos As Long

    SendMessageLngPtr hwnd, EM_GETSEL, start, endpos
    SendMessageLngPtr hwnd, EM_HIDESELECTION, True, 0

    Dim tlen As Long
    tlen = LenB(Text)
    Call SendMessageString(hwnd, WM_SETTEXT, tlen, Text)
    If start >= OffsetsAt Then
        start = start - 1
        endpos = endpos - 1
    ElseIf endpos >= OffsetsAt Then
        endpos = endpos - 1
    End If
    
    SendMessageLngPtr hwnd, EM_SETSEL, ByVal start, ByVal endpos
    SendMessageLngPtr hwnd, EM_HIDESELECTION, False, 0

End Sub




