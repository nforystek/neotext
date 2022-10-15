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

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_FRAME = &H400
Private Const RDW_INVALIDATE = &H1

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hWnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long

Private Const WM_USER = &H400
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_EXGETSEL = (WM_USER + 52)
Private Const EM_GETSEL = &HB0&
Private Const EM_SETSEL = &HB1&
Private Const EM_HIDESELECTION = (WM_USER + 63)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETREDRAW = &HB

Private Type POINTAPI
        x As Long
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
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long



Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const GWL_WNDPROC = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long


Private hWndProp As Long
Private hWndCust As Long
Private hWndProc As Long
Private hWndMSVB As Long
Private hWndCode As String


Public Sub MSVBRedraw(ByVal IsEnabled As Boolean)
    Static DisHwnd As Long
    Dim cnt As Long
    
    If (hWndMSVB <> 0) And ((Not IsEnabled) And (DisHwnd = 0)) Then
        DisHwnd = hWndMSVB

        
'        If Hooks.count > 0 Then
'            For cnt = 1 To Hooks.count
'                Hooks(cnt).SaveVisibility
'               ' SendMessage Hooks(cnt).hWnd, WM_SETREDRAW, 0, ByVal 0&
'            Next
'        End If
        
        SendMessage hWndMSVB, WM_SETREDRAW, 0, ByVal 0&
        
        
    End If
    If (hWndMSVB <> 0) And (IsEnabled And (DisHwnd <> 0)) Then


'        If Hooks.count > 0 Then
'            For cnt = 1 To Hooks.count
'                If Hooks(cnt).Visible Then
'                    Hooks(cnt).Show
'                Else
'                    Hooks(cnt).Hide
'                End If
'            Next
'        End If

        
        SendMessage hWndMSVB, WM_SETREDRAW, 1, ByVal 0&
'        If Hooks.count > 0 Then
'            For cnt = 1 To Hooks.count
'                SendMessage Hooks(cnt).hWnd, WM_SETREDRAW, 1, ByVal 0&
'            Next
'        End If

        RedrawWindow hWndMSVB, ByVal 0&, ByVal 0&, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_FRAME Or RDW_INVALIDATE

        DisHwnd = 0
        
    End If

End Sub

Private Function FinCodeModuleByCaption(ByRef VBInstance As VBE, ByVal Caption As String) As CodeModule
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim cm As CodeModule
    
    Dim Member As Member
    For Each vbproj In VBInstance.VBProjects
        For Each vbcomp In vbproj.VBComponents
            If InStr(Caption, " " & vbcomp.Name & " ") > 0 Then
                Set cm = GetCodeModule2(vbcomp)
                If Not cm Is Nothing Then
                    If cm.CodePane.Window.Caption = Caption Then
                        Set FinCodeModuleByCaption = cm
                        Set cm = Nothing
                        Exit Function
                    End If
                End If
                Set cm = Nothing
            End If
            
        Next
    Next
End Function

Private Sub CleanHooks()
    
    If Hooks.count > 0 Then
        Dim cnt As Long
        cnt = 1
        Do While cnt <= Hooks.count

            If IsWindowVisible(Hooks(cnt).hWnd) = 0 Then
                hWndCode = Replace(hWndCode, "h" & Hooks(cnt).hWnd, "")
                Hooks.Remove cnt
            Else
                cnt = cnt + 1
            End If
        Loop

    End If

End Sub
Public Sub ItterateDialogs(ByRef VBInstance As VBE)
            
    Dim flagCust As Boolean
    Dim flagProc As Boolean
    
    If hWndProc <> 0 Then
        flagProc = True
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
        UpdateAttributeToCommentDescriptions VBInstance.VBProjects
        CleanHooks
    End If

    If (hWndProp <> 0) Then
        FixConditionalCompile hWndProp
        hWndProp = 0
    End If
    
    If Hooks.count > 0 Then

        Dim Frm As FormHWnd
        For Each Frm In Hooks

            Frm.SaveVisibility
            
            If Frm.CodeModule Is Nothing Then
                
                Set Frm.CodeModule = FinCodeModuleByCaption(VBInstance, GetCaption(Frm.hWnd))

            End If

        Next
    End If
    
End Sub

Private Function ItterateDialogsWinEvents(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinEvents = Not SubCheckHwnds(hWnd, lParam)
    If ItterateDialogsWinEvents Then EnumChildWindows hWnd, AddressOf ItterateDialogsWinChildEvents1, lParam
End Function

Private Function ItterateDialogsWinChildEvents1(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinChildEvents1 = Not SubCheckHwnds(hWnd, lParam)
    If ItterateDialogsWinChildEvents1 Then EnumChildWindows hWnd, AddressOf ItterateDialogsWinChildEvents2, lParam
End Function

Private Function ItterateDialogsWinChildEvents2(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinChildEvents2 = Not SubCheckHwnds(hWnd, lParam)
    If ItterateDialogsWinChildEvents2 Then EnumChildWindows hWnd, AddressOf ItterateDialogsWinChildEvents1, lParam
End Function


Private Function PtrObj(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    CopyMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    CopyMemory NewObj, lZero, 4&
End Function

Private Function SubCheckHwnds(ByVal hWnd As Long, ByVal lParam As Long) As Boolean

    Dim pid As Long
    Dim txt As String
    txt = GetCaption(hWnd)
    
    Dim cls As String
    cls = GetClass(hWnd)
    
    If GetWindowThreadProcessId(hWnd, pid) Then
        If pid = VBPID Then
        
            If (txt = "Con&ditional Compilation Arguments:") And hWndProp = 0 Then
                hWndProp = GetWindow(hWnd, GW_HWNDNEXT)
            End If
        
            If (txt = "Customize") And hWndCust = 0 Then
                hWndCust = hWnd
            End If
        
            If (txt = "Procedure Attributes") And hWndProc = 0 Then
                hWndProc = hWnd
            End If
            
            If (InStr(txt, "Microsoft Visual Basic") > 0) And hWndMSVB = 0 Then
                hWndMSVB = hWnd
            End If
            
            If (cls = "VbaWindow") And InStr(hWndCode, "h" & hWnd) = 0 Then
                hWndCode = hWndCode & "h" & hWnd
                
                Dim VBI As VBI
                Set VBI = PtrObj(lParam)
                
                Dim Frm As FormHWnd
                Set Frm = New FormHWnd
                
                Frm.hWnd = hWnd
                
                Frm.SaveVisibility
                
                MSVBRedraw False
                
                Hooks.Add Frm, "h" & hWnd

                Set Frm = Nothing
                Set VBI = Nothing

                MSVBRedraw True
            End If

            
        End If
    End If
    SubCheckHwnds = ((hWndProp <> 0) Or (hWndCust <> 0) Or (hWndProc <> 0)) And (hWndMSVB <> 0)
End Function

Private Sub FixConditionalCompile(ByVal hWnd As Long)

    Dim invar As String
    Dim inval As String
    Dim outCC As String
    Dim inCC As String
    Dim atCC As String

    atCC = GetText(hWnd)
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

    inCC = GetText(hWnd)
    If ((Not (inCC = outCC)) And (Not (inCC = ""))) Then
        Dim offsetat As Long
        Do
            offsetat = offsetat + 1
        Loop While Mid(outCC, offsetat, 1) = Mid(inCC, offsetat, 1) And offsetat < Len(inCC)
            
        SetText hWnd, outCC, offsetat
    End If

End Sub
Private Function GetClass(ByVal hWnd As Long) As String
    Dim cls As String
    Dim lSize As Long
    cls = String(255, Chr(0))
    lSize = Len(cls)
    Call GetClassName(hWnd, cls, lSize)
    GetClass = Trim(Replace(cls, Chr(0), ""))
End Function

Public Function GetCaption(ByVal hWnd As Long) As String
    Dim txt As String
    Dim lSize As Long
    txt = String(255, Chr(0))
    lSize = Len(txt)
    Call GetWindowText(hWnd, txt, lSize)
    GetCaption = Trim(Replace(txt, Chr(0), ""))
End Function
Private Function GetText(ByVal hWnd As Long) As String
    Dim Text As String
    Dim tlen As Long
    tlen = SendMessageStruct(hWnd, WM_GETTEXTLENGTH, 0&, 0&) + 1
    Text = String(tlen, Chr(0)) 'Space(tlen)
    Call SendMessageString(hWnd, WM_GETTEXT, tlen, Text)
    'GetText = Left(Text, tlen)
    GetText = Replace(Text, Chr(0), "")
End Function

Private Sub SetText(ByVal hWnd As Long, ByVal Text As String, ByVal OffsetsAt As Long)

    Dim start As Long
    Dim endpos As Long

    SendMessageLngPtr hWnd, EM_GETSEL, start, endpos
    SendMessageLngPtr hWnd, EM_HIDESELECTION, True, 0

    Dim tlen As Long
    tlen = LenB(Text)
    Call SendMessageString(hWnd, WM_SETTEXT, tlen, Text)
    If start >= OffsetsAt Then
        start = start - 1
        endpos = endpos - 1
    ElseIf endpos >= OffsetsAt Then
        endpos = endpos - 1
    End If
    
    SendMessageLngPtr hWnd, EM_SETSEL, ByVal start, ByVal endpos
    SendMessageLngPtr hWnd, EM_HIDESELECTION, False, 0

End Sub




