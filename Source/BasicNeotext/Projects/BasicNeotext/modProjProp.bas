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


Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Private Const WM_USER = &H400
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_EXGETSEL = (WM_USER + 52)
Private Const EM_GETSEL = &HB0&
Private Const EM_SETSEL = &HB1&
Private Const EM_HIDESELECTION = (WM_USER + 63)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Type POINTAPI
        x As Long
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


Private hWndProp As Long
Private hWndCust As Long
Private hWndProc As Long


Public Sub ItterateDialogs(ByRef VBProjects As VBProjects)
    
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
    
    EnumWindows AddressOf ItterateDialogsWinEvents, VBPID
    
    If hWndCust = 0 And flagCust Then
        SetBNSettings
    End If
    
    If hWndProc = 0 And flagProc Then
        UpdateAttributeToCommentDescriptions VBProjects
    End If

    If (hWndProp <> 0) Then
        FixConditionalCompile hWndProp
        hWndProp = 0
    End If
  
End Sub


Private Function ItterateDialogsWinEvents(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinEvents = Not SubCheckHwnds(hwnd, lParam)
    EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents1, lParam
End Function

Private Function ItterateDialogsWinChildEvents1(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinChildEvents1 = Not SubCheckHwnds(hwnd, lParam)
    EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents2, lParam
End Function

Private Function ItterateDialogsWinChildEvents2(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    ItterateDialogsWinChildEvents2 = Not SubCheckHwnds(hwnd, lParam)
    EnumChildWindows hwnd, AddressOf ItterateDialogsWinChildEvents1, lParam
End Function

Private Function SubCheckHwnds(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim txt As String
    Dim lSize As Long
    txt = String(255, Chr(0)) 'Space$(255)
    lSize = Len(txt)
    Call GetWindowText(hwnd, txt, lSize)
    txt = Trim(Replace(txt, Chr(0), ""))
'    If lSize > 0 Then
'        txt = Left$(txt, lSize)
'    End If
    If GetWindowThreadProcessId(hwnd, lSize) Then
        If lSize = lParam Then
        
            If (txt = "Con&ditional Compilation Arguments:") And hWndProp = 0 Then
                hWndProp = GetWindow(hwnd, GW_HWNDNEXT)
            End If
        
            If (txt = "Customize") And hWndCust = 0 Then
                hWndCust = hwnd
            End If
        
            If (txt = "Procedure Attributes") And hWndProc = 0 Then
                hWndProc = hwnd
            End If
            
        End If
    End If
    SubCheckHwnds = (hWndProp <> 0) Or (hWndCust <> 0) Or (hWndProc <> 0)
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




