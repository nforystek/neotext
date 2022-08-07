Attribute VB_Name = "modWinProc"
#Const modWinProc = -1
Option Explicit
Option Compare Binary

Public Const ftpLocalSize = 32768 '65536 ' - 1
Public Const ftpBufferSize = 16384  ' - 1
Public Const ftpPacketSize = 8192  ' - 1

Private Type CWPSTRUCT
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)

Private Const WM_QUEUESYNC = &H23
Private Const WM_QUERYOPEN = &H13
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16

Private Const WM_POWERBROADCAST = 536
Private Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Private Const PBT_APMSUSPEND As Long = &H4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = -4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Enum AutoRated
    HaltRates = -3
    OpenRates = -1
    SendCalls = 1
    SendWidth = 2
    ReadCalls = 3
    ReadWidth = 4

    SendEvent = 5
    ReadEvent = 6
End Enum

Private mSockets As Collection

Private mSubProc As Long
Private mForm As frmThread

Public Function SocketsInitialize()
    If Not modSockets.WinsockControl Then
        If mForm Is Nothing Then
            Set mForm = New frmThread
            Load mForm
            mSubProc = SetWindowLong(mForm.hwnd, GWL_WNDPROC, AddressOf WinsockEvents)
        End If
    End If
    modSockets.SocketsInitialize
End Function

Public Function SocketsCleanUp()
    modSockets.SocketsCleanUp
    If Not modSockets.WinsockControl Then
        If Not mForm Is Nothing Then
            SetWindowLong mForm.hwnd, GWL_WNDPROC, mSubProc
            Unload mForm
        
            Set mForm = Nothing
        End If
    End If
End Function

Public Function RegisterSocket(ByRef sock As ISocket)
    If mSockets Is Nothing Then
        Set mSockets = New Collection
    End If
'    Dim ptrs As Memory
'    ptrs = ObjectPointers(sock, 3, 12, 1, 20)
        
    
    mSockets.Add sock, "h" & sock.Handle

    
End Function
Public Function UnregisterSocket(ByRef sock As ISocket)

    mSockets.Remove "h" & sock.Handle

    If mSockets.count = 0 Then
        Set mSockets = Nothing
        TermCerts
    End If
    
End Function

Public Property Get hwnd() As Long
    hwnd = mForm.hwnd
End Property

'Private Function WinsockEvents(ByVal args As Long, Optional arg1 As Long, Optional ByVal arg2 As Long, Optional ByVal arg3 As Long, Optional ByVal arg4 As Long) As Long

Private Function WinsockEvents(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error GoTo noproc
On Local Error GoTo noproc
    Dim ret As Long
    
    

'    Dim tmp As CWPSTRUCT
'    Dim allarg As Msg
'    Dim hwnd As Long
'    Dim uMsg As Long
'    Dim wParam As Long
'    Dim lParam As Long
'    RtlMoveMemory ByVal VarPtr(tmp), ByVal VarPtr(args), LenB(tmp)
'    RtlMoveMemory tmp.lParam, ByVal VarPtr(tmp), 4&
'    RtlMoveMemory tmp.wParam, ByVal VarPtr(tmp) + 4, 4&
'    RtlMoveMemory tmp.Message, ByVal VarPtr(tmp) + 8, 4&
'    RtlMoveMemory tmp.hwnd, ByVal VarPtr(tmp) + 12, 4&
'    hwnd = tmp.lParam
'    uMsg = tmp.wParam
'    wParam = tmp.uMsg
'    lParam = tmp.hwnd
'    RtlMoveMemory ByVal VarPtr(allarg.pt), ByVal VarPtr(tmp.lParam), (LenB(tmp) - (LenB(allarg.pt) + LenB(allarg.time)))
'    RtlMoveMemory allarg.hwnd, ByVal VarPtr(tmp), LenB(allarg) + LenB(tmp)
'    hwnd = allarg.hwnd
'    uMsg = allarg.Message
'    wParam = allarg.wParam
'    lParam = allarg.lParam

   ' Debug.Print tmp.hwnd & " " & tmp.uMsg & " " & tmp.wParam & " " & tmp.lParam


    If (mSubProc <> 0) Then

        Dim sck As ISocket
        Set sck = mSockets("h" & wParam)
        vbaObjSetAddref sck, sck.Address
        
        'Set sck = mSockets("h" & wParam)
       ' Debug.Print GlobalSize(ObjPtr(sck))
        Select Case uMsg
            Case WM_POWERBROADCAST
                ret = sck.EventRaised(wParam, WM_POWERBROADCAST)
            Case WM_WINSOCK, WINSOCK_MESSAGE, SOCKET_MESSAGE
                'Debug.Print "WM_WINSOCK " & hwnd & ", " & uMsg & ", " & wParam & ", " & lParam
                ret = sck.EventRaised(wParam, lParam)

            'Case Else

        End Select
        vbaObjSet sck, 0
        
        Set sck = Nothing
      '  vbaObjSetAddref mSockets("h" & wParam), ObjPtr(sck)
        
    
    End If

    
    GoTo exitfunc:
    
Exit Function
noproc:
    Err.Clear
exitfunc:
    RtlMoveMemory WinsockEvents, ByVal VarPtr(ret), 4&
End Function


