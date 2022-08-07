Attribute VB_Name = "modWinProc"
Option Explicit
'TOP DOWN

Option Compare Binary

Option Private Module
Private wc As Boolean
Private uc As New Collection

Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)

Private Const WM_QUEUESYNC = &H23
Private Const WM_QUERYOPEN = &H13
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16

Private Const WM_POWERBROADCAST = 536
Private Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Private Const PBT_APMSUSPEND As Long = &H4

Public Sub ResetByHandle(ByVal handle As Long)
    Dim lPtr As Variant
    Dim sck As Socket
    For Each lPtr In uc
        Set sck = PtrObj(CLng(lPtr))
        If sck.handle = handle Then
            sck.ResetFlags
        End If
        Set sck = Nothing
    Next
End Sub
Private Property Get PtrObj(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    RtlMoveMemory NewObj, lZero, 4&
End Property

Public Property Get WinsockControl() As Boolean
    WinsockControl = wc
End Property
Public Property Let WinsockControl(ByVal NewVal As Boolean)
    wc = NewVal
End Property

Public Function ControlHostCount() As Long
    ControlHostCount = uc.Count
End Function

Public Sub UnSetControlHost(ByVal lPtr As Long, ByRef hWnd As Long)
    uc.Remove "H" & hWnd
    WindowTerminate hWnd
    hWnd = 0
End Sub

Public Sub SetControlHost(ByVal lPtr As Long, ByRef hWnd As Long)
    hWnd = WindowInitialize(AddressOf WndProc)
    uc.Add lPtr, "H" & hWnd
End Sub

Private Function ControlHostExists(ByRef hWnd As Long) As Boolean
    On Error Resume Next
    Dim tmp As Long
    tmp = uc("H" & hWnd)
    If Err Then
        Err.Clear
        ControlHostExists = False
    Else
        ControlHostExists = True
    End If
    On Error GoTo 0
End Function

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo notloaded

'    Static stacking As Long
'    stacking = stacking + 1
'    If stacking > uc.Count Then GoTo notloaded
'
'On Error GoTo selectnone
    Dim objSocket As Socket
    Select Case uMsg
        Case WM_WINSOCK
            Set objSocket = PtrObj(uc("H" & hWnd))
            If Not ((objSocket.Transmission And SocketPaused) = SocketPaused) Then
                WndProc = objSocket.EventRaised(WSAGetSelectEvent(lParam), WSAGetAsyncError(lParam))
            Else
                WndProc = 0
            End If
        Case WM_POWERBROADCAST
            Set objSocket = PtrObj(uc("H" & hWnd))
            If ((objSocket.Transmission And Direction.StandByPause) = Direction.StandByPause) Then
                Select Case wParam
                    Case PBT_APMRESUMEAUTOMATIC
                        If ((objSocket.Transmission And Direction.SocketPaused) = Direction.SocketPaused) Then
                            objSocket.Transmission = objSocket.Transmission - Direction.SocketPaused
                        End If
                    Case PBT_APMSUSPEND
                        If Not ((objSocket.Transmission And Direction.SocketPaused) = Direction.SocketPaused) Then
                            objSocket.Transmission = objSocket.Transmission + Direction.SocketPaused
                        End If
                End Select
            End If
    End Select
    Set objSocket = Nothing
    
    On Error GoTo 0
'    stacking = stacking - 1
    Exit Function
'selectnone:
'    If Err Then Err.Clear
'    On Error GoTo 0
'    Set objSocket = Nothing
'    stacking = stacking - 1
'    WndProc = 0
'    Exit Function
notloaded:
    On Error GoTo 0
    Set objSocket = Nothing
'    stacking = stacking - 1
    WndProc = 0
End Function


