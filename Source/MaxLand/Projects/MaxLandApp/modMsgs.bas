Attribute VB_Name = "modMsgs"
Option Explicit
'TOP DOWN

Public Type ProcInfo
    Proc As Long
    hwnd As Long
    uMsg As Long
    wArg As Long
    lArg As Long
    addr As Long
End Type

Private PeakState As Long
Public MessageCount As Long
Public MessageQueue() As ProcInfo

Public Sub AddProgram(ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long, ByVal addr As Long)
    Static lstMsg As ProcInfo
    If Not Programs Is Nothing Then
        Dim StopHandle As Long
        StopHandle = Programs.Handle
        If Not (lstMsg.uMsg = uMsg And lstMsg.wArg = wArg And lstMsg.lArg = lArg And lstMsg.addr = addr) Then
            Do
                If Programs.Accepts(uMsg) Then
                    
                    If Not (lstMsg.hwnd = Programs.Handle And lstMsg.uMsg = uMsg And lstMsg.wArg = wArg And lstMsg.lArg = lArg And lstMsg.addr = addr) Then
                        MessageCount = MessageCount + 1
                        ReDim Preserve MessageQueue(1 To MessageCount) As ProcInfo
                        MessageQueue(MessageCount).Proc = PeakState
                        MessageQueue(MessageCount).hwnd = Programs.Handle
                        MessageQueue(MessageCount).uMsg = uMsg
                        MessageQueue(MessageCount).wArg = wArg
                        MessageQueue(MessageCount).lArg = lArg
                        MessageQueue(MessageCount).addr = addr
                        lstMsg.Proc = PeakState
                        lstMsg.hwnd = Programs.Handle
                        lstMsg.uMsg = uMsg
                        lstMsg.wArg = wArg
                        lstMsg.lArg = lArg
                        lstMsg.addr = addr
                    End If
                    
                End If
                ShiftPrograms
            Loop Until Programs.Handle = StopHandle
        End If
    End If
End Sub

Public Sub DelMessage(ByVal Number As Long)
    If (MessageCount > 0) Then
        Dim i As Long
        If (Number <= (MessageCount - 2)) And (Number > 0) Then
            For i = Number To MessageCount - 1
                MessageQueue(i) = MessageQueue(i + 1)
            Next
        End If
        MessageCount = MessageCount - 1
        If MessageCount = 0 Then
            Erase MessageQueue
        Else
            ReDim Preserve MessageQueue(1 To MessageCount) As ProcInfo
        End If
        
    End If
End Sub

Public Sub Initialize()
    ReDim MessageQueue(1 To 1) As ProcInfo
    PeakState = 1
End Sub

Public Sub Terminate()
    Erase MessageQueue
End Sub

Public Sub ProcessOrdered()
    Static wmTimer As Single
    If wmTimer = 0 Or (Timer - wmTimer) >= 1 Then
        wmTimer = Timer
        AddProgram WM_TIMER, wmTimer, 0, 0
    End If
                
    Static i As Long
    If Not Programs Is Nothing Then
        i = i + PeakState
        If (i > MessageCount) Then
            i = 1
        ElseIf (i < 1) Then
            i = MessageCount
        End If
        Dim StopHandle As Long
        StopHandle = Programs.Handle
        Do
            If ((i <= MessageCount) And (i > 0)) Then
                Select Case MessageQueue(i).Proc
                    Case 0
                        DelMessage i
                    Case Else
                        Programs.Routine i
                End Select
            End If
            ShiftPrograms
        Loop Until Programs.Handle = StopHandle
    End If
End Sub
Public Function HandleWindowProc(ByVal Proc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long, ByVal addr As Long) As Long
    If (addr <= 4) Then addr = val(AddressOf DefaultWindowProc)
    RtlMoveMemory ByVal VarPtr(HandleWindowProc), CallWindowProc(addr, hwnd, uMsg, wArg, lArg), 4&
End Function

Public Function DefaultWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long) As Long
    'return 1 to specify failure; return 0 to specify success; return -1 to specity unhandle and use default;
    DefaultWindowProc = 0
End Function

Public Function CustomWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long) As Long
    'return 1 to specify failure; return 0 to specify success; return -1 to specity unhandle and use default;
'    Select Case ReturnScenario
'        Case 1
'            CustomWindowProc = -1
'        Case 2
'            CustomWindowProc = 1
'        Case 3
'            CustomWindowProc = 0
'        Case 4
'            CustomWindowProc = Round(Rnd, 0)
'    End Select
'    Select Case CustomWindowProc
'        Case 1
'            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Failure"
'        Case 0
'            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Success"
'        Case -1
'            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Default"
'    End Select
End Function



