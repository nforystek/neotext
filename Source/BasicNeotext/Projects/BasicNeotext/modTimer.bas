Attribute VB_Name = "modTimer"
#Const [True] = -1
#Const [False] = 0
#Const modTimer = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)

Public Sub StartTimer(ByRef obj As Timer)
    SetTimer obj.hwnd, ObjPtr(obj), obj.Interval, AddressOf TimerProc
End Sub
Public Sub StopTimer(ByRef obj As Timer)
    KillTimer obj.hwnd, ObjPtr(obj)
End Sub

Private Function PtrObj(ByVal lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    RtlMoveMemory NewObj, lZero, 4&
End Function

Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lpTimer As Long, ByVal dwTime As Long)
On Error GoTo notloaded
    Debug.Print "TIMER"
    Dim mTimer As Timer
    Set mTimer = PtrObj(lpTimer)
    
    If Not mTimer.Enabled And Not mTimer.Executed Then
        KillTimer hwnd, lpTimer
    ElseIf (Not mTimer.Enabled) Or mTimer.Executed Then
        KillTimer hwnd, lpTimer
        mTimer.Disable
        mTimer.TickTimer
    ElseIf mTimer.Enabled Then
        mTimer.TickTimer
    End If
    
    Set mTimer = Nothing
    Exit Sub
notloaded:
    Err.Clear
    KillTimer hwnd, lpTimer
End Sub
