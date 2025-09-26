Attribute VB_Name = "modTrack"
#Const modTrack = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Const GWL_WNDPROC = -4
Public Const WS_DISABLED = &H8000000

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)
Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'###########################################################################
'###################### BEGIN UNIQUE NON GLOBALS ###########################
'###########################################################################

Public UC As NTNodes10.Collection


Public Property Get ObjectByPtr(ByVal lPtr As Long) As Object
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4
    Set ObjectByPtr = NewObj
    DestroyObject NewObj
End Property

Public Sub DestroyObject(ByRef Obj As Object)
    RtlMoveMemory Obj, 0&, 4
End Sub

Public Sub UnSetControlHost(ByVal lPtr As Long, ByVal hwnd As Long)
    If Not UC Is Nothing Then
        UC.Remove "H" & hwnd
        If UC.Count = 0 Then Set UC = Nothing
    End If
End Sub

Public Sub SetControlHost(ByVal lPtr As Long, ByVal hwnd As Long)
    If UC Is Nothing Then
        Set UC = New NTNodes10.Collection
    End If
    If Not UC.Exists("H" & hwnd) Then
        UC.Add lPtr, "H" & hwnd
    End If
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo clearit

    Select Case uMsg
        Case 953
            Dim objPlayer As Track
            Set objPlayer = ObjectByPtr(Trim(Str(UC("H" & hwnd))))
            objPlayer.NotifySound
            Set objPlayer = Nothing
    End Select

    WndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    Exit Function
clearit:
    WindowTerminate hwnd
End Function

Public Function WindowInitialize(ByVal lpWndProc As Long) As Long

    If InIDE Then
        WindowInitialize = frmMain.hwnd
    Else
        Dim hwnd As Long
        hwnd = CreateWindowEx(ByVal 0&, "Message", "", WS_DISABLED, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, App.hInstance, ByVal 0&)
        
        SetWindowLong hwnd, GWL_WNDPROC, lpWndProc
        
        WindowInitialize = hwnd
    End If
End Function

Public Sub WindowTerminate(ByVal hwnd As Long)
        
    If Not InIDE Then
        DestroyWindow hwnd
    End If

End Sub
