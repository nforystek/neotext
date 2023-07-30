Attribute VB_Name = "modWinProc"


#Const modWinProc = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public UC As Collection

Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Sub UnSetControlHost(ByVal lPtr As Long, ByVal hwnd As Long)
    UC.Remove "H" & hwnd
    If UC.Count = 0 Then Set UC = Nothing
End Sub

Public Sub SetControlHost(ByVal lPtr As Long, ByVal hwnd As Long)
    If UC Is Nothing Then
        Set UC = New Collection
    End If
    UC.Add lPtr, "H" & hwnd
End Sub

Private Function GetObject(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4&
    Set GetObject = NewObj
    RtlMoveMemory NewObj, lZero, 4&
End Function


Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uMsg
        Case 953
            Dim objPlayer As Player
            Set objPlayer = GetObject(UC("H" & hwnd))
            objPlayer.NotifySound
            Set objPlayer = Nothing
    End Select

End Function



