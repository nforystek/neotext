Attribute VB_Name = "modTrack"
#Const modTrack = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public DisableSound As Boolean

Public Track1 As clsAmbient
Public Track2 As clsAmbient

Public UC As Collection

Public Const GWL_WNDPROC = -4
Public Const WS_DISABLED = &H8000000

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub CreateSounds()
    If Not DisableSound Then
    
        Set Track1 = New clsAmbient
        Track1.FileName = AppPath & "Music\start.mp3"
        Track1.LoopEnabled = True
        Track1.TrackVolume = 1000
        
        Set Track2 = New clsAmbient
        Track2.FileName = AppPath & "Music\music.mp3"
        Track2.LoopEnabled = True
        Track2.TrackVolume = 0
    End If
End Sub

Public Sub CleanupSounds()
    If Not DisableSound Then
        Set Track2 = Nothing
        Set Track1 = Nothing
    End If
End Sub

Public Sub RenderAudio()
    If Not DisableSound Then
        Dim o As Long
        If (MenuMode = 0) Then
            For o = 1 To UBound(Level.Starts)
                If (Player.Location.X > LeastX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.X < GreatestX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.Z > LeastZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.Z < GreatestZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4)) Then
                            
                    Track2.FadeOut
                    Track1.FadeIn
                Else
                    Track1.FadeOut
                    Track2.FadeIn
                End If
            Next
        End If
    End If
End Sub

Public Property Get GetPlayerByPtr(ByVal lPtr As Long) As Object
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4
    Set GetPlayerByPtr = NewObj
    DestroyObject NewObj
End Property

Public Sub DestroyObject(ByRef obj As Object)
    RtlMoveMemory obj, 0&, 4
End Sub

Public Sub UnSetControlHost(ByVal lPtr As Long, ByVal hWnd As Long)
    UC.Remove "H" & hWnd
    If UC.Count = 0 Then Set UC = Nothing
End Sub

Public Sub SetControlHost(ByVal lPtr As Long, ByVal hWnd As Long)
    If UC Is Nothing Then
        Set UC = New Collection
    End If
    UC.Add lPtr, "H" & hWnd
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
        Case 953
            Dim objPlayer As clsAmbient
            Set objPlayer = GetPlayerByPtr(Trim(Str(UC("H" & hWnd))))
            objPlayer.NotifySound
            Set objPlayer = Nothing
    End Select

    WndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)

End Function

Public Function WindowInitialize(ByVal lpWndProc As Long) As Long

    Dim hWnd As Long
    hWnd = CreateWindowEx(ByVal 0&, "Message", "", WS_DISABLED, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, App.hInstance, ByVal 0&)
    
    SetWindowLong hWnd, GWL_WNDPROC, lpWndProc
    
    WindowInitialize = hWnd
    
End Function

Public Sub WindowTerminate(ByVal hWnd As Long)
        
    DestroyWindow hWnd

End Sub
