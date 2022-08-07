
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAmbient"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private hWndVal As Long

Private pFileName As String
Private pLoopEnabled As Boolean
Private pLoopTimes As Long
Private pTrackVolume As Long

Private pCount As Long

Public Property Get hwnd() As Long
    hwnd = hWndVal
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal NewVal As String)
    pFileName = NewVal
End Property
Public Property Get LoopEnabled() As Boolean
    LoopEnabled = pLoopEnabled
End Property
Public Property Let LoopEnabled(ByVal NewVal As Boolean)
    pLoopEnabled = NewVal
End Property
Public Property Get LoopTimes() As Long
    LoopTimes = pLoopTimes
End Property
Public Property Let LoopTimes(ByVal NewVal As Long)
    pLoopTimes = NewVal
End Property
Public Property Get TrackVolume() As Long
    TrackVolume = pTrackVolume
End Property
Public Property Let TrackVolume(ByVal NewVal As Long)
    pTrackVolume = NewVal
End Property

Private Function MciCommand(sCommand As String) As String
    Dim s As String * 255
    mciSendString sCommand, s, 255, hWndVal
    MciCommand = Replace(s, Chr(0), "")
End Function
Public Sub StopSound()
    If hWndVal > 0 Then
        MciCommand "stop maxsound" & hWndVal
        MciCommand "close maxsound" & hWndVal
    End If
    pCount = 0
End Sub
Public Sub PlaySound()
    If Not PathExists(pFileName, True) Then
        AddMessage "File not found: " & pFileName
    Else
        pCount = 0
        MciCommand "open """ & pFileName & """ alias maxsound" & hWndVal
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
        MciCommand "play maxsound" & hWndVal & " from 1 notify"
    End If
End Sub

Public Sub FadeIn()
    Static FadeInStart As Single
    If (FadeInStart = 0) Then
        FadeInStart = Timer
    End If
    If pTrackVolume >= 0 And pTrackVolume < 1000 Then
        pTrackVolume = pTrackVolume + 1
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
    Else
        FadeInStart = 0
        pTrackVolume = 1000
    End If
End Sub

Public Sub FadeOut()
    Static FadeOutStart As Single
    If (FadeOutStart = 0) Then
        FadeOutStart = Timer
    End If
    If pTrackVolume > 0 And pTrackVolume <= 1000 Then
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
        pTrackVolume = pTrackVolume - 1
    Else
        FadeOutStart = 0
        pTrackVolume = 0
    End If
End Sub

Friend Sub NotifySound()
    If pLoopEnabled Then
        If (pLoopTimes > 0) Then pCount = pCount + 1
        If (pCount < pLoopTimes) Or (pLoopTimes = 0) Then
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
            MciCommand "play maxsound" & hWndVal & " from 1 notify"
        Else
            StopSound
        End If
    Else
        StopSound
    End If
End Sub

Private Sub InitializeWindow()
    hWndVal = WindowInitialize(AddressOf WndProc)
    SetControlHost ObjPtr(Me), hWndVal
End Sub

Private Sub TerminateWindow()
    UnSetControlHost ObjPtr(Me), hWndVal
    WindowTerminate hWndVal
    hWndVal = 0
End Sub


Private Sub Class_Initialize()
    pTrackVolume = 1000
    InitializeWindow
End Sub

Private Sub Class_Terminate()
    StopSound
    TerminateWindow
End Sub