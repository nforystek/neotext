VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ambient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private hWndVal As Long

Private pFileName As String
Private pLoopEnabled As Boolean
Private pLoopTimes As Long
Private pTrackVolume As Long
Private pIdentity As String
Private pEnable As Boolean

Private pOrigin As D3DVECTOR
Private pRange As Single

Private pCount As Long

Public Property Get Origin() As D3DVECTOR
    Origin = pOrigin
End Property
Public Property Let Origin(ByRef NewVal As D3DVECTOR)
    pOrigin = NewVal
End Property

Public Property Get Range() As Single
    Range = pRange
End Property
Public Property Let Range(ByVal NewVal As Single)
    pRange = NewVal
End Property

Public Property Get Identity() As String
    Identity = pIdentity
End Property
Public Property Let Identity(ByVal NewVal As String)
    pIdentity = NewVal
End Property

Public Property Get Enable() As Boolean
    Enable = pEnable
End Property
Public Property Let Enable(ByVal NewVal As Boolean)
    pEnable = NewVal
End Property

Public Property Get hWnd() As Long
    hWnd = hWndVal
End Property

Public Property Get Filename() As String
    Filename = pFileName
End Property
Public Property Let Filename(ByVal NewVal As String)
    If Filename <> "" Then MciCommand "close maxsound" & hWndVal
    If PathExists(NewVal, True) Then
        pFileName = NewVal
        MciCommand "open """ & pFileName & """ alias maxsound" & hWndVal
    End If
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
    If PathExists(pFileName, True) Then
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
    End If
End Property

Private Function MciCommand(sCommand As String) As String
    Dim S As String * 255
    mciSendString sCommand, S, 255, hWndVal
    MciCommand = Replace(S, Chr(0), "")
End Function
Public Sub StopSound()
    If (hWndVal > 0) Then
        MciCommand "stop maxsound" & hWndVal
    End If
    pCount = 0
End Sub
Public Sub PlaySound()
    If Not PathExists(pFileName, True) Then
        Err.Raise 53, App.EXEName, "File not found"
    Else
        pCount = 0
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
        MciCommand "play maxsound" & hWndVal & " from 1 notify"
    End If
End Sub

Public Sub FadeIn()
    Static FadeInStart As Single
    If pEnable And (Not (pFileName = "")) Then
        If (FadeInStart = 0) Then
            FadeInStart = Timer
        End If
        If (pTrackVolume >= 0) And (pTrackVolume < 1000) Then
            pTrackVolume = pTrackVolume + 1
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
        Else
            FadeInStart = 0
            pTrackVolume = 1000
        End If
    ElseIf (Not (pFileName = "")) Then
        StopSound
    End If
End Sub

Public Sub FadeOut()
    Static FadeOutStart As Single
    If pEnable And (Not (pFileName = "")) Then
        If (FadeOutStart = 0) Then
            FadeOutStart = Timer
        End If
        If (pTrackVolume > 0) And (pTrackVolume <= 1000) Then
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pTrackVolume
            pTrackVolume = pTrackVolume - 1
        Else
            FadeOutStart = 0
            pTrackVolume = 0
        End If
    ElseIf (Not (pFileName = "")) Then
        StopSound
    End If
End Sub

Friend Sub NotifySound()
    If pLoopEnabled And pEnable And (Not (pFileName = "")) Then
        If (pLoopTimes > 0) Then pCount = pCount + 1
        If (pCount < pLoopTimes) Or (pLoopTimes = 0) Then
            MciCommand "play maxsound" & hWndVal & " from 0 notify"
        Else
            StopSound
        End If
    ElseIf (Not (pFileName = "")) Then
        StopSound
    End If
End Sub

Private Sub InitializeWindow()
    hWndVal = WindowInitialize(AddressOf WndProc)
    SetControlHost ObjPtr(Me), hWndVal
End Sub

Private Sub TerminateWindow()
    If Filename <> "" Then MciCommand "close maxsound" & hWndVal
    UnSetControlHost ObjPtr(Me), hWndVal
    WindowTerminate hWndVal
    hWndVal = 0
End Sub


Private Sub Class_Initialize()
    Stats_Ambient_Count = Stats_Ambient_Count + 1
    pTrackVolume = 1000
    pRange = -1
    InitializeWindow
End Sub

Private Sub Class_Terminate()
    StopSound
    TerminateWindow
    Stats_Ambient_Count = Stats_Ambient_Count - 1
End Sub
