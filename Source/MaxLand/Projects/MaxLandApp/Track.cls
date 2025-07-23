VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Track"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private pOrigin As Point
Private pEnabled As Boolean

Private pRange As Single

Private hWndVal As Long

Private pFileName As String
Private pLoopEnabled As Boolean
Private pLoopTimes As Long

Private pVolume As Long

Private pKey As String
Private pSerialize As Boolean

Private pCount As Long

Public Property Get Count() As Long
Count = pCount
End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Public Property Get Loops() As Long
    Loops = pLoopTimes
End Property
Public Property Let Loops(ByVal RHS As Long)
    pLoopTimes = RHS
    pLoopEnabled = True
End Property

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal NewVal As Boolean)
    pEnabled = NewVal
End Property

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef NewVal As Point)
    Set pOrigin = NewVal
End Property

Public Property Get Range() As Single
    Range = pRange
End Property
Public Property Let Range(ByVal NewVal As Single)
    pRange = NewVal
End Property





Public Property Get hwnd() As Long
    hwnd = hWndVal
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal NewVal As String)
    If PathExists(AppPath & "Sounds\" & NewVal, True) Then
        If FileName <> "" Then MciCommand "close maxsound" & hWndVal
        pFileName = NewVal
        MciCommand "open """ & AppPath & "Sounds\" & NewVal & """ alias maxsound" & hWndVal
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
Public Property Get Volume() As Long
    Volume = pVolume
End Property
Public Property Let Volume(ByVal NewVal As Long)
    If (NewVal >= 0) And (NewVal <= 1000) Then
        pVolume = NewVal
        If PathExists(pFileName, True) Then
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pVolume
        End If
    End If
End Property

Private Function MciCommand(sCommand As String) As String
    Dim S As String * 255
    mciSendString sCommand, S, 255, hWndVal
    MciCommand = Replace(S, Chr(0), "")
End Function
Public Sub StopTrack()
    If (hWndVal > 0) Then
        MciCommand "stop maxsound" & hWndVal
    End If
    pCount = 0
End Sub
Public Sub PlayTrack()
    If Not PathExists(AppPath & "Sounds\" & pFileName, True) Then
        Err.Raise 53, App.EXEName, "File not found"
    Else
        pCount = 0
        MciCommand "setaudio maxsound" & hWndVal & " volume to " & pVolume
        MciCommand "play maxsound" & hWndVal & " from 1 notify"
    End If
End Sub

Public Sub FadeIn()
    Static FadeInStart As Single
    If pEnabled And (Not (pFileName = "")) Then
        If (FadeInStart = 0) Then
            FadeInStart = Timer
        End If
        If (pVolume >= 0) And (pVolume < 1000) Then
            pVolume = pVolume + 1
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pVolume
        Else
            FadeInStart = 0
            pVolume = 1000
        End If
    ElseIf (Not (pFileName = "")) Then
        StopTrack
    End If
End Sub

Public Sub FadeOut()
    Static FadeOutStart As Single
    If pEnabled And (Not (pFileName = "")) Then
        If (FadeOutStart = 0) Then
            FadeOutStart = Timer
        End If
        If (pVolume > 0) And (pVolume <= 1000) Then
            MciCommand "setaudio maxsound" & hWndVal & " volume to " & pVolume
            pVolume = pVolume - 1
        Else
            FadeOutStart = 0
            pVolume = 0
        End If
    ElseIf (Not (pFileName = "")) Then
        StopTrack
    End If
End Sub

Friend Sub NotifySound()
    If pLoopEnabled And pEnabled And (Not (pFileName = "")) Then
        If (pLoopTimes > 0) Then pCount = pCount + 1
        If (pCount < pLoopTimes) Or (pLoopTimes = 0) Then
            MciCommand "play maxsound" & hWndVal & " from 0 notify"
        Else
            StopTrack
        End If
    ElseIf (Not (pFileName = "")) Then
        StopTrack
    End If
End Sub

Private Sub InitializeWindow()
    hWndVal = WindowInitialize(AddressOf WndProc)
    SetControlHost ObjPtr(Me), hWndVal
End Sub

Private Sub TerminateWindow()
    If FileName <> "" Then MciCommand "close maxsound" & hWndVal
    UnSetControlHost ObjPtr(Me), hWndVal
    WindowTerminate hWndVal
    hWndVal = 0
End Sub


Private Sub Class_Initialize()
    Stats_Track_Count = Stats_Track_Count + 1
    Set pOrigin = New Point
    pVolume = 1000
    pRange = -1
    Key = "K" & ObjPtr(Me)
    InitializeWindow
End Sub

Private Sub Class_Terminate()
    StopTrack
    TerminateWindow
    Set pOrigin = Nothing
    Stats_Track_Count = Stats_Track_Count - 1
End Sub
