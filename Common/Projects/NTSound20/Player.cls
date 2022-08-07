VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = True
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

Private pIsPlaying As Boolean
Private pCount As Long

Private Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Boolean

Event SoundNotify()

Public Function Beep(ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Boolean
    Beep = Beeper(dwFreq, dwDuration)
End Function

Friend Property Get hwnd() As Long
    hwnd = hWndVal
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal newVal As String)
    pFileName = newVal
End Property
Public Property Get LoopEnabled() As Boolean
    LoopEnabled = pLoopEnabled
End Property
Public Property Let LoopEnabled(ByVal newVal As Boolean)
    pLoopEnabled = newVal
End Property
Public Property Get LoopTimes() As Long
    LoopTimes = pLoopTimes
End Property
Public Property Let LoopTimes(ByVal newVal As Long)
    pLoopTimes = newVal
End Property

Public Property Get IsPlaying() As Boolean
    IsPlaying = pIsPlaying
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
    pIsPlaying = False
    pCount = 0
End Sub
Public Sub PlaySound()
    If Not PathExists(pFileName, True) Then
        Err.Raise 53, App.EXEName, "File not found"
    Else
        pCount = 0
        pIsPlaying = True
        MciCommand "open """ & pFileName & """ alias maxsound" & hWndVal
        MciCommand "play maxsound" & hWndVal & " from 1 notify"
    End If
End Sub

Friend Sub NotifySound()
    If pLoopEnabled Then
        If (pLoopTimes > 0) Then pCount = pCount + 1
        If (pCount < pLoopTimes) Or (pLoopTimes = 0) Then
            MciCommand "play maxsound" & hWndVal & " from 1 notify"
        Else
            StopSound
        End If
    Else
        StopSound
    End If
    RaiseEvent SoundNotify
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
    InitializeWindow
End Sub

Private Sub Class_Terminate()
    StopSound
    TerminateWindow
End Sub
