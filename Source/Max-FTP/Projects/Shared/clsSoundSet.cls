VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Private WithEvents timTick As NTSchedule20.Timer
Attribute timTick.VB_VarHelpID = -1
Private WithEvents sndPlay As NTSound20.Player
Attribute sndPlay.VB_VarHelpID = -1
Private WithEvents sndStop As NTSound20.Player
Attribute sndStop.VB_VarHelpID = -1
Private WithEvents sndTick As NTSound20.Player
Attribute sndTick.VB_VarHelpID = -1
Private WithEvents sndHalt As NTSound20.Player
Attribute sndHalt.VB_VarHelpID = -1

Private nCount As Long
Private tCount As Long
Private hIsOn As Boolean
Private SleepTime As Long

Public Sub SoundPlay()
    sndPlay.PlaySound
End Sub

Public Sub SoundStop()
    sndStop.PlaySound
End Sub

Public Sub SoundTick(ByVal Count As Long)
    tCount = Count
    If (tCount > 0) And (Not timTick.Enabled) Then timTick.Enabled = True
    If (tCount = 0) And timTick.Enabled Then timTick.Enabled = False
End Sub

Public Sub SoundHalt(ByVal IsOn As Boolean)
    If IsOn And Not sndHalt.IsPlaying Then
        sndHalt.PlaySound
    End If
    hIsOn = IsOn
End Sub

Private Sub Class_Initialize()
    
    Set sndPlay = New NTSound20.Player
    Set sndStop = New NTSound20.Player
    Set sndTick = New NTSound20.Player
    Set sndHalt = New NTSound20.Player
    Set timTick = New NTSchedule20.Timer
    
    timTick.Interval = 100
    
    sndPlay.FileName = AppPath & "Sounds\play.wav"
    sndStop.FileName = AppPath & "Sounds\stop.wav"
    sndTick.FileName = AppPath & "Sounds\tick.wav"
    sndHalt.FileName = AppPath & "Sounds\halt.wav"
    
End Sub

Private Sub Class_Terminate()

    sndPlay.StopSound
    sndStop.StopSound
    sndTick.StopSound
    sndHalt.StopSound

    timTick.Enabled = False
    
    Set timTick = Nothing
    Set sndPlay = Nothing
    Set sndStop = Nothing
    Set sndTick = Nothing
    Set sndHalt = Nothing
    
End Sub

Private Sub sndHalt_SoundNotify()
    If hIsOn Then sndHalt.PlaySound
End Sub

Private Sub sndTick_SoundNotify()
    nCount = nCount - 1
    If (nCount = 0) Then
        SleepTime = Timer
    ElseIf (Not hIsOn) Then
        sndTick.PlaySound
    End If
End Sub

Private Sub timTick_OnTicking()
    If ((nCount = 0) And ((Timer - SleepTime) >= 2)) Or (SleepTime = 0) Then
        nCount = tCount
        If (tCount > 0) Then sndTick.PlaySound
    End If
End Sub
