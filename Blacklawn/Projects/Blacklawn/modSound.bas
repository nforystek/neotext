#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modSound"
#Const modSound = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public DisableSound As Boolean

Public Const SOUND_BANG = 0
Public Const SOUND_BOOM = 1
Public Const SOUND_JUICE = 2
Public Const SOUND_LAUNCH = 3
Public Const SOUND_ROLLOUT = 4
Public Const SOUND_SMACK = 5
Public Const SOUND_TOGGLE = 6
Public Const SOUND_NICKEL = 7

Private Waves() As DirectSoundSecondaryBuffer8

Public Sub CreateSounds()
    If Not DisableSound Then
        ReDim Waves(0 To 7) As DirectSoundSecondaryBuffer8
        LoadSound SOUND_BANG, AppPath & "Base\Sound\bang.wav"
        LoadSound SOUND_BOOM, AppPath & "Base\Sound\boom.wav"
        LoadSound SOUND_JUICE, AppPath & "Base\Sound\juice.wav"
        LoadSound SOUND_LAUNCH, AppPath & "Base\Sound\launch.wav"
        LoadSound SOUND_ROLLOUT, AppPath & "Base\Sound\rollout.wav"
        LoadSound SOUND_SMACK, AppPath & "Base\Sound\smack.wav"
        LoadSound SOUND_TOGGLE, AppPath & "Base\Sound\toggle.wav"
        LoadSound SOUND_NICKEL, AppPath & "Base\Sound\nickel.wav"
    End If
End Sub

Public Sub CleanupSounds()
    If Not DisableSound Then

        StopWave SOUND_BANG
        StopWave SOUND_BOOM
        StopWave SOUND_JUICE
        StopWave SOUND_LAUNCH
        StopWave SOUND_ROLLOUT
        StopWave SOUND_SMACK
        StopWave SOUND_TOGGLE
        StopWave SOUND_NICKEL
        Erase Waves
    End If
End Sub

Public Sub PlayWave(ByVal Index As Long, Optional ByVal Repeat As Boolean = False)
    If Not DisableSound Then

        If PlaySound Then
            Waves(Index).Play IIf(Repeat, DSBPLAY_LOOPING, DSBPLAY_DEFAULT)
        End If
    End If
End Sub

Public Sub StopWave(ByVal Index As Long)
    If Not DisableSound Then

        Waves(Index).Stop
        Waves(Index).SetCurrentPosition 0
    End If
End Sub

Private Sub LoadSound(ByVal Index As Long, ByVal FileName As String)
    If Not DisableSound Then

        Dim bufferDesc As DSBUFFERDESC
        Dim waveFormat As WAVEFORMATEX
          
        bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        
        waveFormat.nFormatTag = WAVE_FORMAT_PCM
        waveFormat.nChannels = 2
        waveFormat.lSamplesPerSec = 22050
        waveFormat.nBitsPerSample = 16
        waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
        waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
        
        Set Waves(Index) = DSound.CreateSoundBufferFromFile(Trim(FileName), bufferDesc)
    End If
End Sub

Public Sub RenderAudio()
    If Not DisableSound Then

        Dim dist As Single
        dist = Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0)
        If (dist > (WithInCityLimits * 2)) And (Track1.TrackVolume > 0) Then
            Track1.FadeOut
            Track2.FadeIn
        ElseIf (dist <= (WithInCityLimits * 2)) And (Track2.TrackVolume > 0) Then
            Track2.FadeOut
            Track1.FadeIn
        End If
    End If
End Sub