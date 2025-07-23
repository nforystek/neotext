Attribute VB_Name = "modSound"
#Const modSound = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public DisableSound As Boolean

Public Waves() As DirectSoundSecondaryBuffer8

Public Sub PlayWave(ByVal Index As Long, Optional ByVal Repeat As Boolean = False)
    If (Not DisableSound) Then
        If SoundFX Then
            Waves(Index).Play IIf(Repeat, DSBPLAY_LOOPING, DSBPLAY_DEFAULT)
        Else
            StopWave Index
        End If
    End If
End Sub
Public Sub VolumeWave(ByVal Index As Long, ByVal Dist As Single)
    If (Not DisableSound) Then
        
        Dim div As Single
        Dim r As Single
        
        r = Round(CSng(Sounds(Index).Range - Dist), 3)
        r = Abs(-Sounds(Index).Range + r)
        r = -((Abs(DSBVOLUME_MAX - DSBVOLUME_MIN) / Sounds(Index).Range) * r)

        Waves(Index).SetVolume r
        
    ElseIf (Not DisableSound) Then
        StopWave Index
    End If
End Sub
Public Sub StopWave(ByVal Index As Long)
    If (Not DisableSound) Then
        Waves(Index).Stop
        Waves(Index).SetCurrentPosition 0
    End If
End Sub

Public Sub LoadWave(ByVal Index As Long, ByVal FileName As String)
    If (Not DisableSound) Then

        Dim bufferDesc As DSBUFFERDESC
        Dim waveFormat As WAVEFORMATEX
          
        ReDim Preserve Waves(1 To Index) As DirectSoundSecondaryBuffer8
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


