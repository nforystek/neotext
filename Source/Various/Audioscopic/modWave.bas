Attribute VB_Name = "modWave"
Option Explicit

'LengthOfFile: 706252
'Id: RIFF
'DataType: WAVE
'LengthOfRiff:  706244
'nAvgBytesPerSec:  176400
'nBlockAlign:  4
'nChannels:  2
'nSamplesPerSec:  44100
'wBitsPerSample:  16
'wFormatTag:  1
'LengthOfData:  703248
'LengthOfWave:  703300
'WaveMilliseconds:  3986
'WaveSamples:  175812
'
'?175812 * 4
' 703248
'?176400 * 4
' 705600
'?705600 / 16
' 44100
'
'?703300-703248
' 52
'?706252-706244
' 8
'
'?706252-703248
' 3004
'?706252-703300
' 2952
'?706244-703248
' 2996
'?706244-703300
' 2944
'
'?3004-2952
' 52
'?3004-2996
' 8
'?2996-2944
' 52
'?2952-2944
' 8
'
'
'LengthOfFile: 706252
'Id: RIFF
'DataType: WAVE
'LengthOfRiff:  706244
'nAvgBytesPerSec:  176400
'nBlockAlign:  4
'nChannels:  2
'nSamplesPerSec:  44100
'wBitsPerSample:  16
'wFormatTag:  1
'LengthOfData:  703248
'LengthOfWave:  703300
'WaveMilliseconds:  3986
'WaveSamples:  175812
'
'?WaveSample * 4 = LengthOfData
'
'?nAvgBytesPerSec * 4 = BitsPerSecond
'?BitsPerSecond / 16 =  nSamplesPerSec
'
'?LengthOfWave-LengthOfData = SizeOfHeaders
'?LengthOfFile-LengthOfRiff = SizeOfFileName
'
'?LengthOfFile-LengthOfData = chksum1
'?LengthOfFile-LengthOfWave = chksum2
'?LengthOfRiff-LengthOfData = chksum3
'?LengthOfRiff-LengthOfWave = chksum4
'
'?chksum1-chksum2 = SizeOfHeaders
'?chksum1-chksum3 = SizeOfFileName
'
'?chksum3-chksum4 = SizeOfHeaders
'?chksum2-chksum4 = SizeOfFileName

Type RiffHdr_
    ID          As String * 4       'identifier string = "RIFF"
    Len         As Long             'remaining length *after* this header
    DataType    As String * 4       'type of data. wav = WAVE
End Type


Type ChunkHdr_                      'CHUNK 8-byte header
    ID          As String * 4       'identifier, e.g. "fmt " or "data"
    Len         As Long             'remaining chunk length *after* header
End Type                            'data bytes follow chunk header



Type WAVEFORMATEX                   'FMT Chunk
    wFormatTag      As Integer      'Format category            '    wFormatSpecific As Integer
    nChannels       As Integer      'Number of channels         '    wNumberOfChannels As Integer
    nSamplesPerSec  As Long         'Sampling rate              '    lSamplesPerSecond As Long
    nAvgBytesPerSec As Long         'For buffer estimation      '    lBytesPerSecond As Long
    nBlockAlign     As Integer      'Data block size            '    wChannelBandwidth As Integer
    wBitsPerSample  As Integer                                  '    wBitsPerSample As Integer
End Type

Type FactChunk_                     'Not always present
    dwFileSize As Long              'Number Of Samples
End Type

Type WaveData8bit_
    ChannelData() As Byte           '8bit samples, (channel)(samples)
End Type

Type WaveData16bit_
    ChannelData() As Integer        '16bit samples, (channel)(samples)
End Type



'Dim RiffHdr As RiffHdr_
'Dim FMTChunk As WAVEFORMATEX
'Dim FACTChunk As FactChunk_
Dim ChunkHdr As ChunkHdr_
'Dim WaveData8bit As WaveData8bit_
'Dim WaveData16bit As WaveData16bit_

Public Type AudioFile2

    RiffHdr As RiffHdr_
    FMTChunk As WAVEFORMATEX
    FACTChunk As FactChunk_

    WaveData8bit As WaveData8bit_
    WaveData16bit As WaveData16bit_

End Type
'
'-
'Number of Samples Per Channel is Calculated By:
'(The Length Of The DataChunk devide by the number of channels)
'devide by the number of bytes per sample
'-\


Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

' API Declarations
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Dest As Long, ByVal Source As Long, ByVal Length As Long)

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Public Declare Function PlaySound Lib "WINMM.DLL" Alias "PlaySoundA" (ByRef Sound As Any, ByVal hLib As Long, ByVal lngFlag As Long) As Long      'BOOL

Public PlayingBytes() As Byte
Public SilentBytes() As Byte

Public Declare Function PlaySound2 Lib "winmm.dll" Alias "PlaySoundA" (ByVal Sound As Long, ByVal hLib As Long, ByVal lngFlag As Long) As Long      'BOOL

Public Enum sndConst
    SND_ASYNC = &H1 ' play asynchronously
    SND_LOOP = &H8 ' loop the sound until Next sndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points To a memory file
    SND_NODEFAULT = &H2 ' silence Not default, If sound not found
    SND_NOSTOP = &H10 ' don't stop any currently playing sound
    SND_SYNC = &H0 ' play synchronously (default), halts prog use till done playing
End Enum

Public Function FindTag(ByVal Tag As String, ByRef pStream As NTNodes10.Stream) As Long
    Dim pos As Long
    Dim bal As Long
    pos = -1
    bal = 1
    Do
        pos = pStream.Poll(Asc(Left(Tag, 1)), bal, pos + 1)
        If pos < pStream.Length Then
            If pStream.Poll(Asc(Mid(Tag, 2, 1)), 1, pos + 1) = 0 Then
                If pStream.Poll(Asc(Mid(Tag, 3, 1)), 1, pos + 2) = 0 Then
                    If Not pStream.Poll(Asc(Mid(Tag, 4, 1)), 1, pos + 3) = 0 Then
                        pos = -1
                    End If
                Else
                    pos = -1
                End If
            Else
                pos = -1
            End If
        End If
        bal = bal + 1
    Loop Until pos > -2
    If pos < 0 Then
        FindTag = -1
    Else
        FindTag = pos
    End If
End Function
Public Function ArrayBoundSize(InArray, Optional ByVal Dimension As Integer = 1, Optional ByVal inBytes As Boolean = False) As Single
    On Error Resume Next
    Dim factor As Long
    If inBytes Then
        factor = ArrayElementSize(InArray)
    Else
        factor = 1
    End If
    ArrayBoundSize = CSng((UBound(InArray, Dimension) + -LBound(InArray, Dimension)) + 1) * factor
    If Err Then Err.Clear
    On Error GoTo 0
End Function
Public Function ArrayDimensions(InArray) As Single
    On Error Resume Next
    Dim e As Long
    Dim d As Long
    Do
        If Err.Number = 0 Then d = d + 1
        e = LBound(InArray, d)
    Loop Until Err.Number <> 0
    If Err Then
        d = d - 1
        Err.Clear
    End If
    ArrayDimensions = d
    On Error GoTo 0
End Function

Public Function ArrayTotalCount(InArray, Optional ByVal inBytes As Boolean = False) As Single
    On Error GoTo fail
    Dim d As Long
    d = ArrayDimensions(InArray)
    Do While d > 0
        ArrayTotalCount = ArrayTotalCount + ArrayBoundSize(InArray, d, inBytes)
        d = d - 1
    Loop
fail:
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function ArrayElementSize(InArray) As Single
    On Error GoTo fail
    Select Case TypeName(InArray)
        Case "Byte()"
            ArrayElementSize = 1
        Case "Boolean()"
            ArrayElementSize = 2
        Case "Integer()"
            ArrayElementSize = 2
        Case "Long()"
            ArrayElementSize = 4
        Case "Single()"
            ArrayElementSize = 4
        Case "Double()"
            ArrayElementSize = 8
        Case "Currency()"
            ArrayElementSize = 8
        Case "Decimal()"
            ArrayElementSize = 14
        Case "Date()"
            ArrayElementSize = 8
    End Select
fail:
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function WaveSamples(ByRef af As AudioFile2) As Long
    If af.FMTChunk.wBitsPerSample = 8 Then
        WaveSamples = ArrayBoundSize(af.WaveData8bit.ChannelData, 2)
    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        WaveSamples = ArrayBoundSize(af.WaveData16bit.ChannelData, 2)
    End If
End Function

Public Function LengthOfData(ByRef af As AudioFile2) As Long

    If af.FMTChunk.wBitsPerSample = 8 Then
        LengthOfData = (UBound(af.WaveData8bit.ChannelData, 2) * af.FMTChunk.nChannels)
    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        LengthOfData = ArrayBoundSize(af.WaveData16bit.ChannelData, 1) * ArrayBoundSize(af.WaveData16bit.ChannelData, 2) * 2
    End If

End Function

Public Function WaveMilliseconds(ByRef af As AudioFile2) As Single
    WaveMilliseconds = Int(((LengthOfData(af) / af.FMTChunk.nBlockAlign) / af.FMTChunk.nSamplesPerSec) * 1000)
End Function

Public Function LengthOfWave(ByRef aFile As AudioFile2, Optional ByVal DurationMS As Long = -1) As Long

    'LengthOfWave = LengthOfWave + LenB(aFile.RiffHdr)
    'LengthOfWave = LengthOfWave + 4
    '-or-
    LengthOfWave = LengthOfWave + 12 'RIFF HEADER
    
    
    If aFile.FACTChunk.dwFileSize <> 0 Then

        'LengthOfWave = LengthOfWave + LenB(aFile.FACTChunk.dwFileSize)
        '-or-
        LengthOfWave = LengthOfWave + 8 'fact header

    End If
    
    'LengthOfWave = LengthOfWave + LenB(aFile.FMTChunk)
    'LengthOfWave = LengthOfWave + 8
    '-or-
    LengthOfWave = LengthOfWave + 32 'format header
    
    If DurationMS = -1 Then
        LengthOfWave = LengthOfWave + LengthOfData(aFile)
    ElseIf DurationMS > 0 Then
        LengthOfWave = LengthOfWave + Round(aFile.FMTChunk.nAvgBytesPerSec * (DurationMS / 1000))
    End If

End Function


'Public Function WaveMilliseconds(ByRef aFile As AudioFile) As Long
'    With aFile.Infos(aFile.wInfoLen)
'        Dim factor As Single
'        Dim remain As Single
'        Dim multiplyer As Single'
'        factor = ((.lSamplesPerSecond * .wNumberOfChannels * .wBitsPerSample) \ (.wNumberOfChannels * .wChannelBandwidth))
'        multiplyer = (aFile.Datas(aFile.wDataLen).lBytes \ factor)
'        remain = (aFile.Datas(aFile.wDataLen).lBytes Mod factor)'
'        WaveMilliseconds = ((multiplyer + (((factor / (factor - remain)) / 100))) * 1000)
'    End With
'End Function

'Public Function WaveMilliseconds(ByRef aFile As AudioFile) As Single
'    With aFile.Infos(aFile.wInfoLen)
'        Dim factor As Single
'        Dim remain As Single
'        Dim multiplyer As Single
'        factor = ((.lSamplesPerSecond * .wNumberOfChannels * .wBitsPerSample) \ (.wNumberOfChannels * .wChannelBandwidth))
'        multiplyer = (aFile.Datas(aFile.wDataLen).lBytes \ factor)
'        remain = (aFile.Datas(aFile.wDataLen).lBytes Mod factor)
'        WaveMilliseconds = ((multiplyer + (((factor / (factor - remain)) / 100))) * 1000)
'        If temp - WaveMilliseconds >= 0.0005 Then
'            WaveMilliseconds = WaveMilliseconds + 0.001
'        End If
'        WaveMilliseconds = WaveMilliseconds * 1000
'    End With
'End Function




Public Sub PlayWaveSound(ByRef useBytes() As Byte)
    PlaySound2 ByVal VarPtr(useBytes(LBound(useBytes))), vbNull, SND_MEMORY Or SND_NODEFAULT Or SND_ASYNC
End Sub


Public Function WaveInfosEqual(ByRef Info1 As AudioFile2, ByRef Info2 As AudioFile2) As Boolean
    WaveInfosEqual = (Info1.FMTChunk.nAvgBytesPerSec = Info2.FMTChunk.nAvgBytesPerSec) And _
                         (Info1.FMTChunk.nSamplesPerSec = Info2.FMTChunk.nSamplesPerSec) And _
                          (Info1.FMTChunk.wBitsPerSample = Info2.FMTChunk.wBitsPerSample) And _
                           (Info1.FMTChunk.nBlockAlign = Info2.FMTChunk.nBlockAlign) And _
                            (Info1.FMTChunk.wFormatTag = Info2.FMTChunk.wFormatTag) And _
                             (Info1.FMTChunk.nChannels = Info2.FMTChunk.nChannels)
End Function


Public Sub AddDesc(ByRef audio() As Byte, ByVal Desc As String, Optional ByVal StartIndex As Long = -1)
    If StartIndex = -1 Then
        ReDim Preserve audio(LBound(audio) To UBound(audio) + 4) As Byte
        StartIndex = UBound(audio) - 3
    End If
    audio(StartIndex + 0) = Asc(Mid(Desc, 1, 1))
    audio(StartIndex + 1) = Asc(Mid(Desc, 2, 1))
    audio(StartIndex + 2) = Asc(Mid(Desc, 3, 1))
    audio(StartIndex + 3) = Asc(Mid(Desc, 4, 1))
End Sub

Public Sub AddLong(ByRef audio() As Byte, ByVal lValue As Long)
    ReDim Preserve audio(LBound(audio) To UBound(audio) + 4) As Byte
    RtlMoveMemory VarPtr(audio(UBound(audio) - 3)), ByVal VarPtr(lValue) + 0, 4
End Sub

'Public Function GetHead(ByRef inBytes() As Byte, ByRef Idx As Long) As AudioHead
'    RtlMoveMemory VarPtr(GetHead), ByVal VarPtr(inBytes(Idx)), LenB(GetHead)
'    Idx = Idx + LenB(GetHead)
'End Function
'
'Public Function GetDesc(ByRef header As AudioHead) As String
'    GetDesc = UCase(Chr(header.bPart(0)) & Chr(header.bPart(1)) & Chr(header.bPart(2)) & Chr(header.bPart(3)))
'End Function



'Public Sub WaveFilter(ByRef aFile As AudioFile, ByVal StartTimeMS As Single, ByVal DurationMS As Single, ByRef definedFilters As ScriptControl, ByVal filterFunction As String)
'    If aFile.wDataLen > 0 And aFile.wInfoLen > 0 Then
'        StartTimeMS = Round(aFile.Infos(1).lBytesPerSecond * (StartTimeMS / 1000))
'        DurationMS = Round(aFile.Infos(1).lBytesPerSecond * (DurationMS / 1000))
'
'        Dim i As Long
'        For i = 1 To (DurationMS / 2)
'            aFile.Datas(1).wWave((StartTimeMS / 2) + i) = definedFilters.Eval((Replace(filterFunction, "%", CStr(aFile.Datas(1).wWave((StartTimeMS / 2) + i)))))
'        Next
'    Else
'        Err.Raise 321, , "The file you have selected is not a RIFF nor WAVE file.  Only 16-bits (Standard PCM) wave format is supported."
'    End If
'
'End Sub


'Public Function WaveCombine(ByRef DestWave() As Byte, ByVal DestStartTimeMS As Single, ByRef SourceWave() As Byte, ByVal SourceDurationTimeMS As Single) As Byte()
'
'    Dim aFile1 As AudioFile2
'    Dim aFile2 As AudioFile2
'
'    aFile1 = WaveBytesAsAudio(DestWave)
'    aFile2 = WaveBytesAsAudio(SourceWave)
'
'    If LengthOfData(aFile1) > 0 And _
'        LengthOfData(aFile2) > 0 And aFile1.FMTChunk.wBitsPerSample = aFile2.FMTChunk.wBitsPerSample Then
'
'        DestStartTimeMS = Round(aFile1.FMTChunk.nAvgBytesPerSec * (DestStartTimeMS / 1000))
'        SourceDurationTimeMS = Round(aFile2.FMTChunk.nAvgBytesPerSec * (SourceDurationTimeMS / 1000))
'
'        If aFile1.FMTChunk.wBitsPerSample = 8 Then
'
'            Dim newData8() As Byte
'            ReDim Preserve newData8(0 To aFile1.FMTChunk.nChannels - 1, 0 To (LengthOfData(aFile1) + SourceDurationTimeMS) - 1) As Byte
'
'            If DestStartTimeMS > 0 Then
'                RtlMoveMemory VarPtr(newData8(LBound(newData8, 1), LBound(newData8, 2))), _
'                    ByVal VarPtr(aFile1.WaveData8bit.ChannelData(LBound(aFile1.WaveData8bit.ChannelData, 1), LBound(aFile1.WaveData8bit.ChannelData, 2))), DestStartTimeMS
'            End If
'
'            RtlMoveMemory VarPtr(newData8(LBound(newData8, 1), LBound(newData8, 2) + DestStartTimeMS)), _
'                ByVal VarPtr(aFile2.WaveData8bit.ChannelData(LBound(aFile2.WaveData8bit.ChannelData, 1), LBound(aFile2.WaveData8bit.ChannelData, 2))), SourceDurationTimeMS
'
'            If (LengthOfData(aFile1) - DestStartTimeMS) > 0 Then
'                RtlMoveMemory VarPtr(newData8(LBound(newData8, 1), LBound(newData8, 2) + (DestStartTimeMS + SourceDurationTimeMS))), _
'                    ByVal VarPtr(aFile1.WaveData8bit.ChannelData(LBound(aFile1.WaveData8bit.ChannelData, 1), LBound(aFile1.WaveData8bit.ChannelData, 2) + DestStartTimeMS)), _
'                    (LengthOfData(aFile1) - DestStartTimeMS)
'            End If
'
'            ReDim aFile1.WaveData8bit.ChannelData(LBound(newData8, 1) To UBound(newData8, 1), LBound(newData8, 2) To UBound(newData8, 2))
'            RtlMoveMemory VarPtr(newData8(LBound(newData8, 1), LBound(newData8, 2))), ByVal VarPtr(aFile1.WaveData8bit.ChannelData(LBound(aFile1.WaveData8bit.ChannelData, 1), LBound(aFile1.WaveData8bit.ChannelData, 2))), LengthOfData(aFile1)
'
'
'        ElseIf aFile1.FMTChunk.wBitsPerSample = 16 Then
'
'            Dim newData() As Integer
'            ReDim Preserve newData(0 To aFile1.FMTChunk.nChannels - 1, 0 To ((LengthOfData(aFile1) + SourceDurationTimeMS) \ 2) - 1) As Integer
'
'            If DestStartTimeMS > 0 Then
'                RtlMoveMemory VarPtr(newData(LBound(newData, 1), LBound(newData, 2))), _
'                    ByVal VarPtr(aFile1.WaveData16bit.ChannelData(LBound(aFile1.WaveData16bit.ChannelData, 1), LBound(aFile1.WaveData16bit.ChannelData, 2))), DestStartTimeMS
'            End If
'
'            RtlMoveMemory VarPtr(newData(LBound(newData, 1), LBound(newData, 2) + (DestStartTimeMS / 2))), _
'                ByVal VarPtr(aFile2.WaveData16bit.ChannelData(LBound(aFile2.WaveData16bit.ChannelData, 1), LBound(aFile2.WaveData16bit.ChannelData, 2))), SourceDurationTimeMS
'
'            If (LengthOfData(aFile1) - DestStartTimeMS) > 0 Then
'                RtlMoveMemory VarPtr(newData(LBound(newData, 1), LBound(newData, 2) + ((DestStartTimeMS + SourceDurationTimeMS) \ 2))), _
'                    ByVal VarPtr(aFile1.WaveData16bit.ChannelData(LBound(aFile1.WaveData16bit.ChannelData, 1), LBound(aFile1.WaveData16bit.ChannelData, 2) + (DestStartTimeMS \ 2))), _
'                    (LengthOfData(aFile1) - DestStartTimeMS)
'            End If
'
''            Erase aFile1.WaveData16bit.ChannelData
''            aFile1.WaveData16bit.ChannelData = newData
'            ReDim aFile1.WaveData16bit.ChannelData(LBound(newData, 1) To UBound(newData, 1), LBound(newData, 2) To UBound(newData, 2))
'            RtlMoveMemory VarPtr(newData(LBound(newData, 1), LBound(newData, 2))), ByVal VarPtr(aFile1.WaveData16bit.ChannelData(LBound(aFile1.WaveData16bit.ChannelData, 1), LBound(aFile1.WaveData16bit.ChannelData, 2))), LengthOfData(aFile1)
'
'        End If
'
'        WaveCombine = WaveRecordToBytes(aFile1)
'
'        Erase newData
'        WaveResetRecord aFile1
'        WaveResetRecord aFile2
'    Else
'        Err.Raise 321, , "The file you have selected is not a RIFF nor WAVE file.  Or there is a format mismatch between the wave formats to combine."
'    End If
'End Function





Public Function ReadFileBytes(ByVal FileName As String) As Byte()
    Dim inBytes() As Byte
    Dim ff As Integer
    ff = FreeFile
    Open FileName For Binary Access Read As #ff
        ReDim inBytes(0 To FileLen(FileName) - 1) As Byte
        Get #ff, 1, inBytes
    Close #ff
    ReadFileBytes = inBytes
End Function

Public Sub WriteFileBytes(ByVal FileName As String, ByRef outBytes() As Byte)
    Dim ff As Integer
    ff = FreeFile
    Open FileName For Output As #ff
    Close #ff
    Open FileName For Binary Access Write As #ff
        Put #ff, 1, outBytes
    Close #ff
End Sub

Public Sub WaveResetRecord(ByRef af As AudioFile2)
    With af
        .RiffHdr.ID = ""
        .RiffHdr.DataType = ""
        .RiffHdr.Len = 0
        .FACTChunk.dwFileSize = 0
        .FMTChunk.nAvgBytesPerSec = 0
        .FMTChunk.nBlockAlign = 0
        .FMTChunk.nChannels = 0
        .FMTChunk.nSamplesPerSec = 0
        .FMTChunk.wBitsPerSample = 0
        .FMTChunk.wFormatTag = 0
        Erase .WaveData8bit.ChannelData
        Erase .WaveData16bit.ChannelData
    End With
End Sub

Public Sub DebugWav(ByRef af As AudioFile2)
    With af
        If .RiffHdr.ID <> "RIFF" Or .RiffHdr.DataType <> "WAVE" Then
            Debug.Print "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
            Exit Sub
        End If
        Debug.Print "Id: "; .RiffHdr.ID
        Debug.Print "DataType: "; .RiffHdr.DataType
        Debug.Print "LengthOfRiff: "; .RiffHdr.Len
        If .FACTChunk.dwFileSize <> 0 Then
            Debug.Print "dwFileSize: "; .FACTChunk.dwFileSize
        End If
        Debug.Print "nAvgBytesPerSec: "; .FMTChunk.nAvgBytesPerSec
        Debug.Print "nBlockAlign: "; .FMTChunk.nBlockAlign
        Debug.Print "nChannels: "; .FMTChunk.nChannels
        Debug.Print "nSamplesPerSec: "; .FMTChunk.nSamplesPerSec
        Debug.Print "wBitsPerSample: "; .FMTChunk.wBitsPerSample
        Debug.Print "wFormatTag: "; .FMTChunk.wFormatTag
        Debug.Print "LengthOfData: "; LengthOfData(af)
        Debug.Print "LengthOfWave: "; LengthOfWave(af)
        Debug.Print "WaveMilliseconds: "; WaveMilliseconds(af)
        Debug.Print "WaveSamples: "; WaveSamples(af)
        Debug.Print
    End With
End Sub
Public Function GetDesc(ByRef inBytes() As Byte, ByVal pos As Long) As String
    GetDesc = UCase(Chr(inBytes(pos)) & Chr(inBytes(pos + 1)) & Chr(inBytes(pos + 2)) & Chr(inBytes(pos + 3)))
End Function
Public Function WaveBytesAsAudio(ByRef inBytes() As Byte) As AudioFile2
    Dim pos As Long
    pos = LBound(inBytes)
    Dim Length As Long
    
    Dim af As AudioFile2
    
    If GetDesc(inBytes, pos) <> "RIFF" Or GetDesc(inBytes, pos + 8) <> "WAVE" Then
        MsgBox "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
        Exit Function
    End If


    af.RiffHdr.ID = "RIFF"
    af.RiffHdr.DataType = "WAVE"
    pos = pos + 4
    
    RtlMoveMemory VarPtr(af.RiffHdr.Len), ByVal VarPtr(inBytes(pos)), 4

    pos = pos + 8

    Do Until pos >= UBound(inBytes) - 1

        Select Case GetDesc(inBytes, pos)
            Case "FMT "
                pos = pos + 4
                RtlMoveMemory VarPtr(Length), ByVal VarPtr(inBytes(pos)), 4
                pos = pos + 4
                RtlMoveMemory VarPtr(af.FMTChunk), ByVal VarPtr(inBytes(pos)), Length
                pos = pos + Length

            Case "FACT"
                pos = pos + 4
                RtlMoveMemory VarPtr(Length), ByVal VarPtr(inBytes(pos)), 4
                pos = pos + 4
                
                af.FACTChunk.dwFileSize = Length
                
            Case "DATA"
                pos = pos + 4
                RtlMoveMemory VarPtr(Length), ByVal VarPtr(inBytes(pos)), 4
                pos = pos + 4
                
                'Debug.Print Length / af.FMTChunk.nChannels / 2
               'Stop
               
                If af.FMTChunk.wBitsPerSample = 8 Then
            
                    ReDim af.WaveData8bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To Length / af.FMTChunk.nChannels)
                    RtlMoveMemory VarPtr(af.WaveData8bit.ChannelData(LBound(af.WaveData8bit.ChannelData, 1), LBound(af.WaveData8bit.ChannelData, 2))), ByVal VarPtr(inBytes(pos)), Length
                    pos = pos + LengthOfData(af)
                ElseIf af.FMTChunk.wBitsPerSample = 16 Then

                    ReDim af.WaveData16bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To Length / af.FMTChunk.nChannels / 2)
                    RtlMoveMemory VarPtr(af.WaveData16bit.ChannelData(LBound(af.WaveData16bit.ChannelData, 1), LBound(af.WaveData16bit.ChannelData, 2))), ByVal VarPtr(inBytes(pos)), Length
                    pos = pos + LengthOfData(af)
    
                End If

        End Select

    Loop
    
    WaveBytesAsAudio = af
    
End Function

Public Function WaveRecordToBytes(ByRef af As AudioFile2) As Byte()

    Dim inBytes() As Byte
    ReDim inBytes(1 To 4) As Byte
    AddDesc inBytes, "RIFF", 1
    AddLong inBytes, af.RiffHdr.Len
    AddDesc inBytes, "WAVE"

    AddDesc inBytes, "fmt "
    AddLong inBytes, LenB(af.FMTChunk)
    
    ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FMTChunk)) As Byte
    RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FMTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FMTChunk)
    
    If af.FACTChunk.dwFileSize <> 0 Then
    
        AddDesc inBytes, "fact"
        AddLong inBytes, LenB(af.FACTChunk)
   
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FACTChunk)) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FACTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FACTChunk)
       
    End If
    
    Dim addLen As Long
    If af.FMTChunk.wBitsPerSample = 8 Then
        AddDesc inBytes, "data"
        addLen = ArrayBoundSize(af.WaveData8bit.ChannelData, 1) * ArrayBoundSize(af.WaveData8bit.ChannelData, 2)
        AddLong inBytes, addLen
        
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + addLen) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - addLen + 1)), _
            VarPtr(af.WaveData8bit.ChannelData(LBound(af.WaveData8bit.ChannelData, 1), LBound(af.WaveData8bit.ChannelData, 2))), addLen
            
    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        AddDesc inBytes, "data"
        addLen = (ArrayBoundSize(af.WaveData16bit.ChannelData, 1) * ArrayBoundSize(af.WaveData16bit.ChannelData, 2) * 2)
        AddLong inBytes, addLen

       
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + addLen) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - addLen + 1)), _
            VarPtr(af.WaveData16bit.ChannelData(LBound(af.WaveData16bit.ChannelData, 1), LBound(af.WaveData16bit.ChannelData, 2))), addLen

    End If
    
    WaveRecordToBytes = inBytes
    
End Function

'Public Sub GetWavPart(ByRef af As AudioFile2, ByRef inBytes() As Byte, Optional ByVal OffSet As Single = 0, Optional ByVal Length As Single = -1)
'
'    ReDim inBytes(1 To 4) As Byte
'    AddDesc inBytes, "RIFF", 1
'
'    If Length = -1 Or Length > af.RiffHdr.Len - OffSet Then
'        Length = af.RiffHdr.Len - OffSet
'    End If
'    AddLong inBytes, Length
'
'    AddDesc inBytes, "WAVE"
'
'    AddDesc inBytes, "fmt "
'    AddLong inBytes, LenB(af.FMTChunk)
'
'    ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FMTChunk)) As Byte
'    RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FMTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FMTChunk)
'
'    If af.FACTChunk.dwFileSize <> 0 Then
'
'        AddDesc inBytes, "fact"
'        AddLong inBytes, LenB(af.FACTChunk)
'
'        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FACTChunk)) As Byte
'        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FACTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FACTChunk)
'
'    End If
'
'    Dim addLen As Long
'    If af.FMTChunk.wBitsPerSample = 8 Then
'        AddDesc inBytes, "data"
'
'        addLen = (ArrayBoundSize(af.WaveData8bit.ChannelData, 1) * ArrayBoundSize(af.WaveData8bit.ChannelData, 2)) - OffSet
'
'        If Length > addLen Then
'            Length = addLen
'        Else
'            addLen = Length
'        End If
'
'        AddLong inBytes, addLen
'
'        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + addLen) As Byte
'        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - addLen + 1)), _
'            VarPtr(af.WaveData8bit.ChannelData(LBound(af.WaveData8bit.ChannelData, 1), (LBound(af.WaveData8bit.ChannelData, 2) + OffSet))), addLen
'
'    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
'        AddDesc inBytes, "data"
'        addLen = (ArrayBoundSize(af.WaveData16bit.ChannelData, 1) * ArrayBoundSize(af.WaveData16bit.ChannelData, 2) - OffSet)
'
'        If Length > addLen Then
'            Length = addLen * 2
'        Else
'            addLen = Length * 2
'        End If
'
'        AddLong inBytes, addLen
'
'        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + addLen) As Byte
'        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - addLen + 1)), _
'            VarPtr(af.WaveData16bit.ChannelData(LBound(af.WaveData16bit.ChannelData, 1), LBound(af.WaveData16bit.ChannelData, 2) + (OffSet / 2) + 1)), addLen
'
'    End If
'End Sub

Public Sub WavePCM16bitStereo(ByRef af As AudioFile2)
    With af.FMTChunk
        .wFormatTag = 1
        .nChannels = 2
        .nSamplesPerSec = 44100
        .nAvgBytesPerSec = 176400
        .nBlockAlign = 4
        .wBitsPerSample = 16
    End With
End Sub

Public Function WaveSilentAudio(ByVal DurationMS As Single) As Byte()

    Dim af As AudioFile2
    
    WavePCM16bitStereo af

    Dim inBytes() As Byte
    ReDim inBytes(1 To 4) As Byte
    AddDesc inBytes, "RIFF", 1
    
   ' DurationMS = Round(af.FMTChunk.nAvgBytesPerSec * (DurationMS / 1000))
    DurationMS = (((af.FMTChunk.nSamplesPerSec * (DurationMS / 1000)) * af.FMTChunk.wBitsPerSample) / af.FMTChunk.nBlockAlign)

    If DurationMS > 0 Then
        If af.FMTChunk.wBitsPerSample = 8 Then
            
            ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + DurationMS) As Byte
            ReDim af.WaveData8bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To DurationMS / af.FMTChunk.nChannels)
            
        ElseIf af.FMTChunk.wBitsPerSample = 16 Then
           
            ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + DurationMS * 2) As Byte
            ReDim af.WaveData16bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To (DurationMS / af.FMTChunk.nChannels) / 2)
            
        End If
        AddLong inBytes, LengthOfWave(af)
    Else
        AddLong inBytes, LengthOfWave(af, 0)
    End If
    
    
    AddDesc inBytes, "WAVE"

    AddDesc inBytes, "fmt "
    AddLong inBytes, LenB(af.FMTChunk)
    
    ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FMTChunk)) As Byte
    RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FMTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FMTChunk)
    
    If af.FACTChunk.dwFileSize <> 0 Then
    
        AddDesc inBytes, "fact"
        AddLong inBytes, LenB(af.FACTChunk)
   
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FACTChunk)) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - LenB(af.FACTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FACTChunk)
       
    End If


    If af.FMTChunk.wBitsPerSample = 8 Then
        AddDesc inBytes, "data"
   
        AddLong inBytes, DurationMS
        
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + DurationMS) As Byte
            
    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        AddDesc inBytes, "data"
        
        AddLong inBytes, DurationMS * 2
       
        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + DurationMS * 2) As Byte

    End If
    WaveSilentAudio = inBytes
    
End Function
'Public Function WaveAudioPartial(ByRef af As AudioFile2, Optional ByVal StartMS As Single = 0, Optional ByVal DurationMS As Single = -1) As Byte()
'
'    Dim InBytes() As Byte
'    ReDim InBytes(1 To 4) As Byte
'    AddDesc InBytes, "RIFF", 1
'
'    Dim bStart As Single
'    Dim bDuration As Single
'
'
'    bStart = Round((af.FMTChunk.nSamplesPerSec / 1000) * StartMS)
'    bDuration = Round(af.FMTChunk.nAvgBytesPerSec * (DurationMS / 1000))
'
'    If bDuration = -1 Or bDuration > af.RiffHdr.Len - bStart Then
'        bDuration = af.RiffHdr.Len - bStart
'    End If
'
'
'    AddLong InBytes, bDuration
'
'
'    AddDesc InBytes, "WAVE"
'
'    AddDesc InBytes, "fmt "
'    AddLong InBytes, LenB(af.FMTChunk)
'
'    ReDim Preserve InBytes(LBound(InBytes) To UBound(InBytes) + LenB(af.FMTChunk)) As Byte
'    RtlMoveMemory ByVal VarPtr(InBytes(UBound(InBytes) - LenB(af.FMTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FMTChunk)
'
'    If af.FACTChunk.dwFileSize <> 0 Then
'
'        AddDesc InBytes, "fact"
'        AddLong InBytes, LenB(af.FACTChunk)
'
'        ReDim Preserve InBytes(LBound(InBytes) To UBound(InBytes) + LenB(af.FACTChunk)) As Byte
'        RtlMoveMemory ByVal VarPtr(InBytes(UBound(InBytes) - LenB(af.FACTChunk) + 1)), VarPtr(af.FMTChunk), LenB(af.FACTChunk)
'
'    End If
'
'
'    Dim addLen As Long
'    If af.FMTChunk.wBitsPerSample = 8 Then
'        AddDesc InBytes, "data"
'
'        addLen = (ArrayBoundSize(af.WaveData8bit.ChannelData, 1) * ArrayBoundSize(af.WaveData8bit.ChannelData, 2)) - bStart
'
'        If bDuration > addLen Then
'            bDuration = addLen
'        Else
'            addLen = bDuration
'        End If
'
'        AddLong InBytes, addLen
'
'        ReDim Preserve InBytes(LBound(InBytes) To UBound(InBytes) + addLen) As Byte
'        RtlMoveMemory ByVal VarPtr(InBytes(UBound(InBytes) - addLen + 1)), _
'            VarPtr(af.WaveData8bit.ChannelData(LBound(af.WaveData8bit.ChannelData, 1), (LBound(af.WaveData8bit.ChannelData, 2) + StartMS))), addLen
'
'    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
'        AddDesc InBytes, "data"
'
'
'        AddLong InBytes, bDuration
'
'        ReDim Preserve InBytes(LBound(InBytes) To UBound(InBytes) + bDuration) As Byte
'        RtlMoveMemory ByVal VarPtr(InBytes(UBound(InBytes) - bDuration + 1)), _
'            VarPtr(af.WaveData16bit.ChannelData(LBound(af.WaveData16bit.ChannelData, 1), LBound(af.WaveData16bit.ChannelData, 2) + (bStart / 2) + 1)), bDuration
'
'    End If
'    WaveAudioPartial = InBytes
'
'    DebugWav af
'End Function

Public Function WaveAudioPartial(ByRef af As AudioFile2, Optional ByVal StartMS As Single = 0, Optional ByVal DurationMS As Single = -1) As Byte()

    Dim inBytes() As Byte
    ReDim inBytes(1 To 4) As Byte
    AddDesc inBytes, "RIFF", 1

    Dim bStart As Single
    Dim bDuration As Single


    bStart = Round((af.FMTChunk.nSamplesPerSec / 1000) * StartMS)
    bDuration = Round(af.FMTChunk.nAvgBytesPerSec * (DurationMS / 1000))

    If bDuration = -1 Or bDuration > af.RiffHdr.Len - bStart Then
        bDuration = af.RiffHdr.Len - bStart
    End If


    AddLong inBytes, bDuration


    AddDesc inBytes, "WAVE"

    AddDesc inBytes, "fmt "
    AddLong inBytes, LenB(af.FMTChunk)

    ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FMTChunk)) As Byte
    RtlMoveMemory VarPtr(inBytes(UBound(inBytes) - LenB(af.FMTChunk) + 1)), ByVal VarPtr(af.FMTChunk), LenB(af.FMTChunk)

    If af.FACTChunk.dwFileSize <> 0 Then

        AddDesc inBytes, "fact"
        AddLong inBytes, LenB(af.FACTChunk)

        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + LenB(af.FACTChunk)) As Byte
        RtlMoveMemory VarPtr(inBytes(UBound(inBytes) - LenB(af.FACTChunk) + 1)), ByVal VarPtr(af.FMTChunk), LenB(af.FACTChunk)

    End If


    Dim addLen As Long
    If af.FMTChunk.wBitsPerSample = 8 Then
        AddDesc inBytes, "data"

        addLen = (ArrayBoundSize(af.WaveData8bit.ChannelData, 1) * ArrayBoundSize(af.WaveData8bit.ChannelData, 2)) - bStart

        If bDuration > addLen Then
            bDuration = addLen
        Else
            addLen = bDuration
        End If

        AddLong inBytes, addLen

        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + addLen) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - addLen + 1)), _
            VarPtr(af.WaveData8bit.ChannelData(LBound(af.WaveData8bit.ChannelData, 1), (LBound(af.WaveData8bit.ChannelData, 2) + StartMS))), addLen

    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        AddDesc inBytes, "data"


        AddLong inBytes, bDuration

        ReDim Preserve inBytes(LBound(inBytes) To UBound(inBytes) + bDuration) As Byte
        RtlMoveMemory ByVal VarPtr(inBytes(UBound(inBytes) - bDuration + 1)), _
            VarPtr(af.WaveData16bit.ChannelData(LBound(af.WaveData16bit.ChannelData, 1), LBound(af.WaveData16bit.ChannelData, 2) + (bStart / 2) + 1)), bDuration

    End If
    WaveAudioPartial = inBytes

    DebugWav af
End Function


Public Sub WaveSaveToFile(ByRef af As AudioFile2, FileName As String)

    If af.RiffHdr.ID <> "RIFF" Or af.RiffHdr.DataType <> "WAVE" Then
        MsgBox "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
        Exit Sub
    End If
    
    Dim FreeNum As Integer
        
    FreeNum = FreeFile
    Open FileName For Output As FreeNum
    Close #FreeNum
    Open FileName For Binary Access Write As FreeNum
        
    Put #FreeNum, 1, af.RiffHdr
    
    ChunkHdr.ID = "fmt "
    ChunkHdr.Len = LenB(af.FMTChunk)
    
    Put FreeNum, , ChunkHdr
    Put FreeNum, , af.FMTChunk
    
    If af.FACTChunk.dwFileSize <> 0 Then
    
        ChunkHdr.ID = "fact"
        ChunkHdr.Len = LenB(af.FACTChunk)
        
        Put FreeNum, , ChunkHdr
        Put FreeNum, , af.FACTChunk

    End If
    
    If af.FMTChunk.wBitsPerSample = 8 Then
        ChunkHdr.ID = "data"
        ChunkHdr.Len = UBound(af.WaveData8bit.ChannelData, 1) * UBound(af.WaveData8bit.ChannelData, 2)
        
        Put FreeNum, , ChunkHdr
        Put FreeNum, , af.WaveData8bit.ChannelData
        
    ElseIf af.FMTChunk.wBitsPerSample = 16 Then
        ChunkHdr.ID = "data"
        ChunkHdr.Len = (UBound(af.WaveData16bit.ChannelData, 1) * UBound(af.WaveData16bit.ChannelData, 2) * 2)

        Put FreeNum, , ChunkHdr
        Put FreeNum, , af.WaveData16bit.ChannelData

    End If
    
End Sub

Public Function WaveLoadFromFile(FileName As String) As AudioFile2
    On Error GoTo fail:
    
    Dim af As AudioFile2
    
    Dim FreeNum As Integer

    Dim TmpSeek As Long

    

    
    FreeNum = FreeFile
    Open FileName For Binary Access Read As FreeNum
    
    'Get RIFF header
    Get #FreeNum, 1, af.RiffHdr
    
    If af.RiffHdr.ID <> "RIFF" Or af.RiffHdr.DataType <> "WAVE" Then
        MsgBox "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
        Exit Function
    End If
    
    
    TmpSeek = Seek(FreeNum)
    'Read Each Chunk in File
    Do
        'Save current position in file
        
        'Read Next Chunk Header
        Get #FreeNum, , ChunkHdr

        'Proccess Chunks
        If UCase(ChunkHdr.ID) = "FMT " Then
            Get #FreeNum, , af.FMTChunk
            'Seek #FreeNum, TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
            TmpSeek = TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
        ElseIf UCase(ChunkHdr.ID) = "FACT" Then
            Get #FreeNum, , af.FACTChunk
            'Seek #FreeNum, TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
            TmpSeek = TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
        ElseIf UCase(ChunkHdr.ID) = "DATA" Then

            
            If af.FMTChunk.wBitsPerSample = 8 Then
                ReDim af.WaveData8bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To ChunkHdr.Len / af.FMTChunk.nChannels)
                Get #FreeNum, , af.WaveData8bit.ChannelData
            

            ElseIf af.FMTChunk.wBitsPerSample = 16 Then
 
            
                ReDim af.WaveData16bit.ChannelData(1 To af.FMTChunk.nChannels, 1 To (ChunkHdr.Len / af.FMTChunk.nChannels) / 2)
                Get #FreeNum, , af.WaveData16bit.ChannelData
            Else
                MsgBox "the bits per sample of this file is not supported."
                GoTo fail
                
                
            End If
            TmpSeek = TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
            'Seek #FreeNum, TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
        Else
            'Seek #FreeNum, TmpSeek + 1
            TmpSeek = TmpSeek + 1
        End If
        
        'Test If Last Chunk Has been found
        If TmpSeek >= LOF(FreeNum) Then
            Exit Do
        End If
        
        Seek #FreeNum, TmpSeek

                    
    Loop
       
        
    Close FreeNum
    WaveLoadFromFile = af

    Debug.Print "LengthOfFile: " & FileLen(FileName)
    DebugWav af
    Exit Function
fail:
    Close FreeNum
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
        Err.Clear
    End If
    af.RiffHdr.DataType = ""
End Function

'Public Function DrawWave8(ByRef af As AudioFile2, Picbox As PictureBox)
'    Dim nChannels As Integer    'Number of Channels in WaveForm
'    Dim nSamples As Long        'Number of Samples per channel
'    Dim Loopsamples As Long
'
'    Channels = UBound(af.WaveData8bit.ChannelData, 1)
'    nSamples = UBound(af.WaveData8bit.ChannelData, 2)
'
'    Picbox.ScaleMode = 0
'    Picbox.ScaleHeight = 2 ^ af.FMTChunk.wBitsPerSample
'    Picbox.ScaleWidth = nSamples
'
'    Picbox.CurrentY = (2 ^ af.FMTChunk.wBitsPerSample) / 2
'
'    Picbox.Visible = False
'    For Loopsamples = 1 To nSamples
'        Picbox.Line -(Loopsamples, af.WaveData8bit.ChannelData(1, Loopsamples))
'    Next
'    Picbox.Visible = True
'End Function
'Public Function DrawWave16(ByRef af As AudioFile2, Picbox As PictureBox)
'    Dim nChannels As Integer    'Number of Channels in WaveForm
'    Dim nSamples As Long        'Number of Samples per channel
'    Dim Loopsamples As Long
'
'    nChannels = UBound(af.WaveData16bit.ChannelData, 1)
'    nSamples = UBound(af.WaveData16bit.ChannelData, 2)
'    Picbox.ScaleMode = 0
'    Picbox.ScaleHeight = (2 ^ af.FMTChunk.wBitsPerSample)
'    Picbox.ScaleWidth = nSamples
'
'    Picbox.CurrentY = (2 ^ af.FMTChunk.wBitsPerSample) / 2
'
'    Picbox.Visible = False
'    For Loopsamples = 1 To nSamples
'        Picbox.Line -(Loopsamples, af.WaveData16bit.ChannelData(1, Loopsamples) + (2 ^ af.FMTChunk.wBitsPerSample) / 2)
'    Next
'    Picbox.Visible = True
'End Function
'
