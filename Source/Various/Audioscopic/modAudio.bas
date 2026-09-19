Attribute VB_Name = "modAudio"
Option Explicit
Option Compare Binary
#Const modAudio = -1

Private Function CancelCalled() As Boolean
    CancelCalled = (Not frmMagic.Visible)
End Function
Public Function SoundLevelAtDistance(ByVal distanceFeet As Double, Optional ByVal sourceDB As Double = 0) As Double
    If distanceFeet <= 0 Then
        ' Prevent invalid log(0)
        SoundLevelAtDistance = sourceDB
        Exit Function
    End If
    
    ' Convert natural log to log10
    Const LN10 As Double = 2.30258509299405
    
    Dim dropDB As Double
    dropDB = 20 * (Log(distanceFeet) / LN10)
    
    SoundLevelAtDistance = sourceDB - dropDB
End Function

'example:
'Dim level As Double
'level = SoundLevelAtDistance(100, 5)
'Print level   ' ~86 dB


Public Function SoundLevelAtTime(ByVal sourceDB As Double, ByVal timeSeconds As Double) As Double
    Const speedOfSound As Double = 1125.32808
    Dim distanceFeet As Double
    
    distanceFeet = timeSeconds * speedOfSound
    SoundLevelAtTime = SoundLevelAtDistance(sourceDB, distanceFeet)
End Function

Public Function SoundTravelTimeSeconds(ByVal distanceFeet As Double) As Double
    Const speedOfSound As Double = 1125.32808 ' ft/s
    
    If distanceFeet <= 0 Then
        SoundTravelTimeSeconds = 0
        Exit Function
    End If
    
    SoundTravelTimeSeconds = distanceFeet / speedOfSound
    
End Function


Public Function SoundTravelTimeMilliseconds(ByVal distanceFeet As Double) As Double
    Const speedOfSound As Double = 1125.32808 ' ft/s
    
    If distanceFeet <= 0 Then
        SoundTravelTimeMilliseconds = 0
        Exit Function
    End If
    
    SoundTravelTimeMilliseconds = (distanceFeet / speedOfSound) * 1000

End Function

'example:
'Dim ms As Double
'ms = SoundTravelTimeMilliseconds(5)
'Print ms   ' ~4.44 ms

Public Function BytesToInt(ByVal B1 As Variant, ByVal B2 As Variant) As Variant

    BytesToInt = B1 + (B2 * &H100&)
    If BytesToInt > 32767 Then BytesToInt = BytesToInt - 65536
    BytesToInt = CInt(BytesToInt)
End Function

Public Function BytesToLong(ByVal B1 As Variant, ByVal B2 As Variant, ByRef b3 As Variant, ByRef b4 As Variant) As Variant
    BytesToLong = B1 + (B2 * &H100&)
    BytesToLong = BytesToLong + (b3 * &H10000)
    BytesToLong = BytesToLong + (b4 * &H1000000)
    BytesToLong = CLng(BytesToLong)
End Function
Public Sub IntToByes(ByVal I1 As Integer, ByRef B1 As Variant, ByRef B2 As Variant)
    
    B1 = I1 And &HFF&
    B2 = (I1 \ &H100&) And &HFF&

End Sub

Public Sub LongToBytes(ByVal L1 As Long, ByRef B1 As Variant, ByRef B2 As Variant, ByRef b3 As Variant, ByRef b4 As Variant)

    B1 = L1 And &HFF&
    B2 = (L1 \ &H100&) And &HFF&
    b3 = (L1 \ &H10000) And &HFF&
    b4 = (L1 \ &H1000000) And &HFF&
    
End Sub

Public Function ChangeWaveVolume(ByVal inFile As String, ByVal outFile As String, ByVal dBChange As Double) As Boolean
    On Error GoTo ErrHandler
    
    Dim gain As Double
    gain = 10 ^ (dBChange / 20#)

    Dim b() As Byte
    Dim f As Integer
    f = FreeFile

    ' Load entire WAV file
    Open inFile For Binary As #f
        ReDim b(LOF(f) - 1)
        Get #f, , b
    Close #f

    ' PCM data starts at offset 44 for standard WAV
    Dim i As Long
    Dim sample As Double

    For i = 44 To UBound(b) Step 2
        sample = BytesToInt(b(i), b(i + 1))
        If sample > 32767 Then sample = sample - 65536

        sample = CLng(sample * gain)

        If sample > 32767 Then sample = 32767
        If sample < -32768 Then sample = -32768

        b(i) = sample And &HFF
        b(i + 1) = (sample \ &H100) And &HFF
    Next i

    f = FreeFile
    Open outFile For Binary As #f
        Put #f, , b
    Close #f

    ChangeWaveVolume = True
    Exit Function

ErrHandler:
    Close #f
    ChangeWaveVolume = False
End Function

'Public Function ChangeWaveVolume(ByVal InFile As String, ByVal OutFile As String, ByVal dBChange As Double) As Boolean
'    On Error GoTo ErrHandler
'    'If CancelCalled Then GoTo ErrHandler
'
'    Dim gain As Double
'    gain = 10 ^ (dBChange / 20#)
'
'    Dim b() As Byte
'    Dim f As Integer
'    f = FreeFile
'
'    ' Load entire WAV file
'    Open InFile For Binary As #f
'        ReDim b(LOF(f) - 1)
'        Get #f, 1, b
'    Close #f
'
'    ' PCM data starts at offset 44 for standard WAV
'    Dim i As Long
'    Dim sample As Long
'
'    i = 12
'    Do While i + 3 <= UBound(b) And (Not (UCase(Chr(b(i))) = "D" And UCase(Chr(b(i + 1))) = "A" And UCase(Chr(b(i + 2))) = "T" And UCase(Chr(b(i + 3))) = "A"))
'
'        i = i + 1
'    Loop
'
'    If i + 3 <= UBound(b) Then
'        If (UCase(Chr(b(i))) & UCase(Chr(b(i + 1))) & UCase(Chr(b(i + 2))) & UCase(Chr(b(i + 3))) = "DATA") Then
'
'
'            i = i + 8
'
'
'            'i = 44
'            Do While i <= UBound(b)
'            'For i = 44 To UBound(B) Step 2
'                sample = CLng(BytesToInt(b(i), b(i + 1)))
'                If sample > 32767 Then sample = sample - 65536
'                sample = CLng(sample * gain)
'
'                If sample > 32767 Then sample = 32767
'                If sample < -32768 Then sample = -32768
'
'                b(i) = CInt(sample) And &HFF
'                b(i + 1) = (CInt(sample) \ &H100) And &HFF
'
'                'If CancelCalled Then GoTo ErrHandler
'                i = i + 2
'            Loop
'            'Next i
'
'            f = FreeFile
'            Open OutFile For Binary As #f
'                Put #f, 1, b
'            Close #f
'        Else
'            MsgBox "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
'
'        End If
'    Else
'        MsgBox "The file you have selected is not a RIFF nor WAVE file.  Only Standard PCM wave format is supported."
'    End If
'    ChangeWaveVolume = True
'    Exit Function
'
'ErrHandler:
'    ChangeWaveVolume = False
'End Function


Public Function AddSilenceToWav( _
        ByVal inFile As String, _
        ByVal outFile As String, _
        ByVal SilenceMs As Long, _
        ByVal AddAtEnd As Boolean) As Boolean

    On Error GoTo fail

    Dim f As Integer
    Dim hdr(43) As Byte
    Dim data() As Byte
    Dim silence() As Byte

    f = FreeFile
    Open inFile For Binary As #f

    '--- read header
    Get #f, , hdr

    '--- extract format info
    Dim channels As Integer
    Dim sampleRate As Long
    Dim bitsPerSample As Integer
    Dim blockAlign As Integer
    Dim dataSize As Long

    channels = hdr(22) + hdr(23) * &H100&
    sampleRate = hdr(24) + hdr(25) * &H100& + hdr(26) * &H10000 + hdr(27) * &H1000000
    bitsPerSample = hdr(34) + hdr(35) * &H100&
    blockAlign = hdr(32) + hdr(33) * &H100&
    dataSize = hdr(40) + hdr(41) * &H100& + hdr(42) * &H10000 + hdr(43) * &H1000000

    '--- read PCM data
    ReDim data(dataSize - 1)
    Get #f, , data
    Close #f

    '--- compute silence length in bytes
    Dim bytesPerMs As Long
    bytesPerMs = (sampleRate * blockAlign) \ 1000

    Dim silenceBytes As Long
    silenceBytes = bytesPerMs * SilenceMs

    ReDim silence(silenceBytes - 1)

    '--- build new PCM buffer
    Dim newData() As Byte
    Dim newSize As Long

    If AddAtEnd = False Then
        ' silence at beginning
        newSize = silenceBytes + dataSize
        ReDim newData(newSize - 1)
        ' copy silence
        CopyMemory newData(0), silence(0), silenceBytes
        ' copy original
        CopyMemory newData(silenceBytes), data(0), dataSize
    Else
        ' silence at end
        newSize = dataSize + silenceBytes
        ReDim newData(newSize - 1)
        ' copy original
        CopyMemory newData(0), data(0), dataSize
        ' copy silence
        CopyMemory newData(dataSize), silence(0), silenceBytes
    End If

    '--- update header
    Dim riffSize As Long
    riffSize = 36 + newSize

    hdr(4) = riffSize And &HFF&
    hdr(5) = (riffSize \ &H100&) And &HFF&
    hdr(6) = (riffSize \ &H10000) And &HFF&
    hdr(7) = (riffSize \ &H1000000) And &HFF&

    hdr(40) = newSize And &HFF&
    hdr(41) = (newSize \ &H100&) And &HFF&
    hdr(42) = (newSize \ &H10000) And &HFF&
    hdr(43) = (newSize \ &H1000000) And &HFF&

    '--- write output file
    f = FreeFile
    Open outFile For Binary As #f
    Put #f, , hdr
    Put #f, , newData
    Close #f

    AddSilenceToWav = True
    Exit Function

fail:
    Close #f
    AddSilenceToWav = False
End Function
'Public Function AddWavFiles( _
'        ByVal FileA As String, _
'        ByVal FileB As String, _
'        ByVal OutFile As String) As Boolean
'
'    On Error GoTo Fail
'    If CancelCalled Then GoTo Fail
'
'    Dim hdrA(43) As Byte, hdrB(43) As Byte
'    Dim dataA() As Byte, dataB() As Byte
'    Dim f As Integer
'
'    '--- read header A
'    f = FreeFile
'    Open FileA For Binary As #f
'    Get #f, , hdrA
'    Dim sizeA As Long
'    sizeA = BytesToLong(hdrA(40), hdrA(41), hdrA(42), hdrA(43))
'    ReDim dataA(sizeA - 1)
'    Get #f, , dataA
'    Close #f
'
'    '--- read header B
'    f = FreeFile
'    Open FileB For Binary As #f
'    Get #f, , hdrB
'    Dim sizeB As Long
'    sizeB = BytesToLong(hdrB(40), hdrB(41), hdrB(42), hdrB(43))
'    ReDim dataB(sizeB - 1)
'    Get #f, , dataB
'    Close #f
'
'    '--- extract format (must match)
'    Dim bits As Integer
'    Dim blockAlign As Integer
'
'    bits = BytesToInt(hdrA(34), hdrA(35))
'    blockAlign = BytesToInt(hdrA(32), hdrA(33))
'
'    '--- choose larger size
'    Dim outSize As Long
'    outSize = IIf(sizeA > sizeB, sizeA, sizeB)
'
'    Dim outData() As Byte
'    ReDim outData(outSize - 1)
'
'    Dim i As Long
'
'    If bits = 8 Then
'        '--- 8-bit unsigned PCM
'        For i = 0 To outSize - 1
'            Dim a As Integer, b As Integer, m As Integer
'
'            a = IIf(i < sizeA, dataA(i), 128)
'            b = IIf(i < sizeB, dataB(i), 128)
'
'            m = a + b - 128
'            If m < 0 Then m = 0
'            If m > 255 Then m = 255
'
'            outData(i) = m
'
'            If CancelCalled Then GoTo Fail
'        Next i
'
'    ElseIf bits = 16 Then
'        '--- 16-bit signed PCM
'        For i = 0 To outSize - 2 Step 2
'            Dim sa As Long, sb As Long, sm As Long
'
'            '--- sample A
'            If i < sizeA - 2 Then
'                sa = BytesToInt(dataA(i), dataA(i + 1))
'                If sa > 32767 Then sa = sa - 65536
'            Else
'                sa = 0
'            End If
'
'            '--- sample B
'            If i < sizeB - 2 Then
'                sb = BytesToInt(dataB(i), dataB(i + 1))
'                If sb > 32767 Then sb = sb - 65536
'            Else
'                sb = 0
'            End If
'
'            '--- mix
'            sm = sa + sb
'            If sm < -32768 Then sm = -32768
'            If sm > 32767 Then sm = 32767
'
'            '--- store
'            Dim us As Long
'            If sm < 0 Then
'                us = sm + 65536
'            Else
'                us = sm
'            End If
'
'            outData(i) = us And &HFF&
'            outData(i + 1) = (us \ &H100&) And &HFF&
'            If CancelCalled Then GoTo Fail
'        Next i
'    End If
'
'    '--- update header A (use A's format)
'    Dim riffSize As Long
'    riffSize = 36 + outSize
'
'    hdrA(4) = riffSize And &HFF&
'    hdrA(5) = (riffSize \ &H100&) And &HFF&
'    hdrA(6) = (riffSize \ &H10000) And &HFF&
'    hdrA(7) = (riffSize \ &H1000000) And &HFF&
'
'    hdrA(36) = Asc("d")
'    hdrA(37) = Asc("A")
'    hdrA(38) = Asc("T")
'    hdrA(39) = Asc("A")
'    hdrA(40) = outSize And &HFF&
'    hdrA(41) = (outSize \ &H100&) And &HFF&
'    hdrA(42) = (outSize \ &H10000) And &HFF&
'    hdrA(43) = (outSize \ &H1000000) And &HFF&
'
'    '--- write output
'    f = FreeFile
'    Open OutFile For Binary As #f
'    Put #f, , hdrA
'    Put #f, 45, outData
'    Close #f
'
'    AddWavFiles = True
'    Exit Function
'
'Fail:
'    Close f
'    AddWavFiles = False
'End Function

Public Function MixWavFiles( _
        ByVal FileA As String, _
        ByVal FileB As String, _
        ByVal outFile As String) As Boolean

    On Error GoTo fail

    Dim hdrA(43) As Byte, hdrB(43) As Byte
    Dim dataA() As Byte, dataB() As Byte
    Dim f As Integer

    '--- read header A
    f = FreeFile
    Open FileA For Binary As #f
    Get #f, , hdrA
    Dim sizeA As Long
    sizeA = hdrA(40) + hdrA(41) * &H100& + hdrA(42) * &H10000 + hdrA(43) * &H1000000
    ReDim dataA(sizeA - 1)
    Get #f, , dataA
    Close #f

    '--- read header B
    f = FreeFile
    Open FileB For Binary As #f
    Get #f, , hdrB
    Dim sizeB As Long
    sizeB = hdrB(40) + hdrB(41) * &H100& + hdrB(42) * &H10000 + hdrB(43) * &H1000000
    ReDim dataB(sizeB - 1)
    Get #f, , dataB
    Close #f

    '--- extract format (must match)
    Dim bits As Integer
    Dim blockAlign As Integer

    bits = hdrA(34) + hdrA(35) * &H100&
    blockAlign = hdrA(32) + hdrA(33) * &H100&

    '--- choose larger size
    Dim outSize As Long
    outSize = IIf(sizeA > sizeB, sizeA, sizeB)

    Dim outData() As Byte
    ReDim outData(outSize - 1)

    Dim i As Long

    If bits = 8 Then
        '--- 8-bit unsigned PCM
        For i = 0 To outSize - 1
            Dim a As Integer, b As Integer, m As Integer

            a = IIf(i < sizeA, dataA(i), 128)
            b = IIf(i < sizeB, dataB(i), 128)

            m = (a + b) - 128
            If m < 0 Then m = 0
            If m > 255 Then m = 255

            outData(i) = m
        Next i

    ElseIf bits = 16 Then
        '--- 16-bit signed PCM
        For i = 0 To outSize - 1 Step 2
            Dim sa As Long, sb As Long, sm As Long

            '--- sample A
            If i < sizeA Then
                sa = BytesToInt(dataA(i), dataA(i + 1))
                If sa > 32767 Then sa = sa - 65536
            Else
                sa = 0
            End If

            '--- sample B
            If i < sizeB Then
                sb = BytesToInt(dataB(i), dataB(i + 1))
                If sb > 32767 Then sb = sb - 65536
            Else
                sb = 0
            End If

            '--- mix
            sm = (sa + sb)
            If sm < -32768 Then sm = -32768
            If sm > 32767 Then sm = 32767

            '--- store
            Dim us As Long
            If sm < 0 Then
                us = sm + 65536
            Else
                us = sm
            End If

            outData(i) = us And &HFF&
            outData(i + 1) = (us \ &H100&) And &HFF&
        Next i
    End If

    '--- update header A (use A's format)
    Dim riffSize As Long
    riffSize = 36 + outSize

    hdrA(4) = riffSize And &HFF&
    hdrA(5) = (riffSize \ &H100&) And &HFF&
    hdrA(6) = (riffSize \ &H10000) And &HFF&
    hdrA(7) = (riffSize \ &H1000000) And &HFF&

    hdrA(40) = outSize And &HFF&
    hdrA(41) = (outSize \ &H100&) And &HFF&
    hdrA(42) = (outSize \ &H10000) And &HFF&
    hdrA(43) = (outSize \ &H1000000) And &HFF&

    '--- write output
    f = FreeFile
    Open outFile For Binary As #f
    Put #f, , hdrA
    Put #f, , outData
    Close #f

    MixWavFiles = True
    Exit Function

fail:
    Close #f
    MixWavFiles = False
End Function

Public Function DeductWavFile( _
        ByVal FileA As String, _
        ByVal FileB As String, _
        ByVal outFile As String) As Boolean

    On Error GoTo fail

    Dim hdrA(43) As Byte, hdrB(43) As Byte
    Dim dataA() As Byte, dataB() As Byte
    Dim f As Integer

    '--- read header A
    f = FreeFile
    Open FileA For Binary As #f
    Get #f, , hdrA
    Dim sizeA As Long
    sizeA = hdrA(40) + hdrA(41) * &H100& + hdrA(42) * &H10000 + hdrA(43) * &H1000000
    ReDim dataA(sizeA - 1)
    Get #f, , dataA
    Close #f

    '--- read header B
    f = FreeFile
    Open FileB For Binary As #f
    Get #f, , hdrB
    Dim sizeB As Long
    sizeB = hdrB(40) + hdrB(41) * &H100& + hdrB(42) * &H10000 + hdrB(43) * &H1000000
    ReDim dataB(sizeB - 1)
    Get #f, , dataB
    Close #f

    '--- extract format (must match)
    Dim bits As Integer
    Dim blockAlign As Integer

    bits = hdrA(34) + hdrA(35) * &H100&
    blockAlign = hdrA(32) + hdrA(33) * &H100&

    '--- choose larger size
    Dim outSize As Long
    outSize = IIf(sizeA > sizeB, sizeA, sizeB)

    Dim outData() As Byte
    ReDim outData(outSize - 1)

    Dim i As Long

    If bits = 8 Then
        '--- 8-bit unsigned PCM
        For i = 0 To outSize - 1
            Dim a As Integer, b As Integer, m As Integer

            a = IIf(i < sizeA, dataA(i), 128)
            b = IIf(i < sizeB, dataB(i), 128)

            m = (a - b) - 128
            If m < 0 Then m = 0
            If m > 255 Then m = 255

            outData(i) = m
        Next i

    ElseIf bits = 16 Then
        '--- 16-bit signed PCM
        For i = 0 To outSize - 1 Step 2
            Dim sa As Long, sb As Long, sm As Long

            '--- sample A
            If i < sizeA Then
                sa = BytesToInt(dataA(i), dataA(i + 1))
                If sa > 32767 Then sa = sa - 65536
            Else
                sa = 0
            End If

            '--- sample B
            If i < sizeB Then
                sb = BytesToInt(dataB(i), dataB(i + 1))
                If sb > 32767 Then sb = sb - 65536
            Else
                sb = 0
            End If

            '--- mix
            sm = (sa - sb)
            If sm < -32768 Then sm = -32768
            If sm > 32767 Then sm = 32767

            '--- store
            Dim us As Long
            If sm < 0 Then
                us = sm + 65536
            Else
                us = sm
            End If

            outData(i) = us And &HFF&
            outData(i + 1) = (us \ &H100&) And &HFF&
        Next i
    End If

    '--- update header A (use A's format)
    Dim riffSize As Long
    riffSize = 36 + outSize

    hdrA(4) = riffSize And &HFF&
    hdrA(5) = (riffSize \ &H100&) And &HFF&
    hdrA(6) = (riffSize \ &H10000) And &HFF&
    hdrA(7) = (riffSize \ &H1000000) And &HFF&

    hdrA(40) = outSize And &HFF&
    hdrA(41) = (outSize \ &H100&) And &HFF&
    hdrA(42) = (outSize \ &H10000) And &HFF&
    hdrA(43) = (outSize \ &H1000000) And &HFF&

    '--- write output
    f = FreeFile
    Open outFile For Binary As #f
    Put #f, , hdrA
    Put #f, , outData
    Close #f

    DeductWavFile = True
    Exit Function

fail:
    Close #f
    DeductWavFile = False
End Function

Public Function AverageWavFiles( _
        ByVal FileA As String, _
        ByVal FileB As String, _
        ByVal outFile As String) As Boolean

    On Error GoTo fail

    Dim hdrA(43) As Byte, hdrB(43) As Byte
    Dim dataA() As Byte, dataB() As Byte
    Dim f As Integer

    '--- read header A
    f = FreeFile
    Open FileA For Binary As #f
    Get #f, , hdrA
    Dim sizeA As Long
    sizeA = hdrA(40) + hdrA(41) * &H100& + hdrA(42) * &H10000 + hdrA(43) * &H1000000
    ReDim dataA(sizeA - 1)
    Get #f, , dataA
    Close #f

    '--- read header B
    f = FreeFile
    Open FileB For Binary As #f
    Get #f, , hdrB
    Dim sizeB As Long
    sizeB = hdrB(40) + hdrB(41) * &H100& + hdrB(42) * &H10000 + hdrB(43) * &H1000000
    ReDim dataB(sizeB - 1)
    Get #f, , dataB
    Close #f

    '--- extract format (must match)
    Dim bits As Integer
    Dim blockAlign As Integer

    bits = hdrA(34) + hdrA(35) * &H100&
    blockAlign = hdrA(32) + hdrA(33) * &H100&

    '--- choose larger size
    Dim outSize As Long
    outSize = IIf(sizeA > sizeB, sizeA, sizeB)

    Dim outData() As Byte
    ReDim outData(outSize - 1)

    Dim i As Long

    If bits = 8 Then
        '--- 8-bit unsigned PCM
        For i = 0 To outSize - 1
            Dim a As Integer, b As Integer, m As Integer

            a = IIf(i < sizeA, dataA(i), 128)
            b = IIf(i < sizeB, dataB(i), 128)

            m = ((a + b) / 2) - 128

            If m < 0 Then m = 0
            If m > 255 Then m = 255

            outData(i) = m
        Next i

    ElseIf bits = 16 Then
        '--- 16-bit signed PCM
        For i = 0 To outSize - 1 Step 2
            Dim sa As Long, sb As Long, sm As Long

            '--- sample A
            If i < sizeA Then
                sa = BytesToInt(dataA(i), dataA(i + 1))
                If sa > 32767 Then sa = sa - 65536
            Else
                sa = 0
            End If

            '--- sample B
            If i < sizeB Then
                sb = BytesToInt(dataB(i), dataB(i + 1))
                If sb > 32767 Then sb = sb - 65536
            Else
                sb = 0
            End If

            '--- mix
            sm = (sa + sb) / 2
            
            If sm < -32768 Then sm = -32768
            If sm > 32767 Then sm = 32767

            '--- store
            Dim us As Long
            If sm < 0 Then
                us = sm + 65536
            Else
                us = sm
            End If

            outData(i) = us And &HFF&
            outData(i + 1) = (us \ &H100&) And &HFF&
        Next i
    End If

    '--- update header A (use A's format)
    Dim riffSize As Long
    riffSize = 36 + outSize

    hdrA(4) = riffSize And &HFF&
    hdrA(5) = (riffSize \ &H100&) And &HFF&
    hdrA(6) = (riffSize \ &H10000) And &HFF&
    hdrA(7) = (riffSize \ &H1000000) And &HFF&

    hdrA(40) = outSize And &HFF&
    hdrA(41) = (outSize \ &H100&) And &HFF&
    hdrA(42) = (outSize \ &H10000) And &HFF&
    hdrA(43) = (outSize \ &H1000000) And &HFF&

    '--- write output
    f = FreeFile
    Open outFile For Binary As #f
    Put #f, , hdrA
    Put #f, , outData
    Close #f

    AverageWavFiles = True
    Exit Function

fail:
    Close #f
    AverageWavFiles = False
End Function

Public Function SubtractWavFile( _
        ByVal FileA As String, _
        ByVal FileB As String, _
        ByVal outFile As String) As Boolean

    On Error GoTo fail

    Dim hdrA(43) As Byte, hdrB(43) As Byte
    Dim dataA() As Byte, dataB() As Byte
    Dim f As Integer

    '--- read header A
    f = FreeFile
    Open FileA For Binary As #f
    Get #f, , hdrA
    Dim sizeA As Long
    sizeA = hdrA(40) + hdrA(41) * &H100& + hdrA(42) * &H10000 + hdrA(43) * &H1000000
    ReDim dataA(sizeA - 1)
    Get #f, , dataA
    Close #f

    '--- read header B
    f = FreeFile
    Open FileB For Binary As #f
    Get #f, , hdrB
    Dim sizeB As Long
    sizeB = hdrB(40) + hdrB(41) * &H100& + hdrB(42) * &H10000 + hdrB(43) * &H1000000
    ReDim dataB(sizeB - 1)
    Get #f, , dataB
    Close #f

    '--- extract format (must match)
    Dim bits As Integer
    Dim blockAlign As Integer

    bits = hdrA(34) + hdrA(35) * &H100&
    blockAlign = hdrA(32) + hdrA(33) * &H100&

    '--- choose larger size
    Dim outSize As Long
    outSize = IIf(sizeA > sizeB, sizeA, sizeB)

    Dim outData() As Byte
    ReDim outData(outSize - 1)

    Dim i As Long

    If bits = 8 Then
        '--- 8-bit unsigned PCM
        For i = 0 To outSize - 1
            Dim a As Integer, b As Integer, m As Integer

            a = IIf(i < sizeA, dataA(i), 128)
            b = IIf(i < sizeB, dataB(i), 128)

            If (Abs(a) > Abs(b)) Then
                If a > 0 Then
                    m = (Abs(a) - Abs(b)) - 128
                ElseIf a < 0 Then
                    m = -(Abs(a) - Abs(b)) + 128
                End If
            ElseIf (Abs(a) < Abs(b)) Then
                If b > 0 Then
                    m = (Abs(b) - Abs(a)) - 128
                ElseIf b < 0 Then
                    m = -(Abs(b) - Abs(a)) + 128
                End If
            End If
            
            
            If m < 0 Then m = 0
            If m > 255 Then m = 255

            outData(i) = m
        Next i

    ElseIf bits = 16 Then
        '--- 16-bit signed PCM
        For i = 0 To outSize - 1 Step 2
            Dim sa As Long, sb As Long, sm As Long

            '--- sample A
            If i < sizeA Then
                sa = BytesToInt(dataA(i), dataA(i + 1))
                If sa > 32767 Then sa = sa - 65536
            Else
                sa = 0
            End If

            '--- sample B
            If i < sizeB Then
                sb = BytesToInt(dataB(i), dataB(i + 1))
                If sb > 32767 Then sb = sb - 65536
            Else
                sb = 0
            End If

            '--- mix
            
            If (Abs(sa) > Abs(sb)) Then
                If sa > 0 Then
                    sm = (Abs(sa) - Abs(sb))
                ElseIf sa < 0 Then
                    sm = -(Abs(sa) - Abs(sb))
                End If
            ElseIf (Abs(sa) < Abs(sb)) Then
                If sb > 0 Then
                    sm = (Abs(sb) - Abs(sa))
                ElseIf sb < 0 Then
                    sm = -(Abs(sb) - Abs(sa))
                End If
            End If
            If sm < -32768 Then sm = -32768
            If sm > 32767 Then sm = 32767

            '--- store
            Dim us As Long
            If sm < 0 Then
                us = sm + 65536
            Else
                us = sm
            End If

            outData(i) = us And &HFF&
            outData(i + 1) = (us \ &H100&) And &HFF&
        Next i
    End If

    '--- update header A (use A's format)
    Dim riffSize As Long
    riffSize = 36 + outSize

    hdrA(4) = riffSize And &HFF&
    hdrA(5) = (riffSize \ &H100&) And &HFF&
    hdrA(6) = (riffSize \ &H10000) And &HFF&
    hdrA(7) = (riffSize \ &H1000000) And &HFF&

    hdrA(40) = outSize And &HFF&
    hdrA(41) = (outSize \ &H100&) And &HFF&
    hdrA(42) = (outSize \ &H10000) And &HFF&
    hdrA(43) = (outSize \ &H1000000) And &HFF&

    '--- write output
    f = FreeFile
    Open outFile For Binary As #f
    Put #f, , hdrA
    Put #f, , outData
    Close #f

    SubtractWavFile = True
    Exit Function

fail:
    Close #f
    SubtractWavFile = False
End Function

Public Function GetPeakDBFS(ByVal wavFile As String) As Double
    On Error GoTo ErrHandler
    If CancelCalled Then GoTo ErrHandler
    
    Dim b() As Byte
    Dim f As Integer
    f = FreeFile

    ' Load entire WAV file
    Open wavFile For Binary As #f
        ReDim b(LOF(f) - 1)
        Get #f, , b
    Close #f

    ' PCM data starts at offset 44
    Dim i As Long
    Dim sample As Long
    Dim peak As Long
    peak = 0

    For i = 44 To UBound(b) - 1 Step 2
        sample = BytesToInt(b(i), b(i + 1))
        
        ' Convert unsigned to signed
        If sample > 32767 Then sample = sample - 65536
        
        ' Track absolute peak
        If Abs(sample) > peak Then peak = Abs(sample)
        If CancelCalled Then GoTo ErrHandler
    Next i

    If peak <= 0 Then
        GetPeakDBFS = -9999 ' silence
        Exit Function
    End If

    ' Convert peak to dBFS
    Const LN10 As Double = 2.30258509299405
    Dim db As Double
    db = 20# * (Log(peak / 32767#) / LN10)

    GetPeakDBFS = peak 'db
    Exit Function

ErrHandler:
    Close #f
    GetPeakDBFS = -9999
End Function


'example:
'Dim peakDB As Double
'peakDB = GetPeakDBFS("C:\audio\test.wav")
'Print "Peak level: "; peakDB; " dBFS"


