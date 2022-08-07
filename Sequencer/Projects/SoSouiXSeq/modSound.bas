Attribute VB_Name = "modSound"
#Const modSound = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public directX As New DirectX8
Public directSound As DirectSound8 'Then there is the sub object, DirectSound:

Public Type dxBuffers
    FileName As String
    Buffer As DirectSoundSecondaryBuffer8
End Type



Public Const MAXPNAMELEN = 32             ' Maximum product name length

' Error values for functions used in this sample. See the function for more information
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)     ' device ID out of range
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)     ' invalid parameter passed
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)        ' no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)           ' memory allocation error

Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)     ' device handle is invalid
Public Const MIDIERR_BASE = 64
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)      ' still something playing
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)          ' hardware is still busy
Public Const MIDIERR_BADOPENMODE = (MIDIERR_BASE + 6)       ' operation unsupported w/ open mode

'User-defined variable the stores information about the MIDI output device.
Type MIDIOUTCAPS
   wMid As Integer                   ' Manufacturer identifier of the device driver for the MIDI output device
                                     ' For a list of identifiers, see the Manufacturer Indentifier topic in the
                                     ' Multimedia Reference of the Platform SDK.
   
   wPid As Integer                   ' Product Identifier Product of the MIDI output device. For a list of
                                     ' product identifiers, see the Product Identifiers topic in the Multimedia
                                     ' Reference of the Platform SDK.
   
   vDriverVersion As Long            ' Version number of the device driver for the MIDI output device.
                                     ' The high-order byte is the major version number, and the low-order byte is
                                     ' the minor version number.
                                     
   szPname As String * MAXPNAMELEN   ' Product name in a null-terminated string.
   
   wTechnology As Integer            ' One of the following that describes the MIDI output device:
                                     '     MOD_FMSYNTH-The device is an FM synthesizer.
                                     '     MOD_MAPPER-The device is the Microsoft MIDI mapper.
                                     '     MOD_MIDIPORT-The device is a MIDI hardware port.
                                     '     MOD_SQSYNTH-The device is a square wave synthesizer.
                                     '     MOD_SYNTH-The device is a synthesizer.
                                     
   wVoices As Integer                ' Number of voices supported by an internal synthesizer device. If the
                                     ' device is a port, this member is not meaningful and is set to 0.
                                     
   wNotes As Integer                 ' Maximum number of simultaneous notes that can be played by an internal
                                     ' synthesizer device. If the device is a port, this member is not meaningful
                                     ' and is set to 0.
                                     
   wChannelMask As Integer           ' Channels that an internal synthesizer device responds to, where the least
                                     ' significant bit refers to channel 0 and the most significant bit to channel
                                     ' 15. Port devices that transmit on all channels set this member to 0xFFFF.
                                     
   dwSupport As Long                 ' One of the following describes the optional functionality supported by
                                     ' the device:
                                     '     MIDICAPS_CACHE-Supports patch caching.
                                     '     MIDICAPS_LRVOLUME-Supports separate left and right volume control.
                                     '     MIDICAPS_STREAM-Provides direct support for the midiStreamOut function.
                                     '     MIDICAPS_VOLUME-Supports volume control.
                                     '
                                     ' If a device supports volume changes, the MIDICAPS_VOLUME flag will be set
                                     ' for the dwSupport member. If a device supports separate volume changes on
                                     ' the left and right channels, both the MIDICAPS_VOLUME and the
                                     ' MIDICAPS_LRVOLUME flags will be set for this member.
End Type

Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
' This function retrieves the number of MIDI output devices present in the system.
' The function returns the number of MIDI output devices. A zero return value means
' there are no MIDI devices in the system.

Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
' This function queries a specified MIDI output device to determine its capabilities.
' The function requires the following parameters;
'     uDeviceID-     unsigned integer variable identifying of the MIDI output device. The
'                    device identifier specified by this parameter varies from zero to one
'                    less than the number of devices present. This parameter can also be a
'                    properly cast device handle.
'     lpMidiOutCaps- address of a MIDIOUTCAPS structure. This structure is filled with
'                    information about the capabilities of the device.
'     cbMidiOutCaps- the size, in bytes, of the MIDIOUTCAPS structure. Use the Len
'                    function with the MIDIOUTCAPS variable as the argument to get
'                    this value.
'
' The function returns MMSYSERR_NOERROR if successful or one of the following error values:
'     MMSYSERR_BADDEVICEID    The specified device identifier is out of range.
'     MMSYSERR_INVALPARAM     The specified pointer or structure is invalid.
'     MMSYSERR_NODRIVER       The driver is not installed.
'     MMSYSERR_NOMEM          The system is unable to load mapper string description.

Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
' The function closes the specified MIDI output device. The function requires a
' handle to the MIDI output device. If the function is successful, the handle is no
' longer valid after the call to this function. A successful function call returns
' MMSYSERR_NOERROR.

' A failure returns one of the following:
'     MIDIERR_STILLPLAYING  Buffers are still in the queue.
'     MMSYSERR_INVALHANDLE  The specified device handle is invalid.
'     MMSYSERR_NOMEM        The system is unable to load mapper string description.

Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
' The function opens a MIDI output device for playback. The function requires the
' following parameters
'     lphmo-               Address of an HMIDIOUT handle. This location is filled with a
'                          handle identifying the opened MIDI output device. The handle
'                          is used to identify the device in calls to other MIDI output
'                          functions.
'     uDeviceID-           Identifier of the MIDI output device that is to be opened.
'     dwCallback-          Address of a callback function, an event handle, a thread
'                          identifier, or a handle of a window or thread called during
'                          MIDI playback to process messages related to the progress of
'                          the playback. If no callback is desired, set this value to 0.
'     dwCallbackInstance-  User instance data passed to the callback. Set this value to 0.
'     dwFlags-Callback flag for opening the device. Set this value to 0.
'
' The function returns MMSYSERR_NOERROR if successful or one of the following error values:
'     MIDIERR_NODEVICE-       No MIDI port was found. This error occurs only when the mapper is opened.
'     MMSYSERR_ALLOCATED-     The specified resource is already allocated.
'     MMSYSERR_BADDEVICEID-   The specified device identifier is out of range.
'     MMSYSERR_INVALPARAM-    The specified pointer or structure is invalid.
'     MMSYSERR_NOMEM-         The system is unable to allocate or lock memory.

Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
' This function sends a short MIDI message to the specified MIDI output device. The function
' requires the handle to the MIDI output device and a message is packed into a doubleword
' value with the first byte of the message in the low-order byte. See the code sample for
' how to create this value.
'
' The function returns MMSYSERR_NOERROR if successful or one of the following error values:
'     MIDIERR_BADOPENMODE-  The application sent a message without a status byte to a stream handle.
'     MIDIERR_NOTREADY-     The hardware is busy with other data.
'     MMSYSERR_INVALHANDLE- The specified device handle is invalid.



'**************************************
'Windows API/Global Declarations for :*
'     Make Your Own *WAV* Player! *
'**************************************


Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_APPLICATION = &H80         '  look for application specific association
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_PURGE = &H40               '  purge non-static events for task
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_TYPE_MASK = &H170007
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside

Private Const INVALID_NOTE = -1     ' Code for keyboard keys that we don't handle

'*************************************************************
Public numDevices As Long       ' number of midi output devices
Public curDevice As Long        ' current midi device
Public hmidi As Long            ' midi output handle
Public rc As Long               ' return code
Public midimsg As Long          ' midi output message buffer
'*************************************************************

Public Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Boolean

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal _
    uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

Public Sub StartRecording()
    Dim r As Integer
    Dim Rate As Long: Rate = 44100
    Dim Channels As Integer: Channels = 2
    Dim Resolution As Integer: Resolution = 16
    Dim Alignment As Integer: Alignment = Channels * Resolution / 8
    r = mciSendString("open new type waveaudio alias recorderTemp", 0&, 0, 0)
    r = mciSendString("set recorderTemp time format ms", 0&, 0, 0)
    r = mciSendString("set recorderTemp time format bytes", 0&, 0, 0)
    r = mciSendString("set recorderTemp alignment " & CStr(CInt(Channels * Resolution / 8)) & " bitspersample " & CStr(Resolution) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(CInt(Channels * Resolution / 8) * Rate), 0&, 0, 0)
    r = mciSendString("record recorderTemp", 0&, 0, 0)
End Sub
Public Sub StopRecording()
    Dim r As Integer
    r = mciSendString("stop recorderTemp", 0&, 0, 0)
    r = mciSendString("save recorderTemp" & " " & AppPath & "check.wav", 0&, 0, 0)
    r = mciSendString("close recorderTemp", 0&, 0, 0)
End Sub

'Private Const ShortFreq = 600
'Private Const ShortDura = 50
'Private Const BetweenDura = 25
'Private Const LongFreq = 600
'Private Const LongDura = 200
'Private Const SpaceDura = 75
'
''Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Public Declare Sub Sleep Lib "kernel32" (ByVal msDuration As Long)
'
'Private Sub PlayMoseCode(ByVal text As String)
'    Do Until text = ""
'
'        Select Case LCase(Left(text, 1))
'            Case "a"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "b"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "c"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "d"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "e"
'                Beep ShortFreq, ShortDura
'            Case "f"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "g"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "h"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "i"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "j"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "k"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "l"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "m"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "n"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "o"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "p"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "q"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "r"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "s"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "t"
'                Beep LongFreq, LongDura
'            Case "u"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "v"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "w"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "x"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "y"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "z"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "0"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "1"
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "2"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "3"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'            Case "4"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "5"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "6"
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "7"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "8"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'            Case "9"
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case "."
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case ","
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'            Case "?"
'                Beep ShortFreq, ShortDura
'                Beep ShortFreq, ShortDura
'                Beep LongFreq, LongDura
'                Beep LongFreq, LongDura
'                Beep ShortFreq, ShortDura
'            Case " "
'                Sleep SpaceDura
'        End Select
'        text = Mid(text, 2)
'        Sleep BetweenDura
'        DoEvents
'    Loop
'
'End Sub

Public Sub Main()
    frmMain.Show
    
'    Do While True
'
'        PlayMoseCode "APB National Security Breech"
'
'    Loop
End Sub
'
'Public Sub Main()
'
'End Sub
Public Function GetWinDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(45, Chr(0))
    ret = GetWindowsDirectory(winDir, 45)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    GetWinDir = winDir
End Function

Public Function GetWinTempDir() As String

    On Error Resume Next
    Dim winDir As String
    Dim ret As Long
    winDir = String(45, Chr(0))
    ret = GetTempPath(45, winDir)
    If ret <> 16 Then
        If PathExists(GetWinDir() + "TEMP") Then
            winDir = GetWinDir() + "TEMP\"
        Else
            MkDir GetWinDir() + "TEMP"
            If PathExists(GetWinDir() + "TEMP") Then
                winDir = GetWinDir() + "TEMP\"
            Else
                winDir = ""
            End If
        End If
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWinTempDir = winDir
    If Err Then Err.Clear
    On Error GoTo 0

End Function


' Press the button and send midi start event
Public Sub StartNote(ByVal channel As Integer, ByVal note As Integer, ByVal volume As Integer)
'    If (Key(Index).Value = 1) Then
'        Exit Sub
'    End If
'    Key(Index).Value = 1
    midimsg = &H90 + (note * &H100) + (volume * &H10000) + channel
    midiOutShortMsg hmidi, midimsg
    DoEvents
End Sub

' Raise the button and send midi stop event
Public Sub StopNote(ByVal channel As Integer, ByVal note As Integer)
'    Key(Index).Value = 0
    midimsg = &H80 + (note * &H100) + channel
    midiOutShortMsg hmidi, midimsg
    DoEvents
End Sub





