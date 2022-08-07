VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl Mixer 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12195
   ScaleHeight     =   555
   ScaleWidth      =   12195
   Begin VB.CommandButton Command2 
      Height          =   192
      Left            =   60
      TabIndex        =   1
      Top             =   324
      Width           =   216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9912
      Top             =   144
   End
   Begin VB.CheckBox pMute 
      Height          =   312
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   28
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   24
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   20
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   12
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   8
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      Height          =   312
      Index           =   0
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      Height          =   312
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      Height          =   312
      Index           =   2
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      Height          =   312
      Index           =   3
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   312
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   312
      Index           =   5
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   312
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   312
      Index           =   7
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   9
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   10
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   11
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   13
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   14
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   15
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   16
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   17
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   18
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   19
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   21
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   22
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   23
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   25
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   26
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   27
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   29
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   30
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox pBeat 
      BackColor       =   &H00E0E0E0&
      Height          =   312
      Index           =   31
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "M"
      Top             =   0
      Width           =   315
   End
   Begin MSComDlg.CommonDialog pOpenWave 
      Index           =   0
      Left            =   10236
      Top             =   144
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.Slider pVolume 
      Height          =   216
      Index           =   0
      Left            =   10704
      TabIndex        =   35
      Top             =   108
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   370
      _Version        =   393216
      Max             =   127
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   675
      ScaleHeight     =   270
      ScaleWidth      =   9990
      TabIndex        =   38
      Top             =   15
      Visible         =   0   'False
      Width           =   9990
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   10035
   End
   Begin VB.Label Label2 
      Caption         =   "+"
      Height          =   180
      Left            =   11952
      TabIndex        =   40
      Top             =   -24
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      Height          =   180
      Left            =   10788
      TabIndex        =   39
      Top             =   -24
      Width           =   180
   End
   Begin VB.Label pResource 
      Caption         =   "MIDI Sequence"
      Height          =   192
      Index           =   0
      Left            =   348
      TabIndex        =   36
      Tag             =   "M"
      Top             =   324
      Width           =   10308
   End
   Begin VB.Label pNumber 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   60
      Width           =   435
   End
   Begin VB.Menu mnuPattern 
      Caption         =   "Pattern"
      Begin VB.Menu mnuLoadWave 
         Caption         =   "Mixer Properties"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Mixer"
      End
      Begin VB.Menu mnuClearNotes 
         Caption         =   "&Clear Notes"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Mixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pMixerType As String
Private pWaveBuffer As dxBuffers
Private pWaveLoaded As Boolean
Private pChannel As Integer
Private pBaseNote As Integer
Private pPattern As Long

Private pStartTempo As Single
Private pStartTiming As Single
Private pRecordings As Collection
Private pBeginning As Collection

Public Property Get StartTiming() As Single
    StartTiming = pStartTiming
End Property
Public Property Let StartTiming(ByVal newval As Single)
    pStartTiming = newval
End Property
Public Property Get StartTempo() As Single
   ' If Not IsRunningMode Then
        If pStartTempo = 0 Then pStartTempo = frmMain.Tempo.Value
        StartTempo = pStartTempo
  '  End If
End Property
Public Property Let StartTempo(ByVal newval As Single)
    pStartTempo = newval
End Property

Public Property Get Beginning() As Collection
    Set Beginning = pBeginning
End Property


Public Property Get Recordings() As Collection
    Set Recordings = pRecordings
End Property

Public Property Get MixerType() As String
    MixerType = pMixerType
End Property
Public Property Let MixerType(ByVal newval As String)
    Dim cnt As Long
    Select Case UCase(Trim(newval))
        Case "MIDI", "M"
            For cnt = 0 To 31
                pBeat(cnt).Visible = True
            Next
            Picture1.Visible = False
            Command1.Visible = False
            mnuClearNotes.Visible = False
            pMixerType = "M"
            Resource.Caption = "MIDI Sequencer - [Channel " & (pChannel + 1) & ", Note " & pBaseNote & "]"
        Case "WAVE", "W"
            For cnt = 0 To 31
                pBeat(cnt).Visible = True
            Next
            Picture1.Visible = False
            Command1.Visible = False
            mnuClearNotes.Visible = False
            pMixerType = "W"
            Resource.Caption = "Wave Sequencer"
        Case "PAINO", "P"
            pMixerType = "P"
            Resource.Caption = "Piano Recorder"
          For cnt = 0 To 31
                pBeat(cnt).Visible = False
            Next
            Picture1.Visible = True
            Command1.Visible = True
            mnuClearNotes.Visible = True
        Case "SYNTH", "S"
          For cnt = 0 To 31
                pBeat(cnt).Visible = False
            Next
            pMixerType = "S"
            Resource.Caption = "Synthesizer Draw"
            Picture1.Visible = False
            Command1.Visible = True
            mnuClearNotes.Visible = True
        Case Else
            For cnt = 0 To 31
                pBeat(cnt).Visible = True
            Next
            Picture1.Visible = False
            Command1.Visible = False
            mnuClearNotes.Visible = False
            pMixerType = "W"
            Resource.Caption = "Wave Sequencer"
    End Select
End Property

Public Property Get Mute()
    Set Mute = pMute(0)
End Property

Public Property Get Number()
    Set Number = pNumber
End Property
Public Property Let Pattern(ByVal newval As Long)
    pPattern = newval
End Property

Public Property Get Beat(ByVal Index As Integer)
    Set Beat = pBeat(Index)
End Property

Public Property Get Resource()
    Set Resource = pResource(0)
End Property

Public Property Get channel() As Integer
    channel = pChannel
End Property
Public Property Let channel(ByVal newval As Integer)
    pChannel = newval
End Property

Public Property Get BaseNote() As Integer
    BaseNote = pBaseNote
End Property
Public Property Let BaseNote(ByVal newval As Integer)
    pBaseNote = newval
End Property

Public Property Get volume()
    Set volume = pVolume(0)
End Property

Public Property Get OpenWave()
    Set OpenWave = pOpenWave(0)
End Property
Private Sub ButtonDown(ByVal Button As Integer)
    If Button = 2 Then
        UserControl.PopupMenu mnuPattern
    ElseIf Button = 2 Then
    
    End If
End Sub
Private Function GetFileName(ByVal Str As String) As String
    If InStrRev(Str, "\") > 0 Then
        GetFileName = Mid(Str, InStrRev(Str, "\") + 1)
    Else
        GetFileName = Str
    End If
End Function
    

Private Sub Command2_Click()
    UserControl.PopupMenu mnuPattern
End Sub

Private Sub mnuClearNotes_Click()
    If MsgBox("Are you sure you want to clear all piano notes?", vbYesNo + vbQuestion, "Clear Notes?") = vbYes Then
        ClearCollection pRecordings
        ClearCollection pBeginning
    End If
End Sub

Private Sub mnuLoadWave_Click()
    If MixerType = "W" Then
        On Error Resume Next
        If pOpenWave(0).FileName <> "" Then
            pOpenWave(0).InitDir = Left(pOpenWave(0).FileName, InStrRev(pOpenWave(0).FileName, "\") - 1)
        Else
            pOpenWave(0).InitDir = AppPath
        End If
        pOpenWave(0).CancelError = True
        pOpenWave(0).Filter = "Wave Audio (*.wav)|*.wav|All Files (*.*)|*.*"
        pOpenWave(0).FilterIndex = 1
        pOpenWave(0).DialogTitle = "Select an Audio File"
        pOpenWave(0).ShowOpen
        If Err = 0 Then
            LoadWaveFileName pOpenWave(0).FileName
        End If
    Else
        Load frmMIDIProp
        frmMIDIProp.Combo1.ListIndex = pChannel
        frmMIDIProp.Text1.text = pBaseNote
        frmMIDIProp.Show 1
        If frmMIDIProp.IsOK Then
            pChannel = frmMIDIProp.Combo1.ListIndex
            pBaseNote = CInt(frmMIDIProp.Text1.text)
            Resource.Caption = "MIDI Sequencer - [Channel " & (pChannel + 1) & ", Note " & pBaseNote & "]"
        End If
        Unload frmMIDIProp
    End If
End Sub

Public Sub RecordPiano()
    Do Until pBeginning.Count = 0
        pBeginning.Remove 1
    Loop

    Dim note
    For Each note In pRecordings
        pBeginning.Add note
    Next
    
    StartTempo = frmMain.Tempo.Value
    StartTiming = Timer
    
    'Picture1.Cls
    Timer1.Enabled = True

End Sub

Public Sub RedrawREcordings()

    Picture1.Cls

    Dim note As Variant
    Dim timat As Single
    Dim start As Boolean
    Dim key1 As Integer
    Dim colornote As Long

    For Each note In pBeginning
        timat = CSng(RemoveNextArg(note, ","))
        start = CBool(RemoveNextArg(note, ","))
        key1 = RemoveNextArg(note, ",")

        If pMixerType = "P" Then
            If start Then
    
                Select Case key1
                    Case 1, 3, 6, 8, 10, 13, 15, 18, 20, 22, 25, 27, 30, 32, 24
                        colornote = vbBlack
                    Case Else
                        colornote = vbWhite
                End Select
            Else
                colornote = &H808080
                            
            End If
            Picture1.Line (timat, Picture1.Height / 2)-(timat, Picture1.Height), colornote
            Picture1.Line (timat + Screen.TwipsPerPixelX, Picture1.Height / 2)-(timat + Screen.TwipsPerPixelX, Picture1.Height), colornote
        ElseIf pMixerType = "S" Then
        
        
        End If
    Next
    
End Sub

Private Sub Timer1_Timer()
    
    Dim timePos As Single
    Dim timeUnit As Single
    Dim atPos As Single
    Dim note As Variant
    Dim timat As Single
    Dim start As Boolean
    Dim key1 As Integer
    Dim volume2 As Integer
    Dim BaseNote2 As Integer
    Dim channel2 As Integer
         
            
    timePos = (((Timer - StartTiming) / (frmMain.AtBeat + 1)) * StartTempo)
    timeUnit = timePos / StartTempo
    atPos = ((Picture1.Width / 32) * frmMain.AtBeat) + timeUnit
    
    'Debug.Print atPos
    
    'Debug.Print "REUP " & pRecordings.Count & "  " & pBeginning.Count
    If Timer1.Tag > atPos Then
        Timer1.Tag = atPos
        Do Until pBeginning.Count = 0
            pRecordings.Add pBeginning(1)
            pBeginning.Remove 1
        Loop
    End If

    RedrawREcordings

        Picture1.Line (atPos, 0)-(atPos, Picture1.Height), &HFFFF&
        Picture1.Line (atPos + Screen.TwipsPerPixelX, 0)-(atPos + Screen.TwipsPerPixelX, Picture1.Height), &HFFFF&
    
    
    'If frmMain.Option1(0).Value Then

        'Dim tmp As New Collection
        
        If (pRecordings.Count > 0) Then
            'For Each note In pRecordings

'            Debug.Print "REUP " & pRecordings.Count & "  " & pBeginning.Count
            Do
                note = Trim(pRecordings(1))

                    timat = CSng(RemoveNextArg(note, ","))
                    'tmp.Add note
                    'Debug.Print NextArg(note, ",") & " " & atPos
                    If timat < atPos Then
                        start = CBool(RemoveNextArg(note, ","))
                        If start Then
                            key1 = RemoveNextArg(note, ",")
                            volume2 = RemoveNextArg(note, ",")
                            If volume > 0 And volume2 > 0 Then
                                volume2 = volume2 * (volume / 100)
                            Else
                                volume2 = 0
                            End If
                            BaseNote2 = RemoveNextArg(note, ",")
                            channel2 = RemoveNextArg(note, ",")
                            If pMute(0).Value = 1 Then
                                If GetActiveWindow = frmMain.hwnd Then
                                    If pMixerType = "P" Then
                                        frmMain.StartNoteEx key1, volume2, BaseNote2, channel2, False
                                    ElseIf pMixerType = "S" Then
                                    
                                    End If
                                End If
                            End If
                        Else
                            key1 = RemoveNextArg(note, ",")
                            volume2 = RemoveNextArg(note, ",")
                            BaseNote2 = RemoveNextArg(note, ",")
                            channel2 = RemoveNextArg(note, ",")
                            If pMixerType = "P" Then
                                frmMain.StopNoteEx key1, BaseNote2, channel2, False
                            ElseIf pMixerType = "S" Then
                            
                            End If
                        End If
                        pRecordings.Remove 1
                    Else
                        note = ""
                    End If

            'Next
            Loop Until (pRecordings.Count = 0) Or (note = "")

        End If
    'End If
End Sub

Public Sub MarkPianoNote(ByVal start As Boolean, ByVal key As Integer, ByVal vol As Integer, ByVal BaseNote As Integer, ByVal channel As Integer)
    Dim timePos As Single
    Dim timeUnit As Single
    Dim atPos As Single
    
    timePos = (((Timer - StartTiming) / (frmMain.AtBeat + 1)) * StartTempo)
    timeUnit = timePos / StartTempo
    atPos = ((Picture1.Width / 32) * frmMain.AtBeat) + timeUnit
    
    Dim cnt As Long
    cnt = 1
    If pRecordings.Count > 0 Then
        Do While atPos > CSng(NextArg(pRecordings(cnt), ",")) And (cnt < pRecordings.Count)
            cnt = cnt + 1
        Loop
        If cnt < pRecordings.Count - 1 Then
            pRecordings.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel, , cnt
        Else
            pRecordings.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel
        End If
    Else
        pRecordings.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel
    End If
    cnt = 1
    If pBeginning.Count > 0 Then
        Do While atPos > CSng(NextArg(pBeginning(cnt), ",")) And (cnt < pBeginning.Count)
            cnt = cnt + 1
        Loop
        If cnt < pBeginning.Count - 1 Then
            pBeginning.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel, , cnt
        Else
            pBeginning.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel
        End If
    Else
        pBeginning.Add atPos & "," & start & "," & key & "," & vol & "," & BaseNote & "," & channel
    End If

End Sub

Public Sub StopPiano()
    Timer1.Enabled = False
    Do Until pRecordings.Count = 0
        pRecordings.Remove 1
    Loop
    If pRecordings.Count = 0 And pBeginning.Count > 0 Then
        Dim note
        For Each note In pBeginning
            pRecordings.Add note
        Next
    End If

End Sub
Private Sub MoveNotes()
    If pRecordings.Count = 0 And pBeginning.Count > 0 Then
        Do Until pBeginning.Count = 0
            pRecordings.Add pBeginning(1)
            pBeginning.Remove 1
        Loop
    End If
    
    Do Until pBeginning.Count = 0
        pBeginning.Remove 1
    Loop

    Dim note
    For Each note In pRecordings
        pBeginning.Add note
    Next
End Sub
Public Sub PlayPiano()

    Do Until pBeginning.Count = 0
        pBeginning.Remove 1
    Loop

    Dim note
    For Each note In pRecordings
        pBeginning.Add note
    Next
    
    StartTempo = frmMain.Tempo.Value
    StartTiming = Timer
    
    'Picture1.Cls
    Timer1.Enabled = True
End Sub

Public Sub LoadWaveFileName(ByVal FileName As String)

    DX7LoadSound FileName
    pWaveLoaded = True
    Resource.Caption = "Wave Sequencer - [" & Replace(FileName, AppPath, "") & "]"

End Sub

Public Sub PlaySound(ByVal LoopIt As Byte)
    If pWaveLoaded And pWaveBuffer.FileName <> "" Then
        DX7LoadSound pWaveBuffer.FileName
        'pWaveBuffer.Buffer.SetPan -10000
        'SB(Buffer).Buffer.SetPan 10000
        'SB(Buffer).Buffer.SetPan (100 * PanValue) - 5000
        VolumeLevel (volume.Value / volume.Max) * 100
        pWaveBuffer.Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0
    End If
End Sub
 
Public Sub DX7LoadSound(FileName As String)
      Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
      Dim waveFormat As WAVEFORMATEX 'what sort of buffer to create
      
      bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
      Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC 'These settings should do for almost any app....
      
      waveFormat.nFormatTag = WAVE_FORMAT_PCM
      waveFormat.nChannels = 2    '2 channels
      waveFormat.lSamplesPerSec = 22050
      waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
      waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
      waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
      On Error GoTo Continue
      Set pWaveBuffer.Buffer = directSound.CreateSoundBufferFromFile(Trim(FileName), bufferDesc)
      pWaveBuffer.FileName = Trim(FileName)
      
      Exit Sub
Continue:
    '  MsgBox "Error can't find file: " & FileName
End Sub

Public Sub VolumeLevel(volume As Byte)
    If volume > 0 Then ' stop division by 0
        pWaveBuffer.Buffer.SetVolume (60 * volume) - 6000
    Else
        pWaveBuffer.Buffer.SetVolume -6000
    End If
End Sub
Private Sub mnuRemove_Click()
    If MsgBox("Are you sure you want to REMOVE this mixer?", vbQuestion + vbYesNo, "Remove?") = vbYes Then
        UserControl.Parent.RemoveMixer CInt(pNumber.Caption)
    End If
End Sub

Private Sub pBeat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub pBeat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If MixerType = "M" Then
            Select Case Beat(Index).Caption
                Case " ", ""
                    Beat(Index).Caption = "S"
                    Beat(Index).Value = 1
                Case "S"
                    Beat(Index).Caption = "L"
                    Beat(Index).Value = 1
                Case "L"
                    Beat(Index).Caption = " "
                    Beat(Index).Value = 0
            End Select
        ElseIf MixerType = "W" Then
            Select Case Beat(Index).Caption
                Case " ", ""
                    Beat(Index).Caption = "W"
                    Beat(Index).Value = 1
                Case "W"
                    Beat(Index).Caption = " "
                    Beat(Index).Value = 0
            End Select
        End If
    ElseIf Button = 2 Then
        Beat(Index).Caption = " "
        Beat(Index).Value = 0
    End If
End Sub

Private Sub pMute_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub pNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub pResource_DblClick(Index As Integer)
    mnuLoadWave_Click
End Sub

Private Sub pResource_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub pVolume_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub UserControl_Initialize()
    Set pRecordings = New Collection
    Set pBeginning = New Collection
    Timer1.Tag = 0
    pWaveLoaded = False
    pChannel = 9
    pBaseNote = 35
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDown Button
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 12195
    UserControl.Height = 555
End Sub

Private Sub UserControl_Show()
RedrawREcordings
End Sub

Private Sub UserControl_Terminate()
    Do Until pBeginning.Count = 0
        pBeginning.Remove 1
    Loop
    Set pBeginning = Nothing
    Do Until pRecordings.Count = 0
        pRecordings.Remove 1
    Loop
    Set pRecordings = Nothing
End Sub
