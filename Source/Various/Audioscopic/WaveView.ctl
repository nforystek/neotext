VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl WaveView 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   7710
   Begin VB.PictureBox PicDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "WaveView.ctx":0000
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   720
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   1455
      ScaleWidth      =   2775
      TabIndex        =   7
      Top             =   420
      Width           =   2775
      Begin VB.Line LeftMarker 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   60
         X2              =   60
         Y1              =   360
         Y2              =   1080
      End
      Begin VB.Line RightMarker 
         BorderColor     =   &H00404040&
         X1              =   360
         X2              =   360
         Y1              =   360
         Y2              =   1020
      End
      Begin VB.Line PlayCursor 
         BorderColor     =   &H00E0E0E0&
         Visible         =   0   'False
         X1              =   1200
         X2              =   1200
         Y1              =   300
         Y2              =   900
      End
   End
   Begin VB.CommandButton OpenButton 
      Appearance      =   0  'Flat
      DisabledPicture =   "WaveView.ctx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      MaskColor       =   &H00FF00FF&
      Picture         =   "WaveView.ctx":0914
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton StopButton 
      DisabledPicture =   "WaveView.ctx":0DE6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      MaskColor       =   &H00FF00FF&
      Picture         =   "WaveView.ctx":12B8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2820
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.PictureBox PicRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2295
      ScaleHeight     =   1395
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton PlayButton 
      DisabledPicture =   "WaveView.ctx":178A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "WaveView.ctx":1C5C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2340
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.PictureBox PicSlidebar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      ScaleHeight     =   375
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
      Begin VB.Shape ViewSlider 
         FillColor       =   &H8000000F&
         Height          =   315
         Left            =   120
         Top             =   60
         Width           =   1635
      End
   End
   Begin VB.PictureBox PicVolscale 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3900
      ScaleHeight     =   735
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   2460
      Width           =   315
   End
   Begin VB.Timer Blinkers 
      Interval        =   20
      Left            =   900
      Top             =   2100
   End
   Begin VB.CommandButton ResetButton 
      Appearance      =   0  'Flat
      DisabledPicture =   "WaveView.ctx":212E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      MaskColor       =   &H00FF00FF&
      Picture         =   "WaveView.ctx":2600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin MSScriptControlCtl.ScriptControl FilterScripts 
      Left            =   6060
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin MSComDlg.CommonDialog BrowseForWave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image DragIcon 
      Height          =   480
      Left            =   2940
      Picture         =   "WaveView.ctx":2AD2
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "WaveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declare Color Constants
Private Const cBlue = &HFF0000
Private Const cRed = &HFF
Private Const cGreen = &HFF00
Private Const cYellow = &HFFFF
Private Const cTeal = &H64FF64
Private Const cPurple = &HFF007D
Private Const cGray = &H7D7D7D
Private Const cOrange = &H96FF

'' Public Declarations
''################################
'Private wFileHeader As FileHeader
'Private wFormat As FormatChunk
'Private wChunk As ChunkHeader
'Private I_Data() As Integer
''################################


Private dBitsPerTwip As Single
Private FileLoaded As String
Private peak As Single, Dip As Single

Private SoundPlaying As Long
Private CursorAdjust As Integer

Private TotalTime As Single
Private ChannelTap As Single
Private DotsOnly As Boolean
Private MRatio As Single
Private DBZoom As Single

Private AudioFile As AudioFile2

Private LastX As Single
Private LastY As Single
Private LastBtn As Integer
Private LaststartMS As Single
Private LaststopMS As Single

Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpszFormat$) As Integer
Private MyFormat As Integer

Public Event Click()

Public Property Get FileName() As String
    FileName = FileLoaded
End Property

Public Property Get WaveMarkerStartMS() As Single
    LineRestrictions
    WaveMarkerStartMS = (TotalTime * ((LeftMarker.X1 - Screen.TwipsPerPixelX) / PicDisplay.Width))
End Property
Public Property Get WaveMarkerStopMS() As Single
    LineRestrictions
    WaveMarkerStopMS = (TotalTime * ((RightMarker.X1 + Screen.TwipsPerPixelX) / PicDisplay.Width))
End Property

Public Property Get WaveMarkerPlayMS() As Single
    LineRestrictions
    WaveMarkerPlayMS = (TotalTime * ((PlayCursor.X1 + Screen.TwipsPerPixelX) / PicDisplay.Width))
End Property
Public Property Get WaveLengthMS() As Single
    WaveLengthMS = TotalTime
End Property

Private Function ScrollOver() As Control
    Dim pt As POINTAPI

    Dim rct As RECT
    
    GetCursorPos pt
    
    GetWindowRect PicSlidebar.hwnd, rct

    If ((pt.X > rct.Left) And (pt.X < rct.Right)) And ((pt.Y > rct.Top) And (pt.Y < rct.Bottom)) Then
        Set ScrollOver = PicSlidebar
    Else
        GetWindowRect PicVolscale.hwnd, rct
        If (((pt.X > rct.Left) And (pt.X < rct.Right)) And ((pt.Y > rct.Top) And (pt.Y < rct.Bottom))) Then
            Set ScrollOver = PicVolscale
        Else
            Set ScrollOver = PicDisplay
        End If
    End If

End Function

Friend Sub ScrollUp()

    On Error GoTo catcherr
    
    Dim useScroller As Control
    Set useScroller = ScrollOver

    If Not useScroller Is Nothing Then
        
        Select Case useScroller.Name
                
            Case "PicSlidebar"
                PicSlidebar.Tag = (PicSlidebar.Width / 100)
                MoveViewPort (PicSlidebar.Tag * 2)
                PicSlidebar.Tag = 0
            Case "PicVolscale"
'                If DBZoom < 10 Then
'                    DBZoom = DBZoom + 0.2
'                End If
'                PicDisplay.Height = ((UserControl.Height - PicSlidebar.Height) * DBZoom)
'                PicDisplay.Top = (((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))
'
'                PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height
'
'                RenderTimeLine
'                RenderFileInfo vbBlack
'                RenderCenters
'                CursorAdjust = 0
                
            Case Else

                Dim StartMS As Single
                Dim stopMS As Single

                StartMS = WaveMarkerStartMS
                stopMS = WaveMarkerStopMS
            
                PicDisplay.Width = PicDisplay.Width + (PicDisplay.Width / (Screen.Width / Screen.Height))

                PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height
      
                ViewSlider.Width = (PicSlidebar.Width * ((UserControl.Width - PicVolscale.Width) / PicDisplay.Width))

                If PicDisplay.Left < 0 Then
                    ViewSlider.Left = (PicSlidebar.Width * (-PicDisplay.Left / PicDisplay.Width))
                End If

                If TotalTime > 0 Then
                    LeftMarker.X1 = (PicDisplay.Width * (StartMS / TotalTime))
                    LeftMarker.X2 = LeftMarker.X1
                    RightMarker.X1 = (PicDisplay.Width * (stopMS / TotalTime))
                    RightMarker.X2 = RightMarker.X1
                End If
                
                RenderTimeLine
                RenderFileInfo vbBlack
                RenderCenters
                
                LineRestrictions
                CursorAdjust = 0
                PicDisplay.SetFocus
        End Select

    End If

    Set useScroller = Nothing
    
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub
Friend Sub ScrollDown()

    On Error GoTo catcherr
    
    Dim useScroller As Control
    Set useScroller = ScrollOver

    If Not useScroller Is Nothing Then
        
        Select Case useScroller.Name
            Case "PicSlidebar"
                PicSlidebar.Tag = ((PicSlidebar.Width / 100) * 2)
                MoveViewPort (PicSlidebar.Tag / 2)
                PicSlidebar.Tag = 0
            Case "PicVolscale"
'                If DBZoom > 0.2 Then
'                    DBZoom = DBZoom - 0.2
'                End If
'
'                PicDisplay.Height = ((UserControl.Height - PicSlidebar.Height) * DBZoom)
'                PicDisplay.Top = (((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))
'
'                PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height
'
'                RenderTimeLine
'                RenderFileInfo vbBlack
'                RenderCenters
'                CursorAdjust = 0
                
            Case Else

                Dim StartMS As Single
                Dim stopMS As Single

                StartMS = WaveMarkerStartMS
                stopMS = WaveMarkerStopMS
                
                If (PicDisplay.Width - (PicDisplay.Width / (Screen.Width / Screen.Height))) >= (17 * Screen.TwipsPerPixelX) Then
                    PicDisplay.Width = (PicDisplay.Width - (PicDisplay.Width / (Screen.Width / Screen.Height)))
                Else
                    PicDisplay.Width = (17 * Screen.TwipsPerPixelX)
                End If

                PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height
 
                ViewSlider.Width = (PicSlidebar.Width * ((UserControl.Width - PicVolscale.Width) / PicDisplay.Width))
                
                If PicDisplay.Width < (UserControl.Width - PicVolscale.Width) Then
                    PicDisplay.Left = 0
                    ViewSlider.Left = 0
                Else
                    If PicDisplay.Left < 0 Then
                        ViewSlider.Left = (PicSlidebar.Width * (-PicDisplay.Left / PicDisplay.Width))
                    End If
                    If (PicDisplay.Width + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width) And ((PicDisplay.Left < 0) And (PicDisplay.Width > (UserControl.Width - PicVolscale.Width))) Then
                        PicDisplay.Left = -(PicDisplay.Width - (UserControl.Width - PicVolscale.Width))
                        ViewSlider.Left = PicSlidebar.Width - ViewSlider.Width
                    End If
                End If

                If TotalTime > 0 Then
                    LeftMarker.X1 = (PicDisplay.Width * (StartMS / TotalTime))
                    LeftMarker.X2 = LeftMarker.X1
                    RightMarker.X1 = (PicDisplay.Width * (stopMS / TotalTime))
                    RightMarker.X2 = RightMarker.X1
                End If

                RenderTimeLine
                RenderFileInfo vbBlack
                RenderCenters
                LineRestrictions
                CursorAdjust = 0
                PicDisplay.SetFocus
        End Select

    End If

    Set useScroller = Nothing
    
    Exit Sub
catcherr:
    'MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Private Sub MoveViewPort(ByVal X As Single)

    RenderFileInfo vbWhite
    If PicSlidebar.Tag > 0 Then
    
        If Not (ViewSlider.Left + (X - PicSlidebar.Tag) >= 0) Then
            PicSlidebar.Tag = -ViewSlider.Left
        ElseIf Not (ViewSlider.Left + (X - PicSlidebar.Tag) + ViewSlider.Width < PicSlidebar.Width) Then
            PicSlidebar.Tag = PicSlidebar.Width - (ViewSlider.Left + ViewSlider.Width)
        Else
            PicSlidebar.Tag = (X - PicSlidebar.Tag)
        End If
        
        PicDisplay.Left = PicDisplay.Left - (PicDisplay.Width * (PicSlidebar.Tag / PicSlidebar.Width))
        ViewSlider.Left = ViewSlider.Left + PicSlidebar.Tag
    
        RenderFileInfo vbBlack
        CursorAdjust = 0
        PicDisplay.SetFocus
    End If
End Sub

Private Sub LineRestrictions()
    
    If LeftMarker.X1 > RightMarker.X1 Then
        RightMarker.X2 = LeftMarker.X1
        LeftMarker.X1 = RightMarker.X1
        RightMarker.X1 = RightMarker.X2
        RightMarker.X2 = RightMarker.X1
        LeftMarker.X2 = LeftMarker.X1
    End If

    If LeftMarker.X1 <= 0 Then
        LeftMarker.X1 = Screen.TwipsPerPixelX
        LeftMarker.X2 = Screen.TwipsPerPixelX
    End If

    If RightMarker.X1 >= PicDisplay.Width Then
        RightMarker.X1 = PicDisplay.Width - Screen.TwipsPerPixelX
        RightMarker.X2 = PicDisplay.Width - Screen.TwipsPerPixelX
    End If

End Sub

Public Property Get WaveData(Optional ByVal StartTimeMS As Single = 0, Optional ByVal DurationTimeMS As Single = -1) As Byte()

    On Error GoTo catcherr
    
    If StartTimeMS > TotalTime Or StartTimeMS < 0 Then
        Err.Raise 8, "WaveData", "Invalid or exceeding start time."
    Else
        If DurationTimeMS = -1 Then DurationTimeMS = TotalTime - StartTimeMS
        If StartTimeMS + DurationTimeMS <= TotalTime Then
            WaveData = WaveAudioPartial(AudioFile, StartTimeMS, DurationTimeMS)
            
        Else
            Err.Raise 9, "WaveData", "Invalid or exceeding  duration time."
        End If
    End If
    
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Property

'Public Sub WaveMixIn(ByRef SourceWaveView As WaveView, Optional ByVal DestStartTimeMS As Single = 0, Optional ByVal SourceStartTimeMS As Single = 0, Optional ByVal SourceDurationTimeMS As Single = -1)
'
'    On Error GoTo catcherr
'
'    If ((DestStartTimeMS > TotalTime) Or (DestStartTimeMS < 0)) Or ((SourceStartTimeMS > SourceWaveView.WaveLengthMS) Or (SourceStartTimeMS < 0)) Then
'        Err.Raise 9, "WaveData", "Invalid or exceeding start time."
'    Else
'        If SourceDurationTimeMS = -1 Then SourceDurationTimeMS = SourceWaveView.WaveLengthMS - SourceStartTimeMS
'        If (DestStartTimeMS + SourceDurationTimeMS <= TotalTime) And (SourceStartTimeMS + SourceDurationTimeMS <= SourceWaveView.WaveLengthMS) Then
'            WaveCombine WaveData, DestStartTimeMS, SourceWaveView.WaveData(SourceStartTimeMS, SourceDurationTimeMS), SourceDurationTimeMS
'
'        Else
'            Err.Raise 9, "WaveData", "Invalid or exceeding duration time."
'        End If
'    End If
'
'    Exit Sub
'catcherr:
'    MsgBox Err.Description, vbCritical, "An error occured"
'    Err.Clear
'End Sub

Public Sub WaveStop()

    On Error GoTo catcherr
    
    If SoundPlaying Then
        PlayWaveSound SilentBytes
        
        SoundPlaying = False
        PlayButton.Enabled = True
        StopButton.Enabled = False
    End If

    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

'Public Sub WavePlay()
'
'    On Error GoTo catcherr
'
'    If FileLoaded <> "" And SoundPlay = False Then
'
'        Dim i As Single
'        Dim lTiming As Single
'        Dim startMS As Single
'        Dim stopMS As Single
'
'        Dim X As Single, Y As Single, y1 As Single, h As Single, w As Single
'
'
'        Dim CurrentTrack As Integer
'
'        LineRestrictions
'
'        SoundPlay = True
'        PlayButton.Enabled = False
'        StopButton.Enabled = True
'
'        PlayCursor.X1 = LeftMarker.X1 - Screen.TwipsPerPixelX
'        PlayCursor.X2 = PlayCursor.X1
'
'        startMS = WaveMarkerStartMS
'
'        Dim sndBytes() As Byte
'
'        sndBytes = WaveAudioPartial(AudioFile, startMS, WaveMarkerStopMS - startMS)
'
'        startMS = PlayCursor.X1
'        stopMS = RightMarker.X1 + Screen.TwipsPerPixelX
'
'        If PlayCursor.X1 + PicDisplay.Left >= 0 And PlayCursor.X1 + PicDisplay.Left < (UserControl.Width - PicVolscale.Width) Then
'            If ((LeftMarker.X1 + PicDisplay.Left >= 0) And ((LeftMarker.X1 + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width))) And _
'                ((RightMarker.X1 + PicDisplay.Left >= 0) And ((RightMarker.X1 + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width))) Then
'                CursorAdjust = 0
'            Else
'                CursorAdjust = 1
'            End If
'        Else
'            CursorAdjust = 0
'        End If
'
'        Open "C:\Development\Projects\Audioscopic\partial.wav" For Binary As #1
'        Put 1, , sndBytes
'        Close 1
'        'PlayWaveSound sndBytes
'        lTiming = Timer
'
'        PlayCursor.Visible = True
'
'
'        Do
'
'            If (Timer - lTiming) > 0 And TotalTime > 0 Then
'                PlayCursor.X1 = startMS + (PicDisplay.Width * (((Timer - lTiming) * 1000) / TotalTime))
'            End If
'
'            If CursorAdjust = 1 Then
'                If PlayCursor.X1 + PicDisplay.Left >= ((UserControl.Width - PicVolscale.Width) / 2) - ((17 * Screen.TwipsPerPixelX) / 2) And _
'                    PlayCursor.X1 + PicDisplay.Left < ((UserControl.Width - PicVolscale.Width) / 2) + ((17 * Screen.TwipsPerPixelX) / 2) Then
'
'                    If (PicDisplay.Width + (PicDisplay.Left - (PlayCursor.X1 - PlayCursor.X2))) >= (UserControl.Width - PicVolscale.Width) Then
'
'                        PicDisplay.Left = PicDisplay.Left - (PlayCursor.X1 - PlayCursor.X2)
'                        ViewSlider.Left = (PicSlidebar.Width * (-PicDisplay.Left / PicDisplay.Width))
'                    End If
'
'                End If
'            End If
'
'            PlayCursor.X2 = PlayCursor.X1
'
'            h = (((Abs(Peak) + Abs(Dip)) / (Abs(-32768) + Abs(32767))) / PicDisplay.Height)
'
'            Dim j As Single
'
'            For CurrentTrack = 1 To AudioFile.Infos(1).wNumberOfChannels
'
'                Y = Round((AudioFile.Infos(1).lBytesPerSecond * WaveMarkerPlayMS) / 1000) / AudioFile.Infos(1).wNumberOfChannels
'                y1 = Abs(AudioFile.Datas(1).wWave(Round((AudioFile.Infos(1).lBytesPerSecond * (WaveMarkerPlayMS / 1000)) / AudioFile.Infos(1).wNumberOfChannels)))  '+ (CurrentTrack - 1)))
'
'
''                w = PicRender.Height / (AudioFile.Infos(1).wNumberOfChannels + 1)
''                y1 = ((w + (AudioFile.Datas(1).wWave(i + CurrentTrack) * MRatio)) * CurrentTrack)
'
'              '  w = PicRender.Height / (AudioFile.Infos(1).wNumberOfChannels + 1) '(PicSlidebar.Width / AudioFile.Infos(1).wNumberOfChannels)
'
'
'               ' y1 = ((w + (AudioFile.Datas(1).wWave(i + CurrentTrack) * MRatio)) * CurrentTrack)
'
'               ' y1 = Abs(AudioFile.Datas(1).wWave(Round((AudioFile.Infos(1).lBytesPerSecond * (WaveMarkerPlayMS / 1000)) / AudioFile.Infos(1).wNumberOfChannels)))  '+ (CurrentTrack - 1)))
'
'            '    j = Round((AudioFile.Infos(1).lBytesPerSecond * WaveMarkerPlayMS) / 1000) / AudioFile.Infos(1).wNumberOfChannels
'
'            '    Debug.Print Round((AudioFile.Infos(1).lBytesPerSecond * WaveMarkerPlayMS) / 1000) / AudioFile.Infos(1).wNumberOfChannels
'
'            '    y1 = AudioFile.Datas(1).wWave(j)
'
'                On Error Resume Next
'
'                '351624
'
'                'Print 351972 - 351624
'                '348
'              '  y1 = Abs(AudioFile.Datas(1).wWave(j))  '+ (CurrentTrack - 1)))
'
'                If Err Then
'                    Debug.Print "Error: " & Err.Description
'                    Err.Clear
'                End If
'                h = (y1 / (Abs(Peak) + Abs(Dip)))
'
'                y1 = (h * (Abs(-32768) + Abs(32767)))
'
'                h = ((-100 - Round(h * 100)) + 200) '
'
'                If h > 75 Then
'                    PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(vbYellow, vbGreen, h), BF
'                ElseIf h > 50 Then
'                    PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(&H80C0FF, vbYellow, h), BF
'                Else
'                    PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(vbRed, &H80C0FF, h), BF
'                End If
'
'                PicVolscale.Line ((w * (CurrentTrack - 1)), (PicDisplay.Height * (h / 100)))-((w * CurrentTrack), 0), vbWhite, BF
'
'                PicVolscale.Line ((w * (CurrentTrack - 1)), (PicDisplay.Height * (h / 100)))-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), vbBlack
'
'                PicVolscale.Refresh
'
'            Next
'
'            y1 = 0
'
'            DoEvents
'
'            If SoundPlay = False Then Exit Do
'
'        Loop While PlayCursor.X1 < stopMS
'
'        RenderVolScale
'
'        PlayCursor.Visible = False
'
'        PlayButton.Enabled = True
'        StopButton.Enabled = False
'
'        SoundPlay = False
'
'    End If
'
'    Exit Sub
'catcherr:
'    MsgBox Err.Description, vbCritical, "An error occured"
'    Err.Clear
'End Sub
'
'Public Sub WaveFromData(ByRef wData() As Byte)
'
'    On Error GoTo catcherr
'
'    WaveResetRecord AudioFile
'
'    AudioFile = WaveBytesAsAudio(wData)
'
'    WaveInitial
'
'    Exit Sub
'catcherr:
'    MsgBox Err.Description, vbCritical, "An error occured"
'    Err.Clear
'End Sub
'
Public Sub WavePlay()

    On Error GoTo catcherr
    
    If FileLoaded <> "" And Not SoundPlaying Then
        SoundPlaying = True
        Dim i As Single
        Dim lTiming As Single
        Dim StartMS As Single
        Dim stopMS As Single
        
        Dim X As Single, Y As Single, y1 As Single, h As Single, w As Single
        
        
        Dim CurrentTrack As Integer
        
        LineRestrictions

        
        PlayButton.Enabled = False
        StopButton.Enabled = True

        PlayCursor.X1 = LeftMarker.X1 - Screen.TwipsPerPixelX
        PlayCursor.X2 = PlayCursor.X1
        
        StartMS = WaveMarkerStartMS
        
        PlayingBytes = WaveAudioPartial(AudioFile, StartMS, WaveMarkerStopMS - StartMS)
        
        StartMS = PlayCursor.X1
        stopMS = RightMarker.X1 + Screen.TwipsPerPixelX
        
        If PlayCursor.X1 + PicDisplay.Left >= 0 And PlayCursor.X1 + PicDisplay.Left < (UserControl.Width - PicVolscale.Width) Then
            If ((LeftMarker.X1 + PicDisplay.Left >= 0) And ((LeftMarker.X1 + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width))) And _
                ((RightMarker.X1 + PicDisplay.Left >= 0) And ((RightMarker.X1 + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width))) Then
                CursorAdjust = 0
            Else
                CursorAdjust = 1
            End If
        Else
            CursorAdjust = 0
        End If
        
        PlayWaveSound PlayingBytes
        
        lTiming = Timer
        
        PlayCursor.Visible = True
        Dim nx As Single
        
        Do
            
            If (Timer - lTiming) > 0 And TotalTime > 0 Then
                nx = StartMS + (PicDisplay.Width * (((Timer - lTiming) * 1000) / TotalTime))
            Else
                nx = 0
            End If
        
            If nx < stopMS And nx > 0 Then
                PlayCursor.X1 = nx
                If CursorAdjust = 1 Then
                    If PlayCursor.X1 + PicDisplay.Left >= ((UserControl.Width - PicVolscale.Width) / 2) - ((17 * Screen.TwipsPerPixelX) / 2) And _
                        PlayCursor.X1 + PicDisplay.Left < ((UserControl.Width - PicVolscale.Width) / 2) + ((17 * Screen.TwipsPerPixelX) / 2) Then
                        
                        If (PicDisplay.Width + (PicDisplay.Left - (PlayCursor.X1 - PlayCursor.X2))) >= (UserControl.Width - PicVolscale.Width) Then
                        
                            PicDisplay.Left = PicDisplay.Left - (PlayCursor.X1 - PlayCursor.X2)
                            ViewSlider.Left = (PicSlidebar.Width * (-PicDisplay.Left / PicDisplay.Width))
                        End If
                        
                    End If
                End If
            
                PlayCursor.X2 = PlayCursor.X1
    
                h = (((Abs(peak) + Abs(Dip)) / (Abs(-32768) + Abs(32767))) / PicDisplay.Height)
                   
                For CurrentTrack = 1 To AudioFile.FMTChunk.nChannels
    
                    w = (PicSlidebar.Width / AudioFile.FMTChunk.nChannels)
    
                    Y = Round((AudioFile.FMTChunk.nAvgBytesPerSec * WaveMarkerPlayMS) / 1000) / AudioFile.FMTChunk.nBlockAlign
                    If Y < WaveSamples(AudioFile) Then
                        
                        If AudioFile.FMTChunk.wBitsPerSample = 8 Then
                            y1 = Abs(AudioFile.WaveData8bit.ChannelData(CurrentTrack, Y))  '+ (CurrentTrack - 1)))
                        ElseIf AudioFile.FMTChunk.wBitsPerSample = 16 Then
                            
                            y1 = Abs(AudioFile.WaveData16bit.ChannelData(CurrentTrack, Y)) ' * AudioFile.FMTChunk.nChannels ' + (CurrentTrack - 1)))
                        End If
        
                        h = (y1 / (Abs(peak) + Abs(Dip)))
                        
                        y1 = (h * (Abs(-32768) + Abs(32767)))
        
                        h = ((-100 - Round(h * 100)) + 200)
                        
                        If h > 75 Then
                            PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(vbYellow, vbGreen, h), BF
                        ElseIf h > 50 Then
                            PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(&H80C0FF, vbYellow, h), BF
                        Else
                            PicVolscale.Line ((w * (CurrentTrack - 1)), PicDisplay.Height)-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), Blend(vbRed, &H80C0FF, h), BF
                        End If
                        
                        PicVolscale.Line ((w * (CurrentTrack - 1)), (PicDisplay.Height * (h / 100)))-((w * CurrentTrack), 0), vbWhite, BF
        
                        PicVolscale.Line ((w * (CurrentTrack - 1)), (PicDisplay.Height * (h / 100)))-((w * CurrentTrack), (PicDisplay.Height * (h / 100))), vbBlack
                        
                        PicVolscale.Refresh
                    End If
    
                Next
    
                y1 = 0
    
                DoEvents
            End If
            
            If SoundPlaying = False Then Exit Do
        
        Loop While nx < stopMS
        
        RenderVolScale
        
        PlayCursor.Visible = False

        PlayButton.Enabled = True
        StopButton.Enabled = False

        SoundPlaying = False

    End If

    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Public Sub WaveFromData(ByRef wData() As Byte)

    On Error GoTo catcherr
    
    WaveResetRecord AudioFile
    
    AudioFile = WaveBytesAsAudio(wData)
    
    WaveInitial

    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Public Sub WaveSaveAs(ByVal FileName As String)

    On Error GoTo catcherr
    
    WaveSaveToFile AudioFile, FileName
    FileLoaded = FileName
  
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Public Sub WaveFromFile(ByVal FileName As String)

    On Error GoTo catcherr

    If FileLoaded <> "" Then
        WaveResetRecord AudioFile
        FileLoaded = ""
    End If
    
    AudioFile = WaveLoadFromFile(FileName)
    
    If AudioFile.RiffHdr.DataType = "WAVE" Then
    
        WaveInitial

        FileLoaded = FileName
    End If
      '      DrawWave16 AudioFile, PicDisplay
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub
Private Sub RenderTimeLine()
    If FileLoaded <> "" Then
        Dim i As Single
        Dim X As Single
        
        i = ((PicDisplay.Width / (((LengthOfData(AudioFile) / 2) / AudioFile.FMTChunk.nChannels) / AudioFile.FMTChunk.nSamplesPerSec)) * 0.1)
        If i > 0 Then
            If PicDisplay.Width / i > Screen.TwipsPerPixelX Then
                For X = 1 To PicDisplay.Width Step i
                    PicDisplay.Line (X, (UserControl.Height - PicSlidebar.Height) - ((UserControl.Height - PicSlidebar.Height) * IIf((X / i) Mod 10 = 0, 0.09, IIf((X / i) Mod 5 = 0, 0.06, 0.03))))-(X, (UserControl.Height - PicSlidebar.Height)), vbBlack
                    UserControl.Line (X, (UserControl.Height - PicSlidebar.Height) - ((UserControl.Height - PicSlidebar.Height) * IIf((X / i) Mod 10 = 0, 0.09, IIf((X / i) Mod 5 = 0, 0.06, 0.03))))-(X, (UserControl.Height - PicSlidebar.Height)), vbBlack
                
                Next
            End If
        End If
    End If
End Sub

Private Sub RenderCenters()
    If FileLoaded <> "" Then
        Dim i As Long
        Dim Y As Single
        If (PicDisplay.Width < (UserControl.Width - PicVolscale.Width)) Then
            For i = 1 To AudioFile.FMTChunk.nChannels
                Y = ((UserControl.Height - PicSlidebar.Height) / (AudioFile.FMTChunk.nChannels + 1) * i) - _
                    (((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))

                UserControl.Line (0, Y)-(UserControl.Width - PicVolscale.Width, Y), GetColorShade(i, 0)
            Next
        End If
    End If
End Sub
Private Sub RenderStatus(ByVal Msg As String, Optional ByVal NoCls As Boolean = False)
    If Not NoCls Then PicDisplay.Cls
    
    PicDisplay.CurrentY = 0
    PicDisplay.CurrentX = -PicDisplay.Left + (Screen.TwipsPerPixelX * 3)
    PicDisplay.CurrentY = -(((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))
    PicDisplay.ForeColor = vbBlack
    PicDisplay.Print Msg
  '  PicDisplay.Refresh
    
    UserControl.Cls
    UserControl.ForeColor = vbBlack
    UserControl.CurrentX = 0
    UserControl.CurrentY = (Screen.TwipsPerPixelX * 3)
    UserControl.Print Msg
   ' UserControl.Refresh
    
    DoEvents
End Sub
Private Sub RenderVolScale()
    PicVolscale.Cls
    If FileLoaded <> "" Then
        Dim i As Long
        Dim Y As Single
        PicVolscale.ForeColor = vbBlack
        For i = 1 To AudioFile.FMTChunk.nChannels
            Y = ((UserControl.Height - PicSlidebar.Height) / (AudioFile.FMTChunk.nChannels + 1) * i) - _
                (((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))
        
            PicVolscale.Line (0, Y)-((PicVolscale.Width / 2), Y), vbBlack
            PicVolscale.CurrentX = (PicVolscale.Width - PicVolscale.TextWidth("00"))
            PicVolscale.CurrentY = Y - (PicVolscale.TextHeight("0") / 2)
            If (PicDisplay.Width < (UserControl.Width - PicVolscale.Width)) Then
                UserControl.ForeColor = GetColorShade(i, 0)
                UserControl.CurrentX = PicVolscale.CurrentX
                UserControl.CurrentY = PicVolscale.CurrentY
            End If
            PicVolscale.Print "0"
            
        Next
    End If
End Sub

Private Sub RenderFileInfo(ByVal clr As Long)
    If FileLoaded <> "" Then
        PicDisplay.CurrentX = -PicDisplay.Left: PicDisplay.CurrentY = -(((UserControl.Height - PicSlidebar.Height) / 2) - (PicDisplay.Height / 2))
        PicDisplay.ForeColor = clr
        Dim info As String
        info = CStr((TotalTime / 1000))
        If InStr(info, ".") > 0 Then info = Left(info, InStr(info, ".") - 1) & Mid(info, InStr(info, "."), 3)
        
        Select Case AudioFile.FMTChunk.nChannels
            Case 1
                info = "Standard PCM " & AudioFile.FMTChunk.nSamplesPerSec & " Hertz, " & _
                            AudioFile.FMTChunk.wBitsPerSample & "-bits, " & "Mono, " & LengthOfData(AudioFile) & " Bytes, " & info & " Seconds"
            Case 2
                info = "Standard PCM " & AudioFile.FMTChunk.nSamplesPerSec & " Hertz, " & _
                           AudioFile.FMTChunk.wBitsPerSample & "-bits, " & "Stereo, " & LengthOfData(AudioFile) & " Bytes, " & info & " Seconds"
            Case Is >= 5
                info = "Standard PCM " & AudioFile.FMTChunk.nSamplesPerSec & " Hertz, " & _
                            AudioFile.FMTChunk.wBitsPerSample & "-bits, " & AudioFile.FMTChunk.nChannels & " Channels, " & LengthOfData(AudioFile) & " Bytes, " & info & " Seconds"
        End Select
        PicDisplay.Print info
        UserControl.Cls
        UserControl.ForeColor = vbBlack
        UserControl.CurrentX = 0
        UserControl.CurrentY = 0
        UserControl.Print info
    End If
End Sub


Private Sub WaveInitial()

    Dim i As Single
    Dim d As Single

    RenderStatus "Loading wave file..."

    peak = 0: Dip = 0
    TotalTime = 0
    ChannelTap = 0
    Dim TotalBytes As Long
    TotalBytes = LengthOfData(AudioFile)
                
    If TotalBytes > 0 Then

        Dim elapse As Single
        elapse = Timer

        If AudioFile.FMTChunk.wBitsPerSample = 8 Then
            For d = LBound(AudioFile.WaveData16bit.ChannelData, 2) To UBound(AudioFile.WaveData16bit.ChannelData, 2)
    
                If AudioFile.WaveData8bit.ChannelData(1, d) < Dip Then Dip = CSng(AudioFile.WaveData8bit.ChannelData(1, d))
                If AudioFile.WaveData8bit.ChannelData(1, d) > peak Then peak = CSng(AudioFile.WaveData8bit.ChannelData(1, d))
    
    
                If Timer - elapse > 0.25 Then
                    elapse = Timer
                    RenderStatus "Loading wave file (" & CInt(((d / UBound(AudioFile.WaveData16bit.ChannelData, 2)) * 100)) & "%)..."
                End If
            Next
        Else
            For d = LBound(AudioFile.WaveData16bit.ChannelData, 2) To UBound(AudioFile.WaveData16bit.ChannelData, 2)
            
                If AudioFile.WaveData16bit.ChannelData(1, d) < Dip Then Dip = CSng(AudioFile.WaveData16bit.ChannelData(1, d))
                If AudioFile.WaveData16bit.ChannelData(1, d) > peak Then peak = CSng(AudioFile.WaveData16bit.ChannelData(1, d))
    
    
                If Timer - elapse > 0.25 Then
                    elapse = Timer
                    RenderStatus "Loading wave file (" & CInt(((d / UBound(AudioFile.WaveData16bit.ChannelData, 2)) * 100)) & "%)..."
                End If
            Next
        End If
        
        
        If ChannelTap < AudioFile.FMTChunk.nChannels Then ChannelTap = AudioFile.FMTChunk.nChannels
        
        TotalTime = WaveMilliseconds(AudioFile)
        
        MRatio = -(PicRender.Height / AudioFile.FMTChunk.nChannels)
        
        If (Abs(peak) + Abs(Dip)) <> 0 Then
            MRatio = MRatio / (Abs(peak) + Abs(Dip))
        Else
            MRatio = 1
        End If
        DBZoom = 0
             
        PicRender.Width = ((TotalTime / Screen.TwipsPerPixelY) * (AudioFile.FMTChunk.wBitsPerSample * AudioFile.FMTChunk.nSamplesPerSec))
        
       ' PicRender.Width = ((TotalTime * 1000) * (Screen.TwipsPerPixelX / 100))
       
        PicRender.Height = ((UserControl.Height - PicSlidebar.Height) / ChannelTap)
        PicDisplay.Left = 0
        PicDisplay.Top = 0
        PicDisplay.Width = (UserControl.ScaleWidth - PicVolscale.Width)
        PicDisplay.Height = (UserControl.Height - PicSlidebar.Height)
        
        ViewSlider.Width = PicSlidebar.Width
        
        If FileLoaded = "" Then
            HookObj Me
        Else
            FileLoaded = ""
        End If

        WaveDisplay ".\"

    End If
    
    
    PicDisplay.ForeColor = vbWhite
    PicDisplay.CurrentX = 0: PicDisplay.CurrentY = 0

    LeftMarker.X1 = 0
    LeftMarker.X2 = 0
    LeftMarker.Visible = True

    RightMarker.X1 = PicDisplay.Width
    RightMarker.X2 = PicDisplay.Width
    RightMarker.Visible = True


    LaststartMS = WaveMarkerStartMS
    LaststopMS = WaveMarkerStopMS
        
    PicDisplay.ForeColor = vbBlack
End Sub

Public Sub WaveDisplay(Optional ByVal fname As String = "")

    On Error GoTo catcherr

    If fname <> "" Or FileLoaded <> "" Then

        Dim CurrentTrack As Integer
        Dim i As Single

        Dim X As Single, Y As Single

        Dim y1 As Single

        Dim clrShade As Long
        Dim TotalBytes As Long
      '  TotalBytes = LengthOfData(AudioFile) 'ArrayBoundSize(AudioFile.WaveData16bit.ChannelData, 2) / 2  'LengthOfData(AudioFile)
    
        If AudioFile.FMTChunk.wBitsPerSample = 8 Then
            TotalBytes = (UBound(AudioFile.WaveData8bit.ChannelData, 2) * AudioFile.FMTChunk.nChannels) * (LenB(AudioFile.WaveData8bit.ChannelData(LBound(AudioFile.WaveData8bit.ChannelData, 1), LBound(AudioFile.WaveData8bit.ChannelData, 2))) * AudioFile.FMTChunk.nBlockAlign)
        ElseIf AudioFile.FMTChunk.wBitsPerSample = 16 Then
            TotalBytes = (UBound(AudioFile.WaveData16bit.ChannelData, 2) * AudioFile.FMTChunk.nChannels) * (LenB(AudioFile.WaveData16bit.ChannelData(LBound(AudioFile.WaveData16bit.ChannelData, 1), LBound(AudioFile.WaveData16bit.ChannelData, 2))) * AudioFile.FMTChunk.nBlockAlign)
        End If

        dBitsPerTwip = Screen.TwipsPerPixelX / AudioFile.FMTChunk.nBlockAlign
       ' dBitsPerTwip = AudioFile.Infos(1).wNumberOfChannels * AudioFile.Infos(1).wChannelBandwidth
       ' dBitsPerTwip = ((AudioFile.Datas(1).lBytes / AudioFile.Infos(1).wNumberOfChannels) / PicRender.Width)

        RenderStatus "Rendering wave file..."

        PicRender.Cls

        Dim ub As Long
        Dim lb As Long
        lb = LBound(AudioFile.WaveData16bit.ChannelData, 2)
        ub = UBound(AudioFile.WaveData16bit.ChannelData, 2)
        Dim elapse As Single
        elapse = Timer

        For i = lb To ub '- AudioFile.FMTChunk.nChannels Step AudioFile.FMTChunk.nChannels * Format(dBitsPerTwip, "##000")

            If Timer - elapse > 0.25 Then
                elapse = Timer
                RenderStatus "Rendering wave file (" & CInt((i / ub) * 100) & "%)..."
            End If

            X = X + ((PicRender.Width / (TotalBytes / AudioFile.FMTChunk.nBlockAlign)) * Format(dBitsPerTwip, "##000"))
      '      X = X + (PicRender.Width / (AudioFile.Datas(1).lBytes / AudioFile.Infos(1).wChannelBandwidth))

            For CurrentTrack = 1 To AudioFile.FMTChunk.nChannels
'
'                PicRender.ForeColor = GetColorShade(CurrentTrack, clrShade)
'                If CurrentTrack Mod AudioFile.Infos(1).wNumberOfChannels = 0 Then
'                    Select Case clrShade Mod (AudioFile.Infos(1).wChannelBandwidth + CurrentTrack)
'                        Case 0
'                            clrShade = 1
'                        Case 1
'                            clrShade = 2
'                        Case 2
'                            clrShade = 3
'                        Case 3
'                            clrShade = 4
'                        Case 4
'                            clrShade = 5
'                        Case 5
'                            clrShade = 0
'                    End Select
'                End If

                Y = (PicRender.Height / (AudioFile.FMTChunk.nChannels + 1) * CurrentTrack)

                If AudioFile.FMTChunk.wBitsPerSample = 8 Then
                    y1 = (Y + (AudioFile.WaveData8bit.ChannelData(CurrentTrack, i) * MRatio))
                ElseIf AudioFile.FMTChunk.wBitsPerSample = 16 Then
                    y1 = (Y + (AudioFile.WaveData16bit.ChannelData(CurrentTrack, i) * MRatio))
                End If
                
                PicRender.ForeColor = GetColorShade(CurrentTrack, (Abs(Y - y1) / Round((PicRender.Height / ((AudioFile.FMTChunk.nChannels + 1) * 2)) / 6)))
                
                If Not DotsOnly Then

                    PicRender.Line (X, Y)-(X, y1)

                    If i = 1 Then
                        PicRender.Line (0, Y)-(PicRender.Width, Y)
                        UserControl.Line (0, Y)-(PicRender.Width, Y), PicRender.ForeColor
                    End If
                Else
                    PicRender.Line (X, y1)-(X, y1), , BF

                    If i = 1 Then
                        PicRender.Line (0, Y)-(PicRender.Width, Y)
                        UserControl.Line (0, Y)-(PicRender.Width, Y), PicRender.ForeColor
                    End If

                End If

            Next CurrentTrack

        Next i

        LeftMarker.Visible = True
        RightMarker.Visible = True
        LeftMarker.X1 = 0
        LeftMarker.X2 = 0
        RightMarker.X1 = PicDisplay.Width
        RightMarker.X2 = PicDisplay.Width

        If FileLoaded = "" And fname <> "" Then
            FileLoaded = fname
        End If

        RenderVolScale

        PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height

        PicSlidebar.PaintPicture PicRender.Image, 0, 0, PicSlidebar.Width, PicSlidebar.Height, 0, 0, PicRender.Width, PicRender.Height

        ViewSlider.Width = (PicSlidebar.Width * ((UserControl.Width - PicVolscale.Width) / PicDisplay.Width))
        ViewSlider.Left = 0
        ViewSlider.Visible = True

        RenderTimeLine

        RenderFileInfo vbBlack



        
        ResetButton.Enabled = True
        PlayButton.Enabled = True
        StopButton.Enabled = True
        OpenButton.Enabled = False

    Else
        ResetButton.Enabled = False
        PlayButton.Enabled = False
        StopButton.Enabled = False
        OpenButton.Enabled = True
        LeftMarker.Visible = False
        RightMarker.Visible = False
        PlayCursor.Visible = False

        ViewSlider.Visible = False
        PicSlidebar.Cls
        PicVolscale.Cls
        PicDisplay.Cls
        UserControl.Cls


    End If

    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Private Sub OpenButton_Click()
    On Error Resume Next
    BrowseForWave.CancelError = True
    BrowseForWave.DefaultExt = "*.wav"
    BrowseForWave.Filter = "Wave Format|*.WAV"
    BrowseForWave.FilterIndex = 0
    BrowseForWave.FileName = FileLoaded
    BrowseForWave.ShowOpen
    
    If Err.Number = cdlCancel Then
        Err.Clear
    Else
        On Error GoTo 0
        WaveFromFile BrowseForWave.FileName
    End If
    PicDisplay.SetFocus
End Sub

Private Sub PlayButton_Click()
    PicDisplay.SetFocus
    WavePlay

End Sub

Private Sub StopButton_Click()
    PicDisplay.SetFocus
    WaveStop
End Sub

Private Sub ResetButton_Click()
    PicDisplay.SetFocus
    WaveClose
End Sub

Private Sub PicSlidebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PicSlidebar.Tag = X
    Else
        PicSlidebar.Tag = 0
    End If
End Sub

Private Sub PicSlidebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If PicSlidebar.Tag <> 0 And ViewSlider.Width < PicSlidebar.Width Then
            PicSlidebar.Tag = Abs(PicSlidebar.Tag)
            MoveViewPort X

            PicSlidebar.Tag = -X

        End If
    Else
        PicSlidebar.Tag = 0
    End If
End Sub

Private Sub PicSlidebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X = PicSlidebar.Tag And ViewSlider.Width < PicSlidebar.Width Then
        PicSlidebar.Tag = (ViewSlider.Width / 2)
        MoveViewPort (X - ViewSlider.Left)
    Else
        PicSlidebar.Tag = 0
    End If
End Sub

Private Sub PicVolscale_Click()
    PicDisplay.SetFocus
    DotsOnly = Not DotsOnly
    WaveDisplay FileLoaded
End Sub

Private Sub PicDisplay_Click()
    RaiseEvent Click
End Sub

Private Sub PicDisplay_DblClick()
    RightMarker.X1 = PicDisplay.Width - Screen.TwipsPerPixelX
    RightMarker.X2 = PicDisplay.Width - Screen.TwipsPerPixelX
    LeftMarker.X1 = Screen.TwipsPerPixelX
    LeftMarker.X2 = Screen.TwipsPerPixelX
End Sub

Private Sub PicDisplay_KeyPress(KeyAscii As Integer)

    On Error GoTo catcherr
    
'    If KeyAscii >= Asc("1") And KeyAscii <= Asc("9") Then
'
'        FilterScripts.Reset
'        FilterScripts.AddCode GetSetting(App.Title, "Filters", "VBScript")
'
'        'WaveFilter AudioFile, WaveMarkerStartMS, WaveMarkerStopMS - WaveMarkerStartMS, FilterScripts, "func" & Chr(KeyAscii) & "(%)"
'
'        'WaveDisplay
'    End If

    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Private Sub PicDisplay_LostFocus()
    LastX = 0
    LastY = 0
End Sub

Private Sub PicDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static lX As Single
  
    If ((Button = 1) And (PicDisplay.MousePointer = 0)) Then

        If (lX <> Screen.TwipsPerPixelX) Then
            RightMarker.Tag = LeftMarker.X1
        End If
        LeftMarker.Tag = X

        If LeftMarker.X1 = RightMarker.X1 Then
            RightMarker.Tag = PicDisplay.Width - Screen.TwipsPerPixelX
        ElseIf (RightMarker.X2 = PicDisplay.Width - Screen.TwipsPerPixelX) Then
            LeftMarker.Tag = Screen.TwipsPerPixelX
        End If
        lX = X
        
    ElseIf (Button = 2) Then

    ElseIf (Button = 0) Then
        lX = Screen.TwipsPerPixelX
    End If

End Sub

Private Sub PicDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (X <> LastX And Y <> LastY And Button = 1) And (LastBtn = 1 And LastX <> 0 And LastY <> 0) Then

        PicDisplay.OLEDrag

    Else
        If Button = 0 Then
            If PicDisplay.MousePointer <> 0 Then
                PicDisplay.MousePointer = 0
            End If
        End If
    End If
    
    LastX = X
    LastY = Y
    LastBtn = Button
    PicDisplay.SetFocus
End Sub

Private Sub PicDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (PicDisplay.MousePointer <> 0) Then
    
        If PicDisplay.MousePointer = 99 Then
            PicDisplay.Drag DragConstants.vbEndDrag
        ElseIf PicDisplay.MousePointer = 12 Then
            PicDisplay.Drag DragConstants.vbCancel
        End If
        LeftMarker.Tag = ""
        RightMarker.Tag = ""
    Else

        If IsNumeric(LeftMarker.Tag) Then
            LeftMarker.X1 = LeftMarker.Tag
            LeftMarker.X2 = LeftMarker.Tag
            LeftMarker.Tag = ""
        End If
        If IsNumeric(RightMarker.Tag) Then
            RightMarker.X1 = RightMarker.Tag
            RightMarker.X2 = RightMarker.Tag
            RightMarker.Tag = ""
        End If
        
    End If

End Sub

Private Sub PicDisplay_OLECompleteDrag(Effect As Long)

    Screen.MousePointer = 0

End Sub

Private Sub PicDisplay_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo catcherr
    
    If data.GetFormat(MyFormat) Then
        If Screen.MousePointer = 99 Then
            
'            Dim temp() As Byte
'            temp = data.GetData(MyFormat)
'
'            If Shift Then
'
'
'            Else
'                Dim af As AudioFile2
'                af = WaveBytesAsAudio(temp)
'                If LengthOfData(af) > 0 Then
'                    Dim SourceDurationTimeMS As Single
'                    SourceDurationTimeMS = Round(((LengthOfData(af) / af.FMTChunk.nAvgBytesPerSec) * 1000))
'                    If SourceDurationTimeMS - (WaveMilliseconds(af) - WaveMarkerStartMS) < 0 Then
'                        Err.Raise 8, "WaveData", "Invalid or exceeding duration time."
'                    Else
'                        AudioFile = WaveBytesAsAudio(WaveCombine(WaveData, WaveMarkerStartMS, temp, SourceDurationTimeMS))
'                        WaveInitial
'
'                    End If
'                End If
'            End If
'            Erase temp
            Screen.MousePointer = 0
        End If
    End If
    
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
End Sub

Private Sub DragDetermine()

    Dim pt As POINTAPI
    GetCursorPos pt
    If CStr(WindowFromPoint(pt.X, pt.Y)) = CStr(PicDisplay.hwnd) Then
         Screen.MousePointer = 12
    Else
         Set Screen.MouseIcon = DragIcon.Picture
         Screen.MousePointer = 99
    End If
End Sub
Private Sub PicDisplay_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

    Effect = vbDropEffectMove Or vbDropEffectCopy
    DefaultCursors = False

    DragDetermine
End Sub

Private Sub PicDisplay_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    
    AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
    
    If FileLoaded <> "" Then

        data.Clear
        data.SetData WaveData, MyFormat

        PicDisplay.MousePointer = 12

    End If

End Sub

Private Sub Blinkers_Timer()

    LeftMarker.BorderColor = IIf((LeftMarker.BorderColor = &H808080), &H404040, &H808080)
    RightMarker.BorderColor = IIf((RightMarker.BorderColor = &H808080), &H404040, &H808080)

    PlayCursor.BorderColor = IIf((RightMarker.BorderColor = &H0&), &HE0E0E0, &H0&)
   
    If LeftMarker.Visible <> (FileLoaded <> "") Then LeftMarker.Visible = (FileLoaded <> "")
    If RightMarker.Visible <> (FileLoaded <> "") Then RightMarker.Visible = (FileLoaded <> "")

End Sub


Private Sub UserControl_Click()
    PicDisplay.SetFocus
End Sub

Private Sub UserControl_Hide()
    WaveStop
End Sub

Public Sub WaveClose()

    On Error GoTo catcherr
    
    WaveStop

    If FileLoaded <> "" Then
        FileLoaded = ""
        HookObj Me
    End If

    WaveResetRecord AudioFile
    
    
    WaveDisplay
    Exit Sub
catcherr:
    MsgBox Err.Description, vbCritical, "An error occured"
    Err.Clear
    
End Sub

Private Sub UserControl_Initialize()
   
    SilentBytes = WaveSilentAudio(0)
    MyFormat = RegisterClipboardFormat("audio/wav")
    
    PicSlidebar.Tag = 0
    WaveDisplay
    UserControl_Resize

End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicDisplay.SetFocus
End Sub

Private Sub UserControl_Resize()
    On Error GoTo exitresize
    If UserControl.Width < ((17 * Screen.TwipsPerPixelX) * 2) Then
        UserControl.Width = ((17 * Screen.TwipsPerPixelX) * 2)
    End If
    If UserControl.Height < ((17 * Screen.TwipsPerPixelY) * 2) Then
        UserControl.Height = ((17 * Screen.TwipsPerPixelY) * 2)
    End If
    
    PicDisplay.Left = 0
    PicDisplay.Top = 0
    PicDisplay.Height = (UserControl.ScaleHeight - PicSlidebar.Height)
    PicDisplay.Width = (UserControl.ScaleWidth - PicVolscale.Width)

    PicSlidebar.Height = (17 * Screen.TwipsPerPixelY)
    PicVolscale.Width = (17 * Screen.TwipsPerPixelX)
    PicSlidebar.Left = 0
    PicSlidebar.Top = (UserControl.Height - PicSlidebar.Height)
    PicSlidebar.Width = (UserControl.Width - (OpenButton.Width + PlayButton.Width + StopButton.Width + ResetButton.Width))
    PicVolscale.Left = (UserControl.Width - PicVolscale.Width)
    PicVolscale.Top = 0
    PicVolscale.Height = (UserControl.Height - OpenButton.Height)
    ViewSlider.Top = 0
    ViewSlider.Height = PicSlidebar.Height
    
    ResetButton.Top = (UserControl.Height - ResetButton.Height)
    ResetButton.Left = (UserControl.Width - ResetButton.Width)


    StopButton.Top = (UserControl.Height - PicSlidebar.Height)
    StopButton.Left = (ResetButton.Left - StopButton.Width)
    
    PlayButton.Top = (UserControl.Height - PicSlidebar.Height)
    PlayButton.Left = (StopButton.Left - PlayButton.Width)
    
    OpenButton.Top = (UserControl.Height - PicSlidebar.Height)
    OpenButton.Left = (PlayButton.Left - OpenButton.Width)

    LeftMarker.y1 = 0
    LeftMarker.Y2 = UserControl.Height
    RightMarker.y1 = 0
    RightMarker.Y2 = UserControl.Height
    PlayCursor.y1 = 0
    PlayCursor.Y2 = UserControl.Height


    If FileLoaded <> "" Then

        PicDisplay.PaintPicture PicRender.Image, 0, 0, PicDisplay.Width, PicDisplay.Height, 0, 0, PicRender.Width, PicRender.Height

        PicSlidebar.PaintPicture PicRender.Image, 0, 0, PicSlidebar.Width, PicSlidebar.Height, 0, 0, PicRender.Width, PicRender.Height

        ViewSlider.Width = (PicSlidebar.Width * ((UserControl.Width - PicVolscale.Width) / PicDisplay.Width))
        
        If PicDisplay.Width < (UserControl.Width - PicVolscale.Width) Then
            PicDisplay.Left = 0
            ViewSlider.Left = 0
        Else
            If PicDisplay.Left < 0 Then
                ViewSlider.Left = (PicSlidebar.Width * (-PicDisplay.Left / PicDisplay.Width))
            End If
            If (PicDisplay.Width + PicDisplay.Left) < (UserControl.Width - PicVolscale.Width) _
                And ((PicDisplay.Left < 0) And (PicDisplay.Width > (UserControl.Width - PicVolscale.Width))) Then
                PicDisplay.Left = -(PicDisplay.Width - (UserControl.Width - PicVolscale.Width))
                ViewSlider.Left = PicSlidebar.Width - ViewSlider.Width
            End If
        End If

        If TotalTime > 0 Then
            LeftMarker.X1 = (PicDisplay.Width * (LaststartMS / TotalTime))
            LeftMarker.X2 = LeftMarker.X1
            RightMarker.X1 = (PicDisplay.Width * (LaststopMS / TotalTime))
            RightMarker.X2 = RightMarker.X1
        End If
        RenderVolScale

        RenderTimeLine
        RenderFileInfo vbBlack
        
        LaststartMS = WaveMarkerStartMS
        LaststopMS = WaveMarkerStopMS
        
    End If
    Exit Sub
exitresize:
    Debug.Print "Usercontrol_Resize(Error)"
    Err.Clear
    
End Sub

Private Sub UserControl_Terminate()
    WaveClose
End Sub

Private Function GetColorShade(ByVal Track As Long, ByVal Shade As Long) As Long
    Select Case Track
        Case 1
            Select Case Shade
            'blue
                Case 0
                    GetColorShade = &HFFC0C0
                Case 1
                    GetColorShade = &HFF8080
                Case 2
                    GetColorShade = &HFF0000
                Case 3
                    GetColorShade = &HC00000
                Case 4
                    GetColorShade = &H800000
                Case 5
                    GetColorShade = &H400000
            End Select
        Case 2
            Select Case Shade
            'red
                Case 0
                    GetColorShade = &HC0C0FF
                Case 1
                    GetColorShade = &H8080FF
                Case 2
                    GetColorShade = &HFF&
                Case 3
                    GetColorShade = &HC0&
                Case 4
                    GetColorShade = &H80&
                Case 5
                    GetColorShade = &H40&
            End Select
        Case 3
            Select Case Shade
            'green
                Case 0
                    GetColorShade = &HC0FFC0
                Case 1
                    GetColorShade = &H80FF80
                Case 2
                    GetColorShade = &HFF00&
                Case 3
                    GetColorShade = &HC000&
                Case 4
                    GetColorShade = &H8000&
                Case 5
                    GetColorShade = &H4000&
            End Select
        Case 4
            Select Case Shade
            'yellow
                Case 0
                    GetColorShade = &HC0FFFF
                Case 1
                    GetColorShade = &H80FFFF
                Case 2
                    GetColorShade = &HFFFF&
                Case 3
                    GetColorShade = &HC0C0&
                Case 4
                    GetColorShade = &H8080&
                Case 5
                    GetColorShade = &H4040&
            End Select
        Case 5
            Select Case Shade
            'teal
                Case 0
                    GetColorShade = &HFFFFC0
                Case 1
                    GetColorShade = &HFFFF80
                Case 2
                    GetColorShade = &HFFFF00
                Case 3
                    GetColorShade = &HC0C000
                Case 4
                    GetColorShade = &H808000
                Case 5
                    GetColorShade = &H404000
            End Select
        Case 6
            Select Case Shade
            'purple
                Case 0
                    GetColorShade = &HFFC0FF
                Case 1
                    GetColorShade = &HFF80FF
                Case 2
                    GetColorShade = &HFF00FF
                Case 3
                    GetColorShade = &HC000C0
                Case 4
                    GetColorShade = &H800080
                Case 5
                    GetColorShade = &H400040
            End Select
        Case 7
            Select Case Shade
            'black
                Case 0
                    GetColorShade = &HFFFFFF
                Case 1
                    GetColorShade = &HE0E0E0
                Case 2
                    GetColorShade = &HC0C0C0
                Case 3
                    GetColorShade = &H808080
                Case 4
                    GetColorShade = &H404040
                Case 5
                    GetColorShade = &H0&
            End Select
        Case 8
            Select Case Shade
            'orange
                Case 0
                    GetColorShade = &HC0E0FF
                Case 1
                    GetColorShade = &H80C0FF
                Case 2
                    GetColorShade = &H80FF&
                Case 3
                    GetColorShade = &H40C0&
                Case 4
                    GetColorShade = &H4080&
                Case 5
                    GetColorShade = &H404080
            End Select
    End Select

End Function
