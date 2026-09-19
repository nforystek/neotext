VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "mdiMain"
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12510
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog BrowseForWave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuMagic 
         Caption         =   "&Magic..."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuDash231 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private initDir As String
Private iniFile As String

Private Sub MDIForm_Activate()
    ResetSelection
End Sub
Public Sub ResetSelection()
    If Not Me.ActiveForm Is Nothing Then
        Dim selHwnd As Long
        selHwnd = Me.ActiveForm.hwnd
        Dim frm As Form
        For Each frm In Forms
            If TypeName(frm) = "frmMain" Then
                If frm.hwnd = selHwnd Then
                    frm.BackColor = &H8000000D
                Else
                    frm.BackColor = &H8000000F
                End If
            End If
        Next
    End If
End Sub
Private Sub MDIForm_Load()
    initDir = App.Path
    
   ' ChangeWaveVolume App.Path & "\g note 0 full.wav", App.Path & "\downvol.wav", -3
    'AddSilenceToWav App.Path & "\g note 0 full.wav", App.Path & "\silence.wav", 3000, False
   
   ' End
    
    '    Dim s As New NTNodes10.Stream
'
'    s.Concat StrConv("riff a s data  ftm ", vbFromUnicode)
'    Debug.Print StrConv(s.Partial(0, 1), vbUnicode)
'
'
'
'    Debug.Print FindTag("data", s)
    
    
'    Dim inBytes() As Byte
'    inBytes = ReadFileBytes(App.Path & "\g note 0 full.wav")
'    Dim R As New Riff
'    R.ToBytes = inBytes
'
'
'
'    Set R = Nothing
'
'
'
'
'       End
       
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmMain" Then
            If Not frm Is Nothing Then
                Unload frm
            End If
        End If
    Next
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Click()
    
    If Not (ActiveForm Is Nothing) Then
        mnuClose.Enabled = True
        mnuMagic.Enabled = (Me.ActiveForm.WaveView1.FileName <> "")
        mnuSaveAs.Enabled = (Me.ActiveForm.WaveView1.FileName = "")
        
    Else
        mnuClose.Enabled = False
        mnuMagic.Enabled = False
        mnuSaveAs.Enabled = False
    End If
End Sub

Private Sub mnuMagic_Click()
    frmMagic.Show 1
    
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    BrowseForWave.CancelError = True
    BrowseForWave.DefaultExt = "*.wav"
    BrowseForWave.Filter = "Wave Format|*.WAV"
    BrowseForWave.FilterIndex = 0
    BrowseForWave.FileName = iniFile
    BrowseForWave.initDir = initDir
    
    BrowseForWave.ShowOpen
    
    If Err.Number = cdlCancel Then
        Err.Clear
    Else
        On Error GoTo 0
        iniFile = BrowseForWave.FileName
        initDir = BrowseForWave.FileName
        Dim nwfrm As New frmMain
        nwfrm.WaveView1.WaveFromFile BrowseForWave.FileName
        nwfrm.Show
        nwfrm.WaveView1.SetFocus
        
        Set nwfrm = Nothing
        
    End If
    
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    BrowseForWave.CancelError = True
    BrowseForWave.DefaultExt = "*.wav"
    BrowseForWave.Filter = "Wave Format|*.WAV"
    BrowseForWave.FilterIndex = Me.ActiveForm.WaveView1.FileName
    BrowseForWave.initDir = Me.ActiveForm.WaveView1.FileName
    
    BrowseForWave.ShowSave
    
    If Err.Number = cdlCancel Then
        Err.Clear
    Else
        On Error GoTo 0
        iniFile = BrowseForWave.FileName
        initDir = BrowseForWave.FileName
        Me.ActiveForm.WaveView1.WaveSaveAs BrowseForWave.FileName
        
    End If
End Sub
