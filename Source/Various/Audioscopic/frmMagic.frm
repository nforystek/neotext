VERSION 5.00
Begin VB.Form frmMagic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make magic..."
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   5490
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1095
      TabIndex        =   3
      Text            =   "4"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3975
      TabIndex        =   1
      Top             =   210
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2535
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -135
      Top             =   -15
   End
   Begin VB.Label Label2 
      Caption         =   "(feet)"
      Height          =   255
      Left            =   1845
      TabIndex        =   4
      Top             =   270
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Distance:"
      Height          =   240
      Left            =   255
      TabIndex        =   2
      Top             =   270
      Width           =   735
   End
End
Attribute VB_Name = "frmMagic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private parentLeft As Single
Private parentTop As Single

Private myLeft As Single
Private myTop As Single


Private Sub Command1_Click()
    If IsNumeric(Text1.Text) Then
        If Val(Text1.Text) > 0 Then
            Dim frm As Form
            Set frm = mdiMain.ActiveForm
            Command1.Enabled = False
            Text1.Enabled = False


'            frm.WaveView1.WaveSaveAs frm.WaveView1.FileName
'            Exit Sub
'################################################################################################################################################
'################################################################################################################################################
'################################################################################################################################################
           
            
            
'################################################################################################################################################
'################################################################################################################################################
'################################################################################################################################################
            


'            Dim nfrm As New frmMain
'
'            AddSilenceToWav frm.WaveView1.FileName, App.Path & "\withsilence.wav", SoundTravelTimeMilliseconds(CInt(Text1.Text)), False
'            ChangeWaveVolume frm.WaveView1.FileName, App.Path & "\downvol.wav", SoundLevelAtDistance(CInt(Text1.Text))
'            MixWavFiles App.Path & "\downvol.wav", App.Path & "\withsilence.wav", App.Path & "\combined.wav"
'            BlendWavFiles frm.WaveView1.FileName, App.Path & "\combined.wav", App.Path & "\midwave.wav", SoundTravelTimeMilliseconds(CInt(Text1.Text))
'            SubtractWavFile frm.WaveView1.FileName, App.Path & "\midwave.wav", App.Path & "\halfwave.wav"
'            MixWavFiles App.Path & "\combined.wav", App.Path & "\halfwave.wav", App.Path & "\final.wav"
'            nfrm.WaveView1.WaveFromFile App.Path & "\final.wav"
'
'
'
''################################################################################################################################################
''################################################################################################################################################
''################################################################################################################################################

            Dim nfrm As New frmMain

            ChangeWaveVolume frm.WaveView1.FileName, App.Path & "\downvol.wav", SoundLevelAtDistance(CInt(Text1.Text))
            AddSilenceToWav frm.WaveView1.FileName, App.Path & "\withsilence.wav", SoundTravelTimeMilliseconds(CInt(Text1.Text)), False
            MixWavFiles App.Path & "\downvol.wav", App.Path & "\withsilence.wav", App.Path & "\combined.wav"
            BlendWavFiles frm.WaveView1.FileName, App.Path & "\combined.wav", App.Path & "\final.wav", SoundTravelTimeMilliseconds(CInt(Text1.Text))

            nfrm.WaveView1.WaveFromFile App.Path & "\final.wav"


            
'################################################################################################################################################
'################################################################################################################################################
'################################################################################################################################################

            Me.Hide
   
               
            nfrm.Show
            Set nfrm = Nothing
            
            
            
            
            
            Set frm = Nothing
            Unload Me
        Else
            MsgBox "Value must be above zero."
        End If
    Else
        MsgBox "Value must be numerical."
    End If
End Sub

Private Sub Command2_Click()
    If Command1.Enabled Then
        Unload Me
    Else
        Me.Hide
    End If

End Sub

Private Sub Form_Load()
    Me.Left = (mdiMain.Width / 2) - (Me.Width / 2) + mdiMain.Left
    Me.Top = (mdiMain.Height / 2) - (Me.Height / 2) + mdiMain.Top
    
    myLeft = Me.Left
    myTop = Me.Top
    
    
    parentLeft = mdiMain.Left
    parentTop = mdiMain.Top
    
    Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
    If myLeft <> Me.Left Then
        Dim xdiff As Single
        xdiff = myLeft - Me.Left
        mdiMain.Left = parentLeft - xdiff
        parentLeft = mdiMain.Left
        myLeft = Me.Left
    End If
    If myTop <> Me.Top Then
        Dim ydiff As Single
        ydiff = myTop - Me.Top
        mdiMain.Top = parentTop - ydiff
        parentTop = mdiMain.Top
        myTop = Me.Top
    End If

End Sub
