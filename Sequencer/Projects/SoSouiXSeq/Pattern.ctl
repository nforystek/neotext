VERSION 5.00
Begin VB.UserControl Pattern 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   ScaleHeight     =   675
   ScaleWidth      =   12315
   Begin SoSouiXSeq.Mixer pMixer 
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   979
   End
   Begin VB.Menu mnuPattern 
      Caption         =   "Pattern"
      Visible         =   0   'False
      Begin VB.Menu mnuAddWave 
         Caption         =   "Add Wave Mixer"
      End
      Begin VB.Menu mnuAddMidi 
         Caption         =   "Add MIDI Mixer"
      End
   End
End
Attribute VB_Name = "Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private nextIndex As Long
Public Event Resize()

Public Function AddMixer(ByVal MixerType As String) As Integer
    Load pMixer(nextIndex)
    pMixer(nextIndex).Visible = True
    pMixer(nextIndex).Number.Caption = (pMixer.Count - 1)
    pMixer(nextIndex).Pattern = frmMain.SelPattern
    pMixer(nextIndex).MixerType = MixerType
    nextIndex = nextIndex + 1

    UserControl_Resize

End Function

Public Function RemoveMixer(ByVal Index As Integer)
    If Index > 0 Then
        Dim X
        Dim indexCnt As Integer
        Dim found As Integer
        indexCnt = 0
        found = 0
        For Each X In pMixer
            If X.Index <> 0 Then
                If found = 0 Then
                    indexCnt = indexCnt + 1
                    If indexCnt = Index Then
                        found = X.Index
                    End If
                Else
                    X.Number.Caption = indexCnt
                    indexCnt = indexCnt + 1
                End If
            End If
        Next
        
        If found > 0 Then
            Unload pMixer(found)
            RemoveMixer = Index
        Else
            RemoveMixer = 0
        End If
'        indexCnt = 0
'        For Each X In pMixer
'            X.Number = indexCnt
'            indexCnt = indexCnt + 1
'        Next
        UserControl_Resize
    End If
End Function

Public Property Get Mixers(Optional ByVal Index As Integer = 0)
    If Index = 0 Then
        Set Mixers = pMixer
    Else
        Dim X
        Dim indexCnt As Integer
        Dim found As Integer
        indexCnt = 0
        found = 0
        For Each X In pMixer
            If X.Index <> 0 Then
                If found = 0 Then
                    indexCnt = indexCnt + 1
                    If indexCnt = Index Then
                        found = X.Index
                    End If
                Else
                    X.Number.Caption = indexCnt
                    indexCnt = indexCnt + 1
                End If
            End If
        Next
        
        If found > 0 Then
            Set Mixers = pMixer(found)
        End If
    End If
End Property

Private Sub mnuAddMidi_Click()
    AddMixer "M"
End Sub

Private Sub mnuAddWave_Click()
    AddMixer "W"
End Sub


Private Sub UserControl_Initialize()
    nextIndex = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        UserControl.PopupMenu mnuPattern
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 12315
    Dim X
    Dim top As Long
    top = 60
    For Each X In pMixer
        If X.Index > 0 Then
            X.top = top
            top = top + pMixer(0).Height + 60
        End If
    Next
    UserControl.Height = top + pMixer(0).Height + 60
    RaiseEvent Resize
End Sub

