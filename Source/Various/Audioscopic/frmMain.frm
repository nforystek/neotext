VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "Wave View"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2010
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin Audioscopic.WaveView WaveView1 
      Height          =   1335
      Left            =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2355
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CodeEdit1_Change()
'    SaveSetting App.Title, "Filters", "VBScript", CodeEdit1.Text
End Sub

Private Sub form_Activate()
    mdiMain.ResetSelection
End Sub

Private Sub Form_Load()
    


'    CodeEdit1.Text = GetSetting(App.Title, "Filters", "VBScript", "'VBScript filters apply with" & vbCrLf & _
'                                    "'keys 1 - 9, default is mute" & vbCrLf & vbCrLf & _
'                                    "Function Func1(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func2(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func3(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func4(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func5(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func6(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func7(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func8(x)" & vbCrLf & "End Function" & vbCrLf & vbCrLf & _
'                                    "Function Func9(x)" & vbCrLf & "End Function" & vbCrLf)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    WaveView1.Left = 3 * Screen.TwipsPerPixelX
    WaveView1.Top = 3 * Screen.TwipsPerPixelY
    
    
    WaveView1.Height = Me.ScaleHeight - (6 * Screen.TwipsPerPixelY)
    WaveView1.Width = Me.ScaleWidth - (6 * Screen.TwipsPerPixelX)

'    If CodeEdit1.Visible Then
'
'    WaveView1.Width = ((Me.ScaleWidth / 3) * 2)
'    WaveView1.Height = (Me.ScaleHeight / 2) - (2 * Screen.TwipsPerPixelY)
'
'
'    WaveView2.Width = ((Me.ScaleWidth / 3) * 2)
'    WaveView2.Height = WaveView1.Height
'
'
'    CodeEdit1.Left = ((Me.ScaleWidth / 3) * 2)
'    CodeEdit1.Width = (Me.ScaleWidth / 3)
'    CodeEdit1.Top = 0
'    CodeEdit1.Height = Me.ScaleHeight
'    Else
'
'    WaveView1.Width = Me.ScaleWidth
'    WaveView1.Height = (Me.ScaleHeight / 2) - (2 * Screen.TwipsPerPixelY)
'
'
'    WaveView2.Width = Me.ScaleWidth
'    WaveView2.Height = WaveView1.Height
'
'    End If
    

    If Err Then Err.Clear
End Sub

