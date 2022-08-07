VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4515
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   StartUpPosition =   2  'CenterScreen
   Tag             =   "about"
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   20000
      Left            =   3660
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1140
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   -15
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   8  'Xor Pen
      Height          =   5310
      Left            =   90
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   4320
      Begin VB.Image Image3 
         Height          =   960
         Left            =   1695
         Picture         =   "frmAbout.frx":08CA
         Top             =   30
         Width           =   960
      End
      Begin VB.Image Image2 
         Height          =   1065
         Left            =   600
         Picture         =   "frmAbout.frx":1D0C
         Top             =   1845
         Width           =   3090
      End
      Begin VB.Label prgRegistered 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   3030
         Width           =   4245
      End
      Begin VB.Label prgVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0b1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   3390
         Width           =   4125
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   352
      Index           =   2
      Left            =   1680
      TabIndex        =   0
      Top             =   5430
      Width           =   1072
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TOP DOWN
Option Explicit

Option Compare Binary

Private Type Particle
    X As Single
    Y As Single
    Xv As Single
    Yv As Single
    Life As Integer
    Dead As Boolean
    Color As Long
End Type

Private Type FireWork
    X As Single
    Y As Single
    Height As Integer
    Color As Long
    Exploded As Boolean
    P() As Particle
End Type

Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim BF As BLENDFUNCTION
Dim lBF As Long

Dim FW() As FireWork
Dim FWCount As Integer
Dim RocketSpeed As Integer

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    
        Case 2
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    prgVersion.Caption = "Version: " & AppVersion
    Me.Caption = "About " & AppName & " " & AppVersion
  '  Label1(1).Caption = NeoTextWebSite
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer3.enabled = True
    
    Timer1.enabled = False
    Timer2.enabled = False
    
    Picture2.Visible = True
    Command1(2).Visible = True
    
    frmAbout.Cls
    
End Sub

Private Sub Label1_Click(Index As Integer)
    RunFile AppPath & "Neotext.org.url"
 
End Sub


Private Sub StartFireWork()
    Dim i As Long
    
    For i = 0 To FWCount
        If FW(i).Y = -1 Then
            GoTo MAKEFIREWORK
        End If
    Next i
    
    FWCount = FWCount + 1
    
    ReDim Preserve FW(FWCount)
    i = FWCount
    
MAKEFIREWORK:
    
    With FW(i)
        .X = Int(Rnd * Me.ScaleWidth)
        .Y = Me.ScaleHeight
        .Height = Rnd * Me.ScaleHeight ' - ((Me.ScaleHeight / 2) + Int((Me.ScaleHeight / 2) * Rnd))
        .Color = Int(Rnd * vbWhite)
        .Exploded = False
        ReDim .P(10)
    End With
End Sub

Private Sub DrawFireWork(tFW As FireWork)
    Dim DeadCount As Integer
    Dim RndSpeed As Single
    Dim RndDeg As Single
    Dim i As Long
    
    With tFW
        If .Exploded Then
            For i = 0 To UBound(.P)
                If .P(i).Life > 0 Then
                    .P(i).Life = .P(i).Life - 1
                    .P(i).X = .P(i).X + .P(i).Xv
                    .P(i).Y = .P(i).Y + .P(i).Yv
                    .P(i).Xv = .P(i).Xv / 1.05
                    .P(i).Yv = .P(i).Yv / 1.05 + 0.05
                    PSet (.P(i).X, .P(i).Y), .P(i).Color
                ElseIf .P(i).Life > -40 Then
                    .P(i).Life = .P(i).Life - 1
                    .P(i).X = .P(i).X + .P(i).Xv + (0.5 - Rnd)
                    .P(i).Y = .P(i).Y + .P(i).Yv + 0.1
                    .P(i).Xv = .P(i).Xv / 1.05
                    .P(i).Yv = .P(i).Yv
                    SetPixelV Me.hDC, .P(i).X, .P(i).Y, .P(i).Color
                Else
                    .P(i).Dead = True
                    DeadCount = DeadCount + 1
                End If
            Next i
            
            If DeadCount >= UBound(.P) Then
                .Y = -1
            End If
        Else
            .Y = .Y - RocketSpeed
            If .Y < .Height Then
                Dim ExplosionShape As Integer
                
                ExplosionShape = Int(Rnd * 6)
                
                Select Case ExplosionShape
                    Case 0 'Regular
                        ReDim .P(Int(Rnd * 100) + 100)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = Int(Rnd * 20) + 20
                            
                            RndSpeed = (Rnd * 5)
                            RndDeg = (Rnd * 360) / 57.3
                            
                            .P(i).Xv = RndSpeed * Cos(RndDeg)
                            .P(i).Yv = RndSpeed * Sin(RndDeg)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 1 'Smilely
                        ReDim .P(35)
                        ReDim .P(50)
                        ReDim .P(52)
                        
                        For i = 0 To 35
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = 3 * Cos(((360 / 35) * (i + 1)) / 57.3)
                            .P(i).Yv = 3 * Sin(((360 / 35) * (i + 1)) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        For i = 36 To 50
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = 2 * Cos(((360 / 35) * i + 15) / 57.3)
                            .P(i).Yv = 2 * Sin(((360 / 35) * i + 15) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        With .P(51)
                            .X = tFW.X
                            .Y = tFW.Y
                            .Life = 50
                            .Xv = 2 * Cos(-55 / 57.3)
                            .Yv = 2 * Sin(-55 / 57.3)
                            .Color = tFW.Color
                        End With
                        
                        With .P(52)
                            .X = tFW.X
                            .Y = tFW.Y
                            .Life = 50
                            .Xv = 2 * Cos(-125 / 57.3)
                            .Yv = 2 * Sin(-125 / 57.3)
                            .Color = tFW.Color
                        End With
                        
                        .Exploded = True
                    Case 2 'Star
                        ReDim .P(50)
                        
                        RndDeg = Int(360 * Rnd)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = (i * 0.1) * Cos(((360 / 5) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Yv = (i * 0.1) * Sin(((360 / 5) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 3 'Spiral
                        ReDim .P(50)
                        
                        RndDeg = (360 * Rnd)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = (i * 0.1) * Cos(((360 / 25) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Yv = (i * 0.1) * Sin(((360 / 25) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 4 'Regular Random
                        
                        
                        ReDim .P(Int(Rnd * 100) + 100)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = Int(Rnd * 20) + 20
                            
                            RndSpeed = (Rnd * 5)
                            RndDeg = (Rnd * 360) / 57.3
                            
                            .P(i).Xv = RndSpeed * Cos(RndDeg)
                            .P(i).Yv = RndSpeed * Sin(RndDeg)
                            .P(i).Color = Int(Rnd * vbWhite)
                        Next i
                        
                        .Exploded = True
                End Select
            Else
                SetPixelV Me.hDC, .X, .Y, vbWhite
            End If
        End If
    End With
End Sub


Private Sub prgVersion_Click()

End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    For i = 0 To FWCount
        DrawFireWork FW(i)
    Next i
    
    RtlMoveMemory lBF, BF, 4
    AlphaBlend Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
    Me.Refresh
End Sub

Private Sub Timer2_Timer()
    StartFireWork
End Sub

Private Sub Timer3_Timer()
    Timer3.enabled = False
    
    Timer1.enabled = True
    Timer2.enabled = True

    Picture2.Visible = False
    Command1(2).Visible = False

    Randomize

    RocketSpeed = 4
    Timer2.Interval = 521

    FWCount = -1
    
    Picture1.Width = Me.ScaleWidth
    Picture1.Height = Me.ScaleHeight
    
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 13
        .AlphaFormat = 0
    End With

End Sub
