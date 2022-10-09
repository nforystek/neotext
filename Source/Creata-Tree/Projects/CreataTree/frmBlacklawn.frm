VERSION 5.00
Begin VB.Form frmBlacklawn 
   Caption         =   "The Blacklawn"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   8310
      Left            =   0
      ScaleHeight     =   8250
      ScaleWidth      =   12180
      TabIndex        =   0
      Top             =   0
      Width           =   12240
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   7005
         Top             =   1170
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4980
         Top             =   840
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   750
         TabIndex        =   4
         Top             =   5280
         Width           =   7215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Score: 1st 0(s) - 2nd 0(s) 0(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   3
         Top             =   6525
         Width           =   3120
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   5760
         TabIndex        =   2
         Top             =   30
         Width           =   3555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   7035
      End
      Begin VB.Shape lawnBase 
         BackColor       =   &H00008000&
         BorderColor     =   &H0000C000&
         Height          =   1695
         Left            =   3675
         Shape           =   3  'Circle
         Top             =   3450
         Width           =   1410
      End
      Begin VB.Line sBRight 
         BorderColor     =   &H0000FFFF&
         X1              =   4230
         X2              =   4095
         Y1              =   3240
         Y2              =   3075
      End
      Begin VB.Line sBLeft 
         BorderColor     =   &H0000FFFF&
         X1              =   3690
         X2              =   3900
         Y1              =   3225
         Y2              =   3060
      End
      Begin VB.Line sTRight 
         BorderColor     =   &H0000FFFF&
         X1              =   4125
         X2              =   4305
         Y1              =   2355
         Y2              =   2880
      End
      Begin VB.Line sTLeft 
         BorderColor     =   &H0000FFFF&
         X1              =   3945
         X2              =   3690
         Y1              =   2400
         Y2              =   2895
      End
      Begin VB.Line shipBase 
         BorderColor     =   &H00000000&
         X1              =   5820
         X2              =   5805
         Y1              =   2715
         Y2              =   3120
      End
   End
End
Attribute VB_Name = "frmBlacklawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Const Bit1 = 1
Private Const Bit2 = 2
Private Const Bit3 = 4
Private Const Bit4 = 8
Private Const Bit5 = 16
Private Const Bit6 = 32
Private Const Bit7 = 64
Private Const Bit8 = 128
Private Const Bit9 = 256
Private Const Bit10 = 1024
Private Const Bit11 = 2048

Private dKeyCode As Integer

Private Const TurnSpeed = 4 'The speed the ship turns at

Private Type FireBullet
    X As Long
    Y As Long
    vX As Long
    vY As Long
    Degree As Integer
    Life As Integer
End Type

Private shipFireBullets() As FireBullet
Private Const MaxFireLife = 100 'How long a bullet fires
Private Const FireSpeed = 15 'Increase to make bullets move faster
Private Const FireLag = 10 'Lag on how fast the ship can shot
Private fireLagCount As Integer

Private sSpeedX As Integer
Private sSpeedY As Integer

Private xPixel As Integer 'Set in main, number of twips per pixel X
Private yPixel As Integer 'Set in main, number of twips per pixel Y

Private shipX As Long 'X coordinate of the ship
Private shipY As Long 'Y coordinate of the ship

Private shipDegree As Integer 'The direction angle that the ship is pointing in

Private sStart As String
Private sState As Integer
Private sScore1 As Long
Private sScore2 As Long
Private sScore3 As Long


Private Function MciCommand(sCommand As String) As String
    mciSendString sCommand, 0&, 0, 0
End Function

Private Sub StopMIDI()
    MciCommand "stop tlst"
    MciCommand "close tlst"

End Sub
Private Sub PlayMIDI()
    
    MciCommand "open """ & AppPath & "Engine\tlst.mid" & """ alias tlst"
    MciCommand "play tlst"
End Sub


Private Sub NewFire()
    
    If fireLagCount = 0 Then
        fireLagCount = 1
        
        Dim cnt As Integer
        cnt = UBound(shipFireBullets) + 1
        ReDim Preserve shipFireBullets(cnt) As FireBullet
    
        shipFireBullets(cnt).X = shipBase.X1
        shipFireBullets(cnt).Y = shipBase.Y1
        shipFireBullets(cnt).vX = sSpeedX
        shipFireBullets(cnt).vY = sSpeedY
        shipFireBullets(cnt).Degree = shipDegree
        shipFireBullets(cnt).Life = 0
    End If

End Sub
Private Sub RemoveBulletsIndex(ByRef ListArray() As FireBullet, ByVal Index As Integer)
    Dim cnt As Integer
    cnt = 0
    Do
        If cnt > Index Then
            ListArray(cnt - 1) = ListArray(cnt)
            End If
        cnt = cnt + 1
        Loop Until cnt > UBound(ListArray)
    If Not UBound(ListArray) <= 0 Then
        ReDim Preserve ListArray(UBound(ListArray) - 1) As FireBullet
    End If
End Sub

Private Sub PaintFire()
    Dim R As Long
    Dim A
    Dim X
    Dim Y
    
    Dim cnt As Integer
    If UBound(shipFireBullets) > 0 Then
        cnt = 1
        Do
            
            Picture1.PSet (shipFireBullets(cnt).X, shipFireBullets(cnt).Y), RGB(0, 0, 0)
            
            If shipFireBullets(cnt).Life <= MaxFireLife Then
                
                shipFireBullets(cnt).Life = shipFireBullets(cnt).Life + 1
            
            
                R = FireSpeed
                A = shipFireBullets(cnt).Degree / 57
                
                X = (R * Cos(A))
                Y = (R * Sin(A))
                
            
                shipFireBullets(cnt).X = shipFireBullets(cnt).X + X + shipFireBullets(cnt).vX
                shipFireBullets(cnt).Y = shipFireBullets(cnt).Y + Y + shipFireBullets(cnt).vY
            
                If shipFireBullets(cnt).X < 0 Then shipFireBullets(cnt).X = Picture1.ScaleWidth
                If shipFireBullets(cnt).X > Picture1.ScaleWidth Then shipFireBullets(cnt).X = 0
    
                If shipFireBullets(cnt).Y < 0 Then shipFireBullets(cnt).Y = Picture1.ScaleHeight
                If shipFireBullets(cnt).Y > Picture1.ScaleHeight Then shipFireBullets(cnt).Y = 0
                
                Picture1.PSet (shipFireBullets(cnt).X, shipFireBullets(cnt).Y), RGB(255, 255, 255)
                
            Else
                RemoveBulletsIndex shipFireBullets, cnt
            End If
            
            cnt = cnt + 1
            
        Loop Until cnt > UBound(shipFireBullets)
    End If

End Sub
Private Sub PaintShip()
    Dim R As Long
    Dim A
    Dim X
    Dim Y
        
    R = 20
    
    A = shipDegree / 57
    
    X = (R * Cos(A)) * xPixel
    Y = (R * Sin(A)) * yPixel

    shipBase.X1 = (Picture1.ScaleWidth / 2) + 80 + X
    shipBase.Y1 = (Picture1.ScaleHeight / 2) + 80 + Y
    
    shipBase.X2 = (Picture1.ScaleWidth / 2) + 80
    shipBase.Y2 = (Picture1.ScaleHeight / 2) + 80
        
    
    R = 10
    
    sTLeft.X1 = shipBase.X1
    sTRight.X1 = shipBase.X1
    
    sTLeft.Y1 = shipBase.Y1
    sTRight.Y1 = shipBase.Y1
    
    sBLeft.X1 = shipBase.X2
    sBRight.X1 = shipBase.X2
    
    sBLeft.Y1 = shipBase.Y2
    sBRight.Y1 = shipBase.Y2
    
    A = (shipDegree + 140) / 57
    X = (R * Cos(A)) * xPixel
    Y = (R * Sin(A)) * yPixel
    
    sTLeft.X2 = shipBase.X2 + X
    sTLeft.Y2 = shipBase.Y2 + Y
    
    sBLeft.X2 = shipBase.X2 + X
    sBLeft.Y2 = shipBase.Y2 + Y
    
    A = (shipDegree + 220) / 57
    X = (R * Cos(A)) * xPixel
    Y = (R * Sin(A)) * yPixel
    
    sTRight.X2 = shipBase.X2 + X
    sTRight.Y2 = shipBase.Y2 + Y
    
    sBRight.X2 = shipBase.X2 + X
    sBRight.Y2 = shipBase.Y2 + Y
    
End Sub

Private Sub PaintLawn()

    lawnBase.Top = ((Picture1.ScaleHeight / 2) - (lawnBase.Height / 2)) + shipY
    lawnBase.Left = ((Picture1.ScaleWidth / 2) - (lawnBase.Width / 2)) + shipX
    
End Sub

Private Sub Form_Load()
    If Ini.Exists("Score1") Then sScore1 = Ini.Setting("Score1")
    If Ini.Exists("Score2") Then sScore2 = Ini.Setting("Score2")
    If Ini.Exists("Score3") Then sScore3 = Ini.Setting("Score3")
    
    Label1.Caption = " UP=Forward, DOWN=Reverse, LEFT/RIGHT=Rotate, B=Breaks, SPACEBAR=Act Tough"
    
    Label2.Caption = "A mythical black act returning to start.   " & vbCrLf & _
                    "1 = Warp random to a pitch black sight.   " & vbCrLf & _
                    "2 = Ride out to pitch black by your self.   " & vbCrLf & _
                    "0 = Forfeit timed score, and warp back.   " & vbCrLf
    
    SetState 0

    xPixel = Screen.TwipsPerPixelX
    yPixel = Screen.TwipsPerPixelY
    
    shipX = 0
    shipY = 0
    
    lawnBase.Top = ((Picture1.ScaleHeight / 2) - (lawnBase.Height / 2)) + shipY
    lawnBase.Left = ((Picture1.ScaleWidth / 2) - (lawnBase.Width / 2)) + shipX
    
    sState = 0
    
    shipDegree = 0
    
    ReDim shipFireBullets(0) As FireBullet
    
    PlayMIDI

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Picture1.Top = 0
        Picture1.Left = 0
        Picture1.Width = Me.ScaleWidth
        Picture1.Height = Me.ScaleHeight
        
        Label1.Top = 0
        Label1.Left = 0
        
        Label2.Left = Me.ScaleWidth - Label2.Width
        Label2.Top = 0
        
        Label3.Top = Me.ScaleHeight - Label3.Height
        Label3.Left = (Me.ScaleWidth / 2) - (Label3.Width / 2)
        
        Label4.Top = Me.ScaleHeight - (Label4.Height * 5)
        Label4.Left = (Me.ScaleWidth / 2) - (Label4.Width / 2)
        
        PaintLawn
    End If
End Sub

Public Function RollRight() As Boolean
    shipDegree = shipDegree + TurnSpeed
    If shipDegree > 359 Then shipDegree = 0
End Function

Public Function RollLeft() As Boolean
    shipDegree = shipDegree - TurnSpeed
    If shipDegree < 0 Then shipDegree = 359
End Function

Private Sub KeyUpPressed()
    Dim R As Long
    Dim A
    Dim X
    Dim Y
    
    R = 1
    
    A = shipDegree / 57
    
    X = (R * Cos(A))
    Y = (R * Sin(A))
    
    sSpeedX = sSpeedX - X
    sSpeedY = sSpeedY - Y
    
    If sSpeedX > 300 Then sSpeedX = 300
    If sSpeedX < -300 Then sSpeedX = -300
    If sSpeedY > 300 Then sSpeedY = 300
    If sSpeedY < -300 Then sSpeedY = -300
End Sub

Private Sub KeyDownPressed()
    Dim R As Long
    Dim A
    Dim X
    Dim Y
    
    R = 1
    
    A = shipDegree / 57
    
    X = (R * Cos(A))
    Y = (R * Sin(A))
    
    sSpeedX = sSpeedX + X
    sSpeedY = sSpeedY + Y

    If sSpeedX > 300 Then sSpeedX = 300
    If sSpeedX < -300 Then sSpeedX = -300
    If sSpeedY > 300 Then sSpeedY = 300
    If sSpeedY < -300 Then sSpeedY = -300
End Sub

Private Sub KeyLeftPressed()
    RollLeft
End Sub

Private Sub KeyRightPressed()
    RollRight
End Sub

Private Sub KeySpacePressed()
    NewFire
End Sub
Private Sub KeyBPressed()
    sSpeedX = 0
    sSpeedY = 0
End Sub

Private Property Let BitValue(ByRef bWord As Integer, ByVal bBit As Integer, ByVal nValue As Boolean)
    If (bWord And bBit) And (Not nValue) Then
        bWord = bWord - bBit
    ElseIf (Not (bWord And bBit)) And nValue Then
        bWord = bWord Or bBit
    End If
End Property

Private Property Get BitValue(ByRef bWord As Integer, ByVal bBit As Integer) As Boolean
    BitValue = (bWord And bBit)
End Property


Private Sub Form_Unload(Cancel As Integer)
    If Ini.Exists("Score1") Then Ini.Remove ("Score1")
    If Ini.Exists("Score2") Then Ini.Remove ("Score2")
    If Ini.Exists("Score3") Then Ini.Remove ("Score3")
    Ini.Add "Score1", sScore1
    Ini.Add "Score2", sScore2
    Ini.Add "Score3", sScore3
    
    StopMIDI
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = 38 Then
        BitValue(dKeyCode, Bit1) = True
    End If
    If KeyCode = 40 Then
        BitValue(dKeyCode, Bit2) = True
    End If
    If KeyCode = 39 Then
        BitValue(dKeyCode, Bit3) = True
    End If
    If KeyCode = 37 Then
        BitValue(dKeyCode, Bit4) = True
    End If
    If KeyCode = 32 Then
        BitValue(dKeyCode, Bit5) = True
    End If
    If KeyCode = 66 Then
        BitValue(dKeyCode, Bit6) = True
    End If
    If KeyCode = 27 Then
        BitValue(dKeyCode, Bit7) = True
    End If
    If KeyCode = 49 Then '1
        BitValue(dKeyCode, Bit8) = True
    End If
    If KeyCode = 50 Then '2
        BitValue(dKeyCode, Bit9) = True
    End If
    If KeyCode = 48 Then '0
        BitValue(dKeyCode, Bit10) = True
    End If

End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Then
        BitValue(dKeyCode, Bit1) = False
    End If
    If KeyCode = 40 Then
        BitValue(dKeyCode, Bit2) = False
    End If
    If KeyCode = 39 Then
        BitValue(dKeyCode, Bit3) = False
    End If
    If KeyCode = 37 Then
        BitValue(dKeyCode, Bit4) = False
    End If
    If KeyCode = 32 Then
        BitValue(dKeyCode, Bit5) = False
    End If
    If KeyCode = 66 Then
        BitValue(dKeyCode, Bit6) = False
    End If
    If KeyCode = 27 Then
        BitValue(dKeyCode, Bit7) = False
    End If
    If KeyCode = 49 Then '1
        BitValue(dKeyCode, Bit8) = False
    End If
    If KeyCode = 50 Then '2
        BitValue(dKeyCode, Bit9) = False
    End If
    If KeyCode = 48 Then '0
        BitValue(dKeyCode, Bit10) = False
    End If
End Sub



Private Sub Timer1_Timer()

        If BitValue(dKeyCode, Bit1) Then
            KeyUpPressed
        End If
        If BitValue(dKeyCode, Bit2) Then
            KeyDownPressed
        End If
        If BitValue(dKeyCode, Bit3) Then
            KeyRightPressed
        End If
        If BitValue(dKeyCode, Bit4) Then
            KeyLeftPressed
        End If
        If BitValue(dKeyCode, Bit5) Then
            KeySpacePressed
        End If
        If BitValue(dKeyCode, Bit6) Then
            KeyBPressed
        End If
        If BitValue(dKeyCode, Bit7) Then
            Unload Me
        End If
        If BitValue(dKeyCode, Bit8) Then '1
            sScore1 = 0
            Randomize
            shipY = Int((200000 - 15000 + 1) * Rnd) + 15000
            Randomize
            shipX = Int((200000 - 15000 + 1) * Rnd) + 15000
            KeyBPressed
            SetState 3
        End If
        If BitValue(dKeyCode, Bit9) Then '2
            sScore2 = 0
            sScore3 = 0
            shipY = 0
            shipX = 0
            KeyBPressed
            SetState 1
        End If
        If BitValue(dKeyCode, Bit10) Then '0
            shipY = 0
            shipX = 0
            KeyBPressed
            SetState -1
        End If
        
    If fireLagCount > 0 Then
        fireLagCount = fireLagCount + 1
        If fireLagCount >= FireLag Then fireLagCount = 0
    End If

    shipX = shipX + sSpeedX
    shipY = shipY + sSpeedY

    If shipX > 2000000 Then shipX = 2000000
    If shipX < -2000000 Then shipX = -2000000
    If shipY > 2000000 Then shipY = 2000000
    If shipY < -2000000 Then shipY = -2000000

    PaintLawn
    
    PaintShip

    PaintFire

    Select Case sState
        Case 2, 3
            If Distance <= 700 Then
                SetState 4
            End If
        Case 1
            If Distance >= 15000 Then
                SetState 2
            End If
    End Select
    
    DisplayScore
    
End Sub

Private Function Distance() As Double
    Distance = Sqr(((shipX - 0) ^ 2) + ((shipY - 0) ^ 2))
End Function

Private Sub SetState(ByVal nState As Integer)
    Select Case nState
        Case -1
            Label4.Caption = "You have forfeit your score and no longer have a legit arrival time."
            Timer2.Enabled = True
            If Not (sStart = "") Then
                Select Case sState
                    Case 1, 2
                        sScore2 = 0
                        sScore3 = 0
                    Case 3
                        sScore1 = 0
                End Select
            End If
            sStart = ""
        Case 0
            Label4.Caption = ""
            sStart = ""
        Case 1
            Label4.Caption = "Penalty clock started, run from the circle far enough to start the pitch black clock."
            Timer2.Enabled = True
            sStart = Now
        Case 2, 3
            Label4.Caption = "The pitch black clock has begun, you either return for your score or have a give up."
            Timer2.Enabled = True
            sStart = Now
        Case 4
            Label4.Caption = "Positive issue, the black movie, pushing space ships like, done."
            Timer2.Enabled = True
            sStart = ""
    End Select
    sState = nState
End Sub

Private Sub DisplayScore()
    
    Select Case sState
        Case 1
            sScore2 = DateDiff("s", CDate(sStart), Now)
        Case 2
            sScore3 = DateDiff("s", CDate(sStart), Now)
        Case 3
            sScore1 = DateDiff("s", CDate(sStart), Now)
    End Select
    
    Dim nScore As String
    
    nScore = "Score: 1st " & Trim(sScore1) & "(s) - 2nd " & Trim(sScore2) & "(s) " & Trim(sScore3) & "(s)"

    Label3.Caption = nScore
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Label4.Caption = ""
End Sub
