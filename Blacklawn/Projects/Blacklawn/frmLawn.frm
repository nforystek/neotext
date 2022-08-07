
VERSION 5.00
Begin VB.Form frmLawn 
   Caption         =   "Blacklawn"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15075
   Icon            =   "frmLawn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15075
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   11115
      Left            =   0
      ScaleHeight     =   11055
      ScaleWidth      =   15045
      TabIndex        =   0
      Top             =   45
      Width           =   15105
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1815
         Top             =   1530
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   930
         Top             =   1545
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000C000&
         X1              =   11595
         X2              =   13050
         Y1              =   1815
         Y2              =   1440
      End
      Begin VB.Line Line15 
         BorderColor     =   &H0000C000&
         X1              =   11370
         X2              =   11565
         Y1              =   3495
         Y2              =   1845
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0000C000&
         X1              =   11310
         X2              =   9960
         Y1              =   3540
         Y2              =   3330
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000C000&
         X1              =   8925
         X2              =   9960
         Y1              =   2010
         Y2              =   3300
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000C000&
         X1              =   5970
         X2              =   8925
         Y1              =   2205
         Y2              =   1980
      End
      Begin VB.Line Line11 
         BorderColor     =   &H0000C000&
         X1              =   5565
         X2              =   5940
         Y1              =   3855
         Y2              =   2235
      End
      Begin VB.Line Line10 
         BorderColor     =   &H0000C000&
         X1              =   4695
         X2              =   5550
         Y1              =   5475
         Y2              =   3870
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000C000&
         X1              =   4710
         X2              =   4935
         Y1              =   5505
         Y2              =   6720
      End
      Begin VB.Line Line8 
         BorderColor     =   &H0000C000&
         X1              =   5745
         X2              =   4980
         Y1              =   6315
         Y2              =   6720
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000C000&
         X1              =   5820
         X2              =   6840
         Y1              =   6330
         Y2              =   6855
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000C000&
         X1              =   6900
         X2              =   6825
         Y1              =   8370
         Y2              =   6795
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000C000&
         X1              =   10710
         X2              =   6930
         Y1              =   8580
         Y2              =   8385
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000C000&
         X1              =   12675
         X2              =   10740
         Y1              =   7755
         Y2              =   8550
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         X1              =   13140
         X2              =   12705
         Y1              =   5490
         Y2              =   7785
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         X1              =   14460
         X2              =   13125
         Y1              =   2895
         Y2              =   5445
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         X1              =   13095
         X2              =   14460
         Y1              =   1455
         Y2              =   2865
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
         Left            =   705
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
         Left            =   8385
         TabIndex        =   3
         Top             =   9075
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
         Height          =   420
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   8850
      End
      Begin VB.Line sBRight 
         BorderColor     =   &H0000C0C0&
         X1              =   4230
         X2              =   4095
         Y1              =   3240
         Y2              =   3075
      End
      Begin VB.Line sBLeft 
         BorderColor     =   &H0000C0C0&
         X1              =   3690
         X2              =   3900
         Y1              =   3225
         Y2              =   3060
      End
      Begin VB.Line sTRight 
         BorderColor     =   &H00008080&
         X1              =   4125
         X2              =   4305
         Y1              =   2355
         Y2              =   2880
      End
      Begin VB.Line sTLeft 
         BorderColor     =   &H00008080&
         X1              =   3945
         X2              =   3690
         Y1              =   2400
         Y2              =   2895
      End
      Begin VB.Line shipBase 
         BorderColor     =   &H00000000&
         X1              =   4035
         X2              =   4020
         Y1              =   2400
         Y2              =   2805
      End
      Begin VB.Shape lawnBase 
         BackColor       =   &H00008000&
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   8700
         Shape           =   3  'Circle
         Top             =   4710
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmLawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

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
Private Const Bit12 = 4096

Private dKeyCode As Integer

Private Const TurnSpeed = 4

Private Type FireBullet
    X As Long
    Y As Long
    vX As Long
    vY As Long
    Degree As Integer
    Life As Integer
End Type

Private shipFireBullets() As FireBullet
Private Const MaxFireLife = 100
Private Const FireSpeed = 15
Private Const FireLag = 10
Private fireLagCount As Integer

Private PlayerBounds As Long

Private sSpeedX As Integer
Private sSpeedY As Integer

Private xPixel As Integer
Private yPixel As Integer

Private shipX As Long
Private shipY As Long

Private shipDegree As Integer

Private sStart As String
Private sState As Integer

Private sScore1 As Long
Private sScore2 As Long
Private sScore3 As Long

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
    Dim r As Long
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
            
            
                r = FireSpeed
                A = shipFireBullets(cnt).Degree / 57
                
                X = (r * Cos(A))
                Y = (r * Sin(A))
                
            
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
    Dim r As Long
    Dim A
    Dim X
    Dim Y
        
    r = 20
    
    A = shipDegree / 57
    
    X = (r * Cos(A)) * xPixel
    Y = (r * Sin(A)) * yPixel

    shipBase.X1 = (Picture1.ScaleWidth / 2) + 80 + X
    shipBase.Y1 = (Picture1.ScaleHeight / 2) + 80 + Y
    
    shipBase.X2 = (Picture1.ScaleWidth / 2) + 80
    shipBase.Y2 = (Picture1.ScaleHeight / 2) + 80
        
    
    r = 10
    
    sTLeft.X1 = shipBase.X1
    sTRight.X1 = shipBase.X1
    
    sTLeft.Y1 = shipBase.Y1
    sTRight.Y1 = shipBase.Y1
    
    sBLeft.X1 = shipBase.X2
    sBRight.X1 = shipBase.X2
    
    sBLeft.Y1 = shipBase.Y2
    sBRight.Y1 = shipBase.Y2
    
    A = (shipDegree + 140) / 57
    X = (r * Cos(A)) * xPixel
    Y = (r * Sin(A)) * yPixel
    
    sTLeft.X2 = shipBase.X2 + X
    sTLeft.Y2 = shipBase.Y2 + Y
    
    sBLeft.X2 = shipBase.X2 + X
    sBLeft.Y2 = shipBase.Y2 + Y
    
    A = (shipDegree + 220) / 57
    X = (r * Cos(A)) * xPixel
    Y = (r * Sin(A)) * yPixel
    
    sTRight.X2 = shipBase.X2 + X
    sTRight.Y2 = shipBase.Y2 + Y
    
    sBRight.X2 = shipBase.X2 + X
    sBRight.Y2 = shipBase.Y2 + Y
    
End Sub

Private Sub PaintLawn()
    Dim xDiff As Long
    Dim yDiff As Long
    xDiff = lawnBase.Left
    yDiff = lawnBase.Top
    
    lawnBase.Top = ((Picture1.ScaleHeight / 2) - (lawnBase.height / 2)) + shipY
    lawnBase.Left = ((Picture1.ScaleWidth / 2) - (lawnBase.width / 2)) + shipX
    
    xDiff = xDiff - lawnBase.Left
    yDiff = yDiff - lawnBase.Top
    
    Line1.X1 = Line1.X1 - xDiff
    Line1.X2 = Line1.X2 - xDiff
    Line1.Y1 = Line1.Y1 - yDiff
    Line1.Y2 = Line1.Y2 - yDiff
    
    Line2.X1 = Line2.X1 - xDiff
    Line2.X2 = Line2.X2 - xDiff
    Line2.Y1 = Line2.Y1 - yDiff
    Line2.Y2 = Line2.Y2 - yDiff
    
    Line3.X1 = Line3.X1 - xDiff
    Line3.X2 = Line3.X2 - xDiff
    Line3.Y1 = Line3.Y1 - yDiff
    Line3.Y2 = Line3.Y2 - yDiff
    
    Line4.X1 = Line4.X1 - xDiff
    Line4.X2 = Line4.X2 - xDiff
    Line4.Y1 = Line4.Y1 - yDiff
    Line4.Y2 = Line4.Y2 - yDiff
    
    Line5.X1 = Line5.X1 - xDiff
    Line5.X2 = Line5.X2 - xDiff
    Line5.Y1 = Line5.Y1 - yDiff
    Line5.Y2 = Line5.Y2 - yDiff
    
    Line6.X1 = Line6.X1 - xDiff
    Line6.X2 = Line6.X2 - xDiff
    Line6.Y1 = Line6.Y1 - yDiff
    Line6.Y2 = Line6.Y2 - yDiff
    
    Line7.X1 = Line7.X1 - xDiff
    Line7.X2 = Line7.X2 - xDiff
    Line7.Y1 = Line7.Y1 - yDiff
    Line7.Y2 = Line7.Y2 - yDiff
    
    Line8.X1 = Line8.X1 - xDiff
    Line8.X2 = Line8.X2 - xDiff
    Line8.Y1 = Line8.Y1 - yDiff
    Line8.Y2 = Line8.Y2 - yDiff
    
    Line9.X1 = Line9.X1 - xDiff
    Line9.X2 = Line9.X2 - xDiff
    Line9.Y1 = Line9.Y1 - yDiff
    Line9.Y2 = Line9.Y2 - yDiff
    
    Line10.X1 = Line10.X1 - xDiff
    Line10.X2 = Line10.X2 - xDiff
    Line10.Y1 = Line10.Y1 - yDiff
    Line10.Y2 = Line10.Y2 - yDiff
    
    Line11.X1 = Line11.X1 - xDiff
    Line11.X2 = Line11.X2 - xDiff
    Line11.Y1 = Line11.Y1 - yDiff
    Line11.Y2 = Line11.Y2 - yDiff
    
    Line12.X1 = Line12.X1 - xDiff
    Line12.X2 = Line12.X2 - xDiff
    Line12.Y1 = Line12.Y1 - yDiff
    Line12.Y2 = Line12.Y2 - yDiff
    
    Line13.X1 = Line13.X1 - xDiff
    Line13.X2 = Line13.X2 - xDiff
    Line13.Y1 = Line13.Y1 - yDiff
    Line13.Y2 = Line13.Y2 - yDiff
    
    Line14.X1 = Line14.X1 - xDiff
    Line14.X2 = Line14.X2 - xDiff
    Line14.Y1 = Line14.Y1 - yDiff
    Line14.Y2 = Line14.Y2 - yDiff
    
    Line15.X1 = Line15.X1 - xDiff
    Line15.X2 = Line15.X2 - xDiff
    Line15.Y1 = Line15.Y1 - yDiff
    Line15.Y2 = Line15.Y2 - yDiff
    
    Line16.X1 = Line16.X1 - xDiff
    Line16.X2 = Line16.X2 - xDiff
    Line16.Y1 = Line16.Y1 - yDiff
    Line16.Y2 = Line16.Y2 - yDiff
    
End Sub

Private Sub Form_Load()
    Me.width = (NextArg(Resolution, "x") * Screen.TwipsPerPixelX)
    Me.height = (RemoveArg(Resolution, "x") * Screen.TwipsPerPixelY)
    
    
    Dim db As clsDatabase
    Set db = New clsDatabase
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    db.rsQuery rs, "SELECT Score1, Score2, Score3 FROM Settings WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"

    If Not db.rsEnd(rs) Then
        sScore1 = rs("Score1")
        sScore2 = rs("Score2")
        sScore3 = rs("Score3")
    End If

    db.rsClose rs

    Set rs = Nothing
    Set db = Nothing
    
    sScore1 = 0
    sScore2 = 0
    sScore3 = 0
    
    PlayerBounds = 1000000
    
    
    Label1.Caption = " UP=Forward, DOWN=Reverse, LEFT/RIGHT=Rotate, B=Breaks, SPACEBAR=Act Tough, N=Easy, M=Hard "
    
    Label2.Caption = "A mythical black act returning to start.   " & vbCrLf & _
                    "1 = Warp random to a pitch black sight.   " & vbCrLf & _
                    "2 = Ride out to pitch black by your self.   " & vbCrLf & _
                    "0 = Forfeit timed score, and warp back.   " & vbCrLf
    
    SetState 0

    xPixel = Screen.TwipsPerPixelX
    yPixel = Screen.TwipsPerPixelY
    
    shipX = 0
    shipY = 0
    
    PaintLawn
    
    sState = 0
    
    shipDegree = 0
    
    ReDim shipFireBullets(0) As FireBullet
    
    Set Track1 = New clsAmbient
    
    Track1.FileName = AppPath & "Base\music1.mp3"
    Track1.LoopEnabled = True
    
    If PlayMusic Then Track1.PlaySound
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Picture1.Top = 0
        Picture1.Left = 0
        Picture1.width = Me.ScaleWidth
        Picture1.height = Me.ScaleHeight
        
        Label1.Top = 0
        Label1.Left = 0
        
        Label2.Left = Me.ScaleWidth - Label2.width
        Label2.Top = 0
        
        Label3.Top = Me.ScaleHeight - Label3.height
        Label3.Left = (Me.ScaleWidth / 2) - (Label3.width / 2)
        
        Label4.Top = Me.ScaleHeight - (Label4.height * 5)
        Label4.Left = (Me.ScaleWidth / 2) - (Label4.width / 2)
        
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
    Dim r As Long
    Dim A
    Dim X
    Dim Y
    
    r = 1
    
    A = shipDegree / 57
    
    X = (r * Cos(A))
    Y = (r * Sin(A))
    
    sSpeedX = sSpeedX - X
    sSpeedY = sSpeedY - Y
    
    If sSpeedX > 300 Then sSpeedX = 300
    If sSpeedX < -300 Then sSpeedX = -300
    If sSpeedY > 300 Then sSpeedY = 300
    If sSpeedY < -300 Then sSpeedY = -300
End Sub

Private Sub KeyDownPressed()
    Dim r As Long
    Dim A
    Dim X
    Dim Y
    
    r = 1
    
    A = shipDegree / 57
    
    X = (r * Cos(A))
    Y = (r * Sin(A))
    
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
    If PlayMusic Then Track1.StopSound
    Set Track1 = Nothing

    Dim db As clsDatabase
    Set db = New clsDatabase
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    db.dbQuery "UPDATE Settings SET Score1=" & sScore1 & ", Score2=" & sScore2 & ", Score3=" & sScore3 & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"

    Set rs = Nothing
    Set db = Nothing
    End
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
    If KeyCode = 49 Then
        BitValue(dKeyCode, Bit8) = True
    End If
    If KeyCode = 50 Then
        BitValue(dKeyCode, Bit9) = True
    End If
    If KeyCode = 48 Then
        BitValue(dKeyCode, Bit10) = True
    End If
    If KeyCode = 78 Then
        BitValue(dKeyCode, Bit11) = True
    End If
    If KeyCode = 77 Then
        BitValue(dKeyCode, Bit12) = True
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
    If KeyCode = 49 Then
        BitValue(dKeyCode, Bit8) = False
    End If
    If KeyCode = 50 Then
        BitValue(dKeyCode, Bit9) = False
    End If
    If KeyCode = 48 Then
        BitValue(dKeyCode, Bit10) = False
    End If
    If KeyCode = 78 Then
        BitValue(dKeyCode, Bit11) = False
    End If
    If KeyCode = 77 Then
        BitValue(dKeyCode, Bit12) = False
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
        If BitValue(dKeyCode, Bit8) Then
            sScore1 = 0
            Randomize
            shipY = IIf((Rnd < 0.5), -RandomPositive(15000, PlayerBounds), RandomPositive(15000, PlayerBounds))
            Randomize
            shipX = IIf((Rnd < 0.5), -RandomPositive(15000, PlayerBounds), RandomPositive(15000, PlayerBounds))
            KeyBPressed
            SetState 3
        End If
        If BitValue(dKeyCode, Bit9) Then
            sScore2 = 0
            sScore3 = 0
            shipY = 0
            shipX = 0
            KeyBPressed
            SetState 1
        End If
        If BitValue(dKeyCode, Bit10) Then
            shipY = 0
            shipX = 0
            KeyBPressed
            SetState -1
        End If
        If BitValue(dKeyCode, Bit11) Then
            PlayerBounds = 150000
        End If
        If BitValue(dKeyCode, Bit12) Then
            PlayerBounds = 1000000
        End If
        
    If fireLagCount > 0 Then
        fireLagCount = fireLagCount + 1
        If fireLagCount >= FireLag Then fireLagCount = 0
    End If

    shipX = shipX + sSpeedX
    shipY = shipY + sSpeedY

    If shipX > PlayerBounds Then shipX = PlayerBounds
    If shipX < -PlayerBounds Then shipX = -PlayerBounds
    If shipY > PlayerBounds Then shipY = PlayerBounds
    If shipY < -PlayerBounds Then shipY = -PlayerBounds

    PaintLawn
    
    PaintShip

    PaintFire

    Select Case sState
        Case 2, 3
            If Distance <= 2000 Then
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
            Label4.Caption = "The pitch black clock has begun, you either return for your score or give up zero."
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
    
    nScore = "Score: 1st " & Trim(sScore1) & "s - 2nd " & Trim(sScore2) & "s " & Trim(sScore3) & "s"

    Label3.Caption = nScore
    
#If VBIDE = -1 Then
    
    Label1.Caption = " UP=Forward, DOWN=Reverse, LEFT/RIGHT=Rotate, B=Breaks, SPACEBAR=Act Tough, N=Easy, M=Hard " & vbCrLf & " Distance: " & Dressing(Distance)
#End If
End Sub
Public Function Dressing(ByVal val) As String
    Dim tmp As String
    tmp = val & ""
    If InStr(tmp, ".") > 0 Then
        tmp = Left(tmp, InStr(tmp, ".") - 1)
    End If
    Dressing = tmp
End Function
Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Label4.Caption = ""
End Sub
