VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotation"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   2280
      Top             =   4080
   End
   Begin VB.PictureBox Picture3 
      Height          =   3960
      Left            =   8280
      ScaleHeight     =   3900
      ScaleWidth      =   3900
      TabIndex        =   2
      Top             =   120
      Width           =   3960
   End
   Begin VB.PictureBox Picture2 
      Height          =   3960
      Left            =   4200
      ScaleHeight     =   3900
      ScaleWidth      =   3900
      TabIndex        =   1
      Top             =   120
      Width           =   3960
   End
   Begin VB.PictureBox Picture1 
      Height          =   3960
      Left            =   120
      ScaleHeight     =   3900
      ScaleWidth      =   3900
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "Use any of three buttons on a mouse clicking on any of the canvas above to move the rotaitons around."
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    Me.MousePointer = 2
    
    Form_Paint
End Sub

Private Sub Form_Paint()
    DrawAxis "X Axis", Picture1, Vertex.Y, Vertex.Z, Angles.X, Vector.Y, Vector.Z, Rotate.X, points.Y, points.Z, twists.X
    DrawAxis "Y Axis", Picture2, Vertex.Z, Vertex.X, Angles.Y, Vector.Z, Vector.X, Rotate.Y, points.Z, points.X, twists.Y
    DrawAxis "Z Axis", Picture3, Vertex.X, Vertex.Y, Angles.Z, Vector.X, Vector.Y, Rotate.Z, points.X, points.Y, twists.Z
End Sub

Private Sub DrawAxis(ByVal Axis As String, ByRef PicBox As PictureBox, ByVal pX As Single, ByVal pY As Single, ByVal aZ As Single, ByVal vX As Single, ByVal vY As Single, ByVal rZ As Single, ByVal tX As Single, ByVal tY As Single, ByVal dZ As Single)
    'draw z axis
    PicBox.Cls
    PicBox.CurrentX = ((Screen.TwipsPerPixelX) * 4)
    PicBox.CurrentY = ((Screen.TwipsPerPixelY) * 4)
    PicBox.Print Axis
    Select Case Axis
        Case "X Axis"
            PicBox.CurrentX = ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) - Form1.TextHeight("Y") - ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "Y"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "Z"
        Case "Y Axis"
            PicBox.CurrentX = ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) - Form1.TextHeight("Y") - ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "Z"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "X"
        Case "Z Axis"
            PicBox.CurrentX = ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) - Form1.TextHeight("Y") - ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "X"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + ((Screen.TwipsPerPixelX) * 4)
            PicBox.CurrentY = ((Screen.TwipsPerPixelY) * 4)
            PicBox.Print "Y"
    End Select
    PicBox.Line (0, (PicBox.ScaleHeight / 2))-(PicBox.ScaleWidth, (PicBox.ScaleHeight / 2)), &H808080
    PicBox.Line ((PicBox.ScaleWidth / 2), 0)-((PicBox.ScaleWidth / 2), PicBox.ScaleHeight), &H808080

    PicBox.Circle (pX + (PicBox.ScaleWidth / 2), pY + (PicBox.ScaleHeight / 2)), 100, &H8000&
    PicBox.Circle (vX + (PicBox.ScaleWidth / 2), vY + (PicBox.ScaleHeight / 2)), 100, vbBlue
    PicBox.Circle (tX + (PicBox.ScaleWidth / 2), tY + (PicBox.ScaleHeight / 2)), 100, vbRed
    
    Dim Distance As Single
    aZ = AngleRestrict(aZ)
    Distance = (((pX ^ 2) + (pY ^ 2)) ^ (1 / 2))
    PicBox.Line ((PicBox.ScaleWidth / 2), (PicBox.ScaleHeight / 2))- _
            ((PicBox.ScaleWidth / 2) + (Distance * Sin(aZ)), _
            ((PicBox.ScaleHeight / 2) - (Distance * Cos(aZ)))), &H8000&
    PicBox.Print Round(aZ * DEGREE, 0)

    rZ = AngleRestrict(rZ)
    Distance = (((vX ^ 2) + (vY ^ 2)) ^ (1 / 2))
    PicBox.Line ((PicBox.ScaleWidth / 2), (PicBox.ScaleHeight / 2))- _
            ((PicBox.ScaleWidth / 2) + (Distance * Sin(rZ)), _
            ((PicBox.ScaleHeight / 2) - (Distance * Cos(rZ)))), vbBlue
    PicBox.Print Round((rZ) * DEGREE, 0)

    dZ = AngleRestrict(dZ)
    Distance = (((tX ^ 2) + (tY ^ 2)) ^ (1 / 2))
    PicBox.Line ((PicBox.ScaleWidth / 2), (PicBox.ScaleHeight / 2))- _
            ((PicBox.ScaleWidth / 2) + (Distance * Sin(dZ)), _
            ((PicBox.ScaleHeight / 2) - (Distance * Cos(dZ)))), vbRed
    PicBox.Print Round((dZ) * DEGREE, 0)
    
    Dim added As String
    added = Round(AngleRestrict(aZ + rZ + dZ) * DEGREE, 0)
    PicBox.CurrentY = PicBox.ScaleHeight - PicBox.TextHeight(added)
    PicBox.CurrentX = PicBox.ScaleWidth - PicBox.TextWidth(added)
    PicBox.Print added
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, , X - (Picture1.ScaleWidth / 2), Y - (Picture1.ScaleHeight / 2)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, , X - (Picture1.ScaleWidth / 2), Y - (Picture1.ScaleHeight / 2)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, Y - (Picture2.ScaleHeight / 2), , X - (Picture2.ScaleWidth / 2)
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, Y - (Picture2.ScaleHeight / 2), , X - (Picture2.ScaleWidth / 2)
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, X - (Picture3.ScaleWidth / 2), Y - (Picture3.ScaleHeight / 2)
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, X - (Picture3.ScaleWidth / 2), Y - (Picture3.ScaleHeight / 2)
End Sub

Private Sub MousePoint(Button As Integer, Optional ByRef X As Variant, Optional ByRef Y As Variant, Optional ByRef Z As Variant)

    If Button = 1 Then
    
        If Not IsMissing(X) Then Vector.X = X
        If Not IsMissing(Y) Then Vector.Y = Y
        If Not IsMissing(Z) Then Vector.Z = Z

        Set Rotate = AnglesOfPoint1(Vector)
        Vertex = Vector
        
        Set Vertex = VectorRotateAxis(Vertex, Rotate)
        Set Angles = AnglesOfPoint1(Vertex)
        
    ElseIf Button = 2 Then
    
        If Not IsMissing(X) Then Vertex.X = X
        If Not IsMissing(Y) Then Vertex.Y = Y
        If Not IsMissing(Z) Then Vertex.Z = Z

        Set Angles = AnglesOfPoint2(Vertex)

    ElseIf Button <> 0 Then
        If Not IsMissing(X) Then points.X = X
        If Not IsMissing(Y) Then points.Y = Y
        If Not IsMissing(Z) Then points.Z = Z
        
        Set twists = AnglesOfPoint3(points)

    End If
    
    Form_Paint
End Sub



Private Sub Timer1_Timer()
    Randomize
    Static stopcount As Long
    stopcount = stopcount + 1
    Select Case stopcount Mod 3
        Case 0
            Picture1_MouseMove 1, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture1_MouseMove 2, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture1_MouseMove 3, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
        Case 1
            Picture2_MouseMove 1, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture2_MouseMove 2, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture2_MouseMove 3, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
        Case 2
            Picture3_MouseMove 1, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture3_MouseMove 2, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
            Picture3_MouseMove 3, 0, RandomPositive(Screen.TwipsPerPixelX * 7, Picture1.Width - (Screen.TwipsPerPixelX * 8)), RandomPositive(Screen.TwipsPerPixelY * 7, Picture1.Height - (Screen.TwipsPerPixelY * 8))
    End Select
        
    If stopcount > 9 Then
        Timer1.Enabled = False
    End If
End Sub
