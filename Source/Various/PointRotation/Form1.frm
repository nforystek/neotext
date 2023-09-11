VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   ScaleHeight     =   15795
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   3540
      Left            =   8250
      ScaleHeight     =   3480
      ScaleWidth      =   4125
      TabIndex        =   2
      Top             =   285
      Width           =   4185
   End
   Begin VB.PictureBox Picture2 
      Height          =   3525
      Left            =   4110
      ScaleHeight     =   3465
      ScaleWidth      =   3810
      TabIndex        =   1
      Top             =   210
      Width           =   3870
   End
   Begin VB.PictureBox Picture1 
      Height          =   3360
      Left            =   345
      ScaleHeight     =   3300
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   300
      Width           =   3480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PI As Single = 3.14159265358979
Private Const DEGREE As Single = 180 / PI
Private Const RADIAN As Single = PI / 180

Private Vertex As New Point
Private Angles As New Point
Private Vector As New Point
Private Rotate As New Point

Private Sub Form_Load()
    
    Me.MousePointer = 2
    
    Form_Paint
End Sub

Private Sub Form_Paint()
    DrawAxis "X Axis", Picture1, Vertex.Y, Vertex.Z, Angles.X, Vector.Y, Vector.Z, Rotate.X
    DrawAxis "Y Axis", Picture2, Vertex.Z, Vertex.X, Angles.Y, Vector.Z, Vector.X, Rotate.Y
    DrawAxis "Z Axis", Picture3, Vertex.X, Vertex.Y, Angles.Z, Vector.X, Vector.Y, Rotate.Z
End Sub

Private Sub DrawAxis(ByVal Axis As String, ByRef PicBox As PictureBox, ByVal pX As Single, ByVal pY As Single, ByVal aZ As Single, ByVal vX As Single, ByVal vY As Single, ByVal rZ As Single)
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
    
    Dim added As String
    added = Round(AngleRestrict(aZ + rZ) * DEGREE, 0)
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

    If Button = 2 Then
    
        If Not IsMissing(X) Then Vector.X = X
        If Not IsMissing(Y) Then Vector.Y = Y
        If Not IsMissing(Z) Then Vector.Z = Z

        Set Rotate = AnglesOfPoint(Vector)
        
    ElseIf Button = 1 Then
    
        If Not IsMissing(X) Then Vertex.X = X
        If Not IsMissing(Y) Then Vertex.Y = Y
        If Not IsMissing(Z) Then Vertex.Z = Z

        Set Angles = AnglesOfPoint(Vertex)
                
    End If
    
    Form_Paint
End Sub


Public Function AngleRestrict(ByVal Angle1 As Single) As Single
    Angle1 = Round(Angle1 * DEGREE, 0)
    Do While Angle1 > 360
        Angle1 = Angle1 - 360
    Loop
    Do While Angle1 <= 0
        Angle1 = Angle1 + 360
    Loop
    AngleRestrict = Angle1 * RADIAN
End Function

Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
End Function

Public Function AnglesOfPoint(ByRef Point As Point) As Point
    Static stack As Integer
    stack = stack + 1
    If stack = 1 Then
        '(1,1,1) is high noon
        'to 45 degree sections
        Point.X = Point.X + 1
        Point.Y = Point.Y + 1
        Point.Z = Point.Z + 1
    End If
    Set AnglesOfPoint = New Point
    With AnglesOfPoint
        If stack < 5 Then
            Dim X As Single
            Dim Y As Single
            Dim Z As Single
            'round them off for checking
            '(6 is for single precision)
            X = Round(Point.X, 6)
            Y = Round(Point.Y, 6)
            Z = Round(Point.Z, 6)
            If (X = 0) Then  'slope of 1
                If (Z = 0) Then
                    'must be 360 or 180
                    If (Y > 0) Then
                        .Z = (180 * RADIAN)
                    ElseIf (Y < 0) Then
                        .Z = (360 * RADIAN)
                    End If
                Else
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.Z
                    AnglesOfPoint.Z = Point.X
                    .Z = AnglesOfPoint(AnglesOfPoint).Z
                End If
            ElseIf (Y = 0) Then   'slope of 0
                If (Z = 0) Then
                    'must be 90 or 270
                    If (X > 0) Then
                        .Z = (90 * RADIAN)
                    ElseIf (X < 0) Then
                        .Z = (270 * RADIAN)
                    End If
                Else
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.Z
                    AnglesOfPoint.Z = Point.X
                    .Z = AnglesOfPoint(AnglesOfPoint).Z
                End If
            ElseIf (X <> 0) And (Y <> 0) Then
                Dim Slope As Single
                Dim Dist As Single
                Dim Large As Single
                Dim Least As Single
                Dim Angle As Single
                'find the larger coordinate
                If Abs(Point.X) > Abs(Point.Y) Then
                    Large = Abs(Point.X)
                    Least = Abs(Point.Y)
                Else
                    Least = Abs(Point.X)
                    Large = Abs(Point.Y)
                End If
                Slope = (Least / Large) 'the angle in square form
                '^^ or tangent, tangable to other axis angles' shared axis
                Dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2)) 'distance
                'still traveling for tangents and cosines
                Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
                Least = (((Dist ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
                Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * Slope) * (Large / Dist)) * (Least / Dist))
                '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
                'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
                Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * Slope)  'bulk sine of the angle in 45 degree slices
                '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
                If (Z <> 0) Then 'two or less axis is one rotation
                    Dim ret As Point
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.Z
                    AnglesOfPoint.Z = Point.X
                    Set ret = AnglesOfPoint(AnglesOfPoint)
                    If stack = 2 Then
                        .X = -ret.Z
                    End If
                    If stack = 1 Then
                        .X = -ret.X
                        .Y = ret.Z
                    End If
                    Set ret = Nothing
                End If
                'get the base angle
                '(up to the quardrant)
                If ((X > 0) And (Y > 0)) Then
                    .Z = (90 * RADIAN)
                ElseIf ((X < 0) And (Y > 0)) Then
                    .Z = (180 * RADIAN)
                ElseIf ((X < 0) And (Y < 0)) Then
                    .Z = (270 * RADIAN)
                ElseIf ((X > 0) And (Y < 0)) Then
                    .Z = (360 * RADIAN)
                End If
                'develop the final angle Z for this duel coordinate X,Y axis only
                Angle = (Large + Least)
                If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
                   (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
                   'the angle for 45 to 90 is in reverse, and doesn't start at 45, but because we
                   'are calculating a second 45 of 90, the offset (-1 not 0) is included if inverse
                    Angle = (PI / 4) - Angle
                    'then also add 45 to the base
                    .Z = .Z + (PI / 4)
                End If
                'add it to the base, returing as .Z
                .Z = .Z + Angle
                If stack = 1 Then
                    'reorganization
                    Angle = .Y
                    .Y = .Z
                    .Z = Angle
                    Angle = .X
                    .X = .Y
                    .Y = .Z
                    .Z = Angle
                    Angle = .X
                    .X = .Y
                    .Y = .Z
                    .Z = Angle
                End If
            End If
        End If
    End With
    If stack = 1 Then 'undo
        Point.X = Point.X - 1
        Point.Y = Point.Y - 1
        Point.Z = Point.Z - 1
    End If
    stack = stack - 1
End Function

