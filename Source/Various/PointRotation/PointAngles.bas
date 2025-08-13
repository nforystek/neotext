Attribute VB_Name = "PointAngles"
Option Explicit



Public Function AnglesOfPoint1(ByRef p As Point) As Point

    Set AnglesOfPoint1 = MakePoint(AngleOfPlot(p.Y, p.Z), AngleOfPlot(p.Z, p.X), AngleOfPlot(p.X, p.Y))

End Function

Public Function AnglesOfPoint2(ByRef p As Point) As Point
    Set AnglesOfPoint2 = New Point
    With AnglesOfPoint2
        If Not (p.X = 0 And p.Y = 0 And p.Z = 0) Then
            Dim tmp As New Point
            Set tmp = p
            .X = AngleRestrict(AngleOfPlot(tmp.Y, tmp.Z))
            .Y = AngleRestrict(AngleOfPlot(tmp.Z, tmp.X))
            .Z = AngleRestrict(AngleOfPlot(tmp.X, tmp.Y))
            Set tmp = Nothing
        End If
    End With
End Function

Public Function AnglesOfPoint3(ByRef Point As Point) As Point
    Static stack As Integer
    stack = stack + 1
    If stack = 1 Then
        '(1,1,1) is high noon
        'to 45 degree sections
        Point.X = Point.X + 1
        Point.Y = Point.Y + 1
        Point.Z = Point.Z + 1
    End If
    Set AnglesOfPoint3 = New Point
    With AnglesOfPoint3
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
                    AnglesOfPoint3.X = Point.Y
                    AnglesOfPoint3.Y = Point.Z
                    AnglesOfPoint3.Z = Point.X
                    .Z = AnglesOfPoint3(AnglesOfPoint3).Z
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
                    AnglesOfPoint3.X = Point.Y
                    AnglesOfPoint3.Y = Point.Z
                    AnglesOfPoint3.Z = Point.X
                    .Z = AnglesOfPoint3(AnglesOfPoint3).Z
                End If
            ElseIf (X <> 0) And (Y <> 0) Then
                Dim slope As Single
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
                slope = (Least / Large) 'the angle in square form
                '^^ or tangent, tangable to other axis angles' shared axis
                Dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2)) 'distance
                'still traveling for tangents and cosines
                Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
                Least = (((Dist ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
                Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / Dist)) * (Least / Dist))
                '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
                'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
                Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)  'bulk sine of the angle in 45 degree slices
                '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
                If (Z <> 0) Then 'two or less axis is one rotation
                    Dim ret As Point
                    AnglesOfPoint3.X = Point.Y
                    AnglesOfPoint3.Y = Point.Z
                    AnglesOfPoint3.Z = Point.X
                    Set ret = AnglesOfPoint3(AnglesOfPoint3)
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



