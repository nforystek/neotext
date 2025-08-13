Attribute VB_Name = "MiscGeometry"
Option Explicit

Public Vertex As New Point
Public Angles As New Point

Public Vector As New Point
Public Rotate As New Point

Public points As New Point
Public twists As New Point

Public Const PI As Single = 3.14159265358979
Public Const DEGREE As Single = 180 / PI
Public Const RADIAN As Single = PI / 180

Public Function RealDegreeAngle(ByVal Angle As Double) As Double
    'input an angle, and ensures it is with-in
    '0.001 to 360 degrees, no neg/zero angles.
    Dim tmp As Double
    If Angle > 360 Then 'above 360
        tmp = Angle - 360
        'invalid numbers can hang it,
        'no change, so tmp<>Angle too
        Do While tmp > 360 And tmp <> Angle
            tmp = tmp - 360
        Loop
        Angle = tmp
    End If
    If Angle <= 0 Then 'zero or below
        tmp = Angle + 360
        'invalid numbers can hang it,
        'no change, so tmp<>Angle too
        Do While tmp <= 0 And tmp <> Angle
            tmp = tmp + 360
        Loop
        Angle = tmp
    End If
    RealDegreeAngle = Angle
End Function

Public Function AngleRestrict(ByVal Angle1 As Single) As Single
    Angle1 = Angle1 * DEGREE
    Do While Round(Angle1, 0) <= 0
        Angle1 = Angle1 + 360
    Loop
    Do While Round(Angle1, 0) > 360
        Angle1 = Angle1 - 360
    Loop
    AngleRestrict = Round(Angle1 * RADIAN, 6)
    If AngleRestrict = PI Or AngleRestrict = PI * 2 Or AngleRestrict = 0 Then
        AngleRestrict = (-AngleRestrict + (PI * 2)) + -(PI * 4)
    End If
End Function

Public Function AngleAxisRestrict(ByRef AxisAngles As Point) As Point
    Set AngleAxisRestrict = New Point
    With AngleAxisRestrict
        .X = AngleRestrict(AxisAngles.X)
        .Y = AngleRestrict(AxisAngles.Y)
        .Z = AngleRestrict(AxisAngles.Z)
    End With
End Function



Public Function InvertAngle(ByVal Angle As Double) As Double
    InvertAngle = RealDegreeAngle(360 - Angle) 'a negative in angles
End Function
Public Function AngleDeduction(ByVal a1 As Double, ByVal a2 As Double) As Double

    a1 = AngleRestrict(a1)
    a2 = AngleRestrict(a2)
    
    a1 = a1 * DEGREE
    a2 = a2 * DEGREE

    AngleDeduction = AngleRestrict((a1 - a2) * RADIAN)

End Function

Public Function AngleAxisDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim d1 As Point
    Dim d2 As Point
    Set d1 = AngleAxisRestrict(p1)
    Set d2 = AngleAxisRestrict(p2)
    
    d1.X = d1.X * DEGREE
    d1.Y = d1.Y * DEGREE
    d1.Z = d1.Z * DEGREE
    
    d2.X = d2.X * DEGREE
    d2.Y = d2.Y * DEGREE
    d2.Z = d2.Z * DEGREE
    
    Set AngleAxisDeduction = New Point
    With AngleAxisDeduction
        .X = (d1.X - d2.X) * RADIAN
        .Y = (d1.Y - d2.Y) * RADIAN
        .Z = (d1.Z - d2.Z) * RADIAN
        
        Set AngleAxisDeduction = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
    Set d1 = Nothing
    Set d2 = Nothing

End Function

Public Function AbsoluteFactor(ByVal N As Single) As Single
    'returns -1 if the n is below zero, returns 1 if n is above zero, and 0 if n is zero
    AbsoluteFactor = ((-(AbsoluteValue(N - 1) - N) - (-AbsoluteValue(N + 1) + N)) * 0.5)
End Function

Public Function AbsoluteValue(ByVal N As Single) As Single
    'same as abs(), returns a number as positive quantified
    AbsoluteValue = (-((-(N * -1) * N) ^ (1 / 2) * -1))
End Function

Public Function AbsoluteWhole(ByVal N As Single) As Single
    'returns only the digits to the left of a decimal in any numerical
    'AbsoluteWhole = (AbsoluteValue(n) - (AbsoluteValue(n) - (AbsoluteValue(n) Mod (AbsoluteValue(n) + 1)))) * AbsoluteFactor(n)
    AbsoluteWhole = (N \ 1) 'is also correct
End Function

Public Function AbsoluteDecimal(ByVal N As Single) As Single
    'returns only the digits to the right of a decimal in any numerical
    AbsoluteDecimal = (AbsoluteValue(N) - AbsoluteValue(AbsoluteWhole(N))) * AbsoluteFactor(N)
End Function

Public Function PointNormalize(ByRef v As Point) As Point
    Set PointNormalize = New Point
    With PointNormalize
        .Z = (v.X ^ 2 + v.Y ^ 2 + v.Z ^ 2) ^ (1 / 2)
        If (.Z = 0) Then .Z = 1
        .X = (v.X / .Z)
        .Y = (v.Y / .Z)
        .Z = (v.Z / .Z)
    End With
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = (((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
    If DistanceEx <> 0 Then DistanceEx = DistanceEx ^ (1 / 2)
End Function

Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal N As Single) As Point
    Dim d As Single
    d = DistanceEx(p1, p2)
    Set DistanceSet = New Point
    With DistanceSet
        If Not (d = N) Then
            If ((d > 0) And (N <> 0)) And (Not (d = N)) Then
                .X = p2.X - p1.X
                .Y = p2.Y - p1.Y
                .Z = p2.Z - p1.Z
                .X = (p1.X + (N * (.X / d)))
                .Y = (p1.Y + (N * (.Y / d)))
                .Z = (p1.Z + (N * (.Z / d)))
            ElseIf (N = 0) Then
                .X = p1.X
                .Y = p1.Y
                .Z = p1.Z
            ElseIf (d = 0) Then
                .X = p2.X
                .Y = p2.Y
                .Z = p2.Z + IIf(p2.Z > p1.Z, N, -N)
            End If
        End If
    End With
End Function

Public Function VectorMultiply(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMultiply = New Point
    With VectorMultiply
        .X = (p1.X * p2.X)
        .Y = (p1.Y * p2.Y)
        .Z = (p1.Z * p2.Z)
    End With
End Function
Public Function VectorNormalize(ByRef p1 As Point) As Point
    Set VectorNormalize = New Point
    With VectorNormalize
        .Z = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.Z))
        If (Round(.Z, 6) > 0) Then
            .Z = (1 / .Z)
            .X = (p1.X * .Z)
            .Y = (p1.Y * .Z)
            .Z = (p1.Z * .Z)
        End If
    End With

'    Set VectorNormalize = New Point
'    With VectorNormalize
'        .Z = VectorMagnitude(p1)
'        If .Z <= epsilon Then .Z = 1
'        .X = (p1.X / .Z)
'        .Y = (p1.Y / .Z)
'        .Z = (p1.Z / .Z)
'        If Abs(.X) < epsilon Then .X = 0
'        If Abs(.Y) < epsilon Then .Y = 0
'        If Abs(.Z) < epsilon Then .Z = 0
'    End With
End Function
Public Function LineNormalize(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set LineNormalize = New Point
    With LineNormalize
        .Z = DistanceEx(p1, p2)
        If (.Z > 0) Then
            .Z = (1 / .Z)
            .X = ((p2.X - p1.X) * .Z)
            .Y = ((p2.Y - p1.Y) * .Z)
            .Z = ((p2.Z - p1.Z) * .Z)
        End If
    End With
End Function
Public Function AngleAxisNormalize(ByRef p1 As Point) As Point
    Set AngleAxisNormalize = New Point
    With AngleAxisNormalize
        .Z = (AngleRestrict(p1.X) + AngleRestrict(p1.Y) + AngleRestrict(p1.Z)) / (360 * RADIAN)
        If .Z <> 0 Then
            .X = (p1.X * .Z)
            .Y = (p1.Y * .Z)
            .Z = (p1.Z * .Z)
        End If
    End With
End Function


Public Function AngleInvertRotation(ByVal A As Single) As Single

    AngleInvertRotation = (-(PI * 2) - A + (PI * 4)) ' - PI

End Function
Public Function AngleAddition(ByVal a1 As Single, ByVal a2 As Single) As Single
    AngleAddition = AngleRestrict(a1 + a2)
End Function
Public Function AngleAxisInvert(ByVal p As Point) As Point
    Set AngleAxisInvert = New Point
    With AngleAxisInvert
        .X = AngleInvertRotation(p.X)
        .Y = AngleInvertRotation(p.Y)
        .Z = AngleInvertRotation(p.Z)
    End With
End Function
Public Function AngleAxisAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim p3 As New Point
    Dim P4 As New Point
    Set p3 = AngleAxisRestrict(p1)
    Set P4 = AngleAxisRestrict(p2)
    
    Set AngleAxisAddition = New Point
    With AngleAxisAddition
    
        .X = (p3.X * DEGREE + P4.X * DEGREE) * RADIAN
        .Y = (p3.Y * DEGREE + P4.Y * DEGREE) * RADIAN
        .Z = (p3.Z * DEGREE + P4.Z * DEGREE) * RADIAN
        
        Set AngleAxisAddition = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
End Function

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal N As Single) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .X = (p1.X * N)
        .Y = (p1.Y * N)
        .Z = (p1.Z * N)
    End With
End Function



Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
End Function

'Private Function Hypotenus(ByVal X As Double, ByVal Y As Double) As Double
'    Hypotenus = ((X ^ 2) + (Y ^ 2)) ^ (1 / 2)
'End Function
'
'Private Function Sine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
'    If Not H = 0 Then
'        If X < Y Then
'            Sine = X / H
'        Else
'            Sine = Y / H
'        End If
'    End If
'End Function
'
'Private Function Cosine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
'    If Not H = 0 Then
'        If X < Y Then
'            Cosine = Y / H
'        Else
'            Cosine = X / H
'        End If
'    End If
'End Function
'
'Private Function Tangent(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
'    If X < Y Then
'        If Not Y = 0 Then
'            Tangent = X / Y
'        End If
'    Else
'        If Not X = 0 Then
'            Tangent = Y / X
'        End If
'    End If
'End Function
