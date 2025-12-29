Attribute VB_Name = "Geometry"
Option Explicit


Public Type Vector
    X As Double
    Y As Double
    Z As Double
End Type

Public Const PI As Double = 3.14159265358979
Public Const DEGREE As Double = (180 / PI)
Public Const RADIAN As Double = (PI / 180)


Private Function Hypotenus(ByVal X As Double, ByVal Y As Double) As Double
    Hypotenus = ((X ^ 2) + (Y ^ 2)) ^ (1 / 2)
End Function

Private Function Sine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If Not H = 0 Then
        If X < Y Then
            Sine = X / H
        Else
            Sine = Y / H
        End If
    End If
End Function
Private Function Cosine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If Not H = 0 Then
        If X < Y Then
            Cosine = Y / H
        Else
            Cosine = X / H
        End If
    End If
End Function
Private Function Tangent(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If X < Y Then
        If Not Y = 0 Then
            Tangent = X / Y
        End If
    Else
        If Not X = 0 Then
            Tangent = Y / X
        End If
    End If
End Function

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

Public Function InvertAngle(ByVal Angle As Double) As Double
    InvertAngle = RealDegreeAngle(360 - Angle) 'a negative in angles
End Function

Public Function AngleOfPlot(ByVal Px As Double, ByVal py As Double) As Double
    Dim X As Double
    Dim Y As Double
    X = Round(Px, 12)
    Y = Round(py, 12)
    If (X = 0) Then
        If (Y > 0) Then
            AngleOfPlot = 180
        ElseIf (Y < 0) Then
            AngleOfPlot = 360
        End If
    ElseIf (Y = 0) Then
        If (X > 0) Then
            AngleOfPlot = 90
        ElseIf (X < 0) Then
            AngleOfPlot = 270
        End If
    Else
        If ((X > 0) And (Y > 0)) Then
            AngleOfPlot = (90 * RADIAN)
        ElseIf ((X < 0) And (Y > 0)) Then
            AngleOfPlot = (180 * RADIAN)
        ElseIf ((X < 0) And (Y < 0)) Then
            AngleOfPlot = (270 * RADIAN)
        ElseIf ((X > 0) And (Y < 0)) Then
            AngleOfPlot = (360 * RADIAN)
        End If
        Dim slope As Double
        Dim Large As Double
        Dim Least As Double
        Dim Angle As Double
        If Abs(Px) > Abs(py) Then
            Large = Abs(Px)
            Least = Abs(py)
        Else
            Least = Abs(Px)
            Large = Abs(py)
        End If
        slope = (Least / Large)
        Angle = (((Px ^ 2) + (py ^ 2)) ^ (1 / 2))
        Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2))
        Least = (((Angle ^ 2) - (Least ^ 2)) ^ (1 / 2))
        Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / Angle)) * (Least / Angle))
        Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)
        Angle = Round(Large + Least, 12)
        If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
           (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
            Angle = (PI / 4) - Angle
            AngleOfPlot = AngleOfPlot + (PI / 4)
        End If
        AngleOfPlot = ((AngleOfPlot + Angle) * DEGREE)
    End If
End Function

'Public Sub AngleOfXAxis(ByVal Px As Double, ByVal py As Double, ByVal pz As Double, ByRef ax As Double)
'
'    If Not (Px = 0 And py = 0 And pz = 0) Then
'
'        ax = RealDegreeAngle(AngleOfPlot(py, pz))
'
'    End If
'End Sub
'
'Public Sub AngleOfYAxis(ByVal Px As Double, ByVal py As Double, ByVal pz As Double, ByRef ay As Double)
'
'    If Not (Px = 0 And py = 0 And pz = 0) Then
'
'        ax = RealDegreeAngle(AngleOfPlot(py, pz))
'
'        RotateXAxis Px, py, pz, InvertAngle(ax)
'
'        ay = RealDegreeAngle(AngleOfPlot(pz, Px))
'
'    End If
'End Sub
'
'Public Sub AngleOfZAxis(ByVal Px As Double, ByVal py As Double, ByVal pz As Double, ByRef az As Double)
'    If Not (Px = 0 And py = 0 And pz = 0) Then
'
'        ax = RealDegreeAngle(AngleOfPlot(py, pz))
'
'        RotateXAxis Px, py, pz, InvertAngle(ax)
'
'        ay = RealDegreeAngle(AngleOfPlot(pz, Px))
'
'        RotateYAxis Px, py, pz, InvertAngle(ay)
'
'        az = RealDegreeAngle(AngleOfPlot(Px, py))
'    End If
'End Sub

Public Sub AnglesOfPoint(ByVal Px As Double, ByVal py As Double, ByVal pz As Double, ByRef ax As Double, ByRef ay As Double, ByRef az As Double)
    Dim tx As Double
    Dim ty As Double
    Dim tz As Double

    If Not (Px = 0 And py = 0 And pz = 0) Then
        tx = Px
        ty = py
        tz = pz

        ax = RealDegreeAngle(AngleOfPlot(ty, tz))

        RotateXAxis tx, ty, tz, InvertAngle(ax)

        ay = RealDegreeAngle(AngleOfPlot(tz, tx))

        RotateYAxis tx, ty, tz, InvertAngle(ay)

        az = RealDegreeAngle(AngleOfPlot(tx, ty))

    End If


End Sub

Public Sub RotateXAxis(ByVal Px As Double, ByRef py As Double, ByRef pz As Double, ByVal Angle As Double)
    Dim ty As Double
    Dim tz As Double
    
    Dim CosPhi As Double
    Dim SinPhi As Double
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)

    tz = pz * CosPhi - py * SinPhi
    ty = pz * SinPhi + py * CosPhi

    py = ty
    pz = tz
End Sub

Public Sub RotateYAxis(ByRef Px As Double, ByVal py As Double, ByRef pz As Double, ByVal Angle As Double)
    Dim tx As Double
    Dim tz As Double
    
    Dim CosPhi As Double
    Dim SinPhi As Double
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)

    tx = Px * CosPhi - pz * SinPhi
    tz = Px * SinPhi + pz * CosPhi

    Px = tx
    pz = tz
End Sub

Public Sub RotateZAxis(ByRef Px As Double, ByRef py As Double, ByVal pz As Double, ByVal Angle As Double)
    Dim tx As Double
    Dim ty As Double
    
    Dim CosPhi As Double
    Dim SinPhi As Double
    CosPhi = Cos(Angle)
    SinPhi = Sin(Angle)

    tx = Px * CosPhi - py * SinPhi
    ty = Px * SinPhi + py * CosPhi
    
    Px = tx
    py = ty
End Sub


Public Sub RotateAllAxis(ByRef Px As Double, ByRef py As Double, ByRef pz As Double, ByVal ax As Double, ByVal ay As Double, ByVal az As Double)
'    Dim tx As Double
'    Dim ty As Double
'    Dim tz As Double
'
'    Dim rx As Double
'    Dim ry As Double
'    Dim rz As Double
'
'    ty = Cos(ax) * py - Sin(ax) * pz
'    tz = Sin(ax) * py + Cos(ax) * pz
'    tx = Px
'
'    Px = Sin(ay) * tz + Cos(ay) * tx
'    pz = Cos(ay) * tz - Sin(ay) * tx
'
'    tx = Cos(az) * Px - Sin(az) * ty
'    ty = Sin(az) * Px + Cos(az) * ty
'
'    Px = tx
'    py = ty

    RotateZAxis Px, py, pz, az
    RotateXAxis Px, py, pz, ax
    RotateYAxis Px, py, pz, ay
        
End Sub


Public Function AngleRestrict(ByVal Angle1 As Double) As Double
    Angle1 = Angle1
    Angle1 = RealDegreeAngle(Angle1)
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

Public Function Length(ByRef p1 As Point) As Double
    Length = ((p1.X ^ 2) + (p1.Y ^ 2) + (p1.Z ^ 2))
    If Length <> 0 Then Length = Length ^ (1 / 2)
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Double
    DistanceEx = (((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
    If DistanceEx <> 0 Then DistanceEx = DistanceEx ^ (1 / 2)
End Function

Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal N As Double) As Point
    Dim d As Double
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

Public Function VertexNormalize(ByRef p1 As Point) As Point
    Set VertexNormalize = New Point
    With VertexNormalize
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

Public Function AngleInvertRotation(ByVal A As Double) As Double
    AngleInvertRotation = (-(PI * 2) - A + (PI * 4)) ' - PI
End Function

Public Function AngleAddition(ByVal a1 As Double, ByVal a2 As Double) As Double
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

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal N As Double) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .X = (p1.X * N)
        .Y = (p1.Y * N)
        .Z = (p1.Z * N)
    End With
End Function

Public Function MakePlot(ByVal X As Double, ByVal Y As Double) As Plot
    Set MakePlot = New Plot
    MakePlot.X = X
    MakePlot.Y = Y
End Function

Public Function MakePoint(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
End Function

