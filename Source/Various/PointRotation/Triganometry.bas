Attribute VB_Name = "Triganometry"
Option Explicit

Public Sub RotateXYQuad(ByVal CW1To3 As Long, ByRef pX As Variant, ByRef pY As Variant)
    'Spins 2D coordinates from its quadrant to the clock-
    'wise quadrant 1 to 3 turns from its current quadrant.
    'Rrotating X and Y clock wise 90, 180 or 270 degrees.
    Dim sw As Variant
    Select Case CW1To3
        Case 1
            pY = -pY
            sw = pY
            pY = pX
            pX = sw
        Case 2
            pX = -pX
            sw = pY
            pY = pX
            pX = sw
        Case 3
            pY = -pY
            pX = -pX
    End Select
End Sub

Public Function AngleX(ByVal Angle As Double, ByVal Distance As Double) As Double
    'given the distance and the angle, return the x coordinate
    AngleX = (Distance * Sin(Angle))
End Function

Public Function AngleY(ByVal Angle As Double, ByVal Distance As Double) As Double
    'given the distance and the angle, return the y coordinate
    AngleY = -(Cos(Angle) * Distance)
End Function

Public Function Hypotenuse(ByVal X As Double, ByVal Y As Double) As Double
    'technically same as the 2D distance if from (0,0), or length X to Y
    Hypotenuse = ((X ^ 2) + (Y ^ 2)) ^ (1 / 2)
End Function

Public Function Sine(ByVal pX As Variant, ByVal pY As Variant) As Variant
    'the same as built-in Sin(Angle) only X and Y are the arguments
    RotateXYQuad 1, pX, pY
    If pX = 0 Then
        If pY <> 0 Then
            Sine = CVErr(449) ' Val("0.#IND")
        End If
    ElseIf pY <> 0 Then
        Sine = CDbl(Abs(pY / (((pX ^ 2) + (pY ^ 2)) ^ (1 / 2))))
    End If
    If pY > 0 Then
        If pX = 0 Then
            Sine = CDbl(1)
        ElseIf Sine < 0 Then
            Sine = CDbl(-Sine)
        End If
    ElseIf pY < 0 Then
        If pX = 0 Then
            Sine = CDbl(-1)
        ElseIf Sine > 0 Then
            Sine = CDbl(-Sine)
        End If
    ElseIf pX <> 0 Then
        Sine = CDbl(0)
    End If
End Function

Public Function Cosine(ByVal pX As Variant, ByVal pY As Variant) As Variant
    'the same as built-in Cos(Angle) only X and Y are the arguments
    RotateXYQuad 1, pX, pY
    If pY = 0 Then
        If pX <> 0 Then
            Cosine = CVErr(449) 'Val("1.#IND")
        End If
    ElseIf pX <> 0 Then
        Cosine = CDbl(Abs(pX / (((pX ^ 2) + (pY ^ 2)) ^ (1 / 2))))
    End If
    If pX > 0 Then
        If pY = 0 Then
            Cosine = CDbl(1)
        ElseIf Cosine < 0 Then
            Cosine = CDbl(-Cosine)
        End If
    ElseIf pX < 0 Then
        If pY = 0 Then
            Cosine = CDbl(-1)
        ElseIf Cosine > 0 Then
            Cosine = CDbl(-Cosine)
        End If
    ElseIf pY <> 0 Then
        Cosine = CDbl(0)
    End If
End Function

Public Function Tangent(ByVal pX As Variant, ByVal pY As Variant) As Variant
    'the same as built-in Tan(Angle) only X and Y are the arguments
    RotateXYQuad 1, pX, pY
    If pX = 0 Then
        If pY > 0 Then
            Tangent = CVErr(449) 'Val("1.#IND")
        ElseIf pY < 0 Then
            Tangent = CDbl(1)
        End If
    ElseIf (pY <> 0) Then
        Tangent = CDbl(Abs(pY / pX))
    End If
    If pX = 0 And pY <> 0 Then
        Tangent = CVErr(0)
    ElseIf pY = 0 And pX <> 0 Then
        Tangent = CDbl(0)
    ElseIf (pX > 0 And pY > 0) Or (pX < 0 And pY < 0) Then
        If Tangent < 0 Then Tangent = CDbl(-Tangent)
    ElseIf (pX < 0 And pY > 0) Or (pX > 0 And pY < 0) Then
        If Tangent > 0 Then Tangent = CDbl(-Tangent)
    End If
End Function

Public Function Secant(ByVal pX As Variant, ByVal pY As Variant) As Variant
    Secant = CDbl(Abs(Cosine(pX, pY)))
    If Secant <> 0 Then Secant = CDbl(1 / Secant)
    If pX = 0 Then
        Secant = CVErr(449)
    ElseIf pY = 0 And pX > 0 Then
        Secant = CDbl(1)
    ElseIf pY = 0 And pX < 0 Then
        Secant = CDbl(-1)
    ElseIf pX > 0 And pY <> 0 Then
        If Secant < 0 Then Secant = CDbl(-Secant)
    ElseIf pX < 0 And pY <> 0 Then
        If Secant > 0 Then Secant = CDbl(-Secant)
    End If
End Function

Public Function Cosecant(ByVal pX As Variant, ByVal pY As Variant) As Variant
    Cosecant = CDbl(Abs(Cosine(pX, pY)))
    If Cosecant <> 0 Then Cosecant = CDbl(1 / Cosecant)
    If pY = 0 Then
        Cosecant = CVErr(449)
    ElseIf pX = 0 And pY > 0 Then
        Cosecant = CDbl(1)
    ElseIf pX = 0 And pY < 0 Then
        Cosecant = CDbl(-1)
    ElseIf pY > 0 And pX <> 0 Then
        If Cosecant < 0 Then Cosecant = CDbl(-Cosecant)
    ElseIf pY < 0 And pX <> 0 Then
        If Cosecant > 0 Then Cosecant = CDbl(-Cosecant)
    End If
End Function

Public Function Cotangent(ByVal pX As Variant, ByVal pY As Variant) As Variant
    Cotangent = CDbl(Abs(Tangent(pX, pY)))
    If Cotangent <> 0 Then Cotangent = CDbl(1 / Cotangent)
    If pY = 0 And pX <> 0 Then
        Cotangent = CVErr(449)
    ElseIf pX = 0 And pY <> 0 Then
        Cotangent = CDbl(0)
    ElseIf (pX > 0 And pY > 0) Or (pX < 0 And pY < 0) Then
        If Cotangent < 0 Then Cotangent = CDbl(-Cotangent)
    ElseIf (pX < 0 And pY > 0) Or (pX > 0 And pY < 0) Then
        If Cotangent > 0 Then Cotangent = CDbl(-Cotangent)
    End If
End Function

Public Function PolarAxis(ByVal X As Double, ByVal Y As Double) As Double
    'returns a value if (x, y) falls on a pole that is vertical, horizontal,
    'or diagonal. the value is to the standard clock time format, 12=noon
    If X = 0 Then
        If Y > 0 Then
            PolarAxis = 12
        ElseIf Y < 0 Then
            PolarAxis = 6
        End If
    ElseIf Y = 0 Then
        If X > 0 Then
            PolarAxis = 3
        ElseIf X < 0 Then
            PolarAxis = 9
        End If
    ElseIf Abs(X) = Abs(Y) Then
        If X > 0 And Y > 0 Then
            PolarAxis = 1.5
        ElseIf X > 0 And Y < 0 Then
            PolarAxis = 4.5
        ElseIf X < 0 And Y < 0 Then
            PolarAxis = 7.5
        ElseIf X < 0 And Y > 0 Then
            PolarAxis = 10.5
        End If
    End If
End Function

Public Function OctentAxium(ByVal X As Double, ByVal Y As Double) As Double
    'returns the octent (every 45 degrees of angle) the point
    'falls within the format is the standard clock, 12=noon
    X = Round(X, 2)
    Y = Round(Y, 2)
    If X <> 0 Or Y <> 0 Then
        OctentAxium = PolarAxis(X, Y)
        If OctentAxium = 0 Then
            If Abs(X) > Abs(Y) Then
                If X > 0 And Y > 0 Then
                    OctentAxium = 2
                ElseIf X > 0 And Y < 0 Then
                    OctentAxium = 4
                ElseIf X < 0 And Y < 0 Then
                    OctentAxium = 8
                ElseIf X < 0 And Y > 0 Then
                    OctentAxium = 10
                End If
            ElseIf Abs(X) < Abs(Y) Then
                If X > 0 And Y > 0 Then
                    OctentAxium = 1
                ElseIf X > 0 And Y < 0 Then
                    OctentAxium = 5
                ElseIf X < 0 And Y < 0 Then
                    OctentAxium = 7
                ElseIf X < 0 And Y > 0 Then
                    OctentAxium = 11
                End If
            End If
        End If
    End If
End Function

Public Function AngleOfPlot(ByVal pX As Double, ByVal pY As Double) As Double
    Dim X As Double
    Dim Y As Double
    X = Round(pX, 12)
    Y = Round(pY, 12)
    If (X = 0) Then
        If (Y > 0) Then
            AngleOfPlot = (180 * RADIAN)
        ElseIf (Y < 0) Then
            AngleOfPlot = (360 * RADIAN)
        End If
    ElseIf (Y = 0) Then
        If (X > 0) Then
            AngleOfPlot = (90 * RADIAN)
        ElseIf (X < 0) Then
            AngleOfPlot = (270 * RADIAN)
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
        If Abs(pX) > Abs(pY) Then
            Large = Abs(pX)
            Least = Abs(pY)
        Else
            Least = Abs(pX)
            Large = Abs(pY)
        End If
        slope = (Least / Large)
        Angle = (((pX ^ 2) + (pY ^ 2)) ^ (1 / 2))
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
        AngleOfPlot = (AngleOfPlot + Angle)
    End If
End Function

