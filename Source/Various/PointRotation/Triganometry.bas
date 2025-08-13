Attribute VB_Name = "Triganometry"
Option Explicit

Public Function Sine(ByRef p As Point) As Variant
    'returns the z axis angle of the x and y in p
    If p.X = 0 Then
        If p.Y <> 0 Then
            Sine = CVErr(449) ' Val("0.#IND")
        End If
    ElseIf p.Y <> 0 Then
        Sine = CDbl(Abs(p.Y / (((p.X ^ 2) + (p.Y ^ 2)) ^ (1 / 2))))
    End If
    If p.Y > 0 Then
        If p.X = 0 Then
            Sine = CDbl(1)
        ElseIf Sine < 0 Then
            Sine = CDbl(-Sine)
        End If
    ElseIf p.Y < 0 Then
        If p.X = 0 Then
            Sine = CDbl(-1)
        ElseIf Sine > 0 Then
            Sine = CDbl(-Sine)
        End If
    ElseIf p.X <> 0 Then
        Sine = CDbl(0)
    End If
End Function

Public Function Cosine(ByRef p As Point) As Variant
    'returns the x axis angle of the x and y in p
    If p.Y = 0 Then
        If p.X <> 0 Then
            Cosine = CVErr(449) 'Val("1.#IND")
        End If
    ElseIf p.X <> 0 Then
        Cosine = CDbl(Abs(p.X / (((p.X ^ 2) + (p.Y ^ 2)) ^ (1 / 2))))
    End If
    If p.X > 0 Then
        If p.Y = 0 Then
            Cosine = CDbl(1)
        ElseIf Cosine < 0 Then
            Cosine = CDbl(-Cosine)
        End If
    ElseIf p.X < 0 Then
        If p.Y = 0 Then
            Cosine = CDbl(-1)
        ElseIf Cosine > 0 Then
            Cosine = CDbl(-Cosine)
        End If
    ElseIf p.Y <> 0 Then
        Cosine = CDbl(0)
    End If
End Function


Public Function Tangent(ByRef p As Point) As Variant
    'returns the y axis angle of the x and y in p
    If p.X = 0 Then
        If p.Y > 0 Then
            Tangent = CVErr(449) 'Val("1.#IND")
        ElseIf p.Y < 0 Then
            Tangent = CDbl(1)
        End If
    ElseIf (p.Y <> 0) Then
        Tangent = CDbl(Abs(p.Y / p.X))
    End If
    If p.X = 0 And p.Y <> 0 Then
        Tangent = CVErr(0)
    ElseIf p.Y = 0 And p.X <> 0 Then
        Tangent = CDbl(0)
    ElseIf (p.X > 0 And p.Y > 0) Or (p.X < 0 And p.Y < 0) Then
        If Tangent < 0 Then Tangent = CDbl(-Tangent)
    ElseIf (p.X < 0 And p.Y > 0) Or (p.X > 0 And p.Y < 0) Then
        If Tangent > 0 Then Tangent = CDbl(-Tangent)
    End If
End Function


Public Function Secant(ByRef p As Point) As Variant
    'returns the rabbit hole axis angle of the carrying traveling triganometry
    Secant = CDbl(Abs(Cosine(p)))
    If Secant <> 0 Then Secant = CDbl(1 / Secant)
    If p.X = 0 Then
        Secant = CVErr(449)
    ElseIf p.Y = 0 And p.X > 0 Then
        Secant = CDbl(1)
    ElseIf p.Y = 0 And p.X < 0 Then
        Secant = CDbl(-1)
    ElseIf p.X > 0 And p.Y <> 0 Then
        If Secant < 0 Then Secant = CDbl(-Secant)
    ElseIf p.X < 0 And p.Y <> 0 Then
        If Secant > 0 Then Secant = CDbl(-Secant)
    End If
End Function
Public Function Cosecant(ByRef p As Point) As Variant
    'returns the rabbit hole axis angle of the carrying traveling triganometry
    Cosecant = CDbl(Abs(Cosine(p)))
    If Cosecant <> 0 Then Cosecant = CDbl(1 / Cosecant)
    If p.Y = 0 Then
        Cosecant = CVErr(449)
    ElseIf p.X = 0 And p.Y > 0 Then
        Cosecant = CDbl(1)
    ElseIf p.X = 0 And p.Y < 0 Then
        Cosecant = CDbl(-1)
    ElseIf p.Y > 0 And p.X <> 0 Then
        If Cosecant < 0 Then Cosecant = CDbl(-Cosecant)
    ElseIf p.Y < 0 And p.X <> 0 Then
        If Cosecant > 0 Then Cosecant = CDbl(-Cosecant)
    End If
End Function

Public Function Cotangent(ByRef p As Point) As Variant
    'returns the rabbit hole axis angle of the carrying traveling triganometry
    Cotangent = CDbl(Abs(Tangent(p)))
    If Cotangent <> 0 Then Cotangent = CDbl(1 / Cotangent)
    If p.Y = 0 And p.X <> 0 Then
        Cotangent = CVErr(449)
    ElseIf p.X = 0 And p.Y <> 0 Then
        Cotangent = CDbl(0)
    ElseIf (p.X > 0 And p.Y > 0) Or (p.X < 0 And p.Y < 0) Then
        If Cotangent < 0 Then Cotangent = CDbl(-Cotangent)
    ElseIf (p.X < 0 And p.Y > 0) Or (p.X > 0 And p.Y < 0) Then
        If Cotangent > 0 Then Cotangent = CDbl(-Cotangent)
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

