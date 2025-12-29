Attribute VB_Name = "Mathematics"
Option Explicit

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Double
    RandomPositive = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function Sign(ByVal N As Double) As Double
    Sign = ((-(Abs((N * 99.99) - 1) - (N * 99.99)) - (-Abs((N * 99.99) + 1) + (N * 99.99))) * 0.5)
End Function

Public Function Signn(ByVal Value As Double) As Double
    Signn = ((-((AbsoluteWhole(Value) / 1 <> 0) * 1) + -1) + -(((-AbsoluteWhole(Value) / 1 + -1) = 0) * 1))
End Function

Public Function InvertNumber(ByVal Value As Long, Optional ByVal Whole As Long = 100, Optional ByVal Unit As Long = 1)
    'returns the inverted value of a whole conprised of unit measures, AbsoluteInvert(25) returns 75
    'another example: AbsoluteInvert(0, 16777216) returns the negative of black 0, which is 16777216
    InvertNumber = -(Whole / Unit) + -(Value / Unit) + ((Whole / Unit) * 2)
End Function


Public Function AbsoluteValue(ByVal N As Double) As Double
    'returns the number in N as positive quantified (same as abs())
    AbsoluteValue = (-((-(N * -1) * N) ^ (1 / 2) * -1))
End Function

Public Function AbsoluteFactor(ByVal N As Double) As Double
    'returns -1 if the N is below zero, returns 1 if N is above zero, and 0 if N is zero
    AbsoluteFactor = ((-(AbsoluteValue(N - 1) - N) - (-AbsoluteValue(N + 1) + N)) * 0.5)
End Function

Public Function AbsoluteWhole(ByVal N As Double) As Double
    'returns only the digits to the left of the decimal point in N
    AbsoluteWhole = (AbsoluteValue(N) - (AbsoluteValue(N) - (AbsoluteValue(N) Mod (AbsoluteValue(N) + 1)))) * AbsoluteFactor(N)
    'AbsoluteWhole = (N \ 1) 'this line returns same whole value as well (but doesn't exist in math, only integral programming)
End Function

Public Function AbsoluteDecimal(ByVal N As Double) As Double
    'returns only the digits to the right of the decimal point in N
    AbsoluteDecimal = (AbsoluteValue(N) - AbsoluteValue(AbsoluteWhole(N))) * AbsoluteFactor(N)
End Function

Public Function Large(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (V1 >= V2) Then
            Large = V1
        Else
            Large = V2
        End If
    ElseIf IsMissing(V4) Then
        If ((V2 >= V3) And (V2 >= V1)) Then
            Large = V2
        ElseIf ((V1 >= V3) And (V1 >= V2)) Then
            Large = V1
        Else
            Large = V3
        End If
    Else
        If ((V2 >= V3) And (V2 >= V1) And (V2 >= V4)) Then
            Large = V2
        ElseIf ((V1 >= V3) And (V1 >= V2) And (V1 >= V4)) Then
            Large = V1
        ElseIf ((V3 >= V1) And (V3 >= V2) And (V3 >= V4)) Then
            Large = V3
        Else
            Large = V4
        End If
    End If
End Function

Public Function Least(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (V1 <= V2) Then
            Least = V1
        Else
            Least = V2
        End If
    ElseIf IsMissing(V4) Then
        If ((V2 <= V3) And (V2 <= V1)) Then
            Least = V2
        ElseIf ((V1 <= V3) And (V1 <= V2)) Then
            Least = V1
        Else
            Least = V3
        End If
    Else
        If ((V2 <= V3) And (V2 <= V1) And (V2 <= V4)) Then
            Least = V2
        ElseIf ((V1 <= V3) And (V1 <= V2) And (V1 <= V4)) Then
            Least = V1
        ElseIf ((V3 <= V1) And (V3 <= V2) And (V3 <= V4)) Then
            Least = V3
        Else
            Least = V4
        End If
    End If
End Function

