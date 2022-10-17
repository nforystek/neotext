Attribute VB_Name = "modGeometry"
Option Explicit

Option Compare Binary

Public Const PI As Single = 3.14159265358979
Public Const Epsilon As Double = 0.999999999999999
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2
Public Const RADIAN As Single = PI / 180
Public Const FOOT As Single = 0.1
Public Const FOVY As Single = (FOOT * 8) '4 feet left, and 4 feet right = 0.8
Public Const FAR  As Single = 90000
Public Const NEAR As Single = 0.05 'one millimeter (308.4 per foor) or greater


Public Function v(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set v = New Point
    With v
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = ((((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = ((((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2)) ^ (1 / 2))
End Function

Public Function RandomPositive(ByVal Lowerbound As Long, ByVal Upperbound As Long) As Single
    RandomPositive = CSng((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function

Public Function PlaneNormal(ByRef v0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    Set PlaneNormal = New Point
    PlaneNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(v0, v1), VectorDeduction(v1, v2)))
End Function

Public Function PointNormal(ByRef v As Point) As Single
    PointNormal = Sqr(VectorDotProduct(v, v))
End Function

Public Function NormalizedAxii(ByVal Value As Single) As Single
    NormalizedAxii = ((-CSng(CBool(Round(Value, 0))) + -1) + -CSng(Not CBool(-Round(Value, 0) + -1)))
End Function

Public Function SphereSurfaceArea(ByVal Radii As Single) As Single
     SphereSurfaceArea = (4 * PI * (Radii ^ 2))
End Function

Public Function SphereVolume(ByVal Radii As Single) As Single
    SphereVolume = ((4 / 3) * PI * (Radii ^ 3))
End Function

Public Function SquareCenter(ByRef v0 As Point, ByRef v1 As Point, ByRef v2 As Point, ByRef V3 As Point) As Point
    Set SquareCenter = New Point
    With SquareCenter
        .X = (LeastOf(v0.X, v1.X, v2.X, V3.X) + ((LargeOf(v0.X, v1.X, v2.X, V3.X) - LeastOf(v0.X, v1.X, v2.X, V3.X)) / 2))
        .Y = (LeastOf(v0.Y, v1.Y, v2.Y, V3.Y) + ((LargeOf(v0.Y, v1.Y, v2.Y, V3.Y) - LeastOf(v0.Y, v1.Y, v2.Y, V3.Y)) / 2))
        .Z = (LeastOf(v0.Z, v1.Z, v2.Z, V3.Z) + ((LargeOf(v0.Z, v1.Z, v2.Z, V3.Z) - LeastOf(v0.Z, v1.Z, v2.Z, V3.Z)) / 2))
    End With
End Function

Public Function CirclePermeter(ByVal Radii As Single) As Single
    CirclePermeter = ((Radii * 2) * PI)
End Function
Public Function CubePerimeter(ByVal Edge As Single) As Single
    CubePerimeter = (Edge * 12)
End Function
Public Function CubeSurfaceArea(ByVal Edge As Single) As Single
    CubeSurfaceArea = (6 * (Edge ^ 2))
End Function

Public Function CubeVolume(ByVal Edge As Single) As Single
    CubeVolume = (Edge ^ 3)
End Function

Public Function TrianglePerimeter(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    TrianglePerimeter = (DistanceEx(p1, p2) + DistanceEx(p2, p3) + DistanceEx(p3, p1))
End Function

Public Function TriangleSurfaceArea(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    Dim l1 As Single
    Dim l2 As Single
    Dim l3 As Single
    l1 = ((((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2)) ^ (1 / 2))
    l2 = ((((p2.X - p3.X) ^ 2) + ((p2.Y - p3.Y) ^ 2) + ((p2.Z - p3.Z) ^ 2)) ^ (1 / 2))
    l3 = ((((p3.X - p1.X) ^ 2) + ((p3.Y - p1.Y) ^ 2) + ((p3.Z - p1.Z) ^ 2)) ^ (1 / 2))
    TriangleSurfaceArea = (((((((l1 + l2) - l3) + ((l2 + l3) - l1) + ((l3 + l1) - l2)) * (l1 * l2 * l3)) / (l1 + l2 + l3)) ^ (1 / 2)))
End Function

Public Function TriangleVolume(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    TriangleVolume = TriangleSurfaceArea(p1, p2, p3) 'volume times 12 will equal same in cubic volume for how many 3d prism
    TriangleVolume = ((((TriangleVolume ^ (1 / 3)) ^ 2) ^ 3) / 12) 'a cube is comprised if no prisim shows more then one face
End Function

Public Function TriangleDotProduct(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    TriangleDotProduct = (((VectorDotProduct(p1, VectorSubtraction(p2, p3)) * VectorDotProduct(p2, VectorSubtraction(p1, p3))) ^ (1 / 3)) * 2)
End Function

Public Function TriangleAveraged(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleAveraged = New Point
    With TriangleAveraged
        .X = ((p1.X + p2.X + p3.X) / 3)
        .Y = ((p1.Y + p2.Y + p3.Y) / 3)
        .Z = ((p1.Z + p2.Z + p3.Z) / 3)
    End With
End Function

Public Function TriangleOffset(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleOffset = New Point
    With TriangleOffset
        .X = (LargeOf(p1.X, p2.X, p3.X) - LeastOf(p1.X, p2.X, p3.X))
        .Y = (LargeOf(p1.Y, p2.Y, p3.Y) - LeastOf(p1.Y, p2.Y, p3.Y))
        .Z = (LargeOf(p1.Z, p2.Z, p3.Z) - LeastOf(p1.Z, p2.Z, p3.Z))
    End With
End Function

Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleAxii = New Point
    With TriangleAxii
        Dim o As Point
        Set o = TriangleOffset(p1, p2, p3)
        .X = (LeastOf(p1.X, p2.X, p3.X) + (o.X / 2))
        .Y = (LeastOf(p1.Y, p2.Y, p3.Y) + (o.Y / 2))
        .Z = (LeastOf(p1.Z, p2.Z, p3.Z) + (o.Z / 2))
    End With
End Function

Public Function TriangleNormal(ByRef v0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    Set TriangleNormal = New Point
    Dim o As Point
    Dim d As Single
    With TriangleNormal
        Set o = TriangleOffset(v0, v1, v2)
        d = (o.X + o.Y + o.Z)
        .Z = (((o.X + o.Y) - o.Z) / d)
        .X = (((o.Y + o.Z) - o.X) / d)
        .Y = (((o.Z + o.X) - o.Y) / d)
    End With
End Function

Public Function TriangleAccordance(ByRef v0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    Set TriangleAccordance = New Point
    With TriangleAccordance
        .X = (((v0.X + v1.X) - v2.X) + ((v1.X + v2.X) - v0.X) - ((v2.X + v0.X) - v1.X))
        .Y = (((v0.Y + v1.Y) - v2.Y) + ((v1.Y + v2.Y) - v0.Y) - ((v2.Y + v0.Y) - v1.Y))
        .Z = (((v0.Z + v1.Z) - v2.Z) + ((v1.Z + v2.Z) - v0.Z) - ((v2.Z + v0.Z) - v1.Z))
    End With
End Function

Public Function TriangleDisplace(ByRef v0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    Set TriangleDisplace = New Point
    With TriangleDisplace
        .X = (Abs((Abs(v0.X) + Abs(v1.X)) - Abs(v2.X)) + Abs((Abs(v1.X) + Abs(v2.X)) - Abs(v0.X)) - Abs((Abs(v2.X) + Abs(v0.X)) - Abs(v1.X)))
        .Y = (Abs((Abs(v0.Y) + Abs(v1.Y)) - Abs(v2.Y)) + Abs((Abs(v1.Y) + Abs(v2.Y)) - Abs(v0.Y)) - Abs((Abs(v2.Y) + Abs(v0.Y)) - Abs(v1.Y)))
        .Z = (Abs((Abs(v0.Z) + Abs(v1.Z)) - Abs(v2.Z)) + Abs((Abs(v1.Z) + Abs(v2.Z)) - Abs(v0.Z)) - Abs((Abs(v2.Z) + Abs(v0.Z)) - Abs(v1.Z)))
    End With
End Function

Public Function VectorRise(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRise = (LargeOf(p1.Y, p2.Y) - LeastOf(p1.Y, p2.Y))
End Function

Public Function VectorRun(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRun = DistanceEx(v(p1.X, 0, p1.Z), v(p2.X, 0, p2.Z))
End Function

Public Function VectorSlope(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorSlope = VectorRun(p1, p2)
    If (VectorSlope <> 0) Then
        VectorSlope = (VectorRise(p1, p2) / VectorSlope)
        If (VectorSlope = 0) Then VectorSlope = -CInt(Not ((p1.X = p2.X) And (p1.Y = p2.Y) And (p1.Z = p2.Z)))
    End If
End Function

Public Function VectorYIntercept(ByRef p1 As Point, ByRef p2 As Point) As Single
    With VectorMidPoint(p1, p2)
        VectorYIntercept = VectorSlope(p1, p2)
        VectorYIntercept = -((VectorYIntercept * .X) - .Y)
    End With
End Function

Public Function VectorDotProduct(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorDotProduct = ((p1.X * p2.X) + (p1.Y * p2.Y) + (p1.Z * p2.Z))
End Function

Public Function VectorCrossProduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCrossProduct = New Point
    With VectorCrossProduct
        .X = ((p1.Y * p2.Z) - (p1.Z * p2.Y))
        .Y = ((p1.Z * p2.X) - (p1.X * p2.Z))
        .Z = ((p1.X * p2.Y) - (p1.Y * p2.X))
    End With
End Function

Public Function VectorSubtraction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorSubtraction = New Point
    With VectorSubtraction
        .X = ((p1.X - p2.Z) - (p1.X - p2.Y))
        .Y = ((p1.Y - p2.X) - (p1.Y - p2.Z))
        .Z = ((p1.Z - p2.Y) - (p1.Z - p2.X))
    End With
End Function

Public Function VectorAccordance(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAccordance = New Point
    With VectorAccordance
        .X = (((p1.X + p1.Y) - p2.Z) + ((p1.Z + p1.X) - p2.Y) - ((p1.Y + p1.Z) - p2.X))
        .Y = (((p1.Y + p1.Z) - p2.X) + ((p1.X + p1.Y) - p2.Z) - ((p1.Z + p1.X) - p2.Y))
        .Z = (((p1.Z + p1.X) - p2.Y) + ((p1.Y + p1.Z) - p2.X) - ((p1.X + p1.Y) - p2.Z))
    End With
End Function

Public Function VectorDisplace(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDisplace = New Point
    With VectorDisplace
        .X = (Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.Z)) + Abs((Abs(p1.Z) + Abs(p1.X)) - Abs(p2.Y)) - Abs((Abs(p1.Y) + Abs(p1.Z)) - Abs(p2.X)))
        .Y = (Abs((Abs(p1.Y) + Abs(p1.Z)) - Abs(p2.X)) + Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.Z)) - Abs((Abs(p1.Z) + Abs(p1.X)) - Abs(p2.Y)))
        .Z = (Abs((Abs(p1.Z) + Abs(p1.X)) - Abs(p2.Y)) + Abs((Abs(p1.Y) + Abs(p1.Z)) - Abs(p2.X)) - Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.Z)))
    End With
End Function

Public Function VectorOffset(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorOffset = New Point
    With VectorOffset
        .X = (LargeOf(p1.X, p2.X) - LeastOf(p1.X, p2.X))
        .Y = (LargeOf(p1.Y, p2.Y) - LeastOf(p1.Y, p2.Y))
        .Z = (LargeOf(p1.Z, p2.Z) - LeastOf(p1.Z, p2.Z))
    End With
End Function

Public Function VectorDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDeduction = New Point
    With VectorDeduction
        .X = (p1.X - p2.X)
        .Y = (p1.Y - p2.Y)
        .Z = (p1.Z - p2.Z)
    End With
End Function

Public Function VectorCrossDeduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCrossDeduct = New Point
    With VectorCrossDeduct
        .X = (p1.X - p2.Z)
        .Y = (p1.Y - p2.X)
        .Z = (p1.Z - p2.Y)
    End With
End Function

Public Function VectorAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAddition = New Point
    With VectorAddition
        .X = (p1.X + p2.X)
        .Y = (p1.Y + p2.Y)
        .Z = (p1.Z + p2.Z)
    End With
End Function

Public Function VectorCombination(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCombination = New Point
    With VectorCombination
        .X = ((p1.X + p2.X) / 2)
        .Y = ((p1.Y + p2.Y) / 2)
        .Z = ((p1.Z + p2.Z) / 2)
    End With
End Function

Public Function VectorIsNormal(ByRef p1 As Point) As Boolean
    VectorIsNormal = (Round(p1.X + p1.Y + p1.Z, 0) = 1)
End Function

Public Function VectorNormalize(ByRef p1 As Point) As Point
    Set VectorNormalize = New Point
    With VectorNormalize
        .Z = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.Z))
        If (.Z > 0) Then
            .Z = (1 / .Z)
            .X = (p1.X * .Z)
            .Y = (p1.Y * .Z)
            .Z = (p1.Z * .Z)
        End If
    End With
End Function

Public Function VectorMidPoint(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMidPoint = New Point
    With VectorMidPoint
        .X = ((LargeOf(p1.X, p2.X) - LeastOf(p1.X, p2.X)) / 2) + LeastOf(p1.X, p2.X)
        .Y = ((LargeOf(p1.Y, p2.Y) - LeastOf(p1.Y, p2.Y)) / 2) + LeastOf(p1.Y, p2.Y)
        .Z = ((LargeOf(p1.Z, p2.Z) - LeastOf(p1.Z, p2.Z)) / 2) + LeastOf(p1.Z, p2.Z)
    End With
End Function

Public Function AbsoluteFactor(ByRef n As Single) As Single
    AbsoluteFactor = ((-(Abs(n - 1) - n) - (-Abs(n + 1) + n)) * 0.5)
End Function

Public Function AbsoluteValue(ByRef n As Single) As Single
    AbsoluteValue = (-((-(n * -1) * n) ^ (1 / 2) * -1))
End Function

Public Function AbsoluteWhole(ByRef n As Single) As Single
    AbsoluteWhole = (n \ 1)
End Function

Public Function AbsoluteDecimal(ByRef n As Single) As Single
    AbsoluteDecimal = (n - AbsoluteWhole(n))
End Function

Public Function LargeOf(ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (v1 >= v2) Then
            LargeOf = v1
        Else
            LargeOf = v2
        End If
    ElseIf IsMissing(V4) Then
        If ((v2 >= V3) And (v2 >= v1)) Then
            LargeOf = v2
        ElseIf ((v1 >= V3) And (v1 >= v2)) Then
            LargeOf = v1
        Else
            LargeOf = V3
        End If
    Else
        If ((v2 >= V3) And (v2 >= v1) And (v2 >= V4)) Then
            LargeOf = v2
        ElseIf ((v1 >= V3) And (v1 >= v2) And (v1 >= V4)) Then
            LargeOf = v1
        ElseIf ((V3 >= v1) And (V3 >= v2) And (V3 >= V4)) Then
            LargeOf = V3
        Else
            LargeOf = V4
        End If
    End If
End Function

Public Function LeastOf(ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (v1 <= v2) Then
            LeastOf = v1
        Else
            LeastOf = v2
        End If
    ElseIf IsMissing(V4) Then
        If ((v2 <= V3) And (v2 <= v1)) Then
            LeastOf = v2
        ElseIf ((v1 <= V3) And (v1 <= v2)) Then
            LeastOf = v1
        Else
            LeastOf = V3
        End If
    Else
        If ((v2 <= V3) And (v2 <= v1) And (v2 <= V4)) Then
            LeastOf = v2
        ElseIf ((v1 <= V3) And (v1 <= v2) And (v1 <= V4)) Then
            LeastOf = v1
        ElseIf ((V3 <= v1) And (V3 <= v2) And (V3 <= V4)) Then
            LeastOf = V3
        Else
            LeastOf = V4
        End If
    End If
End Function
