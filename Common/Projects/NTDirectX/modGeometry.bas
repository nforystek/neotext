Attribute VB_Name = "modGeometry"
    Option Explicit

Option Compare Binary

'#If Not DxVBLibA Then
'
'    Public Type D3DVECTOR
'        X As Single
'        Y As Single
'        Z As Single
'    End Type
'
'#End If

Public Const PI As Single = 3.14159265358979
Public Const Epsilon As Double = 0.999999999999999
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2
Public Const Degree As Single = 180 / PI
Public Const RADIAN As Single = PI / 180
Public Const FOOT As Single = 0.1
Public Const FOVY As Single = (FOOT * 8) '4 feet left, and 4 feet right = 0.8
Public Const FAR  As Single = 90000
Public Const NEAR As Single = 0.05 'one millimeter (308.4 per foor) or greater

Public Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR

    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z

End Function

Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set MakePoint = New Point

    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z

End Function

Public Function ToVector(ByRef Point As Point) As D3DVECTOR

    ToVector.X = Point.X
    ToVector.Y = Point.Y
    ToVector.Z = Point.Z

End Function

Public Function ToPoint(ByRef Vector As D3DVECTOR) As Point
    Set ToPoint = New Point
    ToPoint.X = Vector.X
    ToPoint.Y = Vector.Y
    ToPoint.Z = Vector.Z
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = ((((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = ((((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal n As Single) As Point
    Dim dist As Single
    dist = DistanceEx(p1, p2)
    Set DistanceSet = New Point
    With DistanceSet
        If (dist > 0) And (n > 0) Then
            dist = (((n - dist) / (dist / n)) / n)
            .X = p2.X + ((p2.X - p1.X) * dist)
            .Y = p2.Y + ((p2.Y - p1.Y) * dist)
            .Z = p2.Z + ((p2.Z - p1.Z) * dist)
        ElseIf (n = 0) Then
            .X = p1.X
            .Y = p1.Y
            .Z = p1.Z
        ElseIf (dist = 0) Then
            .X = p2.X
            .Y = p2.Y
            .Z = p2.Z + IIf(p2.Z > p1.Z, n, -n)
        End If
    End With
End Function

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    RandomPositive = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function PlaneNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set PlaneNormal = New Point
    Set PlaneNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(v0, V1), VectorDeduction(V1, V2)))
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

Public Function SquareCenter(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Point
    Set SquareCenter = New Point
    With SquareCenter
        .X = (Least(v0.X, V1.X, V2.X, V3.X) + ((Large(v0.X, V1.X, V2.X, V3.X) - Least(v0.X, V1.X, V2.X, V3.X)) / 2))
        .Y = (Least(v0.Y, V1.Y, V2.Y, V3.Y) + ((Large(v0.Y, V1.Y, V2.Y, V3.Y) - Least(v0.Y, V1.Y, V2.Y, V3.Y)) / 2))
        .Z = (Least(v0.Z, V1.Z, V2.Z, V3.Z) + ((Large(v0.Z, V1.Z, V2.Z, V3.Z) - Least(v0.Z, V1.Z, V2.Z, V3.Z)) / 2))
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
    Dim l1 As Single: l1 = DistanceEx(p1, p2)
    Dim l2 As Single: l2 = DistanceEx(p2, p3)
    Dim l3 As Single: l3 = DistanceEx(p3, p1)
    TriangleSurfaceArea = (((((((l1 + l2) - l3) + ((l2 + l3) - l1) + ((l3 + l1) - l2)) * (l1 * l2 * l3)) / (l1 + l2 + l3)) ^ (1 / 2)))
End Function

Public Function TriangleVolume(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    TriangleVolume = TriangleSurfaceArea(p1, p2, p3)
    TriangleVolume = ((((TriangleVolume ^ (1 / 3)) ^ 2) ^ 3) / 12)
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
        .X = (Large(p1.X, p2.X, p3.X) - Least(p1.X, p2.X, p3.X))
        .Y = (Large(p1.Y, p2.Y, p3.Y) - Least(p1.Y, p2.Y, p3.Y))
        .Z = (Large(p1.Z, p2.Z, p3.Z) - Least(p1.Z, p2.Z, p3.Z))
    End With
End Function

Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleAxii = New Point
    With TriangleAxii
        Dim o As Point
        Set o = TriangleOffset(p1, p2, p3)
        .X = (Least(p1.X, p2.X, p3.X) + (o.X / 2))
        .Y = (Least(p1.Y, p2.Y, p3.Y) + (o.Y / 2))
        .Z = (Least(p1.Z, p2.Z, p3.Z) + (o.Z / 2))
    End With
End Function

Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleNormal = New Point
    Dim o As Point
    Dim d As Single
    With TriangleNormal
        Set o = TriangleOffset(v0, V1, V2)
        d = (o.X + o.Y + o.Z)
        .Z = (((o.X + o.Y) - o.Z) / d)
        .X = (((o.Y + o.Z) - o.X) / d)
        .Y = (((o.Z + o.X) - o.Y) / d)
    End With
End Function

Public Function TriangleAccordance(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleAccordance = New Point
    With TriangleAccordance
        .X = (((v0.X + V1.X) - V2.X) + ((V1.X + V2.X) - v0.X) - ((V2.X + v0.X) - V1.X))
        .Y = (((v0.Y + V1.Y) - V2.Y) + ((V1.Y + V2.Y) - v0.Y) - ((V2.Y + v0.Y) - V1.Y))
        .Z = (((v0.Z + V1.Z) - V2.Z) + ((V1.Z + V2.Z) - v0.Z) - ((V2.Z + v0.Z) - V1.Z))
    End With
End Function

Public Function TriangleDisplace(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleDisplace = New Point
    With TriangleDisplace
        .X = (Abs((Abs(v0.X) + Abs(V1.X)) - Abs(V2.X)) + Abs((Abs(V1.X) + Abs(V2.X)) - Abs(v0.X)) - Abs((Abs(V2.X) + Abs(v0.X)) - Abs(V1.X)))
        .Y = (Abs((Abs(v0.Y) + Abs(V1.Y)) - Abs(V2.Y)) + Abs((Abs(V1.Y) + Abs(V2.Y)) - Abs(v0.Y)) - Abs((Abs(V2.Y) + Abs(v0.Y)) - Abs(V1.Y)))
        .Z = (Abs((Abs(v0.Z) + Abs(V1.Z)) - Abs(V2.Z)) + Abs((Abs(V1.Z) + Abs(V2.Z)) - Abs(v0.Z)) - Abs((Abs(V2.Z) + Abs(v0.Z)) - Abs(V1.Z)))
    End With
End Function

Public Function VectorRotateAxis(ByRef p1 As Point, ByRef Angles As Point) As Point
    'AnglesStrict Angles

    Set VectorRotateAxis = VectorRotateX(p1, Angles.X)
    Set VectorRotateAxis = VectorRotateY(VectorRotateAxis, Angles.Y)
    Set VectorRotateAxis = VectorRotateZ(VectorRotateAxis, Angles.Z)

End Function

Public Function VectorRotateX(ByRef p1 As Point, ByVal angle As Single) As Point
    Set VectorRotateX = New Point
    With VectorRotateX
        .X = p1.X
        .Y = p1.Y * Cos(angle)
        .Z = p1.Z * Cos(angle)
    End With
End Function

Public Function VectorRotateY(ByRef p1 As Point, ByVal angle As Single) As Point
    Set VectorRotateY = New Point
    With VectorRotateY
        .X = p1.X * Cos(angle)
        .Y = p1.Y
        .Z = p1.Z * Cos(angle)
    End With
End Function

Public Function VectorRotateZ(ByRef p1 As Point, ByVal angle As Single) As Point
    Set VectorRotateZ = New Point
    With VectorRotateZ
        .X = p1.X * Cos(angle)
        .Y = p1.Y * Cos(angle)
        .Z = p1.Z
    End With
End Function

Public Function VectorAxisAngles(ByRef p1 As Point) As Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles

        .Z = VectorZAngle(p1.X, p1.Y)
        .Y = VectorZAngle(p1.Z, p1.X)
        .X = VectorZAngle(p1.Y, p1.Z)
        
    End With
    AnglesStrict VectorAxisAngles
End Function

Public Function VectorZAngle(ByVal pX As Single, ByVal pY As Single) As Single
    If pX = 0 Then
        If pX > 0 Then
            VectorZAngle = 90
        ElseIf pX < 0 Then
            VectorZAngle = 270
        End If
    ElseIf pX = 0 Then
        If pX > 0 Then
            VectorZAngle = 360
        ElseIf pX < 0 Then
            VectorZAngle = 180
        End If
    Else
        If pX > 0 And pX > 0 Then
            VectorZAngle = (90 * (Abs(pY) / Abs(pX)))
        ElseIf pX < 0 And pX > 0 Then
            VectorZAngle = (90 * (Abs(pX) / Abs(pY))) + 90
        ElseIf pX < 0 And pX < 0 Then
            VectorZAngle = (90 * (Abs(pY) / Abs(pX))) + 180
        ElseIf pX > 0 And pX < 0 Then
            VectorZAngle = (90 * (Abs(pX) / Abs(pY))) + 270
        End If
    End If

    VectorZAngle = VectorZAngle * RADIAN
    AngleStrict VectorZAngle
End Function

Public Function AngleStrict(ByRef p As Single) As Single
'    If (p \ 360) > 0 Then p = (Abs(p) - ((Abs(p) \ 360) * 360)) * AbsoluteFactor(p)
'    If (p < 0) Then p = p + 360
'    Do While p < 0
'        p = p + 360
'    Loop
'    Do While p > 360
'        p = p - 360
'    Loop
    Do While p = 0 Or p <= -(PI * 2)
        p = p + (PI * 2)
    Loop
    Do While p >= (PI * 2)
        p = p - (PI * 2)
    Loop
    AngleStrict = p
End Function

Public Sub AnglesStrict(ByRef p As Point)
    p.X = AngleStrict(p.X)
    p.Y = AngleStrict(p.Y)
    p.Z = AngleStrict(p.Z)
End Sub

'Public Function VectorAngleAxis(ByRef p1 As Point) As Point
'
'    Dim p As Point
'    Set p = VectorNormalize(p1)
'
'    Set VectorAngleAxis = New Point
'    With VectorAngleAxis
'        .X = Round(AngleXDegree(p), 6)
'        .Y = Round(AngleYDegree(p), 6)
'        .Z = Round(AngleZDegree(p), 6)
'    End With
'
'    AngleStrictRadian VectorAngleAxis
'End Function
'
'Public Function AngleXDegree(ByRef p As Point) As Single
'
'    AngleXDegree = ((AngleDegree(p.X - Abs(p.Z), p.Y + p.Z) + (PI - ((PI / 4) * 3) * 1)) + (AngleDegree(p.Y - p.Z, p.X + Abs(p.Z))) + ((PI - ((PI / 4) * 3)) * 3)) + (PI / 2)
'
'End Function
'
'Public Function AngleYDegree(ByRef p As Point) As Single
'
'    AngleYDegree = (((AngleDegree(p.Z - Abs(p.X), p.Z + p.Y) - (PI - ((PI / 4) * 3) * 1)) + (AngleDegree(p.Z - p.Y, p.Z + Abs(p.X))) + ((PI - ((PI / 4) * 3)) * 1)))
'
'End Function
'
'Public Function AngleZDegree(ByRef p As Point) As Single
'
'    AngleZDegree = (((AngleDegree(p.Z - Abs(p.X), (p.Y / 2) + p.X) / 2) + (AngleDegree((p.Y / 2) - p.X, p.Z + Abs(p.X)) / 2)) * 2)
'
'End Function
'
'Public Function AngleDegree(ByVal pX As Single, ByVal pY As Single) As Single
'
'    If pX > 0 And pY > 0 Then
'        AngleDegree = 0
'    ElseIf pX < 0 And pY > 0 Then
'        AngleDegree = (PI / 2)
'    ElseIf pX < 0 And pY < 0 Then
'        AngleDegree = PI
'    ElseIf pX > 0 And pY < 0 Then
'        AngleDegree = ((PI / 2) * 3)
'    End If
'
'    If pX <> 0 And pY <> 0 Then
'        If Abs(pY) / Abs(pX) < 0.5 Then
'            If AngleDegree = ((PI / 2) * 3) Then
'                AngleDegree = AngleDegree + AngleOfTwoPoints(0, 1, Abs(pX), Abs(pY)) + (PI / 2)
'            ElseIf AngleDegree = (PI / 2) Then
'                AngleDegree = (AngleDegree + AngleOfTwoPoints(0, 1, Abs(pX), Abs(pY))) + (PI / 2)
'            Else
'                AngleDegree = AngleDegree + (-PI - AngleOfTwoPoints(0, 1, Abs(pX), Abs(pY))) + (PI * 2)
'            End If
'        Else
'            If AngleDegree = ((PI / 2) * 3) Then
'                AngleDegree = AngleDegree + AngleOfTwoPoints(1, 0, Abs(pY), Abs(pX)) + (PI / 2)
'            ElseIf AngleDegree = (PI / 2) Then
'                AngleDegree = (AngleDegree + AngleOfTwoPoints(1, 0, Abs(pY), Abs(pX))) + (PI / 2)
'            Else
'                AngleDegree = AngleDegree + (-PI - AngleOfTwoPoints(1, 0, Abs(pY), Abs(pX))) + (PI * 2)
'            End If
'        End If
'    End If
'
'    AngleStrictDegree AngleDegree
'End Function
'
'Public Function AngleOfTwoPoints(ByVal p1x As Single, ByVal p1y As Single, ByVal p2x As Single, ByVal p2y As Single) As Single
'    Dim p1 As Single
'    Dim p2 As Single
'    p1 = ((p1x * p2x) + (p1y * p2y))
'    p2 = ((p1x ^ 2 + p1y ^ 2) ^ (1 / 2)) * ((p2x ^ 2 + p2y ^ 2) ^ (1 / 2))
'    If p2 > 0 Then AngleOfTwoPoints = ArcCos(p1 / p2)
'End Function
'
'Public Function ArcSin(ByVal X As Double) As Double
'    ArcSin = Atn(X / Sqr(-X * X + 1))
'End Function
'
'Public Function ArcCos(ByVal X As Double) As Double
'    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
'End Function

Public Function VectorRise(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRise = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
End Function

Public Function VectorRun(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRun = DistanceEx(MakePoint(p1.X, 0, p1.Z), MakePoint(p2.X, 0, p2.Z))
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

Public Function VectorMultiply(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMultiply = New Point
    With VectorMultiply
        .X = (p1.X * p2.X)
        .Y = (p1.Y * p2.Y)
        .Z = (p1.Z * p2.Z)
    End With
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
        .X = (Large(p1.X, p2.X) - Least(p1.X, p2.X))
        .Y = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
        .Z = (Large(p1.Z, p2.Z) - Least(p1.Z, p2.Z))
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

Public Function VectorCrossDeduct(ByRef p1 As Point, ByRef p2 As Point)
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

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .X = (p1.X * n)
        .Y = (p1.Y * n)
        .Z = (p1.Z * n)
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
        .X = ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) / 2) + Least(p1.X, p2.X)
        .Y = ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) / 2) + Least(p1.Y, p2.Y)
        .Z = ((Large(p1.Z, p2.Z) - Least(p1.Z, p2.Z)) / 2) + Least(p1.Z, p2.Z)
    End With
End Function

Public Function VectorNegative(ByRef p1 As Point) As Point
    Set VectorNegative = New Point
    With VectorNegative
        .X = -p1.X
        .Y = -p1.Y
        .Z = -p1.Z
    End With
End Function

Public Function VectorDivision(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorDivision = New Point
    With VectorDivision
        .X = (p1.X / n)
        .Y = (p1.Y / n)
        .Z = (p1.Z / n)
    End With
End Function



Public Function VectorIsNormal(ByRef p1 As Point) As Boolean
    VectorIsNormal = (Round(p1.X + p1.Y + p1.Z, 0) = 1)
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


Public Function AngleQuadrant(ByVal angle As Single) As Single
    If angle > 0 And angle <= 90 Then
        AngleQuadrant = 1
    ElseIf angle > 90 And angle <= 180 Then
        AngleQuadrant = 2
    ElseIf angle > 180 And angle <= 270 Then
        AngleQuadrant = 3
    ElseIf angle > 270 And angle <= 360 Then
        AngleQuadrant = 4
    End If
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

