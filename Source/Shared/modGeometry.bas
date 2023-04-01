Attribute VB_Name = "modGeometry"
Option Explicit

Option Compare Binary

Public Const PI As Single = 3.14159265358979
Public Const epsilon As Double = 0.999999999999999
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2
Public Const DEGREE As Single = 180 / PI
Public Const RADIAN As Single = PI / 180
Public Const FOOT As Single = 0.1
Public Const MILE As Single = 5280 * FOOT
Public Const FOVY As Single = (FOOT * 8) '4 feet left, and 4 feet right = 0.8
Public Const Far  As Single = 900000000
Public Const Near As Single = 0 '0.05 'one millimeter (308.4 per foor) or greater


Public Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.z = z
End Function

Public Function MakePoint(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Point
    Set MakePoint = New Point
    MakePoint.x = x
    MakePoint.y = y
    MakePoint.z = z
End Function

Public Function MakeCoord(ByVal x As Single, ByVal y As Single) As Coord
    Set MakeCoord = New Coord
    MakeCoord.x = x
    MakeCoord.y = y
End Function

Public Function ToCoord(ByRef Vector As D3DVECTOR) As Coord
    Set ToCoord = New Coord
    ToCoord.x = Vector.x
    ToCoord.y = Vector.y
End Function

Public Function ToVector(ByRef Point As Point) As D3DVECTOR
    ToVector.x = Point.x
    ToVector.y = Point.y
    ToVector.z = Point.z
End Function

Public Function ToPoint(ByRef Vector As D3DVECTOR) As Point
    Set ToPoint = New Point
    ToPoint.x = Vector.x
    ToPoint.y = Vector.y
    ToPoint.z = Vector.z
End Function

Public Function ToPlane(ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Range
        
    Dim pNormal As Point
    Set pNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(V2, V1), VectorDeduction(V3, V1)))
        
    Set ToPlane = New Range
    With ToPlane
        .W = VectorDotProduct(pNormal, V1) * -1
        .x = pNormal.x
        .y = pNormal.y
        .z = pNormal.z
    End With
End Function

Public Function ToVec4(ByRef Plane As Range) As D3DVECTOR4
    ToVec4.x = Plane.x
    ToVec4.y = Plane.y
    ToVec4.z = Plane.z
    ToVec4.W = Plane.W
End Function

Public Function DistanceToPlane(ByRef p As Point, ByRef r As Range) As Single
    DistanceToPlane = (r.x * p.x + r.y * p.y + r.z * p.z + r.W) / Sqr(r.x * r.x + r.y * r.y + r.z * r.z)
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = ((((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = ((((p1.x - p2.x) ^ 2) + ((p1.y - p2.y) ^ 2) + ((p1.z - p2.z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal n As Single) As Point
    Dim dist As Single
    dist = DistanceEx(p1, p2)
    Set DistanceSet = New Point
    With DistanceSet
        If Not (dist = n) Then
            If ((dist > 0) And (n > 0)) Then
                .x = Large(p1.x, p2.x) - Least(p1.x, p2.x)
                .y = Large(p1.y, p2.y) - Least(p1.y, p2.y)
                .z = Large(p1.z, p2.z) - Least(p1.z, p2.z)
                .x = (Least(p1.x, p2.x) + (n * (.x / dist)))
                .y = (Least(p1.y, p2.y) + (n * (.y / dist)))
                .z = (Least(p1.z, p2.z) + (n * (.z / dist)))
            ElseIf (n = 0) Then
                .x = p1.x
                .y = p1.y
                .z = p1.z
            ElseIf (dist = 0) Then
                .x = p2.x
                .y = p2.y
                .z = p2.z + IIf(p2.z > p1.z, n, -n)
            End If
        End If
    End With
End Function

Public Function PointOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    Dim r As Range
    Set r = ToPlane(v0, V1, V2)
    PointOnPlane = (r.x * (p.x - v0.x)) + (r.y * (p.y - v0.y)) + (r.z * (p.z - v0.z)) = 0
End Function
Public Function PointSideOfPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    PointSideOfPlane = VectorDotProduct(PlaneNormal(v0, V1, V2), p) > 0
End Function

Public Function PointOnPlaneNearestPoint(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Point

    Set PointOnPlaneNearestPoint = New Point
    With PointOnPlaneNearestPoint
    
        Dim r As Range
        Set r = ToPlane(v0, V1, V2)
        
        Dim n As Point
        Set n = PlaneNormal(v0, V1, V2)
    
        Dim d As Single
        d = DistanceToPlane(p, r)
        
        .x = (d * n.x)
        .y = (d * n.y)
        .z = (d * n.z)
        
    End With
    Set PointOnPlaneNearestPoint = VectorAddition(p, PointOnPlaneNearestPoint)
    
End Function

Public Function LineIntersectPlane(ByRef Plane As Range, PStart As Point, vDir As Point, ByRef VIntersectOut As Point) As Boolean
    Dim q As New Range     'Start Point
    Dim V As New Range       'Vector Direction

    Dim planeQdot As Single 'Dot products
    Dim planeVdot As Single
    
    Dim t As Single         'Part of the equation for a ray P(t) = Q + tV

    
    q.x = PStart.x          'Q is a point and therefore it's W value is 1
    q.y = PStart.y
    q.z = PStart.z
    q.W = 1
    
    V.x = vDir.x            'V is a vector and therefore it's W value is zero
    V.y = vDir.y
    V.z = vDir.z
    V.W = 0
    
  '  ((Plane.X * V.X) + (Plane.Y * V.Y) + (Plane.z * V.z) + (Plane.w * V.w))
    
    planeVdot = ((Plane.x * V.x) + (Plane.y * V.y) + (Plane.z * V.z) + (Plane.W * V.W)) 'D3DXVec4Dot(Plane, V)
    planeQdot = ((Plane.x * q.x) + (Plane.y * q.y) + (Plane.z * q.z) + (Plane.W * q.W)) 'D3DXVec4Dot(Plane, Q)
            
    'If the dotproduct of plane and V = 0 then there is no intersection
    If planeVdot <> 0 Then
        t = Round((planeQdot / planeVdot) * -1, 5)
        
        If VIntersectOut Is Nothing Then Set VIntersectOut = New Point
        
        'This is where the line intersects the plane
        VIntersectOut.x = Round(q.x + (t * V.x), 5)
        VIntersectOut.y = Round(q.y + (t * V.y), 5)
        VIntersectOut.z = Round(q.z + (t * V.z), 5)

        LineIntersectPlane = True
    Else
        'No Collision
        LineIntersectPlane = False
    End If
    
End Function

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    RandomPositive = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function PlaneNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    'returns a vector perpendicular to a plane V, at 0,0,0, with out the local coordinates information
    Set PlaneNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(v0, V1), VectorDeduction(V1, V2)))
End Function

Public Function PointNormalize(ByRef V As Point) As Point
    Set PointNormalize = New Point
    With PointNormalize
        .z = (V.x ^ 2 + V.y ^ 2 + V.z ^ 2) ^ (1 / 2)
        If (.z = 0) Then .z = 1
        .x = (V.x / .z)
        .y = (V.y / .z)
        .z = (V.z / .z)
    End With
End Function
Public Function Sign(ByVal n As Single) As Single
    Sign = ((-(Abs(n - 1) - n) - (-Abs(n + 1) + n)) * 0.5)
End Function

Public Function Signn(ByVal Value As Single) As Single
    Signn = ((-((Value \ 1 <> 0) * 1) + -1) + -(((-Value \ 1 + -1) = 0) * 1))
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
        .x = (Least(v0.x, V1.x, V2.x, V3.x) + ((Large(v0.x, V1.x, V2.x, V3.x) - Least(v0.x, V1.x, V2.x, V3.x)) / 2))
        .y = (Least(v0.y, V1.y, V2.y, V3.y) + ((Large(v0.y, V1.y, V2.y, V3.y) - Least(v0.y, V1.y, V2.y, V3.y)) / 2))
        .z = (Least(v0.z, V1.z, V2.z, V3.z) + ((Large(v0.z, V1.z, V2.z, V3.z) - Least(v0.z, V1.z, V2.z, V3.z)) / 2))
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
        .x = ((p1.x + p2.x + p3.x) / 3)
        .y = ((p1.y + p2.y + p3.y) / 3)
        .z = ((p1.z + p2.z + p3.z) / 3)
    End With
End Function

Public Function TriangleOffset(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleOffset = New Point
    With TriangleOffset
        .x = (Large(p1.x, p2.x, p3.x) - Least(p1.x, p2.x, p3.x))
        .y = (Large(p1.y, p2.y, p3.y) - Least(p1.y, p2.y, p3.y))
        .z = (Large(p1.z, p2.z, p3.z) - Least(p1.z, p2.z, p3.z))
    End With
End Function

Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleAxii = New Point
    With TriangleAxii
        Dim o As Point
        Set o = TriangleOffset(p1, p2, p3)
        .x = (Least(p1.x, p2.x, p3.x) + (o.x / 2))
        .y = (Least(p1.y, p2.y, p3.y) + (o.y / 2))
        .z = (Least(p1.z, p2.z, p3.z) + (o.z / 2))
    End With
End Function

Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleNormal = New Point
    Dim o As Point
    Dim d As Single
    With TriangleNormal
        Set o = TriangleDisplace(v0, V1, V2)
        d = (o.x + o.y + o.z)
        If (d > 0) Then
            .z = (((o.x + o.y) - o.z) / d)
            .x = (((o.y + o.z) - o.x) / d)
            .y = (((o.z + o.x) - o.y) / d)
        End If
    End With
End Function

Public Function TriangleAccordance(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleAccordance = New Point
    With TriangleAccordance
        .x = (((v0.x + V1.x) - V2.x) + ((V1.x + V2.x) - v0.x) - ((V2.x + v0.x) - V1.x))
        .y = (((v0.y + V1.y) - V2.y) + ((V1.y + V2.y) - v0.y) - ((V2.y + v0.y) - V1.y))
        .z = (((v0.z + V1.z) - V2.z) + ((V1.z + V2.z) - v0.z) - ((V2.z + v0.z) - V1.z))
    End With
End Function

Public Function TriangleDisplace(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleDisplace = New Point
    With TriangleDisplace
        .x = (Abs((Abs(v0.x) + Abs(V1.x)) - Abs(V2.x)) + Abs((Abs(V1.x) + Abs(V2.x)) - Abs(v0.x)) - Abs((Abs(V2.x) + Abs(v0.x)) - Abs(V1.x)))
        .y = (Abs((Abs(v0.y) + Abs(V1.y)) - Abs(V2.y)) + Abs((Abs(V1.y) + Abs(V2.y)) - Abs(v0.y)) - Abs((Abs(V2.y) + Abs(v0.y)) - Abs(V1.y)))
        .z = (Abs((Abs(v0.z) + Abs(V1.z)) - Abs(V2.z)) + Abs((Abs(V1.z) + Abs(V2.z)) - Abs(v0.z)) - Abs((Abs(V2.z) + Abs(v0.z)) - Abs(V1.z)))
    End With
End Function

Public Function VectorBalance(ByRef loZero As Point, ByRef hiWhole As Point, ByVal folcrumPercent As Single) As Point
    Set VectorBalance = New Point
    With VectorBalance
        .x = (loZero.x + ((hiWhole.x - loZero.x) * folcrumPercent))
        .y = (loZero.y + ((hiWhole.y - loZero.y) * folcrumPercent))
        .z = (loZero.z + ((hiWhole.z - loZero.z) * folcrumPercent))
    End With
End Function

Public Function TriangleFolcrum(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Point
    Set TriangleFolcrum = New Point
    With TriangleFolcrum
        If (Not p3 Is Nothing) Then
            .x = (p3.x ^ 2)
            .y = (p3.y ^ 2)
            .z = (p3.z ^ 2)
        End If
        .x = (.x + (p1.x ^ 2) + (p2.x ^ 2)) ^ (1 / 2)
        .y = (.y + (p1.y ^ 2) + (p2.y ^ 2)) ^ (1 / 2)
        .z = (.z + (p1.z ^ 2) + (p2.z ^ 2)) ^ (1 / 2)
    End With
End Function

Public Function TriangleOpposite(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    Dim hypo As Single
    TriangleOpposite = DistanceEx(p1, p2)
    hypo = DistanceEx(p2, p3)
    If hypo < TriangleOpposite Then
        TriangleOpposite = ((TriangleOpposite ^ 2) - (hypo ^ 2)) ^ (1 / 2)
    Else
        TriangleOpposite = ((hypo ^ 2) - (TriangleOpposite ^ 2)) ^ (1 / 2)
    End If
End Function

Public Function TriangleAdjacent(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Single
    TriangleAdjacent = DistanceEx(p1, p2)
    If Not p3 Is Nothing Then
        Dim l1 As Single
        Dim l2 As Single
        l1 = DistanceEx(p2, p3)
        l2 = DistanceEx(p3, p1)
        If TriangleAdjacent < l1 Xor TriangleAdjacent < l2 Then
            If TriangleAdjacent < l1 Then
                If l1 > l2 Then
                    TriangleAdjacent = ((TriangleAdjacent ^ 2) - (l2 ^ 2)) ^ (1 / 2)
                Else
                    TriangleAdjacent = ((TriangleAdjacent ^ 2) - (l1 ^ 2)) ^ (1 / 2)
                End If
            Else
                If l1 > l2 Then
                    If l2 > TriangleAdjacent Then
                        TriangleAdjacent = ((l1 ^ 2) - (TriangleAdjacent ^ 2)) ^ (1 / 2)
                    Else
                        TriangleAdjacent = ((l1 ^ 2) - (l2 ^ 2)) ^ (1 / 2)
                    End If
                Else
                    If l1 > TriangleAdjacent Then
                        TriangleAdjacent = ((l2 ^ 2) - (TriangleAdjacent ^ 2)) ^ (1 / 2)
                    Else
                        TriangleAdjacent = ((l2 ^ 2) - (l1 ^ 2)) ^ (1 / 2)
                    End If
                End If
            End If
        End If
    Else
        TriangleAdjacent = (((TriangleAdjacent ^ 2) / 2) ^ (1 / 2))
    End If
End Function

Public Function TriangleHypotenuse(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Single
    TriangleHypotenuse = DistanceEx(p1, p2)
    If p3 Is Nothing Then
        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (TriangleHypotenuse ^ 2)) ^ (1 / 2)
    Else
        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (DistanceEx(p2, p3) ^ 2)) ^ (1 / 2)
    End If
End Function


Public Sub AngleAxisRestrict(ByRef p As Point)
    p.x = AngleRestrict(p.x)
    p.y = AngleRestrict(p.y)
    p.z = AngleRestrict(p.z)
End Sub
Public Function AngleRestrict(ByRef A As Single) As Single
    A = Round(A, 6)
    AngleRestrict = 0
    Do While A < -PI * 2 And AngleRestrict <= 4
        A = A + (PI * 2)
        AngleRestrict = AngleRestrict + 1
    Loop
    AngleRestrict = 0
    Do While A > PI * 2 And AngleRestrict <= 4
        A = A - (PI * 2)
        AngleRestrict = AngleRestrict + 1
    Loop
    AngleRestrict = A
End Function

Public Function VectorOctet(ByRef p As Point) As Single
    VectorOctet = VectorQuadrant(p)
    If p.z < 0 Then
        Select Case VectorOctet
            Case 1
                VectorOctet = 6
            Case 2
                VectorOctet = 7
            Case 3
                VectorOctet = 8
            Case 4
                VectorOctet = 5
        End Select
    End If
End Function

'Public Function ATan2(Y As Single, X As Single) As Single
'  If X > 0 Then
'    ATan2 = Atn(Y / X)
'  ElseIf X < 0 And Y >= 0 Then
'    ATan2 = Atn(Y / X) + PI
'  ElseIf X < 0 And Y < 0 Then
'    ATan2 = Atn(Y / X) - PI
'  ElseIf X = 0 And Y > 0 Then
'    ATan2 = PI / 2
'  ElseIf X = 0 And Y < 0 Then
'    ATan2 = -PI / 2
'  End If
'End Function

'Public Function ATan2(Y As Single, X As Single) As Single
'  If X > 0 Then
'    ATan2 = Atn(Y / X)
'  ElseIf X < 0 And Y >= 0 Then
'    ATan2 = Atn(Y / X) + PI
'  ElseIf X < 0 And Y < 0 Then
'    ATan2 = Atn(Y / X) - PI
'  ElseIf X = 0 And Y > 0 Then
'    ATan2 = PI / 2
'  ElseIf X = 0 And Y < 0 Then
'    ATan2 = -PI / 2
'  End If
'End Function


Public Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
    If Abs(adj) < 0.000001 Then
        ATan2 = PI / 2
    Else
        ATan2 = Abs(Atn(opp / adj))
    End If
    If adj < 0 Then ATan2 = PI - ATan2
    If opp < 0 Then ATan2 = -ATan2
End Function

Public Function InvSin(Number As Single) As Single
    InvSin = -Number * Number + 1
    If InvSin > 0 Then
        InvSin = Sqr(InvSin)
        If InvSin <> 0 Then InvSin = Atn(Number / InvSin)
    Else
        InvSin = 0
    End If
End Function

Public Function VectorRotateAxis(ByRef p1 As Point, ByRef Angles As Point) As Point

'    Set VectorRotateAxis = New Point
'    VectorRotateAxis = VectorRotateX(p1, Angles.X)
'    VectorRotateAxis = VectorRotateY(p1, Angles.Y)
'    VectorRotateAxis = VectorRotateZ(p1, Angles.z)

    Dim S As Single
    Dim tmp As New Point
    Set VectorRotateAxis = New Point
    With VectorRotateAxis
        .y = Cos(Angles.x) * p1.y - Sin(Angles.x) * p1.z
        .z = Sin(Angles.x) * p1.y + Cos(Angles.x) * p1.z
        tmp.x = p1.x
        tmp.y = .y
        tmp.z = .z
        .x = Sin(Angles.y) * tmp.z + Cos(Angles.y) * tmp.x
        .z = Cos(Angles.y) * tmp.z - Sin(Angles.y) * tmp.x
        tmp.x = .x
        tmp.z = .z
        .x = Cos(Angles.z) * tmp.x - Sin(Angles.z) * tmp.y
        .y = Sin(Angles.z) * tmp.x + Cos(Angles.z) * tmp.y
        .z = tmp.z
    End With
    Set tmp = Nothing
    
    
End Function

Public Function VectorRotateX(ByRef p1 As Point, ByVal angle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    
    RadAngle = (angle * RADIAN)
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
    
    Set VectorRotateX = New Point
    With VectorRotateX

        .z = p1.z * CosPhi - p1.y * SinPhi
        .y = p1.z * SinPhi + p1.y * CosPhi
        .x = p1.x


'        .x = p1.x
'        .Y = Cos(angle) * p1.Y - Sin(angle) * p1.z
'        .z = Sin(angle) * p1.Y + Cos(angle) * p1.z
    End With
End Function

Public Function VectorRotateY(ByRef p1 As Point, ByVal angle As Single) As Point

    
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    
    RadAngle = (angle * RADIAN)
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
 
    Set VectorRotateY = New Point
    With VectorRotateY

        .x = p1.x * CosPhi - p1.z * SinPhi
        .z = p1.x * SinPhi + p1.z * CosPhi
        .y = p1.y
    


'        .x = Cos(angle) * p1.x - Sin(angle) * p1.z
'        .z = Sin(angle) * p1.x + Cos(angle) * p1.z
'        .Y = p1.Y
    End With
End Function

Public Function VectorRotateZ(ByRef p1 As Point, ByVal angle As Single) As Point

    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    
    RadAngle = (angle * RADIAN) * -1
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
    
    Set VectorRotateZ = New Point
    With VectorRotateZ

        .x = p1.x * CosPhi - p1.y * SinPhi
        .y = p1.x * SinPhi + p1.y * CosPhi
        .z = p1.z
        

        
'        .Y = Cos(angle) * p1.x - Sin(angle) * p1.Y
'        .x = Sin(angle) * p1.x + Cos(angle) * p1.Y
'        .z = p1.z
    End With
End Function


'Public Function VectorRotateAxis(ByRef p As Point, ByRef Angles As Point) As Point
'    Dim tmp As New Point
'    Set VectorRotateAxis = New Point
'    With VectorRotateAxis
'
'        Dim X As Point
'        Dim Y As Point
'        Dim z As Point
'        Set X = VectorRotateX(p, Angles.X)
'        Set Y = VectorRotateY(X, Angles.Y)
'        Set z = VectorRotateZ(Y, Angles.z)
'        Set VectorRotateAxis = z
'
''        Dim Var32_1 As Long
''        Dim Var32_2 As Long
''        Dim x1 As Long
''        Dim y1 As Long
''        Dim z1 As Long
' '       Dim mp As Point  'temporary
''        Dim S As New Point 'sine
''        Dim c As New Point 'cosine
''
'''            Dim d As Single
'''            d = DistanceEx(MakePoint(0, 0, 0), p)
'''            Dim t As New Point 'tangent
'''            Dim ss As New Point 'cosecant
'''            Dim cc As New Point 'secant
'''            Dim tt As New Point 'cotangent
'''            Dim A As New Point 'angle bulk
'''            Dim r As New Point 'remainder
'''            Dim i As New Point 'in-between
'''
''            Set mp = MakePoint(p.z, p.Y, p.X)
'''            A.x = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'''            r.x = (A.x - ((A.x \ 90) * 90)) * RADIAN
''            S.x = VectorSine(mp)
''            c.x = VectorCosine(mp)
'''            t.x = VectorTangent(mp)
'''            ss.x = VectorCosecant(mp)
'''            cc.x = VectorSecant(mp)
'''            tt.x = VectorCotangent(mp)
'''
''            Set mp = MakePoint(p.x, p.z, p.y)
'''            A.y = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'''            r.y = (A.y - ((A.y \ 90) * 90)) * RADIAN
''            S.y = VectorSine(mp)
''            c.y = VectorCosine(mp)
'''            t.y = VectorTangent(mp)
'''            ss.y = VectorCosecant(mp)
'''            cc.y = VectorSecant(mp)
'''            tt.y = VectorCotangent(mp)
'''
''            Set mp = MakePoint(p.x, p.y, p.z)
'''            A.z = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'''            r.z = (A.z - ((A.z \ 90) * 90)) * RADIAN
''            S.z = VectorSine(mp)
''            c.z = VectorCosine(mp)
'''            t.z = VectorTangent(mp)
'''            ss.z = VectorCosecant(mp)
'''            cc.z = VectorSecant(mp)
'''            tt.z = VectorCotangent(mp)
''
''
''
''
''        Var32_1 = CLng((c.y * p.x) / 256)
''        Var32_2 = CLng((S.y * p.z) / 256)
''        x1 = Var32_1 - Var32_2
''        Var32_1 = CLng((S.y * p.x) / 256)
''        Var32_2 = CLng((c.y * p.z) / 256)
''        z1 = Var32_1 + Var32_2
''        Var32_1 = CLng((c.z * x1) / 256)
''        Var32_2 = CLng((S.z * p.y) / 256)
''        .x = Var32_1 + Var32_2
''        Var32_1 = CLng((c.z * p.y) / 256)
''        Var32_2 = CLng((S.z * x1) / 256)
''        y1 = Var32_1 - Var32_2
''        Var32_1 = CLng((c.x * z1) / 256)
''        Var32_2 = CLng((S.x * y1) / 256)
''        .z = Var32_1 - Var32_2
''        Var32_1 = CLng((S.x * z1) / 256)
''        Var32_2 = CLng((c.x * y1) / 256)
''        .y = Var32_1 + Var32_2
'
''        Dim a As New Point
''        Set a = VectorAxisAngles(p)
''        a.x = a.x + Angles.x
''        a.y = a.y + Angles.y
''        a.z = a.z + Angles.z
'''
''        Dim r1 As Single
''        Dim r2 As Single
''        Dim r3 As Single
''
''        Dim n As Point
''        Set n = VectorNormalize(p)
''
''        Dim n1 As Point
''        Set n1 = VectorNormalize(a)
''
''
''
''        r1 = (p.y ^ 2 + p.z ^ 2) ^ (1 / 2)
''        r2 = (p.x ^ 2 + p.z ^ 2) ^ (1 / 2)
''        r3 = (p.y ^ 2 + p.x ^ 2) ^ (1 / 2)
''
''
''        .x = (r3 * Sin(a.z)) + (r2 * Cos(a.y))
''        .y = (r3 * Cos(a.z)) + (r1 * Sin(a.x))
''        .z = (r2 * Sin(a.y)) + (r1 * Cos(a.x))
'
''        .x = (r1 * Tan(a.x)) * (r3 * Cos(a.z)) * (r2 * Cos(a.y)) * (r1 * Tan(a.x)) * n.x
''        .y = (r2 * Tan(a.y)) * (r3 * Sin(a.z)) * (r1 * Cos(a.x)) * (r2 * Tan(a.y)) * n.y
''        .z = (r3 * Tan(a.z)) * (r2 * Sin(a.y)) * (r1 * Sin(a.x)) * (r3 * Tan(a.z)) * n.z
'
'        'Dim r As Single
'
'
''        r = (p.y ^ 2 + p.z ^ 2) ^ (1 / 2)
''        x.x = p.x
''        x.y = r * Cos(a.x + .x)
''        x.z = r * Sin(a.x + .x)
''
''        r = (p.x ^ 2 + p.z ^ 2) ^ (1 / 2)
''        y.y = p.y
''        y.x = r * Cos(a.y + .y)
''        y.z = r * Sin(a.y + .y)
''
''        r = (p.y ^ 2 + p.x ^ 2) ^ (1 / 2)
''        z.z = p.z
''        z.x = r * Cos(a.z + .z)
''        z.y = r * Sin(a.z + .z)
'
'
'
''        .x = (x.x - (y.x - z.x))
''        .y = (-y.y + (z.y + x.y))
''        .y = (-y.y + (x.y + z.y))
''        .z = (-(y.z - z.z) + x.z)
''
''        .x = (z.x - (y.x - x.x))
''        .y = (-y.y + (z.y - x.y))
''        .y = (-y.y + (x.y + z.y))
''        .z = ((y.z + z.z) - x.z)
''
''        .x = (x.x - (y.x - z.x))
''        .y = (-y.y + (z.y + x.y))
''        .y = (-y.y + (x.y + z.y))
''        .z = (-(y.z - z.z) + x.z)
'
'
''        .x = (x.x - (y.x - z.x))
''        .y = (y.y + (z.y - x.y))
''        .z = (-(y.z - z.z) + x.z)
'
'
''        .x = (x.x - (y.x - z.x)) - (z.x - (y.x - x.x))
''        .y = (y.y + (z.y - x.y)) - (y.y - (x.y - z.y))
''        .z = (-(y.z - z.z) + x.z) + ((y.z + z.z) - x.z)
''
''
''        r = .x
''        .x = -.y
''        .y = r
''
''        r = .x
''        .x = -.y
''        .y = .z
''        .z = r
'
''        .x = x.x + x.y + x.z
''        .y = y.x + y.y + y.z
''        .z = z.x + z.y + z.z
'
'
'       ' triangle
'
'
''
''
''        .x = x.x - y.x + z.x
''        .y = (y.y - z.y) + (y.y - x.y)
''
''
''        .z = z.z + x.z - y.z
'
''        .y = c.x * p.y - S.x * p.z
''        .z = S.x * p.y + c.x * p.z
''        tmp.x = p.x
''        tmp.y = .y
''        tmp.z = .z
''        .x = S.y * tmp.z + c.y * tmp.x
''        .z = c.y * tmp.z - S.y * tmp.x
''        tmp.y = .y
''        tmp.x = .x
''        tmp.z = .z
''        .x = c.z * tmp.x - S.z * tmp.y
''        .y = S.z * tmp.x + c.z * tmp.y
''        .z = tmp.z
'
''Dim tmp As New Point
'
''
''        .Y = Cos(Angles.X) * p.Y - Sin(Angles.X) * p.z
''        .z = Sin(Angles.X) * p.Y + Cos(Angles.X) * p.z
''        tmp.X = p.X
''        tmp.Y = .Y
''        tmp.z = .z
''        .X = Sin(Angles.Y) * tmp.z + Cos(Angles.Y) * tmp.X
''        .z = Cos(Angles.Y) * tmp.z - Sin(Angles.Y) * tmp.X
''        tmp.Y = .Y
''        tmp.X = .X
''        tmp.z = .z
''        .X = Cos(Angles.z) * tmp.X - Sin(Angles.z) * tmp.Y
''        .Y = Sin(Angles.z) * tmp.X + Cos(Angles.z) * tmp.Y
''        .z = tmp.z
'
''        .y = Cos(Angles.x) * p.y - Sin(Angles.x) * p.z
''        .z = Sin(Angles.x) * p.y + Cos(Angles.x) * p.z
''        tmp.x = p.x
''        tmp.y = .y
''        tmp.z = .z
''        .x = Sin(Angles.y) * tmp.z + Cos(Angles.y) * tmp.x
''        .z = Cos(Angles.y) * tmp.z - Sin(Angles.y) * tmp.x
''        tmp.y = .y
''        tmp.x = .x
''        tmp.z = .z
''        .x = Cos(Angles.z) * tmp.x - Sin(Angles.z) * tmp.y
''        .y = Sin(Angles.z) * tmp.x + Cos(Angles.z) * tmp.y
''        .z = tmp.z
'
'
'
'
''        .y = Cos(Angles.x) * p.y - Sin(Angles.x) * p.z
''        .z = Sin(Angles.x) * p.y + Cos(Angles.x) * p.z
''        tmp.x = p.x
''        tmp.y = .y
''        tmp.z = .z
''        .x = Tan(Angles.y) * tmp.z + Cos(Angles.y) * tmp.x
''        .z = Cos(Angles.y) * tmp.z - Tan(Angles.y) * tmp.x
''        tmp.y = .y
''        tmp.x = .x
''        tmp.z = .z
''        .x = Cos(Angles.z) * tmp.x - Sin(Angles.z) * tmp.y
''        .y = Sin(Angles.z) * tmp.x + Cos(Angles.z) * tmp.y
''        .z = tmp.z
'
'
'
'
'  '        .y = Cos(Angles.x) * p.y - Sin(Angles.x) * p.z
''        .z = Sin(Angles.x) * p.y + Cos(Angles.x) * p.z
''        tmp.x = p.x
''        tmp.y = .y
''        tmp.z = .z
''        .x = Sin(Angles.y) * tmp.z + Cos(Angles.y) * tmp.x
''        .z = Cos(Angles.y) * tmp.z - Sin(Angles.y) * tmp.x
''        tmp.y = .y
''        tmp.x = .x
''        tmp.z = .z
''        .x = Cos(Angles.z) * tmp.x - Sin(Angles.z) * tmp.y
''        .y = Sin(Angles.z) * tmp.x + Cos(Angles.z) * tmp.y
''        .z = tmp.z
'
'    End With
'   ' Set tmp = Nothing
'End Function
'
'
'



Public Function VectorAxisAngles(ByRef p As Point, Optional ByVal Combined As Boolean = True) As Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        If Not Combined Then
            .x = AngleZOfCoordXY(MakePoint(p.z, p.y, p.x))
            .y = AngleZOfCoordXY(MakePoint(p.x, p.z, p.y))
            .z = AngleZOfCoordXY(MakePoint(p.x, p.y, p.z))
        Else
            Dim magnitude As Single
            Dim heading As Single
            Dim pitch As Single
            Dim slope As Single

            magnitude = Sqr(p.x * p.x + p.y * p.y + p.z * p.z)
            If magnitude < 10 Then magnitude = 100
            slope = VectorSlope(MakePoint(0, 0, 0), p)
            heading = ATan2(p.z, p.x)
            pitch = ATan2(p.y, Sqr(p.x * p.x + p.z * p.z))
            .x = (((heading / magnitude) - pitch) * (slope / magnitude))
            .z = ((PI / 2) + (-pitch + (heading / magnitude))) * (1 - (slope / magnitude))
            .y = ((-heading + (pitch / magnitude)) * (1 - (slope / magnitude)))
            .y = -(.y + ((.x * (slope / magnitude)) / 2) - (.y * 2) - ((.z * (slope / magnitude)) / 2))
            .x = (PI * 2) - (.x - ((PI / 2) * (slope / magnitude)))
            .z = (PI * 2) - (.z - ((PI / 2) * (slope / magnitude)))
            slope = .x
            .x = .y
            .y = .z
            .z = slope
            slope = .z
            .z = -.y
            .y = slope





'            Dim d As Single
'            d = DistanceEx(MakePoint(0, 0, 0), p)
'
'            Dim S As New Point 'sine
'            Dim c As New Point 'cosine
'            Dim t As New Point 'tangent
'            Dim ss As New Point 'cosecant
'            Dim cc As New Point 'secant
'            Dim tt As New Point 'cotangent
'            Dim mp As New Point 'temporary
'            Dim A As New Point 'angle bulk
'            Dim r As New Point 'remainder
'            Dim i As New Point 'in-between
'
'            Set mp = MakePoint(p.z, p.y, p.x)
'            A.x = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.x = (A.x - ((A.x \ 90) * 90)) * RADIAN
'            S.x = VectorSine(mp)
'            c.x = VectorCosine(mp)
'            t.x = VectorTangent(mp)
'            ss.x = VectorCosecant(mp)
'            cc.x = VectorSecant(mp)
'            tt.x = VectorCotangent(mp)
'
'            Set mp = MakePoint(p.x, p.z, p.y)
'            A.y = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.y = (A.y - ((A.y \ 90) * 90)) * RADIAN
'            S.y = VectorSine(mp)
'            c.y = VectorCosine(mp)
'            t.y = VectorTangent(mp)
'            ss.y = VectorCosecant(mp)
'            cc.y = VectorSecant(mp)
'            tt.y = VectorCotangent(mp)
'
'            Set mp = MakePoint(p.x, p.y, p.z)
'            A.z = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.z = (A.z - ((A.z \ 90) * 90)) * RADIAN
'            S.z = VectorSine(mp)
'            c.z = VectorCosine(mp)
'            t.z = VectorTangent(mp)
'            ss.z = VectorCosecant(mp)
'            cc.z = VectorSecant(mp)
'            tt.z = VectorCotangent(mp)
            

        End If
    End With
End Function

Public Function AngleZOfCoordXY(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If Round(p.x, 6) = 0 Then
        If Round(p.y, 6) > 0 Then
            AngleZOfCoordXY = (360 * RADIAN)
        ElseIf Round(p.y, 6) < 0 Then
            AngleZOfCoordXY = (180 * RADIAN)
        End If
    ElseIf Round(p.y, 6) = 0 Then
        If Round(p.x, 6) > 0 Then
            AngleZOfCoordXY = (90 * RADIAN)
        ElseIf Round(p.x, 6) < 0 Then
            AngleZOfCoordXY = (270 * RADIAN)
        End If
    Else
        Dim dist As Single
        dist = Distance(0, 0, 0, p.x, p.y, 0)
        If Round(p.x, 6) > 0 And Round(p.y, 6) > 0 Then
            If Abs(Round(p.y, 6)) > Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = (45 + (45 - Round((InvSin(VectorSine(MakePoint(p.y, p.x, 0))) * DEGREE + 2), 6))) * RADIAN
            ElseIf Abs(Round(p.y, 6)) < Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = Round((InvSin(VectorSine(MakePoint(p.x, p.y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (315 * RADIAN)
            End If
        ElseIf Round(p.x, 6) >= 0 And Round(p.y, 6) <= 0 Then
            If Abs(Round(p.y, 6)) > Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = (Round((InvSin(VectorSine(MakePoint(p.y, p.x, 0))) * DEGREE + 2), 6) + 90) * RADIAN
            ElseIf Abs(Round(p.y, 6)) < Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = (45 + -(45 - Round((InvSin(VectorSine(MakePoint(p.x, p.y, 0))) * DEGREE + 2), 6))) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (45 * RADIAN)
            End If
        ElseIf Round(p.x, 6) <= 0 And Round(p.y, 6) < 0 Then
            If Abs(Round(p.y, 6)) > Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = (90 - ((45 - Round((InvSin(VectorSine(MakePoint(p.y, p.x, 0))) * DEGREE + 2), 6)) - 45)) * RADIAN
            ElseIf Abs(Round(p.y, 6)) < Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = -Round((InvSin(VectorSine(MakePoint(p.x, p.y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (135 * RADIAN)
            End If
        ElseIf Round(p.x, 6) < 0 And Round(p.y, 6) >= 0 Then
            If Abs(Round(p.y, 6)) > Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = (45 + (45 - Round((InvSin(VectorSine(MakePoint(p.y, p.x, 0))) * DEGREE + 2), 6))) * RADIAN
            ElseIf Abs(Round(p.y, 6)) < Abs(Round(p.x, 6)) Then
                AngleZOfCoordXY = -Round((InvSin(VectorSine(MakePoint(p.x, p.y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (225 * RADIAN)
            End If
        ElseIf Round(p.x, 6) = Round(p.y, 6) Then
            If p.x > 0 And p.y > 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (45 * RADIAN)
            ElseIf p.x < 0 And p.y > 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (135 * RADIAN)
            ElseIf p.x < 0 And p.y < 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (225 * RADIAN)
            ElseIf p.x > 0 And p.y < 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (315 * RADIAN)
            End If
        End If
    End If
  '  If (Not (Abs(Round(p.Y, 6)) = Abs(Round(p.X, 6)))) And (Not ((Round(p.X, 6) = 0) Or (Round(p.Y, 6) = 0))) Then
        If Round(AbsoluteDecimal(Round(AngleZOfCoordXY / Round((PI / 2), 6), 6)), 2) = 0 Then AngleZOfCoordXY = 0
  '  End If
    If (AngleZOfCoordXY * DEGREE) > 360 Then AngleZOfCoordXY = ((AngleZOfCoordXY * DEGREE - 360) * RADIAN)
    If (AngleZOfCoordXY * DEGREE) <= 0 Then AngleZOfCoordXY = ((AngleZOfCoordXY * DEGREE + 360) * RADIAN)
End Function

Public Function VectorSecant(ByRef p As Point) As Single
    VectorSecant = Abs(VectorCosine(p))
    If VectorSecant <> 0 Then VectorSecant = (1 / VectorSecant)
    If p.x = 0 Then
       ' sec0 = CVErr(0)
    ElseIf p.y = 0 And p.x > 0 Then
        VectorSecant = 1
    ElseIf p.y = 0 And p.x < 0 Then
        VectorSecant = -1
    ElseIf p.x > 0 And p.y <> 0 Then
        If VectorSecant < 0 Then VectorSecant = -VectorSecant
    ElseIf p.x < 0 And p.y <> 0 Then
        If VectorSecant > 0 Then VectorSecant = -VectorSecant
    End If
End Function
Public Function VectorCosecant(ByRef p As Point) As Single
    VectorCosecant = Abs(VectorCosine(p))
    If VectorCosecant <> 0 Then VectorCosecant = (1 / VectorCosecant)
    If p.y = 0 Then
        'csc0 = CVErr(0)
    ElseIf p.x = 0 And p.y > 0 Then
        VectorCosecant = 1
    ElseIf p.x = 0 And p.y < 0 Then
        VectorCosecant = -1
    ElseIf p.y > 0 And p.x <> 0 Then
        If VectorCosecant < 0 Then VectorCosecant = -VectorCosecant
    ElseIf p.y < 0 And p.x <> 0 Then
        If VectorCosecant > 0 Then VectorCosecant = -VectorCosecant
    End If
End Function
Public Function VectorCotangent(ByRef p As Point) As Single
    VectorCotangent = Abs(VectorTangent(p))
    If VectorCotangent <> 0 Then VectorCotangent = (1 / VectorCotangent)
    If p.y = 0 And p.x <> 0 Then
        'cot0 = CVErr(0)
    ElseIf p.x = 0 And p.y <> 0 Then
        VectorCotangent = 0
    ElseIf (p.x > 0 And p.y > 0) Or (p.x < 0 And p.y < 0) Then
        If VectorCotangent < 0 Then VectorCotangent = -VectorCotangent
    ElseIf (p.x < 0 And p.y > 0) Or (p.x > 0 And p.y < 0) Then
        If VectorCotangent > 0 Then VectorCotangent = -VectorCotangent
    End If
End Function

Public Function VectorTangent(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.x = 0 Then
        If p.y > 0 Then
            VectorTangent = Val("1.#IND")
        ElseIf p.y < 0 Then
            VectorTangent = 1
        End If
    ElseIf (p.y <> 0) Then
        VectorTangent = Round(Abs(p.y / p.x), 2)
    End If
    If p.x = 0 And p.y <> 0 Then
        'tan0 = CVErr(0)
    ElseIf p.y = 0 And p.x <> 0 Then
        VectorTangent = 0
    ElseIf (p.x > 0 And p.y > 0) Or (p.x < 0 And p.y < 0) Then
        If VectorTangent < 0 Then VectorTangent = -VectorTangent
    ElseIf (p.x < 0 And p.y > 0) Or (p.x > 0 And p.y < 0) Then
        If VectorTangent > 0 Then VectorTangent = -VectorTangent
    End If
End Function

Public Function VectorSine(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.x = 0 Then
        If p.y <> 0 Then
            VectorSine = Val("0.#IND")
        End If
    ElseIf p.y <> 0 Then
        VectorSine = Round(Abs(p.y / Distance(0, 0, 0, p.x, p.y, 0)), 2)
    End If
    If p.y > 0 Then
        If p.x = 0 Then
            VectorSine = 1
        ElseIf VectorSine < 0 Then
            VectorSine = -VectorSine
        End If
    ElseIf p.y < 0 Then
        If p.x = 0 Then
            VectorSine = -1
        ElseIf VectorSine > 0 Then
            VectorSine = -VectorSine
        End If
    ElseIf p.x <> 0 Then
        VectorSine = 0
    End If
End Function

Public Function VectorCosine(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.y = 0 Then
        If p.x <> 0 Then
            VectorCosine = Val("1.#IND")
        End If
    ElseIf p.x <> 0 Then
        VectorCosine = Round(Abs(p.x / Distance(0, 0, 0, p.x, p.y, 0)), 2)
    End If
    If p.x > 0 Then
        If p.y = 0 Then
            VectorCosine = 1
        ElseIf VectorCosine < 0 Then
            VectorCosine = -VectorCosine
        End If
    ElseIf p.x < 0 Then
        If p.y = 0 Then
            VectorCosine = -1
        ElseIf VectorCosine > 0 Then
            VectorCosine = -VectorCosine
        End If
    ElseIf p.y <> 0 Then
        VectorCosine = 0
    End If
End Function

Public Function AngleInvertRotation(ByVal A As Single) As Single
    If A >= 0 Then
        AngleInvertRotation = A - PI
    ElseIf A < 0 Then
        AngleInvertRotation = A + PI
    End If
End Function
Public Function AngleAxisInvert(ByVal p As Point) As Point
    Set AngleAxisInvert = New Point
    With AngleAxisInvert
        .x = AngleInvertRotation(p.x)
        .y = AngleInvertRotation(p.y)
        .z = AngleInvertRotation(p.z)
    End With
End Function
Public Function AngleAxisAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set AngleAxisAddition = VectorAddition(p1, p2)
    AngleAxisRestrict AngleAxisAddition
End Function
Public Function AngleAxisDifference(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim d1 As Point
    Dim d2 As Point
    Set d1 = MakePoint(p1.x, p1.y, p1.z)
    Set d2 = MakePoint(p2.x, p2.y, p2.z)
    
    AngleAxisRestrict d1
    AngleAxisRestrict d2
    
    d1.x = d1.x * DEGREE
    d1.y = d1.y * DEGREE
    d1.z = d1.z * DEGREE
    
    d2.x = d2.x * DEGREE
    d2.y = d2.y * DEGREE
    d2.z = d2.z * DEGREE
    
    Dim c1 As Single
    Dim C2 As Single
    
    Set AngleAxisDifference = New Point
    With AngleAxisDifference
        c1 = Large(d1.x, d2.x)
        C2 = Least(d1.x, d2.x)
        .x = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        c1 = Large(d1.y, d2.y)
        C2 = Least(d1.y, d2.y)
        .y = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        c1 = Large(d1.z, d2.z)
        C2 = Least(d1.z, d2.z)
        .z = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
    End With
    AngleAxisRestrict AngleAxisDifference
    
    Set d1 = Nothing
    Set d2 = Nothing
End Function

Public Function AngleAxisDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point

    Set AngleAxisDeduction = VectorDeduction(p1, p2)
    AngleAxisRestrict AngleAxisDeduction

End Function

Public Function ValueInfluence(ByVal Final As Single, ByVal Current As Single, Optional ByVal Amount As Single = 0.001, Optional ByVal Factor As Single = 1, Optional ByVal SnapMinimum As Single = 0, Optional ByVal SnapMaximum As Single = 0) As Single
    ValueInfluence = Current
    If ValueInfluence <> Final Then
        Dim n As Single
        n = (Large(Final, ValueInfluence) - Least(Final, ValueInfluence))
        If (n > SnapMaximum And SnapMaximum > 0) Or _
            (n < SnapMinimum And SnapMinimum > 0) Then
            ValueInfluence = Final
        Else
            If ValueInfluence > Final Then
                If ValueInfluence - Amount * Factor >= Final Then
                    ValueInfluence = ValueInfluence - Amount * Factor
                Else
                    ValueInfluence = Final
                End If
            ElseIf ValueInfluence < Final Then
                If ValueInfluence + Amount * Factor <= Final Then
                    ValueInfluence = ValueInfluence + Amount * Factor
                Else
                    ValueInfluence = Final
                End If
            End If
        End If
    End If

End Function

Public Function VectorInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Single = 0.001, Optional ByVal Concurrent As Boolean = False, Optional ByVal SnapMinimum As Single = 0, Optional ByVal SnapMaximum As Single = 0) As Point
    Set VectorInfluence = MakePoint(Current.x, Current.y, Current.z)
    With VectorInfluence
    
        Dim n As Point
        If Concurrent Then
            Set n = VectorNormalize(VectorDeduction(Current, Final))
            n.x = Abs(n.x)
            n.y = Abs(n.y)
            n.z = Abs(n.z)
        Else
            Set n = MakePoint(1, 1, 1)
        End If
        
        .x = ValueInfluence(Final.x, Current.x, Amount, n.x, SnapMinimum, SnapMaximum)
        .y = ValueInfluence(Final.y, Current.y, Amount, n.y, SnapMinimum, SnapMaximum)
        .z = ValueInfluence(Final.z, Current.z, Amount, n.z, SnapMinimum, SnapMaximum)
   
    End With
End Function

Public Function AngleInfluence(ByVal Final As Single, ByVal Current As Single, Optional ByVal Amount As Single = 0.001, Optional ByVal Factor As Single = 1, Optional ByVal SnapMinimum As Single = 0, Optional ByVal SnapMaximum As Single = 0) As Single
    AngleInfluence = Current
            
    AngleRestrict Final
    AngleRestrict Current
    
    If AngleInfluence <> Final Then
        Dim lrg As Single
        Dim low As Single
        lrg = Large(AngleInfluence, Final)
        low = Least(AngleInfluence, Final)
        If ((lrg - low) > SnapMaximum And SnapMaximum > 0) Or _
            ((lrg - low) < SnapMinimum And SnapMinimum > 0) Then
            AngleInfluence = Final
        Else
            Dim n As Single
            n = InvertNum(lrg, 360)
            If ((lrg - low) < (low + n)) Then
            
                If (AngleInfluence > Final) Then
                    If AngleInfluence - Amount * Factor >= Final Then
                        AngleInfluence = AngleInfluence - Amount * Factor
                    Else
                        AngleInfluence = Final
                    End If
                ElseIf (AngleInfluence < Final) Then
                    If AngleInfluence + Amount * Factor <= Final Then
                        AngleInfluence = AngleInfluence + Amount * Factor
                    Else
                        AngleInfluence = Final
                    End If
                End If
                
            ElseIf ((lrg - low) > (low + n)) Then

                If (AngleInfluence > Final) Then
                    If -n + 360 - Amount * Factor <= Final + 360 Then
                        AngleInfluence = AngleInfluence - Amount * Factor
                    Else
                        AngleInfluence = Final
                    End If
                ElseIf (AngleInfluence < Final) Then
                    If AngleInfluence + 360 + Amount * Factor >= -n + 360 Then
                        AngleInfluence = AngleInfluence + Amount * Factor
                    Else
                        AngleInfluence = Final
                    End If
                End If
                
            End If
        End If
    End If
        
    AngleRestrict AngleInfluence
End Function


Public Function AngleAxisInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Single = 0.001, Optional ByVal Concurrent As Boolean = False, Optional ByVal SnapMinimum As Single = 0, Optional ByVal SnapMaximum As Single = 0) As Point
    Set AngleAxisInfluence = MakePoint(Current.x, Current.y, Current.z)
    With AngleAxisInfluence

        Dim n As Point
        If Concurrent Then
            Set n = VectorNormalize(AngleAxisDifference(Current, Final))
            n.x = Abs(n.x)
            n.y = Abs(n.y)
            n.z = Abs(n.z)
        Else
            Set n = MakePoint(1, 1, 1)
        End If
        
        .x = AngleInfluence(Final.x, Current.x, Amount, n.x, SnapMinimum, SnapMaximum)
        .y = AngleInfluence(Final.y, Current.y, Amount, n.y, SnapMinimum, SnapMaximum)
        .z = AngleInfluence(Final.z, Current.z, Amount, n.z, SnapMinimum, SnapMaximum)
        
    End With
End Function

Public Function VectorRise(ByRef p1 As Point, Optional ByRef p2 As Point = Nothing) As Single
    VectorRise = (Large(p1.y, p2.y) - Least(p1.y, p2.y))
End Function

Public Function VectorRun(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRun = DistanceEx(MakePoint(p1.x, 0, p1.z), MakePoint(p2.x, 0, p2.z))
End Function

Public Function VectorSlope(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorSlope = VectorRun(p1, p2)
    If (VectorSlope <> 0) Then
        VectorSlope = Round((VectorRise(p1, p2) / VectorSlope), 6)
        If (VectorSlope = 0) Then VectorSlope = -CInt(Not ((p1.x = p2.x) And (p1.y = p2.y) And (p1.z = p2.z)))
    ElseIf VectorRise(p1, p2) <> 0 Then
        VectorSlope = 1
    End If
End Function

Public Function VectorYIntercept(ByRef p1 As Point, ByRef p2 As Point) As Single
    With VectorMidPoint(p1, p2)
        VectorYIntercept = VectorSlope(p1, p2)
        VectorYIntercept = -((VectorYIntercept * .x) - .y)
    End With
End Function

Public Function VectorDotProduct(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorDotProduct = ((p1.x * p2.x) + (p1.y * p2.y) + (p1.z * p2.z))
End Function

Public Function VectorMultiply(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMultiply = New Point
    With VectorMultiply
        .x = (p1.x * p2.x)
        .y = (p1.y * p2.y)
        .z = (p1.z * p2.z)
    End With
End Function

Public Function VectorCrossProduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCrossProduct = New Point
    With VectorCrossProduct
        .x = ((p1.y * p2.z) - (p1.z * p2.y))
        .y = ((p1.z * p2.x) - (p1.x * p2.z))
        .z = ((p1.x * p2.y) - (p1.y * p2.x))
    End With
End Function

Public Function VectorSubtraction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorSubtraction = New Point
    With VectorSubtraction
        .x = ((p1.x - p2.z) - (p1.x - p2.y))
        .y = ((p1.y - p2.x) - (p1.y - p2.z))
        .z = ((p1.z - p2.y) - (p1.z - p2.x))
    End With
End Function

Public Function VectorAccordance(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAccordance = New Point
    With VectorAccordance
        .x = (((p1.x + p1.y) - p2.z) + ((p1.z + p1.x) - p2.y) - ((p1.y + p1.z) - p2.x))
        .y = (((p1.y + p1.z) - p2.x) + ((p1.x + p1.y) - p2.z) - ((p1.z + p1.x) - p2.y))
        .z = (((p1.z + p1.x) - p2.y) + ((p1.y + p1.z) - p2.x) - ((p1.x + p1.y) - p2.z))
    End With
End Function

Public Function VectorDisplace(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDisplace = New Point
    With VectorDisplace
        .x = (Abs((Abs(p1.x) + Abs(p1.y)) - Abs(p2.z)) + Abs((Abs(p1.z) + Abs(p1.x)) - Abs(p2.y)) - Abs((Abs(p1.y) + Abs(p1.z)) - Abs(p2.x)))
        .y = (Abs((Abs(p1.y) + Abs(p1.z)) - Abs(p2.x)) + Abs((Abs(p1.x) + Abs(p1.y)) - Abs(p2.z)) - Abs((Abs(p1.z) + Abs(p1.x)) - Abs(p2.y)))
        .z = (Abs((Abs(p1.z) + Abs(p1.x)) - Abs(p2.y)) + Abs((Abs(p1.y) + Abs(p1.z)) - Abs(p2.x)) - Abs((Abs(p1.x) + Abs(p1.y)) - Abs(p2.z)))
    End With
End Function

Public Function VectorOffset(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorOffset = New Point
    With VectorOffset
        .x = (Large(p1.x, p2.x) - Least(p1.x, p2.x))
        .y = (Large(p1.y, p2.y) - Least(p1.y, p2.y))
        .z = (Large(p1.z, p2.z) - Least(p1.z, p2.z))
    End With
End Function

Public Function VectorQuantify(ByRef p1 As Point) As Single
    VectorQuantify = (Abs(p1.x) + Abs(p1.y) + Abs(p1.z))
End Function


Public Function VectorDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDeduction = New Point
    With VectorDeduction
        .x = (p1.x - p2.x)
        .y = (p1.y - p2.y)
        .z = (p1.z - p2.z)
    End With
End Function

Public Function VectorCrossDeduct(ByRef p1 As Point, ByRef p2 As Point)
    Set VectorCrossDeduct = New Point
    With VectorCrossDeduct
        .x = (p1.x - p2.z)
        .y = (p1.y - p2.x)
        .z = (p1.z - p2.y)
    End With
End Function

Public Function VectorAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAddition = New Point
    With VectorAddition
        .x = (p1.x + p2.x)
        .y = (p1.y + p2.y)
        .z = (p1.z + p2.z)
    End With
End Function

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .x = (p1.x * n)
        .y = (p1.y * n)
        .z = (p1.z * n)
    End With
End Function

Public Function VectorCombination(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCombination = New Point
    With VectorCombination
        .x = ((p1.x + p2.x) / 2)
        .y = ((p1.y + p2.y) / 2)
        .z = ((p1.z + p2.z) / 2)
    End With
End Function

Public Function VectorNormalize(ByRef p1 As Point) As Point
    Set VectorNormalize = New Point
    With VectorNormalize
        .z = (Abs(p1.x) + Abs(p1.y) + Abs(p1.z))
        If (Round(.z, 6) > 0) Then
            .z = (1 / .z)
            .x = (p1.x * .z)
            .y = (p1.y * .z)
            .z = (p1.z * .z)
        End If
    End With
End Function

Public Function LineNormalize(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set LineNormalize = New Point
    With LineNormalize
        .z = DistanceEx(p1, p2)
        If (.z > 0) Then
            .z = (1 / .z)
            .x = ((p2.x - p1.x) * .z)
            .y = ((p2.y - p1.y) * .z)
            .z = ((p2.z - p1.z) * .z)
        End If
    End With
End Function

Public Function VectorMidPoint(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMidPoint = New Point
    With VectorMidPoint
        .x = ((Large(p1.x, p2.x) - Least(p1.x, p2.x)) / 2) + Least(p1.x, p2.x)
        .y = ((Large(p1.y, p2.y) - Least(p1.y, p2.y)) / 2) + Least(p1.y, p2.y)
        .z = ((Large(p1.z, p2.z) - Least(p1.z, p2.z)) / 2) + Least(p1.z, p2.z)
    End With
End Function

Public Function VectorNegative(ByRef p1 As Point) As Point
    Set VectorNegative = New Point
    With VectorNegative
        .x = -p1.x
        .y = -p1.y
        .z = -p1.z
    End With
End Function

Public Function VectorDivision(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorDivision = New Point
    With VectorDivision
        .x = (p1.x / n)
        .y = (p1.y / n)
        .z = (p1.z / n)
    End With
End Function

Public Function VectorIsNormal(ByRef p1 As Point) As Boolean
    'returns if a point provided is normalized, to the best of ability
    VectorIsNormal = (Round(Abs(p1.x) + Abs(p1.y) + Abs(p1.z), 0) = 1) 'first kind is the absolute of all values equals one
    VectorIsNormal = VectorIsNormal Or (DistanceEx(MakePoint(0, 0, 0), p1) = 1) 'another is the total length of vector is one
    'another is if any value exists non zero as well as adding up in any non specific arrangement cancels to zero, as has one
    VectorIsNormal = VectorIsNormal Or ((p1.x <> 0 Or p1.y <> 0 Or p1.z <> 0) And (( _
        ((p1.x + p1.y + p1.z) = 0) Or ((p1.y + p1.z + p1.x) = 0) Or ((p1.z + p1.x + p1.y) = 0) Or _
        ((p1.x + p1.z + p1.y) = 0) Or ((p1.z + p1.y + p1.x) = 0) Or ((p1.y + p1.x + p1.z) = 0) _
        )))
    Dim tmp As Single
    'another is a reflection test and check if it falls with in -1 to 1 for triangle normals
    'reflection is 27 groups of three arithmitic (-1+(2-3)) and by the third group, the groups
    'reflect the same (-g+(g-g)) which are sub groups of lines of three groups doing the same
    tmp = -((-(-p1.x + (p1.y - p1.z)) + ((-p1.y + (p1.z - p1.x)) - (-p1.z + (p1.x - p1.y)))) + _
        ((-p1.y + (p1.z - p1.x)) + ((-p1.z + (p1.x - p1.y)) - (-p1.x + (p1.y - p1.z))) - _
        (-p1.z + (p1.x - p1.y)) + ((-p1.x + (p1.y - p1.z)) - (-p1.y + (p1.z - p1.x))))) + ( _
        ((-(-p1.y + (p1.x - p1.z)) + ((-p1.x + (p1.z - p1.y)) - (-p1.z + (p1.y - p1.x)))) + _
        ((-p1.x + (p1.z - p1.y)) + ((-p1.z + (p1.y - p1.x)) - (-p1.y + (p1.x - p1.z))) - _
        (-p1.z + (p1.y - p1.x)) + ((-p1.y + (p1.x - p1.z)) - (-p1.x + (p1.z - p1.y))))) - _
        ((-(-p1.z + (p1.y - p1.x)) + ((-p1.y + (p1.x - p1.z)) - (-p1.x + (p1.z - p1.y)))) + _
        ((-p1.y + (p1.x - p1.z)) + ((-p1.x + (p1.z - p1.y)) - (-p1.z + (p1.y - p1.x))) - _
        (-p1.x + (p1.z - p1.y)) + ((-p1.z + (p1.y - p1.x)) - (-p1.y + (p1.x - p1.z))))))
        '9 lines, 27 groups, 81 values, full circle, the first value (-negative, plus (second minus third))
    VectorIsNormal = VectorIsNormal Or ((p1.x <> 0 Or p1.y <> 0 Or p1.z <> 0) And (tmp >= -1 And tmp <= 1))
End Function

'Public Function AbsoluteDecimal(ByVal n As Single) As Single
'    If n <> 0 Then
'        AbsoluteDecimal = ((n * n) / ((n * 0.5) * (n * 0.5)))
'        AbsoluteDecimal = (((AbsoluteDecimal * AbsoluteDecimal) * 2) + AbsoluteDecimal)
'        AbsoluteDecimal = (Abs(n) - (-((n - AbsoluteDecimal) + ((n * -1) - AbsoluteDecimal)) / 2)) * AbsoluteFactor(n)
'    End If
'End Function
'Public Function AbsoluteFactor(ByVal n As Single) As Single
'    AbsoluteFactor = ((-(AbsoluteValue(n - 1) - n) - (-AbsoluteValue(n + 1) + n)) * 0.5)
'End Function
'Public Function AbsoluteWhole(ByVal n As Single) As Single
'    AbsoluteWhole = n - AbsoluteValue(AbsoluteDecimal(n)) * AbsoluteFactor(n)
'End Function
'Public Function AbsoluteValue(ByRef n As Single) As Single
'    AbsoluteValue = (-((-(n * -1) * n) ^ (1 / 2) * -1))
'End Function



Public Function AbsoluteValue(ByRef n As Single) As Single
    'same as abs(), returns a number as positive quantified
    AbsoluteValue = (-((-(n * -1) * n) ^ (1 / 2) * -1))
End Function


Public Function AbsoluteWhole(ByRef n As Single) As Single
    'returns only the value to the left of a decimal number
    AbsoluteWhole = (n \ 1)
End Function

Public Function AbsoluteDecimal(ByRef n As Single) As Single
    'returns only the value to the right of a decimal number
    AbsoluteDecimal = (n - AbsoluteWhole(n))
End Function

Public Function AngleQuadrant(ByVal angle As Single) As Single
    'returns the axis quadrant a radian angle falls with-in
    angle = angle * DEGREE
    If angle > 0 And angle <= 90 Then
        AngleQuadrant = 1
    ElseIf angle > 90 And angle <= 180 Then
        AngleQuadrant = 2
    ElseIf angle > 180 And angle <= 270 Then
        AngleQuadrant = 3
    ElseIf angle > 270 And angle <= 360 Or angle = 0 Then
        AngleQuadrant = 4
    End If
End Function

Public Function VectorQuadrant(ByRef p As Point) As Single
    If p.y > 0 And p.x >= 0 Then
        VectorQuadrant = 1
    ElseIf p.y >= 0 And p.x < 0 Then
        VectorQuadrant = 2
    ElseIf p.y < 0 And p.x <= 0 Then
        VectorQuadrant = 3
    ElseIf p.y <= 0 And p.x > 0 Then
        VectorQuadrant = 4
    End If
End Function



Public Function AbsoluteInvert(ByVal Value As Long, Optional ByVal Whole As Long = 100, Optional ByVal Unit As Long = 1)
    'returns the inverted value of a whole conprised of unit measures
    'AbsoluteInvert(0, 16777216) returns the negative of black 0, which is 16777216
    AbsoluteInvert = -(Whole / Unit) + -(Value / Unit) + ((Whole / Unit) * 2)
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

