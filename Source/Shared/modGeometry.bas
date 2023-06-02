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


Public Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.z = z
End Function

Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.z = z
End Function

Public Function MakeCoord(ByVal X As Single, ByVal Y As Single) As Coord
    Set MakeCoord = New Coord
    MakeCoord.X = X
    MakeCoord.Y = Y
End Function

Public Function ToCoord(ByRef Vector As D3DVECTOR) As Coord
    Set ToCoord = New Coord
    ToCoord.X = Vector.X
    ToCoord.Y = Vector.Y
End Function

Public Function ToVector(ByRef Point As Point) As D3DVECTOR
    ToVector.X = Point.X
    ToVector.Y = Point.Y
    ToVector.z = Point.z
End Function

Public Function ToPoint(ByRef Vector As D3DVECTOR) As Point
    Set ToPoint = New Point
    ToPoint.X = Vector.X
    ToPoint.Y = Vector.Y
    ToPoint.z = Vector.z
End Function

Public Function ToPlane(ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Range
        
    Dim pNormal As Point
    Set pNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(V2, V1), VectorDeduction(V3, V1)))
        
    Set ToPlane = New Range
    With ToPlane
        .W = VectorDotProduct(pNormal, V1) * -1
        .X = pNormal.X
        .Y = pNormal.Y
        .z = pNormal.z
    End With
End Function

Public Function ToVec4(ByRef Plane As Range) As D3DVECTOR4
    ToVec4.X = Plane.X
    ToVec4.Y = Plane.Y
    ToVec4.z = Plane.z
    ToVec4.W = Plane.W
End Function

Public Function DistanceToPlane(ByRef p As Point, ByRef r As Range) As Single
    DistanceToPlane = (r.X * p.X + r.Y * p.Y + r.z * p.z + r.W) / Sqr(r.X * r.X + r.Y * r.Y + r.z * r.z)
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = ((((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = ((((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.z - p2.z) ^ 2)) ^ (1 / 2))
End Function

Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal n As Single) As Point
    Dim dist As Single
    dist = DistanceEx(p1, p2)
    Set DistanceSet = New Point
    With DistanceSet
        If Not (dist = n) Then
            If ((dist > 0) And (n > 0)) Then
                .X = Large(p1.X, p2.X) - Least(p1.X, p2.X)
                .Y = Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)
                .z = Large(p1.z, p2.z) - Least(p1.z, p2.z)
                .X = (Least(p1.X, p2.X) + (n * (.X / dist)))
                .Y = (Least(p1.Y, p2.Y) + (n * (.Y / dist)))
                .z = (Least(p1.z, p2.z) + (n * (.z / dist)))
            ElseIf (n = 0) Then
                .X = p1.X
                .Y = p1.Y
                .z = p1.z
            ElseIf (dist = 0) Then
                .X = p2.X
                .Y = p2.Y
                .z = p2.z + IIf(p2.z > p1.z, n, -n)
            End If
        End If
    End With
End Function

Public Function PointOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    Dim r As Range
    Set r = ToPlane(v0, V1, V2)
    PointOnPlane = (r.X * (p.X - v0.X)) + (r.Y * (p.Y - v0.Y)) + (r.z * (p.z - v0.z)) = 0
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
        
        .X = (d * n.X)
        .Y = (d * n.Y)
        .z = (d * n.z)
        
    End With
    Set PointOnPlaneNearestPoint = VectorAddition(p, PointOnPlaneNearestPoint)
    
End Function

Public Function LineIntersectPlane(ByRef Plane As Range, PStart As Point, vDir As Point, ByRef VIntersectOut As Point) As Boolean
    Dim q As New Range     'Start Point
    Dim v As New Range       'Vector Direction

    Dim planeQdot As Single 'Dot products
    Dim planeVdot As Single
    
    Dim t As Single         'Part of the equation for a ray P(t) = Q + tV

    
    q.X = PStart.X          'Q is a point and therefore it's W value is 1
    q.Y = PStart.Y
    q.z = PStart.z
    q.W = 1
    
    v.X = vDir.X            'V is a vector and therefore it's W value is zero
    v.Y = vDir.Y
    v.z = vDir.z
    v.W = 0
    
  '  ((Plane.X * V.X) + (Plane.Y * V.Y) + (Plane.z * V.z) + (Plane.w * V.w))
    
    planeVdot = ((Plane.X * v.X) + (Plane.Y * v.Y) + (Plane.z * v.z) + (Plane.W * v.W)) 'D3DXVec4Dot(Plane, V)
    planeQdot = ((Plane.X * q.X) + (Plane.Y * q.Y) + (Plane.z * q.z) + (Plane.W * q.W)) 'D3DXVec4Dot(Plane, Q)
            
    'If the dotproduct of plane and V = 0 then there is no intersection
    If planeVdot <> 0 Then
        t = Round((planeQdot / planeVdot) * -1, 5)
        
        If VIntersectOut Is Nothing Then Set VIntersectOut = New Point
        
        'This is where the line intersects the plane
        VIntersectOut.X = Round(q.X + (t * v.X), 5)
        VIntersectOut.Y = Round(q.Y + (t * v.Y), 5)
        VIntersectOut.z = Round(q.z + (t * v.z), 5)

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

Public Function PointNormalize(ByRef v As Point) As Point
    Set PointNormalize = New Point
    With PointNormalize
        .z = (v.X ^ 2 + v.Y ^ 2 + v.z ^ 2) ^ (1 / 2)
        If (.z = 0) Then .z = 1
        .X = (v.X / .z)
        .Y = (v.Y / .z)
        .z = (v.z / .z)
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

Public Function SphereToCubeRoot(ByVal Diameter As Single) As Single
    SphereToCubeRoot = (((Diameter ^ 2) / 2) ^ (1 / 2))
    'opposite of CubeToSphereDiameter() if edge1, edge2, and edge3 are the same value,
    'true cube. for instance ((Diameter^2)^(1/3)) equals one eight of any of all three edges
    'surface area of a sphere is still only two dimensions, so we skip ahead cutting down
End Function

Public Function SquareCenter(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Point
    Set SquareCenter = New Point
    With SquareCenter
        'center by adding onto the lowest value of axis with the the middle of the absolute difference of each of axis
        .X = (Least(v0.X, V1.X, V2.X, V3.X) + ((Large(v0.X, V1.X, V2.X, V3.X) - Least(v0.X, V1.X, V2.X, V3.X)) / 2))
        .Y = (Least(v0.Y, V1.Y, V2.Y, V3.Y) + ((Large(v0.Y, V1.Y, V2.Y, V3.Y) - Least(v0.Y, V1.Y, V2.Y, V3.Y)) / 2))
        .z = (Least(v0.z, V1.z, V2.z, V3.z) + ((Large(v0.z, V1.z, V2.z, V3.z) - Least(v0.z, V1.z, V2.z, V3.z)) / 2))
    End With
End Function

Public Function CirclePermeter(ByVal Radii As Single) As Single
    CirclePermeter = ((Radii * 2) * PI)
End Function

Public Function CubeToSphereDiameter(ByVal edge1 As Single, Optional ByVal edge2 As Single = 0, Optional ByVal edge3 As Single = 0) As Single
    'opposite of SphereToCubeRoot(), input is three edges or length, width and height
    'each form of square dimensions among the neighbor is used with a self squared to
    'find all the possible square dimensions making two groups of three, and add together
    'the averages of those groups then root it by two, returning a diameter by the volume
    If edge2 = 0 And edge3 = 0 Then
        CubeToSphereDiameter = (((((edge1 ^ 2) + (edge1 ^ 2) + (edge1 ^ 2)) / 3)) + _
                ((((edge1 ^ 2) + (edge1 ^ 2) + (edge1 ^ 2)) / 3))) ^ (1 / 2)
    Else
        CubeToSphereDiameter = (((((edge1 * edge2) + (edge2 * edge3) + (edge3 * edge1)) / 3)) + _
                ((((edge1 * edge1) + (edge2 * edge2) + (edge3 * edge3)) / 3))) ^ (1 / 2)
    End If
End Function
Public Function CubePerimeter(ByVal edge1 As Single, Optional ByVal edge2 As Single = 0, Optional ByVal edge3 As Single = 0) As Single
    If edge2 = 0 And edge3 = 0 Then
        CubePerimeter = (edge1 * 12)
    Else
        CubePerimeter = (edge1 * 4) + (edge2 * 4) + (edge3 * 4)
    End If
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
        .z = ((p1.z + p2.z + p3.z) / 3)
    End With
End Function

Public Function TriangleOffset(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleOffset = New Point
    With TriangleOffset
        .X = (Large(p1.X, p2.X, p3.X) - Least(p1.X, p2.X, p3.X))
        .Y = (Large(p1.Y, p2.Y, p3.Y) - Least(p1.Y, p2.Y, p3.Y))
        .z = (Large(p1.z, p2.z, p3.z) - Least(p1.z, p2.z, p3.z))
    End With
End Function

Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleAxii = New Point
    With TriangleAxii
        Dim o As Point
        Set o = TriangleOffset(p1, p2, p3)
        .X = (Least(p1.X, p2.X, p3.X) + (o.X / 2))
        .Y = (Least(p1.Y, p2.Y, p3.Y) + (o.Y / 2))
        .z = (Least(p1.z, p2.z, p3.z) + (o.z / 2))
    End With
End Function

Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleNormal = New Point
    Dim o As Point
    Dim d As Single
    With TriangleNormal
        Set o = TriangleDisplace(v0, V1, V2)
        d = (o.X + o.Y + o.z)
        If (d > 0) Then
            .z = (((o.X + o.Y) - o.z) / d)
            .X = (((o.Y + o.z) - o.X) / d)
            .Y = (((o.z + o.X) - o.Y) / d)
        End If
    End With
End Function

Public Function TriangleAccordance(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleAccordance = New Point
    With TriangleAccordance
        .X = (((v0.X + V1.X) - V2.X) + ((V1.X + V2.X) - v0.X) - ((V2.X + v0.X) - V1.X))
        .Y = (((v0.Y + V1.Y) - V2.Y) + ((V1.Y + V2.Y) - v0.Y) - ((V2.Y + v0.Y) - V1.Y))
        .z = (((v0.z + V1.z) - V2.z) + ((V1.z + V2.z) - v0.z) - ((V2.z + v0.z) - V1.z))
    End With
End Function

Public Function TriangleDisplace(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleDisplace = New Point
    With TriangleDisplace
        .X = (Abs((Abs(v0.X) + Abs(V1.X)) - Abs(V2.X)) + Abs((Abs(V1.X) + Abs(V2.X)) - Abs(v0.X)) - Abs((Abs(V2.X) + Abs(v0.X)) - Abs(V1.X)))
        .Y = (Abs((Abs(v0.Y) + Abs(V1.Y)) - Abs(V2.Y)) + Abs((Abs(V1.Y) + Abs(V2.Y)) - Abs(v0.Y)) - Abs((Abs(V2.Y) + Abs(v0.Y)) - Abs(V1.Y)))
        .z = (Abs((Abs(v0.z) + Abs(V1.z)) - Abs(V2.z)) + Abs((Abs(V1.z) + Abs(V2.z)) - Abs(v0.z)) - Abs((Abs(V2.z) + Abs(v0.z)) - Abs(V1.z)))
    End With
End Function

Public Function VectorBalance(ByRef loZero As Point, ByRef hiWhole As Point, ByVal folcrumPercent As Single) As Point
    Set VectorBalance = New Point
    With VectorBalance
        .X = (loZero.X + ((hiWhole.X - loZero.X) * folcrumPercent))
        .Y = (loZero.Y + ((hiWhole.Y - loZero.Y) * folcrumPercent))
        .z = (loZero.z + ((hiWhole.z - loZero.z) * folcrumPercent))
    End With
End Function

Public Function TriangleFolcrum(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Point
    Set TriangleFolcrum = New Point
    With TriangleFolcrum
        If (Not p3 Is Nothing) Then
            .X = (p3.X ^ 2)
            .Y = (p3.Y ^ 2)
            .z = (p3.z ^ 2)
        End If
        .X = (.X + (p1.X ^ 2) + (p2.X ^ 2)) ^ (1 / 2)
        .Y = (.Y + (p1.Y ^ 2) + (p2.Y ^ 2)) ^ (1 / 2)
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


Public Sub AngleAxisRestrict(ByRef AxisAngles As Point)
    AxisAngles.X = AngleRestrict(AxisAngles.X)
    AxisAngles.Y = AngleRestrict(AxisAngles.Y)
    AxisAngles.z = AngleRestrict(AxisAngles.z)
End Sub
Public Function AngleRestrict(ByRef AxisAngle As Single) As Single
    AxisAngle = Round(AxisAngle, 6)
    AngleRestrict = 0
    Do While AxisAngle < -PI * 2 And AngleRestrict <= 4
        AxisAngle = AxisAngle + (PI * 2)
        AngleRestrict = AngleRestrict + 1
    Loop
    AngleRestrict = 0
    Do While AxisAngle > PI * 2 And AngleRestrict <= 4
        AxisAngle = AxisAngle - (PI * 2)
        AngleRestrict = AngleRestrict + 1
    Loop
    AngleRestrict = AxisAngle
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

Public Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
    If Abs(adj) < 0.000001 Then
        ATan2 = PI / 2
    Else
        ATan2 = Abs(Atn(opp / adj))
    End If
    If adj < 0 Then ATan2 = PI - ATan2
    If opp < 0 Then ATan2 = -ATan2
End Function

Public Function InvSin(number As Single) As Single
    InvSin = -number * number + 1
    If InvSin > 0 Then
        InvSin = Sqr(InvSin)
        If InvSin <> 0 Then InvSin = Atn(number / InvSin)
    Else
        InvSin = 0
    End If
End Function

Public Function VectorRotateAxis(ByRef PointToRotate As Point, ByRef RadianAngles As Point) As Point

'    Dim S As Single
'    Dim tmp As New Point
'    Set VectorRotateAxis = New Point
'    With VectorRotateAxis
'        .Y = Cos(RadianAngles.X) * PointToRotate.Y - Sin(RadianAngles.X) * PointToRotate.z
'        .z = Sin(RadianAngles.X) * PointToRotate.Y + Cos(RadianAngles.X) * PointToRotate.z
'        tmp.X = PointToRotate.X
'        tmp.Y = .Y
'        tmp.z = .z
'        .X = Sin(RadianAngles.Y) * tmp.z + Cos(RadianAngles.Y) * tmp.X
'        .z = Cos(RadianAngles.Y) * tmp.z - Sin(RadianAngles.Y) * tmp.X
'        tmp.X = .X
'        tmp.z = .z
'        .X = Cos(RadianAngles.z) * tmp.X - Sin(RadianAngles.z) * tmp.Y
'        .Y = Sin(RadianAngles.z) * tmp.X + Cos(RadianAngles.z) * tmp.Y
'        .z = tmp.z
'    End With
'    Set tmp = Nothing
    

   ' Dim S As Single
    Dim tmp As New Point
    Set VectorRotateAxis = New Point
    With VectorRotateAxis
        .Y = Cos(RadianAngles.X) * PointToRotate.Y - Sin(RadianAngles.X) * PointToRotate.z
        .z = Sin(RadianAngles.X) * PointToRotate.Y + Cos(RadianAngles.X) * PointToRotate.z
        tmp.X = PointToRotate.X
        tmp.Y = .Y
        tmp.z = .z
        .X = Sin(RadianAngles.Y) * tmp.z + Cos(RadianAngles.Y) * tmp.X
        .z = Cos(RadianAngles.Y) * tmp.z - Sin(RadianAngles.Y) * tmp.X
        tmp.X = .X
        tmp.z = .z
        .X = Cos(RadianAngles.z) * tmp.X - Sin(RadianAngles.z) * tmp.Y
        .Y = Sin(RadianAngles.z) * tmp.X + Cos(RadianAngles.z) * tmp.Y
        .z = tmp.z



'    Debug.Print .X; .Y; .z

'            Dim S As New Point 'sine
'            Dim c As New Point 'cosine
'            Dim t As New Point 'tangent
'            Dim ss As New Point 'cosecant
'            Dim cc As New Point 'secant
'            Dim tt As New Point 'cotangent
'            Dim mp As New Point 'temporary
'            Dim A As New Point 'angle
'            Dim r As New Point 'co-angle
'
'            Set mp = MakePoint(PointToRotate.z, PointToRotate.Y, PointToRotate.X)
'            A.X = AngleZOfCoordXY(mp)
'            r.X = AngleRestrict(AngleRestrict(A.X) + AngleRestrict(RadianAngles.X)) * DEGREE
'            A.X = AngleQuadrant(A.X * DEGREE)
'            S.X = VectorSine(mp)
'            c.X = VectorCosine(mp)
'            t.X = VectorTangent(mp)
'            ss.X = VectorCosecant(mp)
'            cc.X = VectorSecant(mp)
'            tt.X = VectorCotangent(mp)
'
'            Set mp = MakePoint(PointToRotate.X, PointToRotate.z, PointToRotate.Y)
'            A.Y = AngleZOfCoordXY(mp)
'            r.Y = AngleRestrict(AngleRestrict(A.Y) + AngleRestrict(RadianAngles.Y)) * DEGREE
'            A.Y = AngleQuadrant(A.Y * DEGREE)
'            S.Y = VectorSine(mp)
'            c.Y = VectorCosine(mp)
'            t.Y = VectorTangent(mp)
'            ss.Y = VectorCosecant(mp)
'            cc.Y = VectorSecant(mp)
'            tt.Y = VectorCotangent(mp)
'
'            Set mp = MakePoint(PointToRotate.X, PointToRotate.Y, PointToRotate.z)
'            A.z = AngleZOfCoordXY(mp)
'            r.z = AngleRestrict(AngleRestrict(A.z) + AngleRestrict(RadianAngles.z)) * DEGREE
'            A.z = AngleQuadrant(A.z * DEGREE)
'            S.z = VectorSine(mp)
'            c.z = VectorCosine(mp)
'            t.z = VectorTangent(mp)
'            ss.z = VectorCosecant(mp)
'            cc.z = VectorSecant(mp)
'            tt.z = VectorCotangent(mp)


            
'            r.X = ((RadianAngles.X * S.X) + ((RadianAngles.Y * c.X) - (RadianAngles.z * ss.z)))
'            r.Y = ((RadianAngles.X * cc.Y) - ((RadianAngles.Y * cc.X) + (RadianAngles.z * S.z)))
'            r.z = ((RadianAngles.z * t.z) - ((RadianAngles.X * tt.Y) + (RadianAngles.Y * tt.X)))
'
'
'            r.Y = ((RadianAngles.X * S.X) + ((RadianAngles.X * c.X) - (RadianAngles.z * c.z)))
'            r.X = ((RadianAngles.Y * ss.Y) - ((RadianAngles.Y * cc.Y) + (RadianAngles.z * cc.z)))
'            r.z = ((RadianAngles.z * t.z) - ((RadianAngles.X * tt.Y) + (RadianAngles.Y * tt.X))) - _
'                    ((RadianAngles.z * tt.z) - ((RadianAngles.Y * c.Y) + (RadianAngles.X * ss.X)))
            
            

'            r.Y = -(((RadianAngles.X * S.X) + ((RadianAngles.X * ss.Y) - (RadianAngles.X * ss.z)))) '****
'
'            r.X = ((RadianAngles.Y * c.Y) + ((RadianAngles.Y * cc.X) - (RadianAngles.Y * cc.z))) - (r.Y * 2)
'
'            r.z = ((RadianAngles.z * t.z) - ((RadianAngles.X * tt.Y) + (RadianAngles.Y * tt.X))) - _
'                    ((RadianAngles.z * tt.z) - ((RadianAngles.Y * c.Y) + (RadianAngles.X * ss.X)))
'
'
'            r.X = AngleRestrict(r.X)
'            r.Y = AngleRestrict(r.Y)
'            r.z = 1 - AngleRestrict(r.z)
'
'            Debug.Print r.X; r.Y; r.z
'            Debug.Print Round(.X, 3); Round(.Y, 3); Round(.z, 3)
'            Debug.Print
'
'            .X = r.X
'            .Y = r.Y
'            .z = r.z
            
    End With
    Set tmp = Nothing
End Function

Public Function VectorRotateX(ByRef PointToRotate As Point, ByVal RadianAngle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(RadianAngle)
    SinPhi = Sin(RadianAngle)
    Set VectorRotateX = New Point
    With VectorRotateX
        .z = PointToRotate.z * CosPhi - PointToRotate.Y * SinPhi
        .Y = PointToRotate.z * SinPhi + PointToRotate.Y * CosPhi
        .X = PointToRotate.X
    End With
End Function

Public Function VectorRotateY(ByRef PointToRotate As Point, ByVal RadianAngle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(RadianAngle)
    SinPhi = Sin(RadianAngle)
    Set VectorRotateY = New Point
    With VectorRotateY
        .X = PointToRotate.X * CosPhi - PointToRotate.z * SinPhi
        .z = PointToRotate.X * SinPhi + PointToRotate.z * CosPhi
        .Y = PointToRotate.Y
    End With
End Function

Public Function VectorRotateZ(ByRef PointToRotate As Point, ByVal RadianAngle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(RadianAngle * -1)
    SinPhi = Sin(RadianAngle * -1)
    Set VectorRotateZ = New Point
    With VectorRotateZ
        .X = PointToRotate.X * CosPhi - PointToRotate.Y * SinPhi
        .Y = PointToRotate.X * SinPhi + PointToRotate.Y * CosPhi
        .z = PointToRotate.z
    End With
End Function


Public Function VectorAxisAngles(ByRef PointToZero As Point, Optional ByVal Combined As Boolean = True) As Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        If Not Combined Then
            .X = AngleZOfCoordXY(MakePoint(PointToZero.z, PointToZero.Y, PointToZero.X))
            .Y = AngleZOfCoordXY(MakePoint(PointToZero.X, PointToZero.z, PointToZero.Y))
            .z = AngleZOfCoordXY(MakePoint(PointToZero.X, PointToZero.Y, PointToZero.z))
        Else
            Dim magnitude As Single
            Dim heading As Single
            Dim pitch As Single
            Dim slope As Single

            magnitude = Sqr(PointToZero.X * PointToZero.X + PointToZero.Y * PointToZero.Y + PointToZero.z * PointToZero.z)
            If magnitude < 10 Then magnitude = 100
            slope = VectorSlope(MakePoint(0, 0, 0), PointToZero)
            heading = ATan2(PointToZero.z, PointToZero.X)
            pitch = ATan2(PointToZero.Y, Sqr(PointToZero.X * PointToZero.X + PointToZero.z * PointToZero.z))
            .X = (((heading / magnitude) - pitch) * (slope / magnitude))
            .z = ((PI / 2) + (-pitch + (heading / magnitude))) * (1 - (slope / magnitude))
            .Y = ((-heading + (pitch / magnitude)) * (1 - (slope / magnitude)))
            .Y = -(.Y + ((.X * (slope / magnitude)) / 2) - (.Y * 2) - ((.z * (slope / magnitude)) / 2))
            .X = (PI * 2) - (.X - ((PI / 2) * (slope / magnitude)))
            .z = (PI * 2) - (.z - ((PI / 2) * (slope / magnitude)))
        End If
    End With
    
    
'            Dim S As New Point 'sine
'            Dim c As New Point 'cosine
'            Dim t As New Point 'tangent
'            Dim ss As New Point 'cosecant
'            Dim cc As New Point 'secant
'            Dim tt As New Point 'cotangent
'            Dim mp As New Point 'temporary
'            Dim A As New Point 'angle
'            Dim r As New Point 'co-angle
'
'            Set mp = MakePoint(p1.z, p1.y, p1.x)
'            A.x = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.x = (A.x - ((A.x \ 90) * 90)) * RADIAN
'            S.x = VectorSine(mp)
'            c.x = VectorCosine(mp)
'            t.x = VectorTangent(mp)
'            ss.x = VectorCosecant(mp)
'            cc.x = VectorSecant(mp)
'            tt.x = VectorCotangent(mp)
'
'            Set mp = MakePoint(p1.x, p1.z, p1.y)
'            A.y = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.y = (A.y - ((A.y \ 90) * 90)) * RADIAN
'            S.y = VectorSine(mp)
'            c.y = VectorCosine(mp)
'            t.y = VectorTangent(mp)
'            ss.y = VectorCosecant(mp)
'            cc.y = VectorSecant(mp)
'            tt.y = VectorCotangent(mp)
'
'            Set mp = MakePoint(p1.x, p1.y, p1.z)
'            A.z = (AngleZOfCoordXY(mp) * DEGREE) * RADIAN
'            r.z = (A.z - ((A.z \ 90) * 90)) * RADIAN
'            S.z = VectorSine(mp)
'            c.z = VectorCosine(mp)
'            t.z = VectorTangent(mp)
'            ss.z = VectorCosecant(mp)
'            cc.z = VectorSecant(mp)
'            tt.z = VectorCotangent(mp)

End Function

Public Function AngleZOfCoordXY(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If Round(p.X, 6) = 0 Then
        'ony have a up and down line of x,y so it must be 360 or 180
        If Round(p.Y, 6) > 0 Then
            AngleZOfCoordXY = (360 * RADIAN)
        ElseIf Round(p.Y, 6) < 0 Then
            AngleZOfCoordXY = (180 * RADIAN)
        End If
    ElseIf Round(p.Y, 6) = 0 Then
        'only have a left or right line of x,y so it must be 90 or 270
        If Round(p.X, 6) > 0 Then
            AngleZOfCoordXY = (90 * RADIAN)
        ElseIf Round(p.X, 6) < 0 Then
            AngleZOfCoordXY = (270 * RADIAN)
        End If
    Else
        Dim dist As Single
        dist = Distance(0, 0, 0, p.X, p.Y, 0)
        'in clockwise starting at 0, or 12, for quadrants
        If Round(p.X, 6) > 0 And Round(p.Y, 6) > 0 Then
            'first quadrant
            If Abs(Round(p.Y, 6)) > Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = (45 + (45 - Round((InvSin(VectorSine(MakePoint(p.Y, p.X, 0))) * DEGREE + 2), 6))) * RADIAN
            ElseIf Abs(Round(p.Y, 6)) < Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = Round((InvSin(VectorSine(MakePoint(p.X, p.Y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (315 * RADIAN)
            End If
        ElseIf Round(p.X, 6) >= 0 And Round(p.Y, 6) <= 0 Then
            'second quadrant
            If Abs(Round(p.Y, 6)) > Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = (Round((InvSin(VectorSine(MakePoint(p.Y, p.X, 0))) * DEGREE + 2), 6) + 90) * RADIAN
            ElseIf Abs(Round(p.Y, 6)) < Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = (45 + -(45 - Round((InvSin(VectorSine(MakePoint(p.X, p.Y, 0))) * DEGREE + 2), 6))) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (45 * RADIAN)
            End If
        ElseIf Round(p.X, 6) <= 0 And Round(p.Y, 6) < 0 Then
            'third quadrant
            If Abs(Round(p.Y, 6)) > Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = (90 - ((45 - Round((InvSin(VectorSine(MakePoint(p.Y, p.X, 0))) * DEGREE + 2), 6)) - 45)) * RADIAN
            ElseIf Abs(Round(p.Y, 6)) < Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = -Round((InvSin(VectorSine(MakePoint(p.X, p.Y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (135 * RADIAN)
            End If
        ElseIf Round(p.X, 6) < 0 And Round(p.Y, 6) >= 0 Then
            'fourth quadrant
            If Abs(Round(p.Y, 6)) > Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = (45 + (45 - Round((InvSin(VectorSine(MakePoint(p.Y, p.X, 0))) * DEGREE + 2), 6))) * RADIAN
            ElseIf Abs(Round(p.Y, 6)) < Abs(Round(p.X, 6)) Then
                AngleZOfCoordXY = -Round((InvSin(VectorSine(MakePoint(p.X, p.Y, 0))) * DEGREE + 2), 6) * RADIAN
            Else
                AngleZOfCoordXY = AngleZOfCoordXY + (225 * RADIAN)
            End If
        ElseIf Round(p.X, 6) = Round(p.Y, 6) Then
            'diagnal line through the 0,0,0
            If p.X > 0 And p.Y > 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (45 * RADIAN)
            ElseIf p.X < 0 And p.Y > 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (135 * RADIAN)
            ElseIf p.X < 0 And p.Y < 0 Then
                AngleZOfCoordXY = AngleZOfCoordXY + (225 * RADIAN)
            ElseIf p.X > 0 And p.Y < 0 Then
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
    If p.X = 0 Then
       ' sec0 = CVErr(0)
    ElseIf p.Y = 0 And p.X > 0 Then
        VectorSecant = 1
    ElseIf p.Y = 0 And p.X < 0 Then
        VectorSecant = -1
    ElseIf p.X > 0 And p.Y <> 0 Then
        If VectorSecant < 0 Then VectorSecant = -VectorSecant
    ElseIf p.X < 0 And p.Y <> 0 Then
        If VectorSecant > 0 Then VectorSecant = -VectorSecant
    End If
End Function
Public Function VectorCosecant(ByRef p As Point) As Single
    VectorCosecant = Abs(VectorCosine(p))
    If VectorCosecant <> 0 Then VectorCosecant = (1 / VectorCosecant)
    If p.Y = 0 Then
        'csc0 = CVErr(0)
    ElseIf p.X = 0 And p.Y > 0 Then
        VectorCosecant = 1
    ElseIf p.X = 0 And p.Y < 0 Then
        VectorCosecant = -1
    ElseIf p.Y > 0 And p.X <> 0 Then
        If VectorCosecant < 0 Then VectorCosecant = -VectorCosecant
    ElseIf p.Y < 0 And p.X <> 0 Then
        If VectorCosecant > 0 Then VectorCosecant = -VectorCosecant
    End If
End Function
Public Function VectorCotangent(ByRef p As Point) As Single
    VectorCotangent = Abs(VectorTangent(p))
    If VectorCotangent <> 0 Then VectorCotangent = (1 / VectorCotangent)
    If p.Y = 0 And p.X <> 0 Then
        'cot0 = CVErr(0)
    ElseIf p.X = 0 And p.Y <> 0 Then
        VectorCotangent = 0
    ElseIf (p.X > 0 And p.Y > 0) Or (p.X < 0 And p.Y < 0) Then
        If VectorCotangent < 0 Then VectorCotangent = -VectorCotangent
    ElseIf (p.X < 0 And p.Y > 0) Or (p.X > 0 And p.Y < 0) Then
        If VectorCotangent > 0 Then VectorCotangent = -VectorCotangent
    End If
End Function

Public Function VectorTangent(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.X = 0 Then
        If p.Y > 0 Then
            VectorTangent = Val("1.#IND")
        ElseIf p.Y < 0 Then
            VectorTangent = 1
        End If
    ElseIf (p.Y <> 0) Then
        VectorTangent = Round(Abs(p.Y / p.X), 2)
    End If
    If p.X = 0 And p.Y <> 0 Then
        'tan0 = CVErr(0)
    ElseIf p.Y = 0 And p.X <> 0 Then
        VectorTangent = 0
    ElseIf (p.X > 0 And p.Y > 0) Or (p.X < 0 And p.Y < 0) Then
        If VectorTangent < 0 Then VectorTangent = -VectorTangent
    ElseIf (p.X < 0 And p.Y > 0) Or (p.X > 0 And p.Y < 0) Then
        If VectorTangent > 0 Then VectorTangent = -VectorTangent
    End If
End Function

Public Function VectorSine(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.X = 0 Then
        If p.Y <> 0 Then
            VectorSine = Val("0.#IND")
        End If
    ElseIf p.Y <> 0 Then
        VectorSine = Round(Abs(p.Y / Distance(0, 0, 0, p.X, p.Y, 0)), 2)
    End If
    If p.Y > 0 Then
        If p.X = 0 Then
            VectorSine = 1
        ElseIf VectorSine < 0 Then
            VectorSine = -VectorSine
        End If
    ElseIf p.Y < 0 Then
        If p.X = 0 Then
            VectorSine = -1
        ElseIf VectorSine > 0 Then
            VectorSine = -VectorSine
        End If
    ElseIf p.X <> 0 Then
        VectorSine = 0
    End If
End Function

Public Function VectorCosine(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.Y = 0 Then
        If p.X <> 0 Then
            VectorCosine = Val("1.#IND")
        End If
    ElseIf p.X <> 0 Then
        VectorCosine = Round(Abs(p.X / Distance(0, 0, 0, p.X, p.Y, 0)), 2)
    End If
    If p.X > 0 Then
        If p.Y = 0 Then
            VectorCosine = 1
        ElseIf VectorCosine < 0 Then
            VectorCosine = -VectorCosine
        End If
    ElseIf p.X < 0 Then
        If p.Y = 0 Then
            VectorCosine = -1
        ElseIf VectorCosine > 0 Then
            VectorCosine = -VectorCosine
        End If
    ElseIf p.Y <> 0 Then
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
        .X = AngleInvertRotation(p.X)
        .Y = AngleInvertRotation(p.Y)
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
    Set d1 = MakePoint(p1.X, p1.Y, p1.z)
    Set d2 = MakePoint(p2.X, p2.Y, p2.z)
    
    AngleAxisRestrict d1
    AngleAxisRestrict d2
    
    d1.X = d1.X * DEGREE
    d1.Y = d1.Y * DEGREE
    d1.z = d1.z * DEGREE
    
    d2.X = d2.X * DEGREE
    d2.Y = d2.Y * DEGREE
    d2.z = d2.z * DEGREE
    
    Dim c1 As Single
    Dim C2 As Single
    
    Set AngleAxisDifference = New Point
    With AngleAxisDifference
        c1 = Large(d1.X, d2.X)
        C2 = Least(d1.X, d2.X)
        .X = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        c1 = Large(d1.Y, d2.Y)
        C2 = Least(d1.Y, d2.Y)
        .Y = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
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
    Set VectorInfluence = MakePoint(Current.X, Current.Y, Current.z)
    With VectorInfluence
    
        Dim n As Point
        If Concurrent Then
            Set n = VectorNormalize(VectorDeduction(Current, Final))
            n.X = Abs(n.X)
            n.Y = Abs(n.Y)
            n.z = Abs(n.z)
        Else
            Set n = MakePoint(1, 1, 1)
        End If
        
        .X = ValueInfluence(Final.X, Current.X, Amount, n.X, SnapMinimum, SnapMaximum)
        .Y = ValueInfluence(Final.Y, Current.Y, Amount, n.Y, SnapMinimum, SnapMaximum)
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
    Set AngleAxisInfluence = MakePoint(Current.X, Current.Y, Current.z)
    With AngleAxisInfluence

        Dim n As Point
        If Concurrent Then
            Set n = VectorNormalize(AngleAxisDifference(Current, Final))
            n.X = Abs(n.X)
            n.Y = Abs(n.Y)
            n.z = Abs(n.z)
        Else
            Set n = MakePoint(1, 1, 1)
        End If
        
        .X = AngleInfluence(Final.X, Current.X, Amount, n.X, SnapMinimum, SnapMaximum)
        .Y = AngleInfluence(Final.Y, Current.Y, Amount, n.Y, SnapMinimum, SnapMaximum)
        .z = AngleInfluence(Final.z, Current.z, Amount, n.z, SnapMinimum, SnapMaximum)
        
    End With
End Function

Public Function VectorRise(ByRef p1 As Point, Optional ByRef p2 As Point = Nothing) As Single
    VectorRise = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
End Function

Public Function VectorRun(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorRun = DistanceEx(MakePoint(p1.X, 0, p1.z), MakePoint(p2.X, 0, p2.z))
End Function

Public Function VectorSlope(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorSlope = VectorRun(p1, p2)
    If (VectorSlope <> 0) Then
        VectorSlope = Round((VectorRise(p1, p2) / VectorSlope), 6)
        If (VectorSlope = 0) Then VectorSlope = -CInt(Not ((p1.X = p2.X) And (p1.Y = p2.Y) And (p1.z = p2.z)))
    ElseIf VectorRise(p1, p2) <> 0 Then
        VectorSlope = 1
    End If
End Function

Public Function VectorYIntercept(ByRef p1 As Point, ByRef p2 As Point) As Single
    With VectorMidPoint(p1, p2)
        VectorYIntercept = VectorSlope(p1, p2)
        VectorYIntercept = -((VectorYIntercept * .X) - .Y)
    End With
End Function

Public Function VectorDotProduct(ByRef p1 As Point, ByRef p2 As Point) As Single
    VectorDotProduct = ((p1.X * p2.X) + (p1.Y * p2.Y) + (p1.z * p2.z))
End Function

Public Function VectorMultiply(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMultiply = New Point
    With VectorMultiply
        .X = (p1.X * p2.X)
        .Y = (p1.Y * p2.Y)
        .z = (p1.z * p2.z)
    End With
End Function

Public Function VectorCrossProduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCrossProduct = New Point
    With VectorCrossProduct
        .X = ((p1.Y * p2.z) - (p1.z * p2.Y))
        .Y = ((p1.z * p2.X) - (p1.X * p2.z))
        .z = ((p1.X * p2.Y) - (p1.Y * p2.X))
    End With
End Function

Public Function VectorSubtraction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorSubtraction = New Point
    With VectorSubtraction
        .X = ((p1.X - p2.z) - (p1.X - p2.Y))
        .Y = ((p1.Y - p2.X) - (p1.Y - p2.z))
        .z = ((p1.z - p2.Y) - (p1.z - p2.X))
    End With
End Function

Public Function VectorAccordance(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAccordance = New Point
    With VectorAccordance
        .X = (((p1.X + p1.Y) - p2.z) + ((p1.z + p1.X) - p2.Y) - ((p1.Y + p1.z) - p2.X))
        .Y = (((p1.Y + p1.z) - p2.X) + ((p1.X + p1.Y) - p2.z) - ((p1.z + p1.X) - p2.Y))
        .z = (((p1.z + p1.X) - p2.Y) + ((p1.Y + p1.z) - p2.X) - ((p1.X + p1.Y) - p2.z))
    End With
End Function

Public Function VectorDisplace(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDisplace = New Point
    With VectorDisplace
        .X = (Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.z)) + Abs((Abs(p1.z) + Abs(p1.X)) - Abs(p2.Y)) - Abs((Abs(p1.Y) + Abs(p1.z)) - Abs(p2.X)))
        .Y = (Abs((Abs(p1.Y) + Abs(p1.z)) - Abs(p2.X)) + Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.z)) - Abs((Abs(p1.z) + Abs(p1.X)) - Abs(p2.Y)))
        .z = (Abs((Abs(p1.z) + Abs(p1.X)) - Abs(p2.Y)) + Abs((Abs(p1.Y) + Abs(p1.z)) - Abs(p2.X)) - Abs((Abs(p1.X) + Abs(p1.Y)) - Abs(p2.z)))
    End With
End Function

Public Function VectorOffset(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorOffset = New Point
    With VectorOffset
        .X = (Large(p1.X, p2.X) - Least(p1.X, p2.X))
        .Y = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
        .z = (Large(p1.z, p2.z) - Least(p1.z, p2.z))
    End With
End Function

Public Function VectorQuantify(ByRef p1 As Point) As Single
    VectorQuantify = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.z))
End Function


Public Function VectorDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDeduction = New Point
    With VectorDeduction
        .X = (p1.X - p2.X)
        .Y = (p1.Y - p2.Y)
        .z = (p1.z - p2.z)
    End With
End Function

Public Function VectorCrossDeduct(ByRef p1 As Point, ByRef p2 As Point)
    Set VectorCrossDeduct = New Point
    With VectorCrossDeduct
        .X = (p1.X - p2.z)
        .Y = (p1.Y - p2.X)
        .z = (p1.z - p2.Y)
    End With
End Function

Public Function VectorAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorAddition = New Point
    With VectorAddition
        .X = (p1.X + p2.X)
        .Y = (p1.Y + p2.Y)
        .z = (p1.z + p2.z)
    End With
End Function

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .X = (p1.X * n)
        .Y = (p1.Y * n)
        .z = (p1.z * n)
    End With
End Function

Public Function VectorCombination(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorCombination = New Point
    With VectorCombination
        .X = ((p1.X + p2.X) / 2)
        .Y = ((p1.Y + p2.Y) / 2)
        .z = ((p1.z + p2.z) / 2)
    End With
End Function

Public Function VectorNormalize(ByRef p1 As Point) As Point
    Set VectorNormalize = New Point
    With VectorNormalize
        .z = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.z))
        If (Round(.z, 6) > 0) Then
            .z = (1 / .z)
            .X = (p1.X * .z)
            .Y = (p1.Y * .z)
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
            .X = ((p2.X - p1.X) * .z)
            .Y = ((p2.Y - p1.Y) * .z)
            .z = ((p2.z - p1.z) * .z)
        End If
    End With
End Function

Public Function VectorMidPoint(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorMidPoint = New Point
    With VectorMidPoint
        .X = ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) / 2) + Least(p1.X, p2.X)
        .Y = ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) / 2) + Least(p1.Y, p2.Y)
        .z = ((Large(p1.z, p2.z) - Least(p1.z, p2.z)) / 2) + Least(p1.z, p2.z)
    End With
End Function

Public Function VectorNegative(ByRef p1 As Point) As Point
    Set VectorNegative = New Point
    With VectorNegative
        .X = -p1.X
        .Y = -p1.Y
        .z = -p1.z
    End With
End Function

Public Function VectorDivision(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorDivision = New Point
    With VectorDivision
        .X = (p1.X / n)
        .Y = (p1.Y / n)
        .z = (p1.z / n)
    End With
End Function

Public Function VectorIsNormal(ByRef p1 As Point) As Boolean
    'returns if a point provided is normalized, to the best of ability
    VectorIsNormal = (Round(Abs(p1.X) + Abs(p1.Y) + Abs(p1.z), 0) = 1) 'first kind is the absolute of all values equals one
    VectorIsNormal = VectorIsNormal Or (DistanceEx(MakePoint(0, 0, 0), p1) = 1) 'another is the total length of vector is one
    'another is if any value exists non zero as well as adding up in any non specific arrangement cancels to zero, as has one
    VectorIsNormal = VectorIsNormal Or ((p1.X <> 0 Or p1.Y <> 0 Or p1.z <> 0) And (( _
        ((p1.X + p1.Y + p1.z) = 0) Or ((p1.Y + p1.z + p1.X) = 0) Or ((p1.z + p1.X + p1.Y) = 0) Or _
        ((p1.X + p1.z + p1.Y) = 0) Or ((p1.z + p1.Y + p1.X) = 0) Or ((p1.Y + p1.X + p1.z) = 0) _
        )))
    Dim tmp As Single
    'another is a reflection test and check if it falls with in -1 to 1 for triangle normals
    'reflection is 27 groups of three arithmitic (-1+(2-3)) and by the third group, the groups
    'reflect the same (-g+(g-g)) which are sub groups of lines of three groups doing the same
    tmp = -((-(-p1.X + (p1.Y - p1.z)) + ((-p1.Y + (p1.z - p1.X)) - (-p1.z + (p1.X - p1.Y)))) + _
        ((-p1.Y + (p1.z - p1.X)) + ((-p1.z + (p1.X - p1.Y)) - (-p1.X + (p1.Y - p1.z))) - _
        (-p1.z + (p1.X - p1.Y)) + ((-p1.X + (p1.Y - p1.z)) - (-p1.Y + (p1.z - p1.X))))) + ( _
        ((-(-p1.Y + (p1.X - p1.z)) + ((-p1.X + (p1.z - p1.Y)) - (-p1.z + (p1.Y - p1.X)))) + _
        ((-p1.X + (p1.z - p1.Y)) + ((-p1.z + (p1.Y - p1.X)) - (-p1.Y + (p1.X - p1.z))) - _
        (-p1.z + (p1.Y - p1.X)) + ((-p1.Y + (p1.X - p1.z)) - (-p1.X + (p1.z - p1.Y))))) - _
        ((-(-p1.z + (p1.Y - p1.X)) + ((-p1.Y + (p1.X - p1.z)) - (-p1.X + (p1.z - p1.Y)))) + _
        ((-p1.Y + (p1.X - p1.z)) + ((-p1.X + (p1.z - p1.Y)) - (-p1.z + (p1.Y - p1.X))) - _
        (-p1.X + (p1.z - p1.Y)) + ((-p1.z + (p1.Y - p1.X)) - (-p1.Y + (p1.X - p1.z))))))
        '9 lines, 27 groups, 81 values, full circle, the first value (-negative, plus (second minus third))
    VectorIsNormal = VectorIsNormal Or ((p1.X <> 0 Or p1.Y <> 0 Or p1.z <> 0) And (tmp >= -1 And tmp <= 1))
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
    If p.Y > 0 And p.X >= 0 Then
        VectorQuadrant = 1
    ElseIf p.Y >= 0 And p.X < 0 Then
        VectorQuadrant = 2
    ElseIf p.Y < 0 And p.X <= 0 Then
        VectorQuadrant = 3
    ElseIf p.Y <= 0 And p.X > 0 Then
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

