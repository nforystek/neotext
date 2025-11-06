Attribute VB_Name = "modGeometry"
Option Explicit

Option Compare Binary

'############################################################################################
'############################################################################################
'############################################################################################
'############################################################################################
        
'distance
'd = ((x2 - x1)^2 + (y2 - y1)^2)^(1/2)

'X displace
'x = (d^2 - (y2 - y1)^2)^(1/2)

'Y displace
'y = (d^2 - (x2 - x1)^2)^(1/2)

'slope
'm = (y / x)
        
'Y-Intercept
'b = -((m * x) - y)

'X & Y coordinates
'y = ((m * x) + b)

'x = ((y - b) / m)

'########### The majority of this library angles are in radians

Public Const PI As Double = 3.14159265359
'Public Const PI As Double = 3.14159265358979
Public Const Epsilon As Double = 0.999999999999999 ' 0.0001 '

Public Const DEGREE As Double = (180 / PI)
Public Const RADIAN As Double = (PI / 180)

Public Const D90 As Double = (PI / 4)
Public Const D180 As Double = (PI / 2)
Public Const D360 As Double = PI
Public Const D720 As Double = (PI * 2)

Public Const FOOT As Double = 0.1
Public Const MILE As Double = 5280 * FOOT

'Public Const FOVY As Double = (FOOT * 8) '(FOOT * 8) '4 feet left, and 4 feet right = 0.8
Public Const FOVY As Double = 1.047198 '2.3561946
Public Const SKYFOVY As Double = (MILE * 4)

Public Const FAR  As Double = 900000000
Public Const NEAR As Double = 0 '0.05 'one millimeter (308.4 per foot) or greater




'Conversions
' ?round(RADIAN,6)
' 0.017453
'?round(pi/2/100,6)
' 0.015708
'?DEGREE
' 57.29578
'?57.29578/(pi/4)
'? 72.9512508123925 /100
' 0.729512508123925
'?(pi/4)
 '0.7853982
 
'Public Function Atan2(Y, X)
'    If X > 0 Then
'        Atan2 = Atn(Y / X)
'    ElseIf X < 0 And Y >= 0 Then
'        Atan2 = Atn(Y / X) + PI
'    ElseIf X < 0 And Y < 0 Then
'        Atan2 = Atn(Y / X) - PI
'    ElseIf X = 0 And Y > 0 Then
'        Atan2 = PI / 2
'    ElseIf X = 0 And Y < 0 Then
'        Atan2 = -PI / 2
'    End If
'End Function




Public Function ZeroRotation() As Point
    Set ZeroRotation = New Point
    With ZeroRotation
        .X = (PI * 2)
        .Y = (PI * 2)
        .Z = (PI * 2)
    End With
    AngleAxisRestrict ZeroRotation
End Function
Public Function MakeVector(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
End Function

Public Function MakePoint(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
End Function

Public Function MakePlot(ByVal X As Double, ByVal Y As Double) As Plot
    Set MakePlot = New Plot
    MakePlot.X = X
    MakePlot.Y = Y
End Function

Public Function ToPlot(ByRef Vector As D3DVECTOR) As Plot
    Set ToPlot = New Plot
    ToPlot.X = Vector.X
    ToPlot.Y = Vector.Y
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

Public Function ToPlane(ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Plane
        
    Dim pNormal As Point
    Set pNormal = VectorCrossProduct(VectorDeduction(V2, V1), VectorDeduction(V3, V1))
    Set pNormal = VectorNormalize(pNormal)
        
    Set ToPlane = New Plane
    With ToPlane
        .W = VectorDotProduct(pNormal, V1) * -1
        .X = pNormal.X
        .Y = pNormal.Y
        .Z = pNormal.Z
    End With
End Function

Public Function ToVec4(ByRef Plane As Plane) As D3DVECTOR4
    ToVec4.X = Plane.X
    ToVec4.Y = Plane.Y
    ToVec4.Z = Plane.Z
    ToVec4.W = Plane.W
End Function

Public Function DistanceToPlane(ByRef p As Point, ByRef r As Plane) As Double
    If Sqr(r.X * r.X + r.Y * r.Y + r.Z * r.Z) <> 0 Then
        DistanceToPlane = (r.X * p.X + r.Y * p.Y + r.Z * p.Z + r.W) / Sqr(r.X * r.X + r.Y * r.Y + r.Z * r.Z)
    End If
End Function

Public Function Distance(ByVal p1x As Double, ByVal p1y As Double, ByVal p1z As Double, ByVal p2x As Double, ByVal p2y As Double, ByVal p2z As Double) As Double
    Distance = (((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2))
    If Distance <> 0 Then Distance = Distance ^ (1 / 2)
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
        
'                .x = ((d * p2.x) + (n * p1.x)) / (d + n)
'                .y = ((d * p2.y) + (n * p1.y)) / (d + n)
'                .z = ((d * p2.z) + (n * p1.z)) / (d + n)
        
'#
'                .x = Large(p1.x, p2.x) - Least(p1.x, p2.x)
'                .y = Large(p1.y, p2.y) - Least(p1.y, p2.y)
'                .z = Large(p1.z, p2.z) - Least(p1.z, p2.z)
'                .x = (Least(p1.x, p2.x) + (n * (.x / d)))
'                .y = (Least(p1.y, p2.y) + (n * (.y / d)))
'                .z = (Least(p1.z, p2.z) + (n * (.z / d)))
'#
                .X = p2.X - p1.X
                .Y = p2.Y - p1.Y
                .Z = p2.Z - p1.Z
                .X = (p1.X + (N * (.X / d)))
                .Y = (p1.Y + (N * (.Y / d)))
                .Z = (p1.Z + (N * (.Z / d)))
'#
                
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

Public Function PointOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    Dim r As Plane
    Set r = ToPlane(v0, V1, V2)
    PointOnPlane = ((r.X * (p.X - v0.X)) + (r.Y * (p.Y - v0.Y)) + (r.Z * (p.Z - v0.Z)) = 0)
End Function

Public Function PointSideOfPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    PointSideOfPlane = VectorDotProduct(PlaneNormal(v0, V1, V2), p) > 0
End Function

Public Function PointNearOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Point
    Set PointNearOnPlane = New Point
    With PointNearOnPlane
        Dim r As Plane
        Set r = ToPlane(v0, V1, V2)
        Dim N As Point
        Set N = PlaneNormal(v0, V1, V2)
        Dim d As Double
        d = DistanceToPlane(p, r)
        .X = p.X - (d * N.X)
        .Y = p.Y - (d * N.Y)
        .Z = p.Z - (d * N.Z)
    End With
End Function

Public Function LinePointByPercent(ByRef p1 As Point, ByRef p2 As Point, ByVal factor As Double) As Point
    Set LinePointByPercent = New Point
    With LinePointByPercent
        .X = Least(p1.X, p2.X) + ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) * factor)
        .Y = Least(p1.Y, p2.Y) + ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) * factor)
        .Z = Least(p1.Z, p2.Z) + ((Large(p1.Z, p2.Z) - Least(p1.Z, p2.Z)) * factor)
    End With
End Function
Public Function LineOpposite(ByVal length1 As Double, ByVal length2 As Double, ByVal length3 As Double) As Double
    LineOpposite = Least(length1, length2, length3)
End Function

Public Function LineAdjacent(ByVal length1 As Double, ByVal length2 As Double, ByVal length3 As Double) As Double
    LineAdjacent = Large(Least(length1, length2), Large(Least(length2, length3), Least(length3, length1)))
End Function

Public Function LineHypotenuse(ByVal length1 As Double, ByVal length2 As Double, ByVal length3 As Double) As Double
    LineHypotenuse = Large(length1, length2, length3)
End Function

Public Function LineIntersectPlane(ByRef Plane As Plane, PStart As Point, vDir As Point, ByRef VIntersectOut As Point) As Boolean
    Dim q As New Plane     'Start Point
    Dim v As New Plane       'Vector Direction

    Dim planeQdot As Double 'Dot products
    Dim planeVdot As Double
    
    Dim t As Double         'Part of the equation for a ray P(t) = Q + tV
    
    q.X = PStart.X          'Q is a point and therefore it's W value is 1
    q.Y = PStart.Y
    q.Z = PStart.Z
    q.W = 1
    
    v.X = vDir.X            'V is a vector and therefore it's W value is zero
    v.Y = vDir.Y
    v.Z = vDir.Z
    v.W = 0
    
    '((Plane.X * V.X) + (Plane.Y * V.Y) + (Plane.z * V.z) + (Plane.R * V.R))
    
    planeVdot = ((Plane.X * v.X) + (Plane.Y * v.Y) + (Plane.Z * v.Z) + (Plane.W * v.W)) 'D3DXVec4Dot(Plane, V)
    planeQdot = ((Plane.X * q.X) + (Plane.Y * q.Y) + (Plane.Z * q.Z) + (Plane.W * q.W)) 'D3DXVec4Dot(Plane, Q)
            
    'If the dotproduct of plane and V = 0 then there is no intersection
    If planeVdot <> 0 Then
        t = Round((planeQdot / planeVdot) * -1, 5)
        
        If VIntersectOut Is Nothing Then Set VIntersectOut = New Point
        
        'This is where the line intersects the plane
        VIntersectOut.X = Round(q.X + (t * v.X), 5)
        VIntersectOut.Y = Round(q.Y + (t * v.Y), 5)
        VIntersectOut.Z = Round(q.Z + (t * v.Z), 5)

        LineIntersectPlane = True
    Else
        'No Collision
        LineIntersectPlane = False
    End If
    
End Function

Public Function LineIntersectLine2DEx(ByRef l1p1 As Point, ByRef l1p2 As Point, ByRef l2p1 As Point, ByRef l2p2 As Point) As Point

    Set LineIntersectLine2DEx = LineIntersectLine2D(l1p1.X, l1p1.Y, l1p2.X, l1p2.Y, l2p1.X, l2p1.Y, l2p2.X, l2p2.Y)

End Function


Public Function LineIntersectLine2D(ByVal l1p1x As Double, ByVal l1p1y As Double, ByVal l1p2x As Double, ByVal l1p2y As Double, ByVal l2p1x As Double, ByVal l2p1y As Double, ByVal l2p2x As Double, ByVal l2p2y As Double) As Point

    Dim b As Double
    b = (((l2p2y - l2p1y) * (l1p2x - l1p1x)) - ((l2p2x - l2p1x) * (l1p2y - l1p1y)))

    If b <> 0 Then

        Dim t As Double
        Dim u As Double

        t = (((l2p2x - l2p1x) * (l1p1y - l2p1y)) - ((l2p2y - l2p1y) * (l1p1x - l2p1x))) / b
        u = (((l2p1y - l1p1y) * (l1p1x - l1p2x)) - ((l2p1x - l1p1x) * (l1p1y - l1p2y))) / b
 
        If t >= 0 And t <= 1 And u >= 0 And u <= 1 Then
            Set LineIntersectLine2D = New Point
            LineIntersectLine2D.X = Lerp(l1p1x, l1p2x, t)
            LineIntersectLine2D.Y = Lerp(l1p1y, l1p2y, t)
            LineIntersectLine2D.Z = t
        End If
    End If

End Function

Public Function TriangleIntersect(ByRef t1p1 As Point, ByRef t1p2 As Point, ByRef t1p3 As Point, ByRef t2p1 As Point, ByRef t2p2 As Point, ByRef t2p3 As Point) As Point
'    'Debug.Print TriangleIntersect(MakePoint(-5, 0, 0), MakePoint(5, 0, -5), MakePoint(5, 0, 5), MakePoint(-2.5, 2.5, 0), MakePoint(2.5, 2.5, 0), MakePoint(0, -2.5, 0))
'    'compute another way of representing triangles, the center, normal and side lengths
'    Dim t1n As New Point
'    Dim t2n As New Point
'    Dim t1a As New Point
'    Dim t2a As New Point
'    Dim t1l As New Point
'    Dim t2l As New Point
'
'    t1n = TriangleNormal(t1p1, t1p2, t1p3)
'    t2n = TriangleNormal(t2p1, t2p2, t2p3)
'
'    'Debug.Print t1n; VectorIsNormal(t1n); t2n; VectorIsNormal(t2n)
'
''    t1n = PlaneNormal(t1p1, t1p2, t1p3)
''    t2n = PlaneNormal(t2p1, t2p2, t2p3)
''    Debug.Print t1n; VectorIsNormal(t1n); t2n; VectorIsNormal(t2n)
'
'    t1a = TriangleAxii(t1p1, t1p2, t1p3)
'    t2a = TriangleAxii(t2p1, t2p2, t2p3)
'
'    Debug.Print t1a; t2a
'
'    t1l.X = DistanceEx(t1p1, t1p2)
'    t1l.Y = DistanceEx(t1p2, t1p3)
'    t1l.Z = DistanceEx(t1p3, t1p1)
'
'    t2l.X = DistanceEx(t2p1, t2p2)
'    t2l.Y = DistanceEx(t2p2, t2p3)
'    t2l.Z = DistanceEx(t2p3, t2p1)
'
'    Debug.Print t1l; t2l
'
'
'    Set TriangleIntersect = New Point
'    With TriangleIntersect
'
'
'    End With
End Function

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Double
    RandomPositive = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function PlaneNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    'returns a vector perpendicular to a plane V, at 0,0,0, with out the local coordinates information
    Set PlaneNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(v0, V1), VectorDeduction(V1, V2)))
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
Public Function Sign(ByVal N As Double) As Double
    Sign = ((-(Abs(N - 1) - N) - (-Abs(N + 1) + N)) * 0.5)
End Function

Public Function Signn(ByVal Value As Double) As Double
    Signn = ((-((AbsoluteWhole(Value) / 1 <> 0) * 1) + -1) + -(((-AbsoluteWhole(Value) / 1 + -1) = 0) * 1))
End Function

Public Function SphereSurfaceArea(ByVal Radii As Double) As Double
     SphereSurfaceArea = (4 * PI * (Radii ^ 2))
End Function

Public Function SphereVolume(ByVal Radii As Double) As Double
    SphereVolume = ((4 / 3) * PI * (Radii ^ 3))
End Function

Public Function SphereToCubeRoot(ByVal Diameter As Double) As Double
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
        .Z = (Least(v0.Z, V1.Z, V2.Z, V3.Z) + ((Large(v0.Z, V1.Z, V2.Z, V3.Z) - Least(v0.Z, V1.Z, V2.Z, V3.Z)) / 2))
    End With
End Function

Public Function CirclePermeter(ByVal Radii As Double) As Double
    CirclePermeter = ((Radii * 2) * PI)
End Function

Public Function CubeToSphereDiameter(ByVal edge1 As Double, Optional ByVal edge2 As Double = 0, Optional ByVal edge3 As Double = 0) As Double
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
Public Function CubePerimeter(ByVal edge1 As Double, Optional ByVal edge2 As Double = 0, Optional ByVal edge3 As Double = 0) As Double
    If edge2 = 0 And edge3 = 0 Then
        CubePerimeter = (edge1 * 12)
    Else
        CubePerimeter = (edge1 * 4) + (edge2 * 4) + (edge3 * 4)
    End If
End Function

Public Function CubeSurfaceArea(ByVal Edge As Double) As Double
    CubeSurfaceArea = (6 * (Edge ^ 2))
End Function

Public Function CubeVolume(ByVal Edge As Double) As Double
    CubeVolume = (Edge ^ 3)
End Function


Public Function TrianglePerimeter(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
    TrianglePerimeter = (DistanceEx(p1, p2) + DistanceEx(p2, p3) + DistanceEx(p3, p1))
End Function

Function TriangleArea1(ByVal A As Double, ByVal b As Double, ByVal C As Double) As Double
    'I'm not sure this is anything correct, it doesn't seem to be acurate the higher it is
    'but it was an attempt to develop it logically using the 1/2 base * height
    'I think the function just under this is more accurate perhaps not though.
    'the reality is more likely as if, a traingle is a 3D object and can't
    'represent more then one face, technically a prisim, that's no 2D area
    
    Dim d As Double
    Dim e As Double
    Dim F As Double
    Dim g As Double
    Dim H As Double

    'make c the largest side, doing so
    'sort us the base in any situation
    If A > C Then
        'swap
        d = A
        A = C
        C = d
    End If
    If b > C Then
        'swap
        d = b
        b = C
        C = d
    End If
    
    If A + b < C Then
        'invalid triangle
        Exit Function
    End If
    
    'now make c the odd side
    'if two sides are equal
    If A = C Then
        d = b
        b = C
        C = d
    End If
    If b = C Then
        d = A
        A = C
        C = d
    End If
    
    'now we have, c is our largest base or
    'it is unique among a and b are equal
    'for our base calcualtions, with that
    
    'let's cut c in two, too form two
    'triabgles too apply the formula:
    'area = ((1/2) * B * H) to each.
    
    'because c is largest or a and b are
    'equal, we can use it as a unit whole
    'when forming two triangles, to find
    'a common center or point where to
    'cut c in two parts, where as a and b
    'are equal that point is exactly half.
    'all other opportunity, both a and b
    'represent a portion in their different
    'and are not larger then C, a percent
    'they may represent then if c is whole
    
    d = (A + b) 'a total unit whole
    
    e = (A / d) 'a percent of the unit that a is
    e = (e * C) 'applied to c for where to split
    
    F = (b / d) 'do it again, for b
    F = (C * F) 'proof rill be same as (c-f)
    
    'Debug.Print (Round(e, 6) = (c - Round(f, 6))) = True
    
    'now two trinagles are formed with a - e and b - f
    'that we can get the heights for with pythagorean
    'as the split in C forms right traingles
    'where a dn b are the hypotenuse
    
    g = (((A ^ 2) - (e ^ 2)) ^ (1 / 2))
    H = (((b ^ 2) - (F ^ 2)) ^ (1 / 2))
    
    'now do the area formula for each
    'area = ((1/2) * B * H)
     
    d = ((1 / 2) * A * g)
    e = ((1 / 2) * b * H)
    
    'finally add the two areas for the
    'original traingles total area
    TriangleArea1 = (d + e)

End Function

Public Function TriangleSurfaceArea(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
    Dim l1 As Double: l1 = DistanceEx(p1, p2)
    Dim l2 As Double: l2 = DistanceEx(p2, p3)
    Dim l3 As Double: l3 = DistanceEx(p3, p1)
    TriangleSurfaceArea = (((((((l1 + l2) - l3) + ((l2 + l3) - l1) + ((l3 + l1) - l2)) * (l1 * l2 * l3)) / (l1 + l2 + l3)) ^ (1 / 2)))
End Function

Public Function TriangleVolume(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
    TriangleVolume = TriangleSurfaceArea(p1, p2, p3)
    TriangleVolume = ((((TriangleVolume ^ (1 / 3)) ^ 2) ^ 3) / 12)
End Function

Public Function TriangleDotProduct(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
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

Public Function TriangleLowestOfAll(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleLowestOfAll = New Point
    With TriangleLowestOfAll
        .X = Least(p1.X, p2.X, p3.X)
        .Y = Least(p1.Y, p2.Y, p3.Y)
        .Z = Least(p1.Z, p2.Z, p3.Z)
    End With
End Function

Public Function TriangleLargestOfAll(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleLargestOfAll = New Point
    With TriangleLargestOfAll
        .X = Large(p1.X, p2.X, p3.X)
        .Y = Large(p1.Y, p2.Y, p3.Y)
        .Z = Large(p1.Z, p2.Z, p3.Z)
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
#If NTDirectX = -1 Then
'Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
'    Set TriangleNormal = New Point
'    Dim o As Point
'    Dim d As Double
'    With TriangleNormal
'        Set o = TriangleDisplace(v0, V1, V2)
'        d = (o.X + o.Y + o.z)
'        If (d <> 0) Then
'            .z = (((o.X + o.Y) - o.z) / d)
'            .X = (((o.Y + o.z) - o.X) / d)
'            .Y = (((o.z + o.X) - o.Y) / d)
'        End If
'    End With
'End Function

'Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
'    Set TriangleNormal = New Point
'    Dim o As Point
'    Dim d As Double
'    With TriangleNormal
'        Set o = TriangleDisplace(v0, V1, V2)
'        d = (o.X + o.Y + o.z)
'        If (d <> 0) Then
'            .z = (((o.X + o.Y) - o.z) / d)
'            .X = (((o.Y + o.z) - o.X) / d)
'            .Y = (((o.z + o.X) - o.Y) / d)
'        End If
'    End With
'End Function

Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    Set TriangleNormal = New Point
    Dim o As Point
    Dim d As Double
    With TriangleNormal
        Set o = TriangleDisplace(v0, V1, V2)
        d = (Abs(o.X) + Abs(o.Y) + Abs(o.Z))
        If (d <> 0) Then
            .Z = (((Abs(o.X) + Abs(o.Y)) - Abs(o.Z)) / d)
            .X = (((Abs(o.Y) + Abs(o.Z)) - Abs(o.X)) / d)
            .Y = (((Abs(o.Z) + Abs(o.X)) - Abs(o.Y)) / d)
        End If
    End With
End Function

#End If

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

Public Function VectorBalance(ByRef loZero As Point, ByRef hiWhole As Point, ByVal FulcrumPercent As Double) As Point
    Set VectorBalance = New Point
    With VectorBalance
        .X = (loZero.X + ((hiWhole.X - loZero.X) * FulcrumPercent))
        .Y = (loZero.Y + ((hiWhole.Y - loZero.Y) * FulcrumPercent))
        .Z = (loZero.Z + ((hiWhole.Z - loZero.Z) * FulcrumPercent))
    End With
End Function

Public Function TriangleFulcrum(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Point
    Set TriangleFulcrum = New Point
    With TriangleFulcrum
        If (Not p3 Is Nothing) Then
            .X = (p3.X ^ 2)
            .Y = (p3.Y ^ 2)
            .Z = (p3.Z ^ 2)
        End If
        .X = (.X + (p1.X ^ 2) + (p2.X ^ 2)) ^ (1 / 2)
        .Y = (.Y + (p1.Y ^ 2) + (p2.Y ^ 2)) ^ (1 / 2)
        .Z = (.Z + (p1.Z ^ 2) + (p2.Z ^ 2)) ^ (1 / 2)
    End With
End Function




'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################

Public Function LineYIntercept(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Double
    '2D by nature of always exists, and not for 3D
    'Y-Intercept
    'b = -((m * x) - y)
    If p1 Is Nothing Then
        LineYIntercept = -((LineSlope2D(p2, p1) * p2.X) - p2.Y)
    Else
        LineYIntercept = -((LineSlope2D(p2, p1) * (p2.X - p1.X)) - (p2.Y - p1.Y))
    End If
End Function

Public Function LineSlope2D(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Double
    'slope
    'm = (y / x)
    If p1 Is Nothing Then
        If p2.X <> 0 Then LineSlope2D = p2.Y / p2.X    'rise over run
    Else
        If (p2.X - p1.X) <> 0 Then LineSlope2D = (p2.Y - p1.Y) / (p2.X - p1.X) 'rise over run
    End If
End Function

Public Function LineSlope3D(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Double
    If p1 Is Nothing Then Set p1 = New Point
     'run is the distance formula excluding the Y coordinate
    LineSlope3D = (((p2.X - p1.X) ^ 2) + ((p2.Z - p1.Z) ^ 2)) ^ (1 / 2)
    If LineSlope3D <> 0 Then 'rise doesn't include x or z, so now it's the same
        LineSlope3D = -((p2.Y - p1.Y) / LineSlope3D) 'rise over run
    Else
        LineSlope3D = 0
    End If
End Function


'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function AngleRestrictDegree(ByRef Angle As Double) As Double
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
    AngleRestrictDegree = Angle
End Function

Public Function AngleAxisRestrictDegree(ByRef Angles As Point) As Point
    '3 axis version of AngleRestrictDegree(Angle)
    Angles.X = AngleRestrictDegree(Angles.X)
    Angles.Y = AngleRestrictDegree(Angles.Y)
    Angles.Z = AngleRestrictDegree(Angles.Z)
    Set AngleAxisRestrictDegree = Angles
End Function

Public Function AngleRestrict(ByRef Angle1 As Double) As Double
    Angle1 = Angle1 * DEGREE
    Angle1 = AngleRestrictDegree(Angle1)
    AngleRestrict = Round(Angle1 * RADIAN, 6)
    If AngleRestrict = PI Or AngleRestrict = PI * 2 Or AngleRestrict = 0 Then
        AngleRestrict = (-AngleRestrict + (PI * 2)) + -(PI * 4)
    End If
End Function



Public Function AngleAxisRestrict(ByRef AxisAngles As Point) As Point
    AxisAngles.X = AngleRestrict(AxisAngles.X)
    AxisAngles.Y = AngleRestrict(AxisAngles.Y)
    AxisAngles.Z = AngleRestrict(AxisAngles.Z)
    Set AngleAxisRestrict = AxisAngles
End Function



'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorAxisAngles(ByRef p As Point) As Point
'    Set VectorAxisAngles = New Point
'    Dim tmp As New Point
'    With VectorAxisAngles
'        If Not (p.X = 0 And p.Y = 0 And p.Z = 0) Then
'            Set tmp = p
'            .X = AngleRestrict(AngleOfPlot(tmp.Y, tmp.Z))
'            .Y = AngleRestrict(AngleOfPlot(tmp.Z, tmp.X))
'            .Z = AngleRestrict(AngleOfPlot(tmp.X, tmp.Y))
'            Set tmp = Nothing
'        End If
'    End With

    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim A As Double
    Dim b As Double
    Dim C As Double
    
    X = p.X
    Y = p.Y
    Z = p.Z
    
    sig_RotationMethod X, Y, Z, A, b, C
    
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        .X = Z
        .Y = b
        .Z = C
    End With
End Function

'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As New Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = Point
'            .X = AngleRestrict(AngleOfPlot(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfPlot(MakePoint(tmp.Z, tmp.X, tmp.Y)))
'            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
'            .Z = AngleRestrict(AngleOfPlot(MakePoint(tmp.X, tmp.Y, tmp.Z)))
'            Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Z))
'            Set tmp = Nothing
'        End If
'    End With
'End Function

'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As New Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = Point
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
'            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
'            .Z = AngleRestrict(AngleOfCoord(MakePoint(tmp.X, tmp.Y, tmp.Z)))
'            Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Z))
'            Set tmp = Nothing
'        End If
'    End With
'End Function

'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As New Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = Point
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
'            Set tmp = Nothing
'        End If
'    End With
'End Function

'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
''            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
''            .Z = AngleRestrict(AngleOfCoord(MakePoint(tmp.X, tmp.Y, tmp.Z)))
'            Set tmp = Nothing
'        End If
'    End With
'End Function

'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################


Public Function VectorRotateAxis2(ByRef Point As Point, ByRef Angles As Point) As Point
  '  Set VectorRotateAxis2 = VectorRotateAxis(Point, Angles)

    
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim A As Double
    Dim b As Double
    Dim C As Double
    
    X = Point.X
    Y = Point.Y
    Z = Point.Z
    

    
        
    sig_RotationMethod X, Y, Z, A, b, C

    X = Point.X
    Y = Point.Y
    Z = Point.Z
    
    A = A + Angles.X
    b = b + Angles.Y
    C = C + Angles.Z

    sig_RotationMethod X, Y, Z, A, b, C

    Set VectorRotateAxis2 = New Point
    With VectorRotateAxis2
        .X = X
        .Y = Y
        .Z = Z
    End With
    
End Function

Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
    Set VectorRotateAxis = VectorRotateAxis2(Point, Angles)
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.Y, tmp.Z, tmp.X), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.Z, tmp.X, tmp.Y), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
End Function
'
'Public Function VectorRotateAxis2(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis2 = tmp
'    Set tmp = Nothing
'End Function


'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As New Point
'    Set VectorRotateAxis = New Point
'    With VectorRotateAxis
'        .Y = Cos(Angles.X) * Point.Y - Sin(Angles.X) * Point.Z
'        .Z = Sin(Angles.X) * Point.Y + Cos(Angles.X) * Point.Z
'        tmp.X = Point.X
'        tmp.Y = .Y
'        tmp.Z = .Z
'        .X = Sin(Angles.Y) * tmp.Z + Cos(Angles.Y) * tmp.X
'        .Z = Cos(Angles.Y) * tmp.Z - Sin(Angles.Y) * tmp.X
'        tmp.X = .X
'        .X = Cos(Angles.Z) * tmp.X - Sin(Angles.Z) * tmp.Y
'        .Y = Sin(Angles.Z) * tmp.X + Cos(Angles.Z) * tmp.Y
'    End With
'End Function
'
'Public Function VectorRotateAxis3(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As New Point
'    Set VectorRotateAxis3 = New Point
'    With VectorRotateAxis3
'        .Y = (Cos(Angles.X) * Point.Y - Sin(Angles.X) * Point.Z)
'        .Z = (Sin(Angles.X) * Point.Y + Cos(Angles.X) * Point.Z)
'        tmp.X = Point.X
'        tmp.Y = .Y
'        tmp.Z = .Z
'        .X = (Sin(Angles.Y) * tmp.Z + Cos(Angles.Y) * tmp.X)
'        .Z = (Cos(Angles.Y) * tmp.Z - Sin(Angles.Y) * tmp.X)
'        tmp.X = .X
'        .X = (Cos(Angles.Z) * tmp.X - Sin(Angles.Z) * tmp.Y)
'        .Y = (Sin(Angles.Z) * tmp.X + Cos(Angles.Z) * tmp.Y)
'    End With
'End Function
'
'
'Public Function VectorRotateAxis4(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = Point
'    If Abs(Angles.Y) > Abs(Angles.X) And Abs(Angles.Y) > Abs(Angles.Z) And (Angles.Y <> 0) Then
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'        Set tmp = VectorRotateAxis4(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, 0, Angles.Z))
'    ElseIf Abs(Angles.X) > Abs(Angles.Y) And Abs(Angles.X) > Abs(Angles.Z) And (Angles.X <> 0) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateAxis4(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(0, Angles.Y, Angles.Z))
'    ElseIf Abs(Angles.Z) > Abs(Angles.Y) And Abs(Angles.Z) > Abs(Angles.X) And (Angles.Z <> 0) Then
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateAxis4(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, Angles.Y, 0))
'    End If
'    Set VectorRotateAxis4 = tmp
'    Set tmp = Nothing
'End Function


'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = Point
'    If Abs(Angles.Y) > Abs(Angles.X) And Abs(Angles.Y) > Abs(Angles.Z) And (Angles.Y <> 0) Then
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, 0, Angles.Z))
'    ElseIf Abs(Angles.X) > Abs(Angles.Y) And Abs(Angles.X) > Abs(Angles.Z) And (Angles.X <> 0) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(0, Angles.Y, Angles.Z))
'    ElseIf Abs(Angles.Z) > Abs(Angles.Y) And Abs(Angles.Z) > Abs(Angles.X) And (Angles.Z <> 0) Then
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, Angles.Y, 0))
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

Public Function VectorRotateAxis5(ByRef Point As Point, ByRef Angles As Point) As Point

    Dim matRoll As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matMat As D3DMATRIX

    D3DXMatrixIdentity matRoll
    D3DXMatrixIdentity matYaw
    D3DXMatrixIdentity matPitch
    D3DXMatrixIdentity matMat

    D3DXMatrixRotationX matPitch, AngleConvertWinToDX3DX(Angles.X)
    D3DXMatrixMultiply matMat, matPitch, matMat

    D3DXMatrixRotationY matYaw, AngleConvertWinToDX3DY(Angles.Y)
    D3DXMatrixMultiply matMat, matYaw, matMat

    D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3DZ(Angles.Z)
    D3DXMatrixMultiply matMat, matRoll, matMat

    Dim vout As D3DVECTOR
    D3DXVec3TransformCoord vout, ToVector(Point), matMat

    Set VectorRotateAxis5 = New Point
    With VectorRotateAxis5
        .X = vout.X
        .Y = vout.Y
        .Z = vout.Z
    End With
    
    sig_RotationMethod Point.X, Point.Y, Point.Z, Angles.X, Angles.Y, Angles.Z
    
End Function




'
'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

Private Function sig_Hypotenus(ByVal X As Double, ByVal Y As Double) As Double
    sig_Hypotenus = ((X ^ 2) + (Y ^ 2)) ^ (1 / 2)
End Function

Private Function sig_Sine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If Not H = 0 Then
        If X < Y Then
            sig_Sine = X / H
        Else
            sig_Sine = Y / H
        End If
    End If
End Function
Private Function sig_Cosine(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If Not H = 0 Then
        If X < Y Then
            sig_Cosine = Y / H
        Else
            sig_Cosine = X / H
        End If
    End If
End Function
Private Function sig_Tangent(ByVal X As Double, ByVal Y As Double, ByVal H As Double) As Double
    If X < Y Then
        If Not Y = 0 Then
            sig_Tangent = X / Y
        End If
    Else
        If Not X = 0 Then
            sig_Tangent = Y / X
        End If
    End If
End Function

Public Function sig_AngleOfPlot(ByVal pX As Double, ByVal pY As Double) As Double
    Dim X As Double
    Dim Y As Double
    X = Round(pX, 12)
    Y = Round(pY, 12)
    If (X = 0) Then
        If (Y > 0) Then
            sig_AngleOfPlot = 180
        ElseIf (Y < 0) Then
            sig_AngleOfPlot = 360
        End If
    ElseIf (Y = 0) Then
        If (X > 0) Then
            sig_AngleOfPlot = 90
        ElseIf (X < 0) Then
            sig_AngleOfPlot = 270
        End If
    Else
        If ((X > 0) And (Y > 0)) Then
            sig_AngleOfPlot = (90 * RADIAN)
        ElseIf ((X < 0) And (Y > 0)) Then
            sig_AngleOfPlot = (180 * RADIAN)
        ElseIf ((X < 0) And (Y < 0)) Then
            sig_AngleOfPlot = (270 * RADIAN)
        ElseIf ((X > 0) And (Y < 0)) Then
            sig_AngleOfPlot = (360 * RADIAN)
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
            sig_AngleOfPlot = sig_AngleOfPlot + (PI / 4)
        End If
        sig_AngleOfPlot = ((sig_AngleOfPlot + Angle) * DEGREE)
    End If
End Function


Private Sub sig_RotationMethod(ByRef X As Double, ByRef Y As Double, ByRef Z As Double, ByRef A As Double, ByRef b As Double, ByRef C As Double)

    Dim H1 As Double
    Dim H2 As Double
    Dim H3 As Double
    
    Dim S As Double
    Dim CS As Double
    Dim CC As Double
    Dim SC As Double
    Dim CSC As Double
    
    H1 = sig_Hypotenus(X, Y)
    H2 = sig_Hypotenus(Y, Z)
    H3 = sig_Hypotenus(Z, X)

    A = sig_AngleOfPlot(X, Y)
    b = sig_AngleOfPlot(Y, Z)
    C = sig_AngleOfPlot(Z, X)
    
    S = (X * sig_Sine(X, Y, H1))
    CS = (Y * sig_Sine(X, Y, H1))
    CC = (-(S / 2) + ((Y * sig_Cosine(Y, Z, H2)) + (X * sig_Cosine(Y, Z, H2))) - (S / 2))

    X = ((-X + CS) + X)
    Y = ((-Y + CC) + Y)
    Z = ((-Z + S) + Z)

    SC = -((sig_Tangent(X, Y, H1) / 2) - sig_Tangent(Y, Z, H3))
    CSC = sig_Tangent(Z, Y, H3) - (sig_Tangent(Z, X, H2) / 2)

    X = (X + (CSC * 2))
    Y = (Y + (((SC / 2) + (CSC / 2)) * 2))
    Z = ((((Z / 2) * 3) + (CSC - (SC / 2)) - (Z - (SC - (CSC / 2)))) * 2)
    
End Sub


'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorRotateX(ByRef Point As Point, ByVal Angle As Double) As Point
    Set VectorRotateX = MakePoint(Point.X, Point.Y, Point.Z)
   ' If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Double
    Dim SinPhi   As Double
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    With VectorRotateX
        .Z = Point.Z * CosPhi - Point.Y * SinPhi
        .Y = Point.Z * SinPhi + Point.Y * CosPhi
        .X = Point.X
    End With
End Function

Public Function VectorRotateY(ByRef Point As Point, ByVal Angle As Double) As Point
    Set VectorRotateY = MakePoint(Point.X, Point.Y, Point.Z)
    'If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Double
    Dim SinPhi   As Double
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    With VectorRotateY
        .X = Point.X * CosPhi - Point.Z * SinPhi
        .Z = Point.X * SinPhi + Point.Z * CosPhi
        .Y = Point.Y
    End With
End Function

Public Function VectorRotateZ(ByRef Point As Point, ByVal Angle As Double) As Point
    Set VectorRotateZ = MakePoint(Point.X, Point.Y, Point.Z)
    'If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Double
    Dim SinPhi   As Double
    CosPhi = Cos(Angle)
    SinPhi = Sin(Angle)
    With VectorRotateZ
        .X = Point.X * CosPhi - Point.Y * SinPhi
        .Y = Point.X * SinPhi + Point.Y * CosPhi
        .Z = Point.Z
    End With
End Function


'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################

Public Function VectorSlope(ByRef p1 As Point, ByRef p2 As Point) As Double
    'this returns the slope FACTOR form, not the literal slope, for instance all perfect
    'diagnals, horizontal and vertical will return a 1, no negatives are returned. ONLY
    'if the points equal to each other will the return be a zero, (rise over run rule)
    VectorSlope = VectorRun(p1, p2) 'horizontal travel
    If (VectorSlope <> 0) Then 'slope is defined as rise over run, rise is vertical travel
        VectorSlope = Round((VectorRise(p1, p2) / VectorSlope), 6)
        If (VectorSlope = 0) Then VectorSlope = -CInt(Not ((p1.X = p2.X) And (p1.Y = p2.Y) And (p1.Z = p2.Z)))
    ElseIf VectorRise(p1, p2) <> 0 Then
        VectorSlope = 1
    End If
End Function

Public Function VectorRise(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Double
    VectorRise = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
End Function
Public Function VectorRun(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Double
    VectorRun = DistanceEx(MakePoint(p1.X, 0, p1.Z), MakePoint(p2.X, 0, p2.Z))
End Function


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
Public Function AngleConvertWinToDX3DX(ByVal Angle As Double) As Double
    AngleConvertWinToDX3DX = AngleRestrict(Angle) '[(((360 - Abs(Angle * DEGREE)) * Sign(Angle * DEGREE)) * RADIAN))
End Function

Public Function AngleConvertWinToDX3DY(ByVal Angle As Double) As Double
    AngleConvertWinToDX3DY = AngleRestrict(Angle) '[(((360 - Abs(Angle * DEGREE)) * Sign(Angle * DEGREE)) * RADIAN))
End Function

Public Function AngleConvertWinToDX3DZ(ByVal Angle As Double) As Double
    AngleConvertWinToDX3DZ = AngleRestrict(Angle) '[(((360 - Abs(Angle * DEGREE)) * Sign(Angle * DEGREE)) * RADIAN))
End Function

Public Function AngleAxisCombine(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim p3 As New Point
    Dim P4 As New Point
    Set p3 = AngleAxisRestrict(AngleAxisInvert(p1))
    Set P4 = AngleAxisRestrict(AngleAxisInvert(p2))
    
    Set AngleAxisCombine = New Point
    With AngleAxisCombine
    
        Set p3 = AngleAxisDeduction(p1, p2)
        Set P4 = AngleAxisDifference(p1, p3)
       .X = P4.X
       .Y = P4.Y
       .Z = -P4.Z
        
        
       ' .X = ((p1.X * p2.X + p3.X * P4.X + p1.X * p3.X + p2.X * P4.X) ^ (1 / 4))
       ' .Y = ((p1.Y * p2.Y + p3.Y * P4.Y + p1.Y * p3.Y + p2.Y * P4.Y) ^ (1 / 4))
       ' .z = ((p1.z * p2.z + p3.z * P4.z + p1.z * p3.z + p2.z * P4.z) ^ (1 / 4))
        
        Set AngleAxisCombine = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
End Function

Public Function AngleAxisDifference(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim d1 As Point
    Dim d2 As Point
    Set d1 = AngleAxisRestrict(MakePoint(p1.X, p1.Y, p1.Z))
    Set d2 = AngleAxisRestrict(MakePoint(p2.X, p2.Y, p2.Z))
    
    d1.X = d1.X * DEGREE
    d1.Y = d1.Y * DEGREE
    d1.Z = d1.Z * DEGREE
    
    d2.X = d2.X * DEGREE
    d2.Y = d2.Y * DEGREE
    d2.Z = d2.Z * DEGREE
    
    Dim c1 As Double
    Dim C2 As Double
    
    Set AngleAxisDifference = New Point
    With AngleAxisDifference
        c1 = Large(d1.X, d2.X)
        C2 = Least(d1.X, d2.X)
        .X = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        c1 = Large(d1.Y, d2.Y)
        C2 = Least(d1.Y, d2.Y)
        .Y = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        c1 = Large(d1.Z, d2.Z)
        C2 = Least(d1.Z, d2.Z)
        .Z = Least(((360 - c1) + C2), (c1 - C2)) * RADIAN
        
        Set AngleAxisDifference = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
    
    Set d1 = Nothing
    Set d2 = Nothing
End Function

Public Function AngleAxisSubtraction(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim d1 As Point
    Dim d2 As Point
    Set d1 = AngleAxisRestrict(MakePoint(p1.X, p1.Y, p1.Z))
    Set d2 = AngleAxisRestrict(MakePoint(p2.X, p2.Y, p2.Z))
    
    d1.X = d1.X * DEGREE
    d1.Y = d1.Y * DEGREE
    d1.Z = d1.Z * DEGREE
    
    d2.X = d2.X * DEGREE
    d2.Y = d2.Y * DEGREE
    d2.Z = d2.Z * DEGREE
    
    Dim c1 As Double
    Dim C2 As Double
    
    Set AngleAxisSubtraction = New Point
    With AngleAxisSubtraction
        .X = (Large(d1.X, d2.X) - Least(d1.X, d2.X)) * RADIAN
        
        .Y = (Large(d1.Y, d2.Y) - Least(d1.Y, d2.Y)) * RADIAN
        
        .Z = (Large(d1.Z, d2.Z) - Least(d1.Z, d2.Z)) * RADIAN
        
        Set AngleAxisSubtraction = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
    Set d1 = Nothing
    Set d2 = Nothing
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
Public Function ValueInfluence(ByVal Final As Double, ByVal Current As Double, Optional ByVal Amount As Double = 0.001, _
                                Optional ByVal factor As Double = 1, Optional ByVal SnapRange As Double = 0) As Double

    If (Not ValueSnapCheck(Final, Current, SnapRange)) Then
        Dim N As Double
        N = Large(Final, Current) - Least(Final, Current)
        If (N <= Abs(SnapRange) And Abs(SnapRange) > 0) Then
            ValueInfluence = Final
        Else
            If Current > Final Then
                If Current - Amount >= Final Then
                    ValueInfluence = Current - Amount
                Else
                    ValueInfluence = Final
                End If
            ElseIf Current < Final Then
                If Current + Amount <= Final Then
                    ValueInfluence = Current + Amount
                Else
                    ValueInfluence = Final
                End If
            End If
        End If
    Else
        ValueInfluence = Final
    End If

End Function

Public Function ValueSnapCheck(ByVal Final As Double, ByVal Current As Double, ByVal SnapRange As Double) As Boolean
    If SnapRange = 0 Or (Current = Final) Then
        ValueSnapCheck = (Current = Final)
    Else
        Dim N As Double
        N = Abs((Large(Final, Current) - Least(Final, Current)))
        If (N <= Abs(SnapRange) And Abs(SnapRange) > 0) Then
            ValueSnapCheck = True
        End If
    End If
End Function

Public Function VectorInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Double = 0.001, _
                                Optional ByVal factor As Double = 1, Optional ByVal Concurrent As Boolean = True, _
                                Optional ByVal SnapRange As Double = 0) As Point
                                
    Set VectorInfluence = VectorDisplace(Current, Final)
    With VectorInfluence
        Dim N As Point
        If Not Concurrent Then
            Set N = VectorNormalize(VectorInfluence)
            N.X = IIf(N.X = 0, 1, N.X) * 100
            N.Y = IIf(N.Y = 0, 1, N.Y) * 100
            N.Z = IIf(N.Z = 0, 1, N.Z) * 100
        Else
            Set N = New Point
            N.X = 100
            N.Y = 100
            N.Z = 100
        End If
   
        .X = ValueInfluence(Final.X, Current.X, Amount * ((VectorInfluence.X * factor) / N.X), SnapRange)
        .Y = ValueInfluence(Final.Y, Current.Y, Amount * ((VectorInfluence.Y * factor) / N.Y), SnapRange)
        .Z = ValueInfluence(Final.Z, Current.Z, Amount * ((VectorInfluence.Z * factor) / N.Z), SnapRange)
   
        Set N = Nothing
    End With
End Function

Public Function AngleInfluence(ByVal Final As Double, ByVal Current As Double, Optional ByVal Amount As Double = 0.001, _
                                Optional ByVal factor As Double = 1, Optional ByVal SnapRange As Double = 0) As Double
        
        Dim a1 As Double
        Dim a2 As Double
        If Not ValueSnapCheck(Final, Current, SnapRange) Then
            a1 = (Least(Current, Final) * DEGREE + (360 - Large(Current, Final) * DEGREE)) * RADIAN
            a2 = (Large(Current, Final) * DEGREE - Least(Current, Final) * DEGREE) * RADIAN
            If a1 < a2 Then
                AngleInfluence = ValueInfluence(a1, 0, Amount, SnapRange)
                a1 = AngleRestrict(Current + AngleInfluence)
                a2 = AngleRestrict(Current - AngleInfluence)
                If AngleInfluence <> 0 Then
                    AngleInfluence = Final
                    If Current > Final Then
                        If a1 > Final Then AngleInfluence = a1
                    ElseIf Current < Final Then
                        If a2 < Final Then AngleInfluence = a2
                    End If
                    AngleInfluence = AngleRestrict(AngleInfluence)
                End If
            ElseIf a1 > a2 Then
                AngleInfluence = ValueInfluence(a2, 0, Amount, SnapRange)
                a1 = AngleRestrict(Current - AngleInfluence)
                a2 = AngleRestrict(Current + AngleInfluence)
                If AngleInfluence <> 0 Then
                    AngleInfluence = Final
                    If Current > Final Then
                        If a1 > Final Then AngleInfluence = a1
                    ElseIf Current < Final Then
                        If a2 < Final Then AngleInfluence = a2
                    End If
                    AngleInfluence = AngleRestrict(AngleInfluence)
                End If
            End If
        End If
End Function

Public Function AngleAxisInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Double = 0.001, _
                                    Optional ByVal factor As Double = 1, Optional ByVal Concurrent As Boolean = True, _
                                    Optional ByVal SnapRange As Double = 0) As Point
    
    Set AngleAxisInfluence = AngleAxisDifference(Current, Final)
    With AngleAxisInfluence
        Dim N As Point
        If Not Concurrent Then
            Set N = AngleAxisNormalize(AngleAxisInfluence)
            N.X = IIf(N.X = 0, 1, N.X) '* 100
            N.Y = IIf(N.Y = 0, 1, N.Y) ' * 100
            N.Z = IIf(N.Z = 0, 1, N.Z) '* 100
        Else
            Set N = New Point
            N.X = 0.01 '100
            N.Y = 0.01 '100
            N.Z = 0.01 ' 100
        End If
        
        .X = AngleInfluence(Final.X, Current.X, Amount, ((.X * factor) / N.X), SnapRange)
        .Y = AngleInfluence(Final.Y, Current.Y, Amount, ((.Y * factor) / N.Y), SnapRange)
        .Z = AngleInfluence(Final.Z, Current.Z, Amount, ((.Z * factor) / N.Z), SnapRange)
        
        Set N = Nothing
    End With
End Function


Public Function AngleAxisInbetween(ByRef ZeroPercent As Point, ByRef OneHundred As Point, Optional ByVal DecimalPercent As Double = 0.5) As Point

    Dim d1 As Point
    Dim d2 As Point

    Set d1 = AngleAxisRestrict(MakePoint(ZeroPercent.X, ZeroPercent.Y, ZeroPercent.Z))
    Set d2 = AngleAxisRestrict(MakePoint(OneHundred.X, OneHundred.Y, OneHundred.Z))
    
    d1.X = d1.X * DEGREE
    d1.Y = d1.Y * DEGREE
    d1.Z = d1.Z * DEGREE
    
    d2.X = d2.X * DEGREE
    d2.Y = d2.Y * DEGREE
    d2.Z = d2.Z * DEGREE
    
    Dim c1 As Double
    Dim C2 As Double
    
    Set AngleAxisInbetween = New Point
    With AngleAxisInbetween
        c1 = Large(d1.X, d2.X)
        C2 = Least(d1.X, d2.X)
        If (c1 - C2) <= ((360 - C2) + c1) Then
            .X = ((c1 - C2) * DecimalPercent)
            If d1.X = c1 Then
                .X = c1 + .X
            Else
                .X = C2 - .X
            End If
        Else
            .X = (((360 - C2) + c1) * DecimalPercent)
            If d1.X = C2 Then
                .X = C2 + .X
            Else
                .X = c1 - .X
            End If
        End If
        .X = .X * RADIAN
        

        
        c1 = Large(d1.Y, d2.Y)
        C2 = Least(d1.Y, d2.Y)
        If (c1 - C2) <= ((360 - C2) + c1) Then
            .Y = ((c1 - C2) * DecimalPercent)
            If d1.Y = c1 Then
                .Y = c1 + .Y
            Else
                .Y = C2 - .Y
            End If
        Else
            .Y = (((360 - C2) + c1) * DecimalPercent)
            If d1.Y = C2 Then
                .Y = C2 + .Y
            Else
                .Y = c1 - .Y
            End If
        End If
        .Y = .Y * RADIAN
        
        
        c1 = Large(d1.Z, d2.Z)
        C2 = Least(d1.Z, d2.Z)
        If (c1 - C2) <= ((360 - C2) + c1) Then
            .Z = ((c1 - C2) * DecimalPercent)
            If d1.Z = c1 Then
                .Z = c1 + .Z
            Else
                .Z = C2 - .Z
            End If
        Else
            .Z = (((360 - C2) + c1) * DecimalPercent)
            If d1.Z = C2 Then
                .Z = C2 + .Z
            Else
                .Z = c1 - .Z
            End If
        End If
        .Z = .Z * RADIAN
                
        
        Set AngleAxisInbetween = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
    
    
    Set d1 = Nothing
    Set d2 = Nothing

End Function

Public Function AngleAxisPercentOf(ByRef AngleAxis As Point, ByVal DecimalPercent As Double) As Point

    Set AngleAxisPercentOf = AngleAxisRestrict(MakePoint(AngleAxis.X, AngleAxis.Y, AngleAxis.Z))
    

    With AngleAxisPercentOf
        
        .X = .X * DEGREE
        .Y = .Y * DEGREE
        .Z = .Z * DEGREE

        .X = .X * DecimalPercent * RADIAN
        .Y = .Y * DecimalPercent * RADIAN
        .Z = .Z * DecimalPercent * RADIAN
        
        Set AngleAxisPercentOf = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
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

Public Function VectorDotProduct(ByRef p1 As Point, ByRef p2 As Point) As Double
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

Public Function VectorRootBy(ByRef p1 As Point, ByVal Power As Double) As Point
    Set VectorRootBy = New Point
    With VectorRootBy
        .X = p1.X ^ (1 / Power)
        .Y = p1.Y ^ (1 / Power)
        .Z = p1.Z ^ (1 / Power)
    End With
End Function

Public Function CrossProductLength(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
    CrossProductLength = ((p1.X - p2.X) * (p2.Y - p2.Y) - (p1.Y - p2.Y) * (p2.Z - p2.Z) - (p1.Z - p2.Z) * (p2.X - p2.X))
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

Public Function VectorQuantify(ByRef p1 As Point) As Double
    VectorQuantify = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.Z))
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

Public Function VectorMultiplyBy(ByRef p1 As Point, ByVal N As Double) As Point
    Set VectorMultiplyBy = New Point
    With VectorMultiplyBy
        .X = (p1.X * N)
        .Y = (p1.Y * N)
        .Z = (p1.Z * N)
    End With
End Function

Public Function VectorExponential(ByRef p1 As Point, ByVal N As Double) As Point
    Set VectorExponential = New Point
    With VectorExponential
        .X = (p1.X ^ N)
        .Y = (p1.Y ^ N)
        .Z = (p1.Z ^ N)
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
Public Function AngleAxisNormalize(ByRef p1 As Point) As Point
'    Set AngleAxisNormalize = New Point
'    With AngleAxisNormalize
'        .Z = (AngleRestrict(p1.X) + AngleRestrict(p1.Y) + AngleRestrict(p1.Z))
'        If .Z <> 0 Then
'            .Z = 1 / 720
'            .X = (p1.X * .Z)
'            .Y = (p1.Y * .Z)
'            .Z = (p1.Z * .Z)
'        End If
'    End With
'
'    Set AngleAxisNormalize = New Point
'    With AngleAxisNormalize
'        .Z = (AngleRestrict(p1.X) + AngleRestrict(p1.Y) + AngleRestrict(p1.Z)) '/ (360 * RADIAN)
'        If .Z <> 0 Then
'            .Z = 1 / 360
'            .X = (p1.X * .Z)
'            .Y = (p1.Y * .Z)
'            .Z = (p1.Z * .Z)
'        End If
'    End With
    
    Set AngleAxisNormalize = New Point
    With AngleAxisNormalize
        .Z = (AngleRestrict(p1.X) + AngleRestrict(p1.Y) + AngleRestrict(p1.Z)) '/ (360 * RADIAN)
        If .Z <> 0 Then
            .Z = 1 / 360
            .X = (p1.X * .Z)
            .Y = (p1.Y * .Z)
            .Z = (p1.Z * .Z)
        End If
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
Public Function VectorSign(ByVal p1 As Point) As Point
    Set VectorSign = New Point
    With VectorSign
        If Abs(p1.X) >= Abs(p1.Y) And Abs(p1.X) >= Abs(p1.Z) Then
            .X = IIf(p1.X > 0, 1, IIf(p1.X < 0, -1, 0))
        End If
        If Abs(p1.Y) >= Abs(p1.Z) And Abs(p1.Y) >= Abs(p1.X) Then
            .Y = IIf(p1.Y > 0, 1, IIf(p1.Y < 0, -1, 0))
        End If
        If Abs(p1.Z) >= Abs(p1.X) And Abs(p1.Z) >= Abs(p1.Y) Then
            .Z = IIf(p1.Z > 0, 1, IIf(p1.Z < 0, -1, 0))
        End If
    End With
End Function
Public Function VectorMagnitude(ByVal p1 As Point) As Double
    VectorMagnitude = (p1.X * p1.X + p1.Y * p1.Y + p1.Z * p1.Z)
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

Public Function VectorDivideBy(ByRef p1 As Point, ByVal N As Double) As Point
    Set VectorDivideBy = New Point
    With VectorDivideBy
        .X = (p1.X / N)
        .Y = (p1.Y / N)
        .Z = (p1.Z / N)
    End With
End Function
Public Function VectorDivision(ByRef p1 As Point, ByRef p2 As Point) As Point
    Set VectorDivision = New Point
    With VectorDivision
        If p2.X <> 0 Then
            .X = (p1.X / p2.X)
        End If
        If p2.Y <> 0 Then
            .Y = (p1.Y / p2.Y)
        End If
        If p2.Z <> 0 Then
            .Z = (p1.Z / p2.Z)
        End If
    End With
End Function
Public Function VectorIsNormal(ByRef p1 As Point) As Boolean
    'returns if a point provided is normalized, to the best of ability
    VectorIsNormal = (Round(Abs(p1.X) + Abs(p1.Y) + Abs(p1.Z), 0) = 1) 'first kind is the absolute of all values equals one
    If VectorIsNormal Then Exit Function
    VectorIsNormal = (DistanceEx(MakePoint(0, 0, 0), p1) = 1)  'another is the total length of vector is one
    If VectorIsNormal Then Exit Function
    'another is if any value exists non zero as well as adding up in any non specific arrangement cancels to zero, as has one
    VectorIsNormal = ((p1.X <> 0 Or p1.Y <> 0 Or p1.Z <> 0) And (( _
        ((p1.X + p1.Y + p1.Z) = 0) Or ((p1.Y + p1.Z + p1.X) = 0) Or ((p1.Z + p1.X + p1.Y) = 0) Or _
        ((p1.X + p1.Z + p1.Y) = 0) Or ((p1.Z + p1.Y + p1.X) = 0) Or ((p1.Y + p1.X + p1.Z) = 0) _
        )))
    If VectorIsNormal Then Exit Function
    'triangle's normal, only the sides are expressed upon each axis
    VectorIsNormal = ((((p1.X - p1.Y) + p1.Z) + ((p1.Y - p1.Z) + p1.X) + ((p1.Z - p1.X) + p1.Y)) = 1)
    If VectorIsNormal Then Exit Function
    Dim tmp As Double
    'another is a reflection test and check if it falls with in -1 to 1 for triangle normals
    'reflection is 27 groups of three arithmitic (-1+(2-3)) and by the third group, the groups
    'reflect the same (-g+(g-g)) which are sub groups of lines of three groups doing the same
    tmp = -((-(-p1.X + (p1.Y - p1.Z)) + ((-p1.Y + (p1.Z - p1.X)) - (-p1.Z + (p1.X - p1.Y)))) + _
        ((-p1.Y + (p1.Z - p1.X)) + ((-p1.Z + (p1.X - p1.Y)) - (-p1.X + (p1.Y - p1.Z))) - _
        (-p1.Z + (p1.X - p1.Y)) + ((-p1.X + (p1.Y - p1.Z)) - (-p1.Y + (p1.Z - p1.X))))) + ( _
        ((-(-p1.Y + (p1.X - p1.Z)) + ((-p1.X + (p1.Z - p1.Y)) - (-p1.Z + (p1.Y - p1.X)))) + _
        ((-p1.X + (p1.Z - p1.Y)) + ((-p1.Z + (p1.Y - p1.X)) - (-p1.Y + (p1.X - p1.Z))) - _
        (-p1.Z + (p1.Y - p1.X)) + ((-p1.Y + (p1.X - p1.Z)) - (-p1.X + (p1.Z - p1.Y))))) - _
        ((-(-p1.Z + (p1.Y - p1.X)) + ((-p1.Y + (p1.X - p1.Z)) - (-p1.X + (p1.Z - p1.Y)))) + _
        ((-p1.Y + (p1.X - p1.Z)) + ((-p1.X + (p1.Z - p1.Y)) - (-p1.Z + (p1.Y - p1.X))) - _
        (-p1.X + (p1.Z - p1.Y)) + ((-p1.Z + (p1.Y - p1.X)) - (-p1.Y + (p1.X - p1.Z))))))
        '9 lines, 27 groups, 81 values, full circle, the first value (-negative, plus (second minus third))
    VectorIsNormal = ((p1.X <> 0 Or p1.Y <> 0 Or p1.Z <> 0) And (tmp >= -1 And tmp <= 1))
End Function
Public Function VectorIsSignOf(ByRef p1 As Point) As Boolean
    VectorIsSignOf = (Abs(p1.X) = 0 Or Abs(p1.X) = 1) And (Abs(p1.Y) = 0 Or Abs(p1.Y) = 1) And (Abs(p1.Z) = 0 Or Abs(p1.Z) = 1) 'sign of a vector
End Function
Public Function AbsoluteFactor(ByVal N As Double) As Double
    'returns -1 if the n is below zero, returns 1 if n is above zero, and 0 if n is zero
    AbsoluteFactor = ((-(AbsoluteValue(N - 1) - N) - (-AbsoluteValue(N + 1) + N)) * 0.5)
End Function

Public Function AbsoluteValue(ByVal N As Double) As Double
    'same as abs(), returns a number as positive quantified
    AbsoluteValue = (-((-(N * -1) * N) ^ (1 / 2) * -1))
End Function

Public Function AbsoluteWhole(ByVal N As Double) As Double
    'returns only the digits to the left of a decimal in any numerical
    'AbsoluteWhole = (AbsoluteValue(n) - (AbsoluteValue(n) - (AbsoluteValue(n) Mod (AbsoluteValue(n) + 1)))) * AbsoluteFactor(n)
    AbsoluteWhole = (N \ 1) 'is also correct
End Function

Public Function AbsoluteDecimal(ByVal N As Double) As Double
    'returns only the digits to the right of a decimal in any numerical
    AbsoluteDecimal = (AbsoluteValue(N) - AbsoluteValue(AbsoluteWhole(N))) * AbsoluteFactor(N)
End Function

Public Function AngleQuadrant(ByVal Angle As Double) As Double
    'returns the axis quadrant a radian angle falls with-in
    Angle = Angle * DEGREE
    If (Angle > 0 And Angle < 90) Or (Angle = 360) Then
        AngleQuadrant = 1
    ElseIf Angle >= 90 And Angle < 180 Then
        AngleQuadrant = 2
    ElseIf Angle >= 180 And Angle < 270 Then
        AngleQuadrant = 3
    ElseIf Angle >= 270 And Angle < 360 Then
        AngleQuadrant = 4
    End If
End Function

Public Function VectorQuadrant(ByRef p As Point) As Double
    'starts at (positive, positive) and goes clockwise
    If (p.Y > 0 And p.X >= 0) Or (p.Y >= 0 And p.X > 0) Then
        VectorQuadrant = 1
    ElseIf (p.Y <= 0 And p.X > 0) Or (p.Y < 0 And p.X >= 0) Then
        VectorQuadrant = 2
    ElseIf (p.Y < 0 And p.X <= 0) Or (p.Y <= 0 And p.X < 0) Then
        VectorQuadrant = 3
    ElseIf (p.Y >= 0 And p.X < 0) Or (p.Y > 0 And p.X <= 0) Then
        VectorQuadrant = 4
    End If
End Function

Public Function VectorOctet(ByRef p As Point) As Double
    VectorOctet = VectorQuadrant(p)
    If p.Z < 0 Then VectorOctet = VectorOctet + 4
End Function


Public Function VectorInbetween(ByRef ZeroPercent As Point, ByRef OneHundred As Point, Optional ByVal DecimalPercent As Double = 0.5) As Point

    Dim c1 As Double
    Dim C2 As Double
    
    Set VectorInbetween = New Point
    With VectorInbetween
        c1 = Large(ZeroPercent.X, OneHundred.X)
        C2 = Least(ZeroPercent.X, OneHundred.X)
        If Abs(c1 - C2) <= Abs(C2 - c1) Then
            .X = ZeroPercent.X + ((c1 - C2) * DecimalPercent)
        Else
            .X = ZeroPercent.X + ((C2 - c1) * DecimalPercent)
        End If

        
        c1 = Large(ZeroPercent.Y, OneHundred.Y)
        C2 = Least(ZeroPercent.Y, OneHundred.Y)
        If Abs(c1 - C2) <= Abs(C2 - c1) Then
            .Y = ZeroPercent.Y + ((c1 - C2) * DecimalPercent)
        Else
            .Y = ZeroPercent.Y + ((C2 - c1) * DecimalPercent)
        End If
        
        
        c1 = Large(ZeroPercent.Z, OneHundred.Z)
        C2 = Least(ZeroPercent.Z, OneHundred.Z)
        If Abs(c1 - C2) <= Abs(C2 - c1) Then
            .Z = ZeroPercent.Z + ((c1 - C2) * DecimalPercent)
        Else
            .Z = ZeroPercent.Z + ((C2 - c1) * DecimalPercent)
        End If

    End With

End Function


Public Function AbsoluteInvert(ByVal Value As Long, Optional ByVal Whole As Long = 100, Optional ByVal Unit As Long = 1)
    'returns the inverted value of a whole conprised of unit measures, AbsoluteInvert(25) returns 75
    'another example: AbsoluteInvert(0, 16777216) returns the negative of black 0, which is 16777216
    AbsoluteInvert = -(Whole / Unit) + -(Value / Unit) + ((Whole / Unit) * 2)
End Function

Public Function Lerp(ByVal A As Double, ByVal b As Double, ByVal t As Double) As Double
    Lerp = A + (b - A) * t
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





'Public Function TriangleOpposite(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Double
'    Dim l1 As Double
'    Dim l2 As Double
'    TriangleOpposite = DistanceEx(p1, p2)
'    l1 = DistanceEx(p2, p3)
'    l1 = DistanceEx(p3, p1)
'    If l1 < TriangleOpposite Then
'        If l2 < l1 And l2 < TriangleOpposite Then
'            TriangleOpposite = ((TriangleOpposite ^ 2) - (l1 ^ 2)) ^ (1 / 2)
'        ElseIf l2 > TriangleOpposite Then
'            TriangleOpposite = ((l2 ^ 2) - (Large(l1, TriangleOpposite) ^ 2)) ^ (1 / 2)
'        Else
'            TriangleOpposite = ((TriangleOpposite ^ 2) - (l2 ^ 2)) ^ (1 / 2)
'        End If
'    Else
'        TriangleOpposite = ((l1 ^ 2) - (TriangleOpposite ^ 2)) ^ (1 / 2)
'    End If
'End Function
'
'Public Function TriangleAdjacent(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Double
'    'provide the Hypotenuse as line p1-p2, or all points to the triangle
'    TriangleAdjacent = DistanceEx(p1, p2)
'    If Not p3 Is Nothing Then
'        Dim l1 As Double
'        Dim l2 As Double
'        l1 = DistanceEx(p2, p3)
'        l2 = DistanceEx(p3, p1)
'        If TriangleAdjacent < l1 Xor TriangleAdjacent < l2 Then
'            If TriangleAdjacent < l1 Then
'                If l1 > l2 Then
'                    TriangleAdjacent = ((TriangleAdjacent ^ 2) - (l2 ^ 2)) ^ (1 / 2)
'                Else
'                    TriangleAdjacent = ((TriangleAdjacent ^ 2) - (l1 ^ 2)) ^ (1 / 2)
'                End If
'            Else
'                If l1 > l2 Then
'                    If l2 > TriangleAdjacent Then
'                        TriangleAdjacent = ((l1 ^ 2) - (TriangleAdjacent ^ 2)) ^ (1 / 2)
'                    Else
'                        TriangleAdjacent = ((l1 ^ 2) - (l2 ^ 2)) ^ (1 / 2)
'                    End If
'                Else
'                    If l1 > TriangleAdjacent Then
'                        TriangleAdjacent = ((l2 ^ 2) - (TriangleAdjacent ^ 2)) ^ (1 / 2)
'                    Else
'                        TriangleAdjacent = ((l2 ^ 2) - (l1 ^ 2)) ^ (1 / 2)
'                    End If
'                End If
'            End If
'        End If
'    Else
'        TriangleAdjacent = (((TriangleAdjacent ^ 2) / 2) ^ (1 / 2))
'    End If
'End Function
'
'Public Function TriangleHypotenuse(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Double
'    TriangleHypotenuse = DistanceEx(p1, p2)
'    If p3 Is Nothing Then
'        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (TriangleHypotenuse ^ 2)) ^ (1 / 2)
'    Else
'        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (DistanceEx(p2, p3) ^ 2)) ^ (1 / 2)
'    End If
'End Function
'
'
'Public Function InvSin(number As Double) As Double
'    InvSin = -number * number + 1
'    If InvSin > 0 Then
'        InvSin = Sqr(InvSin)
'        If InvSin <> 0 Then InvSin = Atn(number / InvSin)
'    Else
'        InvSin = 0
'    End If
'End Function
'Function ArcSin(X As Double) As Double
'    If Abs(X) <> 0 Then
'         If Abs((1 - Sqr(Abs(X)))) <> 0 Then
'            ArcSin = Atn2(X / Sqr(Abs(1 - Sqr(Abs(X)))))
'        End If
'    End If
'End Function
'Public Function PlotOfAngle(ByVal RadiusLength As Double, ByVal AngleInRadian As Double) As Point
'    Set PlotOfAngle = New Point
'    With PlotOfAngle
'
'        Dim p2 As New Point
'        Dim d As New Point 'the point we will modify for the finish
'        Dim Angle As Double 'the angle we'll modify for the finish
'
'        d.X = RadiusLength
'        d.Y = RadiusLength
'
'        'restrict the angle to the radian spectrum, and make it degrees
'        'so that, we shoudn't end up with a 0 or below, nor over 360
'        Angle = Round(AngleInRadian * DEGREE, 0)
'        Do While Angle > 360
'            Angle = Angle - 360
'        Loop
'        Do While Angle <= 0
'            Angle = Angle + 360
'        Loop
'
'        'whole value angles of 90 degrees increments localize to
'        'a quadrant quite easy, and need to be excluded in rotation
'        'to make up for angle diagnal percentages aren't at a curve
'        If Angle = 360 Or Angle = 0 Then 'check for zero anyway
'            d.X = 0
'        ElseIf Angle = 270 Then
'            d.X = -d.X
'            d.Y = 0
'        ElseIf Angle = 180 Then
'            d.Y = -d.Y
'            d.X = 0
'        ElseIf Angle = 90 Then
'            d.X = d.X
'            d.Y = 0
'        ElseIf Angle Mod 45 = 0 Then
'            'if the angle is still an increment of 45,
'            'it must be a diagnal, which is easily so
'            d.X = ((((d.X ^ 2) * 2) ^ (1 / 2)) / 2)
'            d.Y = ((((d.Y ^ 2) * 2) ^ (1 / 2)) / 2)
'            'localize to the quadrant
'            If Angle = 315 Then
'                d.X = -d.X
'            ElseIf Angle = 225 Then
'                d.X = -d.X
'                d.Y = -d.Y
'            ElseIf Angle = 135 Then
'                d.Y = -d.Y
'            End If
'        Else
'
'            Dim A As Double 'a+aa is a whole 100% along the x axis of the angles
'            Dim aa As Double 'percent equaling aa with an equalteral right traingle
'            Dim b As Double 'b+bb is a whole 100% along the y axis of the angles
'            Dim bb As Double 'percent equaling bb with an equalteral right traingle
'            Dim m As Double 'slope of the line of the unit circle initial values
'            Dim an As Double 'angle in definition with the -PI*2, PI, 0, PI, and PI*2 of
'            'radian spectrum (5 angles) excluded and the 0, 45 and 90 out degree spectrum
'            '(3 angles) excluded, overlapping, hence the hard values of ((5/45) / (-3/180))
'
'            Dim t As Double 'temporary
'            Dim p3 As New Point
'
'            'next localize the angle to 45 degrees
'            If Angle Mod 90 <= 45 And Angle Mod 90 > 0 Then
'                t = ((45 - (Angle - ((Angle \ 45) * 45))) / 45)
'            Else
'                t = ((Angle - ((Angle \ 45) * 45)) / 45)
'            End If
'
'            'radians range from -PI*2 to PI * 2, all R MOD PI = 0 is a invalid angle
'            '5 total for the radian spectrum and 3 total in a 90 degrees equivelent
'            'negate them in each others complment spectrum desgrees over a total 180%
'            an = (-t - (t * ((5 / 45) / (-3 / 180)))) / ((PI * 2) - (5 * RADIAN))
'
'            b = (an * RadiusLength)
'            A = (RadiusLength - b)
'
'            aa = (((A ^ 2) / 2) ^ (1 / 2))
'            bb = ((((RadiusLength ^ 2) / 2) ^ (1 / 2)) - aa)
'
'            'p2 is the coordinate if they fall upon a diagnale square like a
'            'diamond, and a unit squares points are (-1,0)-(0,1)-(1,0)-(0,-1)
'            p2.Y = (b + (A / 2)) * IIf(d.X < 0, -1, 1)
'            p2.X = (((aa ^ 2) / 2) ^ (1 / 2)) * IIf(d.Y < 0, -1, 1)
'            p2.z = 0
'
'            'get the slope and hypotenus of our p2 to find out p3 on the
'            'curve and make the unit square a unit circle, but use radius
'            m = LineSlope2D(p2)
'            If (p2.X ^ 2 + p2.Y ^ 2) <> 0 Then
'                t = (Least(p2.Y, p2.X) / ((p2.X ^ 2 + p2.Y ^ 2) ^ (1 / 2)))
'            End If
'            If t * m <> 0 Then
'                p3.X = (p3.X - ((RadiusLength - p2.X) / t * m) + p2.Y)
'                p3.Y = (p3.Y - ((RadiusLength - p2.Y) / t * m) + p2.X)
'            End If
'
'            'do a simple swap
'            t = p3.X
'            p3.X = p3.Y
'            p3.Y = t
'
'            'next just cheat.
'            Set d = DistanceSet(MakePoint(0, 0, 0), p3, RadiusLength)
'            Set p3 = Nothing
'
'            If Angle > 270 And Angle <= 360 Then
'                If Angle Mod 90 <= 45 Then
'                    t = d.X
'                    d.X = d.Y
'                    d.Y = t
'                End If
'                d.X = -Abs(d.X)
'                d.Y = Abs(d.Y)
'           ElseIf Angle > 180 And Angle <= 270 Then
'                If Angle Mod 90 > 45 Then
'                    t = d.X
'                    d.X = d.Y
'                    d.Y = t
'                End If
'                d.Y = -Abs(d.Y)
'                d.X = -Abs(d.X)
'           ElseIf Angle > 90 And Angle <= 180 Then
'                If Angle Mod 90 <= 45 Then
'                    t = d.X
'                    d.X = d.Y
'                    d.Y = t
'                End If
'                d.Y = -Abs(d.Y)
'                d.X = Abs(d.X)
'           ElseIf Angle > 0 And Angle <= 90 Then
'                If Angle Mod 90 > 45 Then
'                    t = d.X
'                    d.X = d.Y
'                    d.Y = t
'                End If
'                d.X = Abs(d.X)
'                d.Y = Abs(d.Y)
'            End If
'        End If
'
'        'set p2 x and y variables to aspect ratio
'        p2.X = 1: p2.Y = 1 'we'll go with 1:1
'
'        .X = d.X * p2.X
'        .Y = d.Y * p2.Y
'
'        Set p2 = Nothing
'        Set d = Nothing
'
'    End With
'End Function
'

'
'Public Function AngleOfWave(ByVal WaveXLo As Double, ByVal WaveYDist As Double, ByVal WaveZHi As Double) As Double
'    Const LargePI As Double = ((((PI / 4) * DEGREE) - 1) * RADIAN)
'    Const LeastPI As Double = ((((PI / 16) * DEGREE) + 2) * RADIAN)
'
'    Dim slope As Double
'    Dim WaveHype1 As Double
'    Dim WaveHype2 As Double
'
'    slope = (WaveXLo / WaveZHi)
'    WaveHype1 = (((WaveZHi ^ 2) - (WaveXLo ^ 2)) ^ (1 / 2))
'    WaveHype2 = (((WaveYDist ^ 2) - (WaveXLo ^ 2)) ^ (1 / 2))
'
'    AngleOfWave = (LargePI * slope) + (((LeastPI * slope) * (WaveHype1 / WaveYDist)) * (WaveHype2 / WaveYDist))
'End Function
'
'Public Function AnglesOfPoint(ByRef Point As Point, Optional ByRef Angles As Point) As Point
'    Static stack As Integer
'    stack = (Abs(stack) + 1) * IIf(stack < 0, -1, 1)
'    If Abs(stack) = 1 Then
'        '(1,1,1) is high noon to 45 degree sections
'        'strangely enough, in case (0,0,0), add it
'        Point.X = (Point.X + 1)
'        Point.Y = (Point.Y + 1)
'        Point.z = (Point.z + 1)
'        If Angles Is Nothing Then
'            Set Angles = New Point
'            stack = -stack
'        End If
'
'    End If
'    Set AnglesOfPoint = New Point
'    With AnglesOfPoint
'        If Abs(stack) < 5 Then
'            Dim X As Double
'            Dim Y As Double
'            Dim z As Double
'            Dim Angle As Double
'            'round them off for checking
'            '(6 is for Double precision)
'            X = Round(Point.X, 6)
'            Y = Round(Point.Y, 6)
'            z = Round(Point.z, 6)
'            If (X = 0) Then  'slope of 1
'                If (z = 0) Then
'                    'must be 360 or 180
'                    If (Y > 0) Then
'                        .z = (180 * RADIAN)
'                    ElseIf (Y < 0) Then
'                        .z = (360 * RADIAN)
'                    End If
'                Else
'                    .z = AnglesOfPoint(MakePoint(Point.Y, Point.z, Point.X), Angles).z
'                    '.z = AnglesOfPoint(Point, Angles).z
'                End If
'            ElseIf (Y = 0) Then   'slope of 0
'                If (z = 0) Then
'                    'must be 90 or 270
'                    If (X > 0) Then
'                        .z = (90 * RADIAN)
'                    ElseIf (X < 0) Then
'                        .z = (270 * RADIAN)
'                    End If
'                Else
'                    .z = AnglesOfPoint(MakePoint(Point.Y, Point.z, Point.X), Angles).z
'                    '.z = AnglesOfPoint(Point, Angles).z
'                End If
'            ElseIf (X <> 0) And (Y <> 0) Then
'                Dim slope As Double
'                Dim dist As Double
'
'                'find the larger coordinate
'
'                dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2)) 'distance
'                Angle = AngleOfWave(IIf(Abs(Point.X) > Abs(Point.Y), Abs(Point.Y), Abs(Point.X)), _
'                            dist, IIf(Abs(Point.X) > Abs(Point.Y), Abs(Point.X), Abs(Point.Y)))
'
'                '(up to the quardrant)
'                If ((X > 0) And (Y > 0)) Then
'                    .z = (90 * RADIAN)
'                ElseIf ((X < 0) And (Y > 0)) Then
'                    .z = (180 * RADIAN)
'                ElseIf ((X < 0) And (Y < 0)) Then
'                    .z = (270 * RADIAN)
'                ElseIf ((X > 0) And (Y < 0)) Then
'                    .z = (360 * RADIAN)
'                End If
'
'                If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
'                   (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
'                   'the angle for 45 to 90 is in reverse, and doesn't start at 45, but because we
'                   'are calculating a second 45 of 90, the offset (-1 not 0) is included if inverse
'                    Angle = (PI / 4) - Angle
'                    'then also add 45 to the base
'                    .z = .z + (PI / 4)
'                End If
'                'add it to the base, returing as .Z
'                .z = .z + Angle
'
'                Dim Ret As Point
'                If (z <> 0) Then 'two or less axis is one rotation
'                    Set Ret = AnglesOfPoint(MakePoint(Point.Y, Point.z, Point.X), Angles)
'                    'Set Ret = AnglesOfPoint(Point, Angles)
'                    If Abs(stack) = 2 Then
'                        .Y = Ret.z
'                    End If
'                    If Abs(stack) = 1 Then
'                        .X = Ret.Y
'                        .Y = Ret.z
'                    End If
'                    Set Ret = Nothing
'                End If
'
'                If Abs(stack) = 1 Then
'                    'reorganization
'                    Angle = .X
'                    .X = .Y
'                    .Y = .z
'                    .z = Angle
'                    Angle = .X
'                    .X = .Y
'                    .Y = .z
'                    .z = Angle
'                End If
'            End If
'
'            If Abs(stack) < 5 Then
'                If Not Angles Is Nothing Then
'
'                    Static pX As Point
'                    Static pY As Point
'                    Static pZ As Point
'                    Select Case Abs(stack)
'                        Case 4
'                            Set pZ = PlotOfAngle(dist, (.z + Angles.z))
'                            'Form1.Picture3.Circle (pZ.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - pZ.y), 8 * Screen.TwipsPerPixelX, &HC000&
'                        Case 3
'                            Set pY = PlotOfAngle(dist, (.z + Angles.Y))
'                            'Form1.Picture2.Circle (pY.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - pY.y), 8 * Screen.TwipsPerPixelX, &HC000&
'                        Case 2
'                            Set pX = PlotOfAngle(dist, (.z + Angles.X))
'                            'Form1.Picture1.Circle (pX.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - pX.y), 8 * Screen.TwipsPerPixelX, &HC000&
'                    End Select
'
'                    If Abs(stack) = 1 And dist > 0 Then
'                        Point.X = Point.X - 1
'                        Point.Y = Point.Y - 1
'                        Point.z = Point.z - 1
'                        Set Ret = New Point
'
'                        Dim l As New Point
'                        l.X = DistanceEx(MakePoint(0, 0, 0), pX)
'                        l.Y = DistanceEx(MakePoint(0, 0, 0), pY)
'                        l.z = DistanceEx(MakePoint(0, 0, 0), pZ)
'
'                        Dim S As New Point
'                        S.X = LineSlope2D(pX)
'                        S.Y = LineSlope2D(pY)
'                        S.z = LineSlope2D(pZ)
'
'                        Dim sx As Double
'                        Dim sy As Double
'                        Dim sz As Double
'                        Dim cx As Double
'                        Dim cy As Double
'                        Dim cz As Double
'                        Dim tx As Double
'                        Dim ty As Double
'                        Dim tz As Double
'
'                        sx = Least(pX.X, pX.Y) / dist
'                        sy = Least(pY.X, pY.Y) / dist
'                        sz = Least(pZ.X, pZ.Y) / dist
'                        cx = Large(pX.X, pX.Y) / dist
'                        cy = Large(pY.X, pY.Y) / dist
'                        cz = Large(pZ.X, pZ.Y) / dist
'                        If Large(pX.X, pX.Y) <> 0 Then tx = Least(pX.X, pX.Y) / Large(pX.X, pX.Y)
'                        If Large(pY.X, pY.Y) <> 0 Then ty = Least(pY.X, pY.Y) / Large(pY.X, pY.Y)
'                        If Large(pZ.X, pZ.Y) <> 0 Then tz = Least(pZ.X, pZ.Y) / Large(pZ.X, pZ.Y)
'
'                        's = (o / h)
'                        '   o = (s * h)
'                        '   h = (((o / s) ^ 2) ^ (1 / 2))
'
'                        'c = (a / h)
'                        '   a = (c * h)
'                        '   h = (((a / c) ^ 2) ^ (1 / 2))
'
'                        't = (o / a)
'                        '   o = (t * a)
'                        '   a = (((o / t) ^ 2) ^ (1 / 2))
'
'        'slope
'        'm = (y / x)
'
'        'Y-Intercept
'        'b = -((m * x) - y)
'
'        'X & Y coordinates
'        'y = ((m * x) + b)
'        'x = ((y - b) / m
'
'
'
'                        Ret.X = pX.X
'                        Ret.Y = pY.X
'
'                        Ret.z = pZ.X
'
'
'
'
''                            Form1.Picture1.Circle (Ret.X + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Ret.Y), 8 * Screen.TwipsPerPixelX, &H4000&
''                            Form1.Picture2.Circle (Ret.Y + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Ret.z), 8 * Screen.TwipsPerPixelX, &H4000&
''                            Form1.Picture3.Circle (Ret.z + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Ret.X), 8 * Screen.TwipsPerPixelX, &H4000&
'
'                        Set pX = Nothing
'                        Set pY = Nothing
'                        Set pZ = Nothing
'
'                        Set Point = Ret
'                        Set Ret = Nothing
'
'                        Point.X = Point.X + 1
'                        Point.Y = Point.Y + 1
'                        Point.z = Point.z + 1
'
'
'                    End If
'
'                End If
'            End If
'
'        End If
'
'        If Abs(stack) = 1 Then 'undo
'
'           ' .z = AngleConvertWinToDX3DZ(.z)
'
'            Point.X = (Point.X - 1)
'            Point.Y = (Point.Y - 1)
'            Point.z = (Point.z - 1)
'
'            If stack < 0 Then Set Angles = Nothing
'        End If
'
'    End With
'    stack = (Abs(stack) - 1) * IIf(stack < 0, -1, 1)
'End Function


'
'                        Dim p As Plot
'                        Select Case Abs(stack)
'                            Case 4
'
'                                Form1.Picture3.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                            Case 3
'                                Set p = PlotOfAngle(dist, Angles.z)
'                                Point.y = p.x * Cos(.z) - p.y * Sin(.z)
'                                Point.x = p.x * Sin(.z) + p.y * Cos(.z)
'                                Form1.Picture2.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                            Case 2
'                                Set p = PlotOfAngle(dist, Angles.z)
'                                Point.y = p.x * Cos(.z) - p.y * Sin(.z)
'                                Point.x = p.x * Sin(.z) + p.y * Cos(.z)
'                                Form1.Picture1.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                        End Select


'                        Dim XOld As Double
'                        Dim YOld As Double
'                        Dim ZOld As Double
'                        Dim XNew As Double
'                        Dim YNew As Double
'                        Dim ZNew As Double
'
'                        Dim SinAngleX As Double
'                        Dim CosAngleX As Double
'                        Dim TanAngleX As Double
'                        Dim SinAngleY As Double
'                        Dim CosAngleY As Double
'                        Dim TanAngleY As Double
'                        Dim SinAngleZ As Double
'                        Dim CosAngleZ As Double
'                        Dim TanAngleZ As Double
'
'                        Dim retMatrix As D3DMATRIX
'                        D3DXMatrixIdentity retMatrix
'
'                        SinAngleX = Sin(Angles.x)
'                        CosAngleX = Cos(Angles.x)
'                        TanAngleX = Tan(Angles.x)
'                        SinAngleY = Sin(Angles.y)
'                        CosAngleY = Cos(Angles.y)
'                        TanAngleY = Tan(Angles.y)
'                        SinAngleZ = Sin(Angles.z)
'                        CosAngleZ = Cos(Angles.z)
'                        TanAngleZ = Tan(Angles.z)
'
'                        Dim i As Byte
'                        For i = 0 To 2
'                            If (i = 0) Then
'                                XOld = 1
'                                YOld = retMatrix.m12
'                                ZOld = retMatrix.m13
'                            ElseIf (i = 1) Then
'                                XOld = retMatrix.m21
'                                YOld = 1
'                                ZOld = retMatrix.m23
'                            ElseIf (i = 2) Then
'                                XOld = retMatrix.m31
'                                YOld = retMatrix.m32
'                                ZOld = 1
'                            End If
'                            YNew = (YOld * CosAngleX) - (ZOld * SinAngleX)
'                            ZNew = (YOld * SinAngleX) + (ZOld * CosAngleX)
'                            XNew = XOld
'
'                            XOld = XNew
'                            YOld = YNew
'                            ZOld = ZNew
'
'                            XNew = ((XOld * CosAngleY) + (ZOld * SinAngleY))
'                            YNew = YOld
'                            ZNew = ((-1 * (XOld * SinAngleY)) + (ZOld * CosAngleY))
'
'                            XOld = XNew
'                            YOld = YNew
'                            ZOld = ZNew
'                            XNew = ((XOld * CosAngleZ) - (YOld * SinAngleZ))
'                            YNew = ((YOld * CosAngleZ) + (XOld * SinAngleZ))
'                            ZNew = ZOld
'
'                            If (i = 0) Then
'                                retMatrix.m11 = XNew
'                                retMatrix.m12 = YNew
'                                retMatrix.m13 = ZNew
'                            ElseIf (i = 1) Then
'                                retMatrix.m21 = XNew
'                                retMatrix.m22 = YNew
'                                retMatrix.m23 = ZNew
'                            ElseIf (i = 2) Then
'                                retMatrix.m31 = XNew
'                                retMatrix.m32 = YNew
'                                retMatrix.m33 = ZNew
'                            End If
'                        Next
'
'                        Dim vin As D3DVECTOR
'                        Dim vout As D3DVECTOR
'                        vin.x = Point.x - 1
'                        vin.y = Point.y - 1
'                        vin.z = Point.z - 1
'
'                        D3DXVec3TransformCoord vout, vin, retMatrix
'
'                        Point.x = vout.x + 1
'                        Point.y = vout.y + 1
'                        Point.z = vout.z + 1


'Function ArcCos(X As Double) As Double
'    If Abs(X) <> 0 Then
'         If Abs((1 - Sqr(Abs(X))) / Abs(X)) <> 0 Then
'            ArcCos = Atn2(Sqr(Abs(1 - Sqr(Abs(X))) / Abs(X)))
'        End If
'    End If
'End Function
'Function Atn2(ByVal X As Double) As Double
'
'    If Abs(-X * X + 1) <> 0 Then
'        If Sqr(Abs(-X * X + 1)) <> 0 Then
'            Atn2 = Atn(X / Abs(Sqr(Abs(-X * X + 1))))
'        End If
'    End If
'
'End Function

'Function ArcCos(X As Double) As Double
'    If Abs(X) <> 0 Then
'         If Abs((1 - Sqr(Abs(X))) / Abs(X)) <> 0 Then
'            ArcCos = Atn2(Sqr(Abs(1 - Sqr(Abs(X))) / Abs(X)))
'        End If
'    End If
'End Function
'Function Atn2(ByVal X As Double) As Double
'
'    If Abs(-X * X + 1) <> 0 Then
'        If Sqr(Abs(-X * X + 1)) <> 0 Then
'            Atn2 = Abs(Atn(X / Abs(Sqr(Abs(-X * X + 1)))))
'        End If
'    End If
'
'End Function
'Function atan3(ByVal i As Double, ByVal r As Double) As Double
'        Dim Theta As Double
'        If ((i >= 0) And (r > 0)) Then
'                Theta = Atn(Abs(i) / Abs(r)) '1st quadrant
'        ElseIf ((i >= 0) And (r < 0)) Then
'                Theta = PI - Atn(Abs(i) / Abs(r)) '2nd quadrant
'        ElseIf ((i < 0) And (r < 0)) Then
'                Theta = PI + Atn(Abs(i) / Abs(r)) '3rd quadrant
'        ElseIf ((i < 0) And (r > 0)) Then
'                Theta = 2 * PI - Atn(Abs(i) / Abs(r)) '4th quadrant
'        ElseIf (r = 0) Then
'                Theta = (PI / 2#) '90 degrees
'        End If
'        atan3 = Theta
'End Function



'        Function atan4(ByVal y As Double, ByVal x As Double) As Double
'            Const PI = 3.14159265358979
'            If x > 0 Then
'              atan4 = Atn(y / x)
'            ElseIf x < 0 Then
'              If y < 0 Then
'                atan4 = Atn(y / x) - PI
'              Else
'                atan4 = Atn(y / x) - PI
'              End If
'            Else  'x=0
'              atan4 = Sgn(y) * PI / 2
'            End If
'        End Function

