Attribute VB_Name = "modGeometry"
Option Explicit

Option Compare Binary
        
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

Public Enum AngleValue
    Whole = 0 'returns the base and the angle combined
    Base = 1 'cloests degree nearest multiples of 45 in radian
    Angle = 2 'base, sine and cosine of the angle added together
End Enum

Public Const PI As Single = 3.14159265358979
Public Const epsilon As Double = 0.0001 '0.999999999999999
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
    ToVector.Z = Point.Z
End Function

Public Function ToPoint(ByRef Vector As D3DVECTOR) As Point
    Set ToPoint = New Point
    ToPoint.X = Vector.X
    ToPoint.Y = Vector.Y
    ToPoint.Z = Vector.Z
End Function

Public Function ToPlane(ByRef V1 As Point, ByRef V2 As Point, ByRef V3 As Point) As Range
        
    Dim pNormal As Point
    Set pNormal = VectorCrossProduct(VectorDeduction(V2, V1), VectorDeduction(V3, V1))
    Set pNormal = VectorNormalize(pNormal)
        
    Set ToPlane = New Range
    With ToPlane
        .r = VectorDotProduct(pNormal, V1) * -1
        .X = pNormal.X
        .Y = pNormal.Y
        .Z = pNormal.Z
    End With
End Function

Public Function ToVec4(ByRef Plane As Range) As D3DVECTOR4
    ToVec4.X = Plane.X
    ToVec4.Y = Plane.Y
    ToVec4.Z = Plane.Z
    ToVec4.r = Plane.r
End Function

Public Function DistanceToPlane(ByRef p As Point, ByRef r As Range) As Single
    If Sqr(r.X * r.X + r.Y * r.Y + r.Z * r.Z) <> 0 Then
        DistanceToPlane = (r.X * p.X + r.Y * p.Y + r.Z * p.Z + r.r) / Sqr(r.X * r.X + r.Y * r.Y + r.Z * r.Z)
    End If
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = (((p2x - p1x) ^ 2) + ((p2y - p1y) ^ 2) + ((p2z - p1z) ^ 2))
    If Distance <> 0 Then Distance = Distance ^ (1 / 2)
End Function

Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = (((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
    If DistanceEx <> 0 Then DistanceEx = DistanceEx ^ (1 / 2)
End Function


Public Function DistanceSet(ByRef p1 As Point, ByVal p2 As Point, ByVal n As Single) As Point
    Dim d As Single
    d = DistanceEx(p1, p2)
    Set DistanceSet = New Point
    With DistanceSet
        If Not (d = n) Then
            If ((d > 0) And (n <> 0)) And (Not (d = n)) Then
        
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
                .X = (p1.X + (n * (.X / d)))
                .Y = (p1.Y + (n * (.Y / d)))
                .Z = (p1.Z + (n * (.Z / d)))
'#
                
            ElseIf (n = 0) Then
                .X = p1.X
                .Y = p1.Y
                .Z = p1.Z
            ElseIf (d = 0) Then
                .X = p2.X
                .Y = p2.Y
                .Z = p2.Z + IIf(p2.Z > p1.Z, n, -n)
            End If
        End If
    End With
End Function

Public Function PointOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    Dim r As Range
    Set r = ToPlane(v0, V1, V2)
    PointOnPlane = ((r.X * (p.X - v0.X)) + (r.Y * (p.Y - v0.Y)) + (r.Z * (p.Z - v0.Z)) = 0)
End Function
Public Function PointSideOfPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Boolean
    PointSideOfPlane = VectorDotProduct(PlaneNormal(v0, V1, V2), p) > 0
End Function

Public Function PointNearOnPlane(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point, ByRef p As Point) As Point
    Set PointNearOnPlane = New Point
    With PointNearOnPlane
        Dim r As Range
        Set r = ToPlane(v0, V1, V2)
        Dim n As Point
        Set n = PlaneNormal(v0, V1, V2)
        Dim d As Single
        d = DistanceToPlane(p, r)
        .X = p.X - (d * n.X)
        .Y = p.Y - (d * n.Y)
        .Z = p.Z - (d * n.Z)
    End With
End Function
Public Function LinePointByPercent(ByRef p1 As Point, ByRef p2 As Point, ByVal Factor As Single) As Point
    Set LinePointByPercent = New Point
    With LinePointByPercent
        .X = Least(p1.X, p2.X) + ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) * Factor)
        .Y = Least(p1.Y, p2.Y) + ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) * Factor)
        .Z = Least(p1.Z, p2.Z) + ((Large(p1.Z, p2.Z) - Least(p1.Z, p2.Z)) * Factor)
    End With
End Function
Public Function LineOpposite(ByVal Length1 As Single, ByVal Length2 As Single, ByVal Length3 As Single) As Single
    LineOpposite = Least(Length1, Length2, Length3)
End Function

Public Function LineAdjacent(ByVal Length1 As Single, ByVal Length2 As Single, ByVal Length3 As Single) As Single
    LineAdjacent = Large(Least(Length1, Length2), Large(Least(Length2, Length3), Least(Length3, Length1)))
End Function

Public Function LineHypotenuse(ByVal Length1 As Single, ByVal Length2 As Single, ByVal Length3 As Single) As Single
    LineHypotenuse = Large(Length1, Length2, Length3)
End Function

Public Function LineIntersectPlane(ByRef Plane As Range, PStart As Point, vDir As Point, ByRef VIntersectOut As Point) As Boolean
    Dim q As New Range     'Start Point
    Dim v As New Range       'Vector Direction

    Dim planeQdot As Single 'Dot products
    Dim planeVdot As Single
    
    Dim t As Single         'Part of the equation for a ray P(t) = Q + tV

    
    q.X = PStart.X          'Q is a point and therefore it's W value is 1
    q.Y = PStart.Y
    q.Z = PStart.Z
    q.r = 1
    
    v.X = vDir.X            'V is a vector and therefore it's W value is zero
    v.Y = vDir.Y
    v.Z = vDir.Z
    v.r = 0
    
  '  ((Plane.X * V.X) + (Plane.Y * V.Y) + (Plane.z * V.z) + (Plane.R * V.R))
    
    planeVdot = ((Plane.X * v.X) + (Plane.Y * v.Y) + (Plane.Z * v.Z) + (Plane.r * v.r)) 'D3DXVec4Dot(Plane, V)
    planeQdot = ((Plane.X * q.X) + (Plane.Y * q.Y) + (Plane.Z * q.Z) + (Plane.r * q.r)) 'D3DXVec4Dot(Plane, Q)
            
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
        .Z = (v.X ^ 2 + v.Y ^ 2 + v.Z ^ 2) ^ (1 / 2)
        If (.Z = 0) Then .Z = 1
        .X = (v.X / .Z)
        .Y = (v.Y / .Z)
        .Z = (v.Z / .Z)
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
        .Z = (Least(v0.Z, V1.Z, V2.Z, V3.Z) + ((Large(v0.Z, V1.Z, V2.Z, V3.Z) - Least(v0.Z, V1.Z, V2.Z, V3.Z)) / 2))
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

'Public Function TriangleNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
'    Set TriangleNormal = New Point
'    Dim o As Point
'    Dim d As Single
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
'    Dim d As Single
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
    Dim d As Single
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

Public Function VectorBalance(ByRef loZero As Point, ByRef hiWhole As Point, ByVal folcrumPercent As Single) As Point
    Set VectorBalance = New Point
    With VectorBalance
        .X = (loZero.X + ((hiWhole.X - loZero.X) * folcrumPercent))
        .Y = (loZero.Y + ((hiWhole.Y - loZero.Y) * folcrumPercent))
        .Z = (loZero.Z + ((hiWhole.Z - loZero.Z) * folcrumPercent))
    End With
End Function

Public Function TriangleFolcrum(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Point
    Set TriangleFolcrum = New Point
    With TriangleFolcrum
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




Public Function AngleAxisRestrict(ByRef AxisAngles As Point) As Point
    Set AngleAxisRestrict = New Point
    With AngleAxisRestrict
        .X = AngleRestrict(AxisAngles.X)
        .Y = AngleRestrict(AxisAngles.Y)
        .Z = AngleRestrict(AxisAngles.Z)
    End With
End Function

'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################

Public Function LineYIntercept(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Single
    '2D by nature of always exists, and not for 3D
    'Y-Intercept
    'b = -((m * x) - y)
    If p1 Is Nothing Then
        LineYIntercept = -((LineSlope2D(p2, p1) * p2.X) - p2.Y)
    Else
        LineYIntercept = -((LineSlope2D(p2, p1) * (p2.X - p1.X)) - (p2.Y - p1.Y))
    End If
End Function

Public Function LineSlope2D(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Single
    'slope
    'm = (y / x)
    If p1 Is Nothing Then
        If p2.X <> 0 Then LineSlope2D = p2.Y / p2.X    'rise over run
    Else
        If (p2.X - p1.X) <> 0 Then LineSlope2D = (p2.Y - p1.Y) / (p2.X - p1.X) 'rise over run
    End If
End Function

Public Function LineSlope3D(ByRef p2 As Point, Optional ByRef p1 As Point = Nothing) As Single
    If p1 Is Nothing Then Set p1 = New Point
     'run is the distance formula excluding the Y coordinate
    LineSlope3D = (((p2.X - p1.X) ^ 2) + ((p2.Z - p1.Z) ^ 2)) ^ (1 / 2)
    If LineSlope3D <> 0 Then 'rise doesn't include x or z, so now it's the same
        LineSlope3D = -((p2.Y - p1.Y) / LineSlope3D) 'rise over run
    Else
        LineSlope3D = 0
    End If
End Function

Public Function AngleRestrict(ByVal Angle1 As Single) As Single
    Angle1 = Angle1 * DEGREE
    Do While Round(Angle1, 0) > 360
        Angle1 = Angle1 - 360
    Loop
    Do While Round(Angle1, 0) <= 0
       Angle1 = Angle1 + 360
    Loop
    AngleRestrict = Angle1 * RADIAN
End Function

Public Function AngleOfCoord(ByRef Coord As Coord) As Single
    Dim X As Single
    Dim Y As Single
    X = Round(Coord.X, 6)
    Y = Round(Coord.Y, 6)
    If (X = 0) Then
        If (Y > 0) Then
            AngleOfCoord = (180 * RADIAN)
        ElseIf (Y < 0) Then
            AngleOfCoord = (360 * RADIAN)
        End If
    ElseIf (Y = 0) Then
        If (X > 0) Then
            AngleOfCoord = (90 * RADIAN)
        ElseIf (X < 0) Then
            AngleOfCoord = (270 * RADIAN)
        End If
    Else
        If ((X > 0) And (Y > 0)) Then
            AngleOfCoord = (90 * RADIAN)
        ElseIf ((X < 0) And (Y > 0)) Then
            AngleOfCoord = (180 * RADIAN)
        ElseIf ((X < 0) And (Y < 0)) Then
            AngleOfCoord = (270 * RADIAN)
        ElseIf ((X > 0) And (Y < 0)) Then
            AngleOfCoord = (360 * RADIAN)
        End If
        Dim slope As Single
        Dim Large As Single
        Dim Least As Single
        Dim Angle As Single
        If Abs(Coord.X) > Abs(Coord.Y) Then
            Large = Abs(Coord.X)
            Least = Abs(Coord.Y)
        Else
            Least = Abs(Coord.X)
            Large = Abs(Coord.Y)
        End If
        slope = (Least / Large)
        Angle = (((Coord.X ^ 2) + (Coord.Y ^ 2)) ^ (1 / 2))
        Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2))
        Least = (((Angle ^ 2) - (Least ^ 2)) ^ (1 / 2))
        Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / Angle)) * (Least / Angle))
        Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)
        Angle = Large + Least
        If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
           (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
            Angle = (PI / 4) - Angle
            AngleOfCoord = AngleOfCoord + (PI / 4)
        End If
        AngleOfCoord = AngleOfCoord + Angle
    End If
End Function

'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorAxisAngles(ByRef Point As Point) As Point
    Dim tmp As New Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
            Set tmp = Point
            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), -.X)
            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), -.Y)
            .Z = AngleRestrict(AngleOfCoord(MakePoint(tmp.X, tmp.Y, tmp.Z)))
            Set tmp = Nothing
        End If
    End With
End Function


'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
    Dim tmp As Point
    Set tmp = Point
    Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
    Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
    Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
    Set VectorRotateAxis = tmp
    Set tmp = Nothing
End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = Point
'    If Abs(Angles.X) > Abs(Angles.Y) And Abs(Angles.X) > Abs(Angles.Z) And (Angles.X <> 0) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(0, Angles.Y, Angles.Z))
'    ElseIf Abs(Angles.Y) > Abs(Angles.X) And Abs(Angles.Y) > Abs(Angles.Z) And (Angles.Y <> 0) Then
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, 0, Angles.Z))
'    ElseIf Abs(Angles.Z) > Abs(Angles.Y) And Abs(Angles.Z) > Abs(Angles.X) And (Angles.Z <> 0) Then
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, Angles.Y, 0))
'    End If
'    Set VectorRotateAxis = tmp
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

'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorRotateX(ByRef Point As Point, ByVal Angle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    Set VectorRotateX = New Point
    With VectorRotateX
        .Z = Point.Z * CosPhi - Point.Y * SinPhi
        .Y = Point.Z * SinPhi + Point.Y * CosPhi
        .X = Point.X
    End With
End Function

Public Function VectorRotateY(ByRef Point As Point, ByVal Angle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    Set VectorRotateY = New Point
    With VectorRotateY
        .X = Point.X * CosPhi - Point.Z * SinPhi
        .Z = Point.X * SinPhi + Point.Z * CosPhi
        .Y = Point.Y
    End With
End Function

Public Function VectorRotateZ(ByRef Point As Point, ByVal Angle As Single) As Point
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(Angle)
    SinPhi = Sin(Angle)
    Set VectorRotateZ = New Point
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

Public Function VectorSlope(ByRef p1 As Point, ByRef p2 As Point) As Single
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

Public Function VectorRise(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Single
    VectorRise = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
End Function
Public Function VectorRun(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Single
    VectorRun = DistanceEx(MakePoint(p1.X, 0, p1.Z), MakePoint(p2.X, 0, p2.Z))
End Function

Public Function VectorSine(ByRef p As Point) As Single
    'returns the z axis angle of the x and y in p
    If p.X = 0 Then
        If p.Y <> 0 Then
            VectorSine = Val("0.#IND")
        End If
    ElseIf p.Y <> 0 Then
        VectorSine = Round(Abs(p.Y / (((p.X ^ 2) + (p.Y ^ 2)) ^ (1 / 2))), 2)
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
        VectorCosine = Round(Abs(p.X / (((p.X ^ 2) + (p.Y ^ 2)) ^ (1 / 2))), 2)
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
Public Function AngleConvertWinToDX3DX(ByVal Angle As Single) As Single
    AngleConvertWinToDX3DX = AngleRestrict(Angle) '[(((360 - Abs(Angle * DEGREE)) * Sign(Angle * DEGREE)) * RADIAN))
End Function

Public Function AngleConvertWinToDX3DY(ByVal Angle As Single) As Single
    AngleConvertWinToDX3DY = AngleRestrict(Angle) '[(((360 - Abs(Angle * DEGREE)) * Sign(Angle * DEGREE)) * RADIAN))
End Function

Public Function AngleConvertWinToDX3DZ(ByVal Angle As Single) As Single
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
    
    Dim c1 As Single
    Dim C2 As Single
    
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
Public Function ValueInfluence(ByVal Final As Single, ByVal Current As Single, Optional ByVal Amount As Single = 0.001, _
                                Optional ByVal Factor As Single = 1, Optional ByVal SnapRange As Single = 0) As Single

    If (Not ValueSnapCheck(Final, Current, SnapRange)) Then
        Dim n As Single
        n = Large(Final, Current) - Least(Final, Current)
        If (n <= Abs(SnapRange) And Abs(SnapRange) > 0) Then
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

Public Function ValueSnapCheck(ByVal Final As Single, ByVal Current As Single, ByVal SnapRange As Single) As Boolean
    If SnapRange = 0 Or (Current = Final) Then
        ValueSnapCheck = (Current = Final)
    Else
        Dim n As Single
        n = Abs((Large(Final, Current) - Least(Final, Current)))
        If (n <= Abs(SnapRange) And Abs(SnapRange) > 0) Then
            ValueSnapCheck = True
        End If
    End If
End Function

Public Function VectorInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Single = 0.001, _
                                Optional ByVal Factor As Single = 1, Optional ByVal Concurrent As Boolean = True, _
                                Optional ByVal SnapRange As Single = 0) As Point
                                
    Set VectorInfluence = VectorDisplace(Current, Final)
    With VectorInfluence
        Dim n As Point
        If Not Concurrent Then
            Set n = VectorNormalize(VectorInfluence)
            n.X = IIf(n.X = 0, 1, n.X) * 100
            n.Y = IIf(n.Y = 0, 1, n.Y) * 100
            n.Z = IIf(n.Z = 0, 1, n.Z) * 100
        Else
            Set n = New Point
            n.X = 100
            n.Y = 100
            n.Z = 100
        End If
   
        .X = ValueInfluence(Final.X, Current.X, Amount * ((VectorInfluence.X * Factor) / n.X), SnapRange)
        .Y = ValueInfluence(Final.Y, Current.Y, Amount * ((VectorInfluence.Y * Factor) / n.Y), SnapRange)
        .Z = ValueInfluence(Final.Z, Current.Z, Amount * ((VectorInfluence.Z * Factor) / n.Z), SnapRange)
   
        Set n = Nothing
    End With
End Function

Public Function AngleInfluence(ByVal Final As Single, ByVal Current As Single, Optional ByVal Amount As Single = 0.001, _
                                Optional ByVal Factor As Single = 1, Optional ByVal SnapRange As Single = 0) As Single
        
        Dim a1 As Single
        Dim a2 As Single
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

Public Function AngleAxisInfluence(ByRef Final As Point, ByRef Current As Point, Optional ByVal Amount As Single = 0.001, _
                                    Optional ByVal Factor As Single = 1, Optional ByVal Concurrent As Boolean = True, _
                                    Optional ByVal SnapRange As Single = 0) As Point
    
    Set AngleAxisInfluence = AngleAxisDifference(Current, Final)
    With AngleAxisInfluence
        Dim n As Point
        If Not Concurrent Then
            Set n = AngleAxisNormalize(AngleAxisInfluence)
            n.X = IIf(n.X = 0, 1, n.X) '* 100
            n.Y = IIf(n.Y = 0, 1, n.Y) ' * 100
            n.Z = IIf(n.Z = 0, 1, n.Z) '* 100
        Else
            Set n = New Point
            n.X = 0.01 '100
            n.Y = 0.01 '100
            n.Z = 0.01 ' 100
        End If
        
        .X = AngleInfluence(Final.X, Current.X, Amount, ((.X * Factor) / n.X), SnapRange)
        .Y = AngleInfluence(Final.Y, Current.Y, Amount, ((.Y * Factor) / n.Y), SnapRange)
        .Z = AngleInfluence(Final.Z, Current.Z, Amount, ((.Z * Factor) / n.Z), SnapRange)
        
        Set n = Nothing
    End With
End Function


Public Function AngleAxisInbetween(ByRef ZeroPercent As Point, ByRef OneHundred As Point, Optional ByVal DecimalPercent As Single = 0.5) As Point

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
    
    Dim c1 As Single
    Dim C2 As Single
    
    Set AngleAxisInbetween = New Point
    With AngleAxisInbetween
        c1 = Large(d1.X, d2.X)
        C2 = Least(d1.X, d2.X)
        .X = (c1 - C2)
        
        c1 = Large(d1.Y, d2.Y)
        C2 = Least(d1.Y, d2.Y)
        .Y = (c1 - C2)
        
        c1 = Large(d1.Z, d2.Z)
        C2 = Least(d1.Z, d2.Z)
        .Z = (c1 - C2)
        
        .X = (.X * DecimalPercent) * RADIAN
        .Y = (.Y * DecimalPercent) * RADIAN
        .Z = (.Z * DecimalPercent) * RADIAN
        Set AngleAxisInbetween = AngleAxisRestrict(MakePoint(.X, .Y, .Z))
    End With
    
    
    
    Set d1 = Nothing
    Set d2 = Nothing

End Function

Public Function AngleAxisPercentOf(ByRef AngleAxis As Point, ByVal DecimalPercent As Single) As Point

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

Public Function CrossProductLength(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
    CrossProductLength = ((p1.X - p2.X) * (p2.Y - p2.Y) - (p1.Y - p2.Y) * (p2z - p2.Z) - (p1.Z - p2.Z) * (p2.X - p2.X))
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

Public Function VectorQuantify(ByRef p1 As Point) As Single
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
Public Function VectorMagnitude(ByVal p1 As Point) As Single
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

Public Function VectorDivision(ByRef p1 As Point, ByVal n As Single) As Point
    Set VectorDivision = New Point
    With VectorDivision
        .X = (p1.X / n)
        .Y = (p1.Y / n)
        .Z = (p1.Z / n)
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
    Dim tmp As Single
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

Public Function AbsoluteFactor(ByVal n As Single) As Single
    'returns -1 if the n is below zero, returns 1 if n is above zero, and 0 if n is zero
    AbsoluteFactor = ((-(AbsoluteValue(n - 1) - n) - (-AbsoluteValue(n + 1) + n)) * 0.5)
End Function

Public Function AbsoluteValue(ByVal n As Single) As Single
    'same as abs(), returns a number as positive quantified
    AbsoluteValue = (-((-(n * -1) * n) ^ (1 / 2) * -1))
End Function

Public Function AbsoluteWhole(ByVal n As Single) As Single
    'returns only the digits to the left of a decimal in any numerical
    'AbsoluteWhole = (AbsoluteValue(n) - (AbsoluteValue(n) - (AbsoluteValue(n) Mod (AbsoluteValue(n) + 1)))) * AbsoluteFactor(n)
    AbsoluteWhole = (n \ 1) 'is also correct
End Function

Public Function AbsoluteDecimal(ByVal n As Single) As Single
    'returns only the digits to the right of a decimal in any numerical
    AbsoluteDecimal = (AbsoluteValue(n) - AbsoluteValue(AbsoluteWhole(n))) * AbsoluteFactor(n)
End Function

Public Function AngleQuadrant(ByVal Angle As Single) As Single
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

Public Function VectorQuadrant(ByRef p As Point) As Single
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

Public Function VectorOctet(ByRef p As Point) As Single
    VectorOctet = VectorQuadrant(p)
    If p.Z < 0 Then VectorOctet = VectorOctet + 4
End Function


Public Function AbsoluteInvert(ByVal Value As Long, Optional ByVal Whole As Long = 100, Optional ByVal Unit As Long = 1)
    'returns the inverted value of a whole conprised of unit measures, AbsoluteInvert(25) returns 75
    'another example: AbsoluteInvert(0, 16777216) returns the negative of black 0, which is 16777216
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





'Public Function TriangleOpposite(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Single
'    Dim l1 As Single
'    Dim l2 As Single
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
'Public Function TriangleAdjacent(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Single
'    'provide the Hypotenuse as line p1-p2, or all points to the triangle
'    TriangleAdjacent = DistanceEx(p1, p2)
'    If Not p3 Is Nothing Then
'        Dim l1 As Single
'        Dim l2 As Single
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
'Public Function TriangleHypotenuse(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Single
'    TriangleHypotenuse = DistanceEx(p1, p2)
'    If p3 Is Nothing Then
'        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (TriangleHypotenuse ^ 2)) ^ (1 / 2)
'    Else
'        TriangleHypotenuse = ((TriangleHypotenuse ^ 2) + (DistanceEx(p2, p3) ^ 2)) ^ (1 / 2)
'    End If
'End Function
'
'
'Public Function InvSin(number As Single) As Single
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
'Public Function CoordOfAngle(ByVal RadiusLength As Single, ByVal AngleInRadian As Single) As Point
'    Set CoordOfAngle = New Point
'    With CoordOfAngle
'
'        Dim p2 As New Point
'        Dim d As New Point 'the point we will modify for the finish
'        Dim Angle As Single 'the angle we'll modify for the finish
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
'            Dim A As Single 'a+aa is a whole 100% along the x axis of the angles
'            Dim aa As Single 'percent equaling aa with an equalteral right traingle
'            Dim b As Single 'b+bb is a whole 100% along the y axis of the angles
'            Dim bb As Single 'percent equaling bb with an equalteral right traingle
'            Dim m As Single 'slope of the line of the unit circle initial values
'            Dim an As Single 'angle in definition with the -PI*2, PI, 0, PI, and PI*2 of
'            'radian spectrum (5 angles) excluded and the 0, 45 and 90 out degree spectrum
'            '(3 angles) excluded, overlapping, hence the hard values of ((5/45) / (-3/180))
'
'            Dim t As Single 'temporary
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
'Public Function AngleOfWave(ByVal WaveXLo As Single, ByVal WaveYDist As Single, ByVal WaveZHi As Single) As Single
'    Const LargePI As Single = ((((PI / 4) * DEGREE) - 1) * RADIAN)
'    Const LeastPI As Single = ((((PI / 16) * DEGREE) + 2) * RADIAN)
'
'    Dim slope As Single
'    Dim WaveHype1 As Single
'    Dim WaveHype2 As Single
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
'            Dim X As Single
'            Dim Y As Single
'            Dim z As Single
'            Dim Angle As Single
'            'round them off for checking
'            '(6 is for single precision)
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
'                Dim slope As Single
'                Dim dist As Single
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
'                            Set pZ = CoordOfAngle(dist, (.z + Angles.z))
'                            'Form1.Picture3.Circle (pZ.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - pZ.y), 8 * Screen.TwipsPerPixelX, &HC000&
'                        Case 3
'                            Set pY = CoordOfAngle(dist, (.z + Angles.Y))
'                            'Form1.Picture2.Circle (pY.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - pY.y), 8 * Screen.TwipsPerPixelX, &HC000&
'                        Case 2
'                            Set pX = CoordOfAngle(dist, (.z + Angles.X))
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
'                        Dim sx As Single
'                        Dim sy As Single
'                        Dim sz As Single
'                        Dim cx As Single
'                        Dim cy As Single
'                        Dim cz As Single
'                        Dim tx As Single
'                        Dim ty As Single
'                        Dim tz As Single
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
'                        Dim p As Coord
'                        Select Case Abs(stack)
'                            Case 4
'
'                                Form1.Picture3.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                            Case 3
'                                Set p = CoordOfAngle(dist, Angles.z)
'                                Point.y = p.x * Cos(.z) - p.y * Sin(.z)
'                                Point.x = p.x * Sin(.z) + p.y * Cos(.z)
'                                Form1.Picture2.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                            Case 2
'                                Set p = CoordOfAngle(dist, Angles.z)
'                                Point.y = p.x * Cos(.z) - p.y * Sin(.z)
'                                Point.x = p.x * Sin(.z) + p.y * Cos(.z)
'                                Form1.Picture1.Circle (Point.x + (Form1.Picture3.ScaleWidth / 2), (Form1.Picture3.ScaleHeight / 2) - Point.y), 8 * Screen.TwipsPerPixelX, &H8000&
'                        End Select


'                        Dim XOld As Single
'                        Dim YOld As Single
'                        Dim ZOld As Single
'                        Dim XNew As Single
'                        Dim YNew As Single
'                        Dim ZNew As Single
'
'                        Dim SinAngleX As Single
'                        Dim CosAngleX As Single
'                        Dim TanAngleX As Single
'                        Dim SinAngleY As Single
'                        Dim CosAngleY As Single
'                        Dim TanAngleY As Single
'                        Dim SinAngleZ As Single
'                        Dim CosAngleZ As Single
'                        Dim TanAngleZ As Single
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
'Function Atn2(ByVal X As Single) As Single
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
'Function Atn2(ByVal X As Single) As Single
'
'    If Abs(-X * X + 1) <> 0 Then
'        If Sqr(Abs(-X * X + 1)) <> 0 Then
'            Atn2 = Abs(Atn(X / Abs(Sqr(Abs(-X * X + 1)))))
'        End If
'    End If
'
'End Function
'Function atan3(ByVal i As Single, ByVal r As Single) As Single
'        Dim Theta As Single
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

