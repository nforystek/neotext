Attribute VB_Name = "modGeometry"
Option Explicit

Option Compare Binary

'Private Enum AngleValue
'    Base = 1
'    slope = 2
'    Sine = 4
'    Cosine = 8
'    Tangent = 16
'    CoTangent = 32
'    Secant = 65
'    CoSecant = 128
'End Enum

Private Enum AngleValue
    Base = 1 'cloests degree nearest multiples of 45 in radian
    Sine = 2 'the majority of the angle with in 0 to 45
    Cosine = 3 'the remainder of the angle with in 0 to 22.5
    Tangent = 4 'the tangent of the sine and cosine of the angle
    Angle = 5 'base, sine and cosine of the angle added together
End Enum

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
Public Function LinePointByPercent(ByRef p1 As Point, ByRef p2 As Point, ByVal Factor As Single) As Point
    Set LinePointByPercent = New Point
    With LinePointByPercent
        .X = Least(p1.X, p2.X) + ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) * Factor)
        .Y = Least(p1.Y, p2.Y) + ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) * Factor)
        .z = Least(p1.z, p2.z) + ((Large(p1.z, p2.z) - Least(p1.z, p2.z)) * Factor)
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

Public Function TriangleIntersect(ByRef t1p1 As Point, ByRef t1p2 As Point, ByRef t1p3 As Point, ByRef t2p1 As Point, ByRef t2p2 As Point, ByRef t2p3 As Point) As Point
    'Debug.Print TriangleIntersect(MakePoint(-5, 0, 0), MakePoint(5, 0, -5), MakePoint(5, 0, 5), MakePoint(-2.5, 2.5, 0), MakePoint(2.5, 2.5, 0), MakePoint(0, -2.5, 0))
    'compute another way of representing triangles, the center, normal and side lengths
    Dim t1n As New Point
    Dim t2n As New Point
    Dim t1a As New Point
    Dim t2a As New Point
    Dim t1l As New Point
    Dim t2l As New Point
    
    t1n = TriangleNormal(t1p1, t1p2, t1p3)
    t2n = TriangleNormal(t2p1, t2p2, t2p3)

    Debug.Print t1n; VectorIsNormal(t1n); t2n; VectorIsNormal(t2n)

'    t1n = PlaneNormal(t1p1, t1p2, t1p3)
'    t2n = PlaneNormal(t2p1, t2p2, t2p3)
'    Debug.Print t1n; VectorIsNormal(t1n); t2n; VectorIsNormal(t2n)
    
    t1a = TriangleAxii(t1p1, t1p2, t1p3)
    t2a = TriangleAxii(t2p1, t2p2, t2p3)
    
    Debug.Print t1a; t2a
    
    t1l.X = DistanceEx(t1p1, t1p2)
    t1l.Y = DistanceEx(t1p2, t1p3)
    t1l.z = DistanceEx(t1p3, t1p1)
    
    t2l.X = DistanceEx(t2p1, t2p2)
    t2l.Y = DistanceEx(t2p2, t2p3)
    t2l.z = DistanceEx(t2p3, t2p1)
     
    Debug.Print t1l; t2l
       
    
    Set TriangleIntersect = New Point
    With TriangleIntersect


    End With
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

Public Function TriangleLowestOfAll(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleLowestOfAll = New Point
    With TriangleLowestOfAll
        .X = Least(p1.X, p2.X, p3.X)
        .Y = Least(p1.Y, p2.Y, p3.Y)
        .z = Least(p1.z, p2.z, p3.z)
    End With
End Function

Public Function TriangleLargestOfAll(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Set TriangleLargestOfAll = New Point
    With TriangleLargestOfAll
        .X = Large(p1.X, p2.X, p3.X)
        .Y = Large(p1.Y, p2.Y, p3.Y)
        .z = Large(p1.z, p2.z, p3.z)
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
        d = (Abs(o.X) + Abs(o.Y) + Abs(o.z))
        If (d <> 0) Then
            .z = (((Abs(o.X) + Abs(o.Y)) - Abs(o.z)) / d)
            .X = (((Abs(o.Y) + Abs(o.z)) - Abs(o.X)) / d)
            .Y = (((Abs(o.z) + Abs(o.X)) - Abs(o.Y)) / d)
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
    Dim l1 As Single
    Dim l2 As Single
    TriangleOpposite = DistanceEx(p1, p2)
    l1 = DistanceEx(p2, p3)
    l1 = DistanceEx(p3, p1)
    If l1 < TriangleOpposite Then
        If l2 < l1 And l2 < TriangleOpposite Then
            TriangleOpposite = ((TriangleOpposite ^ 2) - (l1 ^ 2)) ^ (1 / 2)
        ElseIf l2 > TriangleOpposite Then
            TriangleOpposite = ((l2 ^ 2) - (Large(l1, TriangleOpposite) ^ 2)) ^ (1 / 2)
        Else
            TriangleOpposite = ((TriangleOpposite ^ 2) - (l2 ^ 2)) ^ (1 / 2)
        End If
    Else
        TriangleOpposite = ((l1 ^ 2) - (TriangleOpposite ^ 2)) ^ (1 / 2)
    End If
End Function

Public Function TriangleAdjacent(ByRef p1 As Point, ByRef p2 As Point, Optional ByRef p3 As Point = Nothing) As Single
    'provide the Hypotenuse as line p1-p2, or all points to the triangle
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


Public Function AngleAxisRestrict(ByRef AxisAngles As Point) As Point
    Set AngleAxisRestrict = New Point
    With AngleAxisRestrict
        .X = AngleRestrict(AxisAngles.X)
        .Y = AngleRestrict(AxisAngles.Y)
        .z = AngleRestrict(AxisAngles.z)
    End With
End Function
Public Function AngleRestrict(ByRef AxisAngle As Single) As Single
    'cleans up angles that are beyond bounds, returning a count
    'of how many times full circle was removed, + or -
    'i.e. 720 rotates twice around, and is 360 the same
    'so it would then reutrn 1, changing 720 to 360 too
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

Public Function InvSin(number As Single) As Single
    InvSin = -number * number + 1
    If InvSin > 0 Then
        InvSin = Sqr(InvSin)
        If InvSin <> 0 Then InvSin = Atn(number / InvSin)
    Else
        InvSin = 0
    End If
End Function



'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################
'########################################################################################################################

Public Function VectorRotateAimAt(ByRef Point As Point, ByRef AimAt As Point) As Point
    
    Dim tmp As New Point
    Set VectorRotateAimAt = New Point
    With VectorRotateAimAt
    
    
        

        
        

    End With
    
End Function

Public Function VectorRotateAxis(ByRef PointToRotate As Point, ByRef RadianAngles As Point) As Point
    
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
        tmp.X = .X
        tmp.Y = .Y

    End With
    
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


Public Function LineSlope(ByVal x1 As Single, ByVal y1 As Single, Optional ByVal x2 As Single, Optional ByVal y2 As Single) As Single
    If (y2 - y1) <> 0 Then
        LineSlope = (x2 - x1) / (y2 - y1) 'rise over run
    End If
End Function

Public Function LineSlope3D(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, Optional ByVal x2 As Single, Optional ByVal y2 As Single, Optional ByVal z2 As Single) As Single
    LineSlope3D = (((x2 - x1) ^ 2) + ((z2 - z1) ^ 2)) ^ (1 / 2)
    If LineSlope3D <> 0 Then
        LineSlope3D = -((y2 - y1) / LineSlope3D) 'rise over run
    Else
        LineSlope3D = 0
    End If
End Function

Public Function SquareSlope(ByVal x1 As Single, ByVal y1 As Single, Optional ByVal x2 As Single, Optional ByVal y2 As Single) As Single
    'never returns above 1 or below -1
    x1 = x2 - x1
    y1 = y2 - y1
    If (y1 = 0) Or (x1 = 0 And y1 = 0) Then 'horizontal
        SquareSlope = 0
    ElseIf (x1 = 0) Then 'vertical
        SquareSlope = IIf(x1 < 0 Or y1 < 0, -1, 1)
    Else
        If (y1 = x1) Then 'diagonal
            SquareSlope = IIf(x1 < 0 Or y1 < 0, -0.5, 0.5)
        Else
            SquareSlope = Round((y1 / x1), 6) 'rise over run
        End If
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


Private Function AngleOfCoord(ByRef Coord As Coord) As Single
    Dim X As Single
    Dim Y As Single
    'round them off for checking
    X = Round(Coord.X, 6)
    Y = Round(Coord.Y, 6)
    'returns the z axis angle of the x and y in p
    If (X = 0) Then 'slope of 1
        'must be 360 or 180
        If (Y > 0) Then
            AngleOfCoord = (180 * RADIAN)
        ElseIf (Y < 0) Then
            AngleOfCoord = (360 * RADIAN)
        End If
    ElseIf (Y = 0) Then 'slope of 0
        'must be 90 or 270
        If (X > 0) Then
            AngleOfCoord = (90 * RADIAN)
        ElseIf (X < 0) Then
            AngleOfCoord = (270 * RADIAN)
        End If
    Else
        'get the base angle
        '(up to the quardrant)
        If ((X > 0) And (Y > 0)) Then
            AngleOfCoord = (90 * RADIAN)
        ElseIf ((X < 0) And (Y > 0)) Then
            AngleOfCoord = (180 * RADIAN)
        ElseIf ((X < 0) And (Y < 0)) Then
            AngleOfCoord = (270 * RADIAN)
        ElseIf ((X > 0) And (Y < 0)) Then
            AngleOfCoord = (360 * RADIAN)
        End If
        'all the trickery
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
        slope = (Least / Large) 'the angle in square form
        '^^ or tangent, tangable to other axis angles' shared axis
        Angle = (((Coord.X ^ 2) + (Coord.Y ^ 2)) ^ (1 / 2)) 'distance, for now
        Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
        Least = (((Angle ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
        Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / Angle)) * (Least / Angle))
        '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
        'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
        Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope) 'bulk sine of the angle in 45 degree slices
        '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
        Angle = Large + Least
        If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
           (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
            Angle = (PI / 4) - Angle
            AngleOfCoord = AngleOfCoord + (PI / 4)
        End If
        AngleOfCoord = AngleOfCoord + Angle
    End If
End Function

Public Function AnglesOfPoint(ByRef Point As Point, Optional ByRef Angles As Point) As Point
    Static stack As Integer
    stack = stack + 1
    If stack = 1 Then
        '(1,1,1) is high noon
        'to 45 degree sections
        Point.X = Point.X + 1
        Point.Y = Point.Y + 1
        Point.z = Point.z + 1
    End If
    Set AnglesOfPoint = New Point
    With AnglesOfPoint
        If stack < 5 Then
            Dim X As Single
            Dim Y As Single
            Dim z As Single
            'round them off for checking
            '(6 is for single precision)
            X = Round(Point.X, 6)
            Y = Round(Point.Y, 6)
            z = Round(Point.z, 6)
            If (X = 0) Then  'slope of 1
                If (z = 0) Then
                    'must be 360 or 180
                    If (Y > 0) Then
                        .z = (180 * RADIAN)
                    ElseIf (Y < 0) Then
                        .z = (360 * RADIAN)
                    End If
                Else
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.z
                    AnglesOfPoint.z = Point.X
                    .z = AnglesOfPoint(AnglesOfPoint, Angles).z
                End If
            ElseIf (Y = 0) Then   'slope of 0
                If (z = 0) Then
                    'must be 90 or 270
                    If (X > 0) Then
                        .z = (90 * RADIAN)
                    ElseIf (X < 0) Then
                        .z = (270 * RADIAN)
                    End If
                Else
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.z
                    AnglesOfPoint.z = Point.X
                    .z = AnglesOfPoint(AnglesOfPoint, Angles).z
                End If
            ElseIf (X <> 0) And (Y <> 0) Then
                Dim slope As Single
                Dim dist As Single
                Dim Large As Single
                Dim Least As Single
                Dim Angle As Single
                'find the larger coordinate
                If Abs(Point.X) > Abs(Point.Y) Then
                    Large = Abs(Point.X)
                    Least = Abs(Point.Y)
                Else
                    Least = Abs(Point.X)
                    Large = Abs(Point.Y)
                End If
                slope = (Least / Large) 'the angle in square form
                '^^ or tangent, tangable to other axis angles' shared axis
                dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2)) 'distance
                'still traveling for tangents and cosines
                Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
                Least = (((dist ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
                Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / dist)) * (Least / dist))
                '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
                'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
                Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)  'bulk sine of the angle in 45 degree slices
                '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
                If (z <> 0) Then 'two or less axis is one rotation
                    Dim Ret As Point
                    AnglesOfPoint.X = Point.Y
                    AnglesOfPoint.Y = Point.z
                    AnglesOfPoint.z = Point.X
                    Set Ret = AnglesOfPoint(AnglesOfPoint, Angles)
                    If stack = 2 Then
                        .X = -Ret.z
                    End If
                    If stack = 1 Then
                        .X = -Ret.X
                        .Y = Ret.z
                    End If
                    Set Ret = Nothing
                End If
                'get the base angle
                '(up to the quardrant)
                If ((X > 0) And (Y > 0)) Then
                    .z = (90 * RADIAN)
                ElseIf ((X < 0) And (Y > 0)) Then
                    .z = (180 * RADIAN)
                ElseIf ((X < 0) And (Y < 0)) Then
                    .z = (270 * RADIAN)
                ElseIf ((X > 0) And (Y < 0)) Then
                    .z = (360 * RADIAN)
                End If
                'develop the final angle Z for this duel coordinate X,Y axis only
                Angle = (Large + Least)
                If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
                   (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
                   'the angle for 45 to 90 is in reverse, and doesn't start at 45, but because we
                   'are calculating a second 45 of 90, the offset (-1 not 0) is included if inverse
                    Angle = (PI / 4) - Angle
                    'then also add 45 to the base
                    .z = .z + (PI / 4)
                End If
                'add it to the base, returing as .Z
                .z = .z + Angle

                If stack = 1 Then
                    'reorganization
                    Angle = .Y
                    .Y = .z
                    .z = Angle
                    Angle = .X
                    .X = .Y
                    .Y = .z
                    .z = Angle
                    Angle = .X
                    .X = .Y
                    .Y = .z
                    .z = Angle
                End If
                
                If ((Not (Angles Is Nothing)) And (dist > 0)) Then

                        
'                        Dim p1 As Point
'                        Dim p2 As Point
'                        Dim p3 As Point
'                        Dim tmp As Point
'                        Point.X = Point.X - 1
'                        Point.Y = Point.Y - 1
'                        Point.Z = Point.Z - 1
'                        Set p1 = VectorRotateZ(Point, -Angles.Z)
'                        Set p2 = VectorRotateY(Point, -Angles.Y)
'                        Set p3 = VectorRotateX(Point, -Angles.X)
'
'                        Set tmp = TriangleAxii(p1, p2, p3)
'                        Point.X = tmp.X + 1
'                        Point.Y = tmp.Y + 1
'                        Point.Z = tmp.Z + 1
                        
                    If (.z + Angles.z) <> 0 Then

                        Dim tmp As Point
                        Point.X = Point.X - 1
                        Point.Y = Point.Y - 1
                        Point.z = Point.z - 1
                        Set tmp = VectorRotateZ(Point, -Angles.z)
                        Set tmp = VectorRotateY(tmp, -Angles.Y)
                        Set tmp = VectorRotateX(tmp, -Angles.X)
                        Point.X = tmp.X + 1
                        Point.Y = tmp.Y + 1
                        Point.z = tmp.z + 1
                    End If

'                        Point.X = Point.X - 1
'                        Point.Y = Point.Y - 1
'                        Point.z = Point.z - 1
'
'                        Dim p1 As New Point
'                        Dim tmp As New Point
'                        p1.Y = AngleFunction( Cos(Angles.X) * Point.Y - Sin(Angles.X) * Point.z
'                        p1.z = Sin(Angles.X) * Point.Y + Cos(Angles.X) * Point.z
'                        tmp.X = Point.X
'                        tmp.Y = p1.Y
'                        tmp.z = p1.z
'                        p1.X = Sin(Angles.Y) * tmp.z + Cos(Angles.Y) * tmp.X
'                        p1.z = Cos(Angles.Y) * tmp.z - Sin(Angles.Y) * tmp.X
'                        tmp.X = p1.X
'                        tmp.z = p1.z
'                        p1.X = Cos(Angles.z) * tmp.X - Sin(Angles.z) * tmp.Y
'                        p1.Y = Sin(Angles.z) * tmp.X + Cos(Angles.z) * tmp.Y
'                        p1.z = tmp.z
'                        tmp.X = p1.X
'                        tmp.Y = p1.Y
'
'
'                        Point.X = Point.X + 1
'                        Point.Y = Point.Y + 1
'                        Point.z = Point.z + 1

'                    If (.Z + Angles.Z) <> 0 Then
'                        Point.X = Point.X - 1
'                        Point.Y = Point.Y - 1
'                        Point.Z = Point.Z - 1
'                        Dim p1 As New Point
'                        Dim tmp As New Point
'                        p1.Y = Cos(Angles.X) * Point.Y - Sin(Angles.X) * Point.Z
'                        p1.Z = Sin(Angles.X) * Point.Y + Cos(Angles.X) * Point.Z
'                        tmp.X = Point.X
'                        tmp.Y = p1.Y
'                        tmp.Z = p1.Z
'                        p1.X = Sin(Angles.Y) * tmp.Z + Cos(Angles.Y) * tmp.X
'                        p1.Z = Cos(Angles.Y) * tmp.Z - Sin(Angles.Y) * tmp.X
'                        tmp.X = p1.X
'                        tmp.Z = p1.Z
'                        p1.X = Cos(Angles.Z) * tmp.X - Sin(Angles.Z) * tmp.Y
'                        p1.Y = Sin(Angles.Z) * tmp.X + Cos(Angles.Z) * tmp.Y
'                        p1.Z = tmp.Z
'                        tmp.X = p1.X
'                        tmp.Y = p1.Y
'                        Point.X = tmp.X + ((((tmp.Y - tmp.Z) - tmp.X) + (tmp.Z - tmp.Y)) + tmp.X)
'                        Point.Y = tmp.Y + ((tmp.X - tmp.Z) - tmp.Y) + -((-tmp.Z + tmp.X) - tmp.Y)
'                        Point.Z = tmp.Z + ((((tmp.Y - tmp.X) - tmp.Z) + (tmp.Z - tmp.Y)) + tmp.X)
'                        Point.X = Point.X + 1
'                        Point.Y = Point.Y + 1
'                        Point.Z = Point.Z + 1


'                        Point.X = Point.X - 1
'                        Point.Y = Point.Y - 1
'                        Point.z = Point.z - 1
'                        Dim d As New Point
'                        Angle = Round(AngleRestrict(.z + Angles.z) * DEGREE, 6)
'                        Dim dist1 As Single
'                        dist1 = DistanceEx(MakePoint(0, 0, 0), MakePoint(Point.X, Point.Y, 0))
'                        Dim tmp As Point
'                        Dim p1 As New Point
'                        Dim p2 As New Point
'                        Select Case Angle
'                            Case 45
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(2, 2, 0), dist1)
'                            Case 90
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(2, 0, 0), dist1)
'                            Case 135
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(2, -2, 0), dist1)
'                            Case 180
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(0, -2, 0), dist1)
'                            Case 225
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(-2, -2, 0), dist1)
'                            Case 270
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(-2, 0, 0), dist1)
'                            Case 315
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(-2, 2, 0), dist1)
'                            Case 360
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(0, 2, 0), dist1)
'                            Case Else
'                                If Angle > 270 Then
'                                    p1.Y = 1
'                                    p1.X = -1
'                                    Angle = Round(Angle - ((Angle \ 90) * 90), 2)
'                                    If Angle > 45 Then
'                                        Angle = Angle - 45
'                                        p2.X = 0
'                                        p2.Y = 2
'                                    Else
'                                        p2.X = -2
'                                        p2.Y = 0
'                                        Swap p1, p2
'                                    End If
'                                ElseIf Angle > 180 Then
'                                    p1.Y = -1
'                                    p1.X = -1
'                                    Angle = Round(Angle - ((Angle \ 90) * 90), 2)
'                                    If Angle > 45 Then
'                                        Angle = Angle - 45
'                                        p2.X = -2
'                                        p2.Y = 0
'                                    Else
'                                        p2.X = 0
'                                        p2.Y = -2
'                                        Swap p1, p2
'                                    End If
'                                ElseIf Angle > 90 Then
'                                    p1.Y = -1
'                                    p1.X = 1
'                                    Angle = Round(Angle - ((Angle \ 90) * 90), 2)
'                                    If Angle > 45 Then
'                                        Angle = Angle - 45
'                                        p2.X = 0
'                                        p2.Y = 2
'                                    Else
'                                        p2.X = -2
'                                        p2.Y = 0
'                                        Swap p1, p2
'                                    End If
'                                Else
'                                    p1.Y = 1
'                                    p1.X = 1
'                                    Angle = Round(Angle - ((Angle \ 90) * 90), 2)
'                                    If Angle > 45 Then
'                                        Angle = Angle - 45
'                                        p2.X = 2
'                                        p2.Y = 0
'                                    Else
'                                        p2.X = 0
'                                        p2.Y = 2
'                                        Swap p1, p2
'                                    End If
'                                End If
'                                Set tmp = DistanceSet(p1, p2, DistanceEx(p1, p2) * (Angle / 45))
'                                Set tmp = DistanceSet(MakePoint(0, 0, 0), MakePoint(tmp.X, tmp.Y, 0), dist1)
'                        End Select
'                        Point.X = tmp.X '+ 1
'                        Point.Y = tmp.Y '+ 1
'                        Point.z = Point.z + 1
'                    End If

                End If
                
            End If
        End If
    
        If stack = 1 Then 'undo
            Point.X = Point.X - 1
            Point.Y = Point.Y - 1
            Point.z = Point.z - 1
          
    
'            Dim p1 As New Point
'            Dim p2 As New Point
'            Dim p3 As New Point
'           ' Dim tmp As Point
'
'            p1.X = .X
'            p1.Y = .Y
'            p1.z = .z
'
'            p2.X = .z
'            p2.Y = .X
'            p2.z = .Y
'
'            p3.X = .Y
'            p3.Y = .z
'            p3.z = .X
'
'            Set tmp = TriangleAccordance(p1, p2, p3)
'            .X = tmp.X
'            .Y = tmp.Y
'            .z = tmp.z
            
        End If
    End With
    stack = stack - 1
    End Function

Private Function AngleFunction(ByRef Coord As Coord, ByRef RetType As AngleValue) As Single
    Dim X As Single
    Dim Y As Single
    'round them off for checking
    X = Round(Coord.X, 6)
    Y = Round(Coord.Y, 6)
    'returns the z axis angle of the x and y in p
    If (X = 0) Then 'slope of 1
        'must be 360 or 180
        If (RetType = AngleValue.Base) Or (RetType = AngleValue.Angle) Then
            If (Y > 0) Then
                AngleFunction = (180 * RADIAN)
            ElseIf (Y < 0) Then
                AngleFunction = (360 * RADIAN)
            End If
        End If
    ElseIf (Y = 0) Then 'slope of 0
        'must be 90 or 270
        If (RetType = AngleValue.Base) Or (RetType = AngleValue.Angle) Then
            If (X > 0) Then
                AngleFunction = (90 * RADIAN)
            ElseIf (X < 0) Then
                AngleFunction = (270 * RADIAN)
            End If
        End If
    Else
        'get the base angle
        '(up to the quardrant)
        If (RetType = AngleValue.Base) Or (RetType = AngleValue.Angle) Then
            If ((X > 0) And (Y > 0)) Then
                AngleFunction = (90 * RADIAN)
            ElseIf ((X < 0) And (Y > 0)) Then
                AngleFunction = (180 * RADIAN)
            ElseIf ((X < 0) And (Y < 0)) Then
                AngleFunction = (270 * RADIAN)
            ElseIf ((X > 0) And (Y < 0)) Then
                AngleFunction = (360 * RADIAN)
            End If
        End If
        'all the trickery
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
        slope = (Least / Large) 'the angle in square form
        '^^ or tangent, tangable to other axis angles' shared axis
        If (RetType = AngleValue.Tangent) Then
            AngleFunction = slope
            Exit Function
        End If

        Angle = (((Coord.X ^ 2) + (Coord.Y ^ 2)) ^ (1 / 2)) 'distance
        Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
        Least = (((Angle ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
        Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / Angle)) * (Least / Angle))
        '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
        'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
        If (RetType = AngleValue.Cosine) Then
            AngleFunction = Least
            Exit Function
        End If

        Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope) 'bulk sine of the angle in 45 degree slices
        '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
        If (RetType = AngleValue.Sine) Then
            AngleFunction = Large
            Exit Function
        End If

        Angle = Large + Least
        If Not ((((X > 0 And Y > 0) Or (X < 0 And Y < 0)) And (Abs(Y) < Abs(X))) Or _
           (((X < 0 And Y > 0) Or (X > 0 And Y < 0)) And (Abs(Y) > Abs(X)))) Then
            Angle = (PI / 4) - Angle
            If (RetType = AngleValue.Base) Or (RetType = AngleValue.Angle) Then AngleFunction = AngleFunction + (PI / 4)
        End If

        If (RetType = AngleValue.Angle) Then
            AngleFunction = AngleFunction + Angle
        End If
    End If
End Function

'Public Function AnglesOfPoint(ByRef Point As Point) As Point
'    Static stack As Integer
'    stack = stack + 1
'    If stack = 1 Then
'        '(1,1,1) is high noon
'        'to 45 degree sections
'        'when going clockwise
'        Point.X = Point.X + 1
'        Point.Y = Point.Y + 1
'        Point.z = Point.z + 1
'    End If
'    Set AnglesOfPoint = New Point
'    With AnglesOfPoint
'        If stack < 5 Then
'            Dim X As Single
'            Dim Y As Single
'            Dim z As Single
'            'round them off for checking
'            '(6 is for single precision)
'            X = Round(Point.X, 6)
'            Y = Round(Point.Y, 6)
'            z = Round(Point.z, 6)
'            If (X = 0) Then
'                If (z = 0) Then
'                    'must be 360 or 180
'                    If (Y > 0) Then
'                        .z = (180 * RADIAN)
'                    ElseIf (Y < 0) Then
'                        .z = (360 * RADIAN)
'                    End If
'                Else
'                    AnglesOfPoint.X = Point.Y
'                    AnglesOfPoint.Y = Point.z
'                    AnglesOfPoint.z = Point.X
'                    .X = AnglesOfPoint(AnglesOfPoint).z
'                End If
'            ElseIf (Y = 0) Then
'                If (z = 0) Then
'                    'must be 90 or 270
'                    If (X > 0) Then
'                        .z = (90 * RADIAN)
'                    ElseIf (X < 0) Then
'                        .z = (270 * RADIAN)
'                    End If
'                Else
'                    AnglesOfPoint.X = Point.z
'                    AnglesOfPoint.Y = Point.X
'                    AnglesOfPoint.z = Point.Y
'                    .Y = -AnglesOfPoint(AnglesOfPoint).z
'                End If
'            ElseIf (X <> 0) And (Y <> 0) Then
'                Dim slope As Single
'                Dim dist As Single
'                Dim Large As Single
'                Dim Least As Single
'                Dim Angle As Single
'                'find the larger coordinate
'                If Abs(Point.X) > Abs(Point.Y) Then
'                    Large = Abs(Point.X)
'                    Least = Abs(Point.Y)
'                Else
'                    Least = Abs(Point.X)
'                    Large = Abs(Point.Y)
'                End If
'                slope = (Least / Large) 'the angle in square form
'                '^^ or tangent, tangable to other axis angles' shared axis
'                dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2)) 'distance
'                'still traveling for tangents and cosines
'                Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'hypotenus, acute distance
'                Least = (((dist ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, obtuse to the hypotneus and distance
'                Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / dist)) * (Least / dist))
'                '^^ rounding remainder cosine of the angle, to make up for the bulk sine not suffecient a curve
'                'in 16's, we are also adding the two degrees that are one removed from the pi in 4's done next
'                Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)  'bulk sine of the angle in 45 degree slices
'                '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
'                If (z <> 0) Then 'two or less axis is one rotation
'                    Dim Ret As Point
'                    AnglesOfPoint.X = Point.Y
'                    AnglesOfPoint.Y = Point.z
'                    AnglesOfPoint.z = Point.X
'                    Set Ret = AnglesOfPoint(AnglesOfPoint)
'                    .X = Ret.Y
'                    .Y = Ret.z
'                    .z = Ret.X
'                    Set Ret = Nothing
'                End If
'                'get the base angle
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
'                'develop the final angle Z for this duel coordinate X,Y axis only
'                Angle = (Large + Least)
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
'                If stack = 1 Then
'                    'reorganization
'                    Angle = .Y
'                    .Y = .z
'                    .z = Angle
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
'        End If
'    End With
'    If stack = 1 Then 'undo
'        Point.X = Point.X - 1
'        Point.Y = Point.Y - 1
'        Point.z = Point.z - 1
'    End If
'    stack = stack - 1
'End Function
'Private Function SwapAxisLeft(ByRef Point As Point)
'    Dim W As Single
'    W = Point.X
'    Point.X = Point.Y
'    Point.Y = Point.z
'    Point.z = W
'End Function
'
'Private Function SwapAxisRight(ByRef Point As Point)
'    Dim W As Single
'    W = Point.z
'    Point.z = Point.Y
'    Point.Y = Point.X
'    Point.X = W
'End Function
'
'Private Function AngleRotation(ByRef Point As Point, Optional ByRef Angles As Point) As Point
'    'returns the angles of Point *if Angles is supplied, the Point is modified and
'    'rotated by each axis of Angles, the return is not different if Angles is present
'    Static stack As Integer
'    If (stack < 0) Or (Not (Angles Is Nothing)) Then
'        stack = stack - 1 'using stack in reverse if angles is passed
'        'then we are also rotating Point after calculating it's angles
'        If Angles Is Nothing Then Set Angles = New Point
'    Else 'otherwise stack forward if not also rotating a point
'        stack = stack + 1
'    End If
'    Dim z As Single
'    Set AngleRotation = New Point
'    With AngleRotation
'        If Abs(stack) < 5 Then 'most things are stopped on the
'            '4th pass the fith we aren't doing anything beyond
'            Dim X As Single
'            Dim Y As Single
'            Dim Ret As Point
'            'round them off for checking (6 is for single precision)
'            'you could pass 0 as the second arg for a snap-to effect
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
'                    SwapAxisLeft Point
'                    SwapAxisLeft Angles
'                    Set Ret = AngleRotation(Point, Angles)
'                    SwapAxisRight Point
'                    SwapAxisRight Angles
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
'                    SwapAxisLeft Point
'                    SwapAxisLeft Angles
'                    Set Ret = AngleRotation(Point, Angles)
'                    SwapAxisRight Point
'                    SwapAxisRight Angles
'                End If
'            ElseIf (X <> 0) And (Y <> 0) Then 'these assumptions
'                'are already checked and to escape division by zero
'                Dim dist As Single
'                Dim slope As Single
'                Dim Large As Single
'                Dim Least As Single
'                Dim Angle As Single
'                'find the larger coordinate
'                If Abs(Point.X) > Abs(Point.Y) Then
'                    Large = Abs(Point.X)
'                    Least = Abs(Point.Y)
'                Else
'                    Least = Abs(Point.X)
'                    Large = Abs(Point.Y)
'                End If
'                slope = (Least / Large) 'the angle in square form
'                '^^ or tangent, "tangable" to other axis angles'
'                'i.e. x/y, shares x in z/x, and y in y/z views
'                dist = (((Point.X ^ 2) + (Point.Y ^ 2)) ^ (1 / 2))
'                '^^^distance, from (0,0,0) to (0,0) leaving out Z
'                Large = (((Large ^ 2) - (Least ^ 2)) ^ (1 / 2))
'                '^^^hypotenus made of the larger coord less the
'                'least coord, will always be smaller then dist
'                Least = (((dist ^ 2) - (Least ^ 2)) ^ (1 / 2)) 'arc, between the smaller hypotenus and dist
'                'this will always be median, smaller to dist, but bigger then "large" hypotneus solved for
'                Least = (((((((PI / 16) * DEGREE) + 2) * RADIAN) * slope) * (Large / dist)) * (Least / dist))
'                '^^ rounding remainder of the angle, to make up for the bulk not suffecient a curve in 16th's
'                'we are also adding the two degrees that are one removed from the pi in 4th's done next step
'                Large = (((((PI / 4) * DEGREE) - 1) * RADIAN) * slope)  'bulk of the angle in 45 degree slices
'                '^^ where as 0 and 45 are not logical angles, as they blend portion of neighboring 45 degree slices
'                'develop the 0 to 45 degree angle, still no identity
'                'and is just the same if it is to be a 45 to 90 angle
'                Angle = (Least + Large)
'                If (z <> 0) Then 'two or less axis is one rotation
'                    'and three axis rotations can be done in two axis
'                    SwapAxisLeft Point
'                    SwapAxisLeft Angles
'                    Set Ret = AngleRotation(Point, Angles)
'                    SwapAxisRight Point
'                    SwapAxisRight Angles
'                End If
'                'get the base angle
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
'                'develop the final angle Z for this duel coordinate X,Y axis only
'                Angle = (Large + Least)
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
'                If stack < 0 And Abs(stack) < 5 And dist > 0 Then
'                    'if the angles.z argument is
'                    'passed we will do the rotate
'                    'reusing some variables:
'                    If (.z + Angles.z) <> 0 Then
'                        Dim Angle2 As Single
'                        Dim Dist2 As Single
'                        Dim Large2 As Single
'                        Angle = (.z + Angles.z) * DEGREE
'                        Do While Angle > 360
'                            Angle = Angle - 360
'                        Loop
'                        Do While Angle <= 0
'                            Angle = Angle + 360
'                        Loop
'                        Angle2 = ((Angle - (Angle \ ((PI / 2) * DEGREE)) * _
'                                ((PI / 2) * DEGREE)) / ((PI / 2) * DEGREE))
'                        Dist2 = (((dist ^ 2) * 2) ^ (1 / 2))
'                        If ((Angle > 0) And (Angle <= 90)) Then
'                            Large = (Dist2 * Angle2)
'                            Least = -(Dist2 * (1 - Angle2))
'                        ElseIf ((Angle > 90) And (Angle <= 180)) Then
'                            Large = (Dist2 * (1 - Angle2))
'                            Least = (Dist2 * Angle2)
'                       ElseIf ((Angle > 180) And (Angle <= 270)) Then
'                            Large = -(Dist2 * Angle2)
'                            Least = (Dist2 * (1 - Angle2))
'                        ElseIf ((Angle > 270) And (Angle <= 360)) Then
'                            Large = -(Dist2 * (1 - Angle2))
'                            Least = -(Dist2 * Angle2)
'                        End If
'                        Dist2 = ((((Large ^ 2) + (Least ^ 2)) ^ (1 / 2)) - dist)
'                        Point.X = (Large - (Dist2 * (Large / dist) * (PI / 4)))
'                        Point.Y = (Least - (Dist2 * (Least / dist) * (PI / 4)))
'                    End If
'                End If
'                'reorganization
'                If Abs(stack) = 2 Then
'                    z = .z
'                    .z = .Y
'                    .Y = .X
'                    .X = z
'                ElseIf Abs(stack) = 3 Then
'                    z = .Y
'                    .Y = .z
'                    .z = z
'                ElseIf Abs(stack) = 1 Then
'                    z = .z
'                    .z = .Y
'                    .Y = .X
'                    .X = z
'                    z = .z
'                    .z = .X
'                    .X = z
'                End If
'            End If
'            'sorting out the recursion returns
'            If Not Ret Is Nothing Then
'                If Abs(stack) = 3 Then
'                    .Y = -Ret.z
'                End If
'                If Abs(stack) = 2 Then
'                    .X = -Ret.z
'                    .Y = Ret.Y
'                End If
'                If Abs(stack) = 1 Then
'                    .X = Ret.X
'                    .Y = Ret.Y
'                    .z = -Ret.z
'                End If
'                Set Ret = Nothing
'            End If
'        End If
'    End With
'    If (stack < 0) Then
'        stack = stack + 1
'    Else
'        stack = stack - 1
'    End If
'End Function





Public Function VectorAxisAngles(ByRef Point As Point) As Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        Dim magnitude As Single
        Dim heading As Single
        Dim pitch As Single
        Dim slope As Single

        slope = VectorSlope(MakePoint(0, 0, 0), Point)
'        If slope = 0 Then slope = -1
'        If slope > 0.5 Then slope = -slope
        magnitude = ((Point.X ^ 2 + Point.Y ^ 2 + Point.z ^ 2) ^ (1 / 2))
        If magnitude < 100 Then magnitude = 100
        heading = ATan2(Point.z, Point.X)
        pitch = ATan2(Point.Y, (Point.X * Point.X + Point.z * Point.z) ^ (1 / 2))
        .X = (((heading / magnitude) - pitch) * (slope / magnitude))
        .z = ((PI / 2) + (-pitch + (heading / magnitude))) * (1 - (slope / magnitude))
        .Y = ((-heading + (pitch / magnitude)) * (1 - (slope / magnitude)))
        .Y = -(.Y + ((.X * (slope / magnitude)) / 2) - (.Y * 2) - ((.z * (slope / magnitude)) / 2))
        .X = (PI * 2) - (.X - ((PI / 2) * (slope / magnitude)))
        .z = (PI * 2) - (.z - ((PI / 2) * (slope / magnitude)))

    End With

End Function

Public Sub VectorAxisSwapLeft(ByRef Point As Point)
    Dim tmp As Single
    tmp = Point.X
    Point.X = Point.Y
    Point.Y = Point.z
    Point.z = tmp
End Sub


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

Public Function VectorRise(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Single
    VectorRise = (Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y))
End Function
Public Function VectorRun(ByRef p1 As Point, Optional ByRef p2 As Point = "[0,0,0]") As Single
    VectorRun = DistanceEx(MakePoint(p1.X, 0, p1.z), MakePoint(p2.X, 0, p2.z))
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

    AngleInvertRotation = (-(PI * 2) - A + (PI * 4)) ' - PI

End Function
Public Function AngleAddition(ByVal a1 As Single, ByVal a2 As Single) As Single
    AngleAddition = a1 + a2
    AngleAxisRestrict AngleAddition
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

Public Function AngleAxisCombine(ByRef p1 As Point, ByRef p2 As Point) As Point
    Dim p3 As New Point
    Dim P4 As New Point
    Set p3 = AngleAxisInvert(p1)
    Set P4 = AngleAxisInvert(p2)
    Set AngleAxisCombine = New Point
    With AngleAxisCombine
    
        Set p3 = AngleAxisDeduction(p1, p2)
        Set P4 = AngleAxisDifference(p1, p3)
       .X = P4.X
       .Y = P4.Y
       .z = -P4.z
        
        
       ' .X = ((p1.X * p2.X + p3.X * P4.X + p1.X * p3.X + p2.X * P4.X) ^ (1 / 4))
       ' .Y = ((p1.Y * p2.Y + p3.Y * P4.Y + p1.Y * p3.Y + p2.Y * P4.Y) ^ (1 / 4))
       ' .z = ((p1.z * p2.z + p3.z * P4.z + p1.z * p3.z + p2.z * P4.z) ^ (1 / 4))
        
        
    End With
    AngleAxisRestrict AngleAxisCombine
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
    If VectorIsNormal Then Exit Function
    VectorIsNormal = (DistanceEx(MakePoint(0, 0, 0), p1) = 1)  'another is the total length of vector is one
    If VectorIsNormal Then Exit Function
    'another is if any value exists non zero as well as adding up in any non specific arrangement cancels to zero, as has one
    VectorIsNormal = ((p1.X <> 0 Or p1.Y <> 0 Or p1.z <> 0) And (( _
        ((p1.X + p1.Y + p1.z) = 0) Or ((p1.Y + p1.z + p1.X) = 0) Or ((p1.z + p1.X + p1.Y) = 0) Or _
        ((p1.X + p1.z + p1.Y) = 0) Or ((p1.z + p1.Y + p1.X) = 0) Or ((p1.Y + p1.X + p1.z) = 0) _
        )))
    If VectorIsNormal Then Exit Function
    'triangle's normal, only the sides are expressed upon each axis
    VectorIsNormal = ((((p1.X - p1.Y) + p1.z) + ((p1.Y - p1.z) + p1.X) + ((p1.z - p1.X) + p1.Y)) = 1)
    If VectorIsNormal Then Exit Function
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
    VectorIsNormal = ((p1.X <> 0 Or p1.Y <> 0 Or p1.z <> 0) And (tmp >= -1 And tmp <= 1))
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
    If (p.Y > 0 And p.X >= 0) Or (p.Y >= 0 And p.X > 0) Then
        VectorQuadrant = 1
    ElseIf (p.Y >= 0 And p.X < 0) Or (p.Y > 0 And p.X <= 0) Then
        VectorQuadrant = 2
    ElseIf (p.Y < 0 And p.X <= 0) Or (p.Y <= 0 And p.X < 0) Then
        VectorQuadrant = 3
    ElseIf (p.Y <= 0 And p.X > 0) Or (p.Y < 0 And p.X >= 0) Then
        VectorQuadrant = 4
    End If
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

