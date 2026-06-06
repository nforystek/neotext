Attribute VB_Name = "modMaths"


Option Explicit

Public Function DistanceBetweenTwo3DPoints(p1 As D3DVECTOR, p2 As D3DVECTOR) As Single
'##########################################
'#                                        #
'# This is just really Pythagoras Theorm  #
'#                                        #
'##########################################

    Dim tmpVector As D3DVECTOR
    
    tmpVector.X = p2.X - p1.X
    tmpVector.Y = p2.Y - p1.Y
    tmpVector.Z = p2.Z - p1.Z
    
    DistanceBetweenTwo3DPoints = Sqr(tmpVector.X * tmpVector.X + tmpVector.Y * tmpVector.Y + tmpVector.Z * tmpVector.Z)

End Function
'############################################################################################################
'#                                                                                                          #
'# Function Name    :   Ray Intersect Plane                                                                 #
'# Desc             :   Checks to see if a ray has intersected a plane                                      #
'#                                                                                                          #
'#                      Plane  = The plane to test against                                                  #
'#                      PStart = The starting point of the ray                                              #
'#                      VDir   = A vector representing the direction of the ray                             #
'#                      VIntersectOut = A variable that returns the coordinates of the intersection         #
'#                                                                                                          #
'# Returns          :   TRUE if a collision occurs                                                          #
'#                      FALSE if there is no collision                                                      #
'#                                                                                                          #
'############################################################################################################

Public Function RayIntersectPlane(Plane As D3DVECTOR4, PStart As D3DVECTOR, vDir As D3DVECTOR, ByRef VIntersectOut As D3DVECTOR) As Boolean
    Dim q As D3DVECTOR4     'Start Point
    Dim v As D3DVECTOR4     'Vector Direction

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
    
    planeVdot = D3DXVec4Dot(Plane, v)
    planeQdot = D3DXVec4Dot(Plane, q)
            
    'If the dotproduct of plane and V = 0 then there is no intersection
    If planeVdot <> 0 Then
        t = Round((planeQdot / planeVdot) * -1, 5)
        
        'This is where the line intersects the plane
        VIntersectOut.X = Round(q.X + (t * v.X), 5)
        VIntersectOut.Y = Round(q.Y + (t * v.Y), 5)
        VIntersectOut.Z = Round(q.Z + (t * v.Z), 5)

        RayIntersectPlane = True
    Else
        'No Collision
        RayIntersectPlane = False
    End If
    
End Function
'#########################################################################
'#                                                                       #
'# Function Name    :   Point In Triangle                                #
'# Desc             :   Function tests if a given point is within        #
'#                      a triangle as defined by verticies V1, V2 and V3 #
'#                      P is the point to be tested                      #
'#                                                                       #
'#########################################################################
Public Function PointInTriangle(V1 As D3DVECTOR, V2 As D3DVECTOR, V3 As D3DVECTOR, p As D3DVECTOR)
    'Project all vectors onto the xy plane
    'making this a 2d problem instead of a 3d problem
    
    'If the point is within the triangle in 3d space, it is also within the triangle in 2d
    'space. So it is ok to do this method.
    
    'Note in order for the point in triangle function to work correctly v1, v2 and v3
    'must be in clockwise order OR IT WILL NOT WORK
    
    Dim vertex1 As D3DVECTOR2   'Our 2D verticies
    Dim vertex2 As D3DVECTOR2
    Dim vertex3 As D3DVECTOR2
    
    Dim edge1 As D3DVECTOR2     'Our 2D vectors
    Dim edge2 As D3DVECTOR2
    Dim edge3 As D3DVECTOR2
    
    Dim NormEdge1 As D3DVECTOR2 'Vectors perpendicular to to the edge vectors
    Dim NormEdge2 As D3DVECTOR2
    Dim NormEdge3 As D3DVECTOR2
    
    Dim testvec1 As D3DVECTOR2  'A vector drawn between a triangle vertex and the test point
    Dim testvec2 As D3DVECTOR2
    Dim testvec3 As D3DVECTOR2
    
    Dim dot1 As Single          'Dot products can tell us the angle between two vectors
    Dim dot2 As Single          'they can also tell us what side of a plane a point is on.
    Dim dot3 As Single
    
    Dim PointVertex As D3DVECTOR2   'A 2d version of our testpoint
    
    Dim TriangleNormal As D3DVECTOR         'Which direction is our 3d triangle pointing?
    Dim ABSTriangleNormal As D3DVECTOR      'Absolute values of the above
    
    'Get TriangleNormal
    TriangleNormal = PlaneNormal(V1, V2, V3)

    'What is the greatest absolute value
    ABSTriangleNormal.X = Abs(TriangleNormal.X)
    ABSTriangleNormal.Y = Abs(TriangleNormal.Y)
    ABSTriangleNormal.Z = Abs(TriangleNormal.Z)
    
    'Discard the greatest absolute value and project onto the other
    'remaining planes.
    
    If ABSTriangleNormal.X > ABSTriangleNormal.Y And ABSTriangleNormal.X > ABSTriangleNormal.Z Then
        
        Project3dvectorYZplane PointVertex, p
        Project3dvectorYZplane vertex1, V1
        Project3dvectorYZplane vertex2, V2
        Project3dvectorYZplane vertex3, V3
    
    Else
        If (ABSTriangleNormal.Y > ABSTriangleNormal.X) And (ABSTriangleNormal.Y > ABSTriangleNormal.Z) Then
            Project3dvectorXZplane PointVertex, p
            Project3dvectorXZplane vertex1, V1
            Project3dvectorXZplane vertex2, V2
            Project3dvectorXZplane vertex3, V3
        Else
            Project3dvectorXYplane PointVertex, p
            Project3dvectorXYplane vertex1, V1
            Project3dvectorXYplane vertex2, V2
            Project3dvectorXYplane vertex3, V3
        End If
    End If
    
    'Create the vectors from the verticies
    D3DXVec2Subtract edge1, vertex2, vertex1
    D3DXVec2Subtract edge2, vertex3, vertex2
    D3DXVec2Subtract edge3, vertex1, vertex3
    
    'Create vectors perpendicular to the edge vectors
    OrthangonaliseVec2 NormEdge1, edge1
    OrthangonaliseVec2 NormEdge2, edge2
    OrthangonaliseVec2 NormEdge3, edge3

    'Create vector between the vertex point and the testpoint
    D3DXVec2Subtract testvec1, vertex1, PointVertex
    D3DXVec2Subtract testvec2, vertex2, PointVertex
    D3DXVec2Subtract testvec3, vertex3, PointVertex
    
    
    'Calculate dot products
    dot1 = D3DXVec2Dot(testvec1, NormEdge1)
    dot2 = D3DXVec2Dot(testvec2, NormEdge2)
    dot3 = D3DXVec2Dot(testvec3, NormEdge3)
    
    If dot1 > 0 And dot2 > 0 And dot3 > 0 Then
        PointInTriangle = True
    End If
    
End Function

Public Sub OrthangonaliseVec2(ByRef vout As D3DVECTOR2, vin As D3DVECTOR2)
    vout.X = vin.Y * -1
    vout.Y = vin.X
End Sub

Public Sub Project3dvectorXYplane(ByRef vout As D3DVECTOR2, vin As D3DVECTOR)
    vout.X = vin.X
    vout.Y = vin.Y
End Sub
Public Sub Project3dvectorYZplane(ByRef vout As D3DVECTOR2, vin As D3DVECTOR)
    vout.X = vin.Z
    vout.Y = vin.Y
End Sub
Public Sub Project3dvectorXZplane(ByRef vout As D3DVECTOR2, vin As D3DVECTOR)
    vout.X = vin.X
    vout.Y = vin.Z
End Sub


Public Function Create4DPlaneVectorFromPoints(V1 As D3DVECTOR, V2 As D3DVECTOR, V3 As D3DVECTOR) As D3DVECTOR4
    Dim edge1 As D3DVECTOR
    Dim edge2 As D3DVECTOR
        
    Dim pNormal As D3DVECTOR
        
    D3DXVec3Subtract edge1, V2, V1
    D3DXVec3Subtract edge2, V3, V1
                        
    D3DXVec3Cross pNormal, edge1, edge2 'This is the normal vector
    D3DXVec3Normalize pNormal, pNormal  'This is the scaled normal vector
                        
    'Generate the 4D Plane Vector
    Create4DPlaneVectorFromPoints.r = D3DXVec3Dot(pNormal, V1) * -1
    Create4DPlaneVectorFromPoints.X = pNormal.X
    Create4DPlaneVectorFromPoints.Y = pNormal.Y
    Create4DPlaneVectorFromPoints.Z = pNormal.Z
    
End Function

'Written by someone named witchlord on the gamedev boards
Sub VectorMatrixMultiply(ByRef vDest As D3DVECTOR, ByRef vSrc As D3DVECTOR, ByRef mat As D3DMATRIX)
    
    Dim X As Single, Y As Single, Z As Single, W As Single
    X = vSrc.X * mat.m11 + vSrc.Y * mat.m21 + vSrc.Z * mat.m31 + mat.m41
    Y = vSrc.X * mat.m12 + vSrc.Y * mat.m22 + vSrc.Z * mat.m32 + mat.m42
    Z = vSrc.X * mat.m13 + vSrc.Y * mat.m23 + vSrc.Z * mat.m33 + mat.m43
    W = vSrc.X * mat.m14 + vSrc.Y * mat.m24 + vSrc.Z * mat.m34 + mat.m44

    If Abs(W) < epsilon Then Exit Sub

    vDest.X = X / W
    vDest.Y = Y / W
    vDest.Z = Z / W
End Sub

'Thanks to Jack Hoxley. I've adapted it to work with DirectX 8
Public Sub TranslateMatrix(pMatrix As D3DMATRIX, pVector As D3DVECTOR)
  D3DXMatrixIdentity pMatrix
  pMatrix.m41 = pVector.X
  pMatrix.m42 = pVector.Y
  pMatrix.m43 = pVector.Z
End Sub

'Public Function GenerateTriangleNormal(p0 As D3DVECTOR, p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR
''Variables required
'    Dim v01 As D3DVECTOR        'Vector from points 0 to 1
'    Dim v02 As D3DVECTOR        'Vector from points 0 to 2
'    Dim vNorm As D3DVECTOR      'The final vector
'
''Create the vectors from points 0 to 1 and 0 to 2
'    D3DXVec3Subtract v01, p1, p0
'    D3DXVec3Subtract v02, p2, p0
'
''Get the cross product
'    D3DXVec3Cross vNorm, v01, v02
'
'' Normalize this vector
'    D3DXVec3Normalize vNorm, vNorm
'
'' Return the value
'    GenerateTriangleNormal.X = vNorm.X
'    GenerateTriangleNormal.Y = vNorm.Y
'    GenerateTriangleNormal.z = vNorm.z
'
'End Function


