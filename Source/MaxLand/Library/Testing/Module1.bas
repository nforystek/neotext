Attribute VB_Name = "Module1"

Option Explicit


Public Enum CullingMethod
     CullByFlagSet = 0
     CullBySquares = 1
     CullByInsides = 2
     CullByRanging = 3
     CullByClosest = 4
     CullByCameras = 5
     CullByBehinds = 6
     UseAllCulling = 7
End Enum

'the first three parameters are (X,Y,Z) of the point to be checked against a traingle defined by its normal and center
Public Declare Function PointBehindPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                    ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single, _
                                    ByVal CenterX As Single, ByVal CenterY As Single, ByVal CenterZ As Single) As Boolean

Public Declare Function PointTouchesTriangle Lib "..\Release\maxland.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                    ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single, _
                                    ByVal CenterX As Single, ByVal CenterY As Single, ByVal CenterZ As Single) As Long
                                    
'this next function is a 2D test to see if a point is with in a complex closed shape made of polyDataCount number of
'(X,Y) points, stored in polyDataX and polyDataY, the result is the nearest coordinate where the point crosses over
Public Declare Function PointInPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, _
                                     polyDataX() As Single, polyDataY() As Single, ByVal polyDataCount As Long) As Long
                                    
Public Declare Function PointInsidePointList Lib "..\Release\maxland.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, _
                                    ByRef polyListX As Single, ByRef polyListY As Single, ByVal polyListaCount As Long) As Long
                                   ' polyDataX As Any, polyDataY As Any, ByVal polyDataCount As Long) As Long

'Test() depends on the functions results above, two views of 3d, x/y and y/z, are called for pointinpoly when
'satisfactionis returned, it is a single point nearest the point checked, then combined with point behindpoly
'and passed to test() the determination is complete by Test() results, as of now PointInPoly2 fails to inform

Public Declare Function Test Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean

Public Declare Function Test2 Lib "..\Release\maxland.dll" Alias "Test" _
                                    (ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Boolean
                                
     
''trI_tri_intersect() has been Depreciated to the declare just after it
'Public Declare Function tri_tri_intersect Lib "..\Backup\MaxLandLib.dll" _
'                                    (ByVal v0_0 As Single, ByVal v0_1 As Single, ByVal v0_2 As Single, _
'                                    ByVal v1_0 As Single, ByVal v1_1 As Single, ByVal v1_2 As Single, _
'                                    ByVal v2_0 As Single, ByVal v2_1 As Single, ByVal v2_2 As Single, _
'                                    ByVal u0_0 As Single, ByVal u0_1 As Single, ByVal u0_2 As Single, _
'                                    ByVal u1_0 As Single, ByVal u1_1 As Single, ByVal u1_2 As Single, _
'                                    ByVal u2_0 As Single, ByVal u2_1 As Single, ByVal u2_2 As Single) As Integer
'                                    'THAT WAS FUN

Public Declare Function TriangleCrossSegmentEx Lib "..\Release\maxland.dll" _
                                   (ByVal Ax1 As Single, ByVal Ay1 As Single, ByVal Az1 As Single, _
                                    ByVal Ax2 As Single, ByVal Ay2 As Single, ByVal Az2 As Single, _
                                    ByVal Ax3 As Single, ByVal Ay3 As Single, ByVal Az3 As Single, _
                                    ByVal Bx1 As Single, ByVal By1 As Single, ByVal Bz1 As Single, _
                                    ByVal Bx2 As Single, ByVal By2 As Single, ByVal Bz2 As Single, _
                                    ByVal Bx3 As Single, ByVal By3 As Single, ByVal Bz3 As Single, _
                                    ByRef Px0 As Single, ByRef Py0 As Single, ByRef Pz0 As Single, _
                                    ByRef Px1 As Single, ByRef Py1 As Single, ByRef Pz1 As Single) As Single
'
''Forystek() 3 variants of culling, vistype painting the canvas was the direction of able to process multiple
''views, at it's default I think is the most powerful (usually a check between two traingles of all) but it
''fails with a % resulting no triangles check depending on the camera, the other two are more secure, and all.
''and unfortunatly I had not got as far as I projected, so it only colors vistype once against all flags
'Public Declare Function Forystek Lib "MaxLandLib.dll" _
'                                    (ByVal visType As Long, _
'                                    ByVal lngFaceCount As Long, _
'                                    sngCamera() As Single, _
'                                    sngFaceVis() As Single, _
'                                    sngVertexX() As Single, _
'                                    sngVertexY() As Single, _
'                                    sngVertexZ() As Single, _
'                                    sngScreenX() As Single, _
'                                    sngScreenY() As Single, _
'                                    sngScreenZ() As Single, _
'                                    sngZBuffer() As Single) As Long

Public Declare Function CollisionCull Lib "..\Release\maxland.dll" _
                                    (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single, _
                                    ByRef ApplyCulling As Any) As Long

'this is the project purpose, the collision checker, I am certian when I did this one
'it was object count * two at the least for checking and then the trainlges do so too
'but the visType is a flag that only traingles with the flag are check for collision
Public Declare Function Collision Lib "MaxLandLib.dll" _
                                   (ByVal visType As Long, _
                                    ByVal lngFaceCount As Long, _
                                    sngFaceVis() As Single, _
                                    sngVertexX() As Single, _
                                    sngVertexY() As Single, _
                                    sngVertexZ() As Single, _
                                    ByVal lngFaceNum As Long, _
                                    ByRef lngCollidedBrush As Long, _
                                    ByRef lngCollidedFace As Long) As Boolean

'Public Declare Function Collision2 Lib "..\Release\maxland.dll" Alias "Collision" _
'                                   (ByVal visType As Long, _
'                                    ByVal lngFaceCount As Long, _
'                                    sngFaceVis() As Single, _
'                                    sngVertexX() As Single, _
'                                    sngVertexY() As Single, _
'                                    sngVertexZ() As Single, _
'                                    ByVal lngFaceNum As Long, _
'                                    ByRef lngCollidedBrush As Long, _
'                                    ByRef lngCollidedFace As Long) As Boolean

Public Declare Function CollisionObjectFlag Lib "..\Release\maxland.dll" _
                                   (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single, _
                                    ByVal ObjectIndex As Long) As Long
                                    
Public Declare Function CollisionTriangleFlag Lib "..\Release\maxland.dll" _
                                   (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single, _
                                    ByVal TriangleIndex As Long, _
                                    ByVal TriangleCount As Long) As Long
                                    
Public Declare Function CollisionResetFlag Lib "..\Release\maxland.dll" _
                                   (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single, _
                                    ByVal NewFlag As Long) As Long
                                    
Public Declare Sub CollisionClearFlag Lib "..\Release\maxland.dll" _
                                   (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single)
                                    
Public Declare Function CollisionCheck Lib "..\Release\maxland.dll" _
                                   (ByVal Flag As Long, _
                                    ByVal TriangleTotal As Long, _
                                    ByRef FaceVis As Single, _
                                    ByRef VertexX As Single, _
                                    ByRef VertexY As Single, _
                                    ByRef VertexZ As Single, _
                                    ByVal TriangleIndex As Long, _
                                    ByRef CollidedObjectIndex As Long, _
                                    ByRef CollidedTriangleIndex As Long) As Boolean


Public Declare Function Sign2 Lib "..\Release\maxland.dll" Alias "Sign" (ByVal n As Single) As Single


'The following variables are needed for Forystek() and Collision() culling and collision
'checking it is quite incompatable to prorietary needs (like doubling the data to use
'the functions vs however one has their data stored already could use)
Public TotalObjects As Long
Public TotalFaces As Long
Public TotalTriangles As Long
Public sngTriangleSetData() As Single
'sngTriangleFaceData dimension (,n) where n=# is triangle index
'sngTriangleFaceData dimension (n,) where n=0 is x of the triangle normal
'sngTriangleFaceData dimension (n,) where n=1 is y of the triangle normal
'sngTriangleFaceData dimension (n,) where n=2 is z of the triangle normal
'sngTriangleFaceData dimension (n,) where n=3 custom flag for segragation
'sngTriangleFaceData dimension (n,) where n=4 a object organization index
'sngTriangleFaceData dimension (n,) where n=5 is reserved for flag states

Public sngVertexXAxisData() As Single
'sngVertexXAxisData dimension (,n) where n=# is triangle index
'sngVertexXAxisData dimension (n,) where n=0 is X of the first vertex
'sngVertexXAxisData dimension (n,) where n=1 is X of the second vertex
'sngVertexXAxisData dimension (n,) where n=2 is X of the third vertex
Public sngVertexYAxisData() As Single
'sngVertexYAxisData dimension (,n) where n=# is triangle index
'sngVertexYAxisData dimension (n,) where n=0 is Y of the first vertex
'sngVertexYAxisData dimension (n,) where n=1 is Y of the second vertex
'sngVertexYAxisData dimension (n,) where n=2 is Y of the third vertex
Public sngVertexZAxisData() As Single
'sngVertexZAxisData dimension (,n) where n=# is triangle index
'sngVertexZAxisData dimension (n,) where n=0 is Z of the first vertex
'sngVertexZAxisData dimension (n,) where n=1 is Z of the second vertex
'sngVertexZAxisData dimension (n,) where n=2 is Z of the third vertex


Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Type Point
    x As Single
    y As Single
    z As Single
End Type
Public Type Triangle
    p1 As Point
    p2 As Point
    p3 As Point
    a As Point
    n As Point
    L As Point
End Type

Public Const Epsilon = 0.000001

Public Const Repeats = 5000

Public Sub Main()


    MakeTestData
    
    'Main0
    
   ' Main1
    'Main2
    
   ' Main3
   Main4
End Sub

Public Sub Main4()

    Dim angle As Single
    
    ' Plane normals
    angle = PlaneAngleToPlane(MakePoint(0, 0, 1), MakePoint(0, 0, 1))   ' Same direction ? 0 degrees
    Debug.Print "Angle = "; angle
    
    angle = PlaneAngleToPlane(MakePoint(0, 0, 1), MakePoint(0, 0, -1))  ' Opposite direction ? 180 degrees
    Debug.Print "Angle = "; angle
    
    angle = PlaneAngleToPlane(MakePoint(1, 0, 0), MakePoint(0, 1, 0))   ' Perpendicular ? 90 degrees
    Debug.Print "Angle = "; angle

    
End Sub
Public Function PlaneToPlaneBackface(ByRef n1 As Point, ByRef n2 As Point) As Boolean
    Dim b1 As Point
    b1 = MakePoint(n1.x + n2.x, n1.y + n2.y, n1.z + n2.z)
    PlaneNormalToPlaneTest = ((b1.x < 0.5 And b1.x > -0.5) And (b1.y < 0.5 And b1.y > -0.5) And (b1.z < 0.5 And b1.z > -0.5))
End Function

Public Sub Main3()
    ResetCollisionTestData
    
    
    AddCubeToCollision MakePoint(0, 0, 0), 20, 0
    
    AddTriangleToCollision MakePoint(-15, 0, 0), MakePoint(-15, 10, 0), MakePoint(-15, 0, 10), 0
    
    Debug.Print CollisionCull(1, TotalTriangles, sngTriangleSetData(0, 0), sngVertexXAxisData(0, 0), sngVertexYAxisData(0, 0), sngVertexZAxisData(0, 0), 6)
    Debug.Print CollisionCull(1, TotalTriangles, sngTriangleSetData(0, 0), sngVertexXAxisData(0, 0), sngVertexYAxisData(0, 0), sngVertexZAxisData(0, 0), UseAllCulling)
    Debug.Print CollisionCull(1, TotalTriangles, sngTriangleSetData(0, 0), sngVertexXAxisData(0, 0), sngVertexYAxisData(0, 0), sngVertexZAxisData(0, 0), CullByClosest)

End Sub

Public Sub MakeTestData()
    

    
    'AddCubeToCollision MakePoint(5, 5, 5), 20, 0

    
   ' AddCubeToCollision MakePoint(-40, 0, 0), 20, 0

End Sub

Public Sub AddTriangleToCollision(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, Optional ByVal Flag As Long = 0)

    ReDim Preserve sngTriangleSetData(0 To 5, 0 To TotalTriangles) As Single
    ReDim Preserve sngVertexXAxisData(0 To 2, 0 To TotalTriangles) As Single
    ReDim Preserve sngVertexYAxisData(0 To 2, 0 To TotalTriangles) As Single
    ReDim Preserve sngVertexZAxisData(0 To 2, 0 To TotalTriangles) As Single
    
    sngVertexXAxisData(0, TotalTriangles) = p1.x
    sngVertexYAxisData(0, TotalTriangles) = p1.y
    sngVertexZAxisData(0, TotalTriangles) = p1.z

    sngVertexXAxisData(1, TotalTriangles) = p2.x
    sngVertexYAxisData(1, TotalTriangles) = p2.y
    sngVertexZAxisData(1, TotalTriangles) = p2.z

    sngVertexXAxisData(2, TotalTriangles) = p3.x
    sngVertexYAxisData(2, TotalTriangles) = p3.y
    sngVertexZAxisData(2, TotalTriangles) = p3.z

    Dim pn As Point
    pn = TriangleNormal(p1, p2, p3)
    
    sngTriangleSetData(0, TotalTriangles) = pn.x
    sngTriangleSetData(1, TotalTriangles) = pn.y
    sngTriangleSetData(2, TotalTriangles) = pn.z
    
    sngTriangleSetData(3, TotalTriangles) = Flag
    sngTriangleSetData(4, TotalTriangles) = TotalObjects
    sngTriangleSetData(5, TotalTriangles) = TotalFaces

    TotalTriangles = TotalTriangles + 1
    TotalObjects = TotalObjects + 1

End Sub
Private Sub AddSquareToCollision(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, ByRef p4 As Point, Optional ByVal Flag As Long = 0)

    AddTriangleToCollision p1, p2, p3, Flag
    
    TotalObjects = TotalObjects - 1
                                            
    AddTriangleToCollision p2, p3, p4, Flag
    
    TotalFaces = TotalFaces + 1
End Sub
Private Sub AddCubeToCollision(ByRef Location As Point, ByVal WallSize As Single, Optional ByVal Flag As Long = 0)
    AddSquareToCollision _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + (WallSize / 2), Location.z + (WallSize / 2)), Flag
    
    AddSquareToCollision _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), Flag
    AddSquareToCollision _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), Flag
    AddSquareToCollision _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), Flag
    AddSquareToCollision _
                        MakePoint(Location.x + (WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + -(WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + (WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + (WallSize / 2)), Flag
    AddSquareToCollision _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.x + (WallSize / 2), Location.z + (WallSize / 2)), _
                        MakePoint(Location.x + -(WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), _
                        MakePoint(Location.x + (WallSize / 2), Location.y + (WallSize / 2), Location.z + -(WallSize / 2)), Flag
                            
    TotalObjects = TotalObjects + 1
End Sub

Private Sub PrintFlags()
    Dim cnt As Long
    Dim out As String
    
    For cnt = 0 To TotalTriangles - 1
        out = out & Trim(CStr(sngTriangleSetData(3, cnt)))
    Next
    Debug.Print TotalTriangles & " " & out
    Debug.Print
End Sub
Public Sub ResetCollisionTestData()
    TotalObjects = 0
    TotalFaces = 0
    TotalTriangles = 0
    Erase sngTriangleSetData
    Erase sngVertexXAxisData
    Erase sngVertexYAxisData
    Erase sngVertexZAxisData
End Sub
Public Sub Main0()
    ResetCollisionTestData

    AddCubeToCollision MakePoint(0, 0, 0), 20, 1

    AddCubeToCollision MakePoint(-50, 0, 0), 20, 2



    
    AddTriangleToCollision MakePoint(-20, 0, 0), MakePoint(-20, 20, 0), MakePoint(0, 5, 0), 1
  '  AddTriangleToCollision MakePoint(-20, 50, 0), MakePoint(-20, 70, 0), MakePoint(0, 50, 0), 1

    AddCubeToCollision MakePoint(50, 0, 0), 20, 3
    
    AddCubeToCollision MakePoint(0, -50, 0), 20, 1
    

    
    Dim CollidedObjectIndex As Long
    Dim CollidedTriangleIndex As Long
    
   ' CollisionClearFlag 1, TotalTriangles, sngTriangleSetData(0, 0), sngVertexXAxisData(0, 0), sngVertexYAxisData(0, 0), sngVertexZAxisData(0, 0)

    PrintFlags
    

    Dim elapse As Single
    Dim i As Long
    Dim ret As Boolean
    
    


    elapse = Timer
    For i = 1 To Repeats
        ret = Collision(1, TotalTriangles, sngTriangleSetData, sngVertexXAxisData, sngVertexYAxisData, sngVertexZAxisData, TotalTriangles - 1, CollidedObjectIndex, CollidedTriangleIndex)
    Next
    Debug.Print CollidedObjectIndex; CollidedTriangleIndex
    Debug.Print "Collision: " & (Timer - elapse)

    If ret Then
        Debug.Print "Collision Occurs: " & ret
    Else
        Debug.Print "No Collision."
    End If

    elapse = Timer
    For i = 1 To Repeats
        'ret = CollisionCheck(1, TotalTriangles, sngTriangleSetData(0, 0), sngVertexXAxisData(0, 0), sngVertexYAxisData(0, 0), sngVertexZAxisData(0, 0), TotalTriangles - 1, CollidedObjectIndex, CollidedTriangleIndex)
        'ret = CollisionCheckVB(1, TotalTriangles, sngTriangleSetData, sngVertexXAxisData, sngVertexYAxisData, sngVertexZAxisData, TotalTriangles - 1, CollidedObjectIndex, CollidedTriangleIndex)
        ret = CollisionCheck3(1, TotalTriangles, sngTriangleSetData, sngVertexXAxisData, sngVertexYAxisData, sngVertexZAxisData, 25, CollidedObjectIndex, CollidedTriangleIndex)
    Next
    Debug.Print CollidedObjectIndex; CollidedTriangleIndex
    Debug.Print "CollisionCheck: " & (Timer - elapse)

    If ret Then
        Debug.Print "Collision Occurs: " & ret
    Else
        Debug.Print "No Collision."
    End If

End Sub
Private Function SkipCheck(ByVal Flag As Long, ByRef FaceVis() As Single, ByVal TriangleIndex As Long, ByVal i As Long) As Boolean
    If (Abs(FaceVis(5, i)) <> Flag) Then FaceVis(5, i) = Flag
    SkipCheck = ((FaceVis(3, i) = Flag) And (FaceVis(4, i) <> FaceVis(4, TriangleIndex)) And (FaceVis(5, i) = Flag))
End Function

Public Function CollisionCheckVB(ByVal Flag As Long, ByVal TriangleTotal As Long, ByRef FaceVis() As Single, ByRef VertexX() As Single, ByRef VertexY() As Single, ByRef VertexZ() As Single, ByVal TriangleIndex As Long, ByRef CollidedObjectIndex As Long, ByRef CollidedTriangleIndex As Long) As Boolean

    Dim i As Long, j As Long
    Dim nx As Single, ny As Single, nz As Single
    Dim lx As Single, ly As Single, lz As Single
    Dim cx As Single, cy As Single, cz As Single
    Dim p1 As Point, p2 As Point, p3 As Point
    
    Do While (i < TriangleTotal)
        
        If SkipCheck(Flag, FaceVis, TriangleIndex, i) Then
        
            p1 = MakePoint(VertexX(0, i), VertexY(0, i), VertexZ(0, i))
            p2 = MakePoint(VertexX(1, i), VertexY(1, i), VertexZ(1, i))
            p3 = MakePoint(VertexX(2, i), VertexY(2, i), VertexZ(2, i))
    
            lx = DistanceEx(p1, p2)
            ly = DistanceEx(p2, p3)
            lz = DistanceEx(p3, p1)

            cx = Least(p1.x, p2.x, p3.x)
            cy = Least(p1.y, p2.y, p3.y)
            cz = Least(p1.z, p2.z, p3.z)
            cx = (cx + ((Large(p1.x, p2.x, p3.x) - cx) / 2))
            cy = (cy + ((Large(p1.y, p2.y, p3.y) - cy) / 2))
            cz = (cz + ((Large(p1.z, p2.z, p3.z) - cz) / 2))
            
            nx = FaceVis(0, i)
            ny = FaceVis(1, i)
            nz = FaceVis(2, i)
            
            For j = 0 To 2
                If PointTouchesTriangle(VertexX(j, TriangleIndex) - cx, VertexY(j, TriangleIndex) - cy, VertexZ(j, TriangleIndex) - cz, lx, ly, lz, nx, ny, nz) = -1 Then
                    CollidedObjectIndex = FaceVis(4, i)
                    CollidedTriangleIndex = i
                    CollisionCheckVB = True
                    Exit Function
                End If
            Next

        End If
        i = i + 3
    Loop
    CollisionCheckVB = False
End Function


Public Function CollisionCheck3(ByVal Flag As Long, ByVal TriangleTotal As Long, ByRef FaceVis() As Single, ByRef VertexX() As Single, ByRef VertexY() As Single, ByRef VertexZ() As Single, ByVal TriangleIndex As Long, ByRef CollidedObjectIndex As Long, ByRef CollidedTriangleIndex As Long) As Boolean

    Dim start As Long, sstop As Long
    Dim i As Long, count As Long
    
    Do While (i < TriangleTotal)
        If SkipCheck(Flag, FaceVis, TriangleIndex, i) Then
        
            start = i
            sstop = i

            Do While SkipCheck(Flag, FaceVis, TriangleIndex, sstop)
                sstop = sstop + 1
                If (sstop >= TriangleTotal) Then Exit Do
            Loop
            
            If (sstop > start) Then
                sstop = (sstop - 1)
                count = (sstop - (start - 1))
                If ((start + (count - 1) <= TriangleTotal) And (start <= sstop)) Then
                
                
                   ' Debug.Print start; sstop; count
                  '  Stop 'todo
    
                    
                End If
                i = sstop
            End If
        End If
        i = i + 1
    Loop
    CollisionCheck3 = False
End Function


Public Function Sign(ByVal n As Double) As Double
    'returns the sign of any number which is the multiplication facttr of it's negative (*-1), zero(*0) or positive (*1)
    Sign = ((-(Abs((n * 99.99) - 1) - (n * 99.99)) - (-Abs((n * 99.99) + 1) + (n * 99.99))) * 0.5)
End Function

Public Function PointBehindPoly3(ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                ByVal Length1 As Single, ByVal Length2 As Single, ByVal Length3 As Single, _
                                ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single) As Boolean
    PointBehindPoly3 = (DistanceEx(MakePoint(0, 0, 0), MakePoint(PointX, PointY, PointZ)) <= _
        ((((Length1 + Length2) / 2) + ((Length1 + Length3) / 2) + ((Length2 + Length3) / 2)) / 3)) And _
         ((Sign(PointX) >= Sign(NormalX)) And (Sign(PointY) >= Sign(NormalY)) And (Sign(PointZ) >= Sign(NormalZ)))
End Function


Public Function PointSideOfPlane(ByRef V0 As Point, ByRef v1 As Point, ByRef v2 As Point, ByRef p As Point) As Boolean
    PointSideOfPlane = VectorDotProduct(PlaneNormal(V0, v1, v2), p) > 0
End Function

Public Sub Main2()
    
    Dim i As Long
    
    
    Dim t1 As Triangle
    Dim t2 As Triangle
    Dim p1 As Point
    Dim p2 As Point
    Dim ret As Single
    
    Dim c1 As Point
    Dim c2 As Point
    Dim n1 As Point
    Dim n2 As Point
    Dim l1 As Point
    Dim l2 As Point
    
    t1.p3 = MakePoint(0, 8, 0)
    t1.p2 = MakePoint(15, 7, 6)
    t1.p1 = MakePoint(4, 0, 14)
    t2.p3 = MakePoint(14, 0, -3)
    t2.p2 = MakePoint(6, 12, 1)
    t2.p1 = MakePoint(4, 12, 14)
    
    Dim elapse As Single
    
    elapse = Timer
    For i = 1 To Repeats
        ret = TriTriSegmentEx(t1.p1, t1.p2, t1.p3, t2.p1, t2.p2, t2.p3, p1, p2)
    Next
    Debug.Print "TriTriSegmentEx: " & (Timer - elapse)

    If ret Then
        Debug.Print "Intersection segment: " & ret
        Debug.Print "H=(10,5,0)=(" & Round(p1.x, 0) & "," & Round(p1.y, 0) & "," & Round(p1.z, 0) & ")"
        Debug.Print "I=(0,5,0)=(" & Round(p2.x, 0) & "," & Round(p2.y, 0) & "," & Round(p2.z, 0) & ")"
    Else
        Debug.Print "No intersection."
    End If


    elapse = Timer
    For i = 1 To Repeats
        ret = TriangleCrossSegmentEx(t1.p1.x, t1.p1.y, t1.p1.z, t1.p2.x, t1.p2.y, t1.p2.z, t1.p3.x, t1.p3.y, t1.p3.z, _
                                t2.p1.x, t2.p1.y, t2.p1.z, t2.p2.x, t2.p2.y, t2.p2.z, t2.p3.x, t2.p3.y, t2.p3.z, _
                                p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
    Next
    Debug.Print "TriangleCrossSegmentEx: " & (Timer - elapse)

    If ret Then
        Debug.Print "Intersection segment: " & ret
        Debug.Print "H=(8,7,3)=(" & Round(p1.x, 0) & "," & Round(p1.y, 0) & "," & Round(p1.z, 0) & ")"
        Debug.Print "I=(9,6,6)=(" & Round(p2.x, 0) & "," & Round(p2.y, 0) & "," & Round(p2.z, 0) & ")"
    Else
        Debug.Print "No intersection."
    End If
    

    t1.p1 = MakePoint(0, 0, 0)
    t1.p2 = MakePoint(20, 0, 0)
    t1.p3 = MakePoint(0, 20, 0)
    t2.p1 = MakePoint(-10, 5, 0)
    t2.p2 = MakePoint(10, 5, 10)
    t2.p3 = MakePoint(10, 5, -10)

    

    
    elapse = Timer
    For i = 1 To Repeats
        ret = TriTriSegmentEx(t1.p1, t1.p2, t1.p3, t2.p1, t2.p2, t2.p3, p1, p2)
    Next
    Debug.Print "TriTriSegmentEx: " & (Timer - elapse)
 
    If ret Then
        Debug.Print "Intersection segment: " & ret
        Debug.Print "H=(10,5,0)=(" & Round(p1.x, 0) & "," & Round(p1.y, 0) & "," & Round(p1.z, 0) & ")"
        Debug.Print "I=(0,5,0)=(" & Round(p2.x, 0) & "," & Round(p2.y, 0) & "," & Round(p2.z, 0) & ")"
    Else
        Debug.Print "No intersection."
    End If
 
 
    elapse = Timer
    For i = 1 To Repeats
        ret = TriangleCrossSegmentEx(t1.p1.x, t1.p1.y, t1.p1.z, t1.p2.x, t1.p2.y, t1.p2.z, t1.p3.x, t1.p3.y, t1.p3.z, _
                                t2.p1.x, t2.p1.y, t2.p1.z, t2.p2.x, t2.p2.y, t2.p2.z, t2.p3.x, t2.p3.y, t2.p3.z, _
                                p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
    Next
    Debug.Print "TriangleCrossSegmentEx: " & (Timer - elapse)

    If ret Then
        Debug.Print "Intersection segment: " & ret
        Debug.Print "H=(10,5,0)=(" & Round(p1.x, 0) & "," & Round(p1.y, 0) & "," & Round(p1.z, 0) & ")"
        Debug.Print "I=(0,5,0)=(" & Round(p2.x, 0) & "," & Round(p2.y, 0) & "," & Round(p2.z, 0) & ")"
    Else
        Debug.Print "No intersection."
    End If

End Sub


Public Sub Main1()
          
    'to test the new function DLL's outcome is the
    'same as the compiled one whose source is lost.
    
    Dim n1 As Integer
    Dim n2 As Integer
    Dim n3 As Integer
    Dim n4 As Integer
    Dim n5 As Integer
    Dim n6 As Integer
         
    Dim Px1 As Single
    Dim Py1 As Single
    Dim Pz1 As Single
    Dim vX1 As Single
    Dim vY1 As Single
    Dim vZ1 As Single
    Dim nX1 As Single
    Dim nY1 As Single
    Dim nZ1 As Single
    
    Dim pX2 As Single
    Dim pY2 As Single
    Dim pZ2 As Single
    Dim vX2 As Single
    Dim vY2 As Single
    Dim vZ2 As Single
    Dim nX2 As Single
    Dim nY2 As Single
    Dim nZ2 As Single
    
    Dim PointX As Single
    Dim PointY As Single
    Dim PointZ As Single

    Dim PointListsX(0 To 4) As Single
    Dim PointListsY(0 To 4) As Single
    Dim PointListsZ(0 To 4) As Single
    
    'make an 8x8 square whose center is (0,0)
    'going in the counter-clockwise direction

    Randomize

    Dim t1 As Triangle
    Dim t2 As Triangle
    
    Dim o1 As Point
    Dim o2 As Point
       
    PointListsX(0) = 4: PointListsY(0) = -4
    PointListsX(1) = 4: PointListsY(1) = 4
    PointListsX(2) = -4: PointListsY(2) = 4
    PointListsX(3) = -4: PointListsY(3) = -4
    PointListsX(4) = 4: PointListsY(4) = -4
        
    Dim testCount As Long
    
    Do
        testCount = testCount + 1
        
        Randomize
        DoEvents
        'this main loop is to debug all the frist three declared functions
        'to ensure they produce the same results as the original lost code

        With RandomPoint
            Px1 = .x
            Py1 = .y
            Pz1 = .z
        End With
        With RandomTriangle
            nX1 = .n.x
            nY1 = .n.y
            nZ1 = .n.z
            vX1 = .a.x
            vY1 = .a.y
            vZ1 = .a.z
        End With

        Debug.Print "PointBehindPoly()=" & PointBehindPoly(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1) & _
            " PointTouchesTriangle()=" & PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1) & _
            " PointBehindPoly3()=" & PointBehindPoly3(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)
        If (Not (CVar(PointBehindPoly(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)) = _
            CVar(PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)))) Or _
            (Not (CVar(PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)) = _
            CVar(PointBehindPoly3(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)))) Then testCount = -Abs(testCount)
        'Debug.Print

        'the box is 8x8 centered on (0,0) so we'll use
        'twice it's size and generate random within -8,8
        PointX = (RndNum(0, 16) - 8)
        PointY = (RndNum(0, 16) - 8)
        PointZ = (RndNum(0, 16) - 8)

        
'        Debug.Print "PointInPoly()=" & PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) & "  " & _
'            "PointInsidePointList()=" & PointInsidePointList(PointX, PointY, PointListsX(0), PointListsY(0), 5) & " " & _
'            "PointInPoly3()=" & PointInPoly3(PointX, PointY, PointListsX, PointListsY, 5)
        If (Not (PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) = _
            PointInsidePointList(PointX, PointY, PointListsX(0), PointListsY(0), 5))) Or _
             (Not (PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) = PointInPoly3(PointX, PointY, PointListsX, PointListsY, 5))) Then testCount = Abs(testCount)
        'Debug.Print
        
        'arbitrary arguments, unsigned short return values from PointInPoly that results a percentage with in the
        'scope of a integer max value from zero, indicating the point in the point list it falls inside the poly on
        n1 = Round(RndNum(0, 1), 0)
        n2 = Round(RndNum(0, 1), 0)
        n3 = Round(RndNum(0, 1), 0)


        'use the same square as if it is a cube in 3d,
        'and check each 2D axis for collision using test

        
        'Debug.Print "Test(n1, n2, n3)=" & Test(n1, n2, n3) & " Test2(n1, n2, n3)=" & Test2(n1, n2, n3)
        If Not CVar(Test(n1, n2, n3)) = CVar(Test2(n1, n2, n3)) Then testCount = -Abs(testCount)
        'Debug.Print


'        Debug.Print "Test(n1, n2, n3)=" & Test(n1, n2, n3) & " Test2(n1, n2, n3)=" & Test2(n1, n2, n3) & " Test3(n1, n2, n3)=" & Test3(n1, n2, n3)
'        If (Not (CVar(Test(n1, n2, n3)) = CVar(Test2(n1, n2, n3)))) Or (Not (CVar(Test2(n1, n2, n3)) = CVar(Test3(n1, n2, n3)))) Then Stop
'        Debug.Print

    Loop Until testCount > Repeats Or testCount < 0
    
    If testCount < 0 Then
        Debug.Print "Opps!  Discrepencies found!"
    Else
        Debug.Print "No discrepencies found."
    End If
    
    Debug.Print
End Sub



Public Function RndNum(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
    RndNum = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound) - 1
End Function


Public Function RandomPoint() As Point
    With RandomPoint
        .x = (RndNum(0, 16) - 8)
        .y = (RndNum(0, 16) - 8)
        .z = (RndNum(0, 16) - 8)
    End With
End Function

Public Function RandomTriangle() As Triangle

    With RandomTriangle

        .p1 = RandomPoint
        .p2 = RandomPoint
        .p3 = RandomPoint
        
        .a = TriangleAxii(.p1, .p2, .p3)
        
        .L.x = Distance(.p1, .p2)
        .L.y = Distance(.p2, .p3)
        .L.z = Distance(.p3, .p1)

        .n = PlaneNormal(.p1, .p2, .p3)
        
    End With

End Function


Public Function Test3(ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean
'I have been unsuccessful in the VB6 environment to get this one to act like Test() and Test2()
    Test3 = ((((n1 And n2 + n3) Or (n1 + n2 And n3)) And ((n1 - n2 Or Not n3) - (Not n1 Or n2 - n3))) _
        Or (((n1 - n2 Or n3) And (n1 - n2 Or n3)) + ((n1 Or n2 + Not n3) And (Not n1 + n2 And n3))))
End Function



Public Function PointInPoly3(ByVal Px As Single, ByVal Py As Single, polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long

    If (polyn > 2) Then
        Dim ref As Single
        Dim ret As Single
        Dim result As Long

        ref = ((Px - polyx(0)) * (polyy(1) - polyy(0)) - (Py - polyy(0)) * (polyx(1) - polyx(0)))
        ret = ref
        Dim i As Long
        For i = 1 To polyn - 1
            ref = ((Px - polyx(i)) * (polyy(i) - polyy(i - 1)) - (Py - polyy(i)) * (polyx(i) - polyx(i - 1)))
            If ((ret >= 0) And (ref < 0) And (result = 0)) Then
                result = i
            End If
            ret = ref
        Next
        If ((result = 0) Or (result > polyn)) Then
            PointInPoly3 = 1
        Else
            PointInPoly3 = 0
        End If
    End If

End Function




Public Function Distance(ByRef p1 As Point, ByRef p2 As Point) As Single
    Distance = (((p1.x - p2.x) ^ 2) + ((p1.y - p2.y) ^ 2) + ((p1.z - p2.z) ^ 2))
    If Distance <> 0 Then Distance = Distance ^ (1 / 2)
End Function

Public Function PlaneNormal(ByRef V0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    'returns a vector perpendicular to a plane V, at 0,0,0, with out the local coordinates information
    PlaneNormal = VectorCrossProduct(VectorDeduction(V0, v1), VectorDeduction(v1, v2))
End Function
Public Function MakePoint(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Point
    With MakePoint
        .x = x
        .y = y
        .z = z
    End With
End Function
Private Function VectorNormalize(a As Point) As Point
    Dim L As Single: L = DistanceEx(MakePoint(0, 0, 0), a)
    If L = 0 Then
        VectorNormalize = MakePoint(0, 0, 0)
    Else
        VectorNormalize = MakePoint(a.x / L, a.y / L, a.z / L)
    End If
End Function

Public Function VectorDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    With VectorDeduction
        .x = (p1.x - p2.x)
        .y = (p1.y - p2.y)
        .z = (p1.z - p2.z)
    End With
End Function

Public Function VectorCrossProduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    With VectorCrossProduct
        .x = ((p1.y * p2.z) - (p1.z * p2.y))
        .y = ((p1.z * p2.x) - (p1.x * p2.z))
        .z = ((p1.x * p2.y) - (p1.y * p2.x))
    End With
End Function
Public Function VectorDotProduct(a As Point, B As Point) As Single
    VectorDotProduct = a.x * B.x + a.y * B.y + a.z * B.z
End Function
Public Function VectorAddition(ByRef p1 As Point, ByRef p2 As Point) As Point
    With VectorAddition
        .x = (p1.x + p2.x)
        .y = (p1.y + p2.y)
        .z = (p1.z + p2.z)
    End With
End Function
Private Function TriangleNormal(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    Dim v1 As Point, v2 As Point
    v1 = VectorDeduction(p1, p2)
    v2 = VectorDeduction(p1, p3)
    TriangleNormal = VectorNormalize(VectorCrossProduct(v1, v2))
End Function
Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    With TriangleAxii
        Dim o As Point
        o = TriangleOffset(p1, p2, p3)
        .x = (Least(p1.x, p2.x, p3.x) + (o.x / 2))
        .y = (Least(p1.y, p2.y, p3.y) + (o.y / 2))
        .z = (Least(p1.z, p2.z, p3.z) + (o.z / 2))
    End With
End Function
Public Function TriangleOffset(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    With TriangleOffset
        .x = (Large(p1.x, p2.x, p3.x) - Least(p1.x, p2.x, p3.x))
        .y = (Large(p1.y, p2.y, p3.y) - Least(p1.y, p2.y, p3.y))
        .z = (Large(p1.z, p2.z, p3.z) - Least(p1.z, p2.z, p3.z))
    End With
End Function
Public Function DistanceEx(ByRef p1 As Point, ByRef p2 As Point) As Single
    DistanceEx = (((p1.x - p2.x) ^ 2) + ((p1.y - p2.y) ^ 2) + ((p1.z - p2.z) ^ 2))
    If DistanceEx <> 0 Then DistanceEx = DistanceEx ^ (1 / 2)
End Function
Public Function Large(ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (v1 >= v2) Then
            Large = v1
        Else
            Large = v2
        End If
    ElseIf IsMissing(V4) Then
        If ((v2 >= V3) And (v2 >= v1)) Then
            Large = v2
        ElseIf ((v1 >= V3) And (v1 >= v2)) Then
            Large = v1
        Else
            Large = V3
        End If
    Else
        If ((v2 >= V3) And (v2 >= v1) And (v2 >= V4)) Then
            Large = v2
        ElseIf ((v1 >= V3) And (v1 >= v2) And (v1 >= V4)) Then
            Large = v1
        ElseIf ((V3 >= v1) And (V3 >= v2) And (V3 >= V4)) Then
            Large = V3
        Else
            Large = V4
        End If
    End If
End Function

Public Function Least(ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If (v1 <= v2) Then
            Least = v1
        Else
            Least = v2
        End If
    ElseIf IsMissing(V4) Then
        If ((v2 <= V3) And (v2 <= v1)) Then
            Least = v2
        ElseIf ((v1 <= V3) And (v1 <= v2)) Then
            Least = v1
        Else
            Least = V3
        End If
    Else
        If ((v2 <= V3) And (v2 <= v1) And (v2 <= V4)) Then
            Least = v2
        ElseIf ((v1 <= V3) And (v1 <= v2) And (v1 <= V4)) Then
            Least = v1
        ElseIf ((V3 <= v1) And (V3 <= v2) And (V3 <= V4)) Then
            Least = V3
        Else
            Least = V4
        End If
    End If
End Function


Public Function AreParallel(t1p1 As Point, t1p2 As Point, t1p3 As Point, t2p1 As Point, t2p2 As Point, t2p3 As Point) As Boolean
    Dim n1 As Point, n2 As Point, cross As Point
    n1 = TriangleNormal(t1p1, t1p2, t1p3)
    n2 = TriangleNormal(t2p1, t2p2, t2p3)
    cross = VectorCrossProduct(n1, n2)
    AreParallel = (Abs(cross.x) < Epsilon And Abs(cross.y) < Epsilon And Abs(cross.z) < Epsilon)
End Function

Public Function AreCoplanar(t1p1 As Point, t1p2 As Point, t1p3 As Point, t2p1 As Point, t2p2 As Point, t2p3 As Point) As Boolean
    If Not AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3) Then
        AreCoplanar = False
        Exit Function
    End If
    
    Dim n1 As Point, d As Single
    n1 = TriangleNormal(t1p1, t1p2, t1p3)
    d = -(n1.x * t1p1.x + n1.y * t1p1.y + n1.z * t1p1.z)
    
    AreCoplanar = Abs(n1.x * t2p1.x + n1.y * t2p1.y + n1.z * t2p1.z + d) < Epsilon
End Function

Public Function AreParallelCoplanar(t1p1 As Point, t1p2 As Point, t1p3 As Point, t2p1 As Point, t2p2 As Point, t2p3 As Point) As Boolean
    Dim n1 As Point, n2 As Point, cross As Point
    Dim d As Single, p As Point
    
    ' Normals
    n1 = TriangleNormal(t1p1, t1p2, t1p3)
    n2 = TriangleNormal(t2p1, t2p2, t2p3)
    
    ' Cross product of normals
    cross = VectorCrossProduct(n1, n2)
    
    ' Plane constant from triangle 1
    d = -(n1.x * t1p1.x + n1.y * t1p1.y + n1.z * t1p1.z)
    
    ' Test point from triangle 2
    p = t2p1
    
    ' Single algebraic condition: parallel AND coplanar
    AreParallelCoplanar = _
        (Abs(cross.x) < Epsilon And Abs(cross.y) < Epsilon And Abs(cross.z) < Epsilon) _
        And (Abs(n1.x * p.x + n1.y * p.y + n1.z * p.z + d) < Epsilon)
End Function



' ===== Point-in-triangle test (barycentric) =====
Private Function PointInTriangle(p As Point, V0 As Point, v1 As Point, v2 As Point) As Boolean
    Dim u As Point, V As Point, w As Point
    u = VectorDeduction(v1, V0)
    V = VectorDeduction(v2, V0)
    w = VectorDeduction(p, V0)

    Dim uu As Single, vv As Single, uv As Single
    Dim wu As Single, wv As Single, d As Single

    uu = VectorDotProduct(u, u)
    vv = VectorDotProduct(V, V)
    uv = VectorDotProduct(u, V)
    wu = VectorDotProduct(w, u)
    wv = VectorDotProduct(w, V)

    d = uv * uv - uu * vv
    If Abs(d) < Epsilon Then
        PointInTriangle = False
        Exit Function
    End If

    Dim s As Single, t As Single
    s = (uv * wv - vv * wu) / d
    t = (uv * wu - uu * wv) / d

    PointInTriangle = (s >= -Epsilon And t >= -Epsilon And (s + t) <= 1 + Epsilon)
End Function

' ===== Edge-plane intersection =====
Private Function EdgePlaneIntersect(p As Point, Q As Point, planePoint As Point, PlaneNormal As Point, x As Point) As Boolean
    Dim dir As Point: dir = VectorDeduction(Q, p)
    Dim denom As Single: denom = VectorDotProduct(PlaneNormal, dir)
    If Abs(denom) < Epsilon Then
        EdgePlaneIntersect = False
        Exit Function
    End If

    Dim t As Single
    t = VectorDotProduct(PlaneNormal, VectorDeduction(planePoint, p)) / denom
    If t < -Epsilon Or t > 1 + Epsilon Then
        EdgePlaneIntersect = False
        Exit Function
    End If

    x = VectorAddition(p, MakePoint(dir.x * t, dir.y * t, dir.z * t))
    EdgePlaneIntersect = True
End Function

Public Function Acos(ByVal x As Double) As Double
    ' Clamp input to valid domain [-1, 1]
    If x < -1# Then x = -1#
    If x > 1# Then x = 1#

    ' acos(x) = atan2( sqrt(1 - x^2), x )
    Acos = Atn2(Sqr(1# - x * x), x)
End Function

Public Function PlaneAngleToPlane(ByRef n1 As Point, ByRef n2 As Point) As Single
    Dim dot As Single
    dot = n1.x * n2.x + n1.y * n2.y + n1.z * n2.z
    Dim mag1 As Single
    Dim mag2 As Single
    mag1 = (n1.x * n1.x + n1.y * n1.y + n1.z * n1.z) ^ (1 / 2)
    mag2 = (n2.x * n2.x + n2.y * n2.y + n2.z * n2.z) ^ (1 / 2)
    If (mag1 = 0 Or mag2 = 0) Then PlaneAngleToPlane = 0
    
    Dim cosTheta As Single
    cosTheta = dot / (mag1 * mag2)
    
    If (cosTheta > 1) Then cosTheta = 1
    If (cosTheta < -1) Then cosTheta = -1
    
    Dim angle As Single
    angle = Acos(cosTheta) * 180 / 3.14159265
    
    PlaneAngleToPlane = angle
End Function



Public Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 Then
        Atn2 = Atn(y / x) + Sgn(y) * 3.14159265358979
    Else
        Atn2 = Sgn(y) * 3.14159265358979 / 2
    End If
End Function



'##########################################################################
'##########################################################################
'##########################################################################



Public Function TriTriSegmentEx(ByRef t1p1 As Point, ByRef t1p2 As Point, ByRef t1p3 As Point, ByRef t2p1 As Point, ByRef t2p2 As Point, ByRef t2p3 As Point, ByRef OutP0 As Point, ByRef OutP1 As Point) As Single
    Dim ap As Boolean
    Dim ac As Boolean
    ap = AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3)
    ac = AreCoplanar(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3)
    Dim l1 As Single
    Dim l2 As Single

    If ap And Not ac Then
        TriTriSegmentEx = 0 'parallel triangles but not on the same plane and/or overlapping
    ElseIf ac Then
        'potentially parallel, but on the same plane at any rate, return the overlapping difference from a edge view of the mboth
        'because colliding triangles below are in the positive specture of a integers max value, this will be in the negative spec
        l1 = (DistanceEx(t1p1, t1p2) + DistanceEx(t1p2, t1p3) + DistanceEx(t1p3, t1p1))
        l2 = (DistanceEx(t2p1, t2p2) + DistanceEx(t2p2, t2p3) + DistanceEx(t2p3, t2p1))

        TriTriSegmentEx = (Least(l1, l2) / Large(l1, l2)) * -32768
    Else
        'the triangles are certianly colliding, and must be caught
        'before two edges have penetrated the other, or vice versa
        'and that before this function is called so by time now is
        
        Dim nA As Point, nB As Point
        nA = VectorCrossProduct(VectorDeduction(t1p2, t1p1), VectorDeduction(t1p3, t1p1))
        nB = VectorCrossProduct(VectorDeduction(t2p2, t2p1), VectorDeduction(t2p3, t2p1))
    
        Dim pts(0 To 5) As Point
        Dim c As Integer: c = 0
        Dim x As Point
    
        ' Intersect edges of A with plane of B
        If EdgePlaneIntersect(t1p1, t1p2, t2p1, nB, x) Then If PointInTriangle(x, t2p1, t2p2, t2p3) Then pts(c) = x: c = c + 1
        If EdgePlaneIntersect(t1p2, t1p3, t2p1, nB, x) Then If PointInTriangle(x, t2p1, t2p2, t2p3) Then pts(c) = x: c = c + 1
        If EdgePlaneIntersect(t1p3, t1p1, t2p1, nB, x) Then If PointInTriangle(x, t2p1, t2p2, t2p3) Then pts(c) = x: c = c + 1
    
        ' Intersect edges of B with plane of A
        If EdgePlaneIntersect(t2p1, t2p2, t1p1, nA, x) Then If PointInTriangle(x, t1p1, t1p2, t1p3) Then pts(c) = x: c = c + 1
        If EdgePlaneIntersect(t2p2, t2p3, t1p1, nA, x) Then If PointInTriangle(x, t1p1, t1p2, t1p3) Then pts(c) = x: c = c + 1
        If EdgePlaneIntersect(t2p3, t2p1, t1p1, nA, x) Then If PointInTriangle(x, t1p1, t1p2, t1p3) Then pts(c) = x: c = c + 1
    
        If c < 2 Then
            'this shouldn't happen by prequisit input args as being in collision determined by three 2D views using PointInPoly
            TriTriSegmentEx = 0
            Exit Function
        End If
    
        ' Choose two extreme points along intersection line direction
        Dim dir As Point: dir = VectorNormalize(VectorCrossProduct(nA, nB))
        Dim minProj As Single, maxProj As Single
        Dim minIdx As Integer, maxIdx As Integer
        minProj = VectorDotProduct(dir, pts(0)): maxProj = minProj
        minIdx = 0: maxIdx = 0
    
        Dim i As Integer
        For i = 1 To c - 1
            Dim p As Single: p = VectorDotProduct(dir, pts(i))
            If p < minProj Then minProj = p: minIdx = i
            If p > maxProj Then maxProj = p: maxIdx = i
        Next i
    
        OutP0 = pts(minIdx)
        OutP1 = pts(maxIdx)
        
        TriTriSegmentEx = DistanceEx(OutP0, OutP1)

    End If
End Function

