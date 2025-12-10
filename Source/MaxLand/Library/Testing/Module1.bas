Attribute VB_Name = "Module1"

Option Explicit

'the first three parameters are (X,Y,Z) of the point to be checked against a traingle defined by its normal and center
Public Declare Function PointBehindPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                    ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single, _
                                    ByVal CenterX As Single, ByVal CenterY As Single, ByVal CenterZ As Single) As Boolean

Public Declare Function PointTouchesTriangle Lib "..\Debug\maxland.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                    ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single, _
                                    ByVal CenterX As Single, ByVal CenterY As Single, ByVal CenterZ As Single) As Boolean
                                    
'this next function is a 2D test to see if a point is with in a complex closed shape made of polyDataCount number of
'(X,Y) points, stored in polyDataX and polyDataY, the result is the nearest coordinate where the point crosses over
Public Declare Function PointInPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, _
                                     polyDataX() As Single, polyDataY() As Single, ByVal polyDataCount As Long) As Long
                                    
Public Declare Function PointInsidePointList Lib "..\Debug\maxland.dll" _
                                    (ByVal PointX As Single, ByVal PointY As Single, _
                                    ByVal polyListX As Long, ByVal polyListY As Long, ByVal polyListaCount As Long) As Long
                                   ' polyDataX As Any, polyDataY As Any, ByVal polyDataCount As Long) As Long

'Test() depends on the functions results above, two views of 3d, x/y and y/z, are called for pointinpoly when
'satisfactionis returned, it is a single point nearest the point checked, then combined with point behindpoly
'and passed to test() the determination is complete by Test() results, as of now PointInPoly2 fails to inform

Public Declare Function Test Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean

Public Declare Function Test2 Lib "..\Debug\maxland.dll" Alias "Test" _
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

Public Declare Function TriangleCrossSegmentEx Lib "..\Debug\maxland.dll" _
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
'
'Public Declare Function Culling Lib "..\Debug\maxland.dll" _
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
'
''this is the project purpose, the collision checker, I am certian when I did this one
''it was object count * two at the least for checking and then the trainlges do so too
''but the visType is a flag that only traingles with the flag are check for collision
'Public Declare Function Collision Lib "MaxLandLib.dll" _
'                                   (ByVal visType As Long, _
'                                    ByVal lngFaceCount As Long, _
'                                    sngFaceVis() As Single, _
'                                    sngVertexX() As Single, _
'                                    sngVertexY() As Single, _
'                                    sngVertexZ() As Single, _
'                                    ByVal lngFaceNum As Long, _
'                                    ByRef lngCollidedBrush As Long, _
'                                    ByRef lngCollidedFace As Long) As Boolean
'
'Public Declare Function Collision2 Lib "..\Debug\maxland.dll" Alias "Collision" _
'                                   (ByVal visType As Long, _
'                                    ByVal lngFaceCount As Long, _
'                                    sngFaceVis() As Single, _
'                                    sngVertexX() As Single, _
'                                    sngVertexY() As Single, _
'                                    sngVertexZ() As Single, _
'                                    ByVal lngFaceNum As Long, _
'                                    ByRef lngCollidedBrush As Long, _
'                                    ByRef lngCollidedFace As Long) As Boolean
'
'
''The following variables are needed for Forystek() and Collision() culling and collision
''checking it is quite incompatable to prorietary needs (like doubling the data to use
''the functions vs however one has their data stored already could use)
'Public lngTotalTriangles As Long
'Public sngTriangleFaceData() As Single
''sngTriangleFaceData dimension (,n) where n=# is triangle/face index
''sngTriangleFaceData dimension (n,) where n=0 is x of the face normal
''sngTriangleFaceData dimension (n,) where n=1 is y of the face normal
''sngTriangleFaceData dimension (n,) where n=2 is z of the face normal
''sngTriangleFaceData dimension (n,) where n=3 is custom vistype flag
''sngTriangleFaceData dimension (n,) where n=4 is the object index
''sngTriangleFaceData dimension (n,) where n=4 is the face index
'
'Public sngVertexXAxisData() As Single
'Public sngVertexYAxisData() As Single
'Public sngVertexZAxisData() As Single
''sngVertexXAxisData dimension (,n) where n=# is triangle/face index
''sngVertexXAxisData dimension (n,) where n=0 is X of the first vertex
''sngVertexXAxisData dimension (n,) where n=1 is X of the second vertex
''sngVertexXAxisData dimension (n,) where n=2 is X of the fourth vertex
''sngVertexXAxisData dimension (n,) where n=3 is X of the fith an so on

Public Type Point
    X As Single
    Y As Single
    Z As Single
End Type
Public Type Triangle
    P1 As Point
    P2 As Point
    p3 As Point
    A As Point
    n As Point
    L As Point
End Type

Public Sub Main()
    Main1
    Main2
End Sub

Private Sub ConvertTriangle2(ByRef t1 As Triangle, ByRef c1 As Point, ByRef n1 As Point, ByRef l1 As Point)
    c1.X = (t1.P1.X + t1.P2.X + t1.p3.X) / 3
    c1.Y = (t1.P1.Y + t1.P2.Y + t1.p3.Y) / 3
    c1.Z = (t1.P1.Z + t1.P2.Z + t1.p3.Z) / 3
    n1 = VectorNormalize(TriangleNormal(t1.P1, t1.P2, t1.p3))
    l1.X = Distance(t1.P1, t1.P2)
    l1.Y = Distance(t1.P2, t1.p3)
    l1.Z = Distance(t1.p3, t1.P1)
End Sub
Public Sub Main2()

    Dim t1 As Triangle
    Dim t2 As Triangle
    Dim P1 As Point
    Dim P2 As Point
    Dim Ret As Double
    Dim c1 As Point
    Dim c2 As Point
    Dim n1 As Point
    Dim n2 As Point
    Dim l1 As Point
    Dim l2 As Point
    
    t1.p3 = MakePoint(0, 8, 0)
    t1.P2 = MakePoint(15, 7, 6)
    t1.P1 = MakePoint(4, 0, 14)
    t2.p3 = MakePoint(14, 0, -3)
    t2.P2 = MakePoint(6, 12, 1)
    t2.P1 = MakePoint(4, 12, 14)
        
    Ret = TriangleCrossSegmentEx(t1.P1.X, t1.P1.Y, t1.P1.Z, t1.P2.X, t1.P2.Y, t1.P2.Z, t1.p3.X, t1.p3.Y, t1.p3.Z, _
                        t2.P1.X, t2.P1.Y, t2.P1.Z, t2.P2.X, t2.P2.Y, t2.P2.Z, t2.p3.X, t2.p3.Y, t2.p3.Z, _
                        P1.X, P1.Y, P1.Z, P2.X, P2.Y, P2.Z)
    If Ret Then
        Debug.Print "Intersection segment: " & Ret
        Debug.Print "H=(8,7,3)=(" & Round(P1.X, 0) & "," & Round(P1.Y, 0) & "," & Round(P1.Z, 0) & ")"
        Debug.Print "I=(9,6,6)=(" & Round(P2.X, 0) & "," & Round(P2.Y, 0) & "," & Round(P2.Z, 0) & ")"
    Else
        Debug.Print "No intersection."
    End If
    

    t1.P1 = MakePoint(0, 0, 0)
    t1.P2 = MakePoint(20, 0, 0)
    t1.p3 = MakePoint(0, 20, 0)
    t2.P1 = MakePoint(-10, 5, 0)
    t2.P2 = MakePoint(10, 5, 10)
    t2.p3 = MakePoint(10, 5, -10)
    
    Ret = TriangleCrossSegmentEx(t1.P1.X, t1.P1.Y, t1.P1.Z, t1.P2.X, t1.P2.Y, t1.P2.Z, t1.p3.X, t1.p3.Y, t1.p3.Z, _
                        t2.P1.X, t2.P1.Y, t2.P1.Z, t2.P2.X, t2.P2.Y, t2.P2.Z, t2.p3.X, t2.p3.Y, t2.p3.Z, _
                        P1.X, P1.Y, P1.Z, P2.X, P2.Y, P2.Z)
    If Ret Then
        Debug.Print "Intersection segment: " & Ret
        Debug.Print "H=(10,5,0)=(" & Round(P1.X, 0) & "," & Round(P1.Y, 0) & "," & Round(P1.Z, 0) & ")"
        Debug.Print "I=(0,5,0)=(" & Round(P2.X, 0) & "," & Round(P2.Y, 0) & "," & Round(P2.Z, 0) & ")"
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

    Dim PointListsX(0 To 5) As Single
    Dim PointListsY(0 To 5) As Single
    Dim PointListsZ(0 To 5) As Single
    
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
            Px1 = .X
            Py1 = .Y
            Pz1 = .Z
        End With
        With RandomTriangle
            nX1 = .n.X
            nY1 = .n.Y
            nZ1 = .n.Z
            vX1 = .A.X
            vY1 = .A.Y
            vZ1 = .A.Z
        End With

        Debug.Print "PointBehindPoly()=" & PointBehindPoly(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1) & _
            " PointTouchesTriangle()=" & PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1) & _
            " PointBehindPoly3()=" & PointBehindPoly3(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)
        If (Not (CVar(PointBehindPoly(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)) = _
            CVar(PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)))) Or _
            (Not (CVar(PointTouchesTriangle(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)) = _
            CVar(PointBehindPoly3(Px1, Py1, Pz1, nX1, nY1, nZ1, vX1, vY1, vZ1)))) Then testCount = -Abs(testCount)
        Debug.Print

        'the box is 8x8 centered on (0,0) so we'll use
        'twice it's size and generate random within -8,8
        PointX = (RndNum(0, 16) - 8)
        PointY = (RndNum(0, 16) - 8)
        PointZ = (RndNum(0, 16) - 8)

        
        Debug.Print "PointInPoly()=" & PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) & "  " & _
            "PointInsidePointList()=" & PointInsidePointList(PointX, PointY, ByVal VarPtr(PointListsX(0)), ByVal VarPtr(PointListsY(0)), 5) & " " & _
            "PointInPoly3()=" & PointInPoly3(PointX, PointY, PointListsX, PointListsY, 5)
        If (Not (PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) = _
            PointInsidePointList(PointX, PointY, ByVal VarPtr(PointListsX(0)), ByVal VarPtr(PointListsY(0)), 5))) Or _
             (Not (PointInPoly(PointX, PointY, PointListsX, PointListsY, 5) = PointInPoly3(PointX, PointY, PointListsX, PointListsY, 5))) Then testCount = -Abs(testCount)
        Debug.Print
        
        'arbitrary arguments, unsigned short return values from PointInPoly that results a percentage with in the
        'scope of a integer max value from zero, indicating the point in the point list it falls inside the poly on
        n1 = Round(RndNum(0, 1), 0)
        n2 = Round(RndNum(0, 1), 0)
        n3 = Round(RndNum(0, 1), 0)


        'use the same square as if it is a cube in 3d,
        'and check each 2D axis for collision using test

        
        Debug.Print "Test(n1, n2, n3)=" & Test(n1, n2, n3) & " Test2(n1, n2, n3)=" & Test2(n1, n2, n3)
        If Not CVar(Test(n1, n2, n3)) = CVar(Test2(n1, n2, n3)) Then testCount = -Abs(testCount)
        Debug.Print


'        Debug.Print "Test(n1, n2, n3)=" & Test(n1, n2, n3) & " Test2(n1, n2, n3)=" & Test2(n1, n2, n3) & " Test3(n1, n2, n3)=" & Test3(n1, n2, n3)
'        If (Not (CVar(Test(n1, n2, n3)) = CVar(Test2(n1, n2, n3)))) Or (Not (CVar(Test2(n1, n2, n3)) = CVar(Test3(n1, n2, n3)))) Then Stop
'        Debug.Print

    Loop Until testCount > 1000 Or testCount < 0
    
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
        .X = (RndNum(0, 16) - 8)
        .Y = (RndNum(0, 16) - 8)
        .Z = (RndNum(0, 16) - 8)
    End With
End Function

Public Function RandomTriangle() As Triangle

    With RandomTriangle

        .P1 = RandomPoint
        .P2 = RandomPoint
        .p3 = RandomPoint
        
        .A = TriangleAxii(.P1, .P2, .p3)
        
        .L.X = Distance(.P1, .P2)
        .L.Y = Distance(.P2, .p3)
        .L.Z = Distance(.p3, .P1)

        .n = PlaneNormal(.P1, .P2, .p3)
        
    End With

End Function


Public Function Test3(ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean
'I have been unsuccessful in the VB6 environment to get this one to act like Test() and Test2()
    Test3 = ((((n1 And n2 + n3) Or (n1 + n2 And n3)) And ((n1 - n2 Or Not n3) - (Not n1 Or n2 - n3))) _
        Or (((n1 - n2 Or n3) And (n1 - n2 Or n3)) + ((n1 Or n2 + Not n3) And (Not n1 + n2 And n3))))
End Function

Public Function PointBehindPoly3(ByVal PointX As Single, ByVal PointY As Single, ByVal PointZ As Single, _
                                ByVal Length1 As Single, ByVal Length2 As Single, ByVal Length3 As Single, _
                                ByVal NormalX As Single, ByVal NormalY As Single, ByVal NormalZ As Single) As Boolean

    PointBehindPoly3 = ((PointZ * Length3 + Length2 * PointY + Length1 * PointX) - (Length3 * NormalZ + Length1 * NormalX + Length2 * NormalY) <= 0)
End Function

Public Function PointInPoly3(ByVal pX As Single, ByVal pY As Single, polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long

    If (polyn > 2) Then
        Dim ref As Single
        Dim Ret As Single
        Dim result As Long

        ref = ((pX - polyx(0)) * (polyy(1) - polyy(0)) - (pY - polyy(0)) * (polyx(1) - polyx(0)))
        Ret = ref
        Dim i As Long
        For i = 1 To polyn
            ref = ((pX - polyx(i)) * (polyy(i) - polyy(i - 1)) - (pY - polyy(i)) * (polyx(i) - polyx(i - 1)))
            If ((Ret >= 0) And (ref < 0) And (result = 0)) Then
                result = i
            End If
            Ret = ref
        Next
        If ((result = 0) Or (result > polyn)) Then
            PointInPoly3 = 1 '//todo: this is suppose to return a decimal percent
                                  '                      //of the total polygon points where in is found inside
        Else
            PointInPoly3 = 0
        End If
    End If

End Function




Public Function Distance(ByRef P1 As Point, ByRef P2 As Point) As Single
    Distance = (((P1.X - P2.X) ^ 2) + ((P1.Y - P2.Y) ^ 2) + ((P1.Z - P2.Z) ^ 2))
    If Distance <> 0 Then Distance = Distance ^ (1 / 2)
End Function

Public Function PlaneNormal(ByRef V0 As Point, ByRef v1 As Point, ByRef v2 As Point) As Point
    'returns a vector perpendicular to a plane V, at 0,0,0, with out the local coordinates information
    PlaneNormal = VectorCrossProduct(VectorDeduction(V0, v1), VectorDeduction(v1, v2))
End Function
Public Function MakePoint(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point
    With MakePoint
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
Private Function VectorNormalize(A As Point) As Point
    Dim L As Double: L = DistanceEx(MakePoint(0, 0, 0), A)
    If L = 0 Then
        VectorNormalize = MakePoint(0, 0, 0)
    Else
        VectorNormalize = MakePoint(A.X / L, A.Y / L, A.Z / L)
    End If
End Function

Public Function VectorDeduction(ByRef P1 As Point, ByRef P2 As Point) As Point
    With VectorDeduction
        .X = (P1.X - P2.X)
        .Y = (P1.Y - P2.Y)
        .Z = (P1.Z - P2.Z)
    End With
End Function

Public Function VectorCrossProduct(ByRef P1 As Point, ByRef P2 As Point) As Point
    With VectorCrossProduct
        .X = ((P1.Y * P2.Z) - (P1.Z * P2.Y))
        .Y = ((P1.Z * P2.X) - (P1.X * P2.Z))
        .Z = ((P1.X * P2.Y) - (P1.Y * P2.X))
    End With
End Function
Public Function VectorDotProduct(A As Point, B As Point) As Double
    VectorDotProduct = A.X * B.X + A.Y * B.Y + A.Z * B.Z
End Function
Public Function VectorAddition(ByRef P1 As Point, ByRef P2 As Point) As Point
    With VectorAddition
        .X = (P1.X + P2.X)
        .Y = (P1.Y + P2.Y)
        .Z = (P1.Z + P2.Z)
    End With
End Function
Private Function TriangleNormal(ByRef P1 As Point, ByRef P2 As Point, ByRef p3 As Point) As Point
    Dim v1 As Point, v2 As Point
    v1 = VectorDeduction(P1, P2)
    v2 = VectorDeduction(P1, p3)
    TriangleNormal = VectorCrossProduct(v1, v2)
End Function
Public Function TriangleAxii(ByRef P1 As Point, ByRef P2 As Point, ByRef p3 As Point) As Point
    With TriangleAxii
        Dim o As Point
        o = TriangleOffset(P1, P2, p3)
        .X = (Least(P1.X, P2.X, p3.X) + (o.X / 2))
        .Y = (Least(P1.Y, P2.Y, p3.Y) + (o.Y / 2))
        .Z = (Least(P1.Z, P2.Z, p3.Z) + (o.Z / 2))
    End With
End Function
Public Function TriangleOffset(ByRef P1 As Point, ByRef P2 As Point, ByRef p3 As Point) As Point
    With TriangleOffset
        .X = (Large(P1.X, P2.X, p3.X) - Least(P1.X, P2.X, p3.X))
        .Y = (Large(P1.Y, P2.Y, p3.Y) - Least(P1.Y, P2.Y, p3.Y))
        .Z = (Large(P1.Z, P2.Z, p3.Z) - Least(P1.Z, P2.Z, p3.Z))
    End With
End Function
Public Function DistanceEx(ByRef P1 As Point, ByRef P2 As Point) As Double
    DistanceEx = (((P1.X - P2.X) ^ 2) + ((P1.Y - P2.Y) ^ 2) + ((P1.Z - P2.Z) ^ 2))
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
    AreParallel = (Abs(cross.X) < 0.0001 And Abs(cross.Y) < 0.0001 And Abs(cross.Z) < 0.0001)
End Function

Public Function AreCoplanar(t1p1 As Point, t1p2 As Point, t1p3 As Point, t2p1 As Point, t2p2 As Point, t2p3 As Point) As Boolean
    If Not AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3) Then
        AreCoplanar = False
        Exit Function
    End If
    
    Dim n1 As Point, d As Double
    n1 = TriangleNormal(t1p1, t1p2, t1p3)
    d = -(n1.X * t1p1.X + n1.Y * t1p1.Y + n1.Z * t1p1.Z)
    
    AreCoplanar = Abs(n1.X * t2p1.X + n1.Y * t2p1.Y + n1.Z * t2p1.Z + d) < 0.0001
End Function

Public Function AreParallelCoplanar(t1p1 As Point, t1p2 As Point, t1p3 As Point, t2p1 As Point, t2p2 As Point, t2p3 As Point) As Boolean
    Dim n1 As Point, n2 As Point, cross As Point
    Dim d As Double, p As Point
    
    ' Normals
    n1 = TriangleNormal(t1)
    n2 = TriangleNormal(t2)
    
    ' Cross product of normals
    cross = VectorCrossProduct(n1, n2)
    
    ' Plane constant from triangle 1
    d = -(n1.X * t1p1.X + n1.Y * t1p1.Y + n1.Z * t1p1.Z)
    
    ' Test point from triangle 2
    p = t2p1
    
    ' Single algebraic condition: parallel AND coplanar
    AreParallelCoplanar = _
        (Abs(cross.X) < 0.0001 And Abs(cross.Y) < 0.0001 And Abs(cross.Z) < 0.0001) _
        And (Abs(n1.X * p.X + n1.Y * p.Y + n1.Z * p.Z + d) < 0.0001)
End Function



' ===== Point-in-triangle test (barycentric) =====
Private Function PointInTriangle(p As Point, V0 As Point, v1 As Point, v2 As Point) As Boolean
    Dim u As Point, v As Point, w As Point
    u = VectorDeduction(v1, V0)
    v = VectorDeduction(v2, V0)
    w = VectorDeduction(p, V0)

    Dim uu As Double, vv As Double, uv As Double
    Dim wu As Double, wv As Double, d As Double

    uu = VectorDotProduct(u, u)
    vv = VectorDotProduct(v, v)
    uv = VectorDotProduct(u, v)
    wu = VectorDotProduct(w, u)
    wv = VectorDotProduct(w, v)

    d = uv * uv - uu * vv
    If Abs(d) < 0.000000001 Then
        PointInTriangle = False
        Exit Function
    End If

    Dim s As Double, t As Double
    s = (uv * wv - vv * wu) / d
    t = (uv * wu - uu * wv) / d

    PointInTriangle = (s >= -0.000000001 And t >= -0.000000001 And (s + t) <= 1 + 0.000000001)
End Function

' ===== Edge-plane intersection =====
Private Function EdgePlaneIntersect(p As Point, Q As Point, planePoint As Point, PlaneNormal As Point, X As Point) As Boolean
    Dim dir As Point: dir = VectorDeduction(Q, p)
    Dim denom As Double: denom = VectorDotProduct(PlaneNormal, dir)
    If Abs(denom) < 0.000000001 Then
        EdgePlaneIntersect = False
        Exit Function
    End If

    Dim t As Double
    t = VectorDotProduct(PlaneNormal, VectorDeduction(planePoint, p)) / denom
    If t < -0.000000001 Or t > 1 + 0.000000001 Then
        EdgePlaneIntersect = False
        Exit Function
    End If

    X = VectorAddition(p, MakePoint(dir.X * t, dir.Y * t, dir.Z * t))
    EdgePlaneIntersect = True
End Function

'##########################################################################
'##########################################################################
'##########################################################################


' ===== Main intersection routine =====
Public Function TriTriSegmentEx(ByRef t1p1 As Point, ByRef t1p2 As Point, ByRef t1p3 As Point, ByRef t2p1 As Point, ByRef t2p2 As Point, ByRef t2p3 As Point, ByRef OutP0 As Point, ByRef OutP1 As Point) As Double
    Dim ap As Boolean
    Dim ac As Boolean
    ap = AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3)
    ac = AreCoplanar(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3)
    Dim l1 As Double
    Dim l2 As Double

        
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
        Dim C As Integer: C = 0
        Dim X As Point
    
        ' Intersect edges of A with plane of B
        If EdgePlaneIntersect(t1p1, t1p2, t2p1, nB, X) Then If PointInTriangle(X, t2p1, t2p2, t2p3) Then pts(C) = X: C = C + 1
        If EdgePlaneIntersect(t1p2, t1p3, t2p1, nB, X) Then If PointInTriangle(X, t2p1, t2p2, t2p3) Then pts(C) = X: C = C + 1
        If EdgePlaneIntersect(t1p3, t1p1, t2p1, nB, X) Then If PointInTriangle(X, t2p1, t2p2, t2p3) Then pts(C) = X: C = C + 1
    
        ' Intersect edges of B with plane of A
        If EdgePlaneIntersect(t2p1, t2p2, t1p1, nA, X) Then If PointInTriangle(X, t1p1, t1p2, t1p3) Then pts(C) = X: C = C + 1
        If EdgePlaneIntersect(t2p2, t2p3, t1p1, nA, X) Then If PointInTriangle(X, t1p1, t1p2, t1p3) Then pts(C) = X: C = C + 1
        If EdgePlaneIntersect(t2p3, t2p1, t1p1, nA, X) Then If PointInTriangle(X, t1p1, t1p2, t1p3) Then pts(C) = X: C = C + 1
    
        If C < 2 Then
            'this shouldn't happen by prequisit input args as being in collision determined by three 2D views using PointInPoly
            TriTriSegmentEx = 0
            Exit Function
        End If
    
        ' Choose two extreme points along intersection line direction
        Dim dir As Point: dir = VectorNormalize(VectorCrossProduct(nA, nB))
        Dim minProj As Double, maxProj As Double
        Dim minIdx As Integer, maxIdx As Integer
        minProj = VectorDotProduct(dir, pts(0)): maxProj = minProj
        minIdx = 0: maxIdx = 0
    
        Dim i As Integer
        For i = 1 To C - 1
            Dim p As Double: p = VectorDotProduct(dir, pts(i))
            If p < minProj Then minProj = p: minIdx = i
            If p > maxProj Then maxProj = p: maxIdx = i
        Next i
    
        OutP0 = pts(minIdx)
        OutP1 = pts(maxIdx)
        
        l1 = (DistanceEx(t1p1, t1p2) + DistanceEx(t1p2, t1p3) + DistanceEx(t1p3, t1p1))
        l2 = (DistanceEx(t2p1, t2p2) + DistanceEx(t2p2, t2p3) + DistanceEx(t2p3, t2p1))
           
        TriTriSegmentEx = ((DistanceEx(OutP0, OutP1) / (l1 + l2)) * 32767)
    End If
End Function

