Attribute VB_Name = "Module1"

Option Explicit

'In this code:
'poly is short for polygon, a triangle, but sometimes I also short use it
'for polygon, so triangle(s) is the meaning of poly through out this code
'
'sorry
'       Nicholas Forystek
'

'the first three parameters are (X,Y,Z) of the point to be checked against the poly, which consists
'of the final six parameters, length1, length2 and length3 are the triangles edges, and normalX
'normalY and normalZ are the traingles plane normal, it is assumed the trainlge axis is at (0,0,0)
'so you deduct the axis from point before sending point to this function for accurate results.
Public Declare Function PointBehindPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal pointX As Single, ByVal pointY As Single, ByVal pointZ As Single, _
                                    ByVal length1 As Single, ByVal length2 As Single, ByVal length3 As Single, _
                                    ByVal normalX As Single, ByVal normalY As Single, ByVal normalZ As Single) As Boolean

Public Declare Function PointBehindPoly2 Lib "..\Debug\maxland.dll" Alias "PointBehindPoly" _
                                    (ByVal pointX As Single, ByVal pointY As Single, ByVal pointZ As Single, _
                                    ByVal length1 As Single, ByVal length2 As Single, ByVal length3 As Single, _
                                    ByVal normalX As Single, ByVal normalY As Single, ByVal normalZ As Single) As Boolean
                                    
'this next function is a 2D test to see if a point is with in a complex closed shape made of polyDataCount number of
'(X,Y) points, stored in polyDataX and polyDataY, the result is the nearest coordinate where the point crosses over
Public Declare Function PointInPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal pointX As Single, ByVal pointY As Single, _
                                     polyDataX() As Single, polyDataY() As Single, ByVal polyDataCount As Long) As Long
                                    
Public Declare Function PointInPoly2 Lib "..\Debug\maxland.dll" Alias "PointInPoly" _
                                    (ByVal pointX As Single, ByVal pointY As Single, _
                                    ByVal polyDataX As Long, ByVal polyDataY As Long, ByVal polyDataCount As Long) As Long
                                   ' polyDataX As Any, polyDataY As Any, ByVal polyDataCount As Long) As Long

'Test() depends on the functions results above, two views of 3d, x/y and y/z, are called for pointinpoly when
'satisfactionis returned, it is a single point nearest the point checked, then combined with point behindpoly
'and passed to test() the determination is complete by Test() results, as of now PointInPoly2 fails to inform

Public Declare Function Test Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean

Public Declare Function Test2 Lib "..\Debug\maxland.dll" Alias "Test" _
                                    (ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Boolean

                                    
'using all the function above, we then know based on Test() two triangles are certinaly in collision and therefore are
'allowed to be poassed to this next function which is only nessisary if oyu need the data precise segment of traingel
'collision.  I understand that is confusinng because the name of the function matches another popular one that just
'checks accident that I lost my code, otherwise I would edit this externals and tthe file DLL info
Public Declare Function tri_tri_intersect Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal v0_0 As Single, ByVal v0_1 As Single, ByVal v0_2 As Single, _
                                    ByVal v1_0 As Single, ByVal v1_1 As Single, ByVal v1_2 As Single, _
                                    ByVal v2_0 As Single, ByVal v2_1 As Single, ByVal v2_2 As Single, _
                                    ByVal u0_0 As Single, ByVal u0_1 As Single, ByVal u0_2 As Single, _
                                    ByVal u1_0 As Single, ByVal u1_1 As Single, ByVal u1_2 As Single, _
                                    ByVal u2_0 As Single, ByVal u2_1 As Single, ByVal u2_2 As Single) As Integer
                                    'THAT WAS FUN

Public Declare Function tri_tri_intersect2 Lib "..\Backup\MaxLandLib.dll" Alias "tri_tri_intersect" _
                                    (ByVal v0_0 As Single, ByVal v0_1 As Single, ByVal v0_2 As Single, _
                                    ByVal v1_0 As Single, ByVal v1_1 As Single, ByVal v1_2 As Single, _
                                    ByVal v2_0 As Single, ByVal v2_1 As Single, ByVal v2_2 As Single, _
                                    ByVal u0_0 As Single, ByVal u0_1 As Single, ByVal u0_2 As Single, _
                                    ByVal u1_0 As Single, ByVal u1_1 As Single, ByVal u1_2 As Single, _
                                    ByVal u2_0 As Single, ByVal u2_1 As Single, ByVal u2_2 As Single) As Integer
                                    'THAT WAS FUN

'Forystek() 3 variants of culling, vistype painting the canvas was the direction of able to process multiple
'views, at it's default I think is the most powerful (usually a check between two traingles of all) but it
'fails with a % resulting no triangles check depending on the camera, the other two are more secure, and all.
'and unfortunatly I had not got as far as I projected, so it only colors vistype once against all flags
Public Declare Function Forystek Lib "MaxLandLib.dll" _
                                    (ByVal visType As Long, _
                                    ByVal lngFaceCount As Long, _
                                    sngCamera() As Single, _
                                    sngFaceVis() As Single, _
                                    sngVertexX() As Single, _
                                    sngVertexY() As Single, _
                                    sngVertexZ() As Single, _
                                    sngScreenX() As Single, _
                                    sngScreenY() As Single, _
                                    sngScreenZ() As Single, _
                                    sngZBuffer() As Single) As Long
                                    
Public Declare Function Forystek2 Lib "MaxLandLib.dll" Alias "Forystek" _
                                    (ByVal visType As Long, _
                                    ByVal lngFaceCount As Long, _
                                    sngCamera() As Single, _
                                    sngFaceVis() As Single, _
                                    sngVertexX() As Single, _
                                    sngVertexY() As Single, _
                                    sngVertexZ() As Single, _
                                    sngScreenX() As Single, _
                                    sngScreenY() As Single, _
                                    sngScreenZ() As Single, _
                                    sngZBuffer() As Single) As Long

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
                                    
Public Declare Function Collision2 Lib "MaxLandLib.dll" Alias "Collision" _
                                   (ByVal visType As Long, _
                                    ByVal lngFaceCount As Long, _
                                    sngFaceVis() As Single, _
                                    sngVertexX() As Single, _
                                    sngVertexY() As Single, _
                                    sngVertexZ() As Single, _
                                    ByVal lngFaceNum As Long, _
                                    ByRef lngCollidedBrush As Long, _
                                    ByRef lngCollidedFace As Long) As Boolean


'The following variables are needed for Forystek() and Collision() culling and collision
'checking it is quite incompatable to prorietary needs (like doubling the data to use
'the functions vs however one has their data stored already could use)
Public lngTotalTriangles As Long
Public sngTriangleFaceData() As Single
'sngTriangleFaceData dimension (,n) where n=# is triangle/face index
'sngTriangleFaceData dimension (n,) where n=0 is x of the face normal
'sngTriangleFaceData dimension (n,) where n=1 is y of the face normal
'sngTriangleFaceData dimension (n,) where n=2 is z of the face normal
'sngTriangleFaceData dimension (n,) where n=3 is custom vistype flag
'sngTriangleFaceData dimension (n,) where n=4 is the object index
'sngTriangleFaceData dimension (n,) where n=4 is the face index

Public sngVertexXAxisData() As Single
Public sngVertexYAxisData() As Single
Public sngVertexZAxisData() As Single
'sngVertexXAxisData dimension (,n) where n=# is triangle/face index
'sngVertexXAxisData dimension (n,) where n=0 is X of the first vertex
'sngVertexXAxisData dimension (n,) where n=1 is X of the second vertex
'sngVertexXAxisData dimension (n,) where n=2 is X of the fourth vertex
'sngVertexXAxisData dimension (n,) where n=3 is X of the fith an so on

Public Type Point
    X As Single
    Y As Single
    Z As Single
End Type
Public Type Triangle
    p1 As Point
    p2 As Point
    p3 As Point
    a As Point
    n As Point
    l As Point
End Type

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

        .p1 = RandomPoint
        .p2 = RandomPoint
        .p3 = RandomPoint
        
        .a = TriangleAxii(.p1, .p2, .p3)
        
        .l.X = Distance(.p1, .p2)
        .l.Y = Distance(.p2, .p3)
        .l.Z = Distance(.p3, .p1)

        .n = PlaneNormal(.p1, .p2, .p3)
        
    End With

End Function

Public Sub Main()
          
    'to test the new function DLL's outcome is the
    'same as the compiled one whose source is lost.
    
    Dim n1 As Integer
    Dim n2 As Integer
    Dim n3 As Integer
    Dim n4 As Integer
    Dim n5 As Integer
    Dim n6 As Integer
         
    Dim pX1 As Single
    Dim pY1 As Single
    Dim pZ1 As Single
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
    
    Dim pointX As Single
    Dim pointY As Single
    Dim pointZ As Single

    Dim PointListsX(0 To 5) As Single
    Dim PointListsY(0 To 5) As Single
    Dim PointListsZ(0 To 5) As Single
    
    'make an 8x8 square whose center is (0,0)
    'going in the counter-clockwise direction

    Randomize

    Dim t1 As Triangle
    Dim t2 As Triangle
    
    
    PointListsX(0) = 4: PointListsY(0) = -4
    PointListsX(1) = 4: PointListsY(1) = 4
    PointListsX(2) = -4: PointListsY(2) = 4
    PointListsX(3) = -4: PointListsY(3) = -4
    PointListsX(4) = 4: PointListsY(4) = -4
    
    
    
    Do

        Randomize
        DoEvents
        'this main loop is to debug all the frist three declared functions
        'to ensure they produce the same results as the original lost code

        With RandomPoint
            pX1 = .X
            pY1 = .Y
            pZ1 = .Z
        End With
        With RandomTriangle
            nX1 = .n.X
            nY1 = .n.Y
            nZ1 = .n.Z
            vX1 = .a.X
            vY1 = .a.Y
            vZ1 = .a.Z
        End With

        Debug.Print "PointBehindPoly()=" & PointBehindPoly(pX1, pY1, pZ1, nX1, nY1, nZ1, vX1, vY1, vZ1) & _
            " PointBehindPoly2()=" & PointBehindPoly2(pX1, pY1, pZ1, nX1, nY1, nZ1, vX1, vY1, vZ1)
        If Not (PointBehindPoly(pX1, pY1, pZ1, nX1, nY1, nZ1, vX1, vY1, vZ1) = _
            PointBehindPoly2(pX1, pY1, pZ1, nX1, nY1, nZ1, vX1, vY1, vZ1)) Then Stop
        Debug.Print

        'the box is 8x8 centered on (0,0) so we'll use
        'twice it's size and generate random within -8,8
        pointX = (RndNum(0, 16) - 8)
        pointY = (RndNum(0, 16) - 8)

        Debug.Print "PointInPoly()=" & PointInPoly(pointX, pointY, PointListsX, PointListsY, 5) & "  " & _
            "PointInPoly2()=" & PointInPoly2(pointX, pointY, ByVal VarPtr(PointListsX(0)), ByVal VarPtr(PointListsY(0)), 5) & " " & _
            "PointInPoly3()=" & PointInPoly3(pointX, pointY, PointListsX, PointListsY, 5)
        If (Not (PointInPoly(pointX, pointY, PointListsX, PointListsY, 5) = _
            PointInPoly2(pointX, pointY, ByVal VarPtr(PointListsX(0)), ByVal VarPtr(PointListsY(0)), 5))) Or _
            (Not (PointInPoly(pointX, pointY, PointListsX, PointListsY, 5) = _
            PointInPoly3(pointX, pointY, PointListsX, PointListsY, 5))) Then Stop
        Debug.Print

        'arbitrary arguments, unsigned short return values from PointInPoly that results a percentage with in the
        'scope of a integer max value from zero, indicating the point in the point list it falls inside the poly on
        n1 = RndNum(0, 1)
        n2 = RndNum(0, 1)
        n3 = RndNum(0, 1)

        Debug.Print "Test(n1, n2, n3)=" & Test(n1, n2, n3) & " Test2(n1, n2, n3)=" & Test2(n1, n2, n3)
        If Not CVar(Test(n1, n2, n3)) = CVar(Test2(n1, n2, n3)) Then Stop
        Debug.Print

    Loop While True
    
    
End Sub


Public Function Distance(ByRef p1 As Point, ByRef p2 As Point) As Single
    Distance = (((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
    If Distance <> 0 Then Distance = Distance ^ (1 / 2)
End Function

Public Function PlaneNormal(ByRef v0 As Point, ByRef V1 As Point, ByRef V2 As Point) As Point
    'returns a vector perpendicular to a plane V, at 0,0,0, with out the local coordinates information
    PlaneNormal = VectorNormalize(VectorCrossProduct(VectorDeduction(v0, V1), VectorDeduction(V1, V2)))
End Function
Public Function VectorNormalize(ByRef p1 As Point) As Point
    With VectorNormalize
        .Z = (Abs(p1.X) + Abs(p1.Y) + Abs(p1.Z))
        If (Round(.Z, 6) > 0) Then
            .Z = (1 / .Z)
            .X = (p1.X * .Z)
            .Y = (p1.Y * .Z)
            .Z = (p1.Z * .Z)
        End If
    End With
End Function
Public Function VectorDeduction(ByRef p1 As Point, ByRef p2 As Point) As Point
    With VectorDeduction
        .X = (p1.X - p2.X)
        .Y = (p1.Y - p2.Y)
        .Z = (p1.Z - p2.Z)
    End With
End Function

Public Function VectorCrossProduct(ByRef p1 As Point, ByRef p2 As Point) As Point
    With VectorCrossProduct
        .X = ((p1.Y * p2.Z) - (p1.Z * p2.Y))
        .Y = ((p1.Z * p2.X) - (p1.X * p2.Z))
        .Z = ((p1.X * p2.Y) - (p1.Y * p2.X))
    End With
End Function
Public Function TriangleAxii(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    With TriangleAxii
        Dim o As Point
        o = TriangleOffset(p1, p2, p3)
        .X = (Least(p1.X, p2.X, p3.X) + (o.X / 2))
        .Y = (Least(p1.Y, p2.Y, p3.Y) + (o.Y / 2))
        .Z = (Least(p1.Z, p2.Z, p3.Z) + (o.Z / 2))
    End With
End Function
Public Function TriangleOffset(ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point) As Point
    With TriangleOffset
        .X = (Large(p1.X, p2.X, p3.X) - Least(p1.X, p2.X, p3.X))
        .Y = (Large(p1.Y, p2.Y, p3.Y) - Least(p1.Y, p2.Y, p3.Y))
        .Z = (Large(p1.Z, p2.Z, p3.Z) - Least(p1.Z, p2.Z, p3.Z))
    End With
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
'Public Sub RandomTriangle()
'
'    Dim pX As Single
'    Dim pY As Single
'    Dim pZ As Single
'    Dim vX As Single
'    Dim vY As Single
'    Dim vZ As Single
'    Dim nX As Single
'    Dim nY As Single
'    Dim nZ As Single
'
'    'to setup the sides, first
'    'randomize three points
'    With RandomPoint
'        pX = .X
'        pY = .Y
'        pZ = .Z
'    End With
'
'    With RandomPoint
'        nX = .X
'        nY = .Y
'        nZ = .Z
'    End With
'
'    With RandomPoint
'        vX = .X
'        vY = .Y
'        vZ = .Z
'    End With
'
'    'get the distances
'    vX1 = (((pointX - nX) ^ 2) + ((pointY - nY) ^ 2) + ((pointZ - nZ) ^ 2))
'    If vX1 <> 0 Then
'        vX1 = vX1 ^ (1 / 2)
'    Else
'        vX1 = 0
'    End If
'    vY1 = (((nX - pX) ^ 2) + ((nY - pY) ^ 2) + ((nZ - pZ) ^ 2))
'    If vY1 <> 0 Then
'        vY1 = vY1 ^ (1 / 2)
'    Else
'        vY1 = 0
'    End If
'    vZ1 = (((pX - pointX) ^ 2) + ((pY - pointY) ^ 2) + ((pZ - pointZ) ^ 2))
'    If vZ1 <> 0 Then
'        vZ1 = vZ1 ^ (1 / 2)
'    Else
'        vZ1 = 0
'    End If
'
'    'rerandomize a point
'    pX1 = (RndNum(0, 200) - 100)
'    pY1 = (RndNum(0, 200) - 100)
'    pZ1 = (RndNum(0, 200) - 100)
'
'    'makeup a normal of the poly (n)
'    nX1 = ((RndNum(0, 200) / 100) - 1)
'    nY1 = ((RndNum(0, 200) / 100) - 1)
'    nZ1 = ((RndNum(0, 200) / 100) - 1)
'
'End Sub


'Public Function PointInPoly3(ByVal pX As Single, ByVal pY As Single, polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long
'    If polyn > 2 Then
'        Dim ref As Single
'        Dim ret As Single
'        ref = (pX - polyx(0)) * (polyy(1) - polyy(0)) - (pY - polyy(0)) * (polyx(1) - polyx(0))
'        ret = ref
'        Dim i As Long
'        For i = 1 To polyn
'            ref = ((pX - polyx(i)) * (polyy(i) - polyy(i - 1)) - (pY - polyy(i)) * (polyx(i) - polyx(i - 1)))
'            If ((ret > 0) And (ref < 0)) And PointInPoly3 = 0 Then
'                PointInPoly3 = i
'            End If
'            ret = ref
'        Next
'        If ((PointInPoly3 > 0) Or (PointInPoly3 <= polyn)) Then
'            PointInPoly3 = 1
'        End If
''        If PointInPoly3 <> 0 Then
''            PointInPoly3 = -Int(((ret > 0) And (ref > 0)))
''        Else
''            PointInPoly3 = -Int(((ret > 0) Xor (ref < 0)))
''        End If
'    End If
'End Function
Public Function PointInPoly3(ByVal pX As Single, ByVal pY As Single, polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long

    If (polyn > 2) Then
        Dim ref As Single
        Dim ret As Single
        Dim result As Long
        
        ref = ((pX - polyx(0)) * (polyy(1) - polyy(0)) - (pY - polyy(0)) * (polyx(1) - polyx(0)))
        ret = ref
        Dim i As Long
        For i = 1 To polyn
            ref = ((pX - polyx(i)) * (polyy(i) - polyy(i - 1)) - (pY - polyy(i)) * (polyx(i) - polyx(i - 1)))
            If ((ret >= 0) And (ref < 0) And (result = 0)) Then
                PointInPoly3 = i
            End If
            ret = ref
        Next
        If ((result = 0) Or (result > polyn)) Then
            PointInPoly3 = 1 '//todo: this is suppose to return a decimal percent
                                  '                      //of the total polygon points where in is found inside
        Else
            PointInPoly3 = 0
        End If
    End If
   
End Function
