Attribute VB_Name = "Module1"

Option Explicit

Public Declare Function PointBehindPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal pX As Single, ByVal pY As Single, ByVal pZ As Single, _
                                    ByVal nX As Single, ByVal nY As Single, ByVal nZ As Single, _
                                    ByVal vX As Single, ByVal vY As Single, ByVal vZ As Single) As Boolean

Public Declare Function PointBehindPoly2 Lib "..\Backup\MaxLandLib.dll" Alias "PointBehindPoly" _
                                    (ByVal pX As Single, ByVal pY As Single, ByVal pZ As Single, _
                                    ByVal nX As Single, ByVal nY As Single, ByVal nZ As Single, _
                                    ByVal vX As Single, ByVal vY As Single, ByVal vZ As Single) As Boolean

Public Declare Function PointInPoly Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal pX As Single, ByVal pY As Single, _
                                    polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long
    
Public Declare Function PointInPoly2 Lib "..\Debug\maxland.dll" Alias "PointInPoly" _
                                    (ByVal pX As Single, ByVal pY As Single, _
                                    polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long

Public Declare Function Test Lib "..\Backup\MaxLandLib.dll" _
                                    (ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single) As Boolean

Public Declare Function Test2 Lib "..\Debug\maxland.dll" Alias "Test" _
                                    (ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Boolean

                                    
                                    
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



Private Function RndNum(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
    RndNum = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound) - 1
End Function

Public Sub Main()
          

    Dim n1 As Integer
    Dim n2 As Integer
    Dim n3 As Integer
    
    Dim cnt As Long
    For cnt = 1 To 2000
    
        Randomize
        DoEvents
        
        n1 = RndNum(0, 1)
        n2 = RndNum(0, 1)
        n3 = RndNum(0, 1)
        
        Debug.Print Test(n1, n2, n3); Test2(n1, n2, n3)
        
    Next


    
End Sub
