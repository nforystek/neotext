Attribute VB_Name = "mod3DObj"
Option Explicit

'############################################################################################################
'Derived Exports ############################################################################################
'############################################################################################################
                                    
Private Declare Function Collision Lib "MaxLandLib.dll" _
                                    (ByVal lngStreamFlagValue As _
                                    Long, ByVal lngTotalTriangles As _
                                    Long, sngTriangleFaceData() As _
                                    Single, sngVertexXAxisData() As _
                                    Single, sngVertexYAxisData() As _
                                    Single, sngVertexZAxisData() As _
                                    Single, ByVal lngTriangleToCheck As _
                                    Long, ByRef lngReturnHitObject As _
                                    Long, ByRef lngReturnHitTriangle As Long) As Boolean

'############################################################################################################
'Variable Declare ###########################################################################################
'############################################################################################################

Public ObjectCount As Long
Public TriangleCount As Long
Public TriangleFace() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'TriangleFace dimension (n,t) where n=0 is x of the face normal
'TriangleFace dimension (n,t) where n=1 is y of the face normal
'TriangleFace dimension (n,t) where n=2 is z of the face normal
'TriangleFace dimension (n,t) where n=3 flag for 1st arg of collision(), culling
'TriangleFace dimension (n,t) where n=4 is the object index
'TriangleFace dimension (n,t) where n=5 is the face index

Public VertexXAxis() As Single
Public VertexYAxis() As Single
Public VertexZAxis() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'VertexXAxis dimension (n,t) where n=0 is X of the first vertex
'VertexXAxis dimension (n,t) where n=1 is X of the second vertex
'VertexXAxis dimension (n,t) where n=2 is X of the third vertex
'VertexYAxis dimension (n,t) where n=0 is Y of the first vertex
'VertexYAxis dimension (n,t) where n=1 is Y of the second vertex
'VertexYAxis dimension (n,t) where n=2 is Y of the third vertex
'VertexZAxis dimension (n,t) where n=0 is Z of the first vertex
'VertexZAxis dimension (n,t) where n=1 is Z of the second vertex
'VertexZAxis dimension (n,t) where n=2 is Z of the third vertex

Public VertexDirectX() As MyVertex
Public ScreenDirectX() As MyScreen

Public Points As Points
'Public Rotates As Orbit 'waiting to be applied rotates
'Public Scalars As Orbit 'waiting to be applied scalars
Public Zero As New Point
Public PlayerGyro As New Point
Public PlanetGyro As New Point
Public Localized As Point

'Public Orbits As Orbits 'collection of in non script accessable for the app for all orbits and implmenets of
'Public Ranges As Ranges 'collection of those that are part of the planet object only needed in global cycle
'Public Points As Points 'cache of all points uniquely, so far this just grows and shouldn't accept change them

Public Sub CleanUpObjs()

    Set Localized = Nothing
    Set Points = Nothing
    
    ObjectCount = 0
    TriangleCount = 0
    Erase TriangleFace
    
    Erase VertexXAxis
    Erase VertexYAxis
    Erase VertexZAxis

    Erase VertexDirectX


End Sub

Public Sub CreateObjs()

    Set Points = New Points
    Set Localized = New Point
         
End Sub



' Return the dot product AB · BC.
' Note that AB · BC = |AB| * |BC| * Cos(theta).
Private Function DotProduct( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal cx As Single, ByVal cy As Single _
  ) As Single
    Dim BAx As Single
    Dim BAy As Single
    Dim BCx As Single
    Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BCx = cx - Bx
    BCy = cy - By

    ' Calculate the dot product.
    DotProduct = BAx * BCx + BAy * BCy
End Function


Public Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, _
    ByVal Bx As Single, ByVal By As Single, ByVal Bz As Single, _
    ByVal cx As Single, ByVal cy As Single, ByVal cz As Single _
  ) As Single
    Dim BAx As Single
    Dim BAy As Single
    Dim BAz As Single
    Dim BCx As Single
    Dim BCy As Single
    Dim BCz As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BAz = Az - Bz
    BCx = cx - Bx
    BCy = cy - By
    BCz = cz - Bz
    
    ' Calculate the Z coordinate of the cross product.
    CrossProductLength = BAx * BCy - BAy * BCz - BAz * BCx
End Function

'' Return the angle ABC.
'' Return a value between PI and -PI.
'' Note that the value is the opposite of what you might
'' expect because Y coordinates increase downward.
Public Function GetAngle(ByRef p1 As Point, ByRef p2 As Point) As Single
'ByVal Ax As Single, ByVal Ay As _
    'Single, ByVal Bx As Single, ByVal By As Single, ByVal _
    'Cx As Single, ByVal Cy As Single) As Single
    Dim dot_product As Single
    Dim cross_product As Single

    ' Get the dot product and cross product.
    dot_product = VectorDotProduct(p1, p2) 'dotproduct(p1.x, p1.y, 0, 0, p3.x, p3.y)
    cross_product = DistanceEx(MakePoint(0, 0, 0), VectorCrossProduct(p1, p2)) 'CrossProductLength(p1.x, p1.y, p1.z, 0, 0, 0, p2.x, p2.y, p2.z) 'CrossProductLength(p1.x, p1.y, 0, 0, p3.x, p3.y)

    ' Calculate the angle.
    GetAngle = ATan2(CDbl(cross_product), CDbl(dot_product)) * DEGREE
End Function


Public Function GetAngle2(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Dim XDiff As Double
    Dim YDiff As Double
    Dim TempAngle As Double

    YDiff = Abs(y2 - y1)

    If x1 = x2 And y1 = y2 Then Exit Function

    If YDiff = 0 And x1 < x2 Then
        GetAngle2 = 0
        Exit Function
    ElseIf YDiff = 0 And x1 > x2 Then
        GetAngle2 = 3.14159265358979
        Exit Function
    End If

    XDiff = Abs(x2 - x1)

    TempAngle = Atn(XDiff / YDiff)

    If y2 > y1 Then TempAngle = 3.14159265358979 - TempAngle
    If x2 < x1 Then TempAngle = -TempAngle
    TempAngle = 1.5707963267949 - TempAngle
    If TempAngle < 0 Then TempAngle = 6.28318530717959 + TempAngle

    GetAngle2 = TempAngle
End Function


Public Function GetAngle3(ByRef p1 As Point, ByRef p2 As Point) As Single
If p1.X = p2.X Then
    If p1.Y < p2.Y Then
        GetAngle3 = 90
    Else
        GetAngle3 = 270
    End If
    Exit Function
ElseIf p1.Y = p2.Y Then
    If p1.X < p2.X Then
        GetAngle3 = 0
    Else
        GetAngle3 = 180
    End If
    Exit Function
Else
    GetAngle3 = Atn(VectorSlope(p1, p2))
    GetAngle3 = GetAngle3 * 180 / PI
    If GetAngle3 < 0 Then GetAngle3 = GetAngle3 + 360
    '----------Test for direction--------
    If p1.X > p2.X And GetAngle3 <> 180 Then GetAngle3 = GetAngle3 + 180
    If p1.Y > p2.Y And GetAngle3 = 90 Then GetAngle3 = GetAngle3 + 180
    If GetAngle3 > 360 Then GetAngle3 = GetAngle3 - 360
End If
End Function

Public Function Sine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180
    p_dblVal = Val(p_dblVal * dblRadian)
    Sine = Sin(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Sine = 0
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function Cosine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180
    p_dblVal = Val(p_dblVal * dblRadian)
    Cosine = Cos(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Cosine = 0
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function Tangent(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180

    p_dblVal = Val(p_dblVal * dblRadian)
    Tangent = Tan(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Tangent = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcSine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblSqr As Single
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
    ' xx Prevent division by Zero error

    If dblSqr = 0 Then
        dblSqr = 1E-30
    End If

    ArcSine = Atn(p_dblVal / dblSqr) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcSine = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcCosine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblSqr As Single
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
    ' xx Prevent division by Zero error

    If dblSqr = 0 Then
        dblSqr = 1E-30
    End If

    ArcCosine = (Atn(-p_dblVal / dblSqr) + 2 * Atn(1)) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcCosine = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcTangent(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    ArcTangent = Atn(p_dblVal) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcTangent = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function



Private Sub ApplyOrigin(ByRef Origin As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Origin Is Nothing Then Exit Sub

'    Static vin As D3DVECTOR
'    Static vout As D3DVECTOR
'    Static matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixTranslation matMesh, -Origin.X, -Origin.Y, -Origin.z
            
    Set ApplyTo.Origin = Origin
    Set ApplyTo.Absolute.Origin = ApplyTo.Origin

'    If TypeName(ApplyTo) <> "Planet" Then
'        Dim V As Matter
'        For Each V In ApplyTo.Volume
'
'            vin.X = V.Point1.X
'            vin.Y = V.Point1.Y
'            vin.z = V.Point1.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point1.X = vout.X
'            V.Point1.Y = vout.Y
'            V.Point1.z = vout.z
'
'            vin.X = V.Point2.X
'            vin.Y = V.Point2.Y
'            vin.z = V.Point2.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point2.X = vout.X
'            V.Point2.Y = vout.Y
'            V.Point2.z = vout.z
'
'            vin.X = V.Point3.X
'            vin.Y = V.Point3.Y
'            vin.z = V.Point3.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point3.X = vout.X
'            V.Point3.Y = vout.Y
'            V.Point3.z = vout.z
'
'            'Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'
'        Next
'    End If
    

'    stacked = stacked + 1
'
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'        Next
'    End If
'
'    stacked = stacked - 1
End Sub


Private Sub ApplyRotate(ByRef Degrees As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Degrees Is Nothing Then Exit Sub

'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Static matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixRotationYawPitchRoll matMesh, -Degrees.X, -Degrees.Y, -Degrees.z
    
    Set ApplyTo.Rotate = Degrees
    Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
    
'    If TypeName(ApplyTo) <> "Planet" Then
'        Dim V As Matter
'        For Each V In ApplyTo.Volume
'
'            vin.X = V.Point1.X
'            vin.Y = V.Point1.Y
'            vin.z = V.Point1.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point1.X = vout.X
'            V.Point1.Y = vout.Y
'            V.Point1.z = vout.z
'
'            vin.X = V.Point2.X
'            vin.Y = V.Point2.Y
'            vin.z = V.Point2.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point2.X = vout.X
'            V.Point2.Y = vout.Y
'            V.Point2.z = vout.z
'
'            vin.X = V.Point3.X
'            vin.Y = V.Point3.Y
'            vin.z = V.Point3.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point3.X = vout.X
'            V.Point3.Y = vout.Y
'            V.Point3.z = vout.z
'
'            'Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'
'        Next
'    End If


    
'    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyRotate Degrees, m, ApplyTo
'    Next
'
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyRotate Degrees, m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyRotate Degrees, m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyRotate Degrees, m, ApplyTo
'        Next
'    End If
'    stacked = stacked - 1

End Sub

Private Static Sub ApplyScaled(ByRef Scalar As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Scalar Is Nothing Then Exit Sub

'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Dim matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixScaling matMesh, Scalar.X, Scalar.Y, Scalar.z

    Set ApplyTo.Scaled = Scalar
    Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled

'    Dim V As Matter
'    For Each V In ApplyTo.Volume
'
'        vin.X = V.Point1.X
'        vin.Y = V.Point1.Y
'        vin.z = V.Point1.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point1.X = vout.X
'        V.Point1.Y = vout.Y
'        V.Point1.z = vout.z
'
'        vin.X = V.Point2.X
'        vin.Y = V.Point2.Y
'        vin.z = V.Point2.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point2.X = vout.X
'        V.Point2.Y = vout.Y
'        V.Point2.z = vout.z
'
'        vin.X = V.Point3.X
'        vin.Y = V.Point3.Y
'        vin.z = V.Point3.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point3.X = vout.X
'        V.Point3.Y = vout.Y
'        V.Point3.z = vout.z
'
'        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'    Next

'    stacked = stacked + 1
'
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyScaled VectorAddition(m.Scaled, Scalar), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'        Next
'    End If
'
'    stacked = stacked - 1
End Sub

Private Static Sub ApplyOffset(ByRef Offset As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Offset Is Nothing Then Exit Sub

'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Dim matMesh As D3DMATRIX
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixTranslation matMesh, -ApplyTo.Offset.X + Offset.X, -ApplyTo.Offset.Y + Offset.Y, -ApplyTo.Offset.z + Offset.z

    Set ApplyTo.Offset = Offset
    Set ApplyTo.Absolute.Offset = ApplyTo.Offset

'    Dim V As Matter
'    For Each V In ApplyTo.Volume
'
'        vin.X = V.Point1.X
'        vin.Y = V.Point1.Y
'        vin.z = V.Point1.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point1.X = vout.X
'        V.Point1.Y = vout.Y
'        V.Point1.z = vout.z
'
'        vin.X = V.Point2.X
'        vin.Y = V.Point2.Y
'        vin.z = V.Point2.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point2.X = vout.X
'        V.Point2.Y = vout.Y
'        V.Point2.z = vout.z
'
'        vin.X = V.Point3.X
'        vin.Y = V.Point3.Y
'        vin.z = V.Point3.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point3.X = vout.X
'        V.Point3.Y = vout.Y
'        V.Point3.z = vout.z
'
'        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'    Next

'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyOrigin VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'        Next
'    End If

End Sub

Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
End Sub
Private Sub LocPos(ByRef Origin As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Origin.X <> 0 Or Origin.Y <> 0 Or Origin.z <> 0 Then
        Dim o As Orbit
        Dim m As Molecule
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        LocPos Origin, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            LocPos Origin, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                CommitOrigin ApplyTo, Parent
                If Relative Then
'                    If Not Parent Is Nothing Then
'                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Parent.Rotate), ApplyTo.Rotate))
'                    ElseIf Not Camera.Planet Is Nothing Then
'                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Camera.Planet.Rotate), ApplyTo.Rotate))
'                    Else
'                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, ApplyTo.Rotate)
'                    End If

                    Set ApplyTo.Relative.Origin = Origin
                Else
                    Set ApplyTo.Absolute.Origin = Origin
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    LocPos Origin, Relative, m, ApplyTo
'                Next
        End Select
    End If
End Sub
Public Sub Rotation(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Orientate(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
End Sub
Private Sub RotOri(ByRef Degrees As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Degrees.X <> 0 Or Degrees.Y <> 0 Or Degrees.z <> 0 Then
        Dim m As Molecule
        Dim o As Point
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        RotOri Degrees, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            RotOri Degrees, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                CommitRotate ApplyTo, Parent
                If Relative Then
'                    If Not Parent Is Nothing Then
'                        Set ApplyTo.Relative.Rotate = AngleAxisDeduction(AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Parent.Rotate), Degrees), ApplyTo.Rotate)
'                    ElseIf Not Camera.Planet Is Nothing Then
'                        Set ApplyTo.Relative.Rotate = Degrees
'                    Else
'                        'Set ApplyTo.Relative.Rotate = AngleAxisDeduction(AngleAxisAddition(ApplyTo.Rotate, Degrees), ApplyTo.Rotate)
'                        Set ApplyTo.Relative.Rotate = Degrees
'                    End If

                    Set ApplyTo.Relative.Rotate = Degrees
                Else
                    Set ApplyTo.Absolute.Rotate = Degrees
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    RotOri Degrees, Relative, m, ApplyTo
'                Next

        End Select
    End If
End Sub

Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitScaling(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
End Sub
Private Sub ScaExp(ByRef Scalar As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Abs(Scalar.X) <> 1 Or Abs(Scalar.Y) <> 1 Or Abs(Scalar.z) <> 1 Then
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        ScaExp Scalar, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            ScaExp Scalar, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'change all molecules with in the specified planets range
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Scaled = Scalar
                Else
                    Set ApplyTo.Absolute.Scaled = Scalar
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    ScaExp Scalar, Relative, m, ApplyTo
'                Next

        End Select
    End If
End Sub
Public Sub Displace(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, False, ApplyTo, Parent  'location is changing the origin to absolute
End Sub
Public Sub Balanced(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False
    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True
End Sub
Private Sub DisBal(ByRef Offset As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Offset.X <> 0 Or Offset.Y <> 0 Or Offset.z <> 0 Then
        Dim dist As Single
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        DisBal Offset, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0)
                        If p.Ranges.W - dist > 0 Then
                            DisBal Offset, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'change all molecules with in the specified planets range
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Offset = Offset
                Else
                    Set ApplyTo.Absolute.Offset = Offset
                End If

'                For Each m In RangedMolecules(ApplyTo)
'                    DisBal Offset, Relative, m, ApplyTo
'                Next
        End Select
    End If
End Sub

Private Function RangedMolecules(ByRef ApplyTo As Molecule) As NTNodes10.Collection
    Set RangedMolecules = New NTNodes10.Collection

'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyScaled Scalar, m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyScaled Scalar, m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyScaled Scalar, m, ApplyTo
'        Next
'    End If
    Dim m As Molecule
    Dim dist As Single
    For Each m In Molecules
        If ((m.Parent Is Nothing) And (Not TypeName(ApplyTo) = "Planet")) Or (TypeName(ApplyTo) = "Planet") Then
            If ApplyTo.Ranges.W = -1 Then
                RangedMolecules.Add m, m.Key
            ElseIf ApplyTo.Ranges.W > 0 Then
                dist = Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z)
                If ApplyTo.Ranges.W - dist > 0 Then
                    RangedMolecules.Add m, m.Key
                End If
            End If
        End If
    Next
    For Each m In ApplyTo.Molecules
        If Not RangedMolecules.Exists(m.Key) Then RangedMolecules.Add m, m.Key
    Next
End Function

Public Sub CommitRoutine(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal DoRotate As Boolean, ByVal DoScaled As Boolean, ByVal DoOrigin As Boolean, ByVal DoOffset As Boolean, ByVal DoAbsolute As Boolean, ByVal DoRelative As Boolean)
    'partial to committing a 3d objects properties during calls that may not sum, for retaining other properties needing change first and entirety per frame
    Static stacked As Boolean
    If Not stacked Then
        stacked = True

        'any absolute position comes first, pending is a difference from the actual
        If (Not ApplyTo.Absolute.Origin.Equals(ApplyTo.Origin)) And ((DoOrigin And DoAbsolute) Or ((Not DoOrigin) And (Not DoAbsolute))) Then
            ApplyOrigin VectorDeduction(ApplyTo.Absolute.Origin, ApplyTo.Origin), ApplyTo, Parent
            Set ApplyTo.Absolute.Origin = ApplyTo.Origin
        End If
        If (Not ApplyTo.Absolute.Rotate.Equals(ApplyTo.Rotate)) And ((DoRotate And DoAbsolute) Or ((Not DoRotate) And (Not DoAbsolute))) Then
            ApplyRotate AngleAxisDeduction(ApplyTo.Absolute.Rotate, ApplyTo.Rotate), ApplyTo, Parent
            Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
        End If
        If (Not ApplyTo.Absolute.Scaled.Equals(ApplyTo.Scaled)) And ((DoScaled And DoAbsolute) Or ((Not DoScaled) And (Not DoAbsolute))) Then
            ApplyScaled VectorDeduction(ApplyTo.Absolute.Scaled, ApplyTo.Scaled), ApplyTo, Parent
            Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
        End If
        If (Not ApplyTo.Absolute.Offset.Equals(ApplyTo.Offset)) And ((DoOffset And DoAbsolute) Or ((Not DoOffset) And (Not DoAbsolute))) Then
            ApplyOffset VectorDeduction(ApplyTo.Absolute.Offset, ApplyTo.Offset), ApplyTo, Parent
            Set ApplyTo.Absolute.Offset = ApplyTo.Offset
        End If
        
        'relative positioning comes secondly, pending is there is any value not empty
        If (ApplyTo.Relative.Offset.X <> 0 Or ApplyTo.Relative.Offset.Y <> 0 Or ApplyTo.Relative.Offset.z <> 0) And ((DoOffset And DoRelative) Or ((Not DoOffset) And (Not DoRelative))) Then
            ApplyOffset VectorAddition(ApplyTo.Relative.Offset, ApplyTo.Offset), ApplyTo, Parent
            Set ApplyTo.Relative.Offset = Nothing
        End If
        If (ApplyTo.Relative.Rotate.X <> 0 Or ApplyTo.Relative.Rotate.Y <> 0 Or ApplyTo.Relative.Rotate.z <> 0) And ((DoRotate And DoRelative) Or ((Not DoRotate) And (Not DoRelative))) Then
            ApplyRotate AngleAxisAddition(ApplyTo.Relative.Rotate, ApplyTo.Rotate), ApplyTo, Parent
            Set ApplyTo.Relative.Rotate = Nothing
        End If
        If (ApplyTo.Relative.Origin.X <> 0 Or ApplyTo.Relative.Origin.Y <> 0 Or ApplyTo.Relative.Origin.z <> 0) And ((DoOrigin And DoRelative) Or ((Not DoOrigin) And (Not DoRelative))) Then
            ApplyOrigin VectorAddition(ApplyTo.Relative.Origin, ApplyTo.Origin), ApplyTo, Parent
            Set ApplyTo.Relative.Origin = Nothing
        End If
        If (Abs(ApplyTo.Relative.Scaled.X) <> 1 Or Abs(ApplyTo.Relative.Scaled.Y) <> 1 Or Abs(ApplyTo.Relative.Scaled.z) <> 1) And ((DoScaled And DoRelative) Or ((Not DoScaled) And (Not DoRelative))) Then
            ApplyScaled VectorAddition(ApplyTo.Relative.Scaled, ApplyTo.Scaled), ApplyTo, Parent
            Set ApplyTo.Relative.Scaled = Nothing
        End If
        stacked = False
    End If
End Sub

Public Sub Begin(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    'called once per frame committing changes the last frame has waiting in object properties in entirety
    Dim m As Molecule
    Dim p As Planet

    For Each p In Planets

        CommitRoutine p, Nothing, True, False, False, False, True, False
        CommitRoutine p, Nothing, False, True, False, False, True, False
        CommitRoutine p, Nothing, False, False, True, False, True, False
        CommitRoutine p, Nothing, False, False, False, True, True, False

        CommitRoutine p, Nothing, True, False, False, False, False, True
        CommitRoutine p, Nothing, False, True, False, False, False, True
        CommitRoutine p, Nothing, False, False, True, False, False, True
        CommitRoutine p, Nothing, False, False, False, True, False, True

        Set p.Relative = Nothing

    Next

    For Each m In Molecules

        If m.Parent Is Nothing Then

            CommitRoutine m, Nothing, True, False, False, False, True, False
            CommitRoutine m, Nothing, False, True, False, False, True, False
            CommitRoutine m, Nothing, False, False, True, False, True, False
            CommitRoutine m, Nothing, False, False, False, True, True, False

            CommitRoutine m, Nothing, True, False, False, False, False, True
            CommitRoutine m, Nothing, False, True, False, False, False, True
            CommitRoutine m, Nothing, False, False, True, False, False, True
            CommitRoutine m, Nothing, False, False, False, True, False, True

            Set m.Relative = Nothing

        End If

    Next

End Sub

Public Sub Finish(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    'called once per frame drawing the objects, with out any of the current frame object
    'properties modifying calls included for latent collision checking rollback

    DDevice.SetRenderState D3DRS_ZENABLE, 1

    DDevice.SetRenderState D3DRS_CLIPPING, 1

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetMaterial LucentMaterial
    DDevice.SetTexture 0, Nothing
    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 1, Nothing
    
    Dim p As Planet
    For Each p In Planets
        Iterate p.Molecules, False
    Next

    Iterate Molecules, True
End Sub

Private Sub Iterate(ByRef col As Object, ByVal NoParentOnly As Boolean)

    Dim m As Molecule
    For Each m In col
        If NoParentOnly Then
            If m.Parent Is Nothing Then
                Render m, Nothing
            End If
        Else
            Render m, m.Parent
        End If
    Next

End Sub

Private Sub Render(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    Static stacked As Integer

    Dim matMat As D3DMATRIX

    If stacked = 0 Then
        D3DXMatrixIdentity matMat
    End If

    Dim matRoll As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matScale As D3DMATRIX

    If Not Camera.Planet Is Nothing Then
        D3DXMatrixRotationX matPitch, Camera.Planet.Rotate.X
        D3DXMatrixMultiply matMat, matPitch, matMat

        D3DXMatrixRotationY matYaw, Camera.Planet.Rotate.Y
        D3DXMatrixMultiply matMat, matYaw, matMat

        D3DXMatrixRotationZ matRoll, Camera.Planet.Rotate.z
        D3DXMatrixMultiply matMat, matRoll, matMat
        
        DDevice.SetTransform D3DTS_WORLD, matMat
    End If

    D3DXMatrixTranslation matPos, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z
    D3DXMatrixMultiply matMat, matPos, matMat
            
    DDevice.SetTransform D3DTS_WORLD, matMat
    
    D3DXMatrixRotationX matPitch, ApplyTo.Rotate.X
    D3DXMatrixMultiply matMat, matPitch, matMat

    D3DXMatrixRotationY matYaw, ApplyTo.Rotate.Y
    D3DXMatrixMultiply matMat, matYaw, matMat

    D3DXMatrixRotationZ matRoll, ApplyTo.Rotate.z
    D3DXMatrixMultiply matMat, matRoll, matMat
    
    DDevice.SetTransform D3DTS_WORLD, matMat

    Dim V As Matter
    For Each V In ApplyTo.Volume
        'update the directx and collision array's then render the object
        VertexDirectX((V.TriangleIndex * 3) + 0).X = V.Point1.X '+ ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 0).Y = V.Point1.Y '+ ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 0).z = V.Point1.z '+ ApplyTo.Origin.z

        VertexDirectX((V.TriangleIndex * 3) + 1).X = V.Point2.X '+ ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 1).Y = V.Point2.Y '+ ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 1).z = V.Point2.z '+ ApplyTo.Origin.z

        VertexDirectX((V.TriangleIndex * 3) + 2).X = V.Point3.X '+ ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 2).Y = V.Point3.Y '+ ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 2).z = V.Point3.z '+ ApplyTo.Origin.z

        If Not Parent Is Nothing Then

            VertexDirectX((V.TriangleIndex * 3) + 0).X = VertexDirectX((V.TriangleIndex * 3) + 0).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 0).Y = VertexDirectX((V.TriangleIndex * 3) + 0).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 0).z = VertexDirectX((V.TriangleIndex * 3) + 0).z + Parent.Origin.z

            VertexDirectX((V.TriangleIndex * 3) + 1).X = VertexDirectX((V.TriangleIndex * 3) + 1).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 1).Y = VertexDirectX((V.TriangleIndex * 3) + 1).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 1).z = VertexDirectX((V.TriangleIndex * 3) + 1).z + Parent.Origin.z

            VertexDirectX((V.TriangleIndex * 3) + 2).X = VertexDirectX((V.TriangleIndex * 3) + 2).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 2).Y = VertexDirectX((V.TriangleIndex * 3) + 2).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 2).z = VertexDirectX((V.TriangleIndex * 3) + 2).z + Parent.Origin.z

        End If

        VertexDirectX(V.TriangleIndex * 3 + 0).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 0).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 0).Nz = V.Normal.z

        VertexDirectX(V.TriangleIndex * 3 + 1).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 1).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 1).Nz = V.Normal.z

        VertexDirectX(V.TriangleIndex * 3 + 2).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 2).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 2).Nz = V.Normal.z

        VertexXAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).X
        VertexXAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).X
        VertexXAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).X

        VertexYAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).Y
        VertexYAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).Y
        VertexYAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).Y

        VertexZAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).z
        VertexZAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).z
        VertexZAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).z

         If ApplyTo.Visible And (Not (TypeName(ApplyTo) = "Planet")) Then
             If Not (V.Translucent Or V.Transparent) Then
                 DDevice.SetMaterial GenericMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
                 DDevice.SetTexture 1, Nothing
             Else
                 DDevice.SetMaterial LucentMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
                 DDevice.SetMaterial GenericMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 1, Files(V.TextureIndex).Data
             End If

             DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VertexDirectX((V.TriangleIndex * 3)), Len(VertexDirectX(0))
         End If
    Next

    stacked = stacked + 1
    Dim m As Molecule

    For Each m In ApplyTo.Molecules
        Render m, ApplyTo
    Next
    stacked = stacked - 1
End Sub

Private Function BuildArrays() As Long
    ReDim Preserve TriangleFace(0 To 5, 0 To TriangleCount) As Single
    ReDim Preserve VertexXAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexYAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexZAxis(0 To 2, 0 To TriangleCount) As Single
    BuildArrays = (((TriangleCount + 1) * 3) - 1)
    ReDim Preserve VertexDirectX(0 To BuildArrays) As MyVertex
    ReDim Preserve ScreenDirectX(0 To BuildArrays) As MyScreen
    BuildArrays = BuildArrays - 2
End Function

Public Function CreateMoleculeFace(ByRef TextureFileName As String, ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, ByRef P4 As Point, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Molecule
    If (((Not (p1.Equals(p2) Or p1.Equals(p3) Or p1.Equals(P4))) And _
        (Not (p3.Equals(p2) Or p3.Equals(p1) Or p3.Equals(P4))) And _
        (Not (p2.Equals(p1) Or p2.Equals(p3) Or p2.Equals(P4))) And _
        (Not (P4.Equals(p2) Or P4.Equals(p3) Or P4.Equals(p1)))) And _
        PathExists(TextureFileName, True)) Then
        
        Dim r As New Molecule
        Set r.Volume = CreateVolumeFace(TextureFileName, p1, p2, p3, P4, ScaleX, ScaleY)
        r.Visible = True
        Set CreateMoleculeFace = r
    End If
End Function
Public Function CreateVolumeFace(ByRef TextureFileName As String, ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, ByRef P4 As Point, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Volume
    If (((Not (p1.Equals(p2) Or p1.Equals(p3) Or p1.Equals(P4))) And _
        (Not (p3.Equals(p2) Or p3.Equals(p1) Or p3.Equals(P4))) And _
        (Not (p2.Equals(p1) Or p2.Equals(p3) Or p2.Equals(P4))) And _
        (Not (P4.Equals(p2) Or P4.Equals(p3) Or P4.Equals(p1)))) And _
        PathExists(TextureFileName, True)) Then
        If ScaleX = 0 Then ScaleX = 1
        If ScaleY = 0 Then ScaleY = 1
        
        Dim vol As New Volume
        Dim m As New Matter
        With m
            .TriangleIndex = TriangleCount
            BuildArrays

            .Index1 = PointCache(p1)
            .Index2 = PointCache(p2)
            .Index3 = PointCache(p3)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .V1 = ScaleY
            .U2 = ScaleX
            .V2 = ScaleY
            .U3 = ScaleX

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.z
            VertexZAxis(1, TriangleCount) = .Point2.z
            VertexZAxis(2, TriangleCount) = .Point3.z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 0

            .ObjectIndex = ObjectCount
            .FaceIndex = 0
            .TextureIndex = GetFileIndex(TextureFileName)
            If TextureFileName <> "" Then
                If Files(.TextureIndex).Data Is Nothing Then
                    Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                    ImageDimensions TextureFileName, Files(.TextureIndex).Size
                End If
            End If
            
            VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
            VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
            ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
            ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
            ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
            ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
            ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
            ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
            ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
            ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
            ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
        End With
        TriangleCount = TriangleCount + 1
        vol.Add m
                
        Set m = New Matter
        With m

            .TriangleIndex = TriangleCount
            BuildArrays

            .Index1 = PointCache(p1)
            .Index2 = PointCache(p3)
            .Index3 = PointCache(P4)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .V1 = ScaleY
            .U2 = ScaleX

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.z
            VertexZAxis(1, TriangleCount) = .Point2.z
            VertexZAxis(2, TriangleCount) = .Point3.z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 1

            .ObjectIndex = ObjectCount
            .FaceIndex = 1
            .TextureIndex = GetFileIndex(TextureFileName)
            If TextureFileName <> "" Then
                If Files(.TextureIndex).Data Is Nothing Then
                    Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                    ImageDimensions TextureFileName, Files(.TextureIndex).Size
                End If
            End If
            
            VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
            VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
            ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
            ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
            ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
            ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
            ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
            ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
            ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
            ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
            ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
        End With
        TriangleCount = TriangleCount + 1
        vol.Add m
        
        ObjectCount = ObjectCount + 1
        
        Set CreateVolumeFace = vol

    End If
End Function

Public Function CreateMoleculeLanding(ByRef TextureFileName As String, ByVal OuterRadii As Single, ByVal RadiiSegments As Single, Optional ByVal InnerRadii As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Molecule
    If OuterRadii <= 0 Then
        Err.Raise 8, , "OuterRadii must be above 0."
    ElseIf InnerRadii < 0 Then
        Err.Raise 8, , "InnerRadii must be 0 or above."
    ElseIf OuterRadii < InnerRadii Then
        Err.Raise 8, , "OuterRadii must be below InnerRadii."
    ElseIf RadiiSegments < 3 Then
        Err.Raise 8, , "RadiiSegments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        
        Dim r As New Molecule
        Set r.Volume = CreateVolumeLanding(TextureFileName, OuterRadii, RadiiSegments, InnerRadii, ScaleX, ScaleY)
        r.Visible = True
        Set CreateMoleculeLanding = r
    End If

End Function


Public Function CreateVolumeLanding(ByRef TextureFileName As String, ByVal OuterRadii As Single, ByVal RadiiSegments As Single, Optional ByVal InnerRadii As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Volume
    If OuterRadii <= 0 Then
        Err.Raise 8, , "OuterRadii must be above 0."
    ElseIf InnerRadii < 0 Then
        Err.Raise 8, , "InnerRadii must be 0 or above."
    ElseIf OuterRadii < InnerRadii Then
        Err.Raise 8, , "OuterRadii must be below InnerRadii."
    ElseIf RadiiSegments < 3 Then
        Err.Raise 8, , "RadiiSegments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        Dim i As Long
        Dim g As Single
        Dim A As Double
        Dim l1 As Single
        Dim l2 As Single
        Dim l3 As Single
    
        Dim intX1 As Single
        Dim intX2 As Single
        Dim intX3 As Single
        Dim intX4 As Single
    
        Dim intY1 As Single
        Dim intY2 As Single
        Dim intY3 As Single
        Dim intY4 As Single
        Dim dist1 As Single
        Dim dist2 As Single
        Dim dist3 As Single
        Dim dist4 As Single

        Dim vol As New Volume
        Dim m As Matter

        
        For i = -IIf(InnerRadii > 0, 6, 3) To ((IIf(InnerRadii > 0, 6, 3) * RadiiSegments) - 1) + (IIf(InnerRadii > 0, 6, 3) * 2) Step IIf(InnerRadii > 0, 6, 3)
    
            g = (((360 / RadiiSegments) * (((i + 1) / IIf(InnerRadii > 0, 6, 3)) - 1)) * RADIAN)
    
            intX2 = (OuterRadii * Sin(g))
            intY2 = (-OuterRadii * Cos(g))
            intX3 = (InnerRadii * Sin(g))
            intY3 = (-InnerRadii * Cos(g))
    
            If i >= 0 Then
            
                If (InnerRadii > 0) Then
                    If (i Mod 12) = 0 Then
    
                        dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
                        dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
    
                    End If

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .V1 = dist2
                        .U2 = dist1
                        .V2 = dist4
                        .U3 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
            
                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                    
                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX4, 0, intY4))
                        .Index3 = PointCache(MakePoint(intX3, 0, intY3))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .V1 = dist2
                        .U2 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                    
                Else
    
                    dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * IIf(i Mod 4 = 0, 1, Sin(g))
                    dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * IIf(i Mod 4 = 0, 1, Cos(g))

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .U1 = ((ScaleX / dist1) * (dist1 / 100))
                        .V2 = ((ScaleY / dist2) * (dist2 / 100))
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                End If
    
            End If
    
            intX1 = intX2
            intY1 = intY2
            intX4 = intX3
            intY4 = intY3
    
        Next
        
        ObjectCount = ObjectCount + 1

        Set CreateVolumeLanding = vol
    End If
    
End Function

Public Function CreateMoleculeMesh(ByVal DirectXFileName As String) As Molecule
    If PathExists(DirectXFileName, True) Then
    
        Dim r As New Molecule
        Set r.Volume = CreateVolumeMesh(DirectXFileName)
        r.Visible = True
        Set CreateMoleculeMesh = r
    End If
End Function
Public Function CreateVolumeMesh(ByVal DirectXFileName As String) As Volume
    If PathExists(DirectXFileName, True) Then

        Dim MeshVerticies() As D3DVERTEX
        Dim MeshIndicies() As Integer

        Dim nMaterials As Long
        Dim nMatBuffer As D3DXBuffer
                
        Dim Mesh As D3DXMesh
        Set Mesh = D3DX.LoadMeshFromX(DirectXFileName, D3DXMESH_DYNAMIC, DDevice, Nothing, nMatBuffer, nMaterials)

        Dim Index As Long
        Dim cnt As Long

        Dim VD As D3DVERTEXBUFFER_DESC
        Mesh.GetVertexBuffer.GetDesc VD
        ReDim MeshVerticies(0 To 0) As D3DVERTEX
        ReDim MeshVerticies(0 To ((VD.Size \ Len(MeshVerticies(0))) - 1)) As D3DVERTEX
        D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, VD.Size, 0, MeshVerticies(0)

        Dim ID As D3DINDEXBUFFER_DESC
        Mesh.GetIndexBuffer.GetDesc ID
        ReDim MeshIndicies(0 To 0) As Integer
        ReDim MeshIndicies(0 To ((ID.Size \ Len(MeshIndicies(0))) - 1)) As Integer
        D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, ID.Size, 0, MeshIndicies(0)



'        If nMaterials > 0 Then
'
'            Dim P1 As String
'            Dim P2 As String
'            Dim P3 As String
'            Dim P4 As String
'            Dim p5 As String
'            Dim p6 As String
'
'            Dim l3 As Single
'            Dim l1 As Single
'            Dim l2 As Single
'            Dim l4 As Single
'            Dim l5 As Single
'            Dim l6 As Single
'
'            Dim checked As Long
'
'            Index = 0 'start at last point of first triangle where start = 0
'            Do Until checked = Mesh.GetNumFaces  'go for amount of faces least 3
'
'                 If l1 = 0 Then
'                    l1 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P1 = IIf(checked Mod 4 = 0, "1,0", "4,3")
'                    l4 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P4 = IIf(checked Mod 4 = 0, "2,0", "5,3")
'                End If
'
'                 If l2 = 0 Then
'                    l2 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'                    P2 = IIf(checked Mod 4 = 0, "1,2", "4,5")
'                    l5 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    p5 = IIf(checked Mod 4 = 0, "1,0", "4,3")
'                End If
'
'                 If l3 = 0 Then
'                    l3 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P3 = IIf(checked Mod 4 = 0, "2,0", "5,3")
'                    l6 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'                    p6 = IIf(checked Mod 4 = 0, "1,2", "4,5")
'                End If
'
'                 Index = Index + 1
'
'                 If l1 = l2 And l2 = l3 And l1 <> 0 Then
'
'                     l2 = 0
'                     l6 = l4
'                     l4 = l5
'                     l5 = 0
'                 Else
'
'                    If l1 <> 0 And l2 <> 0 And l3 <> 0 Then
'
'
'
'
'                        Debug.Print "(" & (Index + CLng(NextArg(P1, ","))) & ", " & (Index + CLng(RemoveArg(P2, ","))) & ", " & (Index + CLng(NextArg(P3, ","))) & ") ";
'                        Debug.Print "(" & (Index + CLng(RemoveArg(P4, ","))) & ", " & (Index + CLng(NextArg(p5, ","))) & ", " & (Index + CLng(NextArg(p6, ","))) & ") ";
'
'                        'SurfaceArea = SurfaceArea + (TriangleAreaByLen(l1, l2, l3) + TriangleAreaByLen(l4, l5, l6))
'
'                        'Volume = Volume + (TriangleVolByLen(l1, l2, l3) + TriangleVolByLen(l4, l5, l6))
'
'                        l1 = 0
'                        l2 = 0
'                        l3 = 0
'                        l4 = 0
'                        l5 = 0
'                        l6 = 0
'                        checked = checked + 2
'                    End If
'
'                    Index = Index + 2
'                End If
'
'            Loop
'
'        End If
'
        'SurfaceArea = (SurfaceArea * 2)





        Dim vol As New Volume
        Dim m As Matter

        Dim Verts() As D3DVERTEX
        Dim chk1 As Integer
        Dim chk2 As Integer
        Dim chk3 As Integer
        Dim chk4 As Integer
        Dim chk5 As Integer
        Dim chk6 As Integer
        chk1 = -1

        Const TrianglePerFace As Long = 2
        Dim TextureFileName As String

        Dim txt As Long
        txt = 1

        Index = 0

        Do While Index <= UBound(MeshIndicies)

            If txt < nMaterials Then
                If D3DX.BufferGetTextureName(nMatBuffer, txt) <> "" Then
                    If PathExists(GetFilePath(DirectXFileName) & "\" & D3DX.BufferGetTextureName(nMatBuffer, txt), True) Then
                        TextureFileName = GetFilePath(DirectXFileName) & "\" & D3DX.BufferGetTextureName(nMatBuffer, txt)
                    End If
                End If
            End If
            txt = txt + 1

            Set m = New Matter
            With m
                .TriangleIndex = TriangleCount
                BuildArrays

                .Index1 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 0)).X, _
                    MeshVerticies(MeshIndicies(Index + 0)).Y, _
                    MeshVerticies(MeshIndicies(Index + 0)).z))

                .Index2 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 1)).X, _
                    MeshVerticies(MeshIndicies(Index + 1)).Y, _
                    MeshVerticies(MeshIndicies(Index + 1)).z))

                .Index3 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 2)).X, _
                    MeshVerticies(MeshIndicies(Index + 2)).Y, _
                    MeshVerticies(MeshIndicies(Index + 2)).z))

                Set .Point1 = Points(.Index1)
                Set .Point2 = Points(.Index2)
                Set .Point3 = Points(.Index3)

                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.z
                VertexZAxis(1, TriangleCount) = .Point2.z
                VertexZAxis(2, TriangleCount) = .Point3.z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.z
                TriangleFace(4, TriangleCount) = ObjectCount
                TriangleFace(5, TriangleCount) = (Index \ 4) + 0

                .ObjectIndex = ObjectCount
                .FaceIndex = (Index \ 4) + 0
                .TextureIndex = GetFileIndex(TextureFileName)
                If TextureFileName <> "" Then
                    If Files(.TextureIndex).Data Is Nothing Then
                        Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                        ImageDimensions TextureFileName, Files(.TextureIndex).Size
                    End If
                End If

                VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 0)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 0)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 1)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 1)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 2).tu = MeshVerticies(MeshIndicies(Index + 2)).tu
                VertexDirectX(.TriangleIndex * 3 + 2).tv = MeshVerticies(MeshIndicies(Index + 2)).tv

            End With
            TriangleCount = TriangleCount + 1
            vol.Add m

            Set m = New Matter
            With m
                .TriangleIndex = TriangleCount
                BuildArrays

                .Index1 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 3)).X, _
                    MeshVerticies(MeshIndicies(Index + 3)).Y, _
                    MeshVerticies(MeshIndicies(Index + 3)).z))

                .Index2 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 4)).X, _
                    MeshVerticies(MeshIndicies(Index + 4)).Y, _
                    MeshVerticies(MeshIndicies(Index + 4)).z))

                .Index3 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 5)).X, _
                    MeshVerticies(MeshIndicies(Index + 5)).Y, _
                    MeshVerticies(MeshIndicies(Index + 5)).z))

                Set .Point1 = Points(.Index1)
                Set .Point2 = Points(.Index2)
                Set .Point3 = Points(.Index3)

                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.z
                VertexZAxis(1, TriangleCount) = .Point2.z
                VertexZAxis(2, TriangleCount) = .Point3.z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.z
                TriangleFace(4, TriangleCount) = ObjectCount
                TriangleFace(5, TriangleCount) = (Index \ 4) + 1

                .ObjectIndex = ObjectCount
                .FaceIndex = (Index \ 4) + 1
                .TextureIndex = GetFileIndex(TextureFileName)
                If TextureFileName <> "" Then
                    If Files(.TextureIndex).Data Is Nothing Then
                        Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                        ImageDimensions TextureFileName, Files(.TextureIndex).Size
                    End If
                End If

                VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 3)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 3)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 4)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 4)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 2).tu = MeshVerticies(MeshIndicies(Index + 5)).tu
                VertexDirectX(.TriangleIndex * 3 + 2).tv = MeshVerticies(MeshIndicies(Index + 5)).tv

            End With

            TriangleCount = TriangleCount + 1
            Index = Index + 6
            vol.Add m

        Loop

        ObjectCount = ObjectCount + 1

        Set CreateVolumeMesh = vol
    End If
End Function


Public Function PointCache(ByRef p As Point) As Long
    Points.Add p
    PointCache = Points.Count
    Exit Function
    If Points.Count > 0 Then
        Dim i As Long
        For i = 1 To Points.Count
            If Points(i).Serialize = p.Serialize Then
                PointCache = i
                Set p = Points(i)
                Exit Function
            End If
        Next
    End If
    Points.Add p, p.Serialize
    PointCache = Points.Count
End Function
