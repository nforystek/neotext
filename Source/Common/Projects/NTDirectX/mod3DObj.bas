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


Public CameraAim() As Single
'CameraAim dimension (0,n) is camera position, n=0=x, n=1=y, n=2=z
'CameraAim dimension (1,n) is camera direction, n=0=x, n=1=y, n=2=z
'CameraAim dimension (2,n) is camera up vector, n=0=x, n=1=y, n=2=z

'the following are used in culling with the maxlandlib.dll (forystek)
'Aling with CameraAim they contain information to
Public ScreenX() As Single
Public ScreenY() As Single
Public ScreenZ() As Single
Public ZBuffer() As Single

'below are for DirectX, rendering 3D and 2D are pre built
Public VertexDirectX() As MyVertex
Public ScreenDirectX() As MyScreen

Public Sub CleanUpObjs()
    
    ObjectCount = 0
    TriangleCount = 0
    Erase TriangleFace
    
    Erase VertexXAxis
    Erase VertexYAxis
    Erase VertexZAxis

    Erase VertexDirectX
    Erase ScreenDirectX
    
    Erase ScreenX
    Erase ScreenY
    Erase ScreenZ
    Erase ZBuffer
    
    Erase CameraAim

End Sub

Public Sub CreateObjs()

    ReDim CameraAim(0 To 2, 0 To 2) As Single
         
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
    MsgBox Err.description, vbExclamation
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
    MsgBox Err.description, vbExclamation
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

Private Function CombineOrbits(ByRef o1 As Orbit, ByRef o2 As Orbit) As Orbit
    Set CombineOrbits = New Orbit
    With CombineOrbits
        .Origin = VectorAddition(o1.Origin, o2.Origin)
        .Offset = VectorDeduction(VectorDeduction(VectorAddition(o1.Origin, o1.Offset), VectorAddition(o2.Origin, o2.Offset)), .Origin)
        .Rotate = VectorAddition(o1.Rotate, o2.Rotate)
        .Scaled = VectorAddition(o1.Scaled, o2.Scaled)
        .Ranges.X = o1.Ranges.X + o2.Ranges.X
        .Ranges.Y = o1.Ranges.Y + o2.Ranges.Y
        .Ranges.z = o1.Ranges.z + o2.Ranges.z
        If o1.Ranges.W = -1 Or o2.Ranges.W = -1 Then
            .Ranges.W = -1
        ElseIf o1.Ranges.W > o2.Ranges.W Then
            .Ranges.W = o1.Ranges.W
        Else
            .Ranges.W = o2.Ranges.W
        End If
    End With
End Function




Public Sub RenderMolecules(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
    Dim p As Planet
    Dim m As Molecule
    
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
    
    Dim matRoll As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPos As D3DMATRIX
        
    Dim matMat As D3DMATRIX
    D3DXMatrixIdentity matMat
    DDevice.SetTransform D3DTS_WORLD, matMat
        
    If Not Camera.Planet Is Nothing Then
        
        D3DXMatrixRotationX matPitch, Camera.Planet.Rotate.X
        D3DXMatrixMultiply matMat, matPitch, matMat

        D3DXMatrixRotationY matYaw, Camera.Planet.Rotate.Y
        D3DXMatrixMultiply matMat, matYaw, matMat
        
        D3DXMatrixRotationZ matRoll, Camera.Planet.Rotate.z
        D3DXMatrixMultiply matMat, matRoll, matMat
        
        DDevice.SetTransform D3DTS_WORLD, matMat
        
   End If
    
    RenderOrbits Molecules, True
'    For Each m In Molecules
'        If m.Parent Is Nothing Then
'            AllCommitRoutine m, Nothing
'        End If
'    Next
    
    If Not Camera.Planet Is Nothing Then

        D3DXMatrixIdentity matMat
    
        D3DXMatrixTranslation matPos, Camera.Planet.Origin.X, Camera.Planet.Origin.Y, Camera.Planet.Origin.z
        D3DXMatrixMultiply matMat, matPos, matMat
        
        D3DXMatrixRotationX matPitch, Camera.Planet.Rotate.X
        D3DXMatrixMultiply matMat, matPitch, matMat
        
        D3DXMatrixRotationY matYaw, Camera.Planet.Rotate.Y
        D3DXMatrixMultiply matMat, matYaw, matMat
        
        D3DXMatrixRotationZ matRoll, Camera.Planet.Rotate.z
        D3DXMatrixMultiply matMat, matRoll, matMat
                
        DDevice.SetTransform D3DTS_WORLD, matMat

    End If
    
    For Each p In Planets
        RenderOrbits p.Molecules, False
    Next
'    For Each p In Planets
'        AllCommitRoutine p, Nothing
''        Set ms = RangedMolecules(p)
''        For Each m In ms
''            AllCommitRoutine m, p
''        Next
'    Next
    
End Sub

Private Sub RenderOrbits(ByRef col As Object, ByVal NoParentOnly As Boolean)

    Dim matMat As D3DMATRIX
    D3DXMatrixIdentity matMat
    
    If Not col Is Nothing Then
    
    Dim m As Molecule
    For Each m In col
        If NoParentOnly Then
            If m.Parent Is Nothing Then
                RenderMolecule m, Nothing, matMat
            End If
        Else
            RenderMolecule m, Nothing, matMat
        End If
        
    Next
    End If
End Sub

Private Sub RenderMolecule(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByRef matMat As D3DMATRIX)

    Dim vout As D3DVECTOR

    Dim matPos As D3DMATRIX
    Dim matRoll As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matScale As D3DMATRIX
    Dim matRot As D3DMATRIX
    
    D3DXMatrixTranslation matPos, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z
    D3DXMatrixMultiply matMat, matPos, matMat
   
    D3DXMatrixRotationX matPitch, ApplyTo.Rotate.X
    D3DXMatrixMultiply matMat, matPitch, matMat
     
    D3DXMatrixRotationY matYaw, ApplyTo.Rotate.Y
    D3DXMatrixMultiply matMat, matYaw, matMat

    D3DXMatrixRotationZ matRoll, ApplyTo.Rotate.z
    D3DXMatrixMultiply matMat, matRoll, matMat
    
    D3DXMatrixTranslation matPos, ApplyTo.Offset.X, ApplyTo.Offset.Y, ApplyTo.Offset.z
    D3DXMatrixMultiply matMat, matPos, matMat
    
    D3DXMatrixScaling matScale, ApplyTo.Scaled.X, ApplyTo.Scaled.Y, ApplyTo.Scaled.z
    D3DXMatrixMultiply matScale, matScale, matMat
    
    Dim m As Molecule
    If Not ApplyTo.Molecules Is Nothing Then
        For Each m In ApplyTo.Molecules
            RenderMolecule m, ApplyTo, matMat
        Next
    End If
    
    Dim v As Matter
    If Not ApplyTo.Volume Is Nothing Then
        For Each v In ApplyTo.Volume
    
            D3DXVec3TransformCoord vout, ToVector(v.Point1), matScale
            VertexDirectX((v.TriangleIndex * 3) + 0).X = vout.X
            VertexDirectX((v.TriangleIndex * 3) + 0).Y = vout.Y
            VertexDirectX((v.TriangleIndex * 3) + 0).z = vout.z
    
            D3DXVec3TransformCoord vout, ToVector(v.Point2), matScale
            VertexDirectX((v.TriangleIndex * 3) + 1).X = vout.X
            VertexDirectX((v.TriangleIndex * 3) + 1).Y = vout.Y
            VertexDirectX((v.TriangleIndex * 3) + 1).z = vout.z
    
            D3DXVec3TransformCoord vout, ToVector(v.Point3), matScale
            VertexDirectX((v.TriangleIndex * 3) + 2).X = vout.X
            VertexDirectX((v.TriangleIndex * 3) + 2).Y = vout.Y
            VertexDirectX((v.TriangleIndex * 3) + 2).z = vout.z
    
            VertexDirectX(v.TriangleIndex * 3 + 0).NX = v.Normal.X
            VertexDirectX(v.TriangleIndex * 3 + 0).NY = v.Normal.Y
            VertexDirectX(v.TriangleIndex * 3 + 0).Nz = v.Normal.z
    
            VertexDirectX(v.TriangleIndex * 3 + 1).NX = v.Normal.X
            VertexDirectX(v.TriangleIndex * 3 + 1).NY = v.Normal.Y
            VertexDirectX(v.TriangleIndex * 3 + 1).Nz = v.Normal.z
    
            VertexDirectX(v.TriangleIndex * 3 + 2).NX = v.Normal.X
            VertexDirectX(v.TriangleIndex * 3 + 2).NY = v.Normal.Y
            VertexDirectX(v.TriangleIndex * 3 + 2).Nz = v.Normal.z
    
            VertexXAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).X
            VertexXAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).X
            VertexXAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).X
    
            VertexYAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).Y
            VertexYAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).Y
            VertexYAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).Y
    
            VertexZAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).z
            VertexZAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).z
            VertexZAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).z
    
             If ApplyTo.Visible And (Not (TypeName(ApplyTo) = "Planet")) Then
                 If Not (v.Translucent Or v.Transparent) Then
                     DDevice.SetMaterial GenericMaterial
                     If v.TextureIndex > 0 Then DDevice.SetTexture 0, Files(v.TextureIndex).Data
                     DDevice.SetTexture 1, Nothing
                 Else
                     DDevice.SetMaterial LucentMaterial
                     If v.TextureIndex > 0 Then DDevice.SetTexture 0, Files(v.TextureIndex).Data
                     DDevice.SetMaterial GenericMaterial
                     If v.TextureIndex > 0 Then DDevice.SetTexture 1, Files(v.TextureIndex).Data
                 End If
    
                 DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VertexDirectX((v.TriangleIndex * 3)), Len(VertexDirectX(0))
             End If
        Next
    End If
    
    D3DXMatrixTranslation matPos, -ApplyTo.Offset.X, -ApplyTo.Offset.Y, -ApplyTo.Offset.z
    D3DXMatrixMultiply matMat, matPos, matMat
    
    D3DXMatrixRotationZ matRoll, -ApplyTo.Rotate.z
    D3DXMatrixMultiply matMat, matRoll, matMat
    
    D3DXMatrixRotationY matYaw, -ApplyTo.Rotate.Y
    D3DXMatrixMultiply matMat, matYaw, matMat
    
    D3DXMatrixRotationX matPitch, -ApplyTo.Rotate.X
    D3DXMatrixMultiply matMat, matPitch, matMat
    
    D3DXMatrixTranslation matPos, -ApplyTo.Origin.X, -ApplyTo.Origin.Y, -ApplyTo.Origin.z
    D3DXMatrixMultiply matMat, matPos, matMat
    
End Sub

Public Function RebuildTriangleArray() As Long
    If TriangleCount >= 0 Then
        ReDim Preserve TriangleFace(0 To 5, 0 To TriangleCount) As Single
        
        ReDim Preserve VertexXAxis(0 To 2, 0 To TriangleCount) As Single
        ReDim Preserve VertexYAxis(0 To 2, 0 To TriangleCount) As Single
        ReDim Preserve VertexZAxis(0 To 2, 0 To TriangleCount) As Single

        ReDim Preserve ScreenX(0 To 2, 0 To TriangleCount) As Single
        ReDim Preserve ScreenY(0 To 2, 0 To TriangleCount) As Single
        ReDim Preserve ScreenZ(0 To 2, 0 To TriangleCount) As Single
        ReDim Preserve ZBuffer(0 To 3, 0 To TriangleCount) As Single
        
        RebuildTriangleArray = (((TriangleCount + 1) * 3) - 1)
        ReDim Preserve VertexDirectX(0 To RebuildTriangleArray) As MyVertex
        ReDim Preserve ScreenDirectX(0 To RebuildTriangleArray) As MyScreen
        RebuildTriangleArray = RebuildTriangleArray - 2
    Else
        Erase TriangleFace
        Erase VertexXAxis
        Erase VertexYAxis
        Erase VertexZAxis
        Erase VertexDirectX
        Erase ScreenDirectX
        TriangleCount = -1
    End If
End Function

Public Function RemoveTriangleArray(ByRef TriangleIndex As Long)
    Dim i As Long
    If TriangleIndex < TriangleCount - 1 Then
        For i = TriangleIndex To TriangleCount - 2
            TriangleFace(0, i) = TriangleFace(0, i + 1)
            TriangleFace(1, i) = TriangleFace(1, i + 1)
            TriangleFace(2, i) = TriangleFace(2, i + 1)
            TriangleFace(3, i) = TriangleFace(3, i + 1)
            TriangleFace(4, i) = TriangleFace(4, i + 1)
            TriangleFace(5, i) = TriangleFace(5, i + 1)
            VertexXAxis(0, i) = VertexXAxis(0, i + 1)
            VertexXAxis(1, i) = VertexXAxis(1, i + 1)
            VertexXAxis(2, i) = VertexXAxis(2, i + 1)
            VertexYAxis(0, i) = VertexYAxis(0, i + 1)
            VertexYAxis(1, i) = VertexYAxis(1, i + 1)
            VertexYAxis(2, i) = VertexYAxis(2, i + 1)
            VertexZAxis(0, i) = VertexZAxis(0, i + 1)
            VertexZAxis(1, i) = VertexZAxis(1, i + 1)
            VertexZAxis(2, i) = VertexZAxis(2, i + 1)
        Next
        For i = (TriangleIndex * 3) To ((TriangleCount - 2) * 3)
            VertexDirectX(i) = VertexDirectX(i + 3)
            ScreenDirectX(i) = ScreenDirectX(i + 3)
        Next
    End If
    TriangleCount = TriangleCount - 1
    RebuildTriangleArray
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
        Dim offsetX As Single
        Dim offsetY As Single
        ScaleX = 1
        ScaleY = 1
        offsetX = 0.01
        offsetY = 0.01
        
        Dim vol As New Volume
        Dim m As New Matter
        With m
            .TriangleIndex = TriangleCount
            RebuildTriangleArray

            Set .Point1 = p1
            Set .Point2 = p2
            Set .Point3 = p3

'            .U1 = 0
'            .V1 = ScaleY
'            .U2 = ScaleX
'            .V2 = ScaleY
'            .U3 = ScaleX
'            .V3 = 0

            .U1 = 0
            .V1 = 0
            .U2 = ScaleX
            .V2 = 0
            .U3 = ScaleX
            .V3 = ScaleY
            
            Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
            RebuildTriangleArray

            Set .Point1 = p1
            Set .Point2 = p3
            Set .Point3 = P4
            
'            .U1 = 0
'            .V1 = ScaleY
'            .U2 = ScaleX
'            .V2 = 0
'            .U3 = 0
'            .V3 = 0

            .U1 = 0
            .V1 = 0
            .U2 = ScaleX
            .V2 = ScaleY
            .U3 = 0
            .V3 = ScaleY
            
            Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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

Public Function CreateMoleculeLanding(ByRef TextureFileName As String, ByVal OuterEdge As Single, ByVal Segments As Single, Optional ByVal InnerEdge As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1, Optional DiagnalTexture As Integer = 0) As Molecule
    If OuterEdge <= 0 Then
        Err.Raise 8, , "OuterEdge must be above 0."
    ElseIf InnerEdge < 0 Then
        Err.Raise 8, , "InnerEdge must be 0 or above."
    ElseIf OuterEdge < InnerEdge Then
        Err.Raise 8, , "OuterEdge must be below InnerEdge."
    ElseIf Segments < 3 Then
        Err.Raise 8, , "Segments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        
        Dim r As New Molecule
        Set r.Volume = CreateVolumeLanding(TextureFileName, OuterEdge, Segments, InnerEdge, ScaleX, ScaleY, DiagnalTexture)
        r.Visible = True
        Set CreateMoleculeLanding = r
    End If

End Function


Public Function CreateVolumeLanding(ByRef TextureFileName As String, ByVal OuterEdge As Single, ByVal Segments As Single, Optional ByVal InnerEdge As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1, Optional DiagnalTexture As Integer = 0) As Volume
    If OuterEdge <= 0 Then
        Err.Raise 8, , "OuterEdge must be above 0."
    ElseIf InnerEdge < 0 Then
        Err.Raise 8, , "InnerEdge must be 0 or above."
    ElseIf OuterEdge < InnerEdge Then
        Err.Raise 8, , "OuterEdge must be below InnerEdge."
    ElseIf Segments < 3 Then
        Err.Raise 8, , "Segments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        Dim i As Long
        Dim g As Single
        Dim A As Double
        Dim l1 As Single
        Dim l2 As Single
        Dim l3 As Single
        Dim X As Single
        Dim Y As Single
        
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
        Dim pointsPerFace As Integer
        pointsPerFace = IIf(InnerEdge > 0, 6, 3)
        
        For i = -pointsPerFace To ((pointsPerFace * Segments) - 1) + (pointsPerFace * 2) Step pointsPerFace

            g = (((360 / Segments) * (i / pointsPerFace)) * RADIAN)
            A = (g * DEGREE)
            A = Round(IIf(A < 0, A + 360, IIf(A > 360, A - 360, A)), 0)

            intX2 = (OuterEdge * Sin(A * RADIAN))
            intY2 = (-OuterEdge * Cos(A * RADIAN))
                
            If (InnerEdge > 0) Then
                intX3 = (InnerEdge * Sin(A * RADIAN))
                intY3 = (-InnerEdge * Cos(A * RADIAN))
            End If
            
            If i >= 0 Then
                If (InnerEdge > 0) Then

                    
                    If (i Mod 12) = 0 Then

                        If DiagnalTexture = 1 Then
                            dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) '* IIf(i Mod 4 = 0, 1, -Sin(g)) 'base of trapazoid
                            dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) '* IIf(i Mod 4 = 0, 1, -Cos(g)) 'angled side of trapazoid
                            dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) '* IIf(i Mod 4 = 0, 1, -Sin(g)) 'smaller top edge of trapazoid
                            dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) '* IIf(i Mod 4 = 0, 1, -Cos(g)) 'diagnal inside the trapazoid
                        Else
                            dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                            dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
                            dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                            dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
                        End If

                    End If

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        RebuildTriangleArray

                        If DiagnalTexture = 1 Then
                            Set .Point1 = MakePoint(intX4, 0, intY4)      '1-2
                            Set .Point2 = MakePoint(intX2, 0, intY2)     '|/|
                            Set .Point3 = MakePoint(intX1, 0, intY1)     '4-3
                        Else
                            Set .Point1 = MakePoint(intX2, 0, intY2)
                            Set .Point2 = MakePoint(intX1, 0, intY1)
                            Set .Point3 = MakePoint(intX4, 0, intY4)
                        End If

                        If DiagnalTexture = 1 Then
                            .U1 = dist3
                            .V1 = 0
                            .U2 = 0
                            .V2 = dist2
                            .U3 = dist1
                            .V3 = dist2
                        Else
                            .V1 = dist2
                            .U2 = dist1
                            .V2 = dist4
                            .U3 = dist3
                        End If
                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
                        RebuildTriangleArray

                        If DiagnalTexture = 1 Then
                            Set .Point1 = MakePoint(intX4, 0, intY4)      '     1-2
                            Set .Point2 = MakePoint(intX2, 0, intY2)      '     |/|
                            Set .Point3 = MakePoint(intX3, 0, intY3)      '     4-3
                        Else
                            Set .Point1 = MakePoint(intX2, 0, intY2)
                            Set .Point2 = MakePoint(intX4, 0, intY4)
                            Set .Point3 = MakePoint(intX3, 0, intY3)
                        End If


                        If DiagnalTexture = 1 Then
                            .U1 = dist3
                            .V1 = 0
                            .U2 = 0
                            .V2 = dist2
                            .U3 = 0
                            .V3 = 0
                        Else
                            .V1 = dist2
                            .U2 = dist3
                        End If
                        
                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
               
                    Set m = New Matter
                    With m
                        
                        .TriangleIndex = TriangleCount
                        RebuildTriangleArray
                        
                        If DiagnalTexture = 0 Then
                            Set .Point1 = MakePoint(intX2, 0, intY2)
                            Set .Point2 = MakePoint(intX1, 0, intY1)
                            Set .Point3 = MakePoint(intX4, 0, intY4)
                        Else
                            Set .Point1 = MakePoint(intX2, 0, intY2)
                            Set .Point2 = MakePoint(IIf(Abs(intX1) < Abs(intX2), intX1, intX2), 0, IIf(Abs(intY1) < Abs(intY2), intY1, intY2))
                            Set .Point3 = MakePoint(intX1, 0, intY1)
                        End If
                        

                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)
                        
                        If DiagnalTexture = 0 Then
                            dist1 = DistanceEx(.Point1, .Point2)
                            dist2 = DistanceEx(.Point2, .Point3)
                            dist3 = DistanceEx(.Point3, .Point1)
    
                            .U1 = ((1 / Segments) * 2) * ScaleX
                            .V1 = (VectorSlope(.Point3, .Point1) / 2) * ScaleY
                            .U2 = (-(1 / Segments) * 2) * ScaleX
                            .V2 = (VectorSlope(.Point2, .Point3) / 2) * ScaleY
                            .U3 = 0
                            .V3 = 0
                        End If

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
                        
                        TriangleCount = TriangleCount + 1
                        vol.Add m
                        
                        If DiagnalTexture = 1 Then

                            Dim v As Volume
    
                            dist1 = DistanceEx(.Point2, .Point1)
                            dist2 = DistanceEx(.Point2, .Point3)
    
                            X = (ScaleX * (dist2 / OuterEdge))
                            Y = (ScaleY * (dist1 / OuterEdge))
    
                            .U1 = X
                            .V1 = 0
    
                            .U2 = X
                            .V2 = Y
    
                            .U3 = 0
                            .V3 = Y
    
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
                            
                            If (((intY2 >= 0) And (intX2 > 0)) Or ((intY2 < 0) And (intX2 <= 0))) Then
                                dist1 = DistanceEx(MakePoint(0, 0, intY2), MakePoint(0, 0, intY1))
                                dist2 = DistanceEx(.Point2, MakePoint(0, 0, intY1))
                                Set v = CreateVolumeFace(TextureFileName, .Point1, .Point2, _
                                    MakePoint(0, 0, intY1), MakePoint(0, 0, intY2), _
                                    (ScaleX * (dist1 / OuterEdge)), (ScaleY * (dist2 / OuterEdge)))
                            Else
                                dist1 = DistanceEx(MakePoint(0, 0, intY2), MakePoint(0, 0, intY1))
                                dist2 = DistanceEx(.Point2, MakePoint(0, 0, intY1))
                                Set v = CreateVolumeFace(TextureFileName, .Point3, .Point2, _
                                    MakePoint(0, 0, intY2), MakePoint(0, 0, intY1), _
                                    (ScaleX * (dist1 / OuterEdge)), (ScaleY * (dist2 / OuterEdge)))
                            End If
                                    
                            If Not v Is Nothing Then
                                If v.Count > 0 Then
                                    For Each m In v
                                        vol.Add m
                                    Next
                                    v.Clear
                                End If
                            End If
                        Else
                        
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
                            
                        End If

                    End With
                
                End If
            End If

                
            intX1 = intX2
            intY1 = intY2
            If InnerEdge > 0 Then
                intX4 = intX3
                intY4 = intY3
            End If

        Next

        ObjectCount = ObjectCount + 1

        Set CreateVolumeLanding = vol
    End If
    
End Function


Public Function CreateVolumeLanding2(ByRef TextureFileName As String, ByVal OuterEdge As Single, ByVal Segments As Single, Optional ByVal InnerEdge As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Volume
    If OuterEdge <= 0 Then
        Err.Raise 8, , "OuterEdge must be above 0."
    ElseIf InnerEdge < 0 Then
        Err.Raise 8, , "InnerEdge must be 0 or above."
    ElseIf OuterEdge < InnerEdge Then
        Err.Raise 8, , "OuterEdge must be below InnerEdge."
    ElseIf Segments < 3 Then
        Err.Raise 8, , "Segments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        Dim i As Long
        Dim g As Single
        Dim A As Double
        Dim l1 As Single
        Dim l2 As Single
        Dim l3 As Single
        Dim X As Single
        Dim Y As Single
        
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
        Dim dist5 As Single
        
        Dim vol As New Volume
        Dim m As Matter
        Dim pointsPerFace As Integer
        pointsPerFace = IIf(InnerEdge > 0, 6, 3)
        
        For i = -pointsPerFace To ((pointsPerFace * Segments) - 1) + (pointsPerFace * 2) Step pointsPerFace

            g = (((360 / Segments) * (i / pointsPerFace)) * RADIAN)
            A = (g * DEGREE)
            A = Round(IIf(A < 0, A + 360, IIf(A > 360, A - 360, A)), 0)

            intX2 = (OuterEdge * Sin(A * RADIAN))
            intY2 = (-OuterEdge * Cos(A * RADIAN))
                
            If (InnerEdge > 0) Then
                intX3 = (InnerEdge * Sin(A * RADIAN))
                intY3 = (-InnerEdge * Cos(A * RADIAN))
            End If
            
            If i >= 0 Then
                If (InnerEdge > 0) Then

                    
                    If (i Mod 12) = 0 Then

                        dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) '* IIf(i Mod 4 = 0, 1, -Sin(g)) 'base of trapazoid
                        dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) '* IIf(i Mod 4 = 0, 1, -Cos(g)) 'angled side of trapazoid
                        dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) '* IIf(i Mod 4 = 0, 1, -Sin(g)) 'smaller top edge of trapazoid
                        dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) '* IIf(i Mod 4 = 0, 1, -Cos(g)) 'diagnal inside the trapazoid

                    End If

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        RebuildTriangleArray

                        Set .Point1 = MakePoint(intX4, 0, intY4)     '1-2
                        Set .Point2 = MakePoint(intX2, 0, intY2)     '|/|
                        Set .Point3 = MakePoint(intX1, 0, intY1)     '4-3

                        
                        .U1 = dist3
                        .V1 = 0
                        .U2 = 0
                        .V2 = dist2
                        .U3 = dist1
                        .V3 = dist2
                        
                        
'                        .U1 = dist1
'                        .V1 = 0
'                        .U2 = 0
'                        .V2 = dist4
'                        .U3 = 0
'                        .V3 = 0

'                        .U1 = dist1
'                        .V1 = 0
'                        .U2 = 0
'                        .V2 = dist4
'                        .U3 = 0
'                        .V3 = dist2


                        
                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
                        RebuildTriangleArray

                        Set .Point1 = MakePoint(intX4, 0, intY4)      '     1-2
                        Set .Point2 = MakePoint(intX2, 0, intY2)      '     |/|
                        Set .Point3 = MakePoint(intX3, 0, intY3)    '     4-3


                        .U1 = dist3
                        .V1 = 0
                        .U2 = 0
                        .V2 = dist2
                        .U3 = 0
                        .V3 = 0
                        
                        
'                        .U1 = dist1
'                        .V1 = dist4
'                        .U2 = 0
'                        .V2 = 0
'                        .U3 = dist3
'                        .V3 = 0

'                        .U1 = 0
'                        .V1 = dist1
'                        .U2 = dist4
'                        .V2 = 0
'                        .U3 = 0
'                        .V3 = dist3

'                        .U1 = 0
'                        .V1 = dist1
'                        .U2 = dist2
'                        .V2 = 0
'                        .U3 = 0
'                        .V3 = dist4


                        
                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
               
                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        RebuildTriangleArray

                        Set .Point1 = MakePoint(intX2, 0, intY2)
                        Set .Point2 = MakePoint(intX1, 0, intY1)
                        Set .Point3 = MakePoint(intX4, 0, intY4)


                        dist1 = DistanceEx(.Point1, .Point2)
                        dist2 = DistanceEx(.Point2, .Point3)
                        dist3 = DistanceEx(.Point3, .Point1)

                        .U1 = ((1 / Segments) * 2) * ScaleX
                        .V1 = (VectorSlope(.Point3, .Point1) / 2) * ScaleY
                        .U2 = (-(1 / Segments) * 2) * ScaleX
                        .V2 = (VectorSlope(.Point2, .Point3) / 2) * ScaleY
                        .U3 = 0
                        .V3 = 0

                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
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
            If InnerEdge > 0 Then
                intX4 = intX3
                intY4 = intY3
            End If

        Next

        ObjectCount = ObjectCount + 1

        Set CreateVolumeLanding2 = vol
    End If
    
End Function

'                            Else
'                                Stop
'                            End If

'                            Case 1, 3
'
'
''                            'quadrant two, starts at 90 degrees going clockwise
''                            'point order: bottom-right, top-right, top-left
''
''                            'quadrant four, starts at 270 degrees going clockwise
''                            'point order: toip-right, bottom-right, bottom-left
'
'                            dist1 = DistanceEx(.Point2, .Point1)
'                            dist2 = DistanceEx(.Point2, .Point3)
'
'                            X = (ScaleX * (dist2 / OuterEdge))
'                            Y = (ScaleY * (dist1 / OuterEdge))
'
'                            .U1 = 0
'                            .V1 = X
'
'                            .U2 = Y
'                            .V2 = X
'
'                            .U3 = Y
'                            .V3 = 0
'
'                            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
'                            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
'                            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
'                            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
'                            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
'                            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
'
'                            ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
'                            ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
'                            ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
'                            ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
'                            ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
'                            ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
''                            Debug.Print PointSideOfPlane(.Point1, .Point3, MakePoint((.Point1.X + .Point1.X) / 2, 1, (.Point1.Y + .Point1.Y) / 2), MakePoint(0, 0, intY2))
'
'
'                            dist1 = DistanceEx(.Point2, .Point1)
'                            dist2 = DistanceEx(.Point1, MakePoint(0, 0, intY1))
'
''                            If Not PointSideOfPlane(.Point1, .Point3, MakePoint(((.Point1.X + .Point3.X) / 2), 1, ((.Point1.Y + .Point3.Y) / 2)), MakePoint(0, 0, intY1)) Then
'
'
'                                Set v = CreateVolumeFace(TextureFileName, MakePoint(0, 0, intY2), MakePoint(0, 0, intY1), .Point2, .Point1, _
'                                        (ScaleX * (dist2 / OuterEdge)), (ScaleY * (dist2 / OuterEdge)))
'                                If Not v Is Nothing Then
'                                    If v.Count > 0 Then
'                                        For Each m In v
'                                            vol.Add m
'                                        Next
'                                        v.Clear
'                                    End If
'                                End If
''                            Else
''                             '   Stop
''                            End If
'                        End Select
                        
'                       ' End If
'
'                        End If

               
                'Debug.Print (Round(IIf((g * DEGREE) < 0, (g * DEGREE) + 360, IIf((g * DEGREE) > 360, (g * DEGREE) - 360, (g * DEGREE))), 0) \ 90)
               
                'If ((((i / pointsPerFace) \ pointsPerFace) - 1) \ 2) >= 0 And ((((i / pointsPerFace) \ pointsPerFace) - 1) \ 2) <= 3 Then
               

                    'dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * Sin(g)
                    'dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * Cos(g)
                    
'                    Set m = New Matter
'                    With m
'                        .TriangleIndex = TriangleCount
'                        RebuildTriangleArray
'
'                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
'                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
'                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
'
'                        Set .Point1 = Points(.Index1)
'                        Set .Point2 = Points(.Index2)
'                        Set .Point3 = Points(.Index3)
'
'                        dist1 = DistanceEx(.Point1, .Point2)
'                        dist2 = DistanceEx(.Point2, .Point3)
'                        dist3 = DistanceEx(.Point3, .Point1)
'
''                        A = (g * DEGREE)
''                        A = Round(IIf(A < 0, A + 360, IIf(A > 360, A - 360, A)), 0)
'
''                        .U1 = (1 / Segments) * 2
''                        .V1 = VectorSlope(.Point3, .Point1) / 2
''                        .U2 = -(1 / Segments) * 2
''                        .V2 = VectorSlope(.Point2, .Point3) / 2
''                        .U3 = 0
''                        .V3 = 0
'
'
'                        .U1 = ((1 / Segments) * (ScaleX / 100)) '+ .Point1.X
'                        .V1 = 1 / VectorSlope(.Point3, .Point1) ' /2 '(ScaleY / 100) '+ .Point1.Y
'                        .U2 = ((1 / Segments) * (ScaleY / 100)) '+ .Point2.X
'                        .V2 = 1 / VectorSlope(.Point2, .Point3) '/2' * 100 ' (ScaleY / 100) '+ .Point2.Y
'                        .U3 = 0 ' ((.Point1.X - .Point2.X) * (ScaleX / 100))
'                        .V3 = 0 '((.Point1.Y - .Point2.Y) * (ScaleY / 100))
'
'
''                        X = (dist1 / OuterEdge)
''                        Y = (dist1 / OuterEdge)
''
''                        .U1 = .U1 + X
''                        .V1 = .V1 + Y
''
''                        .U2 = .U2 - Y
''                        .V2 = .V2 - X
''
''                        .U3 = .U3 + Y
''                        .V3 = .V3 - Y
'
'                        Select Case A \ 90 'quadrant clockwize
'                            Case 0
'
'                            Case 1
'
'                            Case 2
'
'                            Case 3
'
'                        End Select
'
''                        .V1 = ((.Point3.X / (OuterEdge * 2)) * (ScaleX / 100))
''                        .U1 = ((.Point2.Y / (OuterEdge * 2)) * (ScaleY / 100))
''
''                        .V3 = ((.Point3.Y / (OuterEdge * 2)) * (ScaleY / 100))
''                        .U3 = ((.Point2.X / (OuterEdge * 2)) * (ScaleX / 100))
''
''                        .V2 = ((.Point1.X / (OuterEdge * 2)) * (ScaleX / 100))
''                        .U2 = ((.Point3.Y / (OuterEdge * 2)) * (ScaleY / 100))
'
''                        .U1 = .U1 - ((OuterEdge - (.Point3.Y)) / (ScaleY / 100))
''                        .V1 = .V1 - ((OuterEdge - (.Point3.X)) / (ScaleX / 100))
''
''                        .U2 = .U2 - ((OuterEdge - (.Point2.Y)) / (ScaleY / 100))
''                        .V2 = .V2 - ((OuterEdge - (.Point2.X)) / (ScaleX / 100))
''
''                        .U3 = .U3 - ((OuterEdge - (.Point1.Y)) / (ScaleY / 100))
''                        .V3 = .V3 - ((OuterEdge - (.Point1.X)) / (ScaleX / 100))
''
''
''                        .U1 = .U1 - ((OuterEdge - (.Point1.X)) / (ScaleX / 100))
''                        .V1 = .V1 - ((OuterEdge - (.Point1.Y)) / (ScaleY / 100))
''
''                        .U2 = .U2 - ((OuterEdge - (.Point2.X)) / (ScaleX / 100))
''                        .V2 = .V2 - ((OuterEdge - (.Point2.Y)) / (ScaleY / 100))
''
''                        .U3 = .U3 - ((OuterEdge - (.Point3.X)) / (ScaleX / 100))
''                        .V3 = .V3 - ((OuterEdge - (.Point3.Y)) / (ScaleY / 100))
''
''
''
''                        .U1 = .U1 - (((OuterEdge * 2) - ((.Point1.X + .Point2.X + .Point3.X) / 3)) / ((ScaleX / 100) / 2))
''                        .V1 = .V1 - (((OuterEdge * 2) - ((.Point1.Y + .Point2.Y + .Point3.Y) / 3)) / ((ScaleY / 100) / 2))
''
''                        .U2 = .U2 + (((OuterEdge * 2) - ((.Point1.X + .Point2.X + .Point3.X) / 3)) / ((ScaleX / 100) / 2))
''                        .V2 = .V2 + (((OuterEdge * 2) - ((.Point1.Y + .Point2.Y + .Point3.Y) / 3)) / ((ScaleY / 100) / 2))
''
''                        .U3 = .U3 - (((OuterEdge * 2) - ((.Point1.X + .Point2.X + .Point3.X) / 3)) / ((ScaleX / 100) / 2))
''                        .V3 = .V3 - (((OuterEdge * 2) - ((.Point1.Y + .Point2.Y + .Point3.Y) / 3)) / ((ScaleY / 100) / 2))
'
'
'                        Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)
'
'                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
'                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
'                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
'
'                        VertexXAxis(0, TriangleCount) = .Point1.X
'                        VertexXAxis(1, TriangleCount) = .Point2.X
'                        VertexXAxis(2, TriangleCount) = .Point3.X
'
'                        VertexYAxis(0, TriangleCount) = .Point1.Y
'                        VertexYAxis(1, TriangleCount) = .Point2.Y
'                        VertexYAxis(2, TriangleCount) = .Point3.Y
'
'                        VertexZAxis(0, TriangleCount) = .Point1.z
'                        VertexZAxis(1, TriangleCount) = .Point2.z
'                        VertexZAxis(2, TriangleCount) = .Point3.z
'
'                        TriangleFace(0, TriangleCount) = .Normal.X
'                        TriangleFace(1, TriangleCount) = .Normal.Y
'                        TriangleFace(2, TriangleCount) = .Normal.z
'                        TriangleFace(4, TriangleCount) = ObjectCount
'                        TriangleFace(5, TriangleCount) = ((pointsPerFace + i) \ pointsPerFace)
'
'                        .ObjectIndex = ObjectCount
'                        .FaceIndex = ((pointsPerFace + i) \ pointsPerFace)
'                        .TextureIndex = GetFileIndex(TextureFileName)
'                        If TextureFileName <> "" Then
'                            If Files(.TextureIndex).Data Is Nothing Then
'                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
'                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
'                            End If
'                        End If
'
'                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
'                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
'                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
'                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
'                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)
'
'                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
'                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
'                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
'                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
'                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)
'
'                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
'                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
'                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
'                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
'                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
'
'                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
'                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
'                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
'                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
'                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
'                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
'                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
'                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
'                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
'                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
'                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
'                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
'                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
'
'                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
'                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
'                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
'                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
'                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
'                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
'
'                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
'                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
'                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
'                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
'                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
'                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
'                    End With
'
'                    TriangleCount = TriangleCount + 1
'                    vol.Add m
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
                RebuildTriangleArray

                Set .Point1 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 0)).X, _
                    MeshVerticies(MeshIndicies(Index + 0)).Y, _
                    MeshVerticies(MeshIndicies(Index + 0)).z)

                Set .Point2 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 1)).X, _
                    MeshVerticies(MeshIndicies(Index + 1)).Y, _
                    MeshVerticies(MeshIndicies(Index + 1)).z)

                Set .Point3 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 2)).X, _
                    MeshVerticies(MeshIndicies(Index + 2)).Y, _
                    MeshVerticies(MeshIndicies(Index + 2)).z)


                Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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
                RebuildTriangleArray

                Set .Point1 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 3)).X, _
                    MeshVerticies(MeshIndicies(Index + 3)).Y, _
                    MeshVerticies(MeshIndicies(Index + 3)).z)

                Set .Point2 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 4)).X, _
                    MeshVerticies(MeshIndicies(Index + 4)).Y, _
                    MeshVerticies(MeshIndicies(Index + 4)).z)

                Set .Point3 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 5)).X, _
                    MeshVerticies(MeshIndicies(Index + 5)).Y, _
                    MeshVerticies(MeshIndicies(Index + 5)).z)


                Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

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


