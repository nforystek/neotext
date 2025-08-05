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

Public Sub RenderObjs(ByRef UserControl As Macroscopic, ByRef Camera As Camera)



End Sub



'' Return the dot product AB · BC.
'' Note that AB · BC = |AB| * |BC| * Cos(theta).
'Private Function DotProduct( _
'    ByVal Ax As Single, ByVal Ay As Single, _
'    ByVal Bx As Single, ByVal By As Single, _
'    ByVal cx As Single, ByVal cy As Single _
'  ) As Single
'    Dim BAx As Single
'    Dim BAy As Single
'    Dim BCx As Single
'    Dim BCy As Single
'
'    ' Get the vectors' coordinates.
'    BAx = Ax - Bx
'    BAy = Ay - By
'    BCx = cx - Bx
'    BCy = cy - By
'
'    ' Calculate the dot product.
'    DotProduct = BAx * BCx + BAy * BCy
'End Function
'
'
'Public Function CrossProductLength( _
'    ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, _
'    ByVal Bx As Single, ByVal By As Single, ByVal Bz As Single, _
'    ByVal cx As Single, ByVal cy As Single, ByVal cz As Single _
'  ) As Single
'    Dim BAx As Single
'    Dim BAy As Single
'    Dim BAz As Single
'    Dim BCx As Single
'    Dim BCy As Single
'    Dim BCz As Single
'
'    ' Get the vectors' coordinates.
'    BAx = Ax - Bx
'    BAy = Ay - By
'    BAz = Az - Bz
'    BCx = cx - Bx
'    BCy = cy - By
'    BCz = cz - Bz
'
'    ' Calculate the Z coordinate of the cross product.
'    CrossProductLength = BAx * BCy - BAy * BCz - BAz * BCx
'End Function
'
''' Return the angle ABC.
''' Return a value between PI and -PI.
''' Note that the value is the opposite of what you might
''' expect because Y coordinates increase downward.
'Public Function GetAngle(ByRef p1 As Point, ByRef p2 As Point) As Single
''ByVal Ax As Single, ByVal Ay As _
'    'Single, ByVal Bx As Single, ByVal By As Single, ByVal _
'    'Cx As Single, ByVal Cy As Single) As Single
'    Dim dot_product As Single
'    Dim cross_product As Single
'
'    ' Get the dot product and cross product.
'    dot_product = VectorDotProduct(p1, p2) 'dotproduct(p1.x, p1.y, 0, 0, p3.x, p3.y)
'    cross_product = DistanceEx(MakePoint(0, 0, 0), VectorCrossProduct(p1, p2)) 'CrossProductLength(p1.x, p1.y, p1.z, 0, 0, 0, p2.x, p2.y, p2.z) 'CrossProductLength(p1.x, p1.y, 0, 0, p3.x, p3.y)
'
'    ' Calculate the angle.
'    GetAngle = ATan2(CDbl(cross_product), CDbl(dot_product)) * DEGREE
'End Function
'
'
'Public Function GetAngle2(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
'    Dim XDiff As Double
'    Dim YDiff As Double
'    Dim TempAngle As Double
'
'    YDiff = Abs(y2 - y1)
'
'    If x1 = x2 And y1 = y2 Then Exit Function
'
'    If YDiff = 0 And x1 < x2 Then
'        GetAngle2 = 0
'        Exit Function
'    ElseIf YDiff = 0 And x1 > x2 Then
'        GetAngle2 = 3.14159265358979
'        Exit Function
'    End If
'
'    XDiff = Abs(x2 - x1)
'
'    TempAngle = Atn(XDiff / YDiff)
'
'    If y2 > y1 Then TempAngle = 3.14159265358979 - TempAngle
'    If x2 < x1 Then TempAngle = -TempAngle
'    TempAngle = 1.5707963267949 - TempAngle
'    If TempAngle < 0 Then TempAngle = 6.28318530717959 + TempAngle
'
'    GetAngle2 = TempAngle
'End Function
'
'
'Public Function GetAngle3(ByRef p1 As Point, ByRef p2 As Point) As Single
'If p1.X = p2.X Then
'    If p1.Y < p2.Y Then
'        GetAngle3 = 90
'    Else
'        GetAngle3 = 270
'    End If
'    Exit Function
'ElseIf p1.Y = p2.Y Then
'    If p1.X < p2.X Then
'        GetAngle3 = 0
'    Else
'        GetAngle3 = 180
'    End If
'    Exit Function
'Else
'    GetAngle3 = Atn(VectorSlope(p1, p2))
'    GetAngle3 = GetAngle3 * 180 / PI
'    If GetAngle3 < 0 Then GetAngle3 = GetAngle3 + 360
'    '----------Test for direction--------
'    If p1.X > p2.X And GetAngle3 <> 180 Then GetAngle3 = GetAngle3 + 180
'    If p1.Y > p2.Y And GetAngle3 = 90 Then GetAngle3 = GetAngle3 + 180
'    If GetAngle3 > 360 Then GetAngle3 = GetAngle3 - 360
'End If
'End Function
'
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


'Public Function Tangent(p_dblVal As Single) As Single
'
'    ' Comments :
'    ' Parameters: p_dblVal -
'    ' Returns: Double -
'    ' Modified :
'    '
'    ' -------------------------
'    'Degree Input Radian Output
'    On Error GoTo PROC_ERR
'    Dim dblPi As Single
'    Dim dblRadian As Single
'    ' xx Calculate the value of Pi.
'    dblPi = 4 * Atn(1)
'    ' xx To convert degrees to radians,
'    'multiply degrees by Pi / 180.
'    dblRadian = dblPi / 180
'
'    p_dblVal = Val(p_dblVal * dblRadian)
'    Tangent = Tan(p_dblVal)
'PROC_EXIT:
'    Exit Function
'PROC_ERR:
'    Tangent = 0
'    'MsgBox Err.Description, vbExclamation
'    Resume PROC_EXIT
'End Function
'
'
'Public Function ArcSine(p_dblVal As Single) As Single
'
'    ' Comments :
'    ' Parameters: p_dblVal -
'    ' Returns: Double -
'    ' Modified :
'    '
'    ' -------------------------
'    'Radian Input Degree Output
'    On Error GoTo PROC_ERR
'    Dim dblSqr As Single
'    Dim dblPi As Single
'    Dim dblDegree As Single
'    ' xx Calculate the value of Pi.
'    dblPi = 4 * Atn(1)
'    ' xx To convert radians to degrees,
'    ' multiply radians by 180/pi.
'    dblDegree = 180 / dblPi
'    p_dblVal = Val(p_dblVal)
'    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
'    ' xx Prevent division by Zero error
'
'    If dblSqr = 0 Then
'        dblSqr = 1E-30
'    End If
'
'    ArcSine = Atn(p_dblVal / dblSqr) * dblDegree
'PROC_EXIT:
'    Exit Function
'PROC_ERR:
'    ArcSine = 0
'    'MsgBox Err.Description, vbExclamation
'    Resume PROC_EXIT
'End Function
'
'
'Public Function ArcCosine(p_dblVal As Single) As Single
'
'    ' Comments :
'    ' Parameters: p_dblVal -
'    ' Returns: Double -
'    ' Modified :
'    '
'    ' -------------------------
'    'Radian Input Degree Output
'    On Error GoTo PROC_ERR
'    Dim dblSqr As Single
'    Dim dblPi As Single
'    Dim dblDegree As Single
'    ' xx Calculate the value of Pi.
'    dblPi = 4 * Atn(1)
'    ' xx To convert radians to degrees,
'    ' multiply radians by 180/pi.
'    dblDegree = 180 / dblPi
'    p_dblVal = Val(p_dblVal)
'    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
'    ' xx Prevent division by Zero error
'
'    If dblSqr = 0 Then
'        dblSqr = 1E-30
'    End If
'
'    ArcCosine = (Atn(-p_dblVal / dblSqr) + 2 * Atn(1)) * dblDegree
'PROC_EXIT:
'    Exit Function
'PROC_ERR:
'    ArcCosine = 0
'    'MsgBox Err.Description, vbExclamation
'    Resume PROC_EXIT
'End Function
'
'
'Public Function ArcTangent(p_dblVal As Single) As Single
'
'    ' Comments :
'    ' Parameters: p_dblVal -
'    ' Returns: Double -
'    ' Modified :
'    '
'    ' -------------------------
'    'Radian Input Degree Output
'    On Error GoTo PROC_ERR
'    Dim dblPi As Single
'    Dim dblDegree As Single
'    ' xx Calculate the value of Pi.
'    dblPi = 4 * Atn(1)
'    ' xx To convert radians to degrees,
'    ' multiply radians by 180/pi.
'    dblDegree = 180 / dblPi
'    p_dblVal = Val(p_dblVal)
'    ArcTangent = Atn(p_dblVal) * dblDegree
'PROC_EXIT:
'    Exit Function
'PROC_ERR:
'    ArcTangent = 0
'    'MsgBox Err.Description, vbExclamation
'    Resume PROC_EXIT
'End Function

Private Function CombineOrbits(ByRef o1 As Orbit, ByRef o2 As Orbit) As Orbit
    Set CombineOrbits = New Orbit
    With CombineOrbits
        .Origin = VectorAddition(o1.Origin, o2.Origin)
        .Offset = VectorDeduction(VectorDeduction(VectorAddition(o1.Origin, o1.Offset), VectorAddition(o2.Origin, o2.Offset)), .Origin)
        .Rotate = VectorAddition(o1.Rotate, o2.Rotate)
        .Scaled = VectorAddition(o1.Scaled, o2.Scaled)
        .Ranges.X = o1.Ranges.X + o2.Ranges.X
        .Ranges.Y = o1.Ranges.Y + o2.Ranges.Y
        .Ranges.Z = o1.Ranges.Z + o2.Ranges.Z
        If o1.Ranges.r = -1 Or o2.Ranges.r = -1 Then
            .Ranges.r = -1
        ElseIf o1.Ranges.r > o2.Ranges.r Then
            .Ranges.r = o1.Ranges.r
        Else
            .Ranges.r = o2.Ranges.r
        End If
    End With
End Function


Public Sub RenderMolecules(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
    Dim p As Planet
    Dim p2 As Planet
    Dim m As Molecule
    
    'called once per frame drawing the objects, with out any of the current frame object
    'properties modifying calls included for latent collision checking rollback


    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

'    DDevice.SetRenderState D3DRS_LIGHTING, 1
'
'    DDevice.SetRenderState D3DRS_CLIPPING, 1
    
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetRenderState D3DRS_ZENABLE, 0
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 0

'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetMaterial LucentMaterial
    DDevice.SetTexture 0, Nothing
    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 1, Nothing

    Dim dist As Single
    Dim dist2 As Single
    
    Dim matRoll As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPos As D3DMATRIX

    
    Dim cnt As Long
'
'    If All.Count > 0 Then
'        cnt = 1
'        Do
'
'            Debug.Print All(cnt).Key;
'            cnt = cnt + 1
'        Loop While cnt <= All.Count
'        Debug.Print
'    End If

    
'    RenderOrbits Molecules, True
'    For Each m In Molecules
'        If m.Parent Is Nothing Then
'            AllCommitRoutine m, Nothing
'        End If
'    Next
''
''    RenderOrbits Planets, False
'
'    For Each p In Planets
'        RenderOrbits p.Molecules, False
'    Next
    
    
'    If Not Camera.Player Is Nothing Then
'
'        D3DXMatrixIdentity matMat
'
'        D3DXMatrixTranslation matPos, -Camera.Player.Absolute.Origin.X, -Camera.Player.Absolute.Origin.Y, -Camera.Player.Absolute.Origin.Z
'        D3DXMatrixMultiply matMat, matPos, matMat
'
'        D3DXMatrixRotationX matPitch, AngleConvertWinToDX3DX(AngleRestrict(-Camera.Player.Absolute.Rotate.Z))
'        D3DXMatrixMultiply matMat, matPitch, matMat
'
'        D3DXMatrixRotationY matYaw, AngleConvertWinToDX3DY(AngleRestrict(-Camera.Player.Absolute.Rotate.X))
'        D3DXMatrixMultiply matMat, matYaw, matMat
'
'        D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3DZ(AngleRestrict(-Camera.Player.Absolute.Rotate.Y))
'        D3DXMatrixMultiply matMat, matRoll, matMat
'
'   End If

    Dim matMat As D3DMATRIX
    D3DXMatrixIdentity matMat


    
    RenderOrbits Molecules, True


    If Planets.Count > 0 Then
        cnt = 0
        Do

            cnt = cnt + 1

            Set p = Planets(cnt)

            If Not Camera.Player Is Nothing Then
                If dist = 0 Then
                    dist = DistanceEx(Planets(cnt).Absolute.Origin, Camera.Player.Absolute.Origin)
                End If
                If cnt < Planets.Count Then

                    dist2 = DistanceEx(Planets(cnt + 1).Absolute.Origin, Camera.Player.Absolute.Origin)
                    If dist2 > dist Then
                        Set p = Planets(cnt + 1)
                        Planets.Remove cnt + 1
                        Planets.Add p, p.Key, cnt
                    Else
                        dist = dist2
                    End If
                End If
            End If


            RenderMolecule p, Nothing, matMat

            RenderOrbits p.Molecules, False


            Set p = Nothing

        Loop Until cnt >= Planets.Count
    End If
    'Debug.Print
   
End Sub

Private Sub RenderOrbits(ByRef col As Object, ByVal NoParentOnly As Boolean, Optional ByVal NoChildren As Boolean)
    
    Dim matMat As D3DMATRIX
    D3DXMatrixIdentity matMat
    
'    If Not col Is Nothing Then
'
'        Dim m As Molecule
'        For Each m In col
'
'            If NoParentOnly Then
'                If m.Parent Is Nothing Then
'                    RenderMolecule m, Nothing, matMat
'                End If
'            Else
'                RenderMolecule m, Nothing, matMat
'            End If
'
'        Next
'    End If


    If Not col Is Nothing Then

        Dim m As Molecule
        Dim m2 As Molecule

        Dim cnt As Long
        Dim cnt2 As Long

        Dim dist As Single
        Static dist2 As Single
        Dim parenttrap As Boolean

        If col.Count > 0 Then
            cnt = 0
            Do

                cnt = cnt + 1

                Set m = col(cnt)

                If Not Camera.Player Is Nothing Then
                    If dist = 0 Then
                        If col(cnt).Parent Is Nothing Then
                            dist = DistanceEx(col(cnt).Absolute.Origin, Camera.Player.Absolute.Origin)
                        Else
                            dist = DistanceEx(VectorAddition(col(cnt).Absolute.Origin, col(cnt).Parent.Absolute.Origin), Camera.Player.Absolute.Origin)
                        End If
                    End If
                    If cnt < col.Count Then

                        If col(cnt + 1).Parent Is Nothing Then
                            dist2 = DistanceEx(col(cnt + 1).Absolute.Origin, Camera.Player.Absolute.Origin)
                        Else
                            dist2 = DistanceEx(VectorAddition(col(cnt + 1).Absolute.Origin, col(cnt + 1).Parent.Absolute.Origin), Camera.Player.Absolute.Origin)
                        End If
                        If (dist2 > dist) Then
                            Set m = col(cnt + 1)
                            col.Remove cnt + 1
                            col.Add m, m.Key, cnt
                            Set m = col(cnt)
                        Else
                            dist = dist2
                        End If
                    End If
                End If

                If NoParentOnly Then
                    If m.Parent Is Nothing Then
                        RenderMolecule m, Nothing, matMat, NoChildren
                    End If
                Else
                    RenderMolecule m, Nothing, matMat, NoChildren
                End If

                Set m = Nothing

            Loop Until cnt >= col.Count
        End If


    End If

End Sub

Private Sub RenderMolecule(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByRef matMat As D3DMATRIX, Optional ByVal NoChildren As Boolean = False)
    
    Dim vout1 As D3DVECTOR
    Dim vout2 As D3DVECTOR
    Dim vout3 As D3DVECTOR
    

    Dim matPos As D3DMATRIX
    Dim matRoll As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matScale As D3DMATRIX


    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matRoll
    D3DXMatrixIdentity matYaw
    D3DXMatrixIdentity matPitch
    D3DXMatrixIdentity matScale

    D3DXMatrixTranslation matPos, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.Z
    D3DXMatrixMultiply matMat, matPos, matMat

    D3DXMatrixRotationX matPitch, AngleConvertWinToDX3DX(ApplyTo.Rotate.X)
    D3DXMatrixMultiply matMat, matPitch, matMat

    D3DXMatrixRotationY matYaw, AngleConvertWinToDX3DY(ApplyTo.Rotate.Y)
    D3DXMatrixMultiply matMat, matYaw, matMat

    D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3DZ(ApplyTo.Rotate.Z)
    D3DXMatrixMultiply matMat, matRoll, matMat

    D3DXMatrixTranslation matPos, ApplyTo.Offset.X, ApplyTo.Offset.Y, ApplyTo.Offset.Z
    D3DXMatrixMultiply matMat, matPos, matMat

    D3DXMatrixScaling matScale, ApplyTo.Scaled.X, ApplyTo.Scaled.Y, ApplyTo.Scaled.Z
    D3DXMatrixMultiply matScale, matScale, matMat
    
    
    If Not Parent Is Nothing Then
        ApplyTo.Moved = ApplyTo.Moved Or Parent.Moved

    End If
    
    Dim v As Matter
    
    If Not ApplyTo.Volume Is Nothing Then
        For Each v In ApplyTo.Volume
    
            
            
            If ApplyTo.Moved Then
            
                D3DXVec3TransformCoord vout1, ToVector(v.Point1), matScale
                VertexDirectX((v.TriangleIndex * 3) + 0).X = vout1.X
                VertexDirectX((v.TriangleIndex * 3) + 0).Y = vout1.Y
                VertexDirectX((v.TriangleIndex * 3) + 0).Z = vout1.Z
        
                D3DXVec3TransformCoord vout2, ToVector(v.Point2), matScale
                VertexDirectX((v.TriangleIndex * 3) + 1).X = vout2.X
                VertexDirectX((v.TriangleIndex * 3) + 1).Y = vout2.Y
                VertexDirectX((v.TriangleIndex * 3) + 1).Z = vout2.Z
        
                D3DXVec3TransformCoord vout3, ToVector(v.Point3), matScale
                VertexDirectX((v.TriangleIndex * 3) + 2).X = vout3.X
                VertexDirectX((v.TriangleIndex * 3) + 2).Y = vout3.Y
                VertexDirectX((v.TriangleIndex * 3) + 2).Z = vout3.Z
                
                Set v.Normal = TriangleNormal(v.Point1, v.Point2, v.Point3)

                VertexDirectX(v.TriangleIndex * 3 + 0).NX = v.Normal.X
                VertexDirectX(v.TriangleIndex * 3 + 0).NY = v.Normal.Y
                VertexDirectX(v.TriangleIndex * 3 + 0).Nz = v.Normal.Z

                VertexDirectX(v.TriangleIndex * 3 + 1).NX = v.Normal.X
                VertexDirectX(v.TriangleIndex * 3 + 1).NY = v.Normal.Y
                VertexDirectX(v.TriangleIndex * 3 + 1).Nz = v.Normal.Z

                VertexDirectX(v.TriangleIndex * 3 + 2).NX = v.Normal.X
                VertexDirectX(v.TriangleIndex * 3 + 2).NY = v.Normal.Y
                VertexDirectX(v.TriangleIndex * 3 + 2).Nz = v.Normal.Z
        
                VertexXAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).X
                VertexXAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).X
                VertexXAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).X
        
                VertexYAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).Y
                VertexYAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).Y
                VertexYAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).Y
        
                VertexZAxis(0, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 0).Z
                VertexZAxis(1, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 1).Z
                VertexZAxis(2, v.TriangleIndex) = VertexDirectX(v.TriangleIndex * 3 + 2).Z


            End If
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
    
   ' Debug.Print ApplyTo.Key;
    If Not NoChildren Then
        Dim m As Molecule
        If Not ApplyTo.Molecules Is Nothing Then
            For Each m In ApplyTo.Molecules
                 RenderMolecule m, ApplyTo, matMat
            Next
        End If
    End If

    
    If ApplyTo.Moved Then ApplyTo.Moved = False
    

    D3DXMatrixTranslation matPos, -ApplyTo.Offset.X, -ApplyTo.Offset.Y, -ApplyTo.Offset.Z
    D3DXMatrixMultiply matMat, matPos, matMat

    D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3DZ(-ApplyTo.Rotate.Z)
    D3DXMatrixMultiply matMat, matRoll, matMat

    D3DXMatrixRotationY matYaw, AngleConvertWinToDX3DY(-ApplyTo.Rotate.Y)
    D3DXMatrixMultiply matMat, matYaw, matMat

    D3DXMatrixRotationX matPitch, AngleConvertWinToDX3DX(-ApplyTo.Rotate.X)
    D3DXMatrixMultiply matMat, matPitch, matMat

    D3DXMatrixTranslation matPos, -ApplyTo.Origin.X, -ApplyTo.Origin.Y, -ApplyTo.Origin.Z
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

            VertexZAxis(0, TriangleCount) = .Point1.Z
            VertexZAxis(1, TriangleCount) = .Point2.Z
            VertexZAxis(2, TriangleCount) = .Point3.Z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.Z
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
            VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
            ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
            ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
            ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
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

            VertexZAxis(0, TriangleCount) = .Point1.Z
            VertexZAxis(1, TriangleCount) = .Point2.Z
            VertexZAxis(2, TriangleCount) = .Point3.Z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.Z
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
            VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
            ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
            ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
            ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
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

                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)

                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X

                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

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

                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)

                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X

                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

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

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z
                        
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

                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)

                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X

                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

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

                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)

                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X

                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

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

                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
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
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Z
                        ScreenDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Z
                        ScreenDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Z
                        ScreenDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z

                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z

                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z

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
        Set Mesh = D3DX.LoadMeshFromX(DirectXFileName, D3DXMESH_POINTS, DDevice, Nothing, nMatBuffer, nMaterials)

        Dim Index As Long
        Dim cnt As Long

        Dim VD As D3DVERTEXBUFFER_DESC
        Mesh.GetVertexBuffer.GetDesc VD
'        VD.FVF = FVF_RENDER
'        VD.Size = FVF_RENDER_SIZE
'        VD.Type = CONST_D3DRESOURCETYPE.D3DRTYPE_INDEXBUFFER
'        VD.Format = D3DFMT_VERTEXDATA
'        VD.Pool = D3DPOOL_DEFAULT
        
        ReDim MeshVerticies(0 To 0) As D3DVERTEX
        ReDim MeshVerticies(0 To ((VD.Size \ Len(MeshVerticies(0))) - 1)) As D3DVERTEX
        D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, VD.Size, 0, MeshVerticies(0)

        Dim ID As D3DINDEXBUFFER_DESC
        Mesh.GetIndexBuffer.GetDesc ID
'        ID.Type = D3DRTYPE_INDEXBUFFER
'        ID.Format = -D3DFMT_VERTEXDATA
'        ID.Pool = D3DPOOL_DEFAULT
        
        ReDim MeshIndicies(0 To 0) As Integer
        ReDim MeshIndicies(0 To ((ID.Size \ Len(MeshIndicies(0))) - 1)) As Integer
        D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, ID.Size, 0, MeshIndicies(0)



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
                    MeshVerticies(MeshIndicies(Index + 0)).Z)

                Set .Point2 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 1)).X, _
                    MeshVerticies(MeshIndicies(Index + 1)).Y, _
                    MeshVerticies(MeshIndicies(Index + 1)).Z)

                Set .Point3 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 2)).X, _
                    MeshVerticies(MeshIndicies(Index + 2)).Y, _
                    MeshVerticies(MeshIndicies(Index + 2)).Z)


                Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.Z
                VertexZAxis(1, TriangleCount) = .Point2.Z
                VertexZAxis(2, TriangleCount) = .Point3.Z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.Z
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
                VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 0)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 0)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 1)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 1)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z
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
                    MeshVerticies(MeshIndicies(Index + 3)).Z)

                Set .Point2 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 4)).X, _
                    MeshVerticies(MeshIndicies(Index + 4)).Y, _
                    MeshVerticies(MeshIndicies(Index + 4)).Z)

                Set .Point3 = MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 5)).X, _
                    MeshVerticies(MeshIndicies(Index + 5)).Y, _
                    MeshVerticies(MeshIndicies(Index + 5)).Z)


                Set .Normal = PlaneNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.Z
                VertexZAxis(1, TriangleCount) = .Point2.Z
                VertexZAxis(2, TriangleCount) = .Point3.Z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.Z
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
                VertexDirectX(.TriangleIndex * 3 + 0).Z = .Point1.Z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.Z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 3)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 3)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Z = .Point2.Z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.Z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 4)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 4)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Z = .Point3.Z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.Z
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


