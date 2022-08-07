Attribute VB_Name = "modLand"
#Const modLand = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

'############################################################################################################
'Derived Exports ############################################################################################
'############################################################################################################
            
'MaxLandLib.dll exports
'extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3);
'Accepts inputs n1 and n2 as retruned from PointInPoly(X,Y) then again for (Z,Y) and n2 as returned from tri_tri_intersect() to return the determination of whether or not the collision is correct and satisfy bitwise and math equalaterally collision precise to real coordination from the preliminary possible collision information the other functions return.
'extern short tri_tri_intersect (unsigned short v0_0, unsigned short v0_1, unsigned short v0_2, unsigned short v1_0, unsigned short v1_1, unsigned short v1_2, unsigned short v2_0, unsigned short v2_1, unsigned short v2_2, unsigned short u0_0, unsigned short u0_1, unsigned short u0_2, unsigned short u1_0, unsigned short u1_1, unsigned short u1_2, unsigned short u2_0, unsigned short u2_1, unsigned short u2_2);
'Accepts two triangle inputs in hyperbolic paraboloid collision form and returns with in the unsiged whole the percentage of each others distance to plane as one value.  **NOTE Assumes the parameter input as triangles are TRUE for collision with one another.
'extern int Forystek (int visType, int lngFaceCount, unsigned short *sngCamera[], unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], unsigned short *sngScreenX[], unsigned short *sngScreenY[], unsigned short *sngScreenZ[], unsigned short *sngZBuffer[]);
'Culling function with three expirimental ways to cull, defined by visType, 0 to 2, returns the difference of input triangles. lngFaceCount, sngCamera[3 x 3], sngFaceVis[6 x lngFaceCount], sngVertexX[3 x lngFaceCount]..Y..Z, sngScreenX[3 x lngFaceCount]..Y..Z, sngZBuffer[4 x lngFaceCount].  The camera is defined by position [0,0]=X, [0,1]=Y, [0,2]=Z, direction [1,0]=X, [1,1]=Y, [1,2]=Z, and upvector [2,0]=X, [2,1]=Y, [2,2]=Z.  sngFaceVis should be initialized to zero, and sngVertex arrays are 3D coordinate equivelent to sngScreen with a screenZ buffer, and Zbuffer for the verticies.
'extern bool PointBehindPoly (unsigned short pX, unsigned short pY, unsigned short pZ, unsigned short nX, unsigned short nY, unsigned short nZ, unsigned short vX, unsigned short vY, unsigned short vZ) ;
'Checks for the presence of a point behind a triangle, the first three inputs are the length of the triangles sides, the next three are the triangles normal, the last three are the point to test with the triangles center removed.
'extern int PointInPoly (int pX, int pY, unsigned short *polyX[], unsigned short *polyY[], int polyN);
'Tests for the presence of a 2D point pX,pY anywhere within a 2D shape defined with a list of points polyX,polyY that has polyN number of coordinates, returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries.
'extern bool Collision (int visType, int lngFaceCount, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace);
'Tests collision of a lngFaceNum against a number of visible faces, lngFaceCount, whose sngFaceVis has been defined with visType as culled with the Forystek function, and returns whether or not a collision occurs also populating the lngCollidedBrush and lngCollidedFace indicating the exact object number (brush) and face number (triangle) that has the collision impact.

Private Declare Function Collision Lib "MaxLandLib" (ByVal visType As _
                                    Long, ByVal lngFaceCount As _
                                    Long, sngFaceVis() As _
                                    Single, sngVertexX() As _
                                    Single, sngVertexY() As _
                                    Single, sngVertexZ() As _
                                    Single, ByVal lngFaceNum As _
                                    Long, ByRef lngCollidedBrush As _
                                    Long, ByRef lngCollidedFace As Long) As Boolean

'############################################################################################################
'Variable Declare ###########################################################################################
'############################################################################################################

Private lngFaceCount As Long
Private sngFaceVis() As Single
'sngFaceVis dimension (,n) where n=# is face number
'sngFaceVis dimension (n,) where n=0 is x of face normal
'sngFaceVis dimension (n,) where n=1 is y of face normal
'sngFaceVis dimension (n,) where n=2 is z of face normal
'sngFaceVis dimension (n,) where n=3 is vis Type, values
'sngFaceVis dimension (n,) where n=4 is gBrush index
'sngFaceVis dimension (n,) where n=4 is gFace index

Private sngVertexX() As Single
Private sngVertexY() As Single
Private sngVertexZ() As Single
'sngVertexX dimension (,n) where n=# is face number
'sngVertexX dimension (n,) where n=0 is faces first vertex.X
'sngVertexX dimension (n,) where n=1 is faces second vertex.X
'sngVertexX dimension (n,) where n=2 is faces third vertex.X
'sngVertexX dimension (n,) where n=3 is faces fourth vertex.X

Private Sunrotated As Single

Private SkyPlaq(0 To 35) As TVERTEX2
Private SkySkin(0 To 4) As Direct3DTexture8
Private SkyVBuf As Direct3DVertexBuffer8

Private PlayerPlaq(0 To 35) As TVERTEX2
Private PlayerSkin As Direct3DTexture8
Private PlayerVBuf As Direct3DVertexBuffer8

Private Sub ReformPlayer()
    Const height As Long = 90
    Const width As Long = 20

    CreateSquareCo PlayerPlaq, 0, 0, 0, MakeVector(-(width / 2), (height / 2), (width / 2)), _
                        MakeVector(-(width / 2), (height / 2), -(width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), -(width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), (width / 2))
    CreateSquareCo PlayerPlaq, 6, 0, 1, MakeVector(-(width / 2), (height / 2), -(width / 2)), _
                        MakeVector((width / 2), (height / 2), -(width / 2)), _
                        MakeVector((width / 2), -(height / 2), -(width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), -(width / 2))
    CreateSquareCo PlayerPlaq, 12, 0, 2, MakeVector((width / 2), (height / 2), -(width / 2)), _
                        MakeVector((width / 2), (height / 2), (width / 2)), _
                        MakeVector((width / 2), -(height / 2), (width / 2)), _
                        MakeVector((width / 2), -(height / 2), -(width / 2))
    CreateSquareCo PlayerPlaq, 18, 0, 3, MakeVector((width / 2), -(height / 2), (width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), (width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), -(width / 2)), _
                        MakeVector((width / 2), -(height / 2), -(width / 2))
    CreateSquareCo PlayerPlaq, 24, 0, 4, MakeVector((width / 2), (height / 2), (width / 2)), _
                        MakeVector(-(width / 2), (height / 2), (width / 2)), _
                        MakeVector(-(width / 2), -(height / 2), (width / 2)), _
                        MakeVector((width / 2), -(height / 2), (width / 2))
    CreateSquareCo PlayerPlaq, 30, 0, 5, MakeVector((width / 2), (height / 2), -(width / 2)), _
                        MakeVector(-(width / 2), (height / 2), -(width / 2)), _
                        MakeVector(-(width / 2), (height / 2), (width / 2)), _
                        MakeVector((width / 2), (height / 2), (width / 2))
                        
    Player.Location.Y = (160 / 2 + 10)
End Sub

Public Sub CreateLand()

    Set PlayerSkin = LoadTexture(AppPath & "Base\invisible.bmp")

    ReformPlayer

    Set PlayerVBuf = DDevice.CreateVertexBuffer(Len(PlayerPlaq(0)) * (UBound(PlayerPlaq) + 1), 0, FVF_VTEXT2, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData PlayerVBuf, 0, Len(PlayerPlaq(0)) * (UBound(PlayerPlaq) + 1), 0, PlayerPlaq(0)

    Set SkySkin(0) = LoadTexture(AppPath & "Base\sky_top.bmp")
    Set SkySkin(1) = LoadTexture(AppPath & "Base\sky_back.bmp")
    Set SkySkin(2) = LoadTexture(AppPath & "Base\sky_left.bmp")
    Set SkySkin(3) = LoadTexture(AppPath & "Base\sky_front.bmp")
    Set SkySkin(4) = LoadTexture(AppPath & "Base\sky_right.bmp")
   ' Set SkySkin(5) = LoadTexture(AppPath & "Base\sky_btm.bmp")
    
    CreateSquare SkyPlaq, 0, MakeVector(-5, -5, 5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(-5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 6, MakeVector(-5, -5, -5), _
                            MakeVector(5, -5, -5), _
                            MakeVector(5, 5, -5), _
                            MakeVector(-5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 12, MakeVector(5, -5, -5), _
                            MakeVector(5, -5, 5), _
                            MakeVector(5, 5, 5), _
                            MakeVector(5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 18, MakeVector(5, -5, -5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(5, -5, 5), 1, 1
    CreateSquare SkyPlaq, 24, MakeVector(5, -5, 5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 30, MakeVector(5, 5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(5, 5, -5), 1, 1
'    CreateSquare SkyPlaq, 36, MakeVector(5, -5, -5), _
'                            MakeVector(-5, -5, -5), _
'                            MakeVector(-5, -5, 5), _
'                            MakeVector(5, -5, 5), 1, 1
                       
    Set SkyVBuf = DDevice.CreateVertexBuffer(Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, FVF_VTEXT2, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SkyVBuf, 0, Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, SkyPlaq(0)
                        
End Sub

Public Function MovePlayer(ByRef Mov As D3DVECTOR)

    Dim cnt As Long
    Dim Index As Long

    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim v3 As D3DVECTOR

    Dim matRotate As D3DMATRIX
    Dim matMove As D3DMATRIX

    D3DXMatrixTranslation matMove, Mov.X, Mov.Y, Mov.Z
    D3DXMatrixRotationYawPitchRoll matRotate, 0, 0, 0

    Index = 0
    For cnt = 0 To 11

        v1.X = PlayerPlaq(Index + 0).X
        v1.Y = PlayerPlaq(Index + 0).Y
        v1.Z = PlayerPlaq(Index + 0).Z

        v2.X = PlayerPlaq(Index + 1).X
        v2.Y = PlayerPlaq(Index + 1).Y
        v2.Z = PlayerPlaq(Index + 1).Z

        v3.X = PlayerPlaq(Index + 2).X
        v3.Y = PlayerPlaq(Index + 2).Y
        v3.Z = PlayerPlaq(Index + 2).Z

        D3DXVec3TransformCoord v1, v1, matRotate
        D3DXVec3TransformCoord v2, v2, matRotate
        D3DXVec3TransformCoord v3, v3, matRotate

        sngVertexX(0, cnt) = v1.X + Player.Location.X
        sngVertexY(0, cnt) = v1.Y + Player.Location.Y
        sngVertexZ(0, cnt) = v1.Z + Player.Location.Z

        sngVertexX(1, cnt) = v2.X + Player.Location.X
        sngVertexY(1, cnt) = v2.Y + Player.Location.Y
        sngVertexZ(1, cnt) = v2.Z + Player.Location.Z

        sngVertexX(2, cnt) = v3.X + Player.Location.X
        sngVertexY(2, cnt) = v3.Y + Player.Location.Y
        sngVertexZ(2, cnt) = v3.Z + Player.Location.Z

        Index = Index + 3
    Next

End Function

Private Function WillCollide(ByRef Mov As D3DVECTOR) As Boolean

    MovePlayer Mov

    Dim retBrushNum As Long
    Dim retFaceNum As Long

    Dim cnt As Long
    For cnt = 0 To 11
        retBrushNum = 0: retFaceNum = 0
        WillCollide = WillCollide Or Collision(0, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, cnt, retBrushNum, retFaceNum)
        If WillCollide Then Exit Function
    Next

End Function

Public Sub AttemptMoves(ByRef Mov As D3DVECTOR)

    Player.Location.X = Player.Location.X + Mov.X
    Player.Location.Y = Player.Location.Y + Mov.Y
    Player.Location.Z = Player.Location.Z + Mov.Z

    If WillCollide(Mov) Then

        Player.Location.X = Player.Location.X - Mov.X
        Player.Location.Y = Player.Location.Y - Mov.Y
        Player.Location.Z = Player.Location.Z - Mov.Z
    End If

    MovePlayer Mov

End Sub

Public Sub RenderLand()

    DDevice.SetMaterial GenericMat

    If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
    End If
    
    DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetVertexShader FVF_VTEXT0
    
    DDevice.SetRenderState D3DRS_FOGENABLE, False
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, False
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT

    Dim matProj As D3DMATRIX
    Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
    DDevice.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0

    DDevice.SetTransform D3DTS_VIEW, matView
    DDevice.SetTransform D3DTS_WORLD, matTemp

    If (Sunrotated = 0) Or (((Timer - Sunrotated) * 0.06) >= 360) Then Sunrotated = Timer

    D3DXMatrixPerspectiveFovLH matProj, PI / 3.5, 0.82, 1, 50
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    D3DXMatrixRotationY matTemp, ((Timer - Sunrotated) * 0.03) * (PI / 180)
    DDevice.SetTransform D3DTS_WORLD, matTemp

    DDevice.SetStreamSource 0, SkyVBuf, Len(SkyPlaq(0))

    DDevice.SetTexture 0, SkySkin(0)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 30, 2
    DDevice.SetTexture 0, SkySkin(1)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    DDevice.SetTexture 0, SkySkin(2)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
    DDevice.SetTexture 0, SkySkin(3)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
    DDevice.SetTexture 0, SkySkin(4)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 24, 2
    
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 0.75, 5, 50000
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    DDevice.SetTransform D3DTS_VIEW, matViewSave
    DDevice.SetTransform D3DTS_WORLD, matTemp

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_FOGENABLE, 1

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    Dim matMove As D3DMATRIX
    Dim matRotate As D3DMATRIX
    
    DDevice.SetStreamSource 0, PlayerVBuf, Len(PlayerPlaq(0))

    DDevice.SetTexture 0, PlayerSkin
    
    D3DXMatrixRotationYawPitchRoll matRotate, -Player.CameraAngle, 0, 0
    D3DXMatrixTranslation matMove, Player.Location.X, Player.Location.Y, Player.Location.Z
    D3DXMatrixMultiply matMove, matRotate, matMove
    DDevice.SetTransform D3DTS_WORLD, matMove

    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12

End Sub

Public Sub CleanupLand()
    Set PlayerSkin = Nothing
    Set PlayerVBuf = Nothing
End Sub

Private Sub CreateSquareCo(ByRef Data() As TVERTEX2, ByVal Index As Long, ByVal BrushIndex As Long, ByVal FaceIndex As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1)

    Dim vn As D3DVECTOR

    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.Z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.Z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.Z

    AddVisFace BrushIndex, FaceIndex, vn, MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z)

    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, 0, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.Z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.Z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.Z

    AddVisFace BrushIndex, FaceIndex, vn, MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z)
End Sub

Public Sub AddVisFace(ByVal BrushIndex As Long, ByVal FaceIndex As Long, ByRef vn As D3DVECTOR, ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR)

    ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

    sngVertexX(0, lngFaceCount) = v0.X
    sngVertexY(0, lngFaceCount) = v0.Y
    sngVertexZ(0, lngFaceCount) = v0.Z

    sngVertexX(1, lngFaceCount) = v1.X
    sngVertexY(1, lngFaceCount) = v1.Y
    sngVertexZ(1, lngFaceCount) = v1.Z

    sngVertexX(2, lngFaceCount) = v2.X
    sngVertexY(2, lngFaceCount) = v2.Y
    sngVertexZ(2, lngFaceCount) = v2.Z

    sngFaceVis(0, lngFaceCount) = vn.X
    sngFaceVis(1, lngFaceCount) = vn.Y
    sngFaceVis(2, lngFaceCount) = vn.Z
    sngFaceVis(3, lngFaceCount) = 0
    sngFaceVis(4, lngFaceCount) = BrushIndex
    sngFaceVis(5, lngFaceCount) = FaceIndex

    lngFaceCount = lngFaceCount + 1

End Sub

Public Sub ResetCollision()
    lngFaceCount = 11
    ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
    ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single
    lngFaceCount = lngFaceCount + 1
End Sub

