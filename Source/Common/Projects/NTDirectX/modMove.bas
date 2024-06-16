Attribute VB_Name = "modMove"
#Const modMove = -1
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

Public Declare Function Collision Lib "MaxLandLib" (ByVal visType As Long, ByVal lngFaceCount As Long, _
                        ByRef sngFaceVis() As Single, ByRef sngVertexX() As Single, ByRef sngVertexY() As Single, ByRef sngVertexZ() As Single, _
                        ByVal lngFaceNum As Long, ByRef lngCollidedBrush As Long, ByRef lngCollidedFace As Long) As Boolean
                        
Public Declare Function Culling Lib "MaxLandLib" Alias "Forystek" (ByVal visType As Long, ByVal lngFaceCount As Long, _
                        ByRef sngCamera() As Single, ByRef sngFaceVis() As Single, ByRef sngVertexX() As Single, ByRef sngVertexY() As Single, ByRef sngVertexZ() As Single, _
                        ByRef sngScreenX() As Single, ByRef sngScreenY() As Single, ByRef sngScreenZ() As Single, ByRef sngZBuffer() As Single) As Long
                        
'############################################################################################################
'Variable Declare ###########################################################################################
'############################################################################################################

Public Const CULL0 = 0
Public Const CULL1 = 1
Public Const CULL2 = 2
Public Const CULL3 = 4
Public Const CULL4 = 3
Public Const CULL5 = 0
Public Const CULL6 = -4

Public lCullCalls As Long
Public lCulledFaces As Long
Public lMovingObjs As Long
Public lFacesShown As Long

Public lngObjCount As Long
Public lngFaceCount As Long

Public lngTestCalls As Long

Public sngFaceVis() As Single
'sngFaceVis dimension (,n) where n=# is face number
'sngFaceVis dimension (n,) where n=0 is x of face normal
'sngFaceVis dimension (n,) where n=1 is y of face normal
'sngFaceVis dimension (n,) where n=2 is z of face normal
'sngFaceVis dimension (n,) where n=3 is vis Type, values
'sngFaceVis dimension (n,) where n=4 is gBrush index
'sngFaceVis dimension (n,) where n=4 is gFace index

Public sngVertexX() As Single
Public sngVertexY() As Single
Public sngVertexZ() As Single
'sngVertexX dimension (,n) where n=# is face number
'sngVertexX dimension (n,) where n=0 is faces first vertex.X
'sngVertexX dimension (n,) where n=1 is faces second vertex.X
'sngVertexX dimension (n,) where n=2 is faces third vertex.X
'sngVertexX dimension (n,) where n=3 is faces fourth vertex.X

Public sngCamera() As Single
'sngCamera dimension (0,n) is camera position, n=0=x, n=1=y, n=2=z
'sngCamera dimension (1,n) is camera direction, n=0=x, n=1=y, n=2=z
'sngCamera dimension (2,n) is camera up vector, n=0=x, n=1=y, n=2=z

Public sngScreenX() As Single
Public sngScreenY() As Single
Public sngScreenZ() As Single
Public sngZBuffer() As Single

Public DebugFace() As MyVertex
Public DebugSkin(0 To 4) As Direct3DTexture8
Public DebugVBuf As Direct3DVertexBuffer8

Public Type MyCulling
    Position As D3DVECTOR
    Direction As D3DVECTOR
    UpVector As D3DVECTOR
    visType As Long
End Type

Public CullingSetup As Integer
Public CullingObject As MyCulling
Public CullingCount As Long
Public Cullings() As MyCulling

Private andCamera() As Single

Private andFaceVis() As Single
Private andVertexX() As Single
Private andVertexY() As Single
Private andVertexZ() As Single

Private andScreenX() As Single
Private andScreenY() As Single
Private andScreenZ() As Single

Private andZBuffer() As Single

Private notCamera() As Single

Private notFaceVis() As Single
Private notVertexX() As Single
Private notVertexY() As Single
Private notVertexZ() As Single

Private notScreenX() As Single
Private notScreenY() As Single
Private notScreenZ() As Single

Private notZBuffer() As Single


Public Sub CreateMove()

    ReDim sngCamera(0 To 2, 0 To 2) As Single

End Sub

Public Sub CleanupMove()


    If CullingCount > 0 Then
        CullingCount = 0
        Erase Cullings
    End If
    
    lngFaceCount = 0
    lngObjCount = 0
    
    Erase sngFaceVis
    
    Erase sngVertexX
    Erase sngVertexY
    Erase sngVertexZ
    
    Erase sngCamera
    
    Erase sngScreenX
    Erase sngScreenY
    Erase sngScreenZ
    Erase sngZBuffer
    
End Sub


Public Sub ComputeNormals()
    Dim cnt As Long
    Dim vn As D3DVECTOR
    
    For cnt = 0 To lngFaceCount - 1
        vn = TriangleNormal(MakeVector(sngVertexX(0, cnt), sngVertexY(0, cnt), sngVertexZ(0, cnt)), _
                            MakeVector(sngVertexX(1, cnt), sngVertexY(1, cnt), sngVertexZ(1, cnt)), _
                            MakeVector(sngVertexX(2, cnt), sngVertexY(2, cnt), sngVertexZ(2, cnt)))
        sngFaceVis(0, cnt) = vn.X
        sngFaceVis(1, cnt) = vn.Y
        sngFaceVis(2, cnt) = vn.Z
    Next
End Sub

Public Sub AddMotion(ByRef Obj As Molecule, ByRef Action As ActionTypes, ByRef Axis As ntobject3d.Point, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional ByVal Identity As String = "")
    Dim act As Motion
    If Identity = "" Then
        Identity = Include.Unnamed(Obj.Motions)
    End If
    
    If Obj.Motions.Exists(Identity) Then
        Set act = Obj.Motions(Identity)
    Else
        Set act = New Motion
        Obj.Motions.Add act, Identity
    End If
    With act
        .Action = Action
        Set .Axis = Axis
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .Latency = Timer
        .Recount = Recount
    End With
End Sub

Public Function DeleteMotion(ByRef Obj As Molecule, ByVal Identity As String) As Boolean
    Dim a As Long
    If Obj.Motion.Exists(Identity) Then
        Obj.Motions.Remove Identity
    End If
End Function


Public Function CalculateMotion(ByRef Motion As Motion, ByRef Action As ActionTypes) As D3DVECTOR

    If (Motion.Action And Action) = Action Then
        
        If Motion.Friction <> 0 Then
            Motion.Emphasis = Motion.Emphasis - (Motion.Emphasis * Motion.Friction)
            If Motion.Emphasis < 0 Then Motion.Emphasis = 0
        End If

        If (Motion.Emphasis > 0.0001) Or (Motion.Emphasis < -0.0001) Then
            CalculateMotion.X = Motion.Axis.X * Motion.Emphasis
            CalculateMotion.Y = Motion.Axis.Y * Motion.Emphasis
            CalculateMotion.Z = Motion.Axis.Z * Motion.Emphasis
        Else
            Motion.Emphasis = 0
        End If
    
    End If
    
End Function

'Public Sub ClearActivities()
'
'    If Molecules.Count > 0 Then
'        Dim o As Long
'        For o = 1 To Molecules.Count
'            Do Until Molecules(o).Motions.Count = 0
'                Molecules(o).Motions.Remove 1
'            Loop
'        Next
'    End If
'End Sub

Public Function AddMeshCollision(ByRef Obj As Molecule, ByVal FaceCount As Long, ByRef Verticies() As D3DVERTEX, ByRef Indicies() As Integer, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError

'#####################################################################################
'############# create face data for a mesh to external compatability #################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim V() As D3DVERTEX

    ReDim V(0 To 2) As D3DVERTEX
    Dim vn As D3DVECTOR
    
    Obj.CollideIndex = lngFaceCount
    Obj.CollideFaces = FaceCount
    AddMeshCollision = lngFaceCount

    Dim matObj As D3DMATRIX
    D3DXMatrixIdentity matObj
    OrientateMolecule matObj, Obj, False

    Index = 0
    For Face = 0 To FaceCount - 1

        For cnt = 0 To 2

            V(cnt).X = Verticies(Indicies(Index + cnt)).X
            V(cnt).Y = Verticies(Indicies(Index + cnt)).Y
            V(cnt).Z = Verticies(Indicies(Index + cnt)).Z

            D3DXVec3TransformCoord vn, ConvertVertexToVector(V(cnt)), matObj
            V(cnt).X = vn.X
            V(cnt).Y = vn.Y
            V(cnt).Z = vn.Z
        Next

        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
        
        vn = TriangleNormal(ConvertVertexToVector(V(0)), ConvertVertexToVector(V(1)), ConvertVertexToVector(V(2)))
        
        For cnt = 0 To 2

            sngVertexX(cnt, lngFaceCount) = V(cnt).X
            sngVertexY(cnt, lngFaceCount) = V(cnt).Y
            sngVertexZ(cnt, lngFaceCount) = V(cnt).Z

        Next

        sngFaceVis(0, lngFaceCount) = vn.X
        sngFaceVis(1, lngFaceCount) = vn.Y
        sngFaceVis(2, lngFaceCount) = vn.Z
        sngFaceVis(3, lngFaceCount) = visType
        sngFaceVis(4, lngFaceCount) = lngObjCount

        sngFaceVis(5, lngFaceCount) = CLng(Replace(CStr(Face / 2), ".5", ""))
        
        lngFaceCount = lngFaceCount + 1

        Index = Index + 3
        
    Next

    Obj.CollideObject = lngObjCount

    lngObjCount = lngObjCount + 1

    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function AddPlaneCollision(ByRef Obj As Plane, ByVal FaceCount As Long, ByRef Verticies() As MyVertex, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError

'#####################################################################################
'############# create face data for a mesh to external compatability #################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim V() As D3DVERTEX

    ReDim V(0 To 2) As D3DVERTEX
    Dim vn As D3DVECTOR

    Obj.CollideIndex = lngFaceCount
    Obj.CollideFaces = FaceCount
    AddPlaneCollision = lngFaceCount
    Dim matObj As D3DMATRIX
    OrientateMolecule matObj, Obj.Molecule, True
    
    Index = 0
    For Face = 0 To FaceCount - 1

        For cnt = 0 To 2

            V(cnt).X = Verticies(Index + cnt).X
            V(cnt).Y = Verticies(Index + cnt).Y
            V(cnt).Z = Verticies(Index + cnt).Z

            D3DXVec3TransformCoord vn, ConvertVertexToVector(V(cnt)), matObj
            V(cnt).X = vn.X
            V(cnt).Y = vn.Y
            V(cnt).Z = vn.Z
        Next

        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
        
        vn = TriangleNormal(ConvertVertexToVector(V(0)), ConvertVertexToVector(V(1)), ConvertVertexToVector(V(2)))
        
        For cnt = 0 To 2

            sngVertexX(cnt, lngFaceCount) = V(cnt).X
            sngVertexY(cnt, lngFaceCount) = V(cnt).Y
            sngVertexZ(cnt, lngFaceCount) = V(cnt).Z

        Next

        sngFaceVis(0, lngFaceCount) = vn.X
        sngFaceVis(1, lngFaceCount) = vn.Y
        sngFaceVis(2, lngFaceCount) = vn.Z
        sngFaceVis(3, lngFaceCount) = visType
        sngFaceVis(4, lngFaceCount) = lngObjCount

        sngFaceVis(5, lngFaceCount) = CLng(Replace(CStr(Face / 2), ".5", ""))
        
        lngFaceCount = lngFaceCount + 1

        Index = Index + 3
        
    Next

    Obj.CollideObject = lngObjCount

    lngObjCount = lngObjCount + 1
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function DelCollision(ByRef Obj As Object, ByVal FaceCount As Long)
On Error GoTo ObjectError

    If Obj.CollideIndex > -1 Then
    
        Dim cnt As Long
        Dim Face As Long
        Dim Index As Long
        
        Index = FaceCount

        If lngFaceCount - Index > 0 Then
    
            For Face = Obj.CollideIndex To lngFaceCount - Index - 1
                sngFaceVis(0, Face) = sngFaceVis(0, Index + Face - 1)
                sngFaceVis(1, Face) = sngFaceVis(1, Index + Face - 1)
                sngFaceVis(2, Face) = sngFaceVis(2, Index + Face - 1)
                sngFaceVis(3, Face) = sngFaceVis(3, Index + Face - 1)
                sngFaceVis(4, Face) = sngFaceVis(4, Index + Face - 1)
                sngFaceVis(5, Face) = sngFaceVis(5, Index + Face - 1)
                sngVertexX(0, Face) = sngVertexX(0, Index + Face - 1)
                sngVertexX(1, Face) = sngVertexX(1, Index + Face - 1)
                sngVertexX(2, Face) = sngVertexX(2, Index + Face - 1)
                sngVertexY(0, Face) = sngVertexY(0, Index + Face - 1)
                sngVertexY(1, Face) = sngVertexY(1, Index + Face - 1)
                sngVertexY(2, Face) = sngVertexY(2, Index + Face - 1)
                sngVertexZ(0, Face) = sngVertexZ(0, Index + Face - 1)
                sngVertexZ(1, Face) = sngVertexZ(1, Index + Face - 1)
                sngVertexZ(2, Face) = sngVertexZ(2, Index + Face - 1)
                
                sngScreenX(0, Face) = sngScreenX(0, Index + Face - 1)
                sngScreenX(1, Face) = sngScreenX(1, Index + Face - 1)
                sngScreenX(2, Face) = sngScreenX(2, Index + Face - 1)
                sngScreenY(0, Face) = sngScreenY(0, Index + Face - 1)
                sngScreenY(1, Face) = sngScreenY(1, Index + Face - 1)
                sngScreenY(2, Face) = sngScreenY(2, Index + Face - 1)
                sngScreenZ(0, Face) = sngScreenZ(0, Index + Face - 1)
                sngScreenZ(1, Face) = sngScreenZ(1, Index + Face - 1)
                sngScreenZ(2, Face) = sngScreenZ(2, Index + Face - 1)
                
                sngZBuffer(0, Face) = sngZBuffer(0, Index + Face - 1)
                sngZBuffer(1, Face) = sngZBuffer(1, Index + Face - 1)
                sngZBuffer(2, Face) = sngZBuffer(2, Index + Face - 1)
                sngZBuffer(3, Face) = sngZBuffer(3, Index + Face - 1)
                
            Next
            
            For cnt = 1 To Molecules.Count
                If Molecules(cnt).CollideIndex > Obj.CollideIndex Then
                    Molecules(cnt).CollideIndex = Molecules(cnt).CollideIndex - Index
                End If
            Next
            For cnt = 1 To Planes.Count
                If Planes(cnt).CollideIndex > Obj.CollideIndex Then
                    Planes(cnt).CollideIndex = Planes(cnt).CollideIndex - Index
                End If
            Next
            
        End If
        
        Obj.CollideIndex = -1
        lngObjCount = lngObjCount - 1
        lngFaceCount = lngFaceCount - Index
        
        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single
    
        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single
    
        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
    End If
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


'
