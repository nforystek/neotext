Attribute VB_Name = "modMove"
#Const modMove = -1
Option Explicit
'TOP DoWN
Option Compare Binary

Option Private Module

'Collision Culling Flag Relation COnstraints
'
'Object
'   Map
'       Player
'           Map
'           Player
'       Object
'           Map
'           Player
'
'   *Self
'       Map
'       Object
'       Player
'   Player
'       Map
'       Object
'   Object
'       Map
'       Player

Public Const CULL0 = 0
Public Const CULL1 = 1
Public Const CULL2 = 2
Public Const CULL3 = 4
Public Const CULL4 = 3
Public Const CULL5 = 0
Public Const CULL6 = -4


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
'Checks for the presence of a point behind a triangle, the first three inputs are the length of the triangles sides, the Next three are the triangles normal, the last three are the point to test with the triangles center removed.

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

Public lCullCalls As Long
Public lCulledFaces As Long
Public lMovingObjs As Long
Public lFacesShown As Long

Public lngObjCount As Long
Public lngFaceCount As Long

Public lngTestCalls As Long

Public sngFaceVis() As Single 'object organization and normals
'sngFaceVis dimension (,n) where n=# is face number (global in count)
'sngFaceVis dimension (n,) where n=0 is x of face normal
'sngFaceVis dimension (n,) where n=1 is y of face normal
'sngFaceVis dimension (n,) where n=2 is z of face normal
'sngFaceVis dimension (n,) where n=3 is vis Type, values (exclude flags)
'sngFaceVis dimension (n,) where n=4 is gBrush index (object number)
'sngFaceVis dimension (n,) where n=4 is gFace index (to vertex arrays)

Public sngVertexX() As Single 'all the 3d data of the collision tests
Public sngVertexY() As Single 'organized by face indexs of 4 vertex
Public sngVertexZ() As Single 'that will be tested for collisions
'sngVertexX dimension (,n) where n=# is face number (global in count)
'sngVertexX dimension (n,) where n=0 is faces first vertex.X
'sngVertexX dimension (n,) where n=1 is faces second vertex.X
'sngVertexX dimension (n,) where n=2 is faces third vertex.X
'sngVertexX dimension (n,) where n=3 is faces fourth vertex.X

Public sngCamera() As Single 'culling exclusion technique
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

'Private andCamera() As Single
'
'Private andFaceVis() As Single
'Private andVertexX() As Single
'Private andVertexY() As Single
'Private andVertexZ() As Single
'
'Private andScreenX() As Single
'Private andScreenY() As Single
'Private andScreenZ() As Single
'
'Private andZBuffer() As Single
'
'Private notCamera() As Single
'
'Private notFaceVis() As Single
'Private notVertexX() As Single
'Private notVertexY() As Single
'Private notVertexZ() As Single
'
'Private notScreenX() As Single
'Private notScreenY() As Single
'Private notScreenZ() As Single
'
'Private notZBuffer() As Single

Public Sub CreateMove()

    ReDim sngCamera(0 To 2, 0 To 2) As Single
    
    Set DebugSkin(0) = LoadTexture(AppPath & "Models\debug0.bmp")
    Set DebugSkin(1) = LoadTexture(AppPath & "Models\debug1.bmp")
    Set DebugSkin(2) = LoadTexture(AppPath & "Models\debug2.bmp")
    Set DebugSkin(3) = LoadTexture(AppPath & "Models\debug4.bmp")
    Set DebugSkin(4) = LoadTexture(AppPath & "Models\debug3.bmp")
    
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


Public Function SetMotion(ByRef act As Motion, ByRef Action As Actions, ByRef dat As Point, ByRef emp As Single) As String
    act.Identity = Replace(modGuid.GUID, "-", "K")
    act.Action = Action
    Set act.Data = dat
    act.Emphasis = emp
    SetMotion = act.Identity
End Function


Public Function MotionExists(ByRef Motions As ntnodes10.Collection, ByVal MGUID As String) As Boolean
    Dim A As Long
    For A = 1 To Motions.Count
        If Motions(A).Identity = MGUID Then
            MotionExists = True
            Exit Function
        End If
    Next
    MotionExists = False
End Function

Public Function ValidMotion(ByRef Motion As Motion) As Boolean
    ValidMotion = (Motion.Identity <> "")
End Function

Public Function CalculateMotion(ByRef Motion As Motion, ByRef Action As Actions) As D3DVECTOR

    If (Action And Motion.Action) = Action Then

        If Motion.Friction <> 0 Then

            Motion.Emphasis = Motion.Emphasis - (Motion.Emphasis * Motion.Friction)
            If Motion.Emphasis < 0 Then
                Motion.Emphasis = 0
                Motion.Identity = ""
            End If
        End If

        If (Motion.Emphasis > 0.001) Or (Motion.Emphasis < -0.001) Then
            CalculateMotion.X = Motion.Data.X * Motion.Emphasis
            CalculateMotion.Y = Motion.Data.Y * Motion.Emphasis
            CalculateMotion.Z = Motion.Data.Z * Motion.Emphasis
        Else
            Motion.Emphasis = 0
        End If

    End If

End Function

Private Sub ApplyMotion(ByRef Obj As Element)
    Dim cnt As Long
    Dim cnt2 As Long
    Dim Offset As D3DVECTOR
    Dim vout As D3DVECTOR
    
    If ((Not (Perspective = Spectator)) And (Obj.CollideObject = Player.CollideObject)) Or (Not (Obj.CollideObject = Player.CollideObject)) Then
        
        If Obj.Gravitational Then
            If Not Obj.OnLadder Then
                If Obj.InLiquid Then
                    D3DXVec3Add vout, ToVector(Obj.Direct), CalculateMotion(LiquidGravityDirect, Directing)
                    Set Obj.Direct = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Twists), CalculateMotion(LiquidGravityRotate, Rotating)
                    Set Obj.Twists = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Scalar), CalculateMotion(LiquidGravityScaled, Scaling)
                    Set Obj.Scalar = ToPoint(vout)
                Else
                    D3DXVec3Add vout, ToVector(Obj.Direct), CalculateMotion(GlobalGravityDirect, Directing)
                    Set Obj.Direct = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Twists), CalculateMotion(GlobalGravityRotate, Rotating)
                    Set Obj.Twists = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Scalar), CalculateMotion(GlobalGravityScaled, Scaling)
                    Set Obj.Scalar = ToPoint(vout)
                End If
            End If
        End If
    End If
    
    If Obj.Effect = Collides.None Then

        If Obj.Motions.Count > 0 Then
            Dim A As Long
            For A = 1 To Obj.Motions.Count
                If ValidMotion(Obj.Motions(A)) Then
                    D3DXVec3Add vout, ToVector(Obj.Direct), CalculateMotion(Obj.Motions(A), Directing)
                    Set Obj.Direct = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Twists), CalculateMotion(Obj.Motions(A), Rotating)
                    Set Obj.Twists = ToPoint(vout)
                    D3DXVec3Add vout, ToVector(Obj.Scalar), CalculateMotion(Obj.Motions(A), Scaling)
                    Set Obj.Scalar = ToPoint(vout)
                End If
            Next
        End If
    End If
End Sub
Public Sub ResetMotion()
    Dim A As Long
    Dim o As Long
    Set Player.Direct = MakePoint(0, 0, 0)
    If Elements.Count > 0 Then
        For o = 1 To Elements.Count
           Set Elements(o).Direct = MakePoint(0, 0, 0)
        Next
    End If
End Sub

Public Sub RenderMotion()
On Error GoTo ObjectError

    RenderMotion2 Player
    
    If Elements.Count > 0 Then
        Dim e1 As Element
        For Each e1 In Elements
    
            RenderMotion2 e1
            
        Next
    End If
    

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub RenderMotion2(ByRef e1 As Element)
    Dim d As Boolean
    Dim o As Long
    Dim A As Long
    Dim act As Motion
    Dim trig As String
    Dim line As String
    Dim id As String

    Do
    Loop Until (Not e1.DeleteMotion(""))
    
    If e1.Visible Then
        ApplyMotion e1

        If e1.Motions.Count > 0 Then
            A = 1
            Do While A <= e1.Motions.Count
                Set act = e1.Motions(A)

                If act.Reactive > -1 Then
                    If (Timer - act.latency) > act.Reactive Then
                        act.latency = Timer
                        
                        act.Emphasis = act.Initials
                        e1.DeleteMotion act.Identity
                        If Not act.Script = "" Then
                            
                            ExecuteScript e1, act.Script
                        
                        End If
                        If act.Recount > -1 Then
                            If act.Recount > 0 Then
                                act.Recount = act.Recount - 1
                                e1.AddMotion act.Action, act.Identity, act.Data, act.Initials, act.Friction, act.Reactive, act.Recount, act.Script
                                'a = a + 1
                            End If
                        Else
                            e1.AddMotion act.Action, act.Identity, act.Data, act.Initials, act.Friction, act.Reactive, act.Recount, act.Script
                            'a = a + 1
                        End If
                        A = A + 1
                        
                    Else
                        A = A + 1
                    End If
                    
                ElseIf ((act.Emphasis = 0) Or (act.Recount = 0)) Then 'And (Not act.Reactive = -1) Then
                    e1.DeleteMotion act.Identity
                Else
                    A = A + 1
                End If
                Set act = Nothing
            Loop
        End If
    End If

End Sub


Public Sub InputMove()
On Error GoTo ObjectError

    lFacesShown = 0
    lMovingObjs = 0
    lngTestCalls = 0
      
    Dim e1 As Element
    
    If ((Perspective = Spectator) Or DebugMode) Then
    
        Player.Origin.X = Player.Origin.X + Player.Direct.X
        Player.Origin.Y = Player.Origin.Y + Player.Direct.Y
        Player.Origin.Z = Player.Origin.Z + Player.Direct.Z
                
    Else
    
        InputMove2 Player

    End If

    If (Elements.Count > 0) And (Not DebugMode) Then
        Dim A As Long
  
        For Each e1 In Elements

            If (e1.Effect = Collides.None) Then
            
                InputMove2 e1
                
            End If
        Next
    End If


    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Public Sub InputMove2(ByRef e1 As Element)

    Dim e2 As Element
        
    If Not e1.Attachments Is Nothing Then
        If e1.Attachments.Count > 0 Then
        'if we have attachments
        
            'originals
            Dim oOrigin As New Point
            Dim oRotate As New Point
            Dim oScaled As New Point
            
            oOrigin = e1.Origin
            oRotate = e1.Rotate
            oScaled = e1.Scaled
            
            'differences
            Dim dOrigin As New Point
            Dim dRotate As New Point
            Dim dScaled As New Point
                   
        End If
    End If
        
    If (e1.CollideIndex > -1) Then
      
        'commit the following object view changes based on then
        'functions called to determine the restrictive nature of
        
        SpinObject e1
        BlowObject e1
        MoveObject e1

        lFacesShown = lFacesShown + e1.CulledFaces 'statistics
        lMovingObjs = lMovingObjs + 1
        
        

        
    Else
        'the freespace changes similar to the functions above with no restrictions
        If (e1.Direct.X <> 0) Or (e1.Direct.Y <> 0) Or (e1.Direct.Z <> 0) Then
            e1.Origin.X = e1.Origin.X + e1.Direct.X
            e1.Origin.Y = e1.Origin.Y + e1.Direct.Y
            e1.Origin.Z = e1.Origin.Z + e1.Direct.Z
        End If
       ' e1.Direct = NoPoint
        If (e1.Twists.X <> 0) Or (e1.Twists.Y <> 0) Or (e1.Twists.Z <> 0) Then
            e1.Rotate.X = e1.Rotate.X + e1.Twists.X
            e1.Rotate.Y = e1.Rotate.Y + e1.Twists.Y
            e1.Rotate.Z = e1.Rotate.Z + e1.Twists.Z
        End If
        'e1.Twists = NoAngle
        If (e1.Scalar.X <> 0) Or (e1.Scalar.Y <> 0) Or (e1.Scalar.Z <> 0) Then
            e1.Scaled.X = e1.Scaled.X + e1.Scalar.X
            e1.Scaled.Y = e1.Scaled.Y + e1.Scalar.Y
            e1.Scaled.Z = e1.Scaled.Z + e1.Scalar.Z
        End If
       ' e1.Scalar = NoPoint
    End If

    'preform boundary restriction tests and adjust accordingly
    If (e1.Origin.Y > SpaceBoundary) Or (e1.Origin.Y < -SpaceBoundary) Then e1.Origin.Y = -e1.Origin.Y
    If (e1.Origin.X > SpaceBoundary) Or (e1.Origin.X < -SpaceBoundary) Then e1.Origin.X = -e1.Origin.X
    If (e1.Origin.Z > SpaceBoundary) Or (e1.Origin.Z < -SpaceBoundary) Then e1.Origin.Z = -e1.Origin.Z
    
    
    If Not e1.Attachments Is Nothing Then
        If e1.Attachments.Count > 0 Then
            'if we have attachments

            'get the differences in the parent objectss changes
            Set dOrigin = VectorDeduction(e1.Origin, oOrigin)
            Set dRotate = VectorMultiplyBy(AngleAxisDeduction(VectorMultiplyBy(e1.Rotate, RADIAN), VectorMultiplyBy(oRotate, RADIAN)), DEGREE)
            Set dScaled = VectorDeduction(e1.Scaled, dScaled)
            
            '## originals ##
            'oOrigin
            'oRotate
            'oScaled
            
            '## differences ##
            'dOrigin
            'dOrigin
            'dScaled
            
            '## finals ##
            'e1.Origin
            'e1.Rotate
            'e1.Scaled
            
            
            For Each e2 In e1.Attachments
                'per each attachment
            
                'make e1's origin the point (0,0,0) according
                'to e2's origin berfore e1 had modifications.
                Set e2.Origin = VectorDeduction(e2.Origin, oOrigin)
                'if the rotation is not blank
                If Not ((Round(oRotate.X, 0) = 360) And (Round(oRotate.Y, 0) = 358) And (Round(oRotate.Z, 0) = 360)) Then
                    'revert the old rotation
                    Set e2.Origin = VectorRotateAxis(e2.Origin, VectorMultiplyBy(oRotate, RADIAN))
                    'rotate to the new rotation
                    Set e2.Origin = VectorRotateAxis(e2.Origin, AngleAxisInvert(VectorMultiplyBy(e1.Rotate, RADIAN)))
                End If
                'restore e2's origin localization of (0,0, 0) at e1's
                'origin to what now it would be after changed e1.origin
                Set e2.Origin = VectorAddition(e2.Origin, e1.Origin)
            
            Next
        End If
    End If

    Set oScaled = Nothing
    Set oRotate = Nothing
    Set oOrigin = Nothing

End Sub

Public Function CoupleMove(ByRef Obj As Element, ByVal objCollision As Long) As Boolean
'###################################################################################
'########## couple the activities of objects in collision with others ##############
'###################################################################################

    Dim A As Long
    Dim cnt As Long
    Dim act As Motion
    If (objCollision > -1) Then
        If (Elements.Count > 0) Then
            Dim e1 As Element
            For Each e1 In Elements

                If (e1.Effect = Collides.None) And (Obj.CollideIndex > -1) Then
                    If (Not e1.CollideObject = Obj.CollideObject) Then
                        If (e1.CollideObject = objCollision) Then
                        'if found to be with the colliding object
                        
                            'add all motions from one to another
                            For A = 1 To Obj.Motions.Count
                                Set act = Obj.Motions(A)
                                e1.AddMotion act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.Script
                            Next

                            e1.Direct = Obj.Direct
                            e1.Scalar = Obj.Scalar
                           
                            'TODO: a bit part of rotation on interaction, if, center of push is withi in averages it's point defined centoid
                            ' then a inverted rotation to the pusher object is correct, otherwise off centoid the motion mor or less determines
                            'the object would rotate until not pushed
                            
                            
                            'average out points in first object
                            'average out points in second object
                            'find distance of those points mid point
                            'find distance of the points vector direction
                            'find the pushing objects vector direction difference
                            'if the angle is with in -45 to  45 pushing the object midly directs it until
                            'if the angle is <-45 or >45 pushing the object rotates it the mild direction
   
                            'then repeat for the vertical axis and such that flat floor would not rotate or
                            'mild direct it could possibly push over like rotate if the weights are included
                            'rather their vertical direction should not change when flat on the floor for now
                            e1.Twists = Obj.Twists
                            
                            CoupleMove = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    End If
End Function

Private Sub MoveObject(ByRef Obj As Element)
On Error GoTo ObjectError

    If Obj.Direct.Equals(NoPoint) Then Exit Sub

    Dim objCollision As Long
    objCollision = -1

    Dim adjust As Single
    Dim visType As Long
    Dim bitType As Long
    bitType = 1
    visType = 2
    adjust = 0.019

    Dim pull As Boolean
    Dim push As Boolean

    Dim tmpset As D3DVECTOR
    Dim newset As D3DVECTOR

    Dim cnt As Long
    Dim cnt2 As Long
    Dim act As Motion
    
'#####################################################################################
'############# preliminary sort the type of space collision checks ###################
'#####################################################################################

    'reset all of the vis flags to zero
    'set to zero, culling ignores them
    Obj.IsMoving = Moving.None
    For cnt = 0 To lngFaceCount - 1
        sngFaceVis(3, cnt) = 0
    Next
   
    Dim e1 As Element
    If (Elements.Count > 0) Then
        'this first look is for laddre effects, visType is a flag for Culling to map which need weaning out
        For Each e1 In Elements 'mark only collidable interests to be flagged as 1 rather then the 0 set above
            If (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.Visible And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = bitType
                Next
            End If
        Next


        If Obj.OnLadder Then 'if we are already on ladder coming in
            Obj.OnLadder = TestCollision(Obj, Actions.None, bitType) 'straight to test
        Else
            Obj.OnLadder = TestCollision(Obj, Actions.None, bitType) 'test as well but..
            If Obj.OnLadder Then 'if this is the first time we are
                Do 'on a ladder coming in, clear the objects motions
                Loop Until Not Obj.DeleteMotion(JumpGUID)
                For cnt = 1 To Portals.Count
                    For cnt2 = 1 To Portals(cnt).Motions.Count
                        Set act = Portals(cnt).Motions(cnt2)
                        Obj.DeleteMotion act.Identity
                        Set act = Nothing
                    Next
                Next
            End If
        End If

        For Each e1 In Elements 'we finished ladder mechanics, now we move on to liquid and clear the bits
            If (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.Visible And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = bitType
                Next
            End If
        Next

        If Obj.InLiquid Then 'the same as ladder, if already liquid;
            Obj.InLiquid = TestCollision(Obj, Actions.None, bitType) 'straight to test
        Else
            Obj.InLiquid = TestCollision(Obj, Actions.None, bitType) 'test as well but..
            If Obj.InLiquid Then 'first time in liquid then
                Do 'delete motions and apeal motions reference
                Loop Until Not Obj.DeleteMotion(JumpGUID)
                For cnt = 1 To Portals.Count
                    For cnt2 = 1 To Portals(cnt).Motions.Count
                        Set act = Portals(cnt).Motions(cnt2)
                        Obj.DeleteMotion act.Identity
                        Set act = Nothing
                    Next
                Next
            End If
        End If

    End If


'#####################################################################################
'############# initial faces data for returning TestCollision info ###################
'#####################################################################################

    'setup the cngCamera() for Culling to work idealy, this could use improvement
    'it is used to help eliminatea majority of triangles from ever being checked
    'by addressing the collision check triangles are within this camera viewport
    'far slower, is a version of visType to eliminate this and thoroughly check all
    sngCamera(0, 0) = Obj.Origin.X
    sngCamera(0, 1) = Obj.Origin.Y + 1
    sngCamera(0, 2) = Obj.Origin.Z

    sngCamera(1, 0) = 1
    sngCamera(1, 1) = -1
    sngCamera(1, 2) = -1

    sngCamera(2, 0) = -1
    sngCamera(2, 1) = 1
    sngCamera(2, 2) = -1

    If lngFaceCount > 0 Then 'apply the camera perspective to wean out triangles (will be reflected in the flag value)
        Obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
        lCullCalls = lCullCalls + 1
    End If

    If (Elements.Count > 0) Then 'after Culling() is called the flags are set for the visTYpe or not
        For Each e1 In Elements 'but since we already did our ladder and liquid, re-exclude them.
            If (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'reset to make culling() ignore these in collision triangles
                Next
            ElseIf (e1.Effect = Collides.Ground) And (e1.CollideIndex > -1) And e1.Visible And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    If Not (((sngFaceVis(0, cnt2) = 0) Or (sngFaceVis(0, cnt2) = 1) Or (sngFaceVis(0, cnt2) = -1)) And _
                        ((sngFaceVis(1, cnt2) = 0) Or (sngFaceVis(1, cnt2) = 1) Or (sngFaceVis(1, cnt2) = -1)) And _
                        ((sngFaceVis(2, cnt2) = 0) Or (sngFaceVis(2, cnt2) = 1) Or (sngFaceVis(2, cnt2) = -1))) Then
                        sngFaceVis(3, cnt2) = visType 'the ground is the only one we want to focus on coming up
                    End If
                Next
            ElseIf (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'reset to make culling() ignore these in collision triangles
                Next
            End If
        Next
    End If
    

'#####################################################################################
'############# predict the Y movements of objects in motion ##########################
'#####################################################################################

    'where the directed info data is
    'only going to be about the Y axis
    tmpset = ToVector(Obj.Direct)
    Obj.Direct.Y = tmpset.Y
    Obj.Direct.X = 0
    Obj.Direct.Z = 0

    'all the collision tests use motion data to modify values of a subset of object change
    'that object change is not applied, and any change that will normally, is ahed of time
    'in a way these are predictions of change, tested for collision 1st before binds them
    If (Obj.Direct.Y <> 0) Then
        'preform check since any Y change exists at all
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y  'no collision then adjust the Y to reflect the change is available
            If Obj.Direct.Y > 0 Then 'and then midify the IsMoving state property of the object
                If Not ((Obj.IsMoving And Moving.Flying) = Moving.Flying) Then Obj.IsMoving = Obj.IsMoving + Moving.Flying
            ElseIf Obj.Direct.Y < 0 Then
                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
            End If
            newset.Y = Obj.Direct.Y
        ElseIf (Obj.Direct.Y < 0) Then 'the y movement is going down
            Do '(x,y may have or not have changed here too cause Y change)
                Obj.Direct.Y = Obj.Direct.Y + adjust 'so, we loop until we find out
                If (Obj.Direct.Y >= 0) Then Exit Do 'of the collision where stands
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Y < 0) Then
                Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y 'change the Y to new data, and adjust the IsMoving state for falling
                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
                newset.Y = Obj.Direct.Y
            End If
        ElseIf (Obj.Direct.Y > 0) Then 'the y movement is going up
            Do '(x,y may have or not have changed here too cause Y change)
                Obj.Direct.Y = Obj.Direct.Y - adjust 'so, we loop until we find out
                If (Obj.Direct.Y <= 0) Then Exit Do 'of the collision where stands
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Y > 0) Then
                Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y 'change the Y to new data, and adjust the IsMoving state for falling
                If Not ((Obj.IsMoving And Moving.Flying) = Moving.Flying) Then Obj.IsMoving = Obj.IsMoving + Moving.Flying
                newset.Y = Obj.Direct.Y
            End If
        End If
    End If
    
'#####################################################################################
'############# adjust face data based on the TestCollision resulted ##################
'#####################################################################################


    If (Elements.Count > 0) Then
        For Each e1 In Elements 'reset the types of Collision effects to be only object to object collision
            If (e1.CollideObject = Obj.CollideObject) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = visType ' on sero here ensures Culling() to consider it left in
                Next
            ElseIf (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'still no ladder checking, we got it complete first thing
                Next
            ElseIf (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'still no kiquid checking, we got it complete first thing
                Next
            End If
        Next
    End If
    
'#####################################################################################
'############# last call to MoveObejct collisions couple Motion here ###############
'#####################################################################################

    CoupleMove Obj, objCollision 'periodic TestCollisions may have occured a collision.
    
    'objCollision is a parameter to TestCollision and it contains the colliding object
    'when if it is, there is one, and it is not itself. Since all positive TestCollision
    'results are used in this full function predicament for motions, then we need to
    'make calls to CoupleMove intermediatly so that inertia is applied correctly
    
'#####################################################################################
'############# predict the X movements of objects in motion ##########################
'#####################################################################################

    Obj.Direct.Y = 0
    Obj.Direct.X = tmpset.X

    'very similar recent code above on the Y axis, we will be doing it
    If (Obj.Direct.X <> 0) Then 'on the X (here) and later on the Z axis
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then 'make the change
            Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust the flags
            If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
            If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
            End If
            newset.X = Obj.Direct.X
        ElseIf (Obj.Direct.X < 0) Then
            Do
                Obj.Direct.X = Obj.Direct.X + adjust
                If (Obj.Direct.X >= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.X < 0) Then 'make the change
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust the flags
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
                End If
                newset.X = Obj.Direct.X
            End If
        ElseIf (Obj.Direct.X > 0) Then
            Do
                Obj.Direct.X = Obj.Direct.X - adjust
                If (Obj.Direct.X <= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.X > 0) Then 'make the change
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust the flags
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
                End If
                newset.X = Obj.Direct.X
            End If
        End If
    End If

'#####################################################################################
'############# predict the Z movements of objects in motion ##########################
'#####################################################################################
    
    Obj.Direct.X = 0
    Obj.Direct.Z = tmpset.Z

    'very similar recent code above on the X and Y axis, we will
    If (Obj.Direct.Z <> 0) Then 'be doing it here on the Z axis
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then 'make the change
            Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'add the movement, and adjust the flags
            If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level 'adjust
            If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
            End If
            newset.Z = Obj.Direct.Z
        ElseIf (Obj.Direct.Z < 0) Then
            Do
                Obj.Direct.Z = Obj.Direct.Z + adjust
                If (Obj.Direct.Z >= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Z < 0) Then 'make the change
                Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'add the movement, and adjust the flags
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
                End If
                newset.Z = Obj.Direct.Z
            End If
        ElseIf (Obj.Direct.Z > 0) Then
            Do
                Obj.Direct.Z = Obj.Direct.Z - adjust
                If (Obj.Direct.Z <= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Z > 0) Then 'make the change
                Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'add the movement, and adjust the flags
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.Z <> newset.Z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving - Moving.Falling
                End If
                newset.Z = Obj.Direct.Z
            End If
        End If
    End If

    Set Obj.Direct = ToPoint(newset)
        
    'everything above here in testing all the moving was for the "newest" point which
    'is the disposition of a prediction from its next motion to the current origin.
    '
    'now below here are some similar testing blocks, but blocks below are setting flags
    'push and pull, to one final possible motion adjustment. the first block defined
    'apart, are testing for (Y + 0.02) in location, prior disposition X and Z are
    'checked to see if it satisfies stepping up. "newest" point will also be modified
    'if the conditions are met to auto motions.  CoupleMOve is a function that at any
    'time a collision is detected between another non Effect objects with collision
    'set, the object is granted all the motions of the moving object to be same paced.
    
'#####################################################################################
'############# before applying predictions couple activities of touching #############
'#####################################################################################
    
    CoupleMove Obj, objCollision 'periodic TestCollisions may have occured a collision.

'#####################################################################################
'############# push/pull of moving objects in Y slope and small step ups #############
'#####################################################################################

    'pull is when the object is on a slope 45 degrees or more, it begins to slide down
    'from gravity.  Push is when an object collides with another non ground element, it
    'chain links to "pushing" the first object furthest from force.  small step up are
    'a vertical wall height the object can automatically drive over, i.e. it's stairs.
    
    If (Not Obj.IsMoving = Moving.None) And _
        (tmpset.X <> newset.X Or tmpset.Z <> newset.Z) And _
        (Not ((Obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling)) Then
        'falling and flying flags are too also check befire here.
        
        Obj.Origin.Y = Obj.Origin.Y + 0.2 'pretend it can step out of it by step up

        Obj.Direct.Y = 0
        Obj.Direct.X = tmpset.X
        Obj.Direct.Z = tmpset.Z

        'the following two flags are the difference
        'between just setting "newst" like above.
        push = True 'one none effect object pushing another
        pull = False 'an object falling diagnal on a slope

        If (Obj.Direct.X <> 0) Or (Obj.Direct.Z <> 0) Then
            'first check for collision and if non exists
            'add them to the actual information data
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                'we need a change of X or Z to consider it a pull, already
                'graivty will take effect to any free falling down objects.
                If Obj.Direct.X <> 0 Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    newset.X = Obj.Direct.X
                    pull = True
                End If
                If Obj.Direct.Z <> 0 Then
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z
                    newset.Z = Obj.Direct.Z
                    pull = True
                End If
            ElseIf (Obj.Direct.X < 0) And (Obj.Direct.Z < 0) Then 'here we do two checks
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    Obj.Direct.Z = Obj.Direct.Z + adjust
                    If ((Obj.Direct.X >= 0) Or (Obj.Direct.Z >= 0)) Then Exit Do
                'slow down the change prediction and check until no collision is found
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) And (Obj.Direct.Z < 0) Then
                    'adjust change and flags to reflect happened
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.Z = Obj.Direct.Z
                    pull = True
                End If

            ElseIf (Obj.Direct.X > 0) And (Obj.Direct.Z > 0) Then 'here we do two checks
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    Obj.Direct.Z = Obj.Direct.Z - adjust
                    If ((Obj.Direct.X <= 0) Or (Obj.Direct.Z <= 0)) Then Exit Do
                'slow down the change prediction and check until no collision is found
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) And (Obj.Direct.Z > 0) Then
                    'adjust change and flags to reflect happened
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.Z = Obj.Direct.Z
                    pull = True
                End If

            ElseIf (Obj.Direct.X < 0) And (Obj.Direct.Z > 0) Then 'here we do two checks
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    Obj.Direct.Z = Obj.Direct.Z - adjust
                    If ((Obj.Direct.X >= 0) Or (Obj.Direct.Z <= 0)) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) And (Obj.Direct.Z > 0) Then
                    'adjust change and flags to reflect happened
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.Z = Obj.Direct.Z
                    pull = True
                End If
            ElseIf (Obj.Direct.X > 0) And (Obj.Direct.Z < 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    Obj.Direct.Z = Obj.Direct.Z + adjust
                    If ((Obj.Direct.X <= 0) Or (Obj.Direct.Z >= 0)) Then Exit Do
                    'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) And (Obj.Direct.Z < 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.Z = Obj.Direct.Z
                    pull = True
                End If
            End If
        End If
        
        Obj.Origin.Y = Obj.Origin.Y - 0.2 'no longer pretending it can step up

        If pull Then push = False

    End If

    Set Obj.Direct = ToPoint(tmpset)

'#####################################################################################
'############# those passing with out pressure couple activities first ###############
'#####################################################################################

    
    CoupleMove Obj, objCollision 'periodic TestCollisions may have occured a collision.


'#####################################################################################
'############# as an object first in motions continues it's push in moved Y ##########
'#####################################################################################


    If push And (Not Obj.IsMoving = Moving.None) And _
        (tmpset.X <> newset.X Or tmpset.Z <> newset.Z) And _
        (Not ((Obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling)) Then

        'where a change existe already, during checks on
        'each axis then occurs the need to change again.
        'so besides the gate IF above this is to do it
        'simgularly on X and Z, which was done above so
        '
        Obj.Origin.Y = Obj.Origin.Y + 0.2 'pretend it can step out of it by step up

        Obj.Direct.Y = 0
        Obj.Direct.X = tmpset.X
        Obj.Direct.Z = tmpset.Z

        push = False

        If (Obj.Direct.X <> 0) Then 'first comes the X axis
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust change and flags to reflect happened
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                newset.X = Obj.Direct.X
                push = True
            ElseIf (Obj.Direct.X < 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    If (Obj.Direct.X >= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust change and flags to reflect happened
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    push = True
                End If

            ElseIf (Obj.Direct.X > 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    If (Obj.Direct.X <= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X 'adjust change and flags to reflect happened
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    push = True
                End If
            End If
        End If
        
        If (Obj.Direct.Z <> 0) Then 'first comes the Z axis
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'adjust change and flags to reflect happened
                If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                newset.Z = Obj.Direct.Z
                push = True
            ElseIf (Obj.Direct.Z < 0) Then
                Do
                    Obj.Direct.Z = Obj.Direct.Z + adjust
                    If (Obj.Direct.Z >= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.Z < 0) Then
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'adjust change and flags to reflect happened
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.Z = Obj.Direct.Z
                    push = True
                End If

            ElseIf (Obj.Direct.Z > 0) Then
                Do
                    Obj.Direct.Z = Obj.Direct.Z - adjust
                    If (Obj.Direct.Z <= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.Z > 0) Then
                    Obj.Origin.Z = Obj.Origin.Z + Obj.Direct.Z 'adjust change and flags to reflect happened
                    If Not ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving + Moving.Level
                    newset.Z = Obj.Direct.Z
                    push = True
                End If

            End If
        End If
        
        Obj.Origin.Y = Obj.Origin.Y - 0.2 'no longer pretending it can step up

    End If


'#####################################################################################
'############# coupled in if pushing or pulling, adjust the X/Z gliding ##############
'#####################################################################################

    'next are some final adjustments to requested "Direct" to reflect what is
    'found possible verses what we recieved in attempted moves for an object.
    'due to zero'ing out directive motions, that may re-adjust our push or pull.
    'they are only needed now in testing skipping this block, when not skipped
    'they may become adjusted to skippin the last block of commented apart code
    
    If (pull Xor push) And (Not ((Obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling)) And _
        ((Obj.IsMoving And Moving.Level) = Moving.Level) Then
        
        Obj.Direct.Y = 0
        Obj.Direct.X = 0
        Obj.Direct.Z = 0

        'slow down the change prediction and check until
        Do While (TestCollision(Obj, Directing, visType, objCollision) = True)
            Obj.Direct.Y = Obj.Direct.Y + adjust
        Loop

        If ((Obj.Direct.Y >= 0) And (Obj.Direct.Y < 0.3)) Or ((Obj.Direct.Y >= 0) And (Obj.Direct.Y <= 0.2)) Then
            Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y 'adjust change and flags to reflect happened
            If Not ((Obj.IsMoving And Moving.Stepping) = Moving.Stepping) Then Obj.IsMoving = Obj.IsMoving + Moving.Stepping
            If ((Obj.IsMoving And Moving.Level) = Moving.Level) Then Obj.IsMoving = Obj.IsMoving - Moving.Level
            newset.Y = Obj.Direct.Y
        End If

    ElseIf ((Obj.IsMoving = Moving.None) And ((tmpset.X = 0 And tmpset.Z = 0) And (newset.X = 0 And newset.Z = 0))) Then

        push = False
        pull = False
        
        Obj.Direct.Y = -adjust
        If Not push Then Obj.Direct.X = adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            pull = True
        Else
            pull = False
            Obj.Direct.Y = 0
            Obj.Direct.X = 0
        End If

        If Not pull Then Obj.Direct.Y = -adjust
        Obj.Direct.Z = adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            push = True
        Else
            push = False
            Obj.Direct.Y = 0
            Obj.Direct.Z = 0
        End If

        If Not pull And Not push Then Obj.Direct.Y = -adjust
        Obj.Direct.X = -adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            pull = (push And Not pull) Or (Not push And Not pull)
        Else
            Obj.Direct.Y = 0
            Obj.Direct.X = 0
        End If

        If Not push And Not pull Then Obj.Direct.Y = -adjust
        Obj.Direct.Z = -adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            push = (pull And Not push) Or (Not push And Not pull)
        Else
            Obj.Direct.Y = 0
            Obj.Direct.Z = 0
        End If


'#####################################################################################
'############# final asjustments made in impressions on self when alone ##############
'#####################################################################################

        'the last check which is to infer movements
        'by the rate of adjust, for steps and steeps
        If (push Xor pull) Or (push And pull) Then

            Obj.Direct.Y = 0

            Do
                Obj.Origin.Y = Obj.Origin.Y - adjust
                If pull Then
                    Obj.Origin.X = Obj.Origin.X + adjust
                    If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                        Obj.Origin.X = Obj.Origin.X - (adjust * 2)
                        If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                            Obj.Origin.Y = Obj.Origin.Y + (adjust / 3)
                        Else
                            Do
                                If Obj.Origin.X + (adjust / 3) <> adjust Then Exit Do
                                Obj.Origin.X = Obj.Origin.X + (adjust / 3)
                            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                            Obj.Origin.X = Obj.Origin.X - (adjust / 3)
                        End If
                    Else
                        Do
                            If Obj.Origin.X - (adjust / 3) <> adjust Then Exit Do
                            Obj.Origin.X = Obj.Origin.X - (adjust / 3)
                        Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                        Obj.Origin.X = Obj.Origin.X + (adjust / 3)
                    End If
                ElseIf push Then

                    Obj.Origin.Z = Obj.Origin.Z + adjust
                    If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                        Obj.Origin.Z = Obj.Origin.Z - (adjust * 2)
                        If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                            Obj.Origin.Y = Obj.Origin.Y + (adjust / 3)
                        Else
                            Do
                                If Obj.Origin.Z + (adjust / 3) <> adjust Then Exit Do
                                Obj.Origin.Z = Obj.Origin.Z + (adjust / 3)
                            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                            Obj.Origin.Z = Obj.Origin.Z - (adjust / 3)
                        End If
                    Else
                        Do
                            If Obj.Origin.Z - (adjust / 3) <> adjust Then Exit Do
                            Obj.Origin.Z = Obj.Origin.Z - (adjust / 3)
                        Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                        Obj.Origin.Z = Obj.Origin.Z + (adjust / 3)
                    End If
                End If

            Loop While (TestCollision(Obj, Directing, visType, objCollision) = True)

        End If
        
    End If


'#####################################################################################
'############# direct activities are primed for Next call to MoveObject  #############
'#####################################################################################


    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
   ' Resume
End Sub

Private Sub SpinObject(ByRef Obj As Element)
On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################

    If Not Obj Is Nothing Then

        If Not TestCollision(Obj, Rotating, 2) Then
                
            Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
            Obj.Rotate.Y = Obj.Rotate.Y + Obj.Twists.Y
            Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
    
            Obj.Rotate = VectorMultiplyBy(AngleAxisRestrict(VectorMultiplyBy(Obj.Rotate, RADIAN)), DEGREE)
            
        End If
        
        Obj.Twists = NoAngle
    
    End If

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub BlowObject(ByRef Obj As Element)
On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################

    If Not Obj Is Nothing Then
    
        If Not TestCollision(Obj, Scaling, 2) Then
        
            Obj.Scaled.X = Obj.Scaled.X + Obj.Scalar.X
            Obj.Scaled.Y = Obj.Scaled.Y + Obj.Scalar.Y
            Obj.Scaled.Z = Obj.Scaled.Z + Obj.Scalar.Z
            
        End If
        
        Obj.Scalar = NoPoint
    
    End If

    
    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function TestCollision(ByRef Obj As Element, ByRef Action As Actions, ByVal visType As Long, Optional ByRef lngCollideObj As Long = -1) As Boolean
On Error GoTo ObjectError


'#####################################################################################
'############# face data is temporary transformed and checked for collision ##########
'#####################################################################################

    If Obj Is Nothing Then Exit Function

    If Action = Rotating And Obj.CollideIndex > -1 Then

'#####################################################################################
'############# in rotation collision we re-adjsut culling view direction #############
'#####################################################################################

        sngCamera(0, 0) = Obj.Origin.X
        sngCamera(0, 1) = Obj.Origin.Y + 1
        sngCamera(0, 2) = Obj.Origin.Z

        sngCamera(1, 0) = 1
        sngCamera(1, 1) = -1
        sngCamera(1, 2) = -1

        sngCamera(2, 0) = -1
        sngCamera(2, 1) = 1
        sngCamera(2, 2) = -1

        If lngFaceCount > 0 Then
            Obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
            lCullCalls = lCullCalls + 1
        End If

    End If


'#####################################################################################
'############# create a transform matrix with the changes applied ####################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim v(2) As D3DVECTOR
    Dim N As D3DVECTOR

    Dim matScale As D3DMATRIX
    Dim matMesh As D3DMATRIX
    Dim matRot As D3DMATRIX
    
    D3DXMatrixIdentity matMesh
    D3DXMatrixIdentity matRot
    D3DXMatrixIdentity matScale

    If Action = Scaling Then
        D3DXMatrixScaling matScale, Obj.Scaled.X + Obj.Scalar.X, Obj.Scaled.Y + Obj.Scalar.Y, Obj.Scaled.Z + Obj.Scalar.Z
    Else
        D3DXMatrixScaling matScale, Obj.Scaled.X, Obj.Scaled.Y, Obj.Scaled.Z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale
    
    If Action = Rotating Then

        D3DXMatrixRotationX matRot, ((Obj.Rotate.X + Obj.Twists.X) * RADIAN)
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
        
        D3DXMatrixRotationY matRot, ((Obj.Rotate.Y + Obj.Twists.Y) * RADIAN)
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
        
        D3DXMatrixRotationZ matRot, ((Obj.Rotate.Z + Obj.Twists.Z) * RADIAN)
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
    Else

        D3DXMatrixRotationX matRot, (Obj.Rotate.X * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationY matRot, (Obj.Rotate.Y * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationZ matRot, (Obj.Rotate.Z * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

    End If

    If Action = Directing Then
        D3DXMatrixTranslation matScale, Obj.Origin.X + Obj.Direct.X, Obj.Origin.Y + Obj.Direct.Y, Obj.Origin.Z + Obj.Direct.Z
    Else
        D3DXMatrixTranslation matScale, Obj.Origin.X, Obj.Origin.Y, Obj.Origin.Z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale
    
            
    If lngFaceCount > 0 And Obj.CollideIndex > -1 And Obj.BoundsIndex > 0 Then
    

'#####################################################################################
'############# update face data with the transformation matrix #######################
'#####################################################################################


        For Face = Obj.CollideIndex To (Obj.CollideIndex + Meshes(Obj.BoundsIndex).Mesh.GetNumFaces) - 1
    
            For cnt = 0 To 2
                
                v(cnt).X = Meshes(Obj.BoundsIndex).Verticies(Index + cnt).X
                v(cnt).Y = Meshes(Obj.BoundsIndex).Verticies(Index + cnt).Y
                v(cnt).Z = Meshes(Obj.BoundsIndex).Verticies(Index + cnt).Z
    
                D3DXVec3TransformCoord v(cnt), v(cnt), matMesh
                
                sngVertexX(cnt, Face) = v(cnt).X
                sngVertexY(cnt, Face) = v(cnt).Y
                sngVertexZ(cnt, Face) = v(cnt).Z

            Next
            
            Index = Index + 3
        Next

'#####################################################################################
'############# per non culled face check and result collision ########################
'#####################################################################################

        Dim lngCollideIdx As Long
        lngCollideIdx = -1
        If Obj.BoundsIndex > 0 Then
            For cnt = Obj.CollideIndex To (Obj.CollideIndex + Meshes(Obj.BoundsIndex).Mesh.GetNumFaces) - 1
                lngTestCalls = lngTestCalls + 1
                lFacesShown = lFacesShown + lngFaceCount
                If lngFaceCount > 0 Then
                    If CBool(Collision(visType, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, cnt, lngCollideObj, lngCollideIdx)) Then
            
                        TestCollision = True
                        GoTo exitfunction
                    End If
                End If
            Next
        End If
    End If
    TestCollision = False

exitfunction:

    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function TestCollisionEx(ByVal FaceNum As Long, ByVal visType As Long, Optional ByRef objCollision As Long = -1, Optional ByRef objFaceIndex As Long = -1) As Boolean
On Error GoTo ObjectError

'#####################################################################################
'############# to the point for simple triangle collsiion checking ###################
'#####################################################################################


    lngTestCalls = lngTestCalls + 1
    lFacesShown = lFacesShown + lngFaceCount
    If lngFaceCount > 0 Then
        If CBool(Collision(visType, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, FaceNum, objCollision, objFaceIndex)) Then
            TestCollisionEx = True
            Exit Function
        End If
    End If

    TestCollisionEx = False

    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function DelCollision(ByRef Obj As Element)
On Error GoTo ObjectError
    Stats_Collision_Count = Stats_Collision_Count - 1
    'Debug.Print "DelCollision"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    If Obj.BoundsIndex > 0 Then
    
        Index = Meshes(Obj.BoundsIndex).Mesh.GetNumFaces
        If lngFaceCount - Index > 0 Then 'Obj.CollideIndex + Index < lngFaceCount Then
    
            For Face = Obj.CollideIndex To lngFaceCount - Index - 1 'Obj.CollideIndex + Index - 1
                sngFaceVis(0, Face) = sngFaceVis(0, Index + Face - 1)
                sngFaceVis(1, Face) = sngFaceVis(1, Index + Face - 1)
                sngFaceVis(2, Face) = sngFaceVis(2, Index + Face - 1)
                sngFaceVis(3, Face) = sngFaceVis(3, Index + Face - 1)
                sngFaceVis(4, Face) = sngFaceVis(4, Index + Face - 1)
                sngFaceVis(5, Face) = sngFaceVis(5, Index + Face - 1)
                
                sngFaceVis(4, Face) = sngFaceVis(4, Face) - 1
                sngFaceVis(5, Face) = sngFaceVis(5, Face) - Index
                
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
            
            Dim e1 As Element
            For Each e1 In Elements
            'For cnt = 1 To Elements.count
                If e1.CollideIndex > Obj.CollideIndex Then
                    e1.CollideIndex = e1.CollideIndex - Index
                End If
            Next
            
    '        If Obj.CollideIndex + Index < lngFaceCount - 2 Then
    '
    '            For Face = Obj.CollideIndex + Index To lngFaceCount - 2
    '                sngFaceVis(0, Face) = sngFaceVis(0, Face + 1)
    '                sngFaceVis(1, Face) = sngFaceVis(1, Face + 1)
    '                sngFaceVis(2, Face) = sngFaceVis(2, Face + 1)
    '                sngFaceVis(3, Face) = sngFaceVis(3, Face + 1)
    '                sngFaceVis(4, Face) = sngFaceVis(4, Face + 1)
    '                sngFaceVis(5, Face) = sngFaceVis(5, Face + 1)
    '                sngVertexX(0, Face) = sngVertexX(0, Face + 1)
    '                sngVertexX(1, Face) = sngVertexX(1, Face + 1)
    '                sngVertexX(2, Face) = sngVertexX(2, Face + 1)
    '                sngVertexY(0, Face) = sngVertexY(0, Face + 1)
    '                sngVertexY(1, Face) = sngVertexY(1, Face + 1)
    '                sngVertexY(2, Face) = sngVertexY(2, Face + 1)
    '                sngVertexZ(0, Face) = sngVertexZ(0, Face + 1)
    '                sngVertexZ(1, Face) = sngVertexZ(1, Face + 1)
    '                sngVertexZ(2, Face) = sngVertexZ(2, Face + 1)
    '
    '                sngScreenX(0, Face) = sngScreenX(0, Face + 1)
    '                sngScreenX(1, Face) = sngScreenX(1, Face + 1)
    '                sngScreenX(2, Face) = sngScreenX(2, Face + 1)
    '                sngScreenY(0, Face) = sngScreenY(0, Face + 1)
    '                sngScreenY(1, Face) = sngScreenY(1, Face + 1)
    '                sngScreenY(2, Face) = sngScreenY(2, Face + 1)
    '                sngScreenZ(0, Face) = sngScreenZ(0, Face + 1)
    '                sngScreenZ(1, Face) = sngScreenZ(1, Face + 1)
    '                sngScreenZ(2, Face) = sngScreenZ(2, Face + 1)
    '
    '                sngZBuffer(0, Face) = sngZBuffer(0, Face + 1)
    '                sngZBuffer(1, Face) = sngZBuffer(1, Face + 1)
    '                sngZBuffer(2, Face) = sngZBuffer(2, Face + 1)
    '                sngZBuffer(3, Face) = sngZBuffer(3, Face + 1)
    '            Next
    '        End If
            
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
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function



Public Function DelCollisionEx(ByRef CollideIndex As Long, ByVal NumFaces As Long)
On Error GoTo ObjectError
    Stats_CollisionEx_Count = Stats_CollisionEx_Count - 1
    'Debug.Print "DelCollisionEx"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    Index = NumFaces
    
    If lngFaceCount - Index > 0 Then 'Obj.CollideIndex + Index < lngFaceCount Then

        For Face = CollideIndex To lngFaceCount - Index - 1 'Obj.CollideIndex + Index - 1
'    If CollideIndex + Index < lngFaceCount Then
'
'        For Face = CollideIndex To CollideIndex + Index - 1
            sngFaceVis(0, Face) = sngFaceVis(0, Index + Face)
            sngFaceVis(1, Face) = sngFaceVis(1, Index + Face)
            sngFaceVis(2, Face) = sngFaceVis(2, Index + Face)
            sngFaceVis(3, Face) = sngFaceVis(3, Index + Face)
            sngFaceVis(4, Face) = sngFaceVis(4, Index + Face)
            sngFaceVis(5, Face) = sngFaceVis(5, Index + Face)
            
            sngFaceVis(4, Face) = sngFaceVis(4, Face) - 1
            sngFaceVis(5, Face) = sngFaceVis(5, Face) - Index
            
            sngVertexX(0, Face) = sngVertexX(0, Index + Face)
            sngVertexX(1, Face) = sngVertexX(1, Index + Face)
            sngVertexX(2, Face) = sngVertexX(2, Index + Face)
            sngVertexY(0, Face) = sngVertexY(0, Index + Face)
            sngVertexY(1, Face) = sngVertexY(1, Index + Face)
            sngVertexY(2, Face) = sngVertexY(2, Index + Face)
            sngVertexZ(0, Face) = sngVertexZ(0, Index + Face)
            sngVertexZ(1, Face) = sngVertexZ(1, Index + Face)
            sngVertexZ(2, Face) = sngVertexZ(2, Index + Face)
            
            sngScreenX(0, Face) = sngScreenX(0, Index + Face)
            sngScreenX(1, Face) = sngScreenX(1, Index + Face)
            sngScreenX(2, Face) = sngScreenX(2, Index + Face)
            sngScreenY(0, Face) = sngScreenY(0, Index + Face)
            sngScreenY(1, Face) = sngScreenY(1, Index + Face)
            sngScreenY(2, Face) = sngScreenY(2, Index + Face)
            sngScreenZ(0, Face) = sngScreenZ(0, Index + Face)
            sngScreenZ(1, Face) = sngScreenZ(1, Index + Face)
            sngScreenZ(2, Face) = sngScreenZ(2, Index + Face)
            
            sngZBuffer(0, Face) = sngZBuffer(0, Index + Face)
            sngZBuffer(1, Face) = sngZBuffer(1, Index + Face)
            sngZBuffer(2, Face) = sngZBuffer(2, Index + Face)
            sngZBuffer(3, Face) = sngZBuffer(3, Index + Face)
            
        Next
        
        Dim e1 As Element
        For Each e1 In Elements
        'For cnt = 1 To Elements.Count
            If e1.CollideIndex > CollideIndex Then
                e1.CollideIndex = e1.CollideIndex - Index
            End If
        Next
        
'        If CollideIndex + Index < lngFaceCount - 2 Then
'
'            For Face = CollideIndex + Index To lngFaceCount - 2
'                sngFaceVis(0, Face) = sngFaceVis(0, Face + 1)
'                sngFaceVis(1, Face) = sngFaceVis(1, Face + 1)
'                sngFaceVis(2, Face) = sngFaceVis(2, Face + 1)
'                sngFaceVis(3, Face) = sngFaceVis(3, Face + 1)
'                sngFaceVis(4, Face) = sngFaceVis(4, Face + 1)
'                sngFaceVis(5, Face) = sngFaceVis(5, Face + 1)
'                sngVertexX(0, Face) = sngVertexX(0, Face + 1)
'                sngVertexX(1, Face) = sngVertexX(1, Face + 1)
'                sngVertexX(2, Face) = sngVertexX(2, Face + 1)
'                sngVertexY(0, Face) = sngVertexY(0, Face + 1)
'                sngVertexY(1, Face) = sngVertexY(1, Face + 1)
'                sngVertexY(2, Face) = sngVertexY(2, Face + 1)
'                sngVertexZ(0, Face) = sngVertexZ(0, Face + 1)
'                sngVertexZ(1, Face) = sngVertexZ(1, Face + 1)
'                sngVertexZ(2, Face) = sngVertexZ(2, Face + 1)
'
'                sngScreenX(0, Face) = sngScreenX(0, Face + 1)
'                sngScreenX(1, Face) = sngScreenX(1, Face + 1)
'                sngScreenX(2, Face) = sngScreenX(2, Face + 1)
'                sngScreenY(0, Face) = sngScreenY(0, Face + 1)
'                sngScreenY(1, Face) = sngScreenY(1, Face + 1)
'                sngScreenY(2, Face) = sngScreenY(2, Face + 1)
'                sngScreenZ(0, Face) = sngScreenZ(0, Face + 1)
'                sngScreenZ(1, Face) = sngScreenZ(1, Face + 1)
'                sngScreenZ(2, Face) = sngScreenZ(2, Face + 1)
'
'                sngZBuffer(0, Face) = sngZBuffer(0, Face + 1)
'                sngZBuffer(1, Face) = sngZBuffer(1, Face + 1)
'                sngZBuffer(2, Face) = sngZBuffer(2, Face + 1)
'                sngZBuffer(3, Face) = sngZBuffer(3, Face + 1)
'            Next
'        End If
        
    End If
    
    CollideIndex = -1
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
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function AddCollision(ByRef Obj As Element, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError
    Stats_Collision_Count = Stats_Collision_Count + 1
'#####################################################################################
'############# create face data for a mesh to external compatability #################
'#####################################################################################
    'Debug.Print "AddCollision"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    Dim v() As D3DVECTOR

    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    Dim vn As D3DVECTOR

    ReDim v(0 To 3) As D3DVECTOR

    If Obj.BoundsIndex > 0 Then
        Obj.CollideIndex = lngFaceCount
        AddCollision = lngFaceCount
    
        Dim FaceCount As Long
        Dim addingFace As Boolean
        
        
        Index = 0
        For Face = 0 To Meshes(Obj.BoundsIndex).Mesh.GetNumFaces - 1
    
            For cnt = 0 To 2
    
                v(cnt).X = Meshes(Obj.BoundsIndex).Verticies(Meshes(Obj.BoundsIndex).Indicies(Index + cnt)).X
                v(cnt).Y = Meshes(Obj.BoundsIndex).Verticies(Meshes(Obj.BoundsIndex).Indicies(Index + cnt)).Y
                v(cnt).Z = Meshes(Obj.BoundsIndex).Verticies(Meshes(Obj.BoundsIndex).Indicies(Index + cnt)).Z
    
                'D3DXVec3TransformCoord vn, v(cnt), matObject
                vn = ToVector(Obj.PointMatrix(ToPoint(v(cnt))))
                
                v(cnt).X = vn.X
                v(cnt).Y = vn.Y
                v(cnt).Z = vn.Z
            Next
    
            ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
            ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
            ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
            ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single
    
            ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
            ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
            ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single
    
            ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
            
            vn = TriangleNormal(v(0), v(1), v(2))
            
            For cnt = 0 To 2
    
                sngVertexX(cnt, lngFaceCount) = v(cnt).X
                sngVertexY(cnt, lngFaceCount) = v(cnt).Y
                sngVertexZ(cnt, lngFaceCount) = v(cnt).Z
    
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
    End If
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function AddCollisionEx(ByRef Verticies() As D3DVECTOR, ByVal NumFaces As Long, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError
    Stats_CollisionEx_Count = Stats_CollisionEx_Count + 1
    'Debug.Print "AddCollisionEx"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim v() As D3DVECTOR

    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    Dim vn As D3DVECTOR

    ReDim v(0 To 3) As D3DVECTOR

    AddCollisionEx = lngFaceCount

    Dim FaceCount As Long
    Dim addingFace As Boolean
    
    Index = 0
    For Face = 0 To NumFaces - 1

        For cnt = 0 To 2
            
            v(cnt).X = Verticies(Index + cnt).X
            v(cnt).Y = Verticies(Index + cnt).Y
            v(cnt).Z = Verticies(Index + cnt).Z
                        
        Next
        
        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single
    
        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
        
        vn = TriangleNormal(v(0), v(1), v(2))

        For cnt = 0 To 2
            
            sngVertexX(cnt, lngFaceCount) = v(cnt).X
            sngVertexY(cnt, lngFaceCount) = v(cnt).Y
            sngVertexZ(cnt, lngFaceCount) = v(cnt).Z

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
    
    lngObjCount = lngObjCount + 1
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub RenderPortals()
     
    If Portals.Count > 0 Then
        Dim t1 As Portal
        
        For Each t1 In Portals
            
            If t1.Enabled Then
            
                RenderPortals2 t1, Player
                
                If Elements.Count > 0 Then
                    Dim e1 As Element
                    
                    For Each e1 In Elements

                        RenderPortals2 t1, e1

                    Next
                End If
                
            End If
        Next
    End If

    Exit Sub

End Sub

Private Sub RenderPortals2(ByRef t1 As Portal, ByRef e1 As Element)
On Error GoTo scripterror

       
    Dim pos As D3DVECTOR
    
    Dim cnt3 As Long
    Dim cnt2 As Long
    
    Dim cnt As Long
    Dim Obj As Long
                            
    Dim A As Long
    Dim act As Motion
    Dim txtobj As String
    Dim errline As Long
    Dim errsource As String

    Dim e2 As Element
        
    If (DistanceEx(e1.Origin, t1.Location) <= t1.Range) Then

        If Not ((t1.Teleport.X = 0) And (t1.Teleport.Y = 0) And (t1.Teleport.Z = 0)) Then
            pos = ToVector(e1.Origin)
            If Elements.Count > 0 Then
                For Each e2 In Elements
                    If e1.Collision Then
                        If e2.CollideIndex > -1 Then
                            If Not e2.CollideIndex = e1.CollideIndex And e2.Gravitational And e2.BoundsIndex > 0 Then
                                For cnt3 = e2.CollideIndex To (e2.CollideIndex + Meshes(e2.BoundsIndex).Mesh.GetNumFaces) - 1
                                    sngFaceVis(3, cnt3) = 1
                                Next
                            ElseIf e2.CollideIndex = e1.CollideIndex And e2.BoundsIndex > 0 Then
                                For cnt3 = e2.CollideIndex To (e2.CollideIndex + Meshes(e2.BoundsIndex).Mesh.GetNumFaces) - 1
                                    sngFaceVis(3, cnt3) = 1
                                Next
                            ElseIf e2.BoundsIndex > 0 Then
                                For cnt3 = e2.CollideIndex To (e2.CollideIndex + Meshes(e2.BoundsIndex).Mesh.GetNumFaces) - 1
                                    sngFaceVis(3, cnt3) = 0
                                Next
                            End If
                        End If
                    End If
                Next
            End If
            e1.Origin = t1.Teleport
            If e1.Collision Then
                If TestCollision(e1, Actions.None, 1) Then
                    Set e1.Origin = ToPoint(pos)
                Else
                    e1.ClearMotions
                End If
            End If
        End If
        If t1.DropsMotions Then
            e1.ClearMotions
        End If
        
        If t1.Motions.Count > 0 Then
            For A = 1 To t1.Motions.Count
                Set act = t1.Motions(A)
                e1.AddMotion act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.Script
            Next
        End If

        
        If Not t1.OnInRange Is Nothing Then
        
            If InStr(t1.OnInRange.AppliesTo & ",", e1.Key & ",") > 0 Or t1.OnInRange.AppliesTo = "" Then
            
                If Not t1.OnInRange.RunFlag Then
                    t1.OnInRange.RunFlag = True
                    errsource = "OnInRange"
                    errline = CLng(t1.OnInRange.StartLine)
                    
                    frmMain.Run t1.OnInRange.RunMethod, e1.Key, errline
                    Debug.Print "OnInRange " & t1.Key & " " & e1.Key
                End If
                
                If Not t1.OnOutRange Is Nothing Then
                    t1.OnOutRange.RunFlag = False
                End If
                
            End If
            
        End If
        

        

    Else
        If Not t1.OnOutRange Is Nothing Then
        
            If InStr(t1.OnOutRange.AppliesTo & ",", e1.Key & ",") > 0 Or t1.OnOutRange.AppliesTo = "" Then

                If Not t1.OnOutRange.RunFlag Then
                    t1.OnOutRange.RunFlag = True
                    errsource = "OnOutRange"
                    errline = CLng(t1.OnOutRange.StartLine)
                    frmMain.Run t1.OnOutRange.RunMethod, e1.Key, errline
                    Debug.Print "OnOutRange " & t1.Key & " " & e1.Key
                                        
                End If
               
                If Not t1.OnInRange Is Nothing Then
                    t1.OnInRange.RunFlag = False
                End If
            End If
            
        End If
        

        

    End If

scripterror:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    If Err.Number <> 0 Then
        DoEvents
    
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & errsource & _
        "\nIn the event starting at Line: " & errline & _
        "\nError: " & Err.Description
        'frmMain.Print "echo An error " & Err.Number & " occurd in " & Err.Source & " at line " & (atLine - 1) & "\n" & Err.Description & "\n" & LastCall
        
        If frmMain.ScriptControl1.Error.Number <> 0 Then frmMain.ScriptControl1.Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End If

End Sub
Private Sub ExecuteScript(ByRef e1 As Element, ByVal EventText As String)
    Dim trig As String
    Dim line As String
    Dim id As String
        
    line = NextArg(EventText, ":")
    trig = RemoveArg(EventText, ":")
    If Left(Trim(trig), 1) = "<" Then
        id = RemoveQuotedArg(trig, "<", ">") & ","
        If ((InStr(id, e1.Key & ",") > 0) And (e1.Key <> "")) Or (id = ",") Then
            If IsNumeric(line) And trig <> "" Then
                frmMain.ExecuteStatement trig, e1.Key, CLng(line)
            ElseIf trig <> "" Then
                frmMain.ExecuteStatement trig, e1.Key
            Else
                frmMain.ExecuteStatement trig, line
            End If
        End If
    Else
        If IsNumeric(line) And trig <> "" Then
            frmMain.ExecuteStatement trig, e1.Key, CLng(line)
        ElseIf trig <> "" Then
            frmMain.ExecuteStatement trig, e1.Key
        Else
            frmMain.ExecuteStatement trig, line
        End If
    End If

End Sub


Private Function GetClosestCamera(Optional ByVal Exclude As String = "") As Long

    Dim cnt As Long
    Dim Dist As Single
    Dim past As Single
    If Cameras.Count > 0 Then
        Static toggle As Boolean
        toggle = Not toggle
        For cnt = IIf(toggle, 1, Cameras.Count) To IIf(toggle, Cameras.Count, 1) Step IIf(toggle, 1, -1)
            Dist = DistanceEx(Player.Origin, Cameras(cnt).Origin)
            If ((Dist <= past) Or (past = 0)) And (InStr(Exclude, cnt & ",") = 0) Then
                GetClosestCamera = cnt
                past = Dist
            End If
        Next
    End If
End Function

Public Sub RenderCameras()
On Error GoTo CameraError

    Dim cnt As Long
    Dim cnt2 As Long
    Dim Dist As Single
    Dim past As Long
    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    
    
    Dim pos As D3DVECTOR
    Dim touched As Boolean
    Dim Face As Long
    Dim ex As String
    
    Dim dot As Single
    Dim v As D3DVECTOR
    Dim N As D3DVECTOR
    
    Dim verts(0 To 2) As D3DVECTOR
    Dim lastCam As Long
    'two quests about cameras
    '1 default projection should be in short range leainant not to turning camera around rather to a any range put projection variance in direction
    '2 movement from one camera to the Next could have a flying adaptation in a swing and out of the constructs way while it flies to genral Next 1
        
    If Perspective = Playmode.CameraMode Then
    
        If Cameras.Count > 0 Then
            
            If (Elements.Count > 0) Then
                Dim e1 As Element
                For Each e1 In Elements
                
                'For cnt = 1 To Elements.count
                    If ((e1.Effect = Collides.Ground) Or (e1.Effect = Collides.InDoor)) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                        For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                            sngFaceVis(3, cnt2) = 1
                        Next
                    ElseIf (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                        For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                            sngFaceVis(3, cnt2) = 0
                        Next
                    End If
                Next
            End If

            cnt = 0
            Player.CameraIndex = 0

            Do
                
                cnt = GetClosestCamera(ex)
                
                touched = False
                        
                If (cnt > 0) Then
                    With Cameras(cnt)
                    
                        verts(0) = ToVector(Player.Origin)
                        verts(1) = VectorAdd(ToVector(Player.Origin), MakeVector(0, -0.01, 0))
                        verts(2) = ToVector(.Origin)
    
                        Face = AddCollisionEx(verts, 1)
                        touched = TestCollisionEx(Face, 1)
                        DelCollisionEx Face, 1
    
                        If (ClassifyPoint(V1, V1, V1, ToVector(Player.Origin)) = 1) Then touched = True
    
    
                        If Not touched Then
                            
                            
                            V1 = VectorSubtract(MakeVector(.Origin.X + Sin(D720 - .Angle), _
                                                                            .Origin.Y - Tan(D720 - .Pitch), _
                                                                            .Origin.Z + Cos(D720 - .Angle)), _
                                                                            ToVector(.Origin))
                                                                            
                            V2 = VectorSubtract(MakeVector(Player.Origin.X - Sin(D720 - .Angle), _
                                                            Player.Origin.Y + Tan(D720 - .Pitch), _
                                                            Player.Origin.Z - Cos(D720 - .Angle)), _
                                                            ToVector(.Origin))
                            
                            If ((V2.X > 0 And V1.X > 0) Or (V2.X < 0 And V1.X < 0)) And _
                                ((V2.Y > 0 And V1.Y > 0) Or (V2.Y < 0 And V1.Y < 0)) And _
                                ((V2.Z > 0 And V1.Z > 0) Or (V2.Z < 0 And V1.Z < 0)) Then
                                touched = False
                                
                                If past <> 0 Then
                                    If DistanceEx(.Origin, Player.Origin) > Dist Then
                                        cnt = past
                                        Dist = DistanceEx(.Origin, Player.Origin)
                                    End If
                                End If
    
                            Else
                                touched = True
                            End If
                            If Not touched Then
    
                                dot = modDecs.VectorDotProduct(V1, V2) / (modDecs.VectorDotProduct(V1, V1) * modDecs.VectorDotProduct(V2, V2))
                            End If
                        End If
                        
                        If Not touched Then
                            If past <> 0 Then
                                If DistanceEx(.Origin, Player.Origin) > Dist Then
                                    cnt = past
                                    Dist = DistanceEx(.Origin, Player.Origin)
                                    ex = ex & cnt & ", "
                                End If
                            End If
    
                            If cnt >= 0 And cnt <= Cameras.Count Then
                                Player.CameraIndex = cnt
                                past = cnt
                                Dist = DistanceEx(.Origin, Player.Origin)
                            End If
                        Else
                            ex = ex & cnt & ", "
                        End If
                    End With
                End If

            Loop Until (cnt = 0) Or (Player.CameraIndex <> 0)
            
            If Player.CameraIndex = 0 And Not lastCam = 0 Then
                Player.CameraIndex = lastCam
            End If
            lastCam = Player.CameraIndex
        
        End If
        
    ElseIf (Not (Player.CameraIndex = 0)) Then
        If Not ((Perspective = Spectator) Or DebugMode) Then
            Player.CameraIndex = 0
        End If
    End If
    
    Exit Sub
CameraError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SortVerticies(ByVal FaceIndex As Long, Optional ByVal VertexCount As Long = 3)
    Dim A As D3DVECTOR
    Dim B As D3DVECTOR
    Dim C As D3DVECTOR
    
    Dim p As D3DVECTOR
    
    Dim cnt As Long
    Dim Angle As Single
    
    Dim smallest As Long
    Dim smallestAngle As Single
    Dim v() As D3DVECTOR
    ReDim v(0 To VertexCount - 1) As D3DVECTOR

    For cnt = 0 To VertexCount - 1
        v(cnt) = MakeVector(sngVertexX(cnt, FaceIndex), sngVertexY(cnt, FaceIndex), sngVertexZ(cnt, FaceIndex))
        C.X = C.X + v(cnt).X
        C.Y = C.Y + v(cnt).Y
        C.Z = C.Z + v(cnt).Z
    Next
    
    C.X = C.X / VertexCount
    C.Y = C.Y / VertexCount
    C.Z = C.Z / VertexCount

    p = GetPlaneNormal(v(0), v(1), v(2))
        
    Dim N As Long
    Dim m As Long
    
    For N = 0 To VertexCount - 1
        
        A = modDecs.VectorNormalize(modDecs.VectorSubtract(v(N), C))
        
        smallest = -1
        smallestAngle = -1
        
        For m = N + 1 To 2
            If Not ClassifyPoint(v(N), C, VectorAdd(C, p), v(m)) = 2 Then 'not back
                B = modDecs.VectorNormalize(modDecs.VectorSubtract(v(m), C))
                
                Angle = modDecs.VectorDotProduct(A, B)
                
                If Angle > smallestAngle Then
                    smallestAngle = Angle
                    smallest = m
        
                End If
            End If
        Next
        
        If smallest = -1 Then Exit Sub
        
        If Not ((N + 1) = smallest) Then
            SwapVector FaceIndex, N + 1, smallest
        End If
    
    Next
    
    A = GetPlaneNormal(v(0), v(1), v(2))
    B = p
    
    If modDecs.VectorDotProduct(A, B) < 0 Then
        ReverseFaceVertices FaceIndex, VertexCount
    End If
    
    sngFaceVis(0, FaceIndex) = A.X
    sngFaceVis(1, FaceIndex) = A.Y
    sngFaceVis(2, FaceIndex) = A.Z

End Sub

Public Function GetPlaneNormal(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR) As D3DVECTOR

    Dim vector1 As D3DVECTOR
    Dim vector2 As D3DVECTOR
    Dim Normal As D3DVECTOR
    Dim Length As Single

    '/*Calculate the Normal*/
    '/*Vector 1*/
    vector1.X = (v0.X - V1.X)
    vector1.Y = (v0.Y - V1.Y)
    vector1.Z = (v0.Z - V1.Z)

    '/*Vector 2*/
    vector2.X = (V1.X - V2.X)
    vector2.Y = (V1.Y - V2.Y)
    vector2.Z = (V1.Z - V2.Z)

    '/*Apply the Cross Product*/
    Normal.X = (vector1.Y * vector2.Z - vector1.Z * vector2.Y)
    Normal.Y = (vector1.Z * vector2.X - vector1.X * vector2.Z)
    Normal.Z = (vector1.X * vector2.Y - vector1.Y * vector2.X)

    '/*Normalize to a unit vector*/
    Length = Sqr(Normal.X * Normal.X + Normal.Y * Normal.Y + Normal.Z * Normal.Z)

    If Length = 0 Then Length = 1

    Normal.X = (Normal.X / Length)
    Normal.Y = (Normal.Y / Length)
    Normal.Z = (Normal.Z / Length)

    GetPlaneNormal = Normal
End Function

Public Function ReverseFaceVertices(ByVal FaceIndex As Long, ByVal VertexCount As Long)

    Dim cnt As Long
    For cnt = 0 To (VertexCount \ 2)
        SwapVector FaceIndex, cnt, (VertexCount - 1) - cnt
        
    Next

End Function

Public Sub SwapVector(ByVal FaceIndex As Long, ByVal FirstIndex As Long, ByVal SecondIndex As Long)
    Dim v As D3DVECTOR
    v.X = sngVertexX(FirstIndex, FaceIndex)
    v.Y = sngVertexY(FirstIndex, FaceIndex)
    v.Z = sngVertexZ(FirstIndex, FaceIndex)
    
    sngVertexX(FirstIndex, FaceIndex) = sngVertexX(SecondIndex, FaceIndex)
    sngVertexY(FirstIndex, FaceIndex) = sngVertexY(SecondIndex, FaceIndex)
    sngVertexZ(FirstIndex, FaceIndex) = sngVertexZ(SecondIndex, FaceIndex)

    sngVertexX(SecondIndex, FaceIndex) = v.X
    sngVertexY(SecondIndex, FaceIndex) = v.Y
    sngVertexZ(SecondIndex, FaceIndex) = v.Z
End Sub





