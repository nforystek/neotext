Attribute VB_Name = "modMove"
#Const modMove = -1
Option Explicit
'TOP DoWN
Option Compare Binary

Option Private Module

Public Enum CameraCollision
    CameraTop = 0
    CameraBack = 1
    CameraLeft = 2
    CameraFront = 3
    CameraRight = 4
    CameraBottom = 5
End Enum


Public Type MyCulling
    Position As D3DVECTOR
    Direction As D3DVECTOR
    UpVector As D3DVECTOR
    visType As Long
End Type


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
                  
                  
'###########################################################################
'###################### BEGIN UNIQUE NON GLOBALS ###########################
'###########################################################################
                  
                  
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

'Public DebugFace() As MyVertex
'Public DebugSkin(0 To 4) As Direct3DTexture8
'Public DebugVBuf As Direct3DVertexBuffer8

Public CullingSetup As Integer
Public CullingObject As MyCulling
Public CullingCount As Long
Public Cullings() As MyCulling


Public Sub CreateMove()

    ReDim sngCamera(0 To 2, 0 To 2) As Single
    
'    Set DebugSkin(0) = LoadTexture(AppPath & "Models\debug0.bmp")
'    Set DebugSkin(1) = LoadTexture(AppPath & "Models\debug1.bmp")
'    Set DebugSkin(2) = LoadTexture(AppPath & "Models\debug2.bmp")
'    Set DebugSkin(3) = LoadTexture(AppPath & "Models\debug4.bmp")
'    Set DebugSkin(4) = LoadTexture(AppPath & "Models\debug3.bmp")
    
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
        vn = modDecs.TriangleNormal(MakeVector(sngVertexX(0, cnt), sngVertexY(0, cnt), sngVertexZ(0, cnt)), _
                            MakeVector(sngVertexX(1, cnt), sngVertexY(1, cnt), sngVertexZ(1, cnt)), _
                            MakeVector(sngVertexX(2, cnt), sngVertexY(2, cnt), sngVertexZ(2, cnt)))
        sngFaceVis(0, cnt) = vn.X
        sngFaceVis(1, cnt) = vn.Y
        sngFaceVis(2, cnt) = vn.Z
    Next
    
    
End Sub


Public Function SetMotion(ByRef act As Motion, ByRef Action As Actions, ByRef dat As Point, ByRef emp As Single) As String
    act.Key = Replace(modGuid.GUID, "-", "K")
    act.Action = Action
    Set act.Data = dat
    act.Emphasis = emp
    SetMotion = act.Key
End Function

Public Function AddMotion(ByRef Motions As NTNodes10.Collection, ByRef Action As Long, ByVal Key As String, ByRef Data As Point, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String

    Dim act As New Motion
    With act
        If Key = "" Then
            .Key = Replace(modGuid.GUID, "-", "K")
        Else
            .Key = Key
        End If
        .Action = Action
        Set .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .latency = Timer
        .Recount = Recount
        .Script = Script
        AddMotion = .Key
    End With
    If Motions Is Nothing Then
        Set Motions = New NTNodes10.Collection
    End If
    If Motions.Count > 0 Then
        Motions.Add act, , 1
    Else
        Motions.Add act
    End If
    
'    If Not pAttachments Is Nothing Then
'        If pAttachments.Count > 0 Then
'            Dim e2 As Element
'            For Each e2 In pAttachments
'                If e2.Motions.Count > 0 Then
'                    e2.Motions.Add act, , 1
'                Else
'                    e2.Motions.Add act
'                End If
'            Next
'
'        End If
'    End If
    
   ' Set act = Nothing

End Function

Public Function DeleteMotion(ByRef Motions As NTNodes10.Collection, ByVal Key As String) As Boolean

    If Not Motions Is Nothing Then
        Dim A As Long
        Dim act As Motion
        A = 1
        Do While A <= Motions.Count
            Set act = Motions(A)
            If act.Key = Key Or (act.Key = "") Then
                Motions.Remove A
                DeleteMotion = True
            Else
                A = A + 1
            End If
            Set act = Nothing
        Loop
    End If
'a prior way that only removed one occurance of the idenity in the collection
'only it started with checking the very last item,and removed it if was idenity
'otherwise it begane a forward iteration, as well it did not check for blank id's
'
'this one is slightly different, it only delete's one occurance and starts from
'the back of the list, iterating all, it also considers blank ID match for removal
'
'    A = Motions.count
'    Do While A > 0
'        Set act = Motions(A)
'        If act.Identity = MGUID Or (act.Identity = "") Then
'            Motions.Remove A
'            DeleteMotion = True
'            Set act = Nothing
'            Exit Function
'        End If
'        A = A - 1
'    Loop
'    Set act = Nothing
'
'
'because of array differences the above was working, yet no longer in objects
'so the choice is a full iteration of the collection from the start and blanks
'are considered matches for removal as well as ID matching, and all matches are
'removed if satisfying those conditions, now it acts like before, with out
'stacking too many motions that do or don't remove bogging down the system

End Function

Public Sub ClearMotions(ByRef Motions As NTNodes10.Collection)
    If Not Motions Is Nothing Then
        Dim act As Motion
        Do While Motions.Count > 0
            Set act = Motions(1)
            Motions.Remove 1
            Set act = Nothing
        Loop
    End If
End Sub

Public Function MotionExists(ByRef Motions As NTNodes10.Collection, ByVal Key As String) As Boolean
    If Not Motions Is Nothing Then
        Dim A As Long
        For A = 1 To Motions.Count
            If Motions(A).Key = Key Then
                MotionExists = True
                Exit Function
            End If
        Next
    End If
    MotionExists = False
End Function

Public Function ValidMotion(ByRef Motion As Motion) As Boolean
    ValidMotion = (Motion.Key <> "")
End Function

Public Function CalculateMotion(ByRef Motion As Motion, ByRef Action As Actions) As D3DVECTOR

    If (Action And Motion.Action) = Action Then

        If Motion.Friction <> 0 Then

            Motion.Emphasis = Motion.Emphasis - (Motion.Emphasis * Motion.Friction)
            If Motion.Emphasis < 0 Then
                Motion.Emphasis = 0
                Motion.Key = ""
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

Private Sub ApplyMotion(ByRef obj As Element, ByVal Action As Actions)
    Dim cnt As Long
    Dim cnt2 As Long
    Dim Offset As D3DVECTOR
    Dim vout As D3DVECTOR
    
    If Not modParse.Player.Element Is Nothing Then
        'gravity
        
        If ((Not (Perspective = Spectator)) And (obj.CollideObject = modParse.Player.Element.CollideObject)) Or (Not (obj.CollideObject = modParse.Player.Element.CollideObject)) Then
            
            If obj.Gravitational Then
                If Not obj.OnLadder Then
                    If obj.InLiquid Then
                        Select Case Action
                            Case (Action And Directing)
                                D3DXVec3Add vout, ToVector(obj.Direct), CalculateMotion(LiquidGravityDirect, Directing)
                                Set obj.Direct = ToPoint(vout)
                            Case (Action And Rotating)
                                D3DXVec3Add vout, ToVector(obj.Twists), CalculateMotion(LiquidGravityRotate, Rotating)
                                Set obj.Twists = ToPoint(vout)
                            Case (Action And Scaling)
                                D3DXVec3Add vout, ToVector(obj.Scalar), CalculateMotion(LiquidGravityScaled, Scaling)
                                Set obj.Scalar = ToPoint(vout)
                        End Select
                    Else
                        Select Case Action
                            Case (Action And Directing)
                                D3DXVec3Add vout, ToVector(obj.Direct), CalculateMotion(GlobalGravityDirect, Directing)
                                Set obj.Direct = ToPoint(vout)
                            Case (Action And Rotating)
                                D3DXVec3Add vout, ToVector(obj.Twists), CalculateMotion(GlobalGravityRotate, Rotating)
                                Set obj.Twists = ToPoint(vout)
                            Case (Action And Scaling)
                                D3DXVec3Add vout, ToVector(obj.Scalar), CalculateMotion(GlobalGravityScaled, Scaling)
                                Set obj.Scalar = ToPoint(vout)
                        End Select
                    End If
                End If
            End If
        End If
    End If
    
    
    If obj.Effect = Collides.Normal Then
        If Not obj.Motions Is Nothing Then
            If obj.Motions.Count > 0 Then
                Dim A As Long
                For A = 1 To obj.Motions.Count
                    If ValidMotion(obj.Motions(A)) Then
                    
                        Select Case Action
                            Case (Action And Directing)
                                D3DXVec3Add vout, ToVector(obj.Direct), CalculateMotion(obj.Motions(A), Directing)
                                Set obj.Direct = ToPoint(vout)
                            Case (Action And Rotating)
                                D3DXVec3Add vout, ToVector(obj.Twists), CalculateMotion(obj.Motions(A), Rotating)
                                Set obj.Twists = ToPoint(vout)
                            Case (Action And Scaling)
                                D3DXVec3Add vout, ToVector(obj.Scalar), CalculateMotion(obj.Motions(A), Scaling)
                                Set obj.Scalar = ToPoint(vout)
                        End Select
                    End If
                Next
            End If
        End If
    End If
End Sub
Public Sub ResetMotion()
    Dim A As Long
    Dim o As Long
   ' If Not modParse.Player.Element Is Nothing Then Set modParse.Player.Element.Direct = MakePoint(0, 0, 0)
    If Elements.Count > 0 Then
        For o = 1 To Elements.Count
           Set Elements(o).Direct = MakePoint(0, 0, 0)
        Next
    End If
End Sub

Public Sub RenderMotion()
On Error GoTo ObjectError

    'If Not modParse.Player.Element Is Nothing Then RenderMotion2 modParse.Player.Element
    
    Dim cnt As Long
    cnt = 1
    Do While cnt <= Elements.Count
    
        RenderMotion2 Elements(cnt)
        cnt = cnt + 1
    Loop


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
    
        ApplyMotion e1, Directing And Scaling
    
        If Not e1.Motions Is Nothing Then
            If e1.Motions.Count > 0 Then
                A = 1
                Do While A <= e1.Motions.Count
                    Set act = e1.Motions(A)
    
                    If act.Reactive > -1 Then
                        If (Timer - act.latency) > act.Reactive Then
                            act.latency = Timer
                            
                            act.Emphasis = act.Initials
                            e1.DeleteMotion act.Key
                            If Not act.Script = "" Then
                                
                                ExecuteScript e1, act.Script
                            
                            End If
                            If act.Recount > -1 Then
                                If act.Recount > 0 Then
                                    act.Recount = act.Recount - 1
                                    e1.AddMotion act.Action, act.Key, act.Data, act.Initials, act.Friction, act.Reactive, act.Recount, act.Script
                                    'a = a + 1
                                End If
                            Else
                                e1.AddMotion act.Action, act.Key, act.Data, act.Initials, act.Friction, act.Reactive, act.Recount, act.Script
                                'a = a + 1
                            End If
                            A = A + 1
                            
                        Else
                            A = A + 1
                        End If
                        
                    ElseIf ((act.Emphasis = 0) Or (act.Recount = 0)) Then 'And (Not act.Reactive = -1) Then
                        e1.DeleteMotion act.Key
                    Else
                        A = A + 1
                    End If
                    Set act = Nothing
                Loop
            End If
        End If
    End If

End Sub


Public Sub InputMove()
On Error GoTo ObjectError

    lFacesShown = 0
    lMovingObjs = 0
    lngTestCalls = 0
    If Not modParse.Player.Element Is Nothing Then
    
        If ((Perspective = Spectator) Or DebugMode) Then
        
            modParse.Player.Element.Origin.X = modParse.Player.Element.Origin.X + modParse.Player.Element.Direct.X
            modParse.Player.Element.Origin.Y = modParse.Player.Element.Origin.Y + modParse.Player.Element.Direct.Y
            modParse.Player.Element.Origin.Z = modParse.Player.Element.Origin.Z + modParse.Player.Element.Direct.Z
                    
        Else
        
            'InputMove2 modParse.Player.Element
    
        End If
        
    End If
    
    Dim cnt As Long
    cnt = 1
    Do While cnt <= Elements.Count

        If (Elements(cnt).Effect = Collides.Normal) Then
        
            InputMove2 Elements(cnt)
            
        End If
        cnt = cnt + 1
    Loop

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub InputMove2(ByRef e1 As Element)

    If ((e1.CollideIndex > -1) And e1.BoundsIndex > 0 And (e1.Effect = Collides.Normal)) Then
   ' If (e1.CollideIndex > -1) And (e1.Effect = 0) Then
    
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
    
        'commit the following object view changes based on then
        'functions called to determine the restrictive nature of
        
        If e1.AttachedTo = "" Then
                
            'elapsed2 = Timer
            SpinObject e1
            'elapsed2 = (Timer - elapsed2)
            'If elapsed2 > 0 Then Debug.Print "SpinObject: " & elapsed2
                    
                    
            'elapsed2 = Timer
            BlowObject e1
            'elapsed2 = (Timer - elapsed2)
            'If elapsed2 > 0 Then Debug.Print "BlowObject: " & elapsed2
                                    
            'elapsed2 = Timer
            MoveObject e1
            'elapsed2 = (Timer - elapsed2)
            'If elapsed2 > 0 Then Debug.Print "MoveObject: " & elapsed2
            
            lFacesShown = lFacesShown + e1.CulledFaces 'statistics
            lMovingObjs = lMovingObjs + 1
                       
        End If
        
        'elapsed2 = Timer
        If Not e1.Attachments Is Nothing Then
            
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
            Dim cnt2 As Long
            Dim e2 As Element
            cnt2 = 1
            Do While cnt2 <= e1.Attachments.Count
                Set e2 = e1.Attachments(cnt2)
                
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
                
                Set e2 = Nothing
                cnt2 = cnt2 + 1
            Loop
                

        End If
    
        Set oScaled = Nothing
        Set oRotate = Nothing
        Set oOrigin = Nothing


        'elapsed2 = (Timer - elapsed2)
        'If elapsed2 > 0 Then Debug.Print "Attachments: " & elapsed2
        
    End If

    If (Not ((e1.CollideIndex > -1) And e1.BoundsIndex > 0 And (e1.Effect = Collides.Normal))) Then
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

    
'    If (e1.Origin.Y > SpaceBoundary) Or (e1.Origin.Y < -SpaceBoundary) Then e1.Origin.Y = -e1.Origin.Y
'    If (e1.Origin.X > SpaceBoundary) Or (e1.Origin.X < -SpaceBoundary) Then e1.Origin.X = -e1.Origin.X
'    If (e1.Origin.Z > SpaceBoundary) Or (e1.Origin.Z < -SpaceBoundary) Then e1.Origin.Z = -e1.Origin.Z
End Sub

Public Function CoupleMove(ByRef obj As Element, ByVal objCollision As Long) As Boolean
'###################################################################################
'########## couple the activities of objects in collision with others ##############
'###################################################################################

    Dim A As Long
    Dim cnt As Long
    Dim act As Motion
    If (objCollision > -1) Then

        Dim e1 As Element
        cnt = 1
        Do While cnt <= Elements.Count
            Set e1 = Elements(cnt)

            If (e1.Effect = Collides.Normal) And (obj.CollideIndex > -1) Then
                If (Not e1.CollideObject = obj.CollideObject) Then
                    If (e1.CollideObject = objCollision) Then
                    'if found to be with the colliding object
                    
                        'add all motions from one to another
                        If Not obj.Motions Is Nothing Then
                            For A = 1 To obj.Motions.Count
                                Set act = obj.Motions(A)
                                If Not e1.MotionExists(act.Key) Then
                                    If act.Action = Directing Then
                                        e1.AddMotion act.Action, act.Key, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.Script
                                    End If
                                End If
                            Next
                        End If
                        
                        e1.Direct = obj.Direct 'setting this seems to "magnetically" couple the directive
                        'actions, if not, a sort of pre push and strole uneven happen like real innertia

                        CoupleMove = True
                        Exit Function
                    End If
                End If
            End If
            Set e1 = Nothing
            cnt = cnt + 1
        Loop

    End If
End Function

Public Function CoupleSpin(ByRef obj As Element, ByVal objCollision As Long) As Boolean
'###################################################################################
'########## couple the activities of objects in collision with others ##############
'###################################################################################

    Dim A As Long
    Dim cnt As Long
    Dim act As Motion
    If (objCollision > -1) Then
        If (Elements.Count > 0) Then
            Dim e1 As Element
            cnt = 1
            Do While cnt <= Elements.Count

                Set e1 = Elements(cnt)

                If (e1.Effect = Collides.Normal) And (obj.CollideIndex > -1) Then
                    If (Not e1.CollideObject = obj.CollideObject) Then
                    
                        If (e1.CollideObject = objCollision) Then
                        'if found to be with the colliding object
                        
                            'add all motions from one to another
                            If Not obj.Motions Is Nothing Then
                                For A = 1 To obj.Motions.Count
                                    Set act = obj.Motions(A)
                                    If Not e1.MotionExists(act.Key) Then
                                        If act.Action = Rotating Then
                                            e1.AddMotion act.Action, act.Key, VectorNegative(act.Data), act.Emphasis, act.Friction, act.Reactive, act.Recount, act.Script
                                        End If
                                    End If
                                Next
                            End If

                            e1.Twists = VectorMultiplyBy(AngleAxisInvert(VectorMultiplyBy(obj.Twists, RADIAN)), DEGREE)
                            
                            CoupleSpin = True
                            Exit Function
                        End If
                    End If
                End If
                Set e1 = Nothing
                cnt = cnt + 1
            Loop
        End If
    End If
End Function

Private Sub MoveObject(ByRef obj As Element)

    If obj.Direct.Equals(NoPoint) Then Exit Sub
   
    Dim inAmt As Integer
    Dim S As Space
'    For Each S In Spaces
'        If S.Boundary > 0 And S.Boundary > S.Range Then
'            If S.InSpace(Obj.Origin) Then
'                If Not S.InSpace(VectorAddition(Obj.Direct, Obj.Origin)) Then
'                    inAmt = inAmt + 1
'                    If inAmt > 1 Then Exit For
'                End If
'            End If
'        End If
'    Next
'    If inAmt > 0 Then
        For Each S In Spaces
            If S.Boundary > 0 And S.Boundary > S.Range Then
                If S.InSpace(obj.Origin) Then
                    If Not S.InSpace(VectorAddition(obj.Direct, obj.Origin)) Then

                        Set obj.Direct = NoPoint
                        

                        Exit Sub
                    
                    End If
                End If
            End If
        Next
'    End If
    
    
On Error GoTo ObjectError

    Dim objCollision As Long
    objCollision = -1

    Dim stepUpStairHeight As Single
    Dim testNudgeAdjust As Single
    
    stepUpStairHeight = 0.2
    testNudgeAdjust = 0.019 '?(Abs(GlobalGravityDirect.Y)/10)
    
    
    Dim visType As Long
    Dim bitType As Long
    bitType = 1
    visType = 2

    Dim pull As Boolean
    Dim push As Boolean

    Dim backup As D3DVECTOR
    Dim newset As D3DVECTOR

    Dim cnt As Long
    Dim cnt2 As Long
    Dim act As Motion
            
'#####################################################################################
'############# preliminary sort the type of space collision checks ###################
'#####################################################################################


    'this next part is a work around to make everything smooth as possible
    'when in collision, it is basically just constantly rotating the bound
    'object in either direction so that it does not snag on laddres, slopes
    'and when running diagnal against a wall, yet sometimes it still snags
    'because the bounds objects are cube like nature, this makes a cylinder

    Dim swapY As Single

    Static Rotator As Single
    Rotator = Rotator + IIf(modParse.Player.Camera.Angle > 0, testNudgeAdjust, -testNudgeAdjust)
    Rotator = AngleRestrict(Rotator * RADIAN) * DEGREE

    swapY = obj.Rotate.Y
    obj.Rotate.Y = Rotator
    Rotator = swapY
    
    'on with probably the longest function I ever made...
    
    
    'reset all of the vis flags to zero
    'set to zero, culling ignores them
    obj.IsMoving = Moving.None
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


        If obj.OnLadder Then 'if we are already on ladder coming in
            obj.OnLadder = TestCollision(obj, Actions.NotDefined, bitType)  'straight to test
        Else
            obj.OnLadder = TestCollision(obj, Actions.NotDefined, bitType) 'test as well but..
            If obj.OnLadder Then 'if this is the first time we are
                Do 'on a ladder coming in, clear the objects motions
                Loop Until Not obj.DeleteMotion(JumpGUID)
                For cnt = 1 To Portals.Count
                    If Not Portals(cnt).Motions Is Nothing Then
                        For cnt2 = 1 To Portals(cnt).Motions.Count
                            Set act = Portals(cnt).Motions(cnt2)
                            obj.DeleteMotion act.Key
                            Set act = Nothing
                        Next
                    End If
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

        If obj.InLiquid Then 'the same as ladder, if already liquid;
            obj.InLiquid = TestCollision(obj, Actions.NotDefined, bitType) 'straight to test
        Else
            obj.InLiquid = TestCollision(obj, Actions.NotDefined, bitType) 'test as well but..
            If obj.InLiquid Then 'first time in liquid then
                Do 'delete motions and apeal motions reference
                Loop Until Not obj.DeleteMotion(JumpGUID)
                For cnt = 1 To Portals.Count
                    If Not Portals(cnt).Motions Is Nothing Then
                        For cnt2 = 1 To Portals(cnt).Motions.Count
                            Set act = Portals(cnt).Motions(cnt2)
                            obj.DeleteMotion act.Key
                            Set act = Nothing
                        Next
                    End If
                Next
            End If
        End If

    End If


'#####################################################################################
'############# initial faces data for returning TestCollision info ###################
'#####################################################################################

    'setup the cngCamera() for Culling to work idealy, this could use improvement
    'it is used to help eliminate a majority of triangles from ever being checked
    'by addressing the collision check triangles are within this camera viewport.
    'it's an eye, up and direction vector whose information is relevant to culling


    sngCamera(0, 0) = obj.Origin.X
    sngCamera(0, 1) = obj.Origin.Y + 1
    sngCamera(0, 2) = obj.Origin.Z

    sngCamera(1, 0) = 1
    sngCamera(1, 1) = -1
    sngCamera(1, 2) = -1

    sngCamera(2, 0) = -1
    sngCamera(2, 1) = 1
    sngCamera(2, 2) = -1

'    sngCamera(0, 0) = Obj.Origin.X
'    sngCamera(0, 1) = Obj.Origin.Y + 2
'    sngCamera(0, 2) = Obj.Origin.Z
'
'    sngCamera(1, 0) = 0
'    sngCamera(1, 1) = -1
'    sngCamera(1, 2) = 0
'
'    sngCamera(2, 0) = 0
'    sngCamera(2, 1) = 0
'    sngCamera(2, 2) = -1
    
    If lngFaceCount > 0 Then 'apply the camera perspective to wean out triangles (will be reflected in the flag value)
        obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
        lCullCalls = lCullCalls + 1
    End If

    If (Elements.Count > 0) Then 'after Culling is called the flags are set for the visTYpe or not
        For Each e1 In Elements 'but since we already did our ladder and liquid, re-exclude them.
            If (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'reset to make culling ignore these in collision triangles
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
                    sngFaceVis(3, cnt2) = 0 'reset to make culling ignore these in collision triangles
                Next
            End If
        Next
    End If
    

'#####################################################################################
'############# predict the Y movements of objects in motion ##########################
'#####################################################################################

    'where the directed info data is
    'only going to be about the Y axis
    backup = ToVector(obj.Direct)
    obj.Direct.Y = backup.Y
    obj.Direct.X = 0
    obj.Direct.Z = 0

    'all the collision tests use motion data to modify values of a subset of object change
    'that object change is not applied, and any change that will normally, is ahed of time
    'in a way these are predictions of change, tested for collision 1st before binds them
    If (obj.Direct.Y <> 0) Then
        'preform check since any Y change exists at all
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then
            obj.Origin.Y = obj.Origin.Y + obj.Direct.Y  'no collision then adjust the Y to reflect the change is available
            If obj.Direct.Y > 0 Then 'and then midify the IsMoving state property of the object
                If Not ((obj.IsMoving And Moving.Flying) = Moving.Flying) Then obj.IsMoving = obj.IsMoving + Moving.Flying
            ElseIf obj.Direct.Y < 0 Then
                If Not ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving + Moving.Falling
            End If
            newset.Y = obj.Direct.Y 'record the difference change to Origin.Y
            objCollision = -1
        ElseIf (obj.Direct.Y < 0) Then 'the y movement is going down
            Do '(x,z may have or not have changed here too cause Y change)
                obj.Direct.Y = obj.Direct.Y + testNudgeAdjust 'so, we loop until we find out
                If (obj.Direct.Y >= 0) Then Exit Do 'of the collision where stands
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.Y < 0) Then
                obj.Origin.Y = obj.Origin.Y + obj.Direct.Y 'change the Y to new data, and adjust the IsMoving state for falling
                If Not ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving + Moving.Falling
                newset.Y = obj.Direct.Y 'record the difference change to Origin.Y
            End If
        ElseIf (obj.Direct.Y > 0) Then 'the y movement is going up
            Do '(x,z may have or not have changed here too cause Y change)
                obj.Direct.Y = obj.Direct.Y - testNudgeAdjust 'so, we loop until we find out
                If (obj.Direct.Y <= 0) Then Exit Do 'of the collision where stands
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.Y > 0) Then
                obj.Origin.Y = obj.Origin.Y + obj.Direct.Y 'change the Y to new data, and adjust the IsMoving state for falling
                If Not ((obj.IsMoving And Moving.Flying) = Moving.Flying) Then obj.IsMoving = obj.IsMoving + Moving.Flying
                newset.Y = obj.Direct.Y 'record the difference change to Origin.Y
            End If
        End If
    End If
    
'#####################################################################################
'############# adjust face data based on the TestCollision resulted ##################
'#####################################################################################


    If (Elements.Count > 0) Then
        For Each e1 In Elements 'reset the types of Collision effects to be only object to object collision
            If (e1.CollideObject = obj.CollideObject) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = visType 'non zero here ensures Culling to consider it left in
                Next
            ElseIf (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'still no ladder checking, we got it complete first thing
                Next
            ElseIf (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
                For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0 'still no liquid checking, we got it complete first thing
                Next
            End If
        Next
    End If
    
'#####################################################################################
'############# last call to MoveObejct collisions couple Motion here ###############
'#####################################################################################

    CoupleMove obj, objCollision 'above was information testing positive for collision
    'when newest.y <> 0, and objCollision is > -1 which is the first check in CoupleMove()
    'and before any of the following check will be done, the above Y axis will be known.
    'therefore making a call to coupleMove, possibly stacks motions on touching objects,
    'and every other call to CoupleMove below is doing the same thing if moves are bound.
    
'#####################################################################################
'############# predict the X movements of objects in motion ##########################
'#####################################################################################

    'If Obj.OnLadder Then testNudgeAdjust = testNudgeAdjust / 0.1
    
    
    obj.Direct.Y = 0
    obj.Direct.X = backup.X

    'very similar recent code above on the Y axis, we will be doing it
    If (obj.Direct.X <> 0) Then 'on the X (here) and later on the Z axis
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then 'make the change
            obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust the flags
            If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
            If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
            End If
            newset.X = obj.Direct.X
            objCollision = -1
        ElseIf (obj.Direct.X < 0) Then
            Do
                obj.Direct.X = obj.Direct.X + testNudgeAdjust
                If (obj.Direct.X >= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.X < 0) Then 'make the change
                obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust the flags
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                    If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
                End If
                newset.X = obj.Direct.X
            End If
        ElseIf (obj.Direct.X > 0) Then
            Do
                obj.Direct.X = obj.Direct.X - testNudgeAdjust
                If (obj.Direct.X <= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.X > 0) Then 'make the change
                obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust the flags
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                    If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
                End If
                newset.X = obj.Direct.X
            End If
        End If
    End If

'#####################################################################################
'############# predict the Z movements of objects in motion ##########################
'#####################################################################################
    
    obj.Direct.X = 0
    obj.Direct.Z = backup.Z

    'very similar recent code above on the X and Y axis, we will
    If (obj.Direct.Z <> 0) Then 'be doing it here on the Z axis
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then 'make the change
            obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'add the movement, and adjust the flags
            If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level 'adjust
            If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
            End If
            newset.Z = obj.Direct.Z
            objCollision = -1
        ElseIf (obj.Direct.Z < 0) Then
            Do
                obj.Direct.Z = obj.Direct.Z + testNudgeAdjust
                If (obj.Direct.Z >= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.Z < 0) Then 'make the change
                obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'add the movement, and adjust the flags
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                    If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
                End If
                newset.Z = obj.Direct.Z
            End If
        ElseIf (obj.Direct.Z > 0) Then
            Do
                obj.Direct.Z = obj.Direct.Z - testNudgeAdjust
                If (obj.Direct.Z <= 0) Then Exit Do
            'until we find back to no movement, or something closer inbetween is colliding
            Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
            If (obj.Direct.Z > 0) Then 'make the change
                obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'add the movement, and adjust the flags
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                If (backup.X <> newset.X) And (backup.Z <> newset.Z) And (Not (backup.Y = newset.Y)) And (Not backup.Y = 0) Then
                    If ((obj.IsMoving And Moving.Falling) = Moving.Falling) Then obj.IsMoving = obj.IsMoving - Moving.Falling
                End If
                newset.Z = obj.Direct.Z
            End If
        End If
    End If

    'If Obj.OnLadder Then testNudgeAdjust = testNudgeAdjust / 10
    
    Set obj.Direct = ToPoint(newset)
        
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
    
    CoupleMove obj, objCollision 'periodic TestCollisions may have occured a collision.

'#####################################################################################
'############# push/pull of moving objects in Y slope and small step ups #############
'#####################################################################################

    'pull is when the object is on a slope 45 degrees or more, it begins to slide down
    'from gravity.  Push is when an object collides with another non ground element, it
    'chain links to "pushing" the first object furthest from force.  small step up are
    'a vertical wall height the object can automatically drive over, i.e. it's stairs.
    
    If (Not obj.IsMoving = Moving.None) And _
        (backup.X <> newset.X Or backup.Z <> newset.Z) And _
        (Not ((obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((obj.IsMoving And Moving.Falling) = Moving.Falling)) Then
        'falling and flying flags are too also check befire here.
        
        obj.Origin.Y = obj.Origin.Y + stepUpStairHeight 'pretend it can step out of it by step up

        obj.Direct.Y = 0
        obj.Direct.X = backup.X
        obj.Direct.Z = backup.Z

        'the following two flags are the difference
        'between just setting "newst" like above.
        push = True 'one none effect object pushing another
        pull = False 'an object falling diagnal on a slope

        If (obj.Direct.X <> 0) Or (obj.Direct.Z <> 0) Then
            'first check for collision and if non exists
            'add them to the actual information data
            If (TestCollision(obj, Directing, visType, objCollision) = False) Then
                'we need a change of X or Z to consider it a pull, already
                'graivty will take effect to any free falling down objects.
                If obj.Direct.X <> 0 Then
                    obj.Origin.X = obj.Origin.X + obj.Direct.X
                    newset.X = obj.Direct.X
                    pull = True
                End If
                If obj.Direct.Z <> 0 Then
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z
                    newset.Z = obj.Direct.Z
                    pull = True
                End If
                objCollision = -1
            ElseIf (obj.Direct.X < 0) And (obj.Direct.Z < 0) Then 'here we do two axis checks at once
                Do
                    obj.Direct.X = obj.Direct.X + testNudgeAdjust
                    obj.Direct.Z = obj.Direct.Z + testNudgeAdjust
                    If ((obj.Direct.X >= 0) Or (obj.Direct.Z >= 0)) Then Exit Do
                'slow down the change prediction and check until no collision is found
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X < 0) And (obj.Direct.Z < 0) Then
                    'adjust change and flags to reflect happened
                    obj.Origin.X = obj.Origin.X + obj.Direct.X
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    newset.Z = obj.Direct.Z
                    pull = True
                End If

            ElseIf (obj.Direct.X > 0) And (obj.Direct.Z > 0) Then 'here we do two axis checks at once
                Do
                    obj.Direct.X = obj.Direct.X - testNudgeAdjust
                    obj.Direct.Z = obj.Direct.Z - testNudgeAdjust
                    If ((obj.Direct.X <= 0) Or (obj.Direct.Z <= 0)) Then Exit Do
                'slow down the change prediction and check until no collision is found
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X > 0) And (obj.Direct.Z > 0) Then
                    'adjust change and flags to reflect happened
                    obj.Origin.X = obj.Origin.X + obj.Direct.X
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    newset.Z = obj.Direct.Z
                    pull = True
                End If

            ElseIf (obj.Direct.X < 0) And (obj.Direct.Z > 0) Then 'here we do two axis checks at once
                Do
                    obj.Direct.X = obj.Direct.X + testNudgeAdjust
                    obj.Direct.Z = obj.Direct.Z - testNudgeAdjust
                    If ((obj.Direct.X >= 0) Or (obj.Direct.Z <= 0)) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X < 0) And (obj.Direct.Z > 0) Then
                    'adjust change and flags to reflect happened
                    obj.Origin.X = obj.Origin.X + obj.Direct.X
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    newset.Z = obj.Direct.Z
                    pull = True
                End If
            ElseIf (obj.Direct.X > 0) And (obj.Direct.Z < 0) Then 'here we do two axis checks at once
                Do
                    obj.Direct.X = obj.Direct.X - testNudgeAdjust
                    obj.Direct.Z = obj.Direct.Z + testNudgeAdjust
                    If ((obj.Direct.X <= 0) Or (obj.Direct.Z >= 0)) Then Exit Do
                    'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X > 0) And (obj.Direct.Z < 0) Then
                    obj.Origin.X = obj.Origin.X + obj.Direct.X
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    newset.Z = obj.Direct.Z
                    pull = True
                End If
            End If
        End If
        
        obj.Origin.Y = obj.Origin.Y - stepUpStairHeight 'no longer pretending it can step up

        If pull Then push = False

    End If

    Set obj.Direct = ToPoint(backup)

'#####################################################################################
'############# those passing with out pressure couple activities first ###############
'#####################################################################################

    
    CoupleMove obj, objCollision 'periodic TestCollisions may have occured a collision.


'#####################################################################################
'############# as an object first in motions continues it's push in moved Y ##########
'#####################################################################################


    If push And (Not obj.IsMoving = Moving.None) And _
        (backup.X <> newset.X Or backup.Z <> newset.Z) And _
        (Not ((obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((obj.IsMoving And Moving.Falling) = Moving.Falling)) Then

        'where a change existe already, during checks on
        'each axis then occurs the need to change again.
        'so besides the gate IF above this is to do it
        'simgularly on X and Z, which was done above so
        '
        obj.Origin.Y = obj.Origin.Y + stepUpStairHeight 'pretend it can step out of it by step up

        obj.Direct.Y = 0
        obj.Direct.X = backup.X
        obj.Direct.Z = backup.Z

        push = False

        If (obj.Direct.X <> 0) Then 'first comes the X axis
            If (TestCollision(obj, Directing, visType, objCollision) = False) Then
                obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust change and flags to reflect happened
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                newset.X = obj.Direct.X
                push = True
                objCollision = -1
            ElseIf (obj.Direct.X < 0) Then
                Do
                    obj.Direct.X = obj.Direct.X + testNudgeAdjust
                    If (obj.Direct.X >= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X < 0) Then
                    obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust change and flags to reflect happened
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    push = True
                End If

            ElseIf (obj.Direct.X > 0) Then
                Do
                    obj.Direct.X = obj.Direct.X - testNudgeAdjust
                    If (obj.Direct.X <= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.X > 0) Then
                    obj.Origin.X = obj.Origin.X + obj.Direct.X 'adjust change and flags to reflect happened
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.X = obj.Direct.X
                    push = True
                End If
            End If
        End If
        
        If (obj.Direct.Z <> 0) Then 'first comes the Z axis
            If (TestCollision(obj, Directing, visType, objCollision) = False) Then
                obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'adjust change and flags to reflect happened
                If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                newset.Z = obj.Direct.Z
                push = True
                objCollision = -1
            ElseIf (obj.Direct.Z < 0) Then
                Do
                    obj.Direct.Z = obj.Direct.Z + testNudgeAdjust
                    If (obj.Direct.Z >= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.Z < 0) Then
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'adjust change and flags to reflect happened
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.Z = obj.Direct.Z
                    push = True
                End If

            ElseIf (obj.Direct.Z > 0) Then
                Do
                    obj.Direct.Z = obj.Direct.Z - testNudgeAdjust
                    If (obj.Direct.Z <= 0) Then Exit Do
                'slow down the change prediction and check until
                Loop Until (TestCollision(obj, Directing, visType, objCollision) = False)
                If (obj.Direct.Z > 0) Then
                    obj.Origin.Z = obj.Origin.Z + obj.Direct.Z 'adjust change and flags to reflect happened
                    If Not ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving + Moving.Level
                    newset.Z = obj.Direct.Z
                    push = True
                End If

            End If
        End If
        
        obj.Origin.Y = obj.Origin.Y - stepUpStairHeight 'no longer pretending it can step up

    End If


'#####################################################################################
'############# coupled in if pushing or pulling, adjust the X/Z gliding ##############
'#####################################################################################

    'next are some final adjustments to requested "Direct" to reflect what is
    'found possible verses what we recieved in attempted moves for an object.
    'due to zero'ing out directive motions, that may re-adjust our push or pull.
    'they are only needed now in testing skipping this block, when not skipped
    'they may become adjusted to skippin the last block of commented apart code
    
    If (pull Xor push) And (Not ((obj.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((obj.IsMoving And Moving.Falling) = Moving.Falling)) And _
        ((obj.IsMoving And Moving.Level) = Moving.Level) Then
        
        obj.Direct.Y = 0
        obj.Direct.X = 0
        obj.Direct.Z = 0

        'slow down the change prediction and check until
        Do While (TestCollision(obj, Directing, visType, objCollision) = True)
            obj.Direct.Y = obj.Direct.Y + testNudgeAdjust
        Loop

        If ((obj.Direct.Y >= 0) And (obj.Direct.Y < 0.3)) Or ((obj.Direct.Y >= 0) And (obj.Direct.Y <= 0.2)) Then
            obj.Origin.Y = obj.Origin.Y + obj.Direct.Y 'adjust change and flags to reflect happened
            If Not ((obj.IsMoving And Moving.Stepping) = Moving.Stepping) Then obj.IsMoving = obj.IsMoving + Moving.Stepping
            If ((obj.IsMoving And Moving.Level) = Moving.Level) Then obj.IsMoving = obj.IsMoving - Moving.Level
            newset.Y = obj.Direct.Y
        End If

    ElseIf ((obj.IsMoving = Moving.None) And ((backup.X = 0 And backup.Z = 0) And (newset.X = 0 And newset.Z = 0))) Then

        push = False
        pull = False
        
        obj.Direct.Y = -testNudgeAdjust
        If Not push Then obj.Direct.X = testNudgeAdjust
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then
            pull = True
            objCollision = -1
        Else
            pull = False
            obj.Direct.Y = 0
            obj.Direct.X = 0
        End If

        If Not pull Then obj.Direct.Y = -testNudgeAdjust
        obj.Direct.Z = testNudgeAdjust
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then
            push = True
            objCollision = -1
        Else
            push = False
            obj.Direct.Y = 0
            obj.Direct.Z = 0
        End If

        If Not pull And Not push Then obj.Direct.Y = -testNudgeAdjust
        obj.Direct.X = -testNudgeAdjust
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then
            pull = (push And Not pull) Or (Not push And Not pull)
            objCollision = -1
        Else
            obj.Direct.Y = 0
            obj.Direct.X = 0
        End If

        If Not push And Not pull Then obj.Direct.Y = -testNudgeAdjust
        obj.Direct.Z = -testNudgeAdjust
        If (TestCollision(obj, Directing, visType, objCollision) = False) Then
            push = (pull And Not push) Or (Not push And Not pull)
            objCollision = -1
        Else
            obj.Direct.Y = 0
            obj.Direct.Z = 0
        End If


'#####################################################################################
'############# final asjustments made in impressions on self when alone ##############
'#####################################################################################

        'the last check which is to infer movements
        'by the rate of adjust, for steps and steeps
        If (push Xor pull) Or (push And pull) Then

            obj.Direct.Y = 0

            Do
                obj.Origin.Y = obj.Origin.Y - testNudgeAdjust
                If pull Then
                    obj.Origin.X = obj.Origin.X + testNudgeAdjust
                    If (TestCollision(obj, Directing, visType, objCollision) = True) Then
                        obj.Origin.X = obj.Origin.X - (testNudgeAdjust * 2)
                        If (TestCollision(obj, Directing, visType, objCollision) = True) Then
                            obj.Origin.Y = obj.Origin.Y + (testNudgeAdjust / 3)
                        Else
                            objCollision = -1
                            Do
                                If obj.Origin.X + (testNudgeAdjust / 3) <> testNudgeAdjust Then Exit Do
                                obj.Origin.X = obj.Origin.X + (testNudgeAdjust / 3)
                            Loop Until (TestCollision(obj, Directing, visType, objCollision) = True)
                            obj.Origin.X = obj.Origin.X - (testNudgeAdjust / 3)
                        End If
                    Else
                        objCollision = -1
                        Do
                            If obj.Origin.X - (testNudgeAdjust / 3) <> testNudgeAdjust Then Exit Do
                            obj.Origin.X = obj.Origin.X - (testNudgeAdjust / 3)
                        Loop Until (TestCollision(obj, Directing, visType, objCollision) = True)
                        obj.Origin.X = obj.Origin.X + (testNudgeAdjust / 3)
                    End If
                ElseIf push Then

                    obj.Origin.Z = obj.Origin.Z + testNudgeAdjust
                    If (TestCollision(obj, Directing, visType, objCollision) = True) Then
                        obj.Origin.Z = obj.Origin.Z - (testNudgeAdjust * 2)
                        If (TestCollision(obj, Directing, visType, objCollision) = True) Then
                            obj.Origin.Y = obj.Origin.Y + (testNudgeAdjust / 3)
                        Else
                            objCollision = -1
                            Do
                                If obj.Origin.Z + (testNudgeAdjust / 3) <> testNudgeAdjust Then Exit Do
                                obj.Origin.Z = obj.Origin.Z + (testNudgeAdjust / 3)
                            Loop Until (TestCollision(obj, Directing, visType, objCollision) = True)
                            obj.Origin.Z = obj.Origin.Z - (testNudgeAdjust / 3)
                        End If
                    Else
                        objCollision = -1
                        Do
                            If obj.Origin.Z - (testNudgeAdjust / 3) <> testNudgeAdjust Then Exit Do
                            obj.Origin.Z = obj.Origin.Z - (testNudgeAdjust / 3)
                        Loop Until (TestCollision(obj, Directing, visType, objCollision) = True)
                        obj.Origin.Z = obj.Origin.Z + (testNudgeAdjust / 3)
                    End If
                End If

            Loop While (TestCollision(obj, Directing, visType, objCollision) = True)

        End If
        
    End If

    swapY = obj.Rotate.Y
    obj.Rotate.Y = Rotator
    Rotator = swapY

    Exit Sub
ObjectError:

    swapY = obj.Rotate.Y
    obj.Rotate.Y = Rotator
    Rotator = swapY
    
'#####################################################################################
'############# direct activities are primed for Next call to MoveObject  #############
'#####################################################################################



    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
   ' Resume
End Sub

Private Sub SpinObject(ByRef obj As Element)

On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################

    If Not obj Is Nothing Then

        If Not TestCollision(obj, Rotating, 2) Then

            obj.Rotate.X = obj.Rotate.X + obj.Twists.X
            obj.Rotate.Y = obj.Rotate.Y + obj.Twists.Y
            obj.Rotate.Z = obj.Rotate.Z + obj.Twists.Z

        End If

        obj.Twists = NoAngle

    End If

Exit Sub

'    Dim e1 As Element
'    Dim cnt2 As Long
'    Dim visType As Long
'
'    visType = 2
'
'    Dim objCollision As Long
'
'    If Not obj Is Nothing Then
'
'
'        If Not obj.Twists.Equals(NoAngle) Then
'
'
'            'this is if at all we have a force in a rotation we need to check/clear
'            Dim backup As New Point
'            backup = obj.Rotate
'
'
'            obj.Rotate.X = obj.Rotate.X + obj.Twists.X
'            obj.Rotate.Y = obj.Rotate.Y + obj.Twists.Y
'            obj.Rotate.Z = obj.Rotate.Z + obj.Twists.Z
'
'
'            If obj.AttachedTo = "" Then
'
'                If (Elements.Count > 0) Then
'                    For Each e1 In Elements 'reset the types of Collision effects to be only object to object collision
'                        If (e1.CollideObject = obj.CollideObject) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
'                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
'                                sngFaceVis(3, cnt2) = visType 'non zero here ensures Culling to consider it left in
'                            Next
'                        ElseIf (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
'                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
'                                sngFaceVis(3, cnt2) = 0 'still no ladder checking, we got it complete first thing
'                            Next
'                        ElseIf (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
'                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
'                                sngFaceVis(3, cnt2) = 0 'still no liquid checking, we got it complete first thing
'                            Next
'                        End If
'                    Next
'                End If
'
'
'                If Not TestCollision(obj, NotDefined, 2, objCollision) Then
'                    'We are able to rotate with the prospective rotation so, commit it.
'
'                    obj.Rotate = VectorMultiplyBy(AngleAxisRestrict(VectorMultiplyBy(obj.Rotate, RADIAN)), DEGREE)
'
'                    obj.Twists = NoAngle
'
'                Else
'
'                    'we've collided with something, couple the spin which will
'                    'turn a collided object in opposing rotation to this one
'                    CoupleSpin obj, objCollision
'
'                   ' Obj.Rotate = backup
'                    obj.Twists = NoAngle
'
'                End If
'            Else
'                obj.Twists = NoAngle
'            End If
'
'        Else
'
'            If (((obj.Direct.Y = 0) And (obj.Direct.X = 0) And (obj.Direct.Z = 0)) Or _
'                ((obj.Direct.Y < 0) And (obj.Direct.X = 0) And (obj.Direct.Z = 0))) And obj.Gravitational Then
'                'only if no other force is applied or only down force
'
'                'Dim backupOrigin As New Point
'                'backupOrigin = Obj.Origin
'
'                Dim newset As New Point
'                Dim backup2 As New Point
'                'Dim newestOrigin As New Point
'                Dim testNudgeAdjust As Single
'
'                backup = obj.Twists
'                backup2 = obj.Rotate
'
'                testNudgeAdjust = (Abs(GlobalGravityRotate.Data.Y) * DEGREE)
'                'test a nudge on each the N,W,E,W, NE, SE, SW and NW poles,
'                'if all pass for can move (free falling) then no don't rotate
'
'                'if only a certian few or more pole pass but not all, then rotate
'                '(i.e. on the edge of an overhanig, or another obj is pushing it)
'                If obj.Key = "pawn3" Then
'                 '   Stop
'                End If
'
'
'
'
''                If (Elements.Count > 0) Then
''                    For Each e1 In Elements 'reset the types of Collision effects to be only object to object collision
''                        If (e1.CollideObject = Obj.CollideObject) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
''                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
''                                sngFaceVis(3, cnt2) = visType 'non zero here ensures Culling to consider it left in
''                            Next
''                        ElseIf (e1.Effect = Collides.Ladder) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
''                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
''                                sngFaceVis(3, cnt2) = 0 'still no ladder checking, we got it complete first thing
''                            Next
''                        ElseIf (e1.Effect = Collides.Liquid) And (e1.CollideIndex > -1) And e1.BoundsIndex > 0 Then
''                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
''                                sngFaceVis(3, cnt2) = 0 'still no liquid checking, we got it complete first thing
''                            Next
''                        End If
''                    Next
''                End If
''
''
''                'Debug.Print "NoData"
''
''                Dim cycle As Long
''                cycle = 1
''                Do While cycle <= 8
''                    Obj.Twists = backup
''
''                    Select Case cycle
''                        Case 1
''                            Obj.Twists.X = Obj.Twists.X + testNudgeAdjust
''                        Case 2
''                            Obj.Twists.X = Obj.Twists.X + -testNudgeAdjust
''                        Case 3
''                            Obj.Twists.Z = Obj.Twists.Z + testNudgeAdjust
''                        Case 4
''                            Obj.Twists.Z = Obj.Twists.Z + -testNudgeAdjust
''                        Case 5
''                            Obj.Twists.X = Obj.Twists.X + (testNudgeAdjust / 2)
''                            Obj.Twists.Z = Obj.Twists.Z + (testNudgeAdjust / 2)
''                        Case 6
''                            Obj.Twists.X = Obj.Twists.X + -(testNudgeAdjust / 2)
''                            Obj.Twists.Z = Obj.Twists.Z + -(testNudgeAdjust / 2)
''                        Case 7
''                            Obj.Twists.X = Obj.Twists.X + -(testNudgeAdjust / 2)
''                            Obj.Twists.Z = Obj.Twists.Z + (testNudgeAdjust / 2)
''                        Case 8
''                            Obj.Twists.X = Obj.Twists.X + (testNudgeAdjust / 2)
''                            Obj.Twists.Z = Obj.Twists.Z + (testNudgeAdjust / 2)
''                    End Select
''
''                    'all the collision tests use motion data to modify values of a subset of object change
''                    'that object change is not applied, and any change that will normally, is ahed of time
''                    'in a way these are predictions of change, tested for collision 1st before binds them
''                    If (Obj.Twists.X <> 0) And (Obj.Twists.Z = 0) Then
''                        'preform check since any Y change exists at all
''                        If (TestCollision(Obj, Rotating, visType, objCollision) = False) Then
''                            Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X  'no collision then adjust the X to reflect the change is available
''                            If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                            newset.X = Obj.Twists.X 'record the difference change to Rotate.X
''                            cycle = 9
''                        ElseIf (Obj.Twists.X < 0) Then 'the y movement is going down
''                            Do '(x,z may have or not have changed here too cause X change)
''                                Obj.Twists.X = Obj.Twists.X + testNudgeAdjust 'so, we loop until we find out
''                                If (Obj.Twists.X >= 0) Then Exit Do 'of the collision where stands
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X < 0) Then
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X 'change the X to new data, and adjust the IsMoving state for falling
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X 'record the difference change to Rotate.X
''                                cycle = 9
''                            End If
''                        ElseIf (Obj.Twists.X > 0) Then 'the y movement is going up
''                            Do '(x,z may have or not have changed here too cause X change)
''                                Obj.Twists.X = Obj.Twists.X - testNudgeAdjust 'so, we loop until we find out
''                                If (Obj.Twists.X <= 0) Then Exit Do 'of the collision where stands
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X > 0) Then
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X 'change the X to new data, and adjust the IsMoving state for falling
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X 'record the difference change to Rotate.X
''                                cycle = 9
''                            End If
''                        End If
''                    ElseIf (Obj.Twists.Z <> 0) And (Obj.Twists.X = 0) Then
''                        'preform check since any Y change exists at all
''                        If (TestCollision(Obj, Rotating, visType, objCollision) = False) Then
''                            Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z  'no collision then adjust the X to reflect the change is available
''                            If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                            newset.Z = Obj.Twists.Z 'record the difference change to Rotate.Z
''                            cycle = 9
''                        ElseIf (Obj.Twists.Z < 0) Then 'the y movement is going down
''                            Do '(x,z may have or not have changed here too cause X change)
''                                Obj.Twists.Z = Obj.Twists.Z + testNudgeAdjust 'so, we loop until we find out
''                                If (Obj.Twists.Z >= 0) Then Exit Do 'of the collision where stands
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.Z < 0) Then
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z 'change the X to new data, and adjust the IsMoving state for falling
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.Z = Obj.Twists.Z 'record the difference change to Rotate.Z
''                                cycle = 9
''                            End If
''                        ElseIf (Obj.Twists.Z > 0) Then 'the y movement is going up
''                            Do '(x,z may have or not have changed here too cause X change)
''                                Obj.Twists.Z = Obj.Twists.Z - testNudgeAdjust  'so, we loop until we find out
''                                If (Obj.Twists.Z <= 0) Then Exit Do 'of the collision where stands
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.Z > 0) Then
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z 'change the X to new data, and adjust the IsMoving state for falling
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.Z = Obj.Twists.Z 'record the difference change to Rotate.Z
''                                cycle = 9
''                            End If
''                        End If
''                    ElseIf (Obj.Twists.X <> 0) Or (Obj.Twists.Z <> 0) Then
''                        'first check for collision and if non exists
''                        'add them to the actual information data
''                        If (TestCollision(Obj, Rotating, visType, objCollision) = False) Then
''                            'we need a change of X or Z to consider it a pull, already
''                            'graivty will take effect to any free falling down objects.
''                            If Obj.Twists.X <> 0 Then
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
''                                newset.X = Obj.Twists.X
''                                cycle = 9
''                            End If
''                            If Obj.Twists.Z <> 0 Then
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
''                                newset.Z = Obj.Twists.Z
''                                cycle = 9
''                            End If
''                        ElseIf (Obj.Twists.X < 0) And (Obj.Twists.Z < 0) Then 'here we do two axis checks at once
''                            Do
''                                Obj.Twists.X = Obj.Twists.X + (testNudgeAdjust / 2)
''                                Obj.Twists.Z = Obj.Twists.Z + (testNudgeAdjust / 2)
''                                If ((Obj.Twists.X >= 0) Or (Obj.Twists.Z >= 0)) Then Exit Do
''                            'slow down the change prediction and check until no collision is found
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X < 0) And (Obj.Twists.Z < 0) Then
''                                'adjust change and flags to reflect happened
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X
''                                newset.Z = Obj.Twists.Z
''                                cycle = 9
''                            End If
''
''                        ElseIf (Obj.Twists.X > 0) And (Obj.Twists.Z > 0) Then 'here we do two axis checks at once
''                            Do
''                                Obj.Twists.X = Obj.Twists.X - (testNudgeAdjust / 2)
''                                Obj.Twists.Z = Obj.Twists.Z - (testNudgeAdjust / 2)
''                                If ((Obj.Twists.X <= 0) Or (Obj.Twists.Z <= 0)) Then Exit Do
''                            'slow down the change prediction and check until no collision is found
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X > 0) And (Obj.Twists.Z > 0) Then
''                                'adjust change and flags to reflect happened
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X
''                                newset.Z = Obj.Twists.Z
''                                cycle = 9
''                            End If
''
''                        ElseIf (Obj.Twists.X < 0) And (Obj.Twists.Z > 0) Then 'here we do two axis checks at once
''                            Do
''                                Obj.Twists.X = Obj.Twists.X + (testNudgeAdjust / 2)
''                                Obj.Twists.Z = Obj.Twists.Z - (testNudgeAdjust / 2)
''                                If ((Obj.Twists.X >= 0) Or (Obj.Twists.Z <= 0)) Then Exit Do
''                            'slow down the change prediction and check until
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X < 0) And (Obj.Twists.Z > 0) Then
''                                'adjust change and flags to reflect happened
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X
''                                newset.Z = Obj.Twists.Z
''                                cycle = 9
''                            End If
''                        ElseIf (Obj.Twists.X > 0) And (Obj.Twists.Z < 0) Then 'here we do two axis checks at once
''                            Do
''                                Obj.Twists.X = Obj.Twists.X - (testNudgeAdjust / 2)
''                                Obj.Twists.Z = Obj.Twists.Z + (testNudgeAdjust / 2)
''                                If ((Obj.Twists.X <= 0) Or (Obj.Twists.Z >= 0)) Then Exit Do
''                                'slow down the change prediction and check until
''                            Loop Until (TestCollision(Obj, Rotating, visType, objCollision) = False)
''                            If (Obj.Twists.X > 0) And (Obj.Twists.Z < 0) Then
''                                Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
''                                Obj.Rotate.Z = Obj.Rotate.Z + Obj.Twists.Z
''                                If Not ((Obj.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.IsMoving = Obj.IsMoving + Moving.Falling
''                                newset.X = Obj.Twists.X
''                                newset.Z = Obj.Twists.Z
''                                cycle = 9
''                            End If
''                        End If
''                    End If
''
''
''                    cycle = cycle + 1
''                Loop
'
''                If cycle = 10 Then 'we had an adjustment
''                    Obj.Twists = newset
''                Else
''                    Obj.Twists = backup
''                End If
''                Obj.Rotate = backup2
'
'            Else
'                'the majority axis of direction movement should only couplemove on a plane
'                'while the others can be tested for slight turns, when SpinObject before MoveObject
'
'
'            End If
'
'        End If
'
'    End If

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub BlowObject(ByRef obj As Element)

    
    If obj.Scalar.Equals(NoPoint) Then Exit Sub
    
On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################

    If Not obj Is Nothing Then
    
        If Not TestCollision(obj, Scaling, 2) Then
        
            obj.Scaled.X = obj.Scaled.X + obj.Scalar.X
            obj.Scaled.Y = obj.Scaled.Y + obj.Scalar.Y
            obj.Scaled.Z = obj.Scaled.Z + obj.Scalar.Z
            
        End If
        
        obj.Scalar = NoPoint
    
    End If

    
    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'Public Sub SetCollisionCamera(ByVal Viewing As CameraCollision, ByRef UsePlayer)
'    Dim p As Point
'    Dim useObj As Element
'    Select Case TypeName(UsePlayer)
'        Case "Boolean"
'            If CBool(UsePlayer) Then
'                sngCamera(0, 0) = modParse.Player.Origin.X
'                sngCamera(0, 1) = modParse.Player.Origin.Y
'                sngCamera(0, 2) = modParse.Player.Origin.Z
'                Set p = VectorRotateY(VectorRotateX(MakePoint(0, 0, 1), modParse.Player.Camera.Pitch), modParse.Player.Camera.Angle)
'            Else
'                sngCamera(0, 0) = 0
'                sngCamera(0, 1) = 0
'                sngCamera(0, 2) = 0
'                Set p = MakePoint(1, 1, 1)
'            End If
'        Case "Element"
'            Set useObj = UsePlayer
'            sngCamera(0, 0) = useObj.Origin.X
'            sngCamera(0, 1) = useObj.Origin.Y
'            sngCamera(0, 2) = useObj.Origin.Z
'            Set p = MakePoint(1, 1, 1)
'        Case "Player"
'            Set useObj = UsePlayer
'            sngCamera(0, 0) = useObj.Origin.X
'            sngCamera(0, 1) = useObj.Origin.Y
'            sngCamera(0, 2) = useObj.Origin.Z
'            Set p = VectorRotateY(VectorRotateX(MakePoint(0, 0, 1), modParse.Player.Camera.Pitch), modParse.Player.Camera.Angle)
'        Case "Nothing"
'            sngCamera(0, 0) = 0
'            sngCamera(0, 1) = 0
'            sngCamera(0, 2) = 0
'            Set p = MakePoint(1, 1, 1)
'    End Select
'
'    Select Case Viewing
'        Case CameraFront
'            sngCamera(2, 0) = 1
'            sngCamera(2, 1) = 0
'            sngCamera(2, 2) = 0
'        Case CameraBack
'            Set p = VectorRotateY(p, 180 * RADIAN)
'            sngCamera(2, 0) = -1
'            sngCamera(2, 1) = 0
'            sngCamera(2, 2) = 0
'        Case CameraLeft
'            Set p = VectorRotateY(p, 270 * RADIAN)
'            sngCamera(2, 0) = 0
'            sngCamera(2, 1) = 1
'            sngCamera(2, 2) = 0
'        Case CameraRight
'            Set p = VectorRotateY(p, 90 * RADIAN)
'            sngCamera(2, 0) = 0
'            sngCamera(2, 1) = -1
'            sngCamera(2, 2) = 0
'        Case CameraBottom
'            Set p = VectorRotateX(p, 90 * RADIAN)
'            sngCamera(2, 0) = 0
'            sngCamera(2, 1) = 0
'            sngCamera(2, 2) = 1
'        Case CameraTop
'            Set p = VectorRotateX(p, 270 * RADIAN)
'            sngCamera(2, 0) = 0
'            sngCamera(2, 1) = 0
'            sngCamera(2, 2) = -1
'    End Select
'
'    sngCamera(1, 0) = p.X
'    sngCamera(1, 1) = p.Y
'    sngCamera(1, 2) = p.Z
'
'End Sub
Public Function TestCollision(ByRef obj As Element, ByRef Action As Actions, ByVal visType As Long, Optional ByRef lngCollideObj As Long = -1) As Boolean
On Error GoTo ObjectError


'#####################################################################################
'############# face data is temporary transformed and checked for collision ##########
'#####################################################################################

    If obj Is Nothing Then Exit Function

    If Action = Rotating Then

'#####################################################################################
'############# in rotation collision we re-adjsut culling view direction #############
'#####################################################################################

      '  SetCollisionCamera CameraBottom, True
        

'        sngCamera(0, 0) = Obj.Origin.X
'        sngCamera(0, 1) = Obj.Origin.Y + 2
'        sngCamera(0, 2) = Obj.Origin.Z
'
'        sngCamera(1, 0) = 0
'        sngCamera(1, 1) = -1
'        sngCamera(1, 2) = 0
'
'        Dim p As Point
'        Set p = VectorRotateY(VectorRotateX(MakePoint(0, 0, 1), modParse.Player.Camera.Pitch), modParse.Player.Camera.Angle)
'
'        sngCamera(2, 0) = Round(p.X, 6)
'        sngCamera(2, 1) = Round(p.Y, 6)
'        sngCamera(2, 2) = Round(p.Z, 6)


        sngCamera(0, 0) = obj.Origin.X
        sngCamera(0, 1) = obj.Origin.Y + 1
        sngCamera(0, 2) = obj.Origin.Z

        sngCamera(1, 0) = 1
        sngCamera(1, 1) = -1
        sngCamera(1, 2) = -1

        sngCamera(2, 0) = -1
        sngCamera(2, 1) = 1
        sngCamera(2, 2) = -1
        
        If lngFaceCount > 0 Then
            obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
            lCullCalls = lCullCalls + 1
        End If

    End If


'#####################################################################################
'############# create a transform matrix with the changes applied ####################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim V(2) As D3DVECTOR
    Dim N As D3DVECTOR

    Dim matScale As D3DMATRIX
    Dim matMesh As D3DMATRIX
    Dim matRot As D3DMATRIX
    
    D3DXMatrixIdentity matMesh
    D3DXMatrixIdentity matRot
    D3DXMatrixIdentity matScale

    
    If (Action And Scaling) = Scaling Then
        D3DXMatrixScaling matScale, (obj.Scaled.X + obj.Scalar.X), (obj.Scaled.Y + obj.Scalar.Y), (obj.Scaled.Z + obj.Scalar.Z)
    Else
        D3DXMatrixScaling matScale, obj.Scaled.X, obj.Scaled.Y, obj.Scaled.Z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale


    If (Action And Directing) = Directing Then
        D3DXMatrixTranslation matScale, (obj.Origin.X + obj.Direct.X), (obj.Origin.Y + obj.Direct.Y), (obj.Origin.Z + obj.Direct.Z)
    Else
        D3DXMatrixTranslation matScale, obj.Origin.X, obj.Origin.Y, obj.Origin.Z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale
    
    If (Action And Rotating) = Rotating Then

        D3DXMatrixRotationX matRot, ((obj.Rotate.X + obj.Twists.X) * RADIAN)
        'D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationY matRot, ((obj.Rotate.Y + obj.Twists.Y) * RADIAN)
        'D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationZ matRot, ((obj.Rotate.Z + obj.Twists.Z) * RADIAN)
        'D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
    Else

        D3DXMatrixRotationX matRot, (obj.Rotate.X * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationY matRot, (obj.Rotate.Y * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationZ matRot, (obj.Rotate.Z * RADIAN)
        D3DXMatrixMultiply matMesh, matRot, matMesh

    End If


    
            
    If lngFaceCount > 0 And obj.CollideIndex > -1 And obj.BoundsIndex > 0 Then
    

'#####################################################################################
'############# update face data with the transformation matrix #######################
'#####################################################################################


        For Face = obj.CollideIndex To (obj.CollideIndex + Meshes(obj.BoundsIndex).Mesh.GetNumFaces) - 1
    
            For cnt = 0 To 2
                
                V(cnt).X = Meshes(obj.BoundsIndex).Verticies(Index + cnt).X
                V(cnt).Y = Meshes(obj.BoundsIndex).Verticies(Index + cnt).Y
                V(cnt).Z = Meshes(obj.BoundsIndex).Verticies(Index + cnt).Z
    
                D3DXVec3TransformCoord V(cnt), V(cnt), matMesh
                
                sngVertexX(cnt, Face) = V(cnt).X
                sngVertexY(cnt, Face) = V(cnt).Y
                sngVertexZ(cnt, Face) = V(cnt).Z

            Next
            
            Index = Index + 3
        Next

'#####################################################################################
'############# per non culled face check and result collision ########################
'#####################################################################################

        Dim lngCollideIdx As Long
        lngCollideIdx = -1
        If obj.BoundsIndex > 0 Then
            For cnt = obj.CollideIndex To (obj.CollideIndex + Meshes(obj.BoundsIndex).Mesh.GetNumFaces) - 1
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
Public Function DelCollision(ByRef obj As Element)
On Error GoTo ObjectError
    Stats_Collision_Count = Stats_Collision_Count - 1
    'Debug.Print "DelCollision"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    
    If obj.BoundsIndex > 0 Then
'        If Not Meshes Is Nothing Then
'        If Obj.BoundsIndex <= UBound(Meshes()) Then
'        If Not Meshes(Obj.BoundsIndex).Mesh Is Nothing Then
            Index = Meshes(obj.BoundsIndex).Mesh.GetNumFaces
'        End If
'        End If
'        End If
        If lngFaceCount - Index > 0 And Index >= UBound(Meshes()) Then 'Obj.CollideIndex + Index < lngFaceCount Then
    
            For Face = obj.CollideIndex To lngFaceCount - Index - 1 'Obj.CollideIndex + Index - 1
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
            
            If Not Elements Is Nothing Then
                Dim e1 As Element
                For Each e1 In Elements
                'For cnt = 1 To Elements.count
                    If e1.CollideIndex > obj.CollideIndex Then
                        e1.CollideIndex = e1.CollideIndex - Index
                    End If
                Next
            End If
            
        End If
        
        obj.CollideIndex = -1
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
'    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
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


Public Function AddCollision(ByRef obj As Element, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError
    Stats_Collision_Count = Stats_Collision_Count + 1
'#####################################################################################
'############# create face data for a mesh to external compatability #################
'#####################################################################################
    'Debug.Print "AddCollision"
    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    Dim V() As D3DVECTOR

    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    Dim vn As D3DVECTOR

    ReDim V(0 To 3) As D3DVECTOR

    If obj.BoundsIndex > 0 Then
        obj.CollideIndex = lngFaceCount
        AddCollision = lngFaceCount
    
        Dim FaceCount As Long
        Dim addingFace As Boolean

        'obj.PrepairMatrix
        obj.ApplyMatrix
        'obj.SetWorldMatrix
        
        Index = 0
        For Face = 0 To Meshes(obj.BoundsIndex).Mesh.GetNumFaces - 1
    
            For cnt = 0 To 2
    
                V(cnt).X = Meshes(obj.BoundsIndex).Verticies(Meshes(obj.BoundsIndex).Indicies(Index + cnt)).X
                V(cnt).Y = Meshes(obj.BoundsIndex).Verticies(Meshes(obj.BoundsIndex).Indicies(Index + cnt)).Y
                V(cnt).Z = Meshes(obj.BoundsIndex).Verticies(Meshes(obj.BoundsIndex).Indicies(Index + cnt)).Z
    
                'D3DXVec3TransformCoord vn, v(cnt), matObject
                vn = ToVector(obj.PointMatrix(ToPoint(V(cnt))))
                
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
            
            vn = modDecs.TriangleNormal(V(0), V(1), V(2))
            
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
    
        obj.CollideObject = lngObjCount
    
        lngObjCount = lngObjCount + 1
    End If
    'Debug.Print obj.Key; obj.CollideObject
    
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
    Dim V() As D3DVECTOR

    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    Dim vn As D3DVECTOR

    ReDim V(0 To 3) As D3DVECTOR

    AddCollisionEx = lngFaceCount

    Dim FaceCount As Long
    Dim addingFace As Boolean
    
    Index = 0
    For Face = 0 To NumFaces - 1

        For cnt = 0 To 2
            
            V(cnt).X = Verticies(Index + cnt).X
            V(cnt).Y = Verticies(Index + cnt).Y
            V(cnt).Z = Verticies(Index + cnt).Z
                        
        Next
        
        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single
    
        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
        
        vn = modDecs.TriangleNormal(V(0), V(1), V(2))

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
    
    lngObjCount = lngObjCount + 1
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub RenderPortals()
     
    Dim cnt As Long
    Dim cnt2 As Long
    cnt = 1
    
    Do While cnt <= Portals.Count
        
        If Portals(cnt).Enabled Then
        
            'If Not modParse.Player.Element Is Nothing Then RenderPortals2 Portals(cnt), modParse.Player.Element
            
            cnt2 = 1
            Do While cnt2 <= Elements.Count And cnt <= Portals.Count
            
                RenderPortals2 Portals(cnt), Elements(cnt2)
                
                cnt2 = cnt2 + 1
            Loop
            
        End If
        
        cnt = cnt + 1
    Loop

End Sub

Private Sub RenderPortals2(ByRef t1 As Portal, ByRef e1 As Element)
On Error GoTo scripterror
       
    Dim pos As D3DVECTOR
    
    Dim cnt3 As Long
    Dim cnt2 As Long
    
    Dim cnt As Long
    Dim obj As Long
                            
    Dim A As Long
    Dim act As Motion
    Dim txtobj As String
    Dim errline As Long
    Dim errsource As String
    Dim portalHit As Boolean

    Dim e2 As Element
        
    portalHit = (DistanceEx(e1.Origin, t1.Location) <= t1.Range)
    
    If (Not (e1.Fulcrums Is Nothing)) And (Not portalHit) Then
        For cnt = 1 To e1.Fulcrums.Count
        
            'portalHit = (DistanceEx(VectorRotateAxis(e1.Fulcrums(cnt), VectorMultiplyBy(e1.Rotate, RADIAN)), t1.Location) <= t1.Range)
            portalHit = (DistanceEx(e1.Fulcrums(cnt), t1.Location) <= t1.Range)
            If portalHit Then Exit For
        Next
    End If
    
    If portalHit Then
        If Not t1.Teleport.Equals(NoPoint) Then
            pos = ToVector(e1.Origin)
            
            cnt = 1
            Do While cnt <= Elements.Count
            
                Set e2 = Elements(cnt)
                
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
                
                Set e2 = Nothing
                cnt = cnt + 1
            Loop
            
            e1.Origin = t1.Teleport
            If e1.Collision Then
               ' If TestCollision(e1, Actions.NotDefined, 2) Then
                
                If TestCollision(e1, Actions.NotDefined, 1) Then
                     'revert if collision occurs after a teleport
                    Set e1.Origin = ToPoint(pos)
                End If
            End If
        End If
        
        If Not t1.OnInRange Is Nothing Then
        
            If InStr(LCase(t1.OnInRange.AppliesTo) & ",", LCase(e1.Key) & ",") > 0 Or t1.OnInRange.AppliesTo = "" Then
            
                If (((t1.OnInRange.EventFlags And 1) <> 1) And (t1.OnInRange.Behavior = Press)) Or (t1.OnInRange.Behavior = Rapid) Then
                
                    If t1.DropsMotions Then
                        e1.ClearMotions
                    End If
                        
                    If (((t1.OnInRange.EventFlags And 1) <> 1) And (t1.OnInRange.Behavior <> Rapid)) Then
                        t1.OnInRange.EventFlags = t1.OnInRange.EventFlags + 1
                        If (t1.OnInRange.Behavior = Locks) Then
                            If ((t1.OnInRange.EventFlags And 2) = 2) Then
                                t1.OnInRange.EventFlags = t1.OnInRange.EventFlags - 2
                            Else
                                t1.OnInRange.EventFlags = t1.OnInRange.EventFlags + 2
                            End If
                        End If
                    End If
                    
                    
                    If (t1.OnInRange.Behavior <> Locks) Or ((t1.OnInRange.Behavior = Locks) And ((t1.OnInRange.EventFlags And 2) = 2)) Then

                        errsource = "OnInRange"
                        errline = CLng(t1.OnInRange.StartLine)
                        frmMain.Run t1.OnInRange.RunScript, e1.Key, errline
                        'Debug.Print "OnInRange " & t1.Key & " " & e1.Key
                    End If
                End If
                
            End If
            
        End If
        
        If Not t1.OnOutRange Is Nothing Then
            If InStr(LCase(t1.OnOutRange.AppliesTo) & ",", LCase(e1.Key) & ",") > 0 Or t1.OnOutRange.AppliesTo = "" Then
                If (((t1.OnOutRange.EventFlags And 1) = 1) And (t1.OnOutRange.Behavior <> Rapid)) Then
                    t1.OnOutRange.EventFlags = t1.OnOutRange.EventFlags - 1
                End If
            End If
        End If

        If Not t1.Motions Is Nothing Then
            If t1.Motions.Count > 0 Then
                For A = 1 To t1.Motions.Count
                    Set act = t1.Motions(A)
                    e1.AddMotion act.Action, act.Key, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.Script
                Next
            End If
        End If
                    
    Else
        If Not t1.OnOutRange Is Nothing Then
        
            If InStr(LCase(t1.OnOutRange.AppliesTo) & ",", LCase(e1.Key) & ",") > 0 Or t1.OnOutRange.AppliesTo = "" Then

                If (((t1.OnOutRange.EventFlags And 1) <> 1) And (t1.OnOutRange.Behavior = Press)) Or (t1.OnOutRange.Behavior = Rapid) Then
                    
                    If (((t1.OnOutRange.EventFlags And 1) <> 1) And (t1.OnOutRange.Behavior <> Rapid)) Then
                        t1.OnOutRange.EventFlags = t1.OnOutRange.EventFlags + 1
                        If (t1.OnOutRange.Behavior = Locks) Then
                            If ((t1.OnOutRange.EventFlags And 2) = 2) Then
                                t1.OnOutRange.EventFlags = t1.OnOutRange.EventFlags - 2
                            Else
                                t1.OnOutRange.EventFlags = t1.OnOutRange.EventFlags + 2
                            End If
                        End If
                    End If
                    
                    If (t1.OnOutRange.Behavior <> Locks) Or ((t1.OnOutRange.Behavior = Locks) And ((t1.OnOutRange.EventFlags And 2) = 2)) Then
                        errsource = "OnOutRange"
                        errline = CLng(t1.OnOutRange.StartLine)
                        frmMain.Run t1.OnOutRange.RunScript, e1.Key, errline
                        'Debug.Print "OnOutRange " & t1.Key & " " & e1.Key
                    End If
                    
                End If

            End If
            
        End If
        
        If Not t1.OnInRange Is Nothing Then
            If InStr(LCase(t1.OnInRange.AppliesTo) & ",", LCase(e1.Key) & ",") > 0 Or t1.OnInRange.AppliesTo = "" Then
                If (((t1.OnInRange.EventFlags And 1) = 1) And (t1.OnInRange.Behavior <> Rapid)) Then
                    t1.OnInRange.EventFlags = t1.OnInRange.EventFlags - 1
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
        ConsoleCommand "echo An error " & Err.Number & " occurd in " & errsource & _
        "\nIn the event starting at Line: " & errline & _
        "\nError: " & Err.Description
        'frmMain.Print "echo An error " & Err.Number & " occurd in " & Err.Source & " at line " & (atLine - 1) & "\n" & Err.Description & "\n" & LastCall
        
        If frmMain.ScriptControl.Error.Number <> 0 Then frmMain.ScriptControl.Error.Clear
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
    If Not modParse.Player.Element Is Nothing Then
        If Cameras.Count > 0 Then
            Static toggle As Boolean
            toggle = Not toggle
            For cnt = IIf(toggle, 1, Cameras.Count) To IIf(toggle, Cameras.Count, 1) Step IIf(toggle, 1, -1)
                Dist = DistanceEx(modParse.Player.Element.Origin, Cameras(cnt).Origin)
                If ((Dist <= past) Or (past = 0)) And (InStr(Exclude, cnt & ",") = 0) Then
                    GetClosestCamera = cnt
                    past = Dist
                End If
            Next
        End If
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
    Dim V As D3DVECTOR
    Dim N As D3DVECTOR
    
    Dim verts(0 To 2) As D3DVECTOR
    Dim lastCam As Long
    'two quests about cameras
    '1 default projection should be in short range leainant not to turning camera around rather to a any range put projection variance in direction
    '2 movement from one camera to the Next could have a flying adaptation in a swing and out of the constructs way while it flies to genral Next 1
        
    If Perspective = Playmode.CameraMode Then
    
        If Not modParse.Player.Element Is Nothing Then
            If Cameras.Count > 0 Then
                
                If (Elements.Count > 0) Then
                    Dim e1 As Element
                    For cnt = 1 To Elements.Count
                        Set e1 = Elements(cnt)
                    
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
                        
                        Set e1 = Nothing
                    Next
                End If
    
                cnt = 0
                modParse.Player.CameraIndex = 0
    
                Do
                    
                    cnt = GetClosestCamera(ex)
                    
                    touched = False
                            
                    If (cnt > 0) Then
                        With Cameras(cnt)
                        
                            verts(0) = ToVector(modParse.Player.Element.Origin)
                            verts(1) = VectorAdd(ToVector(modParse.Player.Element.Origin), MakeVector(0, -0.01, 0))
                            verts(2) = ToVector(.Origin)
        
                            Face = AddCollisionEx(verts, 1)
                            touched = TestCollisionEx(Face, 1)
                            DelCollisionEx Face, 1
        
                            If (ClassifyPoint(V1, V1, V1, ToVector(modParse.Player.Element.Origin)) = 1) Then touched = True
        
        
                            If Not touched Then
                                
                                
                                V1 = VectorSubtract(MakeVector(.Origin.X + Sin(D720 - .Angle), _
                                                                                .Origin.Y - Tan(D720 - .Pitch), _
                                                                                .Origin.Z + Cos(D720 - .Angle)), _
                                                                                ToVector(.Origin))
                                                                                
                                V2 = VectorSubtract(MakeVector(modParse.Player.Element.Origin.X - Sin(D720 - .Angle), _
                                                                modParse.Player.Element.Origin.Y + Tan(D720 - .Pitch), _
                                                                modParse.Player.Element.Origin.Z - Cos(D720 - .Angle)), _
                                                                ToVector(.Origin))
                                
                                If ((V2.X > 0 And V1.X > 0) Or (V2.X < 0 And V1.X < 0)) And _
                                    ((V2.Y > 0 And V1.Y > 0) Or (V2.Y < 0 And V1.Y < 0)) And _
                                    ((V2.Z > 0 And V1.Z > 0) Or (V2.Z < 0 And V1.Z < 0)) Then
                                    touched = False
                                    
                                    If past <> 0 Then
                                        If DistanceEx(.Origin, modParse.Player.Element.Origin) > Dist Then
                                            cnt = past
                                            Dist = DistanceEx(.Origin, modParse.Player.Element.Origin)
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
                                    If DistanceEx(.Origin, modParse.Player.Element.Origin) > Dist Then
                                        cnt = past
                                        Dist = DistanceEx(.Origin, modParse.Player.Element.Origin)
                                        ex = ex & cnt & ", "
                                    End If
                                End If
        
                                If cnt >= 0 And cnt <= Cameras.Count Then
                                    modParse.Player.CameraIndex = cnt
                                    past = cnt
                                    Dist = DistanceEx(.Origin, modParse.Player.Element.Origin)
                                End If
                            Else
                                ex = ex & cnt & ", "
                            End If
                        End With
                    End If
    
                Loop Until (cnt = 0) Or (modParse.Player.CameraIndex <> 0)
                
                If modParse.Player.CameraIndex = 0 And Not lastCam = 0 Then
                    modParse.Player.CameraIndex = lastCam
                End If
                lastCam = modParse.Player.CameraIndex
            
            End If
            
        End If
        
    ElseIf (Not (modParse.Player.CameraIndex = 0)) Then
        If Not ((Perspective = Spectator) Or DebugMode) Then
            modParse.Player.CameraIndex = 0
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
    Dim V() As D3DVECTOR
    ReDim V(0 To VertexCount - 1) As D3DVECTOR

    For cnt = 0 To VertexCount - 1
        V(cnt) = MakeVector(sngVertexX(cnt, FaceIndex), sngVertexY(cnt, FaceIndex), sngVertexZ(cnt, FaceIndex))
        C.X = C.X + V(cnt).X
        C.Y = C.Y + V(cnt).Y
        C.Z = C.Z + V(cnt).Z
    Next
    
    C.X = C.X / VertexCount
    C.Y = C.Y / VertexCount
    C.Z = C.Z / VertexCount

    p = GetPlaneNormal(V(0), V(1), V(2))
        
    Dim N As Long
    Dim m As Long
    
    For N = 0 To VertexCount - 1
        
        A = modDecs.VectorNormalize(modDecs.VectorSubtract(V(N), C))
        
        smallest = -1
        smallestAngle = -1
        
        For m = N + 1 To 2
            If Not ClassifyPoint(V(N), C, VectorAdd(C, p), V(m)) = 2 Then 'not back
                B = modDecs.VectorNormalize(modDecs.VectorSubtract(V(m), C))
                
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
    
    A = GetPlaneNormal(V(0), V(1), V(2))
    B = p
    
    If modDecs.VectorDotProduct(A, B) < 0 Then
        ReverseFaceVertices FaceIndex, VertexCount
    End If
    
    sngFaceVis(0, FaceIndex) = A.X
    sngFaceVis(1, FaceIndex) = A.Y
    sngFaceVis(2, FaceIndex) = A.Z

End Sub

Public Function GetPlaneNormal(ByRef V0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR) As D3DVECTOR

    Dim vector1 As D3DVECTOR
    Dim vector2 As D3DVECTOR
    Dim Normal As D3DVECTOR
    Dim Length As Single

    '/*Calculate the Normal*/
    '/*Vector 1*/
    vector1.X = (V0.X - V1.X)
    vector1.Y = (V0.Y - V1.Y)
    vector1.Z = (V0.Z - V1.Z)

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
    Dim V As D3DVECTOR
    V.X = sngVertexX(FirstIndex, FaceIndex)
    V.Y = sngVertexY(FirstIndex, FaceIndex)
    V.Z = sngVertexZ(FirstIndex, FaceIndex)
    
    sngVertexX(FirstIndex, FaceIndex) = sngVertexX(SecondIndex, FaceIndex)
    sngVertexY(FirstIndex, FaceIndex) = sngVertexY(SecondIndex, FaceIndex)
    sngVertexZ(FirstIndex, FaceIndex) = sngVertexZ(SecondIndex, FaceIndex)

    sngVertexX(SecondIndex, FaceIndex) = V.X
    sngVertexY(SecondIndex, FaceIndex) = V.Y
    sngVertexZ(SecondIndex, FaceIndex) = V.Z
End Sub





