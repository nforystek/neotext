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
        sngFaceVis(2, cnt) = vn.z
    Next
End Sub

Public Sub SwapActivity(ByRef val1 As MyActivity, ByRef val2 As MyActivity)
    Dim tmp As MyActivity
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function SetActivity(ByRef act As MyActivity, ByRef Action As Actions, ByRef dat As D3DVECTOR, ByRef emp As Single) As String
    act.Identity = Replace(modGuid.GUID, "-", "")
    act.Action = Action
    act.Data = dat
    act.Emphasis = emp
    SetActivity = act.Identity
End Function

Public Function AddActivity(ByRef Obj As MyObject, ByRef Action As Actions, ByVal aGUID As String, ByRef Data As D3DVECTOR, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String
    Obj.ActivityCount = Obj.ActivityCount + 1
    ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
    With Obj.Activities(Obj.ActivityCount)
        .Identity = IIf(aGUID = "", Replace(modGuid.GUID, "-", ""), aGUID)
        .Action = Action
        .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .Latency = Timer
        .Recount = Recount
        .OnEvent = Script
        AddActivity = .Identity
    End With
End Function

Public Function AddActivityEx(ByRef Obj As MyPortal, ByRef Action As Actions, ByVal aGUID As String, ByRef Data As D3DVECTOR, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String
    Obj.ActivityCount = Obj.ActivityCount + 1
    ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
    With Obj.Activities(Obj.ActivityCount)
        .Identity = IIf(aGUID = "", Replace(modGuid.GUID, "-", ""), aGUID)
        .Action = Action
        .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .Latency = Timer
        .Recount = Recount
        .OnEvent = Script
        AddActivityEx = .Identity
    End With
End Function

Public Function ActivityExists(ByRef Obj As MyObject, ByVal MGUID As String) As Boolean
    Dim a As Long
    For a = 1 To Obj.ActivityCount
        If Obj.Activities(a).Identity = MGUID Then
            ActivityExists = True
            Exit Function
        End If
    Next
    ActivityExists = False
End Function
Public Function DeleteActivity(ByRef Obj As MyObject, ByVal MGUID As String) As Boolean
    Dim a As Long
    If Obj.ActivityCount > 0 Then
        If Obj.Activities(Obj.ActivityCount).Identity = MGUID Then
            Obj.ActivityCount = Obj.ActivityCount - 1
            If Obj.ActivityCount > 0 Then
                ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
            Else
                Erase Obj.Activities
            End If
            DeleteActivity = True
        Else
            For a = 1 To Obj.ActivityCount
                If Obj.Activities(a).Identity = MGUID Then
                    SwapActivity Obj.Activities(a), Obj.Activities(Obj.ActivityCount)
                    Obj.ActivityCount = Obj.ActivityCount - 1
                    If Obj.ActivityCount > 0 Then
                        ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
                    Else
                        Erase Obj.Activities
                    End If
                    DeleteActivity = True
                    Exit For
                End If
            Next
        End If
    End If
End Function

Public Function ValidActivity(ByRef Activity As MyActivity) As Boolean
    ValidActivity = (Activity.Identity <> "")
End Function

Public Function CalculateActivity(ByRef Activity As MyActivity, ByRef Action As Actions) As D3DVECTOR

    If (Activity.Action And Action) = Action Then
        
        If Activity.Friction <> 0 Then
            Activity.Emphasis = Activity.Emphasis - (Activity.Emphasis * Activity.Friction)
            If Activity.Emphasis < 0 Then
                Activity.Emphasis = 0
                Activity.Identity = ""
            End If
        End If

        If (Activity.Emphasis > 0.0001) Or (Activity.Emphasis < -0.0001) Then
            CalculateActivity.X = Activity.Data.X * Activity.Emphasis
            CalculateActivity.Y = Activity.Data.Y * Activity.Emphasis
            CalculateActivity.z = Activity.Data.z * Activity.Emphasis
        Else
            Activity.Emphasis = 0
        End If
    
    End If
    
End Function

Private Sub ApplyActivity(ByRef Obj As MyObject)
    Dim cnt As Long
    Dim cnt2 As Long
    Dim Offset As D3DVECTOR
    
    If ((Not (Perspective = Spectator)) And (Obj.CollideObject = Player.Object.CollideObject)) Or (Not (Obj.CollideObject = Player.Object.CollideObject)) Then
        
        If Obj.Gravitational Then
            If Not Obj.States.OnLadder Then
                If Obj.States.InLiquid Then
                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(LiquidGravityDirect, Directing)
                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(LiquidGravityRotate, Rotating)
                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(LiquidGravityScaled, Scaling)
                Else
                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(GlobalGravityDirect, Directing)
                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(GlobalGravityRotate, Rotating)
                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(GlobalGravityScaled, Scaling)
                End If
            End If
        End If
    End If
    If Obj.Effect = Collides.None Then
        If Obj.ActivityCount > 0 Then
            Dim a As Long
            For a = 1 To Obj.ActivityCount
                If ValidActivity(Obj.Activities(a)) Then
                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(Obj.Activities(a), Directing)
                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(Obj.Activities(a), Rotating)
                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(Obj.Activities(a), Scaling)
                End If
            Next
        End If
    End If
End Sub
Public Sub ResetMotion()
    Dim a As Long
    Dim o As Long
    Player.Object.Direct = MakeVector(0, 0, 0)
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            Objects(o).Direct = MakeVector(0, 0, 0)
        Next
    End If
End Sub
Public Sub ClearActivities()
    Dim o As Long
    Do Until Player.Object.ActivityCount = 0
        DeleteActivity Player.Object, Player.Object.Activities(1).Identity
    Loop
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            Do Until Objects(o).ActivityCount = 0
                DeleteActivity Objects(o), Objects(o).Activities(1).Identity
            Loop
        Next
    End If
End Sub
Public Sub RenderActive()
On Error GoTo ObjectError

    Dim d As Boolean
    Dim o As Long
    Dim a As Long
    Dim act As MyActivity
    Dim trig As String
    Dim line As String
    Dim id As String
     
    Do
    Loop Until (Not DeleteActivity(Player.Object, ""))
        
    If Player.Object.Visible Then
        ApplyActivity Player.Object

        If Player.Object.ActivityCount > 0 Then

            a = 1
            Do While a <= Player.Object.ActivityCount
                If Player.Object.Activities(a).Reactive > -1 Then
                    If (Timer - Player.Object.Activities(a).Latency) > Player.Object.Activities(a).Reactive Then ' And Player.Object.Activities(a).Reactive > 0 Then
                        Player.Object.Activities(a).Latency = Timer
                        act = Player.Object.Activities(a)
                        act.Emphasis = act.Initials
                        DeleteActivity Player.Object, act.Identity
                        If Not act.OnEvent = "" Then
                            line = NextArg(act.OnEvent, ":")
                            trig = RemoveArg(act.OnEvent, ":")
                            If Left(Trim(trig), 1) = "<" Then
                                id = RemoveQuotedArg(trig, "<", ">") & ","
                                If ((InStr(id, Player.Object.Identity & ",") > 0) And (Player.Object.Identity <> "")) Or (id = ",") Then
                                    ParseLand line, trig
                                End If
                            Else
                                ParseLand line, trig
                            End If
                        End If
                        If act.Recount > -1 Then
                            If act.Recount > 0 Then
                                act.Recount = act.Recount - 1
                                AddActivity Player.Object, act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                            End If
                        Else
                            AddActivity Player.Object, act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                        End If
                        a = a + 1

                    Else
                        a = a + 1
                    End If
                    
                ElseIf ((Player.Object.Activities(a).Emphasis = 0) Or (Player.Object.Activities(a).Recount = 0)) And (Not Player.Object.Activities(a).Reactive = -1) Then
                    DeleteActivity Player.Object, Player.Object.Activities(a).Identity
                Else
                    a = a + 1
                End If
            Loop
    
        End If
    End If

    
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            Do
            Loop Until (Not DeleteActivity(Objects(o), ""))
            
            If Objects(o).Visible Then
                ApplyActivity Objects(o)

                If Objects(o).ActivityCount > 0 Then
                    a = 1
                    Do While a <= Objects(o).ActivityCount
                        If Objects(o).Activities(a).Reactive > -1 Then
                            If (Timer - Objects(o).Activities(a).Latency) > Objects(o).Activities(a).Reactive Then 'And Objects(o).Activities(a).Reactive > 0 Then
                                Objects(o).Activities(a).Latency = Timer
                                act = Objects(o).Activities(a)
                                act.Emphasis = act.Initials
                                DeleteActivity Objects(o), act.Identity
                                If Not act.OnEvent = "" Then
                                    line = NextArg(act.OnEvent, ":")
                                    trig = RemoveArg(act.OnEvent, ":")
                                    If Left(Trim(trig), 1) = "<" Then
                                        id = RemoveQuotedArg(trig, "<", ">") & ","
                                        If ((InStr(id, Objects(o).Identity & ",") > 0) And (Objects(o).Identity <> "")) Or (id = ",") Then
                                            ParseLand line, trig
                                        End If
                                    Else
                                        ParseLand line, trig
                                    End If
                                End If
                                If act.Recount > -1 Then
                                    If act.Recount > 0 Then
                                        act.Recount = act.Recount - 1
                                        AddActivity Objects(o), act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                                    End If
                                Else
                                    AddActivity Objects(o), act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                                End If
                                a = a + 1
                            Else
                                a = a + 1
                            End If
                            
                        ElseIf ((Objects(o).Activities(a).Emphasis = 0) Or (Objects(o).Activities(a).Recount = 0)) And (Not Objects(o).Activities(a).Reactive = -1) Then
                            DeleteActivity Objects(o), Objects(o).Activities(a).Identity
                        Else
                            a = a + 1
                        End If
                    Loop
                End If
            End If
            
        Next
    End If
    

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'    Resume
End Sub

Public Sub InputMove()
On Error GoTo ObjectError

    lFacesShown = 0 'lCulledFaces
    lMovingObjs = 0
    lngTestCalls = 0
      
    Dim cnt As Long
    Dim cnt2 As Long
                        
    If ((Perspective = Spectator) Or DebugMode) Then
    
        Player.Object.Origin.X = Player.Object.Origin.X + Player.Object.Direct.X
        Player.Object.Origin.Y = Player.Object.Origin.Y + Player.Object.Direct.Y
        Player.Object.Origin.z = Player.Object.Origin.z + Player.Object.Direct.z
        
'        If DebugMode Then
'
'            ReDim andCamera(0 To 2, 0 To 2) As Single
'
'            ReDim andFaceVis(0 To 5, 0 To lngFaceCount) As Single
'            ReDim andVertexX(0 To 2, 0 To lngFaceCount) As Single
'            ReDim andVertexY(0 To 2, 0 To lngFaceCount) As Single
'            ReDim andVertexZ(0 To 2, 0 To lngFaceCount) As Single
'
'            ReDim andScreenX(0 To 2, 0 To lngFaceCount) As Single
'            ReDim andScreenY(0 To 2, 0 To lngFaceCount) As Single
'            ReDim andScreenZ(0 To 2, 0 To lngFaceCount) As Single
'
'            ReDim andZBuffer(0 To 3, 0 To lngFaceCount) As Single
'
'            ReDim notCamera(0 To 2, 0 To 2) As Single
'
'            ReDim notFaceVis(0 To 5, 0 To lngFaceCount) As Single
'            ReDim notVertexX(0 To 2, 0 To lngFaceCount) As Single
'            ReDim notVertexY(0 To 2, 0 To lngFaceCount) As Single
'            ReDim notVertexZ(0 To 2, 0 To lngFaceCount) As Single
'
'            ReDim notScreenX(0 To 2, 0 To lngFaceCount) As Single
'            ReDim notScreenY(0 To 2, 0 To lngFaceCount) As Single
'            ReDim notScreenZ(0 To 2, 0 To lngFaceCount) As Single
'
'            ReDim notZBuffer(0 To 3, 0 To lngFaceCount) As Single
'
'            notFaceVis = sngFaceVis
'            notVertexX = sngVertexX
'            notVertexY = sngVertexY
'            notVertexZ = sngVertexZ
'            notScreenX = sngScreenX
'            notScreenY = sngScreenY
'            notScreenZ = sngScreenZ
'            notZBuffer = sngZBuffer
'
'            For cnt2 = 0 To lngFaceCount - 1
'                sngFaceVis(3, cnt2) = 0
'            Next
'
'            Player.Object.CulledFaces = 0
'
'            If CullingSetup > 0 Then
'                notCamera(0, 0) = CullingObject.Position.X
'                notCamera(0, 1) = CullingObject.Position.Y
'                notCamera(0, 2) = CullingObject.Position.z
'
'                notCamera(1, 0) = CullingObject.Direction.X
'                notCamera(1, 1) = CullingObject.Direction.Y
'                notCamera(1, 2) = CullingObject.Direction.z
'
'                notCamera(2, 0) = CullingObject.UpVector.X
'                notCamera(2, 1) = CullingObject.UpVector.Y
'                notCamera(2, 2) = CullingObject.UpVector.z
'
'                Select Case CullingObject.visType
'                Case CULL0
'                        For cnt = 0 To lngFaceCount - 1
'                            notFaceVis(3, cnt) = 0
'                        Next
'                    Case Else
'                        Player.Object.CulledFaces = Culling(CullingObject.visType, lngFaceCount, notCamera, notFaceVis, notVertexX, notVertexY, notVertexZ, notScreenX, notScreenY, notScreenZ, notZBuffer)
'                        For cnt2 = 0 To lngFaceCount - 1
'                            If notFaceVis(3, cnt2) > 1 Then notFaceVis(3, cnt2) = CullingObject.visType
'                        Next
'
'                End Select
'                lCullCalls = lCullCalls + 1
'
'            End If
'
'            If CullingCount > 0 Then
'                For cnt = 1 To CullingCount
'                    andCamera(0, 0) = Cullings(cnt).Position.X
'                    andCamera(0, 1) = Cullings(cnt).Position.Y
'                    andCamera(0, 2) = Cullings(cnt).Position.z
'
'                    andCamera(1, 0) = Cullings(cnt).Direction.X
'                    andCamera(1, 1) = Cullings(cnt).Direction.Y
'                    andCamera(1, 2) = Cullings(cnt).Direction.z
'
'                    andCamera(2, 0) = Cullings(cnt).UpVector.X
'                    andCamera(2, 1) = Cullings(cnt).UpVector.Y
'                    andCamera(2, 2) = Cullings(cnt).UpVector.z
'
'                    andFaceVis = sngFaceVis
'                    andVertexX = sngVertexX
'                    andVertexY = sngVertexY
'                    andVertexZ = sngVertexZ
'                    andScreenX = sngScreenX
'                    andScreenY = sngScreenY
'                    andScreenZ = sngScreenZ
'                    andZBuffer = sngZBuffer
'
'                    Select Case Cullings(cnt).visType
'                        Case 0
'                            For cnt2 = 0 To lngFaceCount - 1
'                                andFaceVis(3, cnt2) = 0
'                            Next
'                        Case Else
'                            Player.Object.CulledFaces = Culling(Cullings(cnt).visType, lngFaceCount, andCamera, andFaceVis, andVertexX, andVertexY, andVertexZ, andScreenX, andScreenY, andScreenZ, andZBuffer)
'
'                        For cnt2 = 0 To lngFaceCount - 1
'                            If andFaceVis(3, cnt2) > 1 Then andFaceVis(3, cnt2) = Cullings(cnt).visType
'                        Next
'
'                    End Select
'                    lCullCalls = lCullCalls + 1
'
'                    If cnt > 1 Then
'
'                        For cnt2 = 0 To lngFaceCount - 1
'                            If andFaceVis(3, cnt2) > 1 Then
'                                notFaceVis(3, cnt2) = andFaceVis(3, cnt2)
'
'                            End If
'
'                        Next
'
'                    End If
'
'                Next
'
'            End If
'
'            sngFaceVis = notFaceVis
'            sngVertexX = notVertexX
'            sngVertexY = notVertexY
'            sngVertexZ = notVertexZ
'            sngScreenX = notScreenX
'            sngScreenY = notScreenY
'            sngScreenZ = notScreenZ
'            sngZBuffer = notZBuffer
'
'        End If
        
    Else
    
        If (Player.Object.CollideIndex > -1) Then
        
            Dim oldorg As D3DVECTOR
            oldorg = Player.Object.Origin

            SpinObject Player.Object
            BlowObject Player.Object
            MoveObject Player.Object

            lFacesShown = lFacesShown + Player.Object.CulledFaces
            lMovingObjs = lMovingObjs + 1
            If (Distance(Player.Object.Origin, MakeVector(0, 0, 0)) > Player.Boundary) Then Player.Object.Origin = oldorg
            
        End If

    End If

    If (ObjectCount > 0) And (Not DebugMode) Then
        Dim a As Long
        Dim act As MyActivity
        For cnt = 1 To ObjectCount
            
            If (Objects(cnt).Effect = Collides.None) Then
                If (Objects(cnt).CollideIndex > -1) Then

                    SpinObject Objects(cnt)
                    BlowObject Objects(cnt)
                    MoveObject Objects(cnt)

                    lFacesShown = lFacesShown + Objects(cnt).CulledFaces
                    lMovingObjs = lMovingObjs + 1
                ElseIf (Objects(cnt).Direct.X <> 0) Or (Objects(cnt).Direct.Y <> 0) Or (Objects(cnt).Direct.z <> 0) Then
                    Objects(cnt).Origin.X = Objects(cnt).Origin.X + Objects(cnt).Direct.X
                    Objects(cnt).Origin.Y = Objects(cnt).Origin.Y + Objects(cnt).Direct.Y
                    Objects(cnt).Origin.z = Objects(cnt).Origin.z + Objects(cnt).Direct.z
                    
                    Objects(cnt).Rotate.X = Objects(cnt).Rotate.X + Objects(cnt).Twists.X
                    Objects(cnt).Rotate.Y = Objects(cnt).Rotate.Y + Objects(cnt).Twists.Y
                    Objects(cnt).Rotate.z = Objects(cnt).Rotate.z + Objects(cnt).Twists.z

                    Objects(cnt).Scaled.X = Objects(cnt).Scaled.X + Objects(cnt).Scalar.X
                    Objects(cnt).Scaled.Y = Objects(cnt).Scaled.Y + Objects(cnt).Scalar.Y
                    Objects(cnt).Scaled.z = Objects(cnt).Scaled.z + Objects(cnt).Scalar.z
                End If
    
                If (Objects(cnt).Origin.X > SpaceBoundary) Or (Objects(cnt).Origin.X < -SpaceBoundary) Then Objects(cnt).Origin.X = -Objects(cnt).Origin.X
                If (Objects(cnt).Origin.z > SpaceBoundary) Or (Objects(cnt).Origin.z < -SpaceBoundary) Then Objects(cnt).Origin.z = -Objects(cnt).Origin.z
            End If
        Next
    End If


    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    'Resume
End Sub

Public Function CoupleMove(ByRef Obj As MyObject, ByVal objCollision As Long) As Boolean
'###################################################################################
'########## couple the activities of objects in collision with others ##############
'###################################################################################

    Dim a As Long
    Dim cnt As Long
    Dim act As MyActivity
    If (objCollision > -1) Then
        If (ObjectCount > 0) Then
            For cnt = 1 To ObjectCount
                If (Objects(cnt).Effect = Collides.None) And (Obj.CollideIndex > -1) Then
                    If (Not Objects(cnt).CollideObject = Obj.CollideObject) Then
                        If (Objects(cnt).CollideObject = objCollision) Then

                            For a = 1 To Obj.ActivityCount
                                act = Obj.Activities(a)
                                AddActivity Objects(cnt), act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                            Next

                            Objects(cnt).Direct = Obj.Direct
                            CoupleMove = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    End If
End Function

Private Sub MoveObject(ByRef Obj As MyObject)
On Error GoTo ObjectError

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

'#####################################################################################
'############# preliminary sort the type of space collision checks ###################
'#####################################################################################

    Obj.States.IsMoving = Moving.None

    For cnt = 0 To lngFaceCount - 1
        sngFaceVis(3, cnt) = 0
    Next

    If (ObjectCount > 0) Then
        For cnt = 1 To ObjectCount
            If (Objects(cnt).Effect = Collides.Ladder) And (Objects(cnt).CollideIndex > -1) And Objects(cnt).Visible Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = bitType
                Next
            End If
        Next

        If Obj.States.OnLadder Then
            Obj.States.OnLadder = TestCollision(Obj, Actions.None, bitType)
        Else
            Obj.States.OnLadder = TestCollision(Obj, Actions.None, bitType)
            If Obj.States.OnLadder Then
                Do
                Loop Until Not DeleteActivity(Obj, JumpGUID)
                For cnt = 1 To PortalCount
                    For cnt2 = 1 To Portals(cnt).ActivityCount
                        DeleteActivity Obj, Portals(cnt).Activities(cnt2).Identity
                    Next

                Next
            End If
        End If

        For cnt = 1 To ObjectCount
            If (Objects(cnt).Effect = Collides.Liquid) And (Objects(cnt).CollideIndex > -1) And Objects(cnt).Visible Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = bitType
                Next
            End If
        Next

        If Obj.States.InLiquid Then
            Obj.States.InLiquid = TestCollision(Obj, Actions.None, bitType)
        Else
            Obj.States.InLiquid = TestCollision(Obj, Actions.None, bitType)
            If Obj.States.InLiquid Then
                Do
                Loop Until Not DeleteActivity(Obj, JumpGUID)
                For cnt = 1 To PortalCount
                    For cnt2 = 1 To Portals(cnt).ActivityCount
                        DeleteActivity Obj, Portals(cnt).Activities(cnt2).Identity
                    Next
                Next
            End If
        End If

    End If


'#####################################################################################
'############# initial faces data for returning TestCollision info ###################
'#####################################################################################

    sngCamera(0, 0) = Obj.Origin.X
    sngCamera(0, 1) = Obj.Origin.Y + 1
    sngCamera(0, 2) = Obj.Origin.z

    sngCamera(1, 0) = 1
    sngCamera(1, 1) = -1
    sngCamera(1, 2) = -1

    sngCamera(2, 0) = -1
    sngCamera(2, 1) = 1
    sngCamera(2, 2) = -1

    Obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
    lCullCalls = lCullCalls + 1

    If (ObjectCount > 0) Then
        For cnt = 1 To ObjectCount
            If (Objects(cnt).Effect = Collides.Ladder) And (Objects(cnt).CollideIndex > -1) Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0
                Next
            ElseIf (Objects(cnt).Effect = Collides.Ground) And (Objects(cnt).CollideIndex > -1) And Objects(cnt).Visible Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    If Not (((sngFaceVis(0, cnt2) = 0) Or (sngFaceVis(0, cnt2) = 1) Or (sngFaceVis(0, cnt2) = -1)) And _
                        ((sngFaceVis(1, cnt2) = 0) Or (sngFaceVis(1, cnt2) = 1) Or (sngFaceVis(1, cnt2) = -1)) And _
                        ((sngFaceVis(2, cnt2) = 0) Or (sngFaceVis(2, cnt2) = 1) Or (sngFaceVis(2, cnt2) = -1))) Then
                        sngFaceVis(3, cnt2) = visType
                    End If
                Next
            ElseIf (Objects(cnt).Effect = Collides.Liquid) And (Objects(cnt).CollideIndex > -1) Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0
                Next
            End If
        Next
    End If
    

'#####################################################################################
'############# predict the Y movements of objects in motion ##########################
'#####################################################################################

    tmpset = Obj.Direct

    Obj.Direct.Y = tmpset.Y
    Obj.Direct.X = 0
    Obj.Direct.z = 0

    If (Obj.Direct.Y <> 0) Then
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y
            If Obj.Direct.Y > 0 Then
                If Not ((Obj.States.IsMoving And Moving.Flying) = Moving.Flying) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Flying
            ElseIf Obj.Direct.Y < 0 Then
                If Not ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Falling
            End If
            newset.Y = Obj.Direct.Y
        ElseIf (Obj.Direct.Y < 0) Then
            Do
                Obj.Direct.Y = Obj.Direct.Y + adjust
                If (Obj.Direct.Y >= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Y < 0) Then
                Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y
                If Not ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Falling
                newset.Y = Obj.Direct.Y
            End If
        ElseIf (Obj.Direct.Y > 0) Then
            Do
                Obj.Direct.Y = Obj.Direct.Y - adjust
                If (Obj.Direct.Y <= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.Y > 0) Then
                Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y
                If Not ((Obj.States.IsMoving And Moving.Flying) = Moving.Flying) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Flying
                newset.Y = Obj.Direct.Y
            End If
        End If
    End If
    

'#####################################################################################
'############# adjust face data based on the TestCollision resulted ##################
'#####################################################################################


    If (ObjectCount > 0) Then
        For cnt = 1 To ObjectCount
            If (Objects(cnt).CollideObject = Obj.CollideObject) And (Objects(cnt).CollideIndex > -1) Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = visType
                Next
            ElseIf (Objects(cnt).Effect = Collides.Ladder) And (Objects(cnt).CollideIndex > -1) Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0
                Next
            ElseIf (Objects(cnt).Effect = Collides.Liquid) And (Objects(cnt).CollideIndex > -1) Then
                For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                    sngFaceVis(3, cnt2) = 0
                Next
            End If
        Next
    End If
    
'#####################################################################################
'############# last call to MoveObejct collisions couple activity here ###############
'#####################################################################################


    CoupleMove Obj, objCollision
    
    
'#####################################################################################
'############# predict the X movements of objects in motion ##########################
'#####################################################################################

    Obj.Direct.Y = 0
    Obj.Direct.X = tmpset.X

    If (Obj.Direct.X <> 0) Then
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
            If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
            If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
            End If
            newset.X = Obj.Direct.X
        ElseIf (Obj.Direct.X < 0) Then
            Do
                Obj.Direct.X = Obj.Direct.X + adjust
                If (Obj.Direct.X >= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.X < 0) Then
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
                End If
                newset.X = Obj.Direct.X
            End If
        ElseIf (Obj.Direct.X > 0) Then
            Do
                Obj.Direct.X = Obj.Direct.X - adjust
                If (Obj.Direct.X <= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.X > 0) Then
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
                End If
                newset.X = Obj.Direct.X
            End If
        End If
    End If

'#####################################################################################
'############# predict the Z movements of objects in motion ##########################
'#####################################################################################

    Obj.Direct.X = 0
    Obj.Direct.z = tmpset.z

    If (Obj.Direct.z <> 0) Then
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
            If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
            If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
            End If
            newset.z = Obj.Direct.z
        ElseIf (Obj.Direct.z < 0) Then
            Do
                Obj.Direct.z = Obj.Direct.z + adjust
                If (Obj.Direct.z >= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.z < 0) Then
                Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
                End If
                newset.z = Obj.Direct.z
            End If
        ElseIf (Obj.Direct.z > 0) Then
            Do
                Obj.Direct.z = Obj.Direct.z - adjust
                If (Obj.Direct.z <= 0) Then Exit Do
            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
            If (Obj.Direct.z > 0) Then
                Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                If (tmpset.X <> newset.X) And (tmpset.z <> newset.z) And (Not (tmpset.Y = newset.Y)) And (Not tmpset.Y = 0) Then
                    If ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Falling
                End If
                newset.z = Obj.Direct.z
            End If
        End If
    End If


    Obj.Direct = newset

'#####################################################################################
'############# before applying predictions couple activities of touching #############
'#####################################################################################

    
    CoupleMove Obj, objCollision
    

'#####################################################################################
'############# push/pull of moving objects in Y slope and small step ups #############
'#####################################################################################

    If (Not Obj.States.IsMoving = Moving.None) And _
        (tmpset.X <> newset.X Or tmpset.z <> newset.z) And _
        (Not ((Obj.States.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling)) Then

        Obj.Origin.Y = Obj.Origin.Y + 0.2

        Obj.Direct.Y = 0
        Obj.Direct.X = tmpset.X
        Obj.Direct.z = tmpset.z

        push = True
        pull = False

        If (Obj.Direct.X <> 0) Or (Obj.Direct.z <> 0) Then
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                If Obj.Direct.X <> 0 Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    newset.X = Obj.Direct.X
                    pull = True
                End If
                If Obj.Direct.z <> 0 Then
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    newset.z = Obj.Direct.z
                    pull = True
                End If
            ElseIf (Obj.Direct.X < 0) And (Obj.Direct.z < 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    Obj.Direct.z = Obj.Direct.z + adjust
                    If ((Obj.Direct.X >= 0) Or (Obj.Direct.z >= 0)) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) And (Obj.Direct.z < 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.z = Obj.Direct.z
                    pull = True
                End If

            ElseIf (Obj.Direct.X > 0) And (Obj.Direct.z > 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    Obj.Direct.z = Obj.Direct.z - adjust
                    If ((Obj.Direct.X <= 0) Or (Obj.Direct.z <= 0)) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) And (Obj.Direct.z > 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.z = Obj.Direct.z
                    pull = True
                End If

            ElseIf (Obj.Direct.X < 0) And (Obj.Direct.z > 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    Obj.Direct.z = Obj.Direct.z - adjust
                    If ((Obj.Direct.X >= 0) Or (Obj.Direct.z <= 0)) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) And (Obj.Direct.z > 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.z = Obj.Direct.z
                    pull = True
                End If
            ElseIf (Obj.Direct.X > 0) And (Obj.Direct.z < 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    Obj.Direct.z = Obj.Direct.z + adjust
                    If ((Obj.Direct.X <= 0) Or (Obj.Direct.z >= 0)) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) And (Obj.Direct.z < 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    newset.z = Obj.Direct.z
                    pull = True
                End If
            End If
        End If

        Obj.Origin.Y = Obj.Origin.Y - 0.2

        If pull Then push = False

    End If

    Obj.Direct = tmpset

'#####################################################################################
'############# those passing with out pressure couple activities first ###############
'#####################################################################################

    
    CoupleMove Obj, objCollision


'#####################################################################################
'############# as an object first in motions continues it's push in moved Y ##########
'#####################################################################################


    If push And (Not Obj.States.IsMoving = Moving.None) And _
        (tmpset.X <> newset.X Or tmpset.z <> newset.z) And _
        (Not ((Obj.States.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling)) Then

        Obj.Origin.Y = Obj.Origin.Y + 0.2

        Obj.Direct.Y = 0
        Obj.Direct.X = tmpset.X
        Obj.Direct.z = tmpset.z

        push = False

        If (Obj.Direct.X <> 0) Then
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                newset.X = Obj.Direct.X
                push = True
            ElseIf (Obj.Direct.X < 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X + adjust
                    If (Obj.Direct.X >= 0) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X < 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    push = True
                End If

            ElseIf (Obj.Direct.X > 0) Then
                Do
                    Obj.Direct.X = Obj.Direct.X - adjust
                    If (Obj.Direct.X <= 0) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.X > 0) Then
                    Obj.Origin.X = Obj.Origin.X + Obj.Direct.X
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.X = Obj.Direct.X
                    push = True
                End If
            End If
        End If

        If (Obj.Direct.z <> 0) Then
            If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
                Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                newset.z = Obj.Direct.z
                push = True
            ElseIf (Obj.Direct.z < 0) Then
                Do
                    Obj.Direct.z = Obj.Direct.z + adjust
                    If (Obj.Direct.z >= 0) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.z < 0) Then
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.z = Obj.Direct.z
                    push = True
                End If

            ElseIf (Obj.Direct.z > 0) Then
                Do
                    Obj.Direct.z = Obj.Direct.z - adjust
                    If (Obj.Direct.z <= 0) Then Exit Do
                Loop Until (TestCollision(Obj, Directing, visType, objCollision) = False)
                If (Obj.Direct.z > 0) Then
                    Obj.Origin.z = Obj.Origin.z + Obj.Direct.z
                    If Not ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Level
                    newset.z = Obj.Direct.z
                    push = True
                End If

            End If
        End If

        Obj.Origin.Y = Obj.Origin.Y - 0.2

    End If


'#####################################################################################
'############# coupled in if pushing or pulling, adjust the X/Z gliding ##############
'#####################################################################################


    If (pull Xor push) And (Not ((Obj.States.IsMoving And Moving.Flying) = Moving.Flying)) And _
        (Not ((Obj.States.IsMoving And Moving.Falling) = Moving.Falling)) And _
        ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then

        Obj.Direct.Y = 0
        Obj.Direct.X = 0
        Obj.Direct.z = 0

        Do While (TestCollision(Obj, Directing, visType, objCollision) = True)
            Obj.Direct.Y = Obj.Direct.Y + adjust
        Loop

        If ((Obj.Direct.Y >= 0) And (Obj.Direct.Y < 0.3)) Or _
            ((Obj.Direct.Y >= 0) And (Obj.Direct.Y <= 0.2)) Then

            Obj.Origin.Y = Obj.Origin.Y + Obj.Direct.Y
            If Not ((Obj.States.IsMoving And Moving.Stepping) = Moving.Stepping) Then Obj.States.IsMoving = Obj.States.IsMoving + Moving.Stepping
            If ((Obj.States.IsMoving And Moving.Level) = Moving.Level) Then Obj.States.IsMoving = Obj.States.IsMoving - Moving.Level
            newset.Y = Obj.Direct.Y
        End If

    ElseIf ((Obj.States.IsMoving = Moving.None) And ((tmpset.X = 0 And tmpset.z = 0) And (newset.X = 0 And newset.z = 0))) Then

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
        Obj.Direct.z = adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            push = True
        Else
            push = False
            Obj.Direct.Y = 0
            Obj.Direct.z = 0
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
        Obj.Direct.z = -adjust
        If (TestCollision(Obj, Directing, visType, objCollision) = False) Then
            push = (pull And Not push) Or (Not push And Not pull)
        Else
            Obj.Direct.Y = 0
            Obj.Direct.z = 0
        End If

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

                    Obj.Origin.z = Obj.Origin.z + adjust
                    If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                        Obj.Origin.z = Obj.Origin.z - (adjust * 2)
                        If (TestCollision(Obj, Directing, visType, objCollision) = True) Then
                            Obj.Origin.Y = Obj.Origin.Y + (adjust / 3)
                        Else
                            Do
                                If Obj.Origin.z + (adjust / 3) <> adjust Then Exit Do
                                Obj.Origin.z = Obj.Origin.z + (adjust / 3)
                            Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                            Obj.Origin.z = Obj.Origin.z - (adjust / 3)
                        End If
                    Else
                        Do
                            If Obj.Origin.z - (adjust / 3) <> adjust Then Exit Do
                            Obj.Origin.z = Obj.Origin.z - (adjust / 3)
                        Loop Until (TestCollision(Obj, Directing, visType, objCollision) = True)
                        Obj.Origin.z = Obj.Origin.z + (adjust / 3)
                    End If
                End If

            Loop While (TestCollision(Obj, Directing, visType, objCollision) = True)

        End If

    End If


'#####################################################################################
'############# direct activities are primed for next call to MoveObject  #############
'#####################################################################################


    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   ' Resume
End Sub

Private Sub SpinObject(ByRef Obj As MyObject)
On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################


    If Not TestCollision(Obj, Rotating, 2) Then
            
        Obj.Rotate.X = Obj.Rotate.X + Obj.Twists.X
        Obj.Rotate.Y = Obj.Rotate.Y + Obj.Twists.Y
        Obj.Rotate.z = Obj.Rotate.z + Obj.Twists.z
   
    End If

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub BlowObject(ByRef Obj As MyObject)
On Error GoTo ObjectError

'#####################################################################################
'############# nothing as fancy as MoveObject for FPS rate/play vs. needs  ###########
'#####################################################################################


    If Not TestCollision(Obj, Scaling, 2) Then
    
        Obj.Scaled.X = Obj.Scaled.X + Obj.Scalar.X
        Obj.Scaled.Y = Obj.Scaled.Y + Obj.Scalar.Y
        Obj.Scaled.z = Obj.Scaled.z + Obj.Scalar.z
        
    End If

    
    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function TestCollision(ByRef Obj As MyObject, ByRef Action As Actions, ByVal visType As Long, Optional ByRef lngCollideObj As Long = -1) As Boolean
On Error GoTo ObjectError


'#####################################################################################
'############# face data is temporary transformed and checked for collision ##########
'#####################################################################################


    If Action = Rotating And (Not Obj.CollideIndex = -1) Then

'#####################################################################################
'############# in rotation collision we re-adjsut culling view direction #############
'#####################################################################################

        sngCamera(0, 0) = Obj.Origin.X
        sngCamera(0, 1) = Obj.Origin.Y + 1
        sngCamera(0, 2) = Obj.Origin.z

        sngCamera(1, 0) = 1
        sngCamera(1, 1) = -1
        sngCamera(1, 2) = -1

        sngCamera(2, 0) = -1
        sngCamera(2, 1) = 1
        sngCamera(2, 2) = -1

        Obj.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
        lCullCalls = lCullCalls + 1

    End If


'#####################################################################################
'############# create a transform matrix with the changes applied ####################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim v(2) As D3DVECTOR
    Dim n As D3DVECTOR

    Dim matScale As D3DMATRIX
    Dim matMesh As D3DMATRIX
    Dim matRot As D3DMATRIX
    
    D3DXMatrixIdentity matMesh
    D3DXMatrixIdentity matRot
    D3DXMatrixIdentity matScale

    If Action = Scaling Then
        D3DXMatrixScaling matScale, Obj.Scaled.X + Obj.Scalar.X, Obj.Scaled.Y + Obj.Scalar.Y, Obj.Scaled.z + Obj.Scalar.z
    Else
        D3DXMatrixScaling matScale, Obj.Scaled.X, Obj.Scaled.Y, Obj.Scaled.z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale
    
    If Action = Rotating Then

        D3DXMatrixRotationX matRot, ((Obj.Rotate.X + Obj.Twists.X) * (PI / 180))
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
        
        D3DXMatrixRotationY matRot, ((Obj.Rotate.Y + Obj.Twists.Y) * (PI / 180))
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
        
        D3DXMatrixRotationZ matRot, ((Obj.Rotate.z + Obj.Twists.z) * (PI / 180))
        D3DXMatrixMultiply matRot, matRot, matMesh
        D3DXMatrixMultiply matMesh, matRot, matMesh
    Else

        D3DXMatrixRotationX matRot, (Obj.Rotate.X * (PI / 180))
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationY matRot, (Obj.Rotate.Y * (PI / 180))
        D3DXMatrixMultiply matMesh, matRot, matMesh

        D3DXMatrixRotationZ matRot, (Obj.Rotate.z * (PI / 180))
        D3DXMatrixMultiply matMesh, matRot, matMesh

    End If

    If Action = Directing Then
        D3DXMatrixTranslation matScale, Obj.Origin.X + Obj.Direct.X, Obj.Origin.Y + Obj.Direct.Y, Obj.Origin.z + Obj.Direct.z
    Else
        D3DXMatrixTranslation matScale, Obj.Origin.X, Obj.Origin.Y, Obj.Origin.z
    End If
    D3DXMatrixMultiply matMesh, matMesh, matScale
    
            
    If lngFaceCount > 0 And Obj.CollideIndex <> -1 Then
    

'#####################################################################################
'############# update face data with the transformation matrix #######################
'#####################################################################################

        For Face = Obj.CollideIndex To (Obj.CollideIndex + Meshes(Obj.MeshIndex).Mesh.GetNumFaces) - 1
    
            For cnt = 0 To 2
                
                v(cnt).X = Meshes(Obj.MeshIndex).Verticies(Index + cnt).X
                v(cnt).Y = Meshes(Obj.MeshIndex).Verticies(Index + cnt).Y
                v(cnt).z = Meshes(Obj.MeshIndex).Verticies(Index + cnt).z
    
                D3DXVec3TransformCoord v(cnt), v(cnt), matMesh
                
                sngVertexX(cnt, Face) = v(cnt).X
                sngVertexY(cnt, Face) = v(cnt).Y
                sngVertexZ(cnt, Face) = v(cnt).z

            Next
            
            Index = Index + 3
        Next

'#####################################################################################
'############# per non culled face check and result collision ########################
'#####################################################################################

        Dim lngCollideIdx As Long
        lngCollideIdx = -1

        For cnt = Obj.CollideIndex To (Obj.CollideIndex + Meshes(Obj.MeshIndex).Mesh.GetNumFaces) - 1
            lngTestCalls = lngTestCalls + 1
            lFacesShown = lFacesShown + lngFaceCount
            If CBool(Collision(visType, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, cnt, lngCollideObj, lngCollideIdx)) Then
    
                TestCollision = True
                GoTo exitfunction
            End If
        Next

    End If
    TestCollision = False

exitfunction:

    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function TestCollisionEx(ByVal FaceNum As Long, ByVal visType As Long, Optional ByRef objCollision As Long = -1, Optional ByRef objFaceIndex As Long = -1) As Boolean
On Error GoTo ObjectError

'#####################################################################################
'############# to the point for simple triangle collsiion checking ###################
'#####################################################################################


    lngTestCalls = lngTestCalls + 1
    lFacesShown = lFacesShown + lngFaceCount
    If CBool(Collision(visType, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, FaceNum, objCollision, objFaceIndex)) Then
        TestCollisionEx = True
        Exit Function
    End If

    TestCollisionEx = False

    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function DelCollision(ByRef Obj As MyObject)
On Error GoTo ObjectError

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    Index = Meshes(Obj.MeshIndex).Mesh.GetNumFaces
    If Obj.CollideIndex + Index < lngFaceCount Then

        For Face = Obj.CollideIndex To Obj.CollideIndex + Index - 1
            sngFaceVis(0, Face) = sngFaceVis(0, Index + Face)
            sngFaceVis(1, Face) = sngFaceVis(1, Index + Face)
            sngFaceVis(2, Face) = sngFaceVis(2, Index + Face)
            sngFaceVis(3, Face) = sngFaceVis(3, Index + Face)
            sngFaceVis(4, Face) = sngFaceVis(4, Index + Face)
            sngFaceVis(5, Face) = sngFaceVis(5, Index + Face)
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
        
        For cnt = 1 To ObjectCount
            If Objects(cnt).CollideIndex > Obj.CollideIndex Then
                Objects(cnt).CollideIndex = Objects(cnt).CollideIndex - Index
            End If
        Next
        
        If Obj.CollideIndex + Index < lngFaceCount - 2 Then
        
            For Face = Obj.CollideIndex + Index To lngFaceCount - 2
                sngFaceVis(0, Face) = sngFaceVis(0, Face + 1)
                sngFaceVis(1, Face) = sngFaceVis(1, Face + 1)
                sngFaceVis(2, Face) = sngFaceVis(2, Face + 1)
                sngFaceVis(3, Face) = sngFaceVis(3, Face + 1)
                sngFaceVis(4, Face) = sngFaceVis(4, Face + 1)
                sngFaceVis(5, Face) = sngFaceVis(5, Face + 1)
                sngVertexX(0, Face) = sngVertexX(0, Face + 1)
                sngVertexX(1, Face) = sngVertexX(1, Face + 1)
                sngVertexX(2, Face) = sngVertexX(2, Face + 1)
                sngVertexY(0, Face) = sngVertexY(0, Face + 1)
                sngVertexY(1, Face) = sngVertexY(1, Face + 1)
                sngVertexY(2, Face) = sngVertexY(2, Face + 1)
                sngVertexZ(0, Face) = sngVertexZ(0, Face + 1)
                sngVertexZ(1, Face) = sngVertexZ(1, Face + 1)
                sngVertexZ(2, Face) = sngVertexZ(2, Face + 1)
                
                sngScreenX(0, Face) = sngScreenX(0, Face + 1)
                sngScreenX(1, Face) = sngScreenX(1, Face + 1)
                sngScreenX(2, Face) = sngScreenX(2, Face + 1)
                sngScreenY(0, Face) = sngScreenY(0, Face + 1)
                sngScreenY(1, Face) = sngScreenY(1, Face + 1)
                sngScreenY(2, Face) = sngScreenY(2, Face + 1)
                sngScreenZ(0, Face) = sngScreenZ(0, Face + 1)
                sngScreenZ(1, Face) = sngScreenZ(1, Face + 1)
                sngScreenZ(2, Face) = sngScreenZ(2, Face + 1)
                
                sngZBuffer(0, Face) = sngZBuffer(0, Face + 1)
                sngZBuffer(1, Face) = sngZBuffer(1, Face + 1)
                sngZBuffer(2, Face) = sngZBuffer(2, Face + 1)
                sngZBuffer(3, Face) = sngZBuffer(3, Face + 1)
            Next
        End If
        
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
    
    Exit Function
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function DelCollisionEx(ByVal CollideIndex As Long, ByVal NumFaces As Long)
On Error GoTo ObjectError

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    
    Index = NumFaces
    If CollideIndex + Index < lngFaceCount Then

        For Face = CollideIndex To CollideIndex + Index - 1
            sngFaceVis(0, Face) = sngFaceVis(0, Index + Face)
            sngFaceVis(1, Face) = sngFaceVis(1, Index + Face)
            sngFaceVis(2, Face) = sngFaceVis(2, Index + Face)
            sngFaceVis(3, Face) = sngFaceVis(3, Index + Face)
            sngFaceVis(4, Face) = sngFaceVis(4, Index + Face)
            sngFaceVis(5, Face) = sngFaceVis(5, Index + Face)
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
        
        For cnt = 1 To ObjectCount
            If Objects(cnt).CollideIndex > CollideIndex Then
                Objects(cnt).CollideIndex = Objects(cnt).CollideIndex - Index
            End If
        Next
        
        If CollideIndex + Index < lngFaceCount - 2 Then
        
            For Face = CollideIndex + Index To lngFaceCount - 2
                sngFaceVis(0, Face) = sngFaceVis(0, Face + 1)
                sngFaceVis(1, Face) = sngFaceVis(1, Face + 1)
                sngFaceVis(2, Face) = sngFaceVis(2, Face + 1)
                sngFaceVis(3, Face) = sngFaceVis(3, Face + 1)
                sngFaceVis(4, Face) = sngFaceVis(4, Face + 1)
                sngFaceVis(5, Face) = sngFaceVis(5, Face + 1)
                sngVertexX(0, Face) = sngVertexX(0, Face + 1)
                sngVertexX(1, Face) = sngVertexX(1, Face + 1)
                sngVertexX(2, Face) = sngVertexX(2, Face + 1)
                sngVertexY(0, Face) = sngVertexY(0, Face + 1)
                sngVertexY(1, Face) = sngVertexY(1, Face + 1)
                sngVertexY(2, Face) = sngVertexY(2, Face + 1)
                sngVertexZ(0, Face) = sngVertexZ(0, Face + 1)
                sngVertexZ(1, Face) = sngVertexZ(1, Face + 1)
                sngVertexZ(2, Face) = sngVertexZ(2, Face + 1)
                
                sngScreenX(0, Face) = sngScreenX(0, Face + 1)
                sngScreenX(1, Face) = sngScreenX(1, Face + 1)
                sngScreenX(2, Face) = sngScreenX(2, Face + 1)
                sngScreenY(0, Face) = sngScreenY(0, Face + 1)
                sngScreenY(1, Face) = sngScreenY(1, Face + 1)
                sngScreenY(2, Face) = sngScreenY(2, Face + 1)
                sngScreenZ(0, Face) = sngScreenZ(0, Face + 1)
                sngScreenZ(1, Face) = sngScreenZ(1, Face + 1)
                sngScreenZ(2, Face) = sngScreenZ(2, Face + 1)
                
                sngZBuffer(0, Face) = sngZBuffer(0, Face + 1)
                sngZBuffer(1, Face) = sngZBuffer(1, Face + 1)
                sngZBuffer(2, Face) = sngZBuffer(2, Face + 1)
                sngZBuffer(3, Face) = sngZBuffer(3, Face + 1)
            Next
        End If
        
    End If
    
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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function AddCollision(ByRef Obj As MyObject, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError

'#####################################################################################
'############# create face data for a mesh to external compatability #################
'#####################################################################################

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim v() As D3DVERTEX

    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim vn As D3DVECTOR

    ReDim v(0 To 3) As D3DVERTEX

    Obj.CollideIndex = lngFaceCount
    AddCollision = lngFaceCount

    Dim FaceCount As Long
    Dim addingFace As Boolean

    Index = 0
    For Face = 0 To Meshes(Obj.MeshIndex).Mesh.GetNumFaces - 1

        For cnt = 0 To 2

            v(cnt).X = Meshes(Obj.MeshIndex).Verticies(Meshes(Obj.MeshIndex).Indicies(Index + cnt)).X
            v(cnt).Y = Meshes(Obj.MeshIndex).Verticies(Meshes(Obj.MeshIndex).Indicies(Index + cnt)).Y
            v(cnt).z = Meshes(Obj.MeshIndex).Verticies(Meshes(Obj.MeshIndex).Indicies(Index + cnt)).z

            D3DXVec3TransformCoord vn, ConvertVertexToVector(v(cnt)), Obj.Matrix
            v(cnt).X = vn.X
            v(cnt).Y = vn.Y
            v(cnt).z = vn.z
        Next

        ReDim Preserve sngFaceVis(0 To 5, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngVertexZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngScreenX(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenY(0 To 2, 0 To lngFaceCount) As Single
        ReDim Preserve sngScreenZ(0 To 2, 0 To lngFaceCount) As Single

        ReDim Preserve sngZBuffer(0 To 3, 0 To lngFaceCount) As Single
        
        vn = TriangleNormal(ConvertVertexToVector(v(0)), ConvertVertexToVector(v(1)), ConvertVertexToVector(v(2)))
        
        For cnt = 0 To 2

            sngVertexX(cnt, lngFaceCount) = v(cnt).X
            sngVertexY(cnt, lngFaceCount) = v(cnt).Y
            sngVertexZ(cnt, lngFaceCount) = v(cnt).z

        Next

        sngFaceVis(0, lngFaceCount) = vn.X
        sngFaceVis(1, lngFaceCount) = vn.Y
        sngFaceVis(2, lngFaceCount) = vn.z
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

Public Function AddCollisionEx(ByRef Verticies() As D3DVECTOR, ByVal NumFaces As Long, Optional ByVal visType As Long = 0) As Long
On Error GoTo ObjectError

    Dim cnt As Long
    Dim Face As Long
    Dim Index As Long
    Dim v() As D3DVECTOR

    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
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
            v(cnt).z = Verticies(Index + cnt).z
                        
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
            sngVertexZ(cnt, lngFaceCount) = v(cnt).z

        Next

        sngFaceVis(0, lngFaceCount) = vn.X
        sngFaceVis(1, lngFaceCount) = vn.Y
        sngFaceVis(2, lngFaceCount) = vn.z
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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub RenderPortals()
On Error GoTo ObjectError

    Dim pos As D3DVECTOR
    
    Dim cnt3 As Long
    Dim cnt2 As Long
    
    Dim cnt As Long
    Dim Obj As Long
    
    Dim id As String
    Dim trig As String
    Dim line As String
                            
    Dim a As Long
    Dim act As MyActivity
    
    If PortalCount > 0 Then
        For cnt = 1 To PortalCount
'            ParseSetGet 0, "$beaconPortal" & cnt & ".x", Portals(cnt).Location.X
'            ParseSetGet 0, "$beaconPortal" & cnt & ".y", Portals(cnt).Location.Y
'            ParseSetGet 0, "$beaconPortal" & cnt & ".z", Portals(cnt).Location.z
            
            If Portals(cnt).Enable Then
                If (Distance(Player.Object.Origin, Portals(cnt).Location) <= Portals(cnt).Range) Then
    
                    If Not ((Portals(cnt).Teleport.X = 0) And (Portals(cnt).Teleport.Y = 0) And (Portals(cnt).Teleport.z = 0)) Then
                        pos = Player.Object.Origin
                        If ObjectCount > 0 Then
                            For cnt2 = 1 To ObjectCount
                                If Objects(cnt2).CollideIndex > -1 Then
                                    If Not Objects(cnt2).CollideIndex = Player.Object.CollideIndex And Objects(cnt2).Gravitational Then
                                        For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                            sngFaceVis(3, cnt3) = 1
                                        Next
                                    ElseIf Objects(cnt2).CollideIndex = Player.Object.CollideIndex Then
                                        For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                            sngFaceVis(3, cnt3) = 1
                                        Next
                                    Else
                                        For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                            sngFaceVis(3, cnt3) = 0
                                        Next
                                    End If
                                End If
                            Next
                        End If
                        Player.Object.Origin = Portals(cnt).Teleport
                        If TestCollision(Player.Object, Actions.None, 1) Then
                            Player.Object.Origin = pos
                            FadeMessage "Teleport error: simultaneous spawn attempt."
                        Else
                            Do While Player.Object.ActivityCount > 0
                                DeleteActivity Player.Object, Player.Object.Activities(1).Identity
                            Loop
                        End If
                    End If
                    If Portals(cnt).ClearActivities Then
                        Do While Player.Object.ActivityCount > 0
                            DeleteActivity Player.Object, Player.Object.Activities(1).Identity
                        Loop
                    End If
                    If Portals(cnt).ActivityCount > 0 Then
                        For a = 1 To Portals(cnt).ActivityCount
                            act = Portals(cnt).Activities(a)
                            AddActivity Player.Object, act.Action, Portals(cnt).Activities(a).Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                        Next
                    
                    End If
                    If Portals(cnt).OnInRange <> "" Then
                        line = NextArg(Portals(cnt).OnInRange, ":")
                        trig = RemoveArg(Portals(cnt).OnInRange, ":")
                        If Left(Trim(trig), 1) = "<" Then
                            id = RemoveQuotedArg(trig, "<", ">") & ","
                            If ((InStr(id, Player.Object.Identity & ",") > 0) And (Player.Object.Identity <> "")) Or (id = ",") Then
                                ParseLand CLng(line), trig
                            End If
                        Else
                            ParseLand CLng(line), trig
                        End If
                    End If
                Else
                    If Portals(cnt).OnOutRange <> "" Then
                        line = NextArg(Portals(cnt).OnOutRange, ":")
                        trig = RemoveArg(Portals(cnt).OnOutRange, ":")
                        If Left(Trim(trig), 1) = "<" Then
                            id = RemoveQuotedArg(trig, "<", ">") & ","
                            If ((InStr(id, Player.Object.Identity & ",") > 0) And (Player.Object.Identity <> "")) Or (id = ",") Then
                                ParseLand CLng(line), trig
                            End If
                        Else
                            ParseLand CLng(line), trig
                        End If
                    End If
                End If

                If ObjectCount > 0 Then
                
                    Dim txtobj As String
                        
                    For Obj = 1 To ObjectCount
                    
                        cnt3 = 0
                        If Objects(Obj).FolcrumCount > 0 Then
                          '  Static added As Boolean
                            
                            For cnt2 = 1 To Objects(Obj).FolcrumCount

                                
                                pos.X = Objects(Obj).Folcrum(cnt2).X
                                pos.Y = Objects(Obj).Folcrum(cnt2).Y
                                pos.z = Objects(Obj).Folcrum(cnt2).z

                    
                                D3DXVec3TransformCoord pos, pos, Objects(Obj).Matrix
                                
'                                If Not added Then
'                                    txtobj = "beacon" & vbCrLf & "{" & vbCrLf
'                                    txtobj = txtobj & "identity beaconFolcrum" & cnt2 & vbCrLf
'                                    txtobj = txtobj & "visible true" & vbCrLf
'                                    txtobj = txtobj & "percentxy 100 100" & vbCrLf
'                                    txtobj = txtobj & "origin " & CSng(pos.X) & " " & CSng(pos.Y) & " " & CSng(pos.z) & vbCrLf
'                                    txtobj = txtobj & "blackalpha" & vbCrLf
'                                    txtobj = txtobj & "filename bubble.bmp" & vbCrLf
'                                    txtobj = txtobj & "beaconlight 1" & vbCrLf
'                                    txtobj = txtobj & "verticallock" & vbCrLf
'                                    txtobj = txtobj & "}" & vbCrLf
'                                    ParseLand 0, txtobj
'                                Else
'                                    ParseSetGet 0, "$beaconFolcrum" & cnt2 & ".x", pos.X
'                                    ParseSetGet 0, "$beaconFolcrum" & cnt2 & ".y", pos.Y
'                                    ParseSetGet 0, "$beaconFolcrum" & cnt2 & ".z", pos.z
'
'                                End If
                                
                                cnt3 = CInt(Distance(MakeVector(pos.X, pos.Y, pos.z), Portals(cnt).Location) <= Portals(cnt).Range)
                                                
                                If cnt3 = -1 Then Exit For
                            Next
                          '  added = True
                            
                        ElseIf (Distance(Objects(Obj).Origin, Portals(cnt).Location) <= Portals(cnt).Range) Then
                            cnt3 = -1
                        End If

                        If cnt3 = -1 Then

                                If Not ((Portals(cnt).Teleport.X = 0) And (Portals(cnt).Teleport.Y = 0) And (Portals(cnt).Teleport.z = 0)) Then
                                    pos = Objects(Obj).Origin
                                    If ObjectCount > 0 Then
                                        For cnt2 = 1 To ObjectCount
                                            If Objects(cnt2).CollideIndex > -1 Then
                                                If Not Objects(cnt2).CollideIndex = Objects(Obj).CollideIndex And Objects(cnt2).Gravitational Then
                                                    For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                                        sngFaceVis(3, cnt3) = 1
                                                    Next
                                                ElseIf Objects(cnt2).CollideIndex = Objects(Obj).CollideIndex Then
                                                    For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                                        sngFaceVis(3, cnt3) = 1
                                                    Next
                                                Else
                                                    For cnt3 = Objects(cnt2).CollideIndex To (Objects(cnt2).CollideIndex + Meshes(Objects(cnt2).MeshIndex).Mesh.GetNumFaces) - 1
                                                        sngFaceVis(3, cnt3) = 0
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                    Objects(Obj).Origin = Portals(cnt).Teleport
                                    If TestCollision(Objects(Obj), Actions.None, 1) Then
                                        Objects(Obj).Origin = pos
                                    Else
                                        Do While Objects(Obj).ActivityCount > 0
                                            DeleteActivity Objects(Obj), Objects(Obj).Activities(1).Identity
                                        Loop
                                    End If
                                End If
                                If Portals(cnt).ClearActivities Then
                                    Do While Objects(Obj).ActivityCount > 0
                                        DeleteActivity Objects(Obj), Objects(Obj).Activities(1).Identity
                                    Loop
                                End If
                                If Portals(cnt).ActivityCount > 0 Then
                                   
                                    For a = 1 To Portals(cnt).ActivityCount
                                        act = Portals(cnt).Activities(a)
                                        AddActivity Objects(Obj), act.Action, Portals(cnt).Activities(a).Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
                                    Next
                                
                                End If
                                If (Portals(cnt).OnInRange <> "") Then
                                    line = NextArg(Portals(cnt).OnInRange, ":")
                                    trig = RemoveArg(Portals(cnt).OnInRange, ":")
                                    If Left(Trim(trig), 1) = "<" Then
                                        id = RemoveQuotedArg(trig, "<", ">") & ","
                                        If ((InStr(id, Objects(Obj).Identity & ",") > 0) And (Objects(Obj).Identity <> "")) Or (id = ",") Then
                                            ParseLand line, trig
                                        End If
                                    Else
                                        ParseLand line, trig
                                    End If
                                End If
                            Else
                                If (Portals(cnt).OnOutRange <> "") Then
                                    line = NextArg(Portals(cnt).OnOutRange, ":")
                                    trig = RemoveArg(Portals(cnt).OnOutRange, ":")
                                    If Left(Trim(trig), 1) = "<" Then
                                        id = RemoveQuotedArg(trig, "<", ">") & ","
                                        If ((InStr(id, Objects(Obj).Identity & ",") > 0) And (Objects(Obj).Identity <> "")) Or (id = ",") Then
                                            ParseLand line, trig
                                        End If
                                    Else
                                        ParseLand line, trig
                                    End If
                                End If
                            'End If
                        End If
                    Next
                End If
                
            End If
        Next
    End If

    Exit Sub
ObjectError:
    If Err.Number = 6 Or Err.Number = 11 Then Resume
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function GetClosestCamera(Optional ByVal Exclude As String = "") As Long

    Dim cnt As Long
    Dim Dist As Single
    Dim past As Single
    If CameraCount > 0 Then
        Static toggle As Boolean
        toggle = Not toggle
        For cnt = IIf(toggle, 1, CameraCount) To IIf(toggle, CameraCount, 1) Step IIf(toggle, 1, -1)
            Dist = Distance(Player.Object.Origin, Cameras(cnt).Location)
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
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    
    
    Dim pos As D3DVECTOR
    Dim touched As Boolean
    Dim Face As Long
    Dim ex As String
    
    Dim dot As Single
    Dim v As D3DVECTOR
    Dim n As D3DVECTOR
    
    Dim verts(0 To 2) As D3DVECTOR
    Dim lastCam As Long
    'two quests about cameras
    '1 default projection should be in short range leainant not to turning camera around rather to a any range put projection variance in direction
    '2 movement from one camera to the next could have a flying adaptation in a swing and out of the constructs way while it flies to genral next 1
        
    If Perspective = Playmode.CameraMode Then
    
        If CameraCount > 0 Then
            
            If (ObjectCount > 0) Then
                For cnt = 1 To ObjectCount
                    If ((Objects(cnt).Effect = Collides.Ground) Or (Objects(cnt).Effect = Collides.InDoor)) And (Objects(cnt).CollideIndex > -1) Then
                        For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
                            sngFaceVis(3, cnt2) = 1
                        Next
                    ElseIf (Objects(cnt).CollideIndex > -1) Then
                        For cnt2 = Objects(cnt).CollideIndex To (Objects(cnt).CollideIndex + Meshes(Objects(cnt).MeshIndex).Mesh.GetNumFaces) - 1
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

                    verts(0) = Player.Object.Origin
                    verts(1) = VectorAdd(Player.Object.Origin, MakeVector(0, -0.01, 0))
                    verts(2) = Cameras(cnt).Location

                    Face = AddCollisionEx(verts, 1)
                    touched = TestCollisionEx(Face, 1)
                    DelCollisionEx Face, 1

                    If (ClassifyPoint(v1, v1, v1, Player.Object.Origin) = 1) Then touched = True


                    If Not touched Then
                        
                        
                        v1 = VectorSubtract(MakeVector(Cameras(cnt).Location.X + Sin(D720 - Cameras(cnt).Angle), _
                                                                        Cameras(cnt).Location.Y - Tan(D720 - Cameras(cnt).Pitch), _
                                                                        Cameras(cnt).Location.z + Cos(D720 - Cameras(cnt).Angle)), _
                                                                        Cameras(cnt).Location)
                                                                        
                        v2 = VectorSubtract(MakeVector(Player.Object.Origin.X - Sin(D720 - Cameras(cnt).Angle), _
                                                        Player.Object.Origin.Y + Tan(D720 - Cameras(cnt).Pitch), _
                                                        Player.Object.Origin.z - Cos(D720 - Cameras(cnt).Angle)), _
                                                        Cameras(cnt).Location)
                        
                        If ((v2.X > 0 And v1.X > 0) Or (v2.X < 0 And v1.X < 0)) And _
                            ((v2.Y > 0 And v1.Y > 0) Or (v2.Y < 0 And v1.Y < 0)) And _
                            ((v2.z > 0 And v1.z > 0) Or (v2.z < 0 And v1.z < 0)) Then
                            touched = False
                            
                            If past <> 0 Then
                                If Distance(Cameras(cnt).Location, Player.Object.Origin) > Dist Then
                                    cnt = past
                                    Dist = Distance(Cameras(cnt).Location, Player.Object.Origin)
                                End If
                            End If

                        Else
                            touched = True
                        End If
                        If Not touched Then

                            dot = VectorDotProduct(v1, v2) / (VectorDotProduct(v1, v1) * VectorDotProduct(v2, v2))
                        End If
                    End If
                    
                    If Not touched Then
                        If past <> 0 Then
                            If Distance(Cameras(cnt).Location, Player.Object.Origin) > Dist Then
                                cnt = past
                                Dist = Distance(Cameras(cnt).Location, Player.Object.Origin)
                                ex = ex & cnt & ", "
                            End If
                        End If

                        If cnt >= 0 And cnt <= CameraCount Then
                            Player.CameraIndex = cnt
                            past = cnt
                            Dist = Distance(Cameras(cnt).Location, Player.Object.Origin)
                        End If
                    Else
                        ex = ex & cnt & ", "
                    End If
                    
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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SortVerticies(ByVal FaceIndex As Long, Optional ByVal VertexCount As Long = 3)
    Dim a As D3DVECTOR
    Dim b As D3DVECTOR
    Dim c As D3DVECTOR
    
    Dim p As D3DVECTOR
    
    Dim cnt As Long
    Dim Angle As Single
    
    Dim smallest As Long
    Dim smallestAngle As Single
    Dim v() As D3DVECTOR
    ReDim v(0 To VertexCount - 1) As D3DVECTOR

    For cnt = 0 To VertexCount - 1
        v(cnt) = MakeVector(sngVertexX(cnt, FaceIndex), sngVertexY(cnt, FaceIndex), sngVertexZ(cnt, FaceIndex))
        c.X = c.X + v(cnt).X
        c.Y = c.Y + v(cnt).Y
        c.z = c.z + v(cnt).z
    Next
    
    c.X = c.X / VertexCount
    c.Y = c.Y / VertexCount
    c.z = c.z / VertexCount

    p = GetPlaneNormal(v(0), v(1), v(2))
        
    Dim n As Long
    Dim m As Long
    
    For n = 0 To VertexCount - 1
        
        a = VectorNormalize(VectorSubtract(v(n), c))
        
        smallest = -1
        smallestAngle = -1
        
        For m = n + 1 To 2
            If Not ClassifyPoint(v(n), c, VectorAdd(c, p), v(m)) = 2 Then 'not back
                b = VectorNormalize(VectorSubtract(v(m), c))
                
                Angle = VectorDotProduct(a, b)
                
                If Angle > smallestAngle Then
                    smallestAngle = Angle
                    smallest = m
        
                End If
            End If
        Next
        
        If smallest = -1 Then Exit Sub
        
        If Not ((n + 1) = smallest) Then
            SwapVector FaceIndex, n + 1, smallest
        End If
    
    Next
    
    a = GetPlaneNormal(v(0), v(1), v(2))
    b = p
    
    If VectorDotProduct(a, b) < 0 Then
        ReverseFaceVertices FaceIndex, VertexCount
    End If
    
    sngFaceVis(0, FaceIndex) = a.X
    sngFaceVis(1, FaceIndex) = a.Y
    sngFaceVis(2, FaceIndex) = a.z

End Sub

Public Function GetPlaneNormal(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR

    Dim vector1 As D3DVECTOR
    Dim vector2 As D3DVECTOR
    Dim Normal As D3DVECTOR
    Dim Length As Single

    '/*Calculate the Normal*/
    '/*Vector 1*/
    vector1.X = (v0.X - v1.X)
    vector1.Y = (v0.Y - v1.Y)
    vector1.z = (v0.z - v1.z)

    '/*Vector 2*/
    vector2.X = (v1.X - v2.X)
    vector2.Y = (v1.Y - v2.Y)
    vector2.z = (v1.z - v2.z)

    '/*Apply the Cross Product*/
    Normal.X = (vector1.Y * vector2.z - vector1.z * vector2.Y)
    Normal.Y = (vector1.z * vector2.X - vector1.X * vector2.z)
    Normal.z = (vector1.X * vector2.Y - vector1.Y * vector2.X)

    '/*Normalize to a unit vector*/
    Length = Sqr(Normal.X * Normal.X + Normal.Y * Normal.Y + Normal.z * Normal.z)

    If Length = 0 Then Length = 1

    Normal.X = (Normal.X / Length)
    Normal.Y = (Normal.Y / Length)
    Normal.z = (Normal.z / Length)

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
    v.z = sngVertexZ(FirstIndex, FaceIndex)
    
    sngVertexX(FirstIndex, FaceIndex) = sngVertexX(SecondIndex, FaceIndex)
    sngVertexY(FirstIndex, FaceIndex) = sngVertexY(SecondIndex, FaceIndex)
    sngVertexZ(FirstIndex, FaceIndex) = sngVertexZ(SecondIndex, FaceIndex)

    sngVertexX(SecondIndex, FaceIndex) = v.X
    sngVertexY(SecondIndex, FaceIndex) = v.Y
    sngVertexZ(SecondIndex, FaceIndex) = v.z
End Sub

