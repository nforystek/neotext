Attribute VB_Name = "modMirror"
Option Explicit

Public Mirroring As Boolean

Private Mirrors As NTNodes10.Collection

Private MirrorPoint As D3DVECTOR
Private MirrorNormal As D3DVECTOR

Public Sub BeginMirrors()

    If Camera Is Nothing Then Exit Sub

    Dim e As Board
    Dim i As Long
    Dim L As Single

    Dim dm As D3DDISPLAYMODE
    Dim pal As PALETTEENTRY
    Dim rct As RECT
    Dim CullMode As Long


    
  '  CullMode = DDevice.GetRenderState(D3DRS_CULLMODE)
   ' DDevice.SetRenderState D3DRS_CULLMODE, IIf(CullMode = D3DCULL_CW, D3DCULL_CCW, IIf(CullMode = D3DCULL_CCW, D3DCULL_CW, D3DCULL_NONE))
    
    If Not Mirrors Is Nothing Then Mirrors.Clear

    If Boards.Count > 0 Then
        For i = 1 To Boards.Count
            Set e = Boards(i)

            If e.Visible And e.Mirror And PointSideOfPlane(e.Point1, e.Point2, e.Point3, Camera.Origin) Then


                L = Distance(Camera.Origin.X, Camera.Origin.Y, Camera.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                If L <= FAR Then

                    If Mirrors Is Nothing Then Set Mirrors = New NTNodes10.Collection

                    
                    
                    DViewPort.X = 0
                    DViewPort.Y = 0
                    DViewPort.Width = 128
                    DViewPort.Height = 128
                    
                    DSurface.BeginScene ReflectRenderTarget, DViewPort

                
                '#################################################################
                '#### SetupWorld configures the matricies used for DirectX 3D ####
                '#################################################################
                'elapsed = Timer
                SetupWorld
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
                
                
                '##################################################################
                '#### RenderMotion prepairs a subsets of preliminary movements ####
                '##################################################################
                'elapsed = Timer
                RenderMotion
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderMotion: " & elapsed
                                
            

                '#########################################################
                '#### RenderSpaces the skies/planes that may be setup ####
                '#########################################################
                'elapsed = Timer
                RenderSpaces
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
                
                
                '########################################################
                '#### RenderWorld renders all the mesh based objects ####
                '########################################################
                'elapsed = Timer
                RenderWorld
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
                

                '##########################################################
                '#### RenderPlayer renders the player's element object ####
                '##########################################################
                'elapsed = Timer
                RenderPlayer
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
                

                '##################################################################
                '#### RenderBoards renders any visible texture boards or walls ####
                '##################################################################
                'elapsed = Timer
                RenderBoards
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed


                '##################################################################
                '#### RenderLucent renders alphablent and translucent textures ####
                '##################################################################
                'elapsed = Timer
                RenderLucent
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed


                '#############################################################
                '#### RenderBeacons renders forward faced texture beacons ####
                '#############################################################
                'elapsed = Timer
                RenderBeacons
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed


                '###############################################################
                '#### RenderCameras moves the view camera if in camera mode ####
                '###############################################################
                'elapsed = Timer
                RenderCameras
                'elapsed = (Timer - elapsed)
                'If elapsed > 0 Then Debug.Print "RenderCameras: " & elapsed
                
                


'                '#################################################################
'                '#### SetupWorld configures the matricies used for DirectX 3D ####
'                '#################################################################
'                'elapsed = Timer
'                SetupWorld
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
'
'
''                '##################################################################
''                '#### RenderMotion prepairs a subsets of preliminary movements ####
''                '##################################################################
''                'elapsed = Timer
''                RenderMotion
''                'elapsed = (Timer - elapsed)
''                'If elapsed > 0 Then Debug.Print "RenderMotion: " & elapsed
'
'
'
'                '#########################################################
'                '#### RenderSpaces the skies/planes that may be setup ####
'                '#########################################################
'                'elapsed = Timer
'                RenderSpaces
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
'
'
'                '########################################################
'                '#### RenderWorld renders all the mesh based objects ####
'                '########################################################
'                'elapsed = Timer
'                RenderWorld
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
'
'
'                '##########################################################
'                '#### RenderPlayer renders the player's element object ####
'                '##########################################################
'                'elapsed = Timer
'                RenderPlayer
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
'
'
'                '##################################################################
'                '#### RenderBoards renders any visible texture boards or walls ####
'                '##################################################################
'                'elapsed = Timer
'                RenderBoards
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
'
'
''                '################################################################
''                '#### RenderMirrors renders mirrors collected by BeginMirros ####
''                '################################################################
''                 'elapsed = Timer
''                RenderMirrors
''                'elapsed = (Timer - elapsed)
''                'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
'
'
'                '##################################################################
'                '#### RenderLucent renders alphablent and translucent textures ####
'                '##################################################################
'                'elapsed = Timer
'                RenderLucent
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
'
'
'                '#############################################################
'                '#### RenderBeacons renders forward faced texture beacons ####
'                '#############################################################
'                'elapsed = Timer
'                RenderBeacons
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
'
'
''                '###########################################################
''                '#### RenderPortals handles all the Portal based events ####
''                '###########################################################
''                'elapsed = Timer
''                RenderPortals
''                'elapsed = (Timer - elapsed)
''                'If elapsed > 0 Then Debug.Print "RenderPortals: " & elapsed
'
'
'                '###############################################################
'                '#### RenderCameras moves the view camera if in camera mode ####
'                '###############################################################
'                'elapsed = Timer
'                RenderCameras
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderCameras: " & elapsed




''                    'elapsed = Timer
''                    SetupWorld
''                    'elapsed = (Timer - elapsed)
''                    'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
'
'
'                    'elapsed = Timer
'                    SetupMirror e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
'
'
'                    '#########################################################
'                    '#### RenderSpaces the skies/planes that may be setup ####
'                    '#########################################################
'                    'elapsed = Timer
'                    RenderSpaces e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
'
'
'                    '########################################################
'                    '#### RenderWorld renders all the mesh based objects ####
'                    '########################################################
'                    'elapsed = Timer
'                    RenderWorld e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
'
'
'                    '##########################################################
'                    '#### RenderPlayer renders the player's element object ####
'                    '##########################################################
'                    'elapsed = Timer
'                    RenderPlayer e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
'
'
'                    '##################################################################
'                    '#### RenderBoards renders any visible texture boards or walls ####
'                    '##################################################################
'                    'elapsed = Timer
'                    RenderBoards e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
'
'
'                    '##################################################################
'                    '#### RenderLucent renders alphablent and translucent textures ####
'                    '##################################################################
'                    'elapsed = Timer
'                    RenderLucent e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
'
'
'                    '#############################################################
'                    '#### RenderBeacons renders forward faced texture beacons ####
'                    '#############################################################
'                    'elapsed = Timer
'                    RenderBeacons e
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed

                                
            

'                '#########################################################
'                '#### RenderSpaces the skies/planes that may be setup ####
'                '#########################################################
'                'elapsed = Timer
'                RenderSpaces
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
'
'
'                '########################################################
'                '#### RenderWorld renders all the mesh based objects ####
'                '########################################################
'                'elapsed = Timer
'                RenderWorld
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
'
'
'                '##########################################################
'                '#### RenderPlayer renders the player's element object ####
'                '##########################################################
'                'elapsed = Timer
'                RenderPlayer
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
'
'
'
'
'
'
'
'
'
'                '##################################################################
'                '#### RenderBoards renders any visible texture boards or walls ####
'                '##################################################################
'                'elapsed = Timer
'                RenderBoards
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
'
'
'                '##################################################################
'                '#### RenderLucent renders alphablent and translucent textures ####
'                '##################################################################
'                'elapsed = Timer
'                RenderLucent
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
'
'
'                '#############################################################
'                '#### RenderBeacons renders forward faced texture beacons ####
'                '#############################################################
'                'elapsed = Timer
'                RenderBeacons
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
'
'
'                '###########################################################
'                '#### RenderPortals handles all the Portal based events ####
'                '###########################################################
'                'elapsed = Timer
'                RenderPortals
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderPortals: " & elapsed
'
'
'                '###############################################################
'                '#### RenderCameras moves the view camera if in camera mode ####
'                '###############################################################
'                'elapsed = Timer
'                RenderCameras
'                'elapsed = (Timer - elapsed)
'                'If elapsed > 0 Then Debug.Print "RenderCameras: " & elapsed
                
          
          
                    
                    DSurface.EndScene
                    

                    
                    DDevice.GetDisplayMode dm

                    rct.Top = 0
                    rct.Left = 0

                    rct.Right = DViewPort.Width
                    rct.Bottom = DViewPort.Height

                    D3DX.SaveSurfaceToFile GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp", D3DXIFF_BMP, ReflectRenderTarget, pal, rct
                     
                    Mirrors.Add D3DX.CreateTextureFromFileEx(DDevice, GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp", _
                        DViewPort.Width, DViewPort.Height, D3DX_FILTER_NONE, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
                        D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0), Boards.Key(i)
                    Kill GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp"
                
                End If

            End If
            Set e = Nothing
        Next
    End If
   ' DDevice.SetRenderState D3DRS_CULLMODE, CullMode
    
End Sub


Public Sub RenderMirrors()

    DDevice.SetRenderState D3DRS_ZENABLE, 1

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

'    Dim matWorld As D3DMATRIX
'    D3DXMatrixIdentity matWorld
'    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    If Player.Camera Is Nothing Then Exit Sub
    
    Dim e As Board
    Dim i As Long
    Dim L As Single
    If Not Mirrors Is Nothing Then
    
        If Boards.Count > 0 Then
            For i = 1 To Boards.Count
                Set e = Boards(i)
    
                If e.Visible And e.Mirror Then
                
                    L = Distance(Camera.Element.Origin.X, Camera.Element.Origin.Y, Camera.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                    If L <= FAR Then
    
                        If Mirrors.Exists(Boards.Key(i)) Then
    
    
                            DDevice.SetMaterial GenericMaterial
                            DDevice.SetTexture 0, Mirrors.Item(Boards.Key(i))
                            DDevice.SetTexture 1, Nothing
    
                            e.Render
    
                        End If
    
                    End If
    
                End If
                Set e = Nothing
            Next
        End If
    End If
End Sub


Public Sub SetupMirror(ByRef Mirror As Board)
On Error GoTo WorldError

    If Player.Camera Is Nothing Then Exit Sub

'    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX

    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matRoll As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matLook As D3DMATRIX

    Dim matWorld As D3DMATRIX
    Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matLook

    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity matTemp
    D3DXMatrixIdentity matRotation
    D3DXMatrixIdentity matPitch



    D3DXMatrixRotationY matRotation, 0
    D3DXMatrixRotationX matPitch, 0.5
    D3DXMatrixMultiply matWorld, matRotation, matPitch
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    
    
'    'Mirror plane: at z = 0, facing +Z
'    Dim MirrorPlane As Plane
'    Set MirrorPlane = ToPlane(Mirror.Point1, Mirror.Point2, Mirror.Point3)
'
'
'
'    Dim rPos As Point
'    Dim rLook As Point
'
'    ' Reflect camera position & look direction
'    Set rPos = ReflectPoint(Camera.Origin, MirrorPlane)
'    Set rLook = ReflectPoint(Camera.Rotate, MirrorPlane)
'
'
'
'    Dim MirrorView As D3DMATRIX
'
'
'
'
'    D3DXMatrixLookAtLH MirrorView, ToVector(rPos), ToVector(rLook), MakeVector(0, 1, 0)
'    DDevice.SetTransform D3DTS_VIEW, MirrorView
        


    

    'Set vec = Nothing
    
    Exit Sub
WorldError:
    If Err.Number = 6 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Resume
End Sub

'Private Function NormalizeVector(V As D3DVECTOR) As D3DVECTOR
'    Dim L As Single
'    L = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
'    If L <> 0! Then
'        NormalizeVector.X = V.X / L
'        NormalizeVector.Y = V.Y / L
'        NormalizeVector.Z = V.Z / L
'    Else
'        NormalizeVector = V
'    End If
'End Function
'
'Private Function ReflectPointAcrossPlane(p As D3DVECTOR, _
'                                         PlanePoint As D3DVECTOR, _
'                                         PlaneNormal As D3DVECTOR) As D3DVECTOR
'    Dim V As D3DVECTOR
'    Dim dot As Single
'
'    V.X = p.X - PlanePoint.X
'    V.Y = p.Y - PlanePoint.Y
'    V.Z = p.Z - PlanePoint.Z
'
'    dot = V.X * PlaneNormal.X + V.Y * PlaneNormal.Y + V.Z * PlaneNormal.Z
'
'    ReflectPointAcrossPlane.X = p.X - 2! * dot * PlaneNormal.X
'    ReflectPointAcrossPlane.Y = p.Y - 2! * dot * PlaneNormal.Y
'    ReflectPointAcrossPlane.Z = p.Z - 2! * dot * PlaneNormal.Z
'End Function
'
'Private Function ReflectVector(V As D3DVECTOR, N As D3DVECTOR) As D3DVECTOR
'    Dim dot As Single
'    dot = V.X * N.X + V.Y * N.Y + V.Z * N.Z
'
'    ReflectVector.X = V.X - 2! * dot * N.X
'    ReflectVector.Y = V.Y - 2! * dot * N.Y
'    ReflectVector.Z = V.Z - 2! * dot * N.Z
'End Function


'Private Function VecDot(a As Vec3, b As Vec3) As Single
'    VecDot = a.X * b.X + a.Y * b.Y + a.Z * b.Z
'End Function
'
'Private Function VecSub(a As Vec3, b As Vec3) As Vec3
'    VecSub.X = a.X - b.X
'    VecSub.Y = a.Y - b.Y
'    VecSub.Z = a.Z - b.Z
'End Function
'
'Private Function VecAdd(a As Vec3, b As Vec3) As Vec3
'    VecAdd.X = a.X + b.X
'    VecAdd.Y = a.Y + b.Y
'    VecAdd.Z = a.Z + b.Z
'End Function
'
'Private Function VecScale(a As Vec3, s As Single) As Vec3
'    VecScale.X = a.X * s
'    VecScale.Y = a.Y * s
'    VecScale.Z = a.Z * s
'End Function

' Reflect a point across a plane
Private Function ReflectPoint(p As Point, pl As Plane) As Point
    Dim Dist As Single
    Dist = modGeometry.VectorDotProduct(pl, p) + pl.W
    Set ReflectPoint = VectorDeduction(p, VectorMultiplyBy(pl, 2 * Dist))
End Function


