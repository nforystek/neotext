Attribute VB_Name = "modMirror"
Option Explicit

Private Mirrors As NTNodes10.Collection

Public Sub BeginMirrors()

    Dim e As Board
    Dim i As Long
    Dim L As Single

    Dim dm As D3DDISPLAYMODE
    Dim pal As PALETTEENTRY
    Dim rct As RECT

    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX

    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matWorld As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matTemp As D3DMATRIX
    
    If Not Mirrors Is Nothing Then Mirrors.Clear
    If Boards.Count > 0 Then
        For i = 1 To Boards.Count
            Set e = Boards(i)

            If e.Visible And e.Mirror Then


                L = Distance(Player.Element.Origin.X, Player.Element.Origin.Y, Player.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                If L <= FAR Then

                    If Mirrors Is Nothing Then Set Mirrors = New NTNodes10.Collection

                    DViewPort.Width = 128
                    DViewPort.Height = 128


                    DSurface.BeginScene ReflectRenderTarget, DViewPort




                    
'                    D3DXMatrixIdentity matWorld
'                    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'                    D3DXMatrixMultiply matTemp, matWorld, matWorld
'                    D3DXMatrixRotationY matRotation, AngleInvertRotation(0.5)
'                    D3DXMatrixRotationX matPitch, 0.5
'                    D3DXMatrixIdentity matWorld
'                    D3DXMatrixMultiply matLook, matRotation, matPitch
'                    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'                    If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex > 0 And Player.CameraIndex <= Cameras.Count)) Or (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0)) Then
'
'                        D3DXMatrixRotationY matRotation, AngleInvertRotation(Cameras(Player.CameraIndex).Angle)
'                        D3DXMatrixRotationX matPitch, Cameras(Player.CameraIndex).Pitch
'                        D3DXMatrixMultiply matLook, matRotation, matPitch
'
'                        D3DXMatrixTranslation matPos, -e.Origin.X, -e.Origin.Y, -e.Origin.Z
'                        D3DXMatrixMultiply matLook, matPos, matLook
'                       ' D3DXMatrixTranslation matPos, e.Origin.X, e.Origin.Y + 0.2, e.Origin.Z
'
'                    Else
'
'                        D3DXMatrixRotationY matRotation, AngleInvertRotation(Player.Camera.Angle)
'                        D3DXMatrixRotationX matPitch, Player.Camera.Pitch
'                        D3DXMatrixMultiply matLook, matRotation, matPitch
'
'                        If Player.Camera.Pitch > 0 Then
'
'                            D3DXMatrixTranslation matPos, -e.Origin.X, -e.Origin.Y, -e.Origin.Z
'                            D3DXMatrixMultiply matLook, matPos, matLook
'                        Else
'                            D3DXMatrixTranslation matPos, -e.Origin.X, -e.Origin.Y + 0.2, -e.Origin.Z
'                            D3DXMatrixMultiply matLook, matPos, matLook
'
'                        End If
'
'                    End If
'
'
'
'                    DDevice.SetTransform D3DTS_VIEW, matLook
'
'
'                    D3DXMatrixPerspectiveFovLH matProj, FOVY, AspectRatio, 0.01, FadeDistance
'                    DDevice.SetTransform D3DTS_PROJECTION, matProj
                

'
''                    'elapsed = Timer
''                    SetupWorld
''                    'elapsed = (Timer - elapsed)
''                    'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
'
'                    'elapsed = Timer
'                    RenderSpaces
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
'
'                    'elapsed = Timer
'                    RenderWorld
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
'
'                    'elapsed = Timer
'                    RenderPlayer
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
'
'                    'elapsed = Timer
'                    RenderBoards
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
'
'                    'elapsed = Timer
'                    RenderLucent
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
'
'                    'elapsed = Timer
'                    RenderBeacons
'                    'elapsed = (Timer - elapsed)
'                    'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
                    

                   ' DSurface.EndScene
                    
'                    RenderSpacesClose

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
End Sub


Public Sub RenderMirrors()

    Dim e As Board
    Dim i As Long
    Dim L As Single

    If Boards.Count > 0 Then
        For i = 1 To Boards.Count
            Set e = Boards(i)

            If e.Visible And e.Mirror Then
            
                L = Distance(Player.Element.Origin.X, Player.Element.Origin.Y, Player.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                If L <= FAR Then

                    If Mirrors.Exists(Boards.Key(i)) Then

                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        DDevice.SetMaterial GenericMaterial
                        DDevice.SetTexture 0, Mirrors.Item(Boards.Key(i))
                        DDevice.SetTexture 1, Nothing

                        'DDevice.SetStreamSource 0, Faces(e.FaceIndex).VBuffer, Len(Faces(e.FaceIndex).Verticies(0))
                        'DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
                        e.Render

                    End If

                End If

            End If
            Set e = Nothing
        Next
    End If
End Sub


Public Sub SetupMirror(ByRef Mirror As Board)
On Error GoTo WorldError

    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX

    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matWorld As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matTemp As D3DMATRIX
    
    If PointSideOfPlane(Mirror.Point1, Mirror.Point2, Mirror.Point3, Player.Element.Origin) Then Exit Sub
    
    
    D3DXMatrixIdentity matWorld

    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    D3DXMatrixMultiply matTemp, matWorld, matWorld
    D3DXMatrixRotationY matRotation, 0.5
    D3DXMatrixRotationX matPitch, 0.5
    D3DXMatrixIdentity matWorld
    D3DXMatrixMultiply matLook, matRotation, matPitch
    DDevice.SetTransform D3DTS_WORLD, matWorld

    If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex > 0 And Player.CameraIndex <= Cameras.Count)) Or (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0)) Then
        
        D3DXMatrixRotationY matRotation, Cameras(Player.CameraIndex).Angle
        D3DXMatrixRotationX matPitch, Cameras(Player.CameraIndex).Pitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        D3DXMatrixTranslation matPos, -Cameras(Player.CameraIndex).Origin.X, -Cameras(Player.CameraIndex).Origin.Y, -Cameras(Player.CameraIndex).Origin.Z
        D3DXMatrixMultiply matLook, matPos, matLook
        D3DXMatrixTranslation matPos, -Player.Element.Origin.X, -Player.Element.Origin.Y + 0.2, -Player.Element.Origin.Z
        
    Else
    
        D3DXMatrixRotationY matRotation, Player.Camera.Angle
        D3DXMatrixRotationX matPitch, Player.Camera.Pitch
        D3DXMatrixMultiply matLook, matRotation, matPitch

        If Player.Camera.Pitch > 0 Then
            D3DXMatrixTranslation matPos, -Player.Element.Origin.X, -Player.Element.Origin.Y, -Player.Element.Origin.Z
            D3DXMatrixMultiply matLook, matPos, matLook
        Else
            D3DXMatrixTranslation matPos, -Player.Element.Origin.X, -Player.Element.Origin.Y + 0.2, -Player.Element.Origin.Z
            D3DXMatrixMultiply matLook, matPos, matLook
        
        End If
        
    End If
    
    lCulledFaces = 0
    lCullCalls = 0
    
    If ((Perspective = Playmode.ThirdPerson) Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not (((Perspective = Spectator) Or DebugMode) And (Player.CameraIndex > 0))) Then
    
        If (CameraClip Or ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0))) And (Not ((Perspective = Spectator) Or DebugMode)) Then

            If ((Perspective = Playmode.CameraMode) And (Player.CameraIndex = 0)) Then

                Player.Element.Twists.Y = 3
            
            End If
        
            Dim cnt As Long
            Dim cnt2 As Long

            Dim Face As Long
            Dim Zoom As Single
            Dim factor As Single
            Dim e1 As Element

            Dim verts(0 To 2) As D3DVECTOR
            Dim touched As Boolean
            Dim V As Point
            
            'initialie sngFaceVis for camera collision checking
            For cnt = 1 To lngFaceCount - 1
                           ' On Error GoTo isdivcheck0
                                sngFaceVis(3, cnt) = 0
                             '   GoTo notdivcheck0
'isdivcheck0:
                                'If Err.Number = 11 Then Resume
'notdivcheck0:
                               ' If Err Then Err.Clear
                              '  On Error GoTo 0
                
            Next

            'commence the camera clip collision checking, this is what keeps
            'the camera from being inside of the level seeing out backfaces
            
            Zoom = 0.2
            factor = 0.5

            Do

                verts(0) = MakeVector(Player.Element.Origin.X, _
                                            Player.Element.Origin.Y - 0.2, _
                                            Player.Element.Origin.Z)

                verts(1) = MakeVector(Player.Element.Origin.X - (Sin(D720 - Player.Camera.Angle) * (Zoom + factor)), _
                                            Player.Element.Origin.Y - 0.2 + (Tan(D720 - Player.Camera.Pitch) * (Zoom + factor)), _
                                            Player.Element.Origin.Z - (Cos(D720 - Player.Camera.Angle) * (Zoom + factor)))

                verts(2) = MakeVector(Player.Element.Origin.X - (Sin(D720 - Player.Camera.Angle)), _
                                      Player.Element.Origin.Y - 0.1 + (Tan(D720 - Player.Camera.Pitch) * Zoom), _
                                      Player.Element.Origin.Z - (Cos(D720 - Player.Camera.Angle)))

'                Set v = VectorNegative(VectorRotateY(VectorRotateX(MakePoint(0, 0, -1), Player.Camera.Pitch), Player.Camera.Angle))
'
'                sngCamera(1, 0) = v.X
'                sngCamera(1, 1) = v.Y
'                sngCamera(1, 2) = v.Z
'
'                Set v = VectorNegative(VectorRotateY(VectorRotateX(MakePoint(0, 1, 0), Player.Camera.Pitch), Player.Camera.Angle))
'
'                sngCamera(2, 0) = v.X
'                sngCamera(2, 1) = v.Y
'                sngCamera(2, 2) = v.Z
'
'                sngCamera(0, 0) = Player.Element.Origin.X
'                sngCamera(0, 1) = Player.Element.Origin.Y
'                sngCamera(0, 2) = Player.Element.Origin.Z

                sngCamera(0, 0) = Player.Element.Origin.X
                sngCamera(0, 1) = Player.Element.Origin.Y
                sngCamera(0, 2) = Player.Element.Origin.Z

                sngCamera(1, 0) = 1
                sngCamera(1, 1) = -1
                sngCamera(1, 2) = -1

                sngCamera(2, 0) = -1
                sngCamera(2, 1) = 1
                sngCamera(2, 2) = -1
                
                If lngFaceCount > 0 Then
                    lCulledFaces = lCulledFaces + Culling(2, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
                    lCullCalls = lCullCalls + 1
                End If

                If (Elements.Count > 0) Then
                    For cnt = 1 To Elements.Count
                        Set e1 = Elements(cnt)
                    'For Each e1 In Elements
                    
                    'For cnt = 1 To Elements.Count
                        If ((Not (e1.Effect = Collides.Ground)) And (Not (e1.Effect = Collides.InDoor))) And (e1.CollideIndex > -1) And (e1.BoundsIndex > 0) Then
                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                      '      On Error GoTo isdivcheck1
                                sngFaceVis(3, cnt2) = 0
                           '     GoTo notdivcheck1
'isdivcheck1:
                             '   If Err.Number = 11 Then Resume
'notdivcheck1:
                            '    If Err Then Err.Clear
                           '     On Error GoTo 0
                                
                            Next
                        ElseIf (e1.Effect = Collides.Ground) And (e1.CollideIndex > -1) And e1.Visible And (e1.BoundsIndex > 0) Then
                            For cnt2 = e1.CollideIndex To (e1.CollideIndex + Meshes(e1.BoundsIndex).Mesh.GetNumFaces) - 1
                                If Not (((sngFaceVis(0, cnt2) = 0) Or (sngFaceVis(0, cnt2) = 1) Or (sngFaceVis(0, cnt2) = -1)) And _
                                    ((sngFaceVis(1, cnt2) = 0) Or (sngFaceVis(1, cnt2) = 1) Or (sngFaceVis(1, cnt2) = -1)) And _
                                    ((sngFaceVis(2, cnt2) = 0) Or (sngFaceVis(2, cnt2) = 1) Or (sngFaceVis(2, cnt2) = -1))) Then
                                    sngFaceVis(3, cnt2) = 2
                                End If
                            Next
                        End If
                        
                        Set e1 = Nothing
                    Next
                    If (Player.Element.CollideIndex > -1) And (Player.Element.BoundsIndex > 0) And (Player.Element.BoundsIndex > 0) Then
                        For cnt2 = Player.Element.CollideIndex To (Player.Element.CollideIndex + Meshes(Player.Element.BoundsIndex).Mesh.GetNumFaces) - 1
                          '  On Error GoTo isdivcheck2
                                sngFaceVis(3, cnt2) = 0
                                'GoTo notdivcheck2
'isdivcheck2:
   '                             If Err.Number = 11 Then Resume
'notdivcheck2:
                          '      If Err Then Err.Clear
                          '      On Error GoTo 0

                        Next
                    End If
                End If

                Face = AddCollisionEx(verts, 1)
                touched = TestCollisionEx(Face, 2)
                DelCollisionEx Face, 1

                If ((Not touched) And (Zoom < Player.Camera.Zoom)) Then Zoom = Zoom + factor

            Loop Until ((touched) Or (Zoom >= Player.Camera.Zoom))

            If (touched And (Zoom > 0.2)) Then Zoom = Zoom + -factor

            D3DXMatrixTranslation matTemp, 0, 0.2, Zoom
            D3DXMatrixMultiply matView, matLook, matTemp

            'all said and done, if the zoom is under a certian val the
            'toon is in the way, so change it to wireframe see through
            Player.Element.WireFrame = (Zoom < 0.8)

        Else
        
            D3DXMatrixTranslation matTemp, 0, 0, IIf(Not ((Perspective = Spectator) Or DebugMode), Player.Camera.Zoom, 0)
            D3DXMatrixMultiply matView, matLook, matTemp
        End If
        DDevice.SetTransform D3DTS_VIEW, matView
        
        D3DXMatrixMultiply matLook, matView, matTemp
    Else
        DDevice.SetTransform D3DTS_VIEW, matLook
    End If

    
    D3DXMatrixPerspectiveFovLH matProj, FOVY, AspectRatio, 0.01, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
    


 
    D3DXMatrixInverse matWorld, 1, matWorld
    
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
 
    D3DXMatrixInverse matLook, 1, matLook
    
    DDevice.SetTransform D3DTS_VIEW, matLook
    
 
    D3DXMatrixInverse matProj, 1, matProj
    
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    
       
    
    
    
    
    Exit Sub
WorldError:
    If Err.Number = 6 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Resume
End Sub



