Attribute VB_Name = "modMirror"
Option Explicit

Public Mirroring As Boolean

Private Mirrors As NTNodes10.Collection

Public Sub BeginMirrors()

    If Camera.Element Is Nothing Then Exit Sub

    Dim e As Board
    Dim i As Long
    Dim L As Single

    Dim dm As D3DDISPLAYMODE
    Dim pal As PALETTEENTRY
    Dim rct As RECT
    Dim CullMode As Long

        
    If Not Mirrors Is Nothing Then Mirrors.Clear

    If Boards.Count > 0 Then
        For i = 1 To Boards.Count
            Set e = Boards(i)

            If e.Visible And e.Mirror And PointSideOfPlane(e.Point1, e.Point2, e.Point3, Camera.Element.Origin) Then


                L = Distance(Camera.Element.Origin.X, Camera.Element.Origin.Y, Camera.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                If L <= FAR Then

                    If Mirrors Is Nothing Then Set Mirrors = New NTNodes10.Collection
                    
                    
                    CullMode = DDevice.GetRenderState(D3DRS_CULLMODE)
                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
                    
                    
                    DViewPort.X = 0
                    DViewPort.Y = 0
                    DViewPort.Width = 128
                    DViewPort.Height = 128
                    
                    DSurface.BeginScene ReflectRenderTarget, DViewPort

                    'elapsed = Timer
                    SetupWorld e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
                    

                    '#########################################################
                    '#### RenderSpaces the skies/planes that may be setup ####
                    '#########################################################
                    'elapsed = Timer
                    RenderSpaces e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderSpaces: " & elapsed
    
    
                    '########################################################
                    '#### RenderWorld renders all the mesh based objects ####
                    '########################################################
                    'elapsed = Timer
                    RenderWorld e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderWorld: " & elapsed
    
    
                    '##########################################################
                    '#### RenderPlayer renders the player's element object ####
                    '##########################################################
                    'elapsed = Timer
                    RenderPlayer e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderPlayer: " & elapsed
    
    
                    '##################################################################
                    '#### RenderBoards renders any visible texture boards or walls ####
                    '##################################################################
                    'elapsed = Timer
                    RenderBoards e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderBoards: " & elapsed
    
    
                    '##################################################################
                    '#### RenderLucent renders alphablent and translucent textures ####
                    '##################################################################
                    'elapsed = Timer
                    RenderLucent e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "RenderLucent: " & elapsed
    
    
                    '#############################################################
                    '#### RenderBeacons renders forward faced texture beacons ####
                    '#############################################################
                    'elapsed = Timer
                    RenderBeacons e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "ReanderBeacons: " & elapsed
                    
                    
                    
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

                    DDevice.SetRenderState D3DRS_CULLMODE, CullMode
                
                End If

            End If
            Set e = Nothing
        Next
    End If
End Sub


Public Sub RenderMirrors()

    DDevice.SetRenderState D3DRS_ZENABLE, 1

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

    Dim matWorld As D3DMATRIX
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    If Camera.Element Is Nothing Then Exit Sub
    
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

    If Camera.Element Is Nothing Then Exit Sub

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

    Dim vec As Point
    Set vec = VectorDeduction(Camera.Element.Origin, Mirror.Origin)

    Dim norm As Point
    Set norm = PlaneNormal(Mirror.Point1, Mirror.Point2, Mirror.Point3)

    Dim Angle As Single
    Dim Pitch As Single

    Angle = AngleOfPlot(-norm.X, -norm.Z)

    Pitch = Mirror.Origin.Y - Camera.Element.Origin.Y

   ' Set norm = VertexNormalize(modGeometry.VectorCrossProduct(norm, vec))


    D3DXMatrixTranslation matPos, Mirror.Origin.X, Mirror.Origin.Y, Mirror.Origin.Z
    D3DXMatrixMultiply matLook, matPos, matLook
    DDevice.SetTransform D3DTS_VIEW, matLook

   ' Set norm = modGeometry.VectorMultiply(norm, MakePoint(Pitch, Angle, 0))


    D3DXMatrixRotationY matRotation, AngleInvertRotation(Angle)
    D3DXMatrixMultiply matLook, matRotation, matLook

'    D3DXMatrixRotationX matPitch, Pitch
'    D3DXMatrixMultiply matLook, matPitch, matLook


    DDevice.SetTransform D3DTS_VIEW, matLook


    
    
    D3DXMatrixPerspectiveFovLH matProj, FOVY * 2, AspectRatio, 0.01, FadeDistance
    DDevice.SetTransform D3DTS_PROJECTION, matProj
    

    Set vec = Nothing
    
    Exit Sub
WorldError:
    If Err.Number = 6 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Resume
End Sub



