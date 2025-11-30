Attribute VB_Name = "modMirror"
Option Explicit

Public Mirroring As Boolean

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

            If e.Visible And e.Mirror And PointSideOfPlane(e.Point1, e.Point2, e.Point3, Player.Element.Origin) Then


                L = Distance(Player.Element.Origin.X, Player.Element.Origin.Y, Player.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
                If L <= FAR Then

                    If Mirrors Is Nothing Then Set Mirrors = New NTNodes10.Collection

                    DViewPort.width = 128
                    DViewPort.height = 128

                    DSurface.BeginScene DefaultRenderTarget, DViewPort

                    'elapsed = Timer
                    SetupMirror e
                    'elapsed = (Timer - elapsed)
                    'If elapsed > 0 Then Debug.Print "SetupWorld: " & elapsed
                    
                    DSurface.EndScene
                    
                    
                    DDevice.GetDisplayMode dm

                    rct.Top = 0
                    rct.Left = 0

                    rct.Right = DViewPort.width
                    rct.Bottom = DViewPort.height

                    D3DX.SaveSurfaceToFile GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp", D3DXIFF_BMP, DefaultRenderTarget, pal, rct
                     
                    Mirrors.Add D3DX.CreateTextureFromFileEx(DDevice, GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp", _
                        DViewPort.width, DViewPort.height, D3DX_FILTER_NONE, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
                        D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0), Boards.Key(i)
                    Kill GetTemporaryFolder & "\" & Boards.Key(i) & ".bmp"
                    
                End If

            End If
            Set e = Nothing
        Next
    End If
End Sub


Public Sub RenderMirrors()

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault


    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Dim e As Board
    Dim i As Long
    Dim L As Single
    If Not Mirrors Is Nothing Then
    
        If Boards.Count > 0 Then
            For i = 1 To Boards.Count
                Set e = Boards(i)
    
                If e.Visible And e.Mirror Then
                
                    L = Distance(Player.Element.Origin.X, Player.Element.Origin.Y, Player.Element.Origin.Z, e.Origin.X, e.Origin.Y, e.Origin.Z)
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

    Dim matView As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matProj As D3DMATRIX

    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matWorld As D3DMATRIX

    Dim matPos As D3DMATRIX
    Dim matTemp As D3DMATRIX
    
    
'    D3DXMatrixIdentity matWorld
'    DDevice.SetTransform D3DTS_WORLD, matWorld
    
    
'    D3DXMatrixMultiply matTemp, matWorld, matWorld
'    D3DXMatrixRotationY matRotation, 0.5
'    D3DXMatrixRotationX matPitch, 0.5
'    D3DXMatrixIdentity matWorld
'    D3DXMatrixMultiply matLook, matRotation, matPitch
'    DDevice.SetTransform D3DTS_WORLD, matLook

    Dim vec As Point
    Set vec = VectorDeduction(Mirror.Origin, Player.Element.Origin)
    
    Dim norm As Point
    Set norm = PlaneNormal(Mirror.Point1, Mirror.Point2, Mirror.Point3)
    
    Dim angle As Single
    Dim pitch As Single
     
    angle = AngleOfPlot(vec.X, vec.Z)
    
    pitch = (Player.Element.Origin.Y - Mirror.Origin.Y)

    D3DXMatrixRotationY matRotation, angle
    D3DXMatrixRotationX matPitch, pitch
    D3DXMatrixMultiply matLook, matRotation, matPitch
    
    D3DXMatrixTranslation matPos, -Mirror.Origin.Z, -Mirror.Origin.X, -Mirror.Origin.Y
    D3DXMatrixMultiply matLook, matPos, matLook
    

    DDevice.SetTransform D3DTS_WORLD, matLook
    
    


'    D3DXMatrixInverse matLook, 1, matLook
    
    Set vec = Nothing

   ' DDevice.SetTransform D3DTS_VIEW, matWorld
        
'    D3DXMatrixPerspectiveFovLH matProj, 0.01, AspectRatio, 0.01, FadeDistance
'    DDevice.SetTransform D3DTS_PROJECTION, matProj

    
    Exit Sub
WorldError:
    If Err.Number = 6 Then Resume
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Resume
End Sub



