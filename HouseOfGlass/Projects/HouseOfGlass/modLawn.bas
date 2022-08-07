Attribute VB_Name = "modLawn"
#Const modLawn = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Type MyMesh
    mesh As D3DXMesh
    mat() As D3DMATERIAL8
    Tex() As Direct3DTexture8
    XYZ() As D3DVERTEX
    idx() As Integer
    
    MCount As Long
    
    FileName As String
    
    Origin As D3DVECTOR
    Rotate As D3DVECTOR
    Scaled As D3DVECTOR
   
End Type

Public Type MyWall
    ObjectIndex As Long
    Rotated As D3DVECTOR
    Origin As D3DVECTOR
End Type

Public Type MyPlace
    Point1 As D3DVECTOR
    Point2 As D3DVECTOR
    Point3 As D3DVECTOR
    Point4 As D3DVECTOR
End Type

Public Type MyLevel
    Walls() As MyWall
    matWall() As D3DMATRIX
    
    Starts() As MyPlace
    Ends() As MyPlace
    
    LastScore As Single
    BestScore As Single
    
    Elapsed As Single
    Loaded As String
End Type


Public Lights() As D3DLIGHT8
Public nLight As Long

Public Objects() As MyMesh
Public nObject As Long

Private PlaneSkin As Direct3DTexture8
Private PlanePlaq() As TVERTEX2
Private PlaneVBuf As Direct3DVertexBuffer8

Public matTemp As D3DMATRIX
Public matObj() As D3DMATRIX
Public matLit() As D3DMATRIX

Private PressF1Skin As Direct3DTexture8
Private PressF1Plaq(0 To 4) As TVERTEX1

Public AtLvl As Long
Public Level As MyLevel

Public Sub RenderLawn()
    
    Dim fogVal As Long
    Dim t As Boolean
    Dim o As Long
    Dim i As Long
                     
    If (MenuMode = 2) Then
        If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(10, 10, 10) Then
            DDevice.SetRenderState D3DRS_AMBIENT, RGB(10, 10, 10)
        End If
    Else
        If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
            DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
        End If
    End If
    
    DDevice.SetVertexShader FVF_VTEXT2

    D3DXMatrixIdentity matTemp
    DDevice.SetTransform D3DTS_WORLD, matTemp

    DDevice.SetTexture 0, PlaneSkin
    DDevice.SetStreamSource 0, PlaneVBuf, Len(PlanePlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    
    D3DXMatrixIdentity matTemp
    DDevice.SetVertexShader FVF_VTEXT2
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    
    If Not (Level.Loaded = "") Then
        
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ZENABLE, 0
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
        
        
       
       
        For o = 1 To UBound(Level.Walls)
            DDevice.SetTransform D3DTS_WORLD, Level.matWall(o)
            For i = 0 To Objects(Level.Walls(o).ObjectIndex).MCount - 1
                If Not (Objects(Level.Walls(o).ObjectIndex).Tex(i) Is Nothing) Then
                    If Objects(Level.Walls(o).ObjectIndex).Tex(i).GetPriority = 1 Then
                        DDevice.SetMaterial Objects(Level.Walls(o).ObjectIndex).mat(i)
                        DDevice.SetTexture 0, Objects(Level.Walls(o).ObjectIndex).Tex(i)
                        Objects(Level.Walls(o).ObjectIndex).mesh.DrawSubset i
                    End If
                End If
            Next
        Next


        
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        
        For o = 1 To UBound(Level.Walls)
            DDevice.SetTransform D3DTS_WORLD, Level.matWall(o)
            For i = 0 To Objects(Level.Walls(o).ObjectIndex).MCount - 1
                If Not (Objects(Level.Walls(o).ObjectIndex).Tex(i) Is Nothing) Then
                    If Objects(Level.Walls(o).ObjectIndex).Tex(i).GetPriority = 0 Then
                        DDevice.SetMaterial Objects(Level.Walls(o).ObjectIndex).mat(i)
                        DDevice.SetTexture 0, Objects(Level.Walls(o).ObjectIndex).Tex(i)
                        Objects(Level.Walls(o).ObjectIndex).mesh.DrawSubset i
                    End If
                End If
            Next
        Next

         
    '    DDevice.SetRenderState D3DRS_CLIPPING, CONST_D3DCLIPFLAGS.D3DCS_PLANE0
    '    DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, CONST_D3DCLIPPLANEFLAGS.D3DCLIPPLANE0
    '
      '  DDevice.SetRenderTarget DefaultRenderTarget, DefaultStencilDepth, ByVal 0&

        

        
      '  DDevice.SetRenderState D3DRS_CLIPPING, CONST_D3DCLIPFLAGS.D3DCS_PLANE1
      '  DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, CONST_D3DCLIPPLANEFLAGS.D3DCLIPPLANE1
    
       'DDevice.SetRenderTarget ReflectRenderTarget, ReflectStencilDepth, ByVal 0&
        

        
        Dim vc As D3DVECTOR
        vc = SquareCenter(Level.Ends(1).Point1, Level.Ends(1).Point2, Level.Ends(1).Point3, Level.Ends(1).Point4)

        Dim matRotate As D3DMATRIX

        D3DXMatrixRotationY matRotate, 180 * (PI / 180)
        D3DXMatrixMultiply matRotate, matObj(5), matRotate
        
        D3DXMatrixTranslation matTemp, vc.X, vc.Y, vc.Z
        D3DXMatrixMultiply matTemp, matRotate, matTemp
                        
        DDevice.SetTransform D3DTS_WORLD, matTemp

        For i = 0 To Objects(5).MCount - 1
            If Not (Objects(5).Tex(i) Is Nothing) Then
                DDevice.SetMaterial GenericMat
                DDevice.SetTexture 0, Objects(5).Tex(i)
                Objects(5).mesh.DrawSubset i
            End If
        Next
        
    Else
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        
        DDevice.SetTransform D3DTS_WORLD, matObj(nObject)
        For i = 0 To Objects(5).MCount - 1
            If Not (Objects(5).Tex(i) Is Nothing) Then
                DDevice.SetMaterial GenericMat
                DDevice.SetTexture 0, Objects(5).Tex(i)
                Objects(5).mesh.DrawSubset i
            End If
        Next
    
    End If
    
    If (MenuMode = 0) Then
        CenterMessage ""
        If Level.Elapsed = 0 Then
            Dim inStart As Boolean
            
            For o = 1 To UBound(Level.Starts)
                If (Player.Location.X > LeastX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.X < GreatestX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.Z > LeastZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                    Player.Location.Z < GreatestZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4)) Then
                    inStart = True
                End If
            Next
            If (Not inStart) Then
                Level.Elapsed = Timer
            Else
                db.rsQuery rs, "SELECT * FROM Scores WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND LevelNum=" & AtLvl & ";"
                If Not db.rsEnd(rs) Then
                    CenterMessage "Best Score: " & rs("BestTime")
                End If
                
            End If
        End If
    ElseIf (MenuMode = -1) Then
        CenterMessage vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Press F1 to start."
    ElseIf (MenuMode = 1) Then
        CenterMessage ""
        AtLvl = AtLvl + 1
        If Not PathExists(AppPath & "Levels\level" & AtLvl & ".hog", True) Then AtLvl = 1
        db.rsQuery rs, "SELECT * FROM Scores WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND LevelNum=" & AtLvl & ";"

        FadeMessage "Level " & AtLvl & IIf(Not db.rsEnd(rs), " (Press F3 to skip)", "")
        LoadLevel AppPath & "Levels\level" & AtLvl & ".hog"
        
        MenuMode = 0
    ElseIf (MenuMode = 2) Then
        CenterMessage vbCrLf & vbCrLf & "Your Score: " & Level.LastScore & vbCrLf & "Best Score: " & Level.BestScore & vbCrLf & vbCrLf & "Press F1 to replay level." & vbCrLf & "Press F2 for next level."
    End If
    
    If (Not (Level.Loaded = "")) And (MenuMode = 0) Then
        For o = 1 To UBound(Level.Ends)
        
            If Player.Location.X > LeastX(Level.Ends(o).Point1, Level.Ends(o).Point2, Level.Ends(o).Point3, Level.Ends(o).Point4) And _
                Player.Location.X < GreatestX(Level.Ends(o).Point1, Level.Ends(o).Point2, Level.Ends(o).Point3, Level.Ends(o).Point4) And _
                Player.Location.Z > LeastZ(Level.Ends(o).Point1, Level.Ends(o).Point2, Level.Ends(o).Point3, Level.Ends(o).Point4) And _
                Player.Location.Z < GreatestZ(Level.Ends(o).Point1, Level.Ends(o).Point2, Level.Ends(o).Point3, Level.Ends(o).Point4) Then
                
                If Level.Elapsed > 0 Then
                
                    Level.LastScore = GetElapsed
                    If Level.LastScore < Level.BestScore Or Level.BestScore = 0 Then
                        Level.BestScore = Level.LastScore
                    End If
                        
                    db.rsQuery rs, "SELECT * FROM Scores WHERE LevelNum=" & Replace(Replace(Replace(LCase(Level.Loaded), "level", ""), ".hog", ""), "'", "''") & ";"
                    If Not db.rsEnd(rs) Then
                        If rs("BestTime") < Level.BestScore Then
                            Level.BestScore = rs("BestTime")
                        Else
                            db.dbQuery "UPDATE Scores SET BestTime=" & Level.BestScore & " WHERE LevelNum=" & Replace(Replace(Replace(LCase(Level.Loaded), "level", ""), ".hog", ""), "'", "''") & ";"
                        End If
                    Else
                        db.dbQuery "INSERT INTO Scores (Username, LevelNum, BestTime) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "', " & Replace(Replace(Replace(LCase(Level.Loaded), "level", ""), ".hog", ""), "'", "''") & ", " & Level.BestScore & ");"
                    End If
                    
                    Level.Elapsed = 0
                Else
                    db.rsQuery rs, "SELECT * FROM Scores WHERE LevelNum=" & Replace(Replace(Replace(LCase(Level.Loaded), "level", ""), ".hog", ""), "'", "''") & ";"
                    If Not db.rsEnd(rs) Then
                        If rs("BestTime") < Level.BestScore Then
                            Level.BestScore = rs("BestTime")
                        End If
                    End If
                End If
                MenuMode = 2
                
            End If
        Next
    End If
End Sub

Public Sub CreateLawn(Optional ByVal inText As String = "")
    Dim r As Single
    Dim o As Long
    Dim i As Long
    Dim vn As D3DVECTOR
    
    Dim inArg() As String
    Dim inItem As String
    Dim inName As String
    Dim inData As String
    
    If inText = "" Then
        inText = Replace(ReadFile(AppPath & "Base\Base.px"), vbTab, "")
    End If
    
    Do Until inText = ""
    
        inItem = RemoveNextArg(inText, vbCrLf)
        If (Not (inItem = "")) And (Not (Left(inItem, 1) = ";")) Then
            inData = RemoveQuotedArg(inText, "{", "}")
            Select Case inItem
                Case "parse"
                    Do Until inData = ""
                        inName = RemoveNextArg(inData, vbCrLf)
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "filename"
                                     CreateLawn Replace(ReadFile(AppPath & "Base\Lawn\" & inArg(1) & ".px"), vbTab, "")
                            End Select
                        End If
                    Loop
                    
                Case "light"
                    nLight = nLight + 1
                    ReDim Preserve Lights(1 To nLight) As D3DLIGHT8
                    ReDim Preserve matLit(1 To nLight) As D3DMATRIX

                    Do Until inData = ""
                        inName = RemoveNextArg(inData, vbCrLf)
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "position"
                                    Lights(nLight).Position = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "direction"
                                    Lights(nLight).Direction = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "diffuse"
                                    Lights(nLight).diffuse.a = CSng(inArg(1))
                                    Lights(nLight).diffuse.r = CSng(inArg(2))
                                    Lights(nLight).diffuse.g = CSng(inArg(3))
                                    Lights(nLight).diffuse.b = CSng(inArg(4))
                                Case "ambient"
                                    Lights(nLight).Ambient.a = CSng(inArg(1))
                                    Lights(nLight).Ambient.r = CSng(inArg(2))
                                    Lights(nLight).Ambient.g = CSng(inArg(3))
                                    Lights(nLight).Ambient.b = CSng(inArg(4))
                                Case "specular"
                                    Lights(nLight).specular.a = CSng(inArg(1))
                                    Lights(nLight).specular.r = CSng(inArg(2))
                                    Lights(nLight).specular.g = CSng(inArg(3))
                                    Lights(nLight).specular.b = CSng(inArg(4))
                                Case "attenuation"
                                    Lights(nLight).Attenuation0 = CSng(inArg(1))
                                    Lights(nLight).Attenuation1 = CSng(inArg(2))
                                    Lights(nLight).Attenuation2 = CSng(inArg(3))
                                Case "phi"
                                    Lights(nLight).Phi = CSng(inArg(1))
                                Case "theta"
                                    Lights(nLight).Theta = CSng(inArg(1))
                                Case "falloff"
                                    Lights(nLight).Falloff = CSng(inArg(1))
                                Case "range"
                                    Lights(nLight).Range = inArg(1)
                                Case "type"
                                    Lights(nLight).Type = inArg(1)
                                    Lights(nLight).Attenuation0 = 0
                                    Lights(nLight).Attenuation1 = 0.05
                                    Lights(nLight).Attenuation2 = 0
                            End Select
                        End If
                    Loop
                    
                    D3DXMatrixIdentity matTemp
                    D3DXMatrixIdentity matLit(nLight)
                    D3DXMatrixScaling matLit(nLight), 1, 1, 1
                    D3DXMatrixTranslation matTemp, Lights(nLight).Position.X, Lights(nLight).Position.Y, Lights(nLight).Position.Z
                    D3DXMatrixMultiply matLit(nLight), matLit(nLight), matTemp
                    
                    DDevice.SetTransform D3DTS_WORLD, matLit(nLight)
                    DDevice.SetLight nLight, Lights(nLight)
                    DDevice.LightEnable nLight, True
                   
                Case "object"
                    nObject = nObject + 1
                    ReDim Preserve Objects(1 To nObject) As MyMesh
                    ReDim Preserve matObj(1 To nObject) As D3DMATRIX
                    
                    Do Until inData = ""
                        inName = RemoveNextArg(inData, vbCrLf)
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "origin"
                                    Objects(nObject).Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "filename"
                                    Objects(nObject).FileName = inArg(1)
                                Case "scale"
                                    Objects(nObject).Scaled = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                                Case "rotate"
                                    Objects(nObject).Rotate = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
                            End Select
                        End If
                    Loop
                    
                    If PathExists(AppPath & "Base\" & Objects(nObject).FileName & ".x", True) Then
                        CreateMesh AppPath & "Base\" & Objects(nObject).FileName & ".x", Objects(nObject).mesh, Objects(nObject).Scaled, Objects(nObject).mat, Objects(nObject).Tex, Objects(nObject).XYZ, Objects(nObject).idx, Objects(nObject).MCount
                        
                        D3DXMatrixIdentity matTemp
                        D3DXMatrixIdentity matObj(nObject)
                        
                        D3DXMatrixScaling matTemp, Objects(nObject).Scaled.X, Objects(nObject).Scaled.Y, Objects(nObject).Scaled.Z
                        D3DXMatrixMultiply matObj(nObject), matObj(nObject), matTemp
                        
                        D3DXMatrixRotationX matTemp, Objects(nObject).Rotate.X * (PI / 180)
                        D3DXMatrixMultiply matObj(nObject), matObj(nObject), matTemp
                        D3DXMatrixRotationY matTemp, Objects(nObject).Rotate.Y * (PI / 180)
                        D3DXMatrixMultiply matObj(nObject), matObj(nObject), matTemp
                        D3DXMatrixRotationZ matTemp, Objects(nObject).Rotate.Z * (PI / 180)
                        D3DXMatrixMultiply matObj(nObject), matObj(nObject), matTemp

                        D3DXMatrixTranslation matTemp, Objects(nObject).Origin.X, Objects(nObject).Origin.Y, Objects(nObject).Origin.Z
                        D3DXMatrixMultiply matObj(nObject), matObj(nObject), matTemp
                    Else
                        ReDim Objects(nObject).Tex(0 To 0) As Direct3DTexture8
                        ReDim Objects(nObject).mat(0 To 0) As D3DMATERIAL8
                    End If
                
                Case "plane"

                    ReDim PlanePlaq(0 To 5) As TVERTEX2
                    Do Until inData = ""
                        inName = RemoveNextArg(inData, vbCrLf)
                        If (Not (inName = "")) Then
                            inArg() = Split(Trim(inName), " ")
                            Select Case inArg(0)
                                Case "filename"
                                  Set PlaneSkin = LoadTexture(AppPath & "Base\" & inArg(1))
                            End Select
                        End If
                    Loop

                    CreateSquare PlanePlaq, 0, MakeVector(FadeDistance, -1, -FadeDistance), _
                                                MakeVector(-FadeDistance, -1, -FadeDistance), _
                                                MakeVector(-FadeDistance, -1, FadeDistance), _
                                                MakeVector(FadeDistance, -1, FadeDistance), _
                                                (FadeDistance * 2) / 128, (FadeDistance * 2) / 128
                                                
                    Set PlaneVBuf = DDevice.CreateVertexBuffer(Len(PlanePlaq(0)) * 6, 0, FVF_VTEXT2, D3DPOOL_DEFAULT)
                    D3DVertexBuffer8SetData PlaneVBuf, 0, Len(PlanePlaq(0)) * 6, 0, PlanePlaq(0)
    
            End Select
        End If
    Loop
    
End Sub

Private Function CreateMesh(ByVal FileName As String, mesh As D3DXMesh, Scaled As D3DVECTOR, MeshMaterials() As D3DMATERIAL8, MeshTextures() As Direct3DTexture8, MeshVerticies() As D3DVERTEX, MeshIndicies() As Integer, nMaterials As Long)
    Dim TextureName As String
    Dim MtrlBuffer As D3DXBuffer

    Set mesh = D3DX.LoadMeshFromX(FileName, D3DXMESH_DONOTCLIP, DDevice, Nothing, MtrlBuffer, nMaterials)
    
    ReDim MeshMaterials(0 To nMaterials - 1) As D3DMATERIAL8
    ReDim MeshTextures(0 To nMaterials - 1) As Direct3DTexture8

    Dim d As ImgDimType
    Dim t As String
    
    Dim q As Integer
    For q = 0 To nMaterials - 1

        D3DX.BufferGetMaterial MtrlBuffer, q, MeshMaterials(q)
        MeshMaterials(q).Ambient = MeshMaterials(q).diffuse
   
        TextureName = D3DX.BufferGetTextureName(MtrlBuffer, q)
        If (TextureName <> "") Then
            If ImageDimensions(AppPath & "Base\" & TextureName, d, t) Then
                Set MeshTextures(q) = D3DX.CreateTextureFromFileEx(DDevice, AppPath & "Base\" & TextureName, d.width, d.height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Transparent, ByVal 0, ByVal 0)
                
                MeshTextures(q).SetPriority -CLng(CBool(Left(TextureName, 5) = "glass"))
            Else
                Debug.Print "IMAGE ERROR: ImageDimensions"
            End If
        End If
        
    Next

    Set MtrlBuffer = Nothing

    Dim vd As D3DVERTEXBUFFER_DESC
    mesh.GetVertexBuffer.GetDesc vd

    ReDim MeshVerticies(0 To ((vd.Size \ 32) - 1)) As D3DVERTEX
    D3DVertexBuffer8GetData mesh.GetVertexBuffer, 0, vd.Size, 0, MeshVerticies(0)

    Dim ID As D3DINDEXBUFFER_DESC
    mesh.GetIndexBuffer.GetDesc ID

    ReDim MeshIndicies(0 To ((ID.Size \ 2) - 1)) As Integer
    D3DIndexBuffer8GetData mesh.GetIndexBuffer, 0, ID.Size, 0, MeshIndicies(0)

End Function

Private Sub AddCollisionObject(ByVal Brush As Long, ByRef Origin As D3DVECTOR, Scaled As D3DVECTOR, ByRef Rotated As D3DVECTOR, MeshVerticies() As D3DVERTEX, MeshIndicies() As Integer)

    Dim matRotate As D3DMATRIX
    D3DXMatrixRotationYawPitchRoll matRotate, Rotated.Y, 0, 0
    
    Dim matScale As D3DMATRIX
    D3DXMatrixScaling matScale, Scaled.X, Scaled.Y, Scaled.Z

    Dim vn As D3DVECTOR
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim v3 As D3DVECTOR

    Dim cnt As Long
    For cnt = 0 To UBound(MeshIndicies) Step 3

        v1 = MakeVector(MeshVerticies(MeshIndicies(cnt + 0)).X, MeshVerticies(MeshIndicies(cnt + 0)).Y, MeshVerticies(MeshIndicies(cnt + 0)).Z)
        v2 = MakeVector(MeshVerticies(MeshIndicies(cnt + 1)).X, MeshVerticies(MeshIndicies(cnt + 1)).Y, MeshVerticies(MeshIndicies(cnt + 1)).Z)
        v3 = MakeVector(MeshVerticies(MeshIndicies(cnt + 2)).X, MeshVerticies(MeshIndicies(cnt + 2)).Y, MeshVerticies(MeshIndicies(cnt + 2)).Z)

        D3DXVec3TransformCoord v1, v1, matScale
        D3DXVec3TransformCoord v2, v2, matScale
        D3DXVec3TransformCoord v3, v3, matScale

        D3DXVec3TransformCoord v1, v1, matRotate
        D3DXVec3TransformCoord v2, v2, matRotate
        D3DXVec3TransformCoord v3, v3, matRotate
        
        v1.X = v1.X + Origin.X
        v2.X = v2.X + Origin.X
        v3.X = v3.X + Origin.X

        v1.Y = v1.Y + Origin.Y
        v2.Y = v2.Y + Origin.Y
        v3.Y = v3.Y + Origin.Y

        v1.Z = v1.Z + Origin.Z
        v2.Z = v2.Z + Origin.Z
        v3.Z = v3.Z + Origin.Z

        vn = TriangleNormal(v1, v2, v3)
        
        AddVisFace Brush, ((cnt + 1) \ 3), vn, v1, v2, v3

    Next

End Sub


Public Sub LoadLevel(ByVal FileName As String)
    ReDim Level.Walls(0 To 0) As MyWall
    ReDim Level.matWall(0 To 0) As D3DMATRIX
    ReDim Level.Starts(0 To 0) As MyPlace
    ReDim Level.Ends(0 To 0) As MyPlace
    Level.Loaded = ""
    
    ResetCollision
    
    ParseLevel Replace(ReadFile(FileName), vbTab, "")

    Player.Location = SquareCenter(Level.Starts(1).Point1, Level.Starts(1).Point2, Level.Starts(1).Point3, Level.Starts(1).Point4)
    Player.Location.Y = (160 / 2 + 10)
    
    Player.CameraAngle = 135
    
    Level.BestScore = 0
    Level.LastScore = 0
    
    Level.Loaded = GetFileTitle(FileName)

    If Not DisableSound Then
        Track2.VolumeTo 0
        Track1.VolumeTo 1000
    End If
End Sub

Private Sub ParseLevel(ByVal inData As String)

    Dim X1 As Single
    Dim X2 As Single
    Dim Y1 As Single
    Dim Y2 As Single
    Dim DT As Long
    
    Dim PosX As Single
    Dim PosZ As Single
        
    Dim inLine As String
    Do Until inData = ""
        inLine = RemoveNextArg(inData, vbCrLf)
        X1 = RemoveNextArg(inLine, ":")
        Y1 = RemoveNextArg(inLine, ":")
        X2 = RemoveNextArg(inLine, ":")
        Y2 = RemoveNextArg(inLine, ":")
        DT = RemoveNextArg(inLine, ":")
        
        If DT <= 3 Then
        
            ReDim Preserve Level.Walls(0 To UBound(Level.Walls) + 1) As MyWall
            ReDim Preserve Level.matWall(0 To UBound(Level.matWall) + 1) As D3DMATRIX
            Level.Walls(UBound(Level.Walls)).ObjectIndex = DT + 1

            If X1 > X2 Then PosX = X2 Else PosX = X1
            If Y1 > Y2 Then PosZ = Y2 Else PosZ = Y1
            
            PosX = (PosX * (FrameScale / 2))
            PosZ = (PosZ * (FrameScale / 2))

            If Y1 = Y2 Then
                Level.Walls(UBound(Level.Walls)).Rotated.Y = 90
            Else
                PosZ = PosZ + (FrameScale / 4)
                PosX = PosX - (FrameScale / 4)
            End If
            
            Level.Walls(UBound(Level.Walls)).Origin.X = PosX
            Level.Walls(UBound(Level.Walls)).Origin.Z = PosZ

            D3DXMatrixIdentity matTemp
            D3DXMatrixIdentity Level.matWall(UBound(Level.matWall))

            D3DXMatrixMultiply Level.matWall(UBound(Level.matWall)), Level.matWall(UBound(Level.matWall)), matObj(Level.Walls(UBound(Level.Walls)).ObjectIndex)
            
            D3DXMatrixRotationX matTemp, Level.Walls(UBound(Level.Walls)).Rotated.X * (PI / 180)
            D3DXMatrixMultiply Level.matWall(UBound(Level.matWall)), Level.matWall(UBound(Level.matWall)), matTemp
            D3DXMatrixRotationY matTemp, Level.Walls(UBound(Level.Walls)).Rotated.Y * (PI / 180)
            D3DXMatrixMultiply Level.matWall(UBound(Level.matWall)), Level.matWall(UBound(Level.matWall)), matTemp
            D3DXMatrixRotationZ matTemp, Level.Walls(UBound(Level.Walls)).Rotated.Z * (PI / 180)
            D3DXMatrixMultiply Level.matWall(UBound(Level.matWall)), Level.matWall(UBound(Level.matWall)), matTemp

            D3DXMatrixTranslation matTemp, Level.Walls(UBound(Level.Walls)).Origin.X, Level.Walls(UBound(Level.Walls)).Origin.Y, Level.Walls(UBound(Level.Walls)).Origin.Z
            D3DXMatrixMultiply Level.matWall(UBound(Level.matWall)), Level.matWall(UBound(Level.matWall)), matTemp
            
            AddCollisionObject UBound(Level.Walls), Level.Walls(UBound(Level.Walls)).Origin, Objects(Level.Walls(UBound(Level.Walls)).ObjectIndex).Scaled, _
                Level.Walls(UBound(Level.Walls)).Rotated, Objects(Level.Walls(UBound(Level.Walls)).ObjectIndex).XYZ, Objects(Level.Walls(UBound(Level.Walls)).ObjectIndex).idx
                    
        ElseIf (DT = 4) Then  'starting
            ReDim Preserve Level.Starts(0 To UBound(Level.Starts) + 1) As MyPlace
            
            If X1 > X2 Then Swap X1, X2
            If Y1 > Y2 Then Swap Y1, Y2
            
            Level.Starts(UBound(Level.Starts)).Point1.X = (X1 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Starts(UBound(Level.Starts)).Point1.Z = (Y1 * (FrameScale / 2))
            
            Level.Starts(UBound(Level.Starts)).Point2.X = (X1 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Starts(UBound(Level.Starts)).Point2.Z = (Y2 * (FrameScale / 2))
   
            Level.Starts(UBound(Level.Starts)).Point3.X = (X2 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Starts(UBound(Level.Starts)).Point3.Z = (Y2 * (FrameScale / 2))
            
            Level.Starts(UBound(Level.Starts)).Point4.X = (X2 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Starts(UBound(Level.Starts)).Point4.Z = (Y1 * (FrameScale / 2))
            
        ElseIf DT = 5 Then 'ending
            ReDim Preserve Level.Ends(0 To UBound(Level.Ends) + 1) As MyPlace
            
            If X1 > X2 Then Swap X1, X2
            If Y1 > Y2 Then Swap Y1, Y2
            
            Level.Ends(UBound(Level.Ends)).Point1.X = (X1 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Ends(UBound(Level.Ends)).Point1.Z = (Y1 * (FrameScale / 2))
            
            Level.Ends(UBound(Level.Ends)).Point2.X = (X1 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Ends(UBound(Level.Ends)).Point2.Z = (Y2 * (FrameScale / 2))
   
            Level.Ends(UBound(Level.Ends)).Point3.X = (X2 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Ends(UBound(Level.Ends)).Point3.Z = (Y2 * (FrameScale / 2))
            
            Level.Ends(UBound(Level.Ends)).Point4.X = (X2 * (FrameScale / 2)) - (FrameScale / 4)
            Level.Ends(UBound(Level.Ends)).Point4.Z = (Y1 * (FrameScale / 2))
        End If
        
    Loop

End Sub

Private Sub DebugFVF(ByRef desc As D3DVERTEXBUFFER_DESC)
    Dim fvf As Long
    fvf = desc.fvf
    If CheckFVF(fvf, D3DFVF_DIFFUSE) Then
        Debug.Print "?D3DFVF_DIFFUSE"
    End If
    If CheckFVF(fvf, D3DFVF_LASTBETA_UBYTE4) Then
        Debug.Print "?D3DFVF_LASTBETA_UBYTE4"
    End If
    If CheckFVF(fvf, D3DFVF_NORMAL) Then
        Debug.Print "?D3DFVF_NORMAL"
    End If
    If CheckFVF(fvf, D3DFVF_POSITION_MASK) Then
        Debug.Print "?D3DFVF_POSITION_MASK"
    End If
    If CheckFVF(fvf, D3DFVF_PSIZE) Then
        Debug.Print "?D3DFVF_PSIZE"
    End If
    If CheckFVF(fvf, D3DFVF_RESERVED0) Then
        Debug.Print "?D3DFVF_RESERVED0"
    End If
    If CheckFVF(fvf, D3DFVF_RESERVED2) Then
        Debug.Print "?D3DFVF_RESERVED2"
    End If
    If CheckFVF(fvf, D3DFVF_SPECULAR) Then
        Debug.Print "?D3DFVF_SPECULAR"
    End If
    If CheckFVF(fvf, D3DFVF_TEX0) Then
        Debug.Print "?D3DFVF_TEX0"
    End If
    If CheckFVF(fvf, D3DFVF_TEX1) Then
        Debug.Print "?D3DFVF_TEX1"
    End If
    If CheckFVF(fvf, D3DFVF_TEX2) Then
        Debug.Print "?D3DFVF_TEX2"
    End If
    If CheckFVF(fvf, D3DFVF_TEX3) Then
        Debug.Print "?D3DFVF_TEX3"
    End If
    If CheckFVF(fvf, D3DFVF_TEX4) Then
        Debug.Print "?D3DFVF_TEX4"
    End If
    If CheckFVF(fvf, D3DFVF_TEX5) Then
        Debug.Print "?D3DFVF_TEX5"
    End If
    If CheckFVF(fvf, D3DFVF_TEX6) Then
        Debug.Print "?D3DFVF_TEX6"
    End If
    If CheckFVF(fvf, D3DFVF_TEX7) Then
        Debug.Print "?D3DFVF_TEX7"
    End If
    If CheckFVF(fvf, D3DFVF_TEX8) Then
        Debug.Print "?D3DFVF_TEX8"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE1_0) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE1_0"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE1_1) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE1_1"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE1_2) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE1_2"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE1_3) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE1_3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE2_0) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE2_0"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE2_1) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE21_1"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE2_2) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE2_2"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE2_3) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE2_3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE3_0) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE3_0"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE3_1) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE3_1"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE3_2) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE3_2"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE3_3) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE3_3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE3_3) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE3_3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE4_0) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE4_0"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE4_1) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE4_1"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE4_2) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE4_2"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOORDSIZE4_3) Then
        Debug.Print "?D3DFVF_TEXCOORDSIZE4_3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOUNT_MASK) Then
        Debug.Print "?D3DFVF_TEXCOUNT_MASK"
    End If
    If CheckFVF(fvf, D3DFVF_TEXCOUNT_SHIFT) Then
        Debug.Print "?D3DFVF_TEXCOUNT_SHIFT"
    End If
    If CheckFVF(fvf, D3DFVF_TEXTUREFORMAT1) Then
        Debug.Print "?D3DFVF_TEXTUREFORMAT1"
    End If
    If CheckFVF(fvf, D3DFVF_TEXTUREFORMAT2) Then
        Debug.Print "?D3DFVF_TEXTUREFORMAT2"
    End If
    If CheckFVF(fvf, D3DFVF_TEXTUREFORMAT3) Then
        Debug.Print "?D3DFVF_TEXTUREFORMAT3"
    End If
    If CheckFVF(fvf, D3DFVF_TEXTUREFORMAT4) Then
        Debug.Print "?D3DFVF_TEXTUREFORMAT4"
    End If
    If CheckFVF(fvf, D3DFVF_VERTEX) Then
        Debug.Print "?D3DFVF_VERTEX"
    End If
    If CheckFVF(fvf, D3DFVF_XYZ) Then
        Debug.Print "?D3DFVF_XYZ"
    End If
    If CheckFVF(fvf, D3DFVF_XYZB1) Then
        Debug.Print "?D3DFVF_XYZB1"
    End If
    If CheckFVF(fvf, D3DFVF_XYZB2) Then
        Debug.Print "?D3DFVF_XYZB2"
    End If
    If CheckFVF(fvf, D3DFVF_XYZB3) Then
        Debug.Print "?D3DFVF_XYZB3"
    End If
    If CheckFVF(fvf, D3DFVF_XYZB4) Then
        Debug.Print "?D3DFVF_XYZB4"
    End If
    If CheckFVF(fvf, D3DFVF_XYZB5) Then
        Debug.Print "?D3DFVF_XYZB5"
    End If
    If CheckFVF(fvf, D3DFVF_XYZRHW) Then
        Debug.Print "?D3DFVF_XYZRHW"
    End If
    Debug.Print ""
End Sub
Private Function CheckFVF(ByRef val As Long, ByRef fvf As Long) As Boolean
    If ((val Or fvf) = fvf) Then
        CheckFVF = True
        val = val - fvf
    End If
End Function
Public Sub CleanupLawn()
    Dim q As Integer
    Dim o As Integer
    If UBound(Objects) - LBound(Objects) > 0 Then
        For o = 1 To UBound(Objects)
            For q = LBound(Objects(o).Tex) To UBound(Objects(o).Tex)
                Set Objects(o).Tex(q) = Nothing
            Next q

            ReDim Objects(o).mat(0 To 0) As D3DMATERIAL8
            ReDim Objects(o).Tex(0 To 0) As Direct3DTexture8
            ReDim Objects(o).XYZ(0 To 0) As D3DVERTEX
            ReDim Objects(o).idx(0 To 0) As Integer
            Set Objects(o).mesh = Nothing
        Next
        ReDim Objects(1 To 1) As MyMesh
        nObject = 0
    End If

    If Not (Level.Loaded = "") Then
        ReDim Level.Walls(0 To 0) As MyWall
        ReDim Level.matWall(0 To 0) As D3DMATRIX
        ReDim Level.Starts(0 To 0) As MyPlace
        ReDim Level.Ends(0 To 0) As MyPlace
        Level.Loaded = ""
    End If
    
    If UBound(Lights) - LBound(Lights) > 0 Then
        ReDim Lights(1 To 1) As D3DLIGHT8
        nLight = 0
    End If
    
End Sub



