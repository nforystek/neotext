Attribute VB_Name = "modProj"
Option Explicit

Public Type MyFile
    Data As Direct3DTexture8
    Path As String
    Size As ImageDimensions
End Type

Public Files() As MyFile
Public FileCount As Long

Public Lights() As D3DLIGHT8
Public LightCount As Long

Private Mirrors As NTNodes10.Collection
Public worldRotate As New Point

Public Sub CreateProj()
        
    frmMain.Startup
    
    If ScriptRoot = "" Then
        If PathExists(CurDir & "\Index.vbx") Then
            ScriptRoot = CurDir
        ElseIf PathExists(AppPath(False) & "Index.vbx") Then
            ScriptRoot = Left(AppPath(False), Len(AppPath(False)) - 1)
        ElseIf PathExists(AppPath(True) & "Index.vbx") Then
            ScriptRoot = Left(AppPath(True), Len(AppPath(True)) - 1)
        ElseIf PathExists(GetFilePath(AppEXE(False)) & "\Index.vbx") Then
            ScriptRoot = GetFilePath(AppEXE(False))
        End If
        If ScriptRoot = "" Then
            ScriptRoot = modFolders.SearchPath("Index.vbx", True, CurDir, FirstOnly)
            If ScriptRoot <> "" Then ScriptRoot = GetFilePath(ScriptRoot)
        End If
    End If

    If PathExists(ScriptRoot & "\Index.vbx") Then
        ParseScript ScriptRoot & "\Index.vbx"
        
    End If
    
End Sub

Private Sub RemoveIterate(ByRef Molecules As Molecules)
    If Not Molecules Is Nothing Then
        Dim m As Molecule
        Do While Molecules.Count > 0
            RemoveIterate Molecules(1).Molecules
            Molecules.Remove 1
        Loop
    End If
End Sub
Public Sub CleanUpProj()
    Dim ser As String
    ser = frmMain.Serialize
    If ser <> "" Then WriteFile ScriptRoot & "\Serial.xml", ser
    
    frmMain.ScriptControl1.Reset
    
    
    Do While Planets.Count > 0
        RemoveIterate Planets(1).Molecules
        Planets.Remove 1
    Loop

    Do While Molecules.Count > 0
        RemoveIterate Molecules(1).Molecules
        Molecules.Remove 1
    Loop
    
    Do While Brilliants.Count > 0
        Brilliants.Remove 1
    Loop

    Do While Billboards.Count > 0
        Billboards.Remove 1
    Loop
    
    Do While OnEvents.Count > 0
        OnEvents.Remove 1
    Loop
    
    Do While All.Count > 0
        All.Remove 1
    Loop
    Dim o As Long
    
    If LightCount > 0 Then
        Erase Lights
        LightCount = 0
    End If
    
    If FileCount > 0 Then
        For o = 1 To FileCount
            Set Files(o).Data = Nothing
            Files(o).Path = ""
        Next
        Erase Files
        FileCount = 0
    End If

End Sub

Public Sub RenderBrilliants(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
    
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetVertexShader FVF_RENDER
    
    Dim fogSTate As Boolean
    fogSTate = DDevice.GetRenderState(D3DRS_FOGENABLE)
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_ZENABLE, False
    
    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld

    Dim b As Brilliant
    Dim l As Long
    If Brilliants.Count > 0 Then
        For l = 1 To Brilliants.Count
            Set b = Brilliants(l)
            b.UpdateValues
            If Brilliants(l).SunLight Then
                DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(0, Lights(Brilliants(l).LightIndex).Diffuse.A * 164 + Lights(Brilliants(l).LightIndex).Diffuse.r * 255, _
                    Lights(Brilliants(l).LightIndex).Diffuse.A * 164 + Lights(Brilliants(l).LightIndex).Diffuse.g * 255, Lights(Brilliants(l).LightIndex).Diffuse.A * 164 + Lights(Brilliants(l).LightIndex).Diffuse.b * 255)

            End If
            Set b = Nothing
        Next
    End If

End Sub

Public Function BlendValue(ByVal StartMinimum As Single, ByVal StartMaximum As Single, ByVal StartFactor As Single, ByVal StopMinimum As Single, ByVal StopMaximum As Single, ByVal StopFactor As Single, ByVal CurrentFactor As Single) As Single

    BlendValue = (((StartMaximum - StartMinimum) * StartFactor) + ((((StopMaximum - StopMinimum) * StopFactor) - ((StartMaximum - StartMinimum) * StartFactor)) * CurrentFactor))
    
End Function
Private Sub SubRenderWorldSetup(ByRef UserControl As Macroscopic, ByRef Camera As Camera, ByVal StartOrStop As Boolean)
    Static matSave As D3DMATRIX
    If StartOrStop Then
        'do start
        DDevice.GetTransform D3DTS_VIEW, matSave
        matView = matSave
        matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
        DDevice.SetTransform D3DTS_VIEW, matView
        D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
            ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), 0, Far
        DDevice.SetTransform D3DTS_PROJECTION, matProj
    Else
        'do stop
        DDevice.SetTransform D3DTS_VIEW, matSave
        matView = matSave
        DDevice.SetTransform D3DTS_WORLD, matWorld
        D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
            ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), Near, Far
        DDevice.SetTransform D3DTS_PROJECTION, matProj
    End If
End Sub

Private Sub SubRenderPlateau(ByRef UserControl As Macroscopic, ByRef Camera As Camera, ByRef p As Planet)
    With p
   ' Debug.Print p.Key
        Dim matPlane As D3DMATRIX
        Dim matRot As D3DMATRIX
        Dim matPos As D3DMATRIX
        Dim matScale As D3DMATRIX
        Dim matYaw As D3DMATRIX
        Dim matPitch As D3DMATRIX
        Dim matRoll As D3DMATRIX
                
        D3DXMatrixIdentity matPlane
        D3DXMatrixIdentity matPos
        D3DXMatrixIdentity matYaw
        D3DXMatrixIdentity matPitch
        D3DXMatrixIdentity matRoll
        D3DXMatrixIdentity matScale
        
                
        DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(255, 255, 255, 255)
        
        'render the round portion first
        If p.Volume.Count > 0 Then
        
            Dim lineto As Point
            'If (p.PlateauInfinite Or p.PlateauHole) Then 'first 12 triangles is the rolling backdrop
                
               
            Dim testFar As Long
            testFar = Far '10 * MILE
            
            If (Not ((Abs(Camera.Player.Origin.X - p.Origin.X) > (testFar / 2)) Or (Abs(Camera.Player.Origin.Y - p.Origin.Y) > (testFar / 2)) Or (Abs(Camera.Player.Origin.z - p.Origin.z) > (testFar / 2)))) And p.PlateauHole Then
                'draws hole type of plane if in range of the hole, else the infinite plane is used further down
            
                D3DXMatrixTranslation matPos, p.Origin.X, p.Origin.Y, p.Origin.z
                D3DXMatrixMultiply matPlane, matPlane, matPos
                
                DDevice.SetTransform D3DTS_WORLD, matPlane
        
                D3DXMatrixRotationZ matRoll, p.Rotate.z
                D3DXMatrixMultiply matPlane, matRoll, matPlane
        
                D3DXMatrixRotationY matYaw, p.Rotate.Y
                D3DXMatrixMultiply matPlane, matYaw, matPlane
                
                D3DXMatrixRotationX matPitch, p.Rotate.X
                D3DXMatrixMultiply matPlane, matPitch, matPlane
                
                D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
                D3DXMatrixMultiply matPlane, matScale, matPlane
        
                DDevice.SetTransform D3DTS_WORLD, matPlane
        
                With p.Volume((p.Volume.Count / 3) + 1)
        
                    SetRenderBlends .Transparent, .Translucent
                    If Not (.Translucent Or .Transparent) Then
                        DDevice.SetMaterial GenericMaterial
                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                        DDevice.SetTexture 1, Nothing
                    Else
                        DDevice.SetMaterial LucentMaterial
                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                        DDevice.SetMaterial GenericMaterial
                        If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
                    End If
                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, ((p.Volume.Count / 3) * 2), VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
        
                End With
        
            Else
                If Not ((Abs(Camera.Player.Origin.X - p.Origin.X) > (testFar / 2)) Or (Abs(Camera.Player.Origin.Y - p.Origin.Y) > (testFar / 2)) Or (Abs(Camera.Player.Origin.z - p.Origin.z) > (testFar / 2))) Then
                'draws island and doughnut stype plane
                
                    If ((p.Rows > 0) And (p.Columns > 0)) And (p.PlateauIsland Or p.PlateauDoughnut) Then
                    'draws a grid of island and doughnut stype plane it one exists a grid
                    
                        Dim j As Long
                        Dim X As Single
                        Dim z As Single
                        X = ((p.Rows \ 2) * ((p.OuterEdge * 2) + p.Field))
                        z = ((p.Columns \ 2) * ((p.OuterEdge * 2) + p.Field))
        
                        For j = 0 To ((p.Rows * p.Columns) - 1)
            
                            If j Mod p.Columns = 0 Then
                                z = z - ((p.OuterEdge * 2) + p.Field)
                                X = -((p.Rows \ 2) * ((p.OuterEdge * 2) + p.Field))
                            Else
                                X = X + ((p.OuterEdge * 2) + p.Field)
                            End If
                            
                            With p.Volume(1)
                                
                                D3DXMatrixTranslation matPos, (p.Origin.X \ (testFar / 2)) * (testFar / 2) + X, _
                                    (p.Origin.Y \ (testFar / 2)) * (testFar / 2), (p.Origin.z \ (testFar / 2)) * (testFar / 2) + z
                                D3DXMatrixMultiply matPlane, matPlane, matPos
            
                                DDevice.SetTransform D3DTS_WORLD, matPlane
            
                                D3DXMatrixRotationZ matRoll, p.Rotate.z
                                D3DXMatrixMultiply matPlane, matRoll, matPlane
            
                                D3DXMatrixRotationY matYaw, p.Rotate.Y
                                D3DXMatrixMultiply matPlane, matYaw, matPlane
            
                                D3DXMatrixRotationX matPitch, p.Rotate.X
                                D3DXMatrixMultiply matPlane, matPitch, matPlane
            
                                D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
                                D3DXMatrixMultiply matPlane, matScale, matPlane
                                        
                                DDevice.SetTransform D3DTS_WORLD, matPlane
        
                                SetRenderBlends .Transparent, .Translucent
                                If Not (.Translucent Or .Transparent) Then
                                    DDevice.SetMaterial GenericMaterial
                                    If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                                    DDevice.SetTexture 1, Nothing
                                Else
                                    DDevice.SetMaterial LucentMaterial
                                    If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                                    DDevice.SetMaterial GenericMaterial
                                    If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
                                End If
                                DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, p.Volume.Count, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
        
                            End With
            
                        Next
                    
                    ElseIf (p.PlateauIsland Or p.PlateauDoughnut) Then
                    'draws island and doughnut stype plane where as no grid of them exists
                                        
                        D3DXMatrixTranslation matPos, p.Origin.X, p.Origin.Y, p.Origin.z
                        D3DXMatrixMultiply matPlane, matPlane, matPos
                        
                        DDevice.SetTransform D3DTS_WORLD, matPlane
        
                        D3DXMatrixRotationZ matRoll, p.Rotate.z
                        D3DXMatrixMultiply matPlane, matRoll, matPlane
        
                        D3DXMatrixRotationY matYaw, p.Rotate.Y
                        D3DXMatrixMultiply matPlane, matYaw, matPlane
                        
                        D3DXMatrixRotationX matPitch, p.Rotate.X
                        D3DXMatrixMultiply matPlane, matPitch, matPlane
                        
                        D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
                        D3DXMatrixMultiply matPlane, matScale, matPlane
        
                        DDevice.SetTransform D3DTS_WORLD, matPlane
                
                        With p.Volume(1)
        
                            SetRenderBlends .Transparent, .Translucent
                            If Not (.Translucent Or .Transparent) Then
                                DDevice.SetMaterial GenericMaterial
                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                                DDevice.SetTexture 1, Nothing
                            Else
                                DDevice.SetMaterial LucentMaterial
                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                                DDevice.SetMaterial GenericMaterial
                                If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
                            End If
                            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, p.Volume.Count, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
        
                        End With
                    End If
        
                End If
                If (p.PlateauInfinite Or p.PlateauHole) Then
                'draws a infinite stype plane, no escaping it
        
                    D3DXMatrixTranslation matPos, ((Camera.Player.Origin.X - p.Origin.X) \ (testFar / 2)) * (testFar / 2), _
                         (p.Origin.Y \ (testFar / 2)) * (testFar / 2), ((Camera.Player.Origin.z - p.Origin.z) \ (testFar / 2)) * (testFar / 2)
                    D3DXMatrixMultiply matPlane, matPlane, matPos
                    
                    DDevice.SetTransform D3DTS_WORLD, matPlane
        
                    D3DXMatrixRotationZ matRoll, p.Rotate.z
                    D3DXMatrixMultiply matPlane, matRoll, matPlane
        
                    D3DXMatrixRotationY matYaw, p.Rotate.Y
                    D3DXMatrixMultiply matPlane, matYaw, matPlane
                    
                    D3DXMatrixRotationX matPitch, p.Rotate.X
                    D3DXMatrixMultiply matPlane, matPitch, matPlane
                    
                    D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
                    D3DXMatrixMultiply matPlane, matScale, matPlane
        
                    DDevice.SetTransform D3DTS_WORLD, matPlane
            
                    With p.Volume(1)
        
                        SetRenderBlends .Transparent, .Translucent
                        If Not (.Translucent Or .Transparent) Then
                            DDevice.SetMaterial GenericMaterial
                            If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                            DDevice.SetTexture 1, Nothing
                        Else
                            DDevice.SetMaterial LucentMaterial
                            If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
                            DDevice.SetMaterial GenericMaterial
                            If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
                        End If
                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, p.Volume.Count, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
        
                    End With
                End If
        
        
            End If
        
        End If
    End With
    
End Sub

Public Sub SubRenderWorld(ByRef UserControl As Macroscopic, ByRef Camera As Camera, ByRef p As Planet, ByVal RelativeFactor As Single)

    Dim matPlane As D3DMATRIX
    Dim matRot As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matScale As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matRoll As D3DMATRIX
            
    D3DXMatrixIdentity matPlane
    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matYaw
    D3DXMatrixIdentity matPitch
    D3DXMatrixIdentity matRoll
    D3DXMatrixIdentity matScale
    
    If (Not p.Follow) Then
        
        D3DXMatrixTranslation matPos, -p.Origin.X, -p.Origin.Y, -p.Origin.z
        D3DXMatrixMultiply matPlane, matPlane, matPos
                
    Else
    
        D3DXMatrixTranslation matPos, 0, 0, 0
        D3DXMatrixMultiply matPlane, matPlane, matPos
        
    End If
        
    DDevice.SetTransform D3DTS_WORLD, matPlane
    
    D3DXMatrixRotationZ matRoll, p.Rotate.z
    D3DXMatrixMultiply matPlane, matRoll, matPlane

    D3DXMatrixRotationY matYaw, p.Rotate.Y
    D3DXMatrixMultiply matPlane, matYaw, matPlane
    
    D3DXMatrixRotationX matPitch, p.Rotate.X
    D3DXMatrixMultiply matPlane, matPitch, matPlane
    
    D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
    D3DXMatrixMultiply matPlane, matScale, matPlane

    DDevice.SetTransform D3DTS_WORLD, matPlane
        
    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault
               
    Dim i As Long
    Dim rsam As Single
    
    rsam = DDevice.GetRenderState(D3DRS_AMBIENT)
          

    If p.Alphablend Then
       
        DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(RelativeFactor * 255, 255 * RelativeFactor, 255 * RelativeFactor, 255 * RelativeFactor)
     
     
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVDESTCOLOR And D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTALPHA And D3DBLEND_INVSRCCOLOR
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
        
        For i = 1 To p.Volume.Count Step 2
            DDevice.SetMaterial GenericMaterial
            DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
            DDevice.SetMaterial GenericMaterial
            DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
        Next
            
    ElseIf (Not p.Translucent) And (Not p.Alphablend) Then
        
        DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB((1 - RelativeFactor) * 255, 255 * RelativeFactor, 255 * RelativeFactor, 255 * RelativeFactor)
        
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
                                            
        For i = 1 To p.Volume.Count Step 2
            DDevice.SetMaterial LucentMaterial
            DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
            DDevice.SetMaterial GenericMaterial
            DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
        Next
            
    Else

        DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(RelativeFactor * 255, 192 * RelativeFactor, 192 * RelativeFactor, 192 * RelativeFactor)

        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

         For i = 1 To p.Volume.Count Step 2
            DDevice.SetMaterial LucentMaterial
            DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
            DDevice.SetMaterial GenericMaterial
            DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
        Next
        
    End If
    
    DDevice.SetRenderState D3DRS_AMBIENT, rsam
    
'    Else
'
'        'rsam = DDevice.GetRenderState(D3DRS_AMBIENT)
'        'DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB((RelativeFactor) * 255, 255 * RelativeFactor, 255 * RelativeFactor, 255 * RelativeFactor)
'
'        If p.Transparent Then
'
'            DDevice.SetRenderState D3DRS_ZENABLE, 1
'            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'
''            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
''            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'            For i = 1 To p.Volume.Count Step 2
'                DDevice.SetMaterial LucentMaterial
'                DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
'                DDevice.SetMaterial GenericMaterial
'                DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
'                DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
'            Next
'
'        ElseIf p.Translucent Then
'
'        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
'        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTALPHA
''
''            DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
''            DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'
'          '  DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
'        'DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTALPHA
'
'            DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'            DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
'
'            For i = 1 To p.Volume.Count Step 2
'                DDevice.SetMaterial LucentMaterial
'                DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
'                DDevice.SetMaterial GenericMaterial
'                DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
'                DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
'            Next
'
'        End If
'
'      '  DDevice.SetRenderState D3DRS_AMBIENT, rsam
'
'    End If

End Sub

Public Sub RenderPlanets(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
       
    Dim dist As Single
    Dim dist2 As Single
    Dim dist3 As Single
    Dim dist4 As Single
    Dim i As Long
    Dim j As Long
    Dim p As Planet
    Dim p2 As Planet
    Dim rsam As Single
    Dim v As Long
    
    Dim onkey As String 'setup camera key during this function state
  
'    Dim total As Long
'    Dim sumx As Single
'    Dim sumz As Single
'    Dim sumy As Single
'    Dim sum As Single
'    Dim maxdist As Single

    Dim aimingAt As Planet
    Dim nearest As Planet
    Dim tmp As Point
    
    If Not Camera.Planet Is Nothing Then onkey = Camera.Planet.Key
    If Not Camera.Player Is Nothing Then AngleAxisRestrict Camera.Player.Rotate
    Camera.BuildColor
        

'#####################################################################################################
'#####################################################################################################
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT

'    DDevice.SetMaterial GenericMaterial
'    DDevice.SetTexture 1, Nothing
'    DDevice.SetMaterial LucentMaterial
'    DDevice.SetTexture 0, Nothing
    
    SubRenderWorldSetup UserControl, Camera, True  'must be called again, tiwce per one call
    
    If (Planets.Count > 0) And (Not Camera.Player Is Nothing) Then
        'first loop through all, we render worlds that
        'are opaque, and calculate up stuff for plateaus
        
        For Each p In Planets
            
            Select Case p.Form
                Case World
                    If p.Visible Then
                        If Not p.Follow And Not Camera.Player Is Nothing Then
                            dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, Camera.Player.Origin.X, Camera.Player.Origin.Y, Camera.Player.Origin.z)
                        ElseIf Not Camera.Planet Is Nothing Then
                            dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, Camera.Planet.Origin.X, Camera.Planet.Origin.Y, Camera.Planet.Origin.z)
                        End If
                        dist = p.RelativeColorFactor(dist)
                        Camera.BuildColor p.color.RGB, dist

                        If (dist > 0) And (Not (p.Translucent Or p.Transparent)) Then
                            

                           ' Debug.Print p.Key; dist

                            SubRenderWorld UserControl, Camera, p, dist

                        Else
                        
                        End If
                    End If
                Case Plateau

                    dist = DistanceEx(Camera.Player.Origin, p.Origin)

                    If p.Visible Then
                        dist2 = p.RelativeColorFactor(dist)
                        Camera.BuildColor p.color.RGB, dist2
                        
'                        'get running sums for universe calsulations for
'                        'planes in collision to scale against spread out
'                        total = total + 1 'hold total of plateau only
'                        sum = sum + ((p.InnerEdge * 2) + (p.OuterEdge * 2))
'                        If sumx < Abs(p.Origin.X) Then sumx = Abs(p.Origin.X)
'                        If sumy < Abs(p.Origin.Y) Then sumy = Abs(p.Origin.Y)
'                        If sumz < Abs(p.Origin.z) Then sumz = Abs(p.Origin.z)
                        
                        'find nearest planet (dist4 holds closest last dist)
                        If dist < dist4 Or dist4 = 0 Then
                            dist4 = dist
                            Set nearest = p
                        End If
                        
                        'find aiming at plaet (dist3 holds closest last aiming at dist)
                        Set tmp = VectorAxisAngles(VectorNegative(VectorDeduction(Camera.Player.Origin, p.Origin)))
                        Set tmp = AngleAxisDifference(VectorAxisAngles(VectorRotateAxis(MakePoint(0, 0, 1), Camera.Player.Rotate)), tmp)
                        
                        dist2 = VectorQuantify(tmp)
                        If dist3 = 0 Or dist2 < dist3 Then
                            dist3 = dist2
                            Set aimingAt = p
                        End If
                        'set camera.planet based on if we enter/leave radius
                        If onkey = p.Key Then
                            If dist > p.OuterEdge + p.Field Then
                                Set Camera.Planet = Nothing
                            End If
                        Else
                            If dist <= p.OuterEdge And onkey = "" Then
                                Set Camera.Planet = p
                            Else
                                
                            End If
                        End If
                        
                    End If

            End Select
            
        Next
        
'        'find greatest dist of x/y/z for max universe size
'        If sumx >= sumy And sumx >= sumz Then maxdist = sumx * 2
'        If sumy >= sumx And sumy >= sumz Then maxdist = sumy * 2
'        If sumz >= sumy And sumz >= sumx Then maxdist = sumz * 2
                
    End If

    SubRenderWorldSetup UserControl, Camera, False  'the second call returns the view state

'#####################################################################################################

    DDevice.SetRenderState D3DRS_LIGHTING, 1
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    
'    DDevice.SetRenderState D3DRS_AMBIENT, Camera.color.ARGB
   ' DDevice.SetRenderState D3DRS_FOGCOLOR, D3DColorARGB(1, red, g, b)

'    DDevice.SetRenderState D3DRS_FOGENABLE, 1
'    DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
'    DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_NONE
'    DDevice.SetRenderState D3DRS_RANGEFOGENABLE, False
'    DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(50 / 2)
'    DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(50)
'    DDevice.SetRenderState D3DRS_FOGDENSITY, 0
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetTexture 1, Nothing
  
    If (Planets.Count > 0) And (Not Camera.Player Is Nothing) Then

        If onkey = "" And (Not Camera.Planet Is Nothing) Then onkey = Camera.Planet.Key
        'sort the planet we are on, to the last one to be rendered
        If (onkey <> "") And Planets(Planets.Count).Key <> onkey Then
            Set p = Planets(onkey)
            Planets.Remove onkey
            Planets.Add p, onkey
        End If
        
        'at this point we have nearest planet, aiming at planet and planet we are on
'        Debug.Print "Nearest: " & nearest.Key;
'        Debug.Print " Aimingat: " & aimingAt.Key;
'        Debug.Print " OnPlanet: " & onkey
        
        DDevice.SetRenderState D3DRS_ZENABLE, 0
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

        'DDevice.SetRenderState D3DRS_ZENABLE, 1
        'DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

        i = 1
        
        Do While i <= Planets.Count
            Set p = Planets(i)
            
            If p.Visible Then
                
                If (p.Form = World) Then
                    If Camera.Player Is Nothing Then
                        dist = 0
                   ElseIf Not Camera.Planet Is Nothing Then
                        dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, Camera.Planet.Origin.X, Camera.Planet.Origin.Y, Camera.Planet.Origin.z)
                        dist = p.RelativeColorFactor(dist)
                   End If
'                    Dist = Distance(p.Origin.X, 0, p.Origin.z, 0, Camera.Player.Origin.Y, 0)
                    
                    
                    If ((dist <= 1) And (dist > 0) And (i < Planets.Count)) And (p.Translucent Or p.Transparent) Then
                        'hot sort the list to opaque World, not on plateau, then
                        'alpha blend world and finally on plateau is already last
                        j = i
                        Do
                            If (Planets(j + 1).Form = Plateau And (Not Planets(j + 1) = onkey)) Then
                                
                                Set p2 = Planets(j + 1)
                                Planets.Remove j + 1
                                If j < Planets.Count Then
                                    Planets.Add p2, p2.Key, j
                                Else
                                    Planets.Add p2, p2.Key
                                End If
                                Set p2 = Nothing
                            Else
                                Exit Do
                            End If
                            j = j + 1
                        Loop While (j < Planets.Count)
                        
                        Set p = Planets(i)
                        
                        dist = Distance(p.Origin.X, 0, p.Origin.z, 0, Camera.Player.Origin.Y, 0)
                        dist = p.RelativeColorFactor(dist)
                        
                    End If
                    
                    If p.Form = World And (dist <= 1) And (dist > 0) And (p.Translucent Or p.Transparent) Then

                        dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, Camera.Player.Origin.X, Camera.Player.Origin.Y, Camera.Player.Origin.z)
                        dist = p.RelativeColorFactor(dist)
                    
                        If ((dist <= 1) And (dist > 0)) Then
                            'draw the world with the alpha blend a relativefactor
                          '  Debug.Print p.Key; dist
                            
                            Camera.BuildColor p.color.RGB, dist
                    
                            SubRenderWorldSetup UserControl, Camera, True  'must be called again, tiwce per one call
    
                            SubRenderWorld UserControl, Camera, p, dist
                            
                            SubRenderWorldSetup UserControl, Camera, False 'the second call returns the view state
                               
                        End If

                    End If
                End If

                If (p.Form = Plateau) And onkey <> p.Key Then
                    
                    Set p.Rotate = VectorAxisAngles(VectorDeduction(Camera.Player.Origin, p.Origin))
                    Set p.Absolute.Rotate = VectorAxisAngles(VectorDeduction(Camera.Player.Origin, p.Origin))
                    
                    SubRenderPlateau UserControl, Camera, p
                    
                End If
            End If
            
            i = i + 1
            If i < Planets.Count Then Set p = Planets(i)
        Loop
        
        If p.Key = onkey Then
            'finally render the planet we are on

            SubRenderPlateau UserControl, Camera, p
        
        End If

        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

    End If

    DDevice.SetRenderState D3DRS_CLIPPING, 1

    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

End Sub

 


Private Sub SetRenderBlends(ByVal Transparent As Boolean, ByVal Translucent As Boolean)

    If DDevice.GetRenderState(D3DRS_ALPHATESTENABLE) = 0 Then DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
    If Transparent Then
        If DDevice.GetRenderState(D3DRS_SRCBLEND) <> D3DBLEND_DESTALPHA Then DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
        If DDevice.GetRenderState(D3DRS_DESTBLEND) <> D3DBLEND_DESTCOLOR Then DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
        If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) <> 0 Then DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    End If

    If Translucent Then
        If DDevice.GetRenderState(D3DRS_SRCBLEND) <> D3DBLEND_DESTALPHA Then DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
        If DDevice.GetRenderState(D3DRS_DESTBLEND) <> D3DBLEND_SRCALPHA Then DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
        If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) = 0 Then DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    End If
    
    If Not (Translucent Or Transparent) Then
        If DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) <> 0 Then DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        If DDevice.GetRenderState(D3DRS_SRCBLEND) <> D3DBLEND_SRCALPHA Then DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        If DDevice.GetRenderState(D3DRS_DESTBLEND) <> D3DBLEND_INVSRCALPHA Then DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Private Sub BackOfTheLine(ByRef line As NTNodes10.Collection)
    Dim Key As String
    Dim Obj As Object
    Key = line.Key(1)
    Set Obj = line.Item(1)
    line.Remove 1
    line.Add Obj, Key
End Sub

Public Sub BeginMirrors(ByRef UserControl As Macroscopic, ByRef Camera As Camera)

'    Dim e As Billboard
'    Dim i As Long
'    Dim l As Single
'
'    Dim dm As D3DDISPLAYMODE
'    Dim pal As PALETTEENTRY
'    Dim rct As RECT
'
'    If Not Mirrors Is Nothing Then Mirrors.Clear
'    If Billboards.Count > 0 Then
'        For i = 1 To Billboards.Count
'            Set e = Billboards(i)
'
'            If e.Visible And ((e.Form And ThreeDimensions) = ThreeDimensions) Then
'
'                If Not (Faces(e.FaceIndex).VBuffer Is Nothing) Then
'                    If Not (e.Transposing Is Nothing) Then
'
'                        l = Distance(MoleculeView.Origin.X, MoleculeView.Origin.Y, MoleculeView.Origin.Z, e.Center.X, e.Center.Y, e.Center.Z)
'                        If l <= FAR Then
'
'                            If Mirrors Is Nothing Then Set Mirrors = New NTNodes10.Collection
'
'                            DViewPort.Width = 128
'                            DViewPort.Height = 128
'
'                            DSurface.BeginScene DefaultRenderTarget, DViewPort
'                            BeginWorld UserControl, e.Transposing
'
'                            RenderPlanets UserControl, e.Transposing
'                            RenderObject UserControl, e.Transposing
'                            DSurface.EndScene
'
'                            DDevice.GetDisplayMode dm
'
'                            rct.Top = 0
'                            rct.Left = 0
'
'                            rct.Right = DViewPort.Width
'                            rct.Bottom = DViewPort.Height
'
'                            D3DX.SaveSurfaceToFile GetTemporaryFolder & "\" & Billboards.key(i) & ".bmp", D3DXIFF_BMP, DefaultRenderTarget, pal, rct
'                            Mirrors.Add D3DX.CreateTextureFromFileEx(DDevice, GetTemporaryFolder & "\" & Billboards.key(i) & ".bmp", _
'                                DViewPort.Width, DViewPort.Height, D3DX_FILTER_NONE, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
'                                D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0), Billboards.key(i)
'                            Kill GetTemporaryFolder & "\" & Billboards.key(i) & ".bmp"
'
'                        End If
'
'                    End If
'                End If
'            End If
'            Set e = Nothing
'        Next
'    End If
End Sub


Public Sub RenderMirror(ByRef UserControl As Macroscopic, ByRef Camera As Camera)

'    Dim e As Billboard
'    Dim i As Long
'    Dim l As Single
'
'    If Billboards.Count > 0 Then
'        For i = 1 To Billboards.Count
'            Set e = Billboards(i)
'
'            If e.Visible And ((e.Form And ThreeDimensions) = ThreeDimensions) Then
'
'                If Not (Faces(e.FaceIndex).VBuffer Is Nothing) Then
'                    If Not (e.Transposing Is Nothing) Then
'
'                        l = Distance(MoleculeView.Origin.X, MoleculeView.Origin.Y, MoleculeView.Origin.Z, e.Center.X, e.Center.Y, e.Center.Z)
'                        If l <= FAR Then
'
'                            If Mirrors.Exists(Billboards.key(i)) Then
'
'                                DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'                                DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                                DDevice.SetMaterial GenericMaterial
'                                DDevice.SetTexture 0, Mirrors.Item(Billboards.key(i))
'                                DDevice.SetTexture 1, Nothing
'
'                                DDevice.SetStreamSource 0, Faces(e.FaceIndex).VBuffer, Len(Faces(e.FaceIndex).Verticies(0))
'                                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'
'                            End If
'
'                        End If
'
'                    End If
'                End If
'            End If
'            Set e = Nothing
'        Next
'    End If
End Sub



