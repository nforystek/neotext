Attribute VB_Name = "modProj"
Option Explicit

Public Type MyFile
    Data As Direct3DTexture8
    path As String
    Size As ImageDimensions
End Type

Public Files() As MyFile
Public FileCount As Long

Public Lights() As D3DLIGHT8
Public LightCount As Long

Private Mirrors As NTNodes10.Collection
Public worldRotate As New Point

Public Sub CreateProj()

    If ScriptRoot <> "" Then
        frmMain.Serialize ParseScript(ScriptRoot & "\Index.vbx")
    Else
        frmMain.Serialize ""
    End If

    
    
    
End Sub

Public Sub CleanUpProj()
    Dim ser As String
    ser = frmMain.Serialize
    If ser <> "" Then WriteFile ScriptRoot & "\Serial.xml", ser
    
    frmMain.ScriptControl1.Reset
    
    Do While Planets.Count > 0
        Planets.Remove 1
    Loop

    Do While Brilliants.Count > 0
        Brilliants.Remove 1
    Loop

    Do While Molecules.Count > 0
        Molecules.Remove 1
    Loop

    Do While Billboards.Count > 0
        Billboards.Remove 1
    Loop
    
    Do While OnEvents.Count > 0
        OnEvents.Remove 1
    Loop
    
    Do While Include.Count > 0
        Include.Remove 1
    Loop
    Dim o As Long
    
    If LightCount > 0 Then
        Erase Lights
        LightCount = 0
    End If
    
    If FileCount > 0 Then
        For o = 1 To FileCount
            Set Files(o).Data = Nothing
            Files(o).path = ""
        Next
        Erase Files
        FileCount = 0
    End If

    If Not StopGame Then frmMain.Startup
End Sub

Public Sub RenderBrilliants(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    
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

Public Sub RenderPlanets(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT

    
'    Dim matSave As D3DMATRIX
'    DDevice.GetTransform D3DTS_VIEW, matSave
'    matView = matSave
'    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
'    DDevice.SetTransform D3DTS_VIEW, matView
'
'    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
'        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), 0, Far
'    DDevice.SetTransform D3DTS_PROJECTION, matProj
        

    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 1, Nothing
    DDevice.SetMaterial LucentMaterial
    DDevice.SetTexture 0, Nothing




    Dim cutoff As Single
    Dim cnt As Long
    Dim dist As Single
    Dim dist2 As Single
    Dim dist3 As Single
    Dim dist4 As Single


    Dim V As Matter
    Dim i As Long
    Dim Render As Boolean

    Dim p As Planet
    If Planets.Count > 0 Then
        For Each p In Planets

            If p.Visible Then
                Select Case p.Form
                    Case World

                        'If DistanceEx(MoleculeView.Origin, p.Origin) <  p.OuterRadius Then

                            dist = Distance(p.Origin.X, 0, p.Origin.z, 0, MoleculeView.Origin.Y, 0)
                            If dist < p.OuterRadius Then
                                If dist > (p.OuterRadius / 4) Then
                                    SetRenderBlends False, True

                                    dist = 1 - (1 * ((MoleculeView.Origin.Y - (p.OuterRadius / 4)) / (p.OuterRadius - (p.OuterRadius / 4))))
                                    Render = True
                                   ' DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(dist, Planets(onKey).Color.Red, Planets(onKey).Color.Green, Planets(onKey).Color.Blue)
                                Else
                                    Render = True
                                    dist = 1
                                    SetRenderBlends p.Transparent, p.Translucent
                                End If

                            Else
                                Render = True
                                dist = 1
                                SetRenderBlends p.Transparent, p.Translucent
                            End If

'                        Else
'                            render = False
'                            dist = 0
'                            SetRenderBlends p.Transparent, p.Translucent
'                        End If


                        If Render Then
                            If (Not (p.Transparent Or p.Translucent)) Then


'                               GenericMaterial.Ambient.A = 1
'                               GenericMaterial.Ambient.r = 1
'                               GenericMaterial.Ambient.g = 1
'                               GenericMaterial.Ambient.b = 1

                                DDevice.SetMaterial GenericMaterial
                                DDevice.SetTexture 1, Nothing
                                For i = 1 To p.Volume.Count Step 2
                                    DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
                               Next
'                               GenericMaterial.Ambient.A = 1
'                               GenericMaterial.Ambient.r = 1
'                               GenericMaterial.Ambient.g = 1
'                               GenericMaterial.Ambient.b = 1
'                               DDevice.SetMaterial GenericMaterial
                            Else
                               ' Debug.Print dist



                                DDevice.SetMaterial LucentMaterial
                                For i = 1 To p.Volume.Count Step 2
                                    DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
                               Next

                               DDevice.SetMaterial GenericMaterial
                               For i = 1 To p.Volume.Count Step 2
                                    DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
                               Next



                            End If
                        End If

                    Case Plateau




                End Select
            End If
        Next
    End If

'    DDevice.SetTransform D3DTS_VIEW, matSave
'    matView = matSave
'
'
'    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
'        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), Near, Far
'    DDevice.SetTransform D3DTS_PROJECTION, matProj
   



    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_FOGENABLE, 0
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetRenderState D3DRS_FOGCOLOR, DDevice.GetRenderState(D3DRS_AMBIENT)
    

    DDevice.SetRenderState D3DRS_ZENABLE, 1
  

    
    Dim pX As Single
    Dim pY As Single
    Dim pZ As Single
    

    'setup lastOrigin value
    Static lastOrigin As Point
    If lastOrigin Is Nothing Then
        Set lastOrigin = New Point
        lastOrigin.X = MoleculeView.Origin.X
        lastOrigin.Y = MoleculeView.Origin.Y
        lastOrigin.z = MoleculeView.Origin.z
    End If
    
    Static elapse As Single
    If elapse = 0 Then elapse = Timer
    
    
    Dim onkey As String 'setup camera key during this function state
    If Not Camera.Planet Is Nothing Then onkey = Camera.Planet.Key
    
    
    Dim lowScale As Single
    lowScale = 0.4
    Const hiScale As Single = 1
    Dim atScale As Single

    Dim sumx As Single
    Dim sumz As Single
    Dim sumy As Single
    Dim sum As Single
    Dim hi As Single
    Dim lo As Single
    Dim gap As Single

    Dim pan As Single
    Dim total As Long
    
    Dim prior As Long
    
    Dim closekey As String
    Dim Localized As New Point

    
    
    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim aimingAt As Planet
    Dim nearest As Planet
    Dim lastNearest As Planet
    

    Dim r As Range

    Dim loc As Point
    Dim norm As Point
    Dim norm2 As Point
    Dim Extend As Point
    Dim tmp As Point
    Static initials As Boolean
        
    Dim Scaled As Point
    Dim Rotate As Point
    Dim Origin As Point
    Dim dot As Single
    Dim maxdist As Single
    AngleAxisRestrict MoleculeView.Rotate
    Dim p2 As Planet

    DDevice.SetTexture 1, Nothing
    dist3 = 0
    dist4 = 0
    If Planets.Count > 0 Then
        
        total = 0
        i = 1

        Do While i <= Planets.Count

            Set p = Planets(i)
        'For Each p In Planets

            If p.Visible Then
                If (p.Form = Plateau) Then
                    total = total + 1 'hold total of plateau only

                    'get running sums for universe calsulations
                    sum = sum + ((p.InnerRadius * 2) + (p.OuterRadius * 2))
                    If sumx < Abs(p.Origin.X) Then sumx = Abs(p.Origin.X)
                    If sumy < Abs(p.Origin.Y) Then sumy = Abs(p.Origin.Y)
                    If sumz < Abs(p.Origin.z) Then sumz = Abs(p.Origin.z)
                    
                    'distance to center of planet
                    dist = DistanceEx(MoleculeView.Origin, p.Origin)

                    'find nearest planet (dist4 holds closest last dist)
                    If dist < dist4 Or dist4 = 0 Then
                        dist4 = dist
                        
                        Planets.Remove i
                        If Planets.Count > 0 Then
                            Planets.Add p, p.Key, 1
                        Else
                            Planets.Add p, p.Key
                        End If

                        Set nearest = p
                    End If
                    
                    'find aiming at plaet (dist3 holds closest last aiming at dist)
                    Set tmp = VectorAxisAngles(VectorNegative(VectorDeduction(MoleculeView.Origin, p.Origin)))
                    Set tmp = AngleAxisDifference(VectorAxisAngles(VectorRotateAxis(MakePoint(0, 0, 1), MoleculeView.Rotate)), tmp)


                    
                    dist2 = VectorQuantify(tmp)
                    If dist3 = 0 Or dist2 < dist3 Then
                            
                        dist3 = dist2
                        Set aimingAt = p

                    End If
                    'Debug.Print p.Key & " " & dist2

                    'using dist2 to keep dist trhough out
                   ' dist2 = DistanceEx(p.Origin, p.Volume(1).Point2) 'getting radius

'Debug.Print p.Key; PointSideOfPlane(p.Volume(1).Point1, p.Volume(1).Point2, p.Volume(1).Point3, MoleculeView.Origin)

                    'set camera.planet based on if we enter/leave radius
                    If onkey = p.Key Then

                        If dist > p.InnerRadius Then
                            Set Camera.Planet = Nothing
                        End If

                        
                    Else
                        
                        If dist <= p.InnerRadius And onkey = "" Then
                                    
                            Set Camera.Planet = p

                        Else
                            
                        End If
                    End If
                    
                    
                End If
            End If
            i = i + 1
            
        Loop
        'Next
        
        'find greatest dist of x/y/z for max universe size
        If sumx >= sumy And sumx >= sumz Then maxdist = sumx * 2
        If sumy >= sumx And sumy >= sumz Then maxdist = sumy * 2
        If sumz >= sumy And sumz >= sumx Then maxdist = sumz * 2
        
        If onkey = "" And (Not Camera.Planet Is Nothing) Then onkey = Camera.Planet.Key
        'move the planet we are on to the last to be rendered
        If (onkey <> "") And Planets(Planets.Count).Key <> onkey Then
            Set p = Planets(onkey)
            Planets.Remove onkey
            If Planets.Count > 0 Then
                Planets.Add p, onkey, 1
            Else
                Planets.Add p, onkey
            End If
        End If
        
        Debug.Print "Nearest: " & nearest.Key;
        Debug.Print " Aimingat: " & aimingAt.Key;
        Debug.Print " OnPlanet: " & onkey
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

        i = Planets.Count
        cnt = 0
        Do Until i = 0
            Set p = Planets(i)
            
        'For Each p In Planets

            If p.Visible Then
                If (p.Form = Plateau) Then
                    cnt = cnt + 1 'the number of total in plateau we are on

                    p.Scaled.z = 1
                    p.Scaled.Y = p.Scaled.z
                    p.Scaled.X = p.Scaled.z


                    
                    Dim matPlane As D3DMATRIX
                    Dim matRot As D3DMATRIX
                    Dim matPos As D3DMATRIX
                    Dim matScale As D3DMATRIX
                    Dim matYaw As D3DMATRIX
                    Dim matPitch As D3DMATRIX
                    Dim matRoll As D3DMATRIX
                    D3DXMatrixIdentity matPlane

                    D3DXMatrixIdentity matRot
                    D3DXMatrixIdentity matPos
                    D3DXMatrixIdentity matYaw
                    D3DXMatrixIdentity matPitch
                    D3DXMatrixIdentity matRoll
                    
                    If onkey <> p.Key Then
                       ' Rotation VectorAxisAngles(VectorDeduction(MoleculeView.Origin, p.Origin)), p
                        Set p.Rotate = VectorAxisAngles(VectorDeduction(MoleculeView.Origin, p.Origin))
                        

                    Else
                   '    Orientate MakePoint(0.01, 0.01, 0.01), p
                   '    Orientate MakePoint(0.01, 0, 0), p
                     '   Orientate MakePoint(0, 0.01, 0), p
                      ' Orientate MakePoint(0, 0, 0.01), p

                      
                    
                    End If



                    If onkey = p.Key Then

                     '   Orientate MakePoint(0.01, 0.01, 0.01), p
                      Orientate MakePoint(0.01, 0, 0), p
                '   Orientate MakePoint(0, 0.01, 0), p
                   '    Orientate MakePoint(0, 0, 0.01), p

                      
                    
                    End If
                    
                    
                    DDevice.SetRenderState D3DRS_ZENABLE, 0
                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
                    



                    D3DXMatrixRotationX matPitch, p.Rotate.X
                    D3DXMatrixMultiply matPlane, matPitch, matPlane

                    D3DXMatrixRotationY matYaw, p.Rotate.Y
                    D3DXMatrixMultiply matPlane, matYaw, matPlane
            
                    D3DXMatrixRotationZ matRoll, p.Rotate.z
                    D3DXMatrixMultiply matPlane, matRoll, matPlane

                    D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.z
                    D3DXMatrixMultiply matPlane, matScale, matPlane

                    D3DXMatrixTranslation matPos, p.Origin.X, p.Origin.Y, p.Origin.z
                    D3DXMatrixMultiply matPlane, matPlane, matPos




                    DDevice.SetTransform D3DTS_WORLD, matPlane
                    
                    
                    'render the round portion first
                    With p.Volume(1 + IIf(p.InnerRadius = 0, 2, 0))

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
                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, p.Volume.Count - IIf(p.InnerRadius = 0, 2, 0), VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
                    End With


                    'render the square portion second, if applies
                    If p.InnerRadius = 0 Then
                        DDevice.SetRenderState D3DRS_ZENABLE, 1
                        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

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
                            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))

                        End With
                    End If
                    
                End If
            End If
            i = i - 1
        Loop


        'reset lastorigin value to current for next time
        lastOrigin.X = MoleculeView.Origin.X
        lastOrigin.Y = MoleculeView.Origin.Y
        lastOrigin.z = MoleculeView.Origin.z
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW



    End If





    'DDevice.SetTransform D3DTS_VIEW, matSave
'    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
'        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), Near, Far
'    DDevice.SetTransform D3DTS_PROJECTION, matProj
    


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

 


'Public Sub RenderPlanets(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
'
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
'
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT
'
'    Dim matProj As D3DMATRIX
'    Dim matView As D3DMATRIX, matSave As D3DMATRIX
'
'    DDevice.GetTransform D3DTS_VIEW, matSave
'    matView = matSave
'    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
'    DDevice.SetTransform D3DTS_VIEW, matView
'
'    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
'        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), NEAR, FAR
'    DDevice.SetTransform D3DTS_PROJECTION, matProj
'    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'
'    DDevice.SetMaterial GenericMaterial
'    DDevice.SetTexture 1, Nothing
'    DDevice.SetMaterial LucentMaterial
'    DDevice.SetTexture 0, Nothing
'
'    Dim onKey As String
'    If Not Camera.Planet Is Nothing Then onKey = Camera.Planet.Key
'
'
'    Static worldRotate As Single
'    worldRotate = (worldRotate - 0.0001)
'    If worldRotate <= 0 Then worldRotate = (worldRotate + (PI * 2))
'
'
''    Dim cutoff As Single
'    Dim cnt As Long
'    Dim dist As Single
'    Dim dist2 As Single
'    Dim dist3 As Single
'
'    Dim v As Matter
'    Dim i As Long
'    Dim render As Boolean
'
'    Dim p As Planet
'    If Planets.Count > 0 Then
'        For Each p In Planets
'
'            If p.Visible Then
'                Select Case p.Form
'                    Case World
'
'                        If onKey <> "" Then
'                            If Planets(onKey).Origin.Equals(p.Origin) Then
'                                dist = Distance(p.Origin.X, 0, p.Origin.Z, 0, MoleculeView.Origin.Y, 0)
'                                If dist < Planets(onKey).OuterRadius Then
'                                    If dist > Planets(onKey).LevelLow Then
'                                        SetRenderBlends False, True
'
'                                        dist = 1 - (1 * ((MoleculeView.Origin.Y - Planets(onKey).LevelLow) / (Planets(onKey).LevelTop - Planets(onKey).LevelLow)))
'                                        render = True
'                                       ' DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(dist, Planets(onKey).Color.Red, Planets(onKey).Color.Green, Planets(onKey).Color.Blue)
'                                    Else
'                                        render = True
'                                        dist = 1
'                                        SetRenderBlends p.Transparent, p.Translucent
'                                    End If
'
'                                Else
'                                    render = False
'                                    dist = 0
'                                    SetRenderBlends p.Transparent, p.Translucent
'                                End If
'                            Else
'                                render = False
'                                dist = 0
'                                SetRenderBlends p.Transparent, p.Translucent
'                            End If
'                        Else
'                            render = False
'                            dist = 0
'                            SetRenderBlends p.Transparent, p.Translucent
'                        End If
'
'
'                        If render Then
'                            If (Not (p.Transparent Or p.Translucent)) Then
'
'
''                               GenericMaterial.Ambient.A = 1
''                               GenericMaterial.Ambient.r = 1
''                               GenericMaterial.Ambient.g = 1
''                               GenericMaterial.Ambient.b = 1
'
'                                DDevice.SetMaterial GenericMaterial
'                                DDevice.SetTexture 1, Nothing
'                                For i = 1 To p.Volume.Count Step 2
'                                    DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
'                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
'                               Next
''                               GenericMaterial.Ambient.A = 1
''                               GenericMaterial.Ambient.r = 1
''                               GenericMaterial.Ambient.g = 1
''                               GenericMaterial.Ambient.b = 1
''                               DDevice.SetMaterial GenericMaterial
'                            Else
'                                Debug.Print dist
'
'
'
'                                DDevice.SetMaterial LucentMaterial
'                                For i = 1 To p.Volume.Count Step 2
'                                    DDevice.SetTexture 0, Files(p.Volume(i).TextureIndex).Data
'                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
'                               Next
'
'                              DDevice.SetMaterial GenericMaterial
'                                For i = 1 To p.Volume.Count Step 2
'                                    DDevice.SetTexture 1, Files(p.Volume(i).TextureIndex).Data
'                                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX(p.Volume(i).TriangleIndex * 3), Len(VertexDirectX(0))
'                               Next
'
'
'
'                            End If
'                        End If
'
'                    Case Plateau
'
'
'
'
'                End Select
'            End If
'        Next
'    End If
'
'    DDevice.SetTransform D3DTS_VIEW, matSave
'    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'    D3DXMatrixPerspectiveFovLH matProj, FOVY, ((((CSng(RemoveArg(Resolution, "x")) / CSng(NextArg(Resolution, "x"))) + _
'        ((CSng(UserControl.Height) / VB.Screen.TwipsPerPixelY) / (CSng(UserControl.Width) / VB.Screen.TwipsPerPixelX))) / modGeometry.PI) * 2), NEAR, FAR
'    DDevice.SetTransform D3DTS_PROJECTION, matProj
'
'
'
'    DDevice.SetRenderState D3DRS_LIGHTING, 1
'    DDevice.SetRenderState D3DRS_FOGENABLE, 1
'
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetRenderState D3DRS_FOGCOLOR, DDevice.GetRenderState(D3DRS_AMBIENT)
'
'  ' DDevice.SetRenderState D3DRS_AMBIENT, Camera.Color
'    DDevice.SetRenderState D3DRS_ZENABLE, 1
'
'
'
''    Dim vin As D3DVECTOR
''    Dim vout As D3DVECTOR
'    Dim matPlane As D3DMATRIX
'    Dim matRot As D3DMATRIX
'    Dim matScale As D3DMATRIX
'    Dim matPos As D3DMATRIX
'    Dim matGlobe As D3DMATRIX
'    Dim matYaw As D3DMATRIX
'    Dim matPitch As D3DMATRIX
'    Dim matRoll As D3DMATRIX
'
'    Dim pX As Single
'    Dim pY As Single
'
'    Dim aim As Planet
'
'    Static lastOrigin As Point
'    If lastOrigin Is Nothing Then
'        Set lastOrigin = New Point
'        lastOrigin.X = MoleculeView.Origin.X
'        lastOrigin.Y = MoleculeView.Origin.Y
'        lastOrigin.Z = MoleculeView.Origin.Z
'    End If
'
'
'    Dim lowScale As Single
'    lowScale = 0.9
'    Const hiScale As Single = 1
'    Dim atScale As Single
'
'    Dim topLevel As Single
'    Dim midLevel As Single
'    Dim lowLevel As Single
'    Dim maxLevel As Single
'
'    Dim ordnal As Long
'
'    Dim sumx As Single
'    Dim sumz As Single
'    Dim sumy As Single
'
'
'    Dim sum As Single
'    Dim hi As Single
'    Dim lo As Single
'    Dim gap As Single
'    Dim pan As Single
'
'
'    Dim aimdiff As Point
'    Dim swirl As Point
'    Dim tracker As Point
'
'    Static down As Point
'
'    If down Is Nothing Then Set down = MakePoint((PI / 3), 0, 0)
'
'    lo = FAR
'    hi = 0
'
'    DDevice.SetTexture 1, Nothing
'
'    If Planets.Count > 0 Then
'
'        For Each p In Planets
'
'            If p.Visible Then
'                If (p.Form = Plateau) Then
'                    cnt = cnt + 1
'
'                    If hi < p.Width * p.Length Then hi = p.Width * p.Length
'                    If lo > p.Width * p.Length Then lo = p.Width * p.Length
'
'                    dist = DistanceEx(p.Origin, VectorDeduction(MoleculeView.Origin, VectorRotateAxis(MakePoint(0, 0, _
'                            DistanceEx(MoleculeView.Origin, p.Origin)), VectorNegative(MoleculeView.Rotate))))
'
'                    If dist < dist3 Or aim Is Nothing Then
'                        Set aim = p
'                        dist3 = dist
'                        Set aimdiff = VectorDeduction(p.Origin, MoleculeView.Origin)
'                    End If
'
''                    dist = DistanceEx(MoleculeView.Origin, p.Origin)
''                    If ((dist < dist2 Or dist2 = 0) And dist < (p.OuterRadius / 2)) And (Camera.Planet Is Nothing) Then
''                        dist2 = dist
''                        Set Camera.Planet = p
''                        onKey = p.Key
''                    End If
'
'                    sum = sum + p.OuterRadius
'                    sumx = sumx + Abs(p.Origin.X)
'                    sumy = sumy + Abs(p.Origin.Y)
'                    sumz = sumz + Abs(p.Origin.Z)
'
'                    If p.Key = onKey Then ordnal = cnt
'
'                End If
'            End If
'
'        Next
'        If sum = 0 Then sum = 1
'        If sumx = 0 Then sumx = 1
'        If sumy = 0 Then sumy = 1
'        If sumz = 0 Then sumz = 1
'
'        gap = hi - lo
'
'        DDevice.SetRenderState D3DRS_ZENABLE, 0
'        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'        If onKey = "" Then
'            If (MoleculeView.Origin.Z - lastOrigin.Z > 0) Or (MoleculeView.Origin.Y - lastOrigin.Y > 0) Or (MoleculeView.Origin.X - lastOrigin.X > 0) Then
'
'                    If Not aim Is Nothing Then
'
'                        If Distance(MoleculeView.Origin.X, 0, MoleculeView.Origin.Z, aim.Origin.X, 0, aim.Origin.Z) > aim.LevelMid Then
'
'                            MoleculeView.Origin.X = MoleculeView.Origin.X + aimdiff.X / 90
'                            MoleculeView.Origin.Z = MoleculeView.Origin.Z + aimdiff.Z / 90
'
'                        Else
'                            Set Camera.Planet = aim
'                            onKey = aim.Key
'
'                            MoleculeView.Origin.Y = (Planets(onKey).LevelTop - 0.001)
'
'                        End If
'
'                    End If
'
'            End If
'
'            pan = PI * 2
'
'        End If
'        pan = PI * 2
'
'        If onKey <> "" Then
'
'            lowLevel = (Planets(onKey).OuterRadius / 90)
'
'            dist = Distance(0, MoleculeView.Origin.Y, 0, 0, Planets(onKey).Origin.Y, 0)
'
'            maxLevel = (((Planets(onKey).OuterRadius / 90) * 2) + (Planets(onKey).OuterRadius / 10))
'            topLevel = ((Planets(onKey).OuterRadius / 90) + (Planets(onKey).OuterRadius / 10))
'            midLevel = (Planets(onKey).OuterRadius / 15)
'
'            If tracker Is Nothing Then Set tracker = MakePoint((PI + Planets(onKey).Rotate.Y), ((PI / 2) - Planets(onKey).Rotate.X), (PI + Planets(onKey).Rotate.Z))
'
'           ' pan = (PI * 2) - (PI / (MoleculeView.Origin.Y / topLevel))
'        End If
'
'        DDevice.SetRenderState D3DRS_ZENABLE, 0
'        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'        cnt = 0
'        For Each p In Planets
'
'            If p.Visible Then
'                If (p.Form = Plateau) Then
'                    If p.Key <> onKey Or (p.Key = onKey And MoleculeView.Origin.Y > ( p.OuterRadius/4)) Then
'
'                        cnt = cnt + 1
'
'                        If onKey = "" Then
'                            lowLevel = (p.OuterRadius / 90)
'                            maxLevel = (((p.OuterRadius / 90) * 2) + (p.OuterRadius / 10))
'                            topLevel = ((p.OuterRadius / 90) + (p.OuterRadius / 10))
'                            midLevel = (p.OuterRadius / 15)
'                        End If
'
'                        D3DXMatrixIdentity matPos
'                        D3DXMatrixIdentity matScale
'                        D3DXMatrixIdentity matRot
'                        D3DXMatrixIdentity matPlane
'                        D3DXMatrixIdentity matGlobe
'
'                        atScale = ((gap / sum) / (pan / (p.Length * p.Width)))
'
'                        If p.Key = onKey Then
'                            Set p.Offset = p.Origin
'                        Else
'                            Set p.Offset = DistanceSet(MoleculeView.Origin, p.Origin, topLevel)
'                        End If
'
'                        D3DXMatrixScaling matScale, atScale, atScale, atScale
'                        D3DXMatrixRotationYawPitchRoll matRot, (PI + MoleculeView.Rotate.Y), ((PI / 2) - MoleculeView.Rotate.X), (PI + MoleculeView.Rotate.Z)
'                        D3DXMatrixMultiply matRot, matScale, matRot
'                        D3DXMatrixTranslation matPos, p.Offset.X, p.Offset.Y, p.Offset.Z
'                        D3DXMatrixMultiply matPlane, matRot, matPos
'
'                        DDevice.SetTransform D3DTS_WORLD, matPlane
'
'                        With p.Volume(5)
'                            If Not (.Translucent Or .Transparent) Then
'                                DDevice.SetMaterial GenericMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                                DDevice.SetTexture 1, Nothing
'                            Else
'                                DDevice.SetMaterial LucentMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                                DDevice.SetMaterial GenericMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
'                            End If
'
'                            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 360, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
'                        End With
'
'                        D3DXMatrixTranslation matPos, -p.Offset.X, -p.Offset.Y, -p.Offset.Z
'                        D3DXMatrixMultiply matPlane, matPlane, matPos
'
'                    End If
'                End If
'            End If
'
'        Next
'
'
'        If onKey <> "" Then
'            DDevice.SetRenderState D3DRS_ZENABLE, 1
'            DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'
'            lowScale = ((gap / sum) / (pan / (Planets(onKey).Length * Planets(onKey).Width)))
'
'            If MoleculeView.Origin.Y < 4 Then MoleculeView.Origin.Y = 4
'
'            If dist > topLevel Then
'                atScale = lowScale
'
'                Set Camera.Planet = Nothing
'
'                MoleculeView.Origin.Y = (Planets(onKey).LevelTop + 0.001)
'
'            ElseIf dist > lowLevel Then
'
'                'atScale = (hiScale - lowScale) * (InvertNum(MoleculeView.Origin.Y, Planets(onKey).LevelTop) / Planets(onKey).LevelTop)
'                atScale = Round(hiScale - (((MoleculeView.Origin.Y - lowLevel) / (topLevel - lowLevel)) * (hiScale - lowScale)), 6)
''                If dist > midLevel Then
''
''                    If (MoleculeView.Origin.Y - lastOrigin.Y > 0) And (topLevel - dist) < lowLevel Then
''
''                        Set aimdiff = VectorDeduction(MakePoint((PI - ((PI / 4) * 2)), 0, 0), MoleculeView.Rotate)
''
''                        Set aimdiff = VectorMultiplyBy(aimdiff, (MoleculeView.Origin.Y - lastOrigin.Y) / (topLevel - dist))
''
''                      '  Set norm = VectorDivision(norm, 2)
''
''
''                    ElseIf MoleculeView.Origin.Equals(lastOrigin) Or (MoleculeView.Origin.Y - lastOrigin.Y > 0) Then
''
''                        Set aimdiff = VectorDeduction(MakePoint((PI - ((PI / 4) * 2)), 0, 0), MoleculeView.Rotate)
''
''                        If aimdiff.X > 0.0001 Then
''                            aimdiff.X = 0.0001
''                        ElseIf aimdiff.X < -0.0001 Then
''                            aimdiff.X = -0.0001
''                        End If
''
''                        If aimdiff.Y > 0.0001 Then
''                            aimdiff.Y = 0.0001
''                        ElseIf aimdiff.Y < -0.0001 Then
''                            aimdiff.Y = -0.0001
''                        End If
''
''                        If aimdiff.Z > 0.0001 Then
''                             aimdiff.Z = 0.0001
''                        ElseIf aimdiff.Z < -0.0001 Then
''                             aimdiff.Z = -0.0001
''                        End If
''                    Else
''                        Set aimdiff = MakePoint(0, 0, 0)
''                    End If
''
''                    If aimdiff.X <> 0 Then
''                        MoleculeView.Rotate.X = MoleculeView.Rotate.X + aimdiff.X
''                    End If
''
''                    If aimdiff.Y <> 0 Then
''                        MoleculeView.Rotate.Y = MoleculeView.Rotate.Y + aimdiff.Y
''                    End If
''
''                    If aimdiff.Z <> 0 Then
''                        MoleculeView.Rotate.Z = MoleculeView.Rotate.Z + aimdiff.Z
''                    End If
''
''                End If
'
'
'            Else
'                atScale = hiScale
'
'            End If
'
'            atScale = Round(atScale, 6)
'  '          If atScale > 0 Then
'
'            D3DXMatrixScaling matScale, atScale, atScale, atScale
''Else
''            px = (MoleculeView.Origin.X \ (Planets(onKey).Width / 3))
''            py = (MoleculeView.Origin.Z \ (Planets(onKey).Length / 3))
''            D3DXMatrixTranslation matPos, MoleculeView.Origin.X - (px * (Planets(onKey).Width / 3)), 0, MoleculeView.Origin.Z - (py * (Planets(onKey).Length / 3))
'
'            pX = (MoleculeView.Origin.X \ (Planets(onKey).Width / 3))
'            pY = (MoleculeView.Origin.Z \ (Planets(onKey).Length / 3))
'
'            If MoleculeView.Origin.Y > Planets(onKey).LevelLow Then
'
'            D3DXMatrixRotationYawPitchRoll matRot, (PI + MoleculeView.Rotate.Y), ((PI / 2) - MoleculeView.Rotate.X), (PI + MoleculeView.Rotate.Z)
'            D3DXMatrixTranslation matPos, MoleculeView.Origin.X - (pX * (Planets(onKey).Width / 3)), 0, MoleculeView.Origin.Z - (pY * (Planets(onKey).Length / 3))
'D3DXMatrixMultiply matPos, matPos, matScale
'            Else
'
'            D3DXMatrixTranslation matPos, MoleculeView.Origin.X - (pX * (Planets(onKey).Width / 3)), 0, MoleculeView.Origin.Z - (pY * (Planets(onKey).Length / 3))
'
'            End If
'
'
'
''End If
'
'            D3DXMatrixMultiply matPlane, matPos, matRot
'
'
'            DDevice.SetTransform D3DTS_WORLD, matPlane
'
'            If dist < lowLevel Then
'
'                'restrict to not below planet plane
'                If MoleculeView.Origin.Y < Planets(onKey).Origin.Y + 1 Then
'                    MoleculeView.Origin.Y = Planets(onKey).Origin.Y + 1
'                End If
'
'                With Planets(onKey).Volume(1)
'                    SetRenderBlends .Transparent, .Translucent
'                    If Not (.Translucent Or .Transparent) Then
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetTexture 1, Nothing
'                    Else
'                        DDevice.SetMaterial LucentMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
'                    End If
'
'                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
'
'                End With
'
'                DDevice.SetRenderState D3DRS_ZENABLE, 1
'                DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'            Else
'
'
'                With Planets(onKey).Volume(5)
'                    SetRenderBlends .Transparent, .Translucent
'                    If Not (.Translucent Or .Transparent) Then
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetTexture 1, Nothing
'                    Else
'                        DDevice.SetMaterial LucentMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
'                    End If
'
'                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 360, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
'                End With
'
'            End If
'
'            If dist < midLevel Then
'
'                With Planets(onKey).Volume(3)
'
'                    DDevice.SetRenderState D3DRS_ZENABLE, 0
'                    SetRenderBlends .Transparent, .Translucent
'                    If Not (.Translucent Or .Transparent) Then
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetTexture 1, Nothing
'                    Else
'                        DDevice.SetMaterial LucentMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                        DDevice.SetMaterial GenericMaterial
'                        If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
'                    End If
'
'                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
'                    DDevice.SetRenderState D3DRS_ZENABLE, 1
'
'                End With
'
'            End If
'
''            If MoleculeView.Origin.Y >= maxLevel Then
''
''                MoleculeView.Origin.Y = maxLevel
''
''            End If
'
'        End If
'
'        lastOrigin.X = MoleculeView.Origin.X 'set last altitude
'        lastOrigin.Y = MoleculeView.Origin.Y
'        lastOrigin.Z = MoleculeView.Origin.Z
'
'
'        DDevice.SetRenderState D3DRS_ZENABLE, 1
'        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'
'    End If
'
'    D3DXMatrixTranslation matPlane, 0, 0, 0
'    DDevice.SetTransform D3DTS_WORLD, matPlane
'
'    DDevice.SetRenderState D3DRS_CLIPPING, 1
'
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
'
'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'End Sub


 
' D3DXVECTOR3 toCam = camPos - spherePos;
' D3DXVECTOR3 fwdVector;
' D3DXVec3Normalize( &fwdVector, &toCam );
'
' D3DXVECTOR3 upVector( 0.0f, 1.0f, 0.0f );
' D3DXVECTOR3 sideVector;
' D3DXVec3CrossProduct( &sideVector, &upVector, &fwdVector );
' D3DXVec3CrossProduct( &upVector, &sideVector, &fwdVector );
'
' D3DXVec3Normalize( &upVector, &toCam );
' D3DXVec3Normalize( &sideVector, &toCam );
'
' D3DXMATRIX orientation( sideVector.x, sideVector.y, sideVector.z, 0.0f,
'                         upVector.x,   upVector.y,   upVector.z,   0.0f,
'                         fwdVector.x,  fwdVector.y,  fwdVector.z,  0.0f,
'                         spherePos.x,  spherePos.y,  spherePos.z,  1.0f );

                     '  Rotate.X = (PI + MoleculeView.Rotate.Y)
                    '   Rotate.Y = ((PI / 2) - MoleculeView.Rotate.X)
                    '   Rotate.Z = (PI + MoleculeView.Rotate.Z)
                        

                       ' Set Rotate = MakePoint(-MoleculeView.Rotate.X, -MoleculeView.Rotate.Y, 0)
   
'                        Set tmp = PlaneNormal(VectorRotateAxis(p.Volume(1).Point1, p.Rotate), _
'                                                VectorRotateAxis(p.Volume(1).Point2, p.Rotate), _
'                                                VectorRotateAxis(p.Volume(1).Point3, p.Rotate))
'                        Set tmp = VectorDeduction(tmp, VectorNormalize(VectorDeduction(MoleculeView.Origin, p.Origin)))
'                       ' Set tmp = VectorDeduction(tmp, MakePoint(p.Ranges.X, p.Ranges.Y, p.Ranges.Z))
'
'                        p.Ranges.Y = -tmp.X
'                        p.Ranges.X = -tmp.Y
'                        p.Ranges.Z = 0

                        
'                    Set tmp = PlaneNormal(VectorRotateAxis(p.Volume(1).Point1, p.Rotate), _
'                                            VectorRotateAxis(p.Volume(1).Point2, p.Rotate), _
'                                            VectorRotateAxis(p.Volume(1).Point3, p.Rotate))
'                    Set tmp = VectorDeduction(VectorNormalize(VectorDeduction(MoleculeView.Origin, p.Origin)), tmp)
'                    p.Ranges.Y = -tmp.X
'                    p.Ranges.X = -tmp.Y
'                    p.Ranges.Z = 0
                    

                        'Set p.Rotate = VectorAddition(p.Rotate, Change)
                        

                        'Rotation Rotate, p
                        
                        
'                        Rotate.Y = p.Rotate.Y
'
'                        Rotate.X = p.Rotate.X
'
'                        Rotate.Z = p.Rotate.Z
                        

                        'Rotate.X = Rotate.X + 0.01
                        'If Rotate.X > PI * 2 Then Rotate.X = -PI * 2
                        

                        
                        
'                        'dist = DistanceEx(Localized, p.Origin)
'                        Set Origin = DistanceSet(MoleculeView.Origin, p.Origin,  p.OuterRadius)
'
'                        Set Origin = MakePoint((p.Origin.X - (Sin(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))), 0, (p.Origin.Z - (Cos(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))))

                    '    Set Rotate = p.Rotate

'
'                        Set Rotate = MakePoint(MoleculeView.Rotate.Y, MoleculeView.Rotate.X, 0)
''
'                     Set tmp = VectorAxisAngles(MakePoint(-MoleculeView.Origin.Y, -MoleculeView.Origin.X, 0))
'
'                        p.Ranges.X = tmp.X
'                        p.Ranges.Y = tmp.Y
'                        p.Ranges.Z = tmp.Z
'
'                        Set Rotate = p.Rotate
                        
'                       ' Set Rotate = VectorNormalize(VectorDeduction(MoleculeView.Origin, p.Origin))
'                       ' Rotate.X = Rotate.Y - (PI / 2)
'
'                        'Set Rotate = VectorAddition(Rotate, MakePoint(-PI, -(PI - ((PI / 2) * 3)), (PI / 2)))
'
'
'

                       ' Debug.Print p.Distance
                            
    '            Set v.Point1 = VectorRotateAxis(v.Point1, MakePoint(PI, (PI - ((PI / 2) * 3)), -(PI / 2)))
    '            Set v.Point2 = VectorRotateAxis(v.Point2, MakePoint(PI, PI - ((PI / 2) * 3), -(PI / 2)))
    '            Set v.Point3 = VectorRotateAxis(v.Point3, MakePoint(PI, PI - ((PI / 2) * 3), -(PI / 2)))
'                        Set Origin = MkaePoint((pX * (p.InnerRadius / 3)), 0, (pY * (p.InnerRadius / 3)))
'                        Set Rotate = p.Rotate
'                        Set Scaled = MakePoint(1, 1, 1)
'                        D3DXMatrixScaling matScale, Scaled.X, Scaled.Y, Scaled.Z
'                        D3DXMatrixMultiply matPlane, matPlane, matScale
'                        D3DXMatrixTranslation matPos, Origin.X, Origin.Y, Origin.Z
'                        D3DXMatrixMultiply matPlane, matPlane, matPos
'                        D3DXMatrixRotationYawPitchRoll matRot, Rotate.X, Rotate.Y, Rotate.Z
'                        D3DXMatrixMultiply matPlane, matPlane, matRot


                        
                     '   Debug.Print tmp
                        
                        'Debug.Print tmp
                      '  p.Ranges.X = 0
                      '  p.Ranges.Y = (PI / 2)
                      '  p.Ranges.Z = 0
                        


                       ' Set p.Rotate = VectorAddition(p.Rotate, Change)
                       ' Set Rotate = p.Rotate

                        'Rotation Rotate, p
                        
                        
                        'Set Rotate = MakePoint(-PI, -(PI - ((PI / 2) * 3)), (PI / 2))

'
'                        'Set Origin = MakePoint((p.Origin.X - (Sin(D720 - MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))), p.Origin.Y, (p.Origin.Z - (Cos(D720 - MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))))
'
'                      '  Set Origin = MakePoint((p.Origin.X - (Sin(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))), 0, (p.Origin.Z - (Cos(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))))

'Set Origin = MakePoint((p.Origin.X - (Sin(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))), 0, (p.Origin.Z - (Cos(D720 - -MoleculeView.Rotate.X) * ((p.InnerRadius * 2) / (PI * 2.5)))))

                        'Set norm = PlaneNormal(p.Volume(1).Point1, p.Volume(1).Point2, p.Volume(1).Point3)
''
'                        Set Rotate = New Point
                      ' Set tmp = VectorNormalize(VectorDeduction(MoleculeView.Origin, p.Origin))
'
'                        Rotate.X = tmp.X
'
''
'
                       ' Set tmp = VectorNormalize(VectorDeduction(tmp, norm))

                      '  p.Ranges.X = tmp.X
                      '  p.Ranges.Y = tmp.Y
                    '    p.Ranges.Z = tmp.Z
                        

                        'Rotate.X = (AngleOfCoord(MakeCoord(tmp.Y, tmp.Z)) - AngleOfCoord(MakeCoord(tmp.X, tmp.Y))) - AngleOfCoord(MakeCoord(tmp.X, tmp.Z))

                        'Rotate.Y = AngleOfCoord(MakeCoord(tmp.Y, tmp.Z)) + -(AngleOfCoord(MakeCoord(tmp.X, tmp.Z)) + AngleOfCoord(MakeCoord(tmp.X, tmp.Z))) + AngleOfCoord(MakeCoord(tmp.Y, tmp.Z))

                     'Rotate.Z = -(-AngleOfCoord(MakeCoord(tmp.X, tmp.Y)) + AngleOfCoord(MakeCoord(tmp.Y, tmp.Z))) - AngleOfCoord(MakeCoord(tmp.X, tmp.Z))

'
'                        Set Rotate = tmp
'                        Set Rotate = VectorNormalize(MakePoint(Abs(tmp.Z) - PI, Abs(tmp.Y), Abs(tmp.X) + PI))
'
'
'
'                        Rotate.Z = AngleOfCoord(MakePoint(tmp.Y, tmp.X, 0))
'
'                        Set Rotate = MakePoint(-MoleculeView.Origin.X, -MoleculeView.Origin.Y, 0)
'
'Debug.Print Origin.X; Origin.Y; Origin.Z




'                        D3DXMatrixRotationYawPitchRoll matBeacon, IIf(Beacons(o).HorizontalLock, 0, -Player.CameraAngle), -Player.CameraPitch, 0
'
'                        D3DXMatrixScaling matScale, 1, 1, 1
'                        D3DXMatrixMultiply matBeacon, matBeacon, matScale
'
'                        D3DXMatrixTranslation matPos, (Beacons(o).Origins(l).X - (Sin(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.Height / (PI * 2.5)))), Beacons(o).Origins(l).Y, (Beacons(o).Origins(l).Z - (Cos(D720 - IIf(Beacons(o).HorizontalLock, 0, Player.CameraAngle)) * (Beacons(o).Dimension.Height / (PI * 2.5))))
'                        D3DXMatrixMultiply matBeacon, matBeacon, matPos
'



''                    Set norm = PlaneNormal(p.Volume(3).Point1, p.Volume(3).Point2, p.Volume(3).Point3)
''                    Set norm2 = PlaneNormal(MoleculeView.Origin, p.Volume(3).Point2, p.Volume(3).Point3)
''
''
''                    Set tmp = VectorCrossProduct(norm, norm2)
'
'                    'Set tmp = VectorNormalize(VectorDeduction(MoleculeView.Origin, p.Origin))
'
'
'                        Set tmp = VectorAxisAngles(VectorDeduction(MoleculeView.Origin, p.Origin))
'
'                       'tmp.X = -MoleculeView.Rotate.Y - ((PI * 2) / 4)
'
'
'                    Rotation tmp, p

                       ' Debug.Print p.Distance


'            If Not aim Is Nothing Then
'                'auto center on aim planet
'                If (MoleculeView.Origin.Z - lastOrigin.Z > 0) Or MoleculeView.Origin.Equals(lastOrigin) Then
'
'                    If (MoleculeView.Origin.Z - lastOrigin.Z) > 0 Then
'                        dist3 = 0.001
'                    Else
'                        dist3 = 0.0001
'                    End If
'
'
'                    If norm.X > dist3 Then
'                        MoleculeView.Rotate.X = MoleculeView.Rotate.X + dist3
'                    ElseIf norm.X < -dist3 Then
'                        MoleculeView.Rotate.X = MoleculeView.Rotate.X + -dist3
'                    End If
'
'                    If norm.Y > dist3 Then
'                        MoleculeView.Rotate.Y = MoleculeView.Rotate.Y + dist3
'                    ElseIf norm.Y < -dist3 Then
'                        MoleculeView.Rotate.Y = MoleculeView.Rotate.Y + -dist3
'                    End If
'
'                    If norm.Z > dist3 Then
'                        MoleculeView.Rotate.Z = MoleculeView.Rotate.Z + dist3
'                    ElseIf norm.Z < -dist3 Then
'                        MoleculeView.Rotate.Z = MoleculeView.Rotate.Z + -dist3
'                    End If
'
'
'                End If
'
'            End If

'            cnt = 0
'            For Each p In Planets
'
'                If p.Visible Then
'                    If (p.Form = Plateau) Then
'                        cnt = cnt + 1
'
'                        D3DXMatrixIdentity matPos
'                        D3DXMatrixIdentity matScale
'                        D3DXMatrixIdentity matRot
'                        D3DXMatrixIdentity matPlane
'                        D3DXMatrixIdentity matGlobe
'
'
''                        atScale = Round((((((360 / Planets.Count) * cnt) * ((sum - p.OuterRadius) / sum)) / topLevel)) * lowScale * (p.Height / p.OuterRadius), 6)
''
''                        Set swirl = MakePoint((PI + MoleculeView.Rotate.Y), ((PI / 2) - MoleculeView.Rotate.X), (PI + MoleculeView.Rotate.Z))
''                        Set norm = VectorAddition(MoleculeView.Origin, VectorRotateAxis(MakePoint(0, 0, topLevel), MoleculeView.Rotate))
''
''                        D3DXMatrixScaling matScale, atScale, atScale, atScale
''                        D3DXMatrixRotationYawPitchRoll matRot, swirl.X, swirl.Y, swirl.Z
''                        D3DXMatrixMultiply matRot, matScale, matRot
''                        D3DXMatrixTranslation matPos, norm.X, norm.Y, norm.Z
''                        D3DXMatrixMultiply matRot, matRot, matPos
''
''                        Set globe = VectorDeduction(globe, MoleculeView.Rotate)
''                        Set MoleculeView.Rotate = MakePoint(0, 0, 0)
''
''                        D3DXMatrixRotationYawPitchRoll matGlobe, globe.Y, (globe.X - ((PI / Planets.Count) * cnt)), globe.Z
''                        D3DXMatrixMultiply matPlane, matRot, matGlobe
'
'                        DDevice.SetTransform D3DTS_WORLD, matPlane
'                        SetRenderBlends p.Transparent, p.Translucent
'
'                        With p.Volume(5)
'                            If Not (.Translucent Or .Transparent) Then
'                                DDevice.SetMaterial GenericMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                                DDevice.SetTexture 1, Nothing
'                            Else
'                                DDevice.SetMaterial LucentMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 0, Files(.TextureIndex).Data
'                                DDevice.SetMaterial GenericMaterial
'                                If .TextureIndex > 0 Then DDevice.SetTexture 1, Files(.TextureIndex).Data
'                            End If
'
'                            DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 360, VertexDirectX((.TriangleIndex * 3)), Len(VertexDirectX(0))
'                        End With
'
'                    End If
'                End If
'
'            Next
'
'        Else
'
'            lowLevel = (Planets(onKey).OuterRadius / 90)
'
'            dist = Distance(0, MoleculeView.Origin.Y, 0, 0, Planets(onKey).Origin.Y, 0)
'
'            maxLevel = (((Planets(onKey).OuterRadius / 90) * 2) + (Planets(onKey).OuterRadius / 10))
'            topLevel = ((Planets(onKey).OuterRadius / 90) + (Planets(onKey).OuterRadius / 10))
'            midLevel = (Planets(onKey).OuterRadius / 15)
'
'            If swirl Is Nothing Then Set swirl = MakePoint((PI + Planets(onKey).Rotate.Y), ((PI / 2) - Planets(onKey).Rotate.X), (PI + Planets(onKey).Rotate.Z))
'        End If
     
'                        If onKey = "" Then
'
'
'                            maxLevel = (((p.OuterRadius / 90) * 2) + (p.OuterRadius / 10))
'                            topLevel = ((p.OuterRadius / 90) + (p.OuterRadius / 10))
'                            midLevel = (p.OuterRadius / 15)
'                            lowLevel = (p.OuterRadius / 90)
'                        End If
'
'                        atScale = Round((((((360 / Planets.Count) * cnt) * ((sum - p.OuterRadius) / sum)) / topLevel)) * lowScale * p.Length, 6)
'
'                        Set p.Scaled = MakePoint(atScale, atScale, atScale)
'                        atScale = (p.OuterRadius / (p.OuterRadius * Planets.Count))
''                        atScale = ((gap / sum) / (pan / (p.Length * p.Width)))
''
'                        If p.Key = onKey Then
'                            Set p.Offset = p.Origin
'                        Else
'                            Set p.Offset = DistanceSet(MoleculeView.Origin, p.Origin, topLevel)
'                        End If
'
'
'                      '  Set p.Offset = VectorRotateAxis(DistanceSet(MakePoint(0, 0, 0), VectorNormalize(p.Origin), -topLevel), VectorNegative(VectorMultiplyBy(p.Origin, atScale)))
'                        'Set p.Offset = VectorAddition(MoleculeView.Origin, VectorRotateAxis(MakePoint(0, 0, -topLevel), MakePoint(0, PI, 0)))
'
'                        Set p.Rotate = VectorNormalize(VectorNegative(VectorDeduction(p.Origin, MoleculeView.Origin)))
'                       ' p.Rotate.X = -MoleculeView.Rotate.Y - ((PI * 2) / 4)
'
'                        If onKey = "" Then
'
'                          '  Set p.Scaled = MakePoint(atScale, atScale, atScale)
'
'                          '  Set p.Offset = VectorAddition(p.Origin, VectorRotateAxis(MakePoint(0, 0, topLevel), MakePoint(0, PI, 0)))
'                          '  Set p.Rotate = VectorNormalize(VectorNegative(VectorDeduction(p.Origin, MoleculeView.Origin)))
'
'                            D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.Z
'
'                            D3DXMatrixTranslation matPos, p.Offset.X, p.Offset.Y, p.Offset.Z
'                            D3DXMatrixMultiply matPos, matScale, matPos
'
'                            D3DXMatrixRotationYawPitchRoll matRot, (PI + MoleculeView.Rotate.Y), ((PI / 2) - MoleculeView.Rotate.X), (PI + MoleculeView.Rotate.Z)
'                            'D3DXMatrixRotationYawPitchRoll matRot, p.Rotate.Y, p.Rotate.X, p.Rotate.Z
'                            D3DXMatrixMultiply matPlane, matScale, matRot
'                        Else
'
'
'                            D3DXMatrixScaling matScale, p.Scaled.X, p.Scaled.Y, p.Scaled.Z
'
'                            D3DXMatrixRotationYawPitchRoll matRot, (PI + MoleculeView.Rotate.Y), ((PI / 2) - MoleculeView.Rotate.X), (PI + MoleculeView.Rotate.Z)
'                            'D3DXMatrixRotationYawPitchRoll matRot, p.Rotate.Y, p.Rotate.X, p.Rotate.Z
'                            D3DXMatrixMultiply matRot, matScale, matRot
'
'                            D3DXMatrixTranslation matPos, p.Offset.X, p.Offset.Y, p.Offset.Z
'                            D3DXMatrixMultiply matPlane, matRot, matPos
'
'                        End If

'            lowScale = Round((((((360 / Planets.Count) * ordnal) * ((sum - Planets(onKey).OuterRadius) / sum)) / topLevel)) * lowScale * (Planets(onKey).Height / Planets(onKey).OuterRadius), 6)
'
'
'            If dist >= topLevel Then
'                atScale = lowScale
'                Set Camera.Planet = Nothing
'
''                MoleculeView.Origin.Y = ((Planets(onKey).OuterRadius / 90) + (Planets(onKey).OuterRadius / 10)) + 1
''
''
''                Set globe = VectorDeduction(MoleculeView.Rotate, VectorDeduction(MakePoint(-(PI / 3) - ((PI / Planets.Count) * ordnal), 0, 0), VectorNegative(globe)))
''
''                Set globe = VectorDeduction(globe, MoleculeView.Rotate)
''
''                MoleculeView.Rotate.X = 0
''                MoleculeView.Rotate.Y = 0
''                MoleculeView.Rotate.Z = 0
'
'
'            ElseIf dist > lowLevel Then
'
'                atScale = Round(hiScale - (((MoleculeView.Origin.Y - lowLevel) / (topLevel - lowLevel)) * (hiScale - lowScale)), 6)
'
'
'                If dist > midLevel Then
'
'                    If (MoleculeView.Origin.Y - lastOrigin.Y > 0) And (topLevel - dist) < lowLevel Then
'
'                        Set norm = VectorDeduction(MakePoint((PI - ((PI / 4) * 2)), 0, 0), MoleculeView.Rotate)
'
'                        Set norm = VectorMultiplyBy(norm, (MoleculeView.Origin.Y - lastOrigin.Y) / (topLevel - dist))
'
'                      '  Set norm = VectorDivision(norm, 2)
'
'
'                    ElseIf MoleculeView.Origin.Equals(lastOrigin) Or (MoleculeView.Origin.Y - lastOrigin.Y > 0) Then
'
'                        Set norm = VectorDeduction(MakePoint((PI - ((PI / 4) * 2)), 0, 0), MoleculeView.Rotate)
'
'                        If norm.X > 0.0001 Then
'                            norm.X = 0.0001
'                        ElseIf norm.X < -0.0001 Then
'                            norm.X = -0.0001
'                        End If
'
'                        If norm.Y > 0.0001 Then
'                            norm.Y = 0.0001
'                        ElseIf norm.Y < -0.0001 Then
'                            norm.Y = -0.0001
'                        End If
'
'                        If norm.Z > 0.0001 Then
'                             norm.Z = 0.0001
'                        ElseIf norm.Z < -0.0001 Then
'                             norm.Z = -0.0001
'                        End If
'                    Else
'                        Set norm = MakePoint(0, 0, 0)
'                    End If
'
'                    If norm.X <> 0 Then
'                        MoleculeView.Rotate.X = MoleculeView.Rotate.X + norm.X
'                    End If
'
'                    If norm.Y <> 0 Then
'                        MoleculeView.Rotate.Y = MoleculeView.Rotate.Y + norm.Y
'                    End If
'
'                    If norm.Z <> 0 Then
'                        MoleculeView.Rotate.Z = MoleculeView.Rotate.Z + norm.Z
'                    End If
'
'                End If
'
'
'            Else
'                atScale = hiScale
'
'            End If


'Public Sub OrientateMolecule(ByRef matMolecule As D3DMATRIX, ByRef m As Molecule, ByVal Planet As Boolean)
'    Dim matMesh As D3DMATRIX
'
'    D3DXMatrixIdentity matMesh
'    If Planet Then
'
'        If Not m.Relative Then
'            D3DXMatrixTranslation matMolecule, m.Origin.X - m.Offset.X, m.Origin.Y - m.Offset.Y, m.Origin.Z - m.Offset.Z
'            D3DXMatrixMultiply matMesh, matMolecule, matMesh
'            D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        End If
'
'        D3DXMatrixRotationX matMesh, ((m.Rotate.X) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        D3DXMatrixRotationY matMesh, ((m.Rotate.Y) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        D3DXMatrixRotationZ matMesh, ((m.Rotate.Z) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'
'        If m.Relative Then
'            D3DXMatrixTranslation matMesh, m.Origin.X + (m.Offset.X * 2), m.Origin.Y + (m.Offset.Y * 2), m.Origin.Z + (m.Offset.Z * 2)
'            D3DXMatrixMultiply matMesh, matMolecule, matMesh
'            D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        End If
'
'        DDevice.SetTransform D3DTS_WORLD, matMolecule
'
'    Else
'
'        If Not m.Relative Then
'            D3DXMatrixTranslation matMesh, m.Origin.X + m.Offset.X, m.Origin.Y + m.Offset.Y, m.Origin.Z + m.Offset.Z
'            D3DXMatrixMultiply matMesh, matMolecule, matMesh
'            D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        End If
'
'        D3DXMatrixScaling matMolecule, m.Scaled.X, m.Scaled.Y, m.Scaled.Z
'
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'
'        D3DXMatrixRotationX matMesh, ((m.Rotate.X) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        D3DXMatrixRotationY matMesh, ((m.Rotate.Y) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        D3DXMatrixRotationZ matMesh, ((m.Rotate.Z) * RADIAN)
'        D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'
'        D3DXMatrixMultiply matMesh, matMolecule, matMesh
'
'        If m.Relative Then
'            D3DXMatrixTranslation matMolecule, m.Origin.X + m.Offset.X, m.Origin.Y + m.Offset.Y, m.Origin.Z + m.Offset.Z
'            D3DXMatrixMultiply matMesh, matMolecule, matMesh
'            D3DXMatrixMultiply matMolecule, matMesh, matMolecule
'        End If
'
'        DDevice.SetTransform D3DTS_WORLD, matMesh
'
'        D3DXMatrixRotationX matMolecule, -((m.Rotate.X) * RADIAN)
'        D3DXMatrixMultiply matMesh, matMolecule, matMesh
'        D3DXMatrixRotationY matMolecule, -((m.Rotate.Y) * RADIAN)
'        D3DXMatrixMultiply matMesh, matMolecule, matMesh
'        D3DXMatrixRotationZ matMolecule, -((m.Rotate.Z) * RADIAN)
'        D3DXMatrixMultiply matMesh, matMolecule, matMesh
'
'    End If
'End Sub
'Private Sub CouplingMolecule(ByRef m As Molecule, ByRef m3 As Molecule)
'    If ((m.Collision And Coupling) = Coupling) And ((m3.Collision And Coupling) = Coupling) Then
'        If (m.Motions.Count > 0) Then
'            Dim cnt2 As Long
'            For cnt2 = 1 To m.Motions.Count
'                Select Case m.Motions.Item(cnt2)
'                    Case "Gravity"
'                    Case "Liquid"
'                    Case Else
'                        Select Case m.Motions.Item(cnt2).Action
'                            Case Direct
'                                If Not m3.Motions.Exists(m.Motions.key(cnt2)) Then
'                                    m3.Motions.Add m.Motions.Item(cnt2).Clone, m.Motions.key(cnt2)
'                                End If
'                            Case Rotate
'                                If Not m3.Motions.Exists(m.Motions.key(cnt2)) Then
'                                    m3.Motions.Add m.Motions.Item(cnt2).Clone, m.Motions.key(cnt2)
'                                    m3.Motions.Item(m.Motions.key(cnt2)).Axis.Invert
'                                End If
'                            Case Scalar
'                        End Select
'                End Select
'            Next
'        End If
'
'    End If
'End Sub
'Private Function CollideTestMolecule(ByRef m As Molecule, ByRef m3 As Molecule) As Boolean
'On Error GoTo ObjectError
'
'    If ((lngFaceCount > 0) And (m.CollideIndex > -1) And (m3.CollideIndex > -1)) Then
'
'        Dim visType As Long
'        visType = 2
'
'        Dim cnt As Long
'
'        For cnt = 0 To lngFaceCount - 1
'            sngFaceVis(3, cnt) = 0
'        Next
'
'        For cnt = m3.CollideIndex To (m3.CollideIndex + m3.CollideFaces) - 1
'            sngFaceVis(3, cnt) = visType
'        Next
'
'        '#####################################################################################
'        '############# in rotation collision we re-adjsut culling view direction #############
'        '#####################################################################################
'
''        sngCamera(0, 0) = m.Origin.X
''        sngCamera(0, 1) = m.Origin.Y + 1
''        sngCamera(0, 2) = m.Origin.Z
''
''        sngCamera(1, 0) = 1
''        sngCamera(1, 1) = -1
''        sngCamera(1, 2) = -1
''
''        sngCamera(2, 0) = -1
''        sngCamera(2, 1) = 1
''        sngCamera(2, 2) = -1
''
''        m.CulledFaces = Culling(visType, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer)
''Debug.Print m.CulledFaces
'
'        '#####################################################################################
'        '############# create a transform matrix with the changes applied ####################
'        '#####################################################################################
''        For cnt = m.CollideIndex To (m.CollideIndex + m.CollideFaces) - 1
''            sngFaceVis(3, cnt) = 0
''        Next
'
'        Dim cnt2 As Long
'        Dim Face As Long
'        Dim Index As Long
'        Dim V(2) As D3DVECTOR
'
'        Dim matMesh As D3DMATRIX
'        DDevice.GetTransform D3DTS_WORLD, matMesh
'
'        If ((m.FaceIndex = -1) And (Not m.MeshIndex = -1)) Or (m.FaceIndex > 0) Then
'
'            If (m.FaceIndex > 0) Then
'                OrientateMolecule matMesh, m, True
'            End If
'
'            '#####################################################################################
'            '############# update face data with the transformation matrix #######################
'            '#####################################################################################
'
'            For Face = m.CollideIndex To (m.CollideIndex + m.CollideFaces) - 1
'
'                For cnt = 0 To 2
'
'                    V(cnt).X = Meshes(m.MeshIndex).Verticies(Index + cnt).X
'                    V(cnt).Y = Meshes(m.MeshIndex).Verticies(Index + cnt).Y
'                    V(cnt).Z = Meshes(m.MeshIndex).Verticies(Index + cnt).Z
'
'                    D3DXVec3TransformCoord V(cnt), V(cnt), matMesh
'
'                    sngVertexX(cnt, Face) = V(cnt).X
'                    sngVertexY(cnt, Face) = V(cnt).Y
'                    sngVertexZ(cnt, Face) = V(cnt).Z
'
'                Next
'
'                Index = Index + 3
'            Next
'
'        Else
'            Exit Function
'        End If
'
'        '#####################################################################################
'        '############# per non culled face check and result collision ########################
'        '#####################################################################################
'
'        Dim lngCollideIdx As Long
'        lngCollideIdx = -1
'
'        Dim objCollision As Long
'        objCollision = -1
'
'        For cnt = m.CollideIndex To (m.CollideIndex + m.CollideFaces) - 1
'
'            If CBool(Collision(visType, lngFaceCount, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, cnt, objCollision, lngCollideIdx)) Then
'
'                CouplingMolecule m, m3
'
'                If Not m.OnCollide Is Nothing Then
'                    CollideTestMolecule = m.OnCollide.RunEvent
'                Else
'                    CollideTestMolecule = True
'                End If
'
'                GoTo exitfunction
'            End If
'        Next
'
'
'
'    End If
'
'exitfunction:
'
'    Exit Function
'ObjectError:
'    If Err.Number = 6 Or Err.Number = 11 Then Resume
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'
'End Function
'
'Private Function RangeTestMolecule(ByRef m As Molecule, ByRef m3 As Molecule) As Boolean
''ByRef m As Molecule, ByRef m3 As Molecule, ByVal InorOut As Boolean) As Boolean
'    Dim matTest1 As D3DMATRIX
'    Dim matTest2 As D3DMATRIX
'    Dim orgin1 As D3DVECTOR
'    Dim orgin2 As D3DVECTOR
'
'    D3DXMatrixIdentity matTest1
'    D3DXMatrixIdentity matTest2
'
'    If Not m3 Is Nothing Then
'
'        orgin1.X = ((m.Origin.X - m.Offset.X) * Sin(m.Rotate.X)) + (m.Offset.X * Sin(m.Rotate.X))
'        orgin1.Y = ((m.Origin.Y - m.Offset.Y) * Tan(m.Rotate.Y)) + (m.Offset.Y * Tan(m.Rotate.Y))
'        orgin1.Z = ((m.Origin.Z - m.Offset.Z) * Cos(m.Rotate.Z)) + (m.Offset.Z * Cos(m.Rotate.Z))
'
'        orgin2.X = ((m3.Origin.X - m3.Offset.X) * Tan(m3.Rotate.X)) + (m3.Offset.X * Sin(m3.Rotate.X))
'        orgin2.Y = ((m3.Origin.Y - m3.Offset.Y) * Cos(m3.Rotate.Y)) + (m3.Offset.Y * Tan(m3.Rotate.Y))
'        orgin2.Z = ((m3.Origin.Z - m3.Offset.Z) * Sin(m3.Rotate.Z)) + (m3.Offset.Z * Cos(m3.Rotate.Z))
'
'        If (DistanceEx(orgin1, orgin2) <= m.Range + m3.Range) Then
'
'            CouplingMolecule m, m3
'
'            If Not m.OnInRange Is Nothing Then RangeTestMolecule = m.OnInRange.RunEvent
'
'        Else
'            If Not m.OnOutRange Is Nothing Then RangeTestMolecule = m.OnOutRange.RunEvent
'
'        End If
'    End If
'End Function
'
'Private Function MoleculeIterateEvent(ByRef Col As NTNodes10.Collection, ByRef m4 As Variant, ByRef m As Molecule) As Boolean
'    Dim m2 As Variant
'    Dim m3 As Molecule
'    Dim p As Planet
'    For m2 = 1 To Col.Count
'        If TypeName(Col.Item(m2)) = "Molecule" Then
'            If Col.key(m2) <> Molecules.key(m4) Then
'                Set m3 = Col.Item(m2)
'            End If
'        ElseIf TypeName(Col.Item(m2)) = "Planet" Then
'            Set p = Col.Item(m2)
'            If p.MoleculeKey <> Molecules.key(m4) Then
'                Set m3 = p.Molecule
'            End If
'        ElseIf TypeName(Col.Item(m2)) = "String" Then
'            If Col.Item(m2) <> Molecules.key(m4) Then
'                Set m3 = Molecules(Col.Item(m2))
'            End If
'        End If
'        If Not m3 Is Nothing Then
'            If ((m.Collision And Ranged) = Ranged) Then
'                MoleculeIterateEvent = RangeTestMolecule(m, m3)
'            End If
'            If (m.Collision > Ranged) Then
'                MoleculeIterateEvent = CollideTestMolecule(m, m3)
'            End If
'            Set m3 = Nothing
'        End If
'        If Not p Is Nothing Then Set p = Nothing
'        If MoleculeIterateEvent Then Exit Function
'    Next
'End Function
'
'Private Function TestCollision(ByRef m4 As Variant, ByRef m As Molecule) As Boolean
'    If (CInt(m.Collision) > 0) Then
'        'Through at the least
'        'this function should used cached couple collision info if it had occurd already in a iteration
'
'        If (Not (m.OnCollide Is Nothing)) Then
'            If Not m.OnCollide.ApplyTo Is Nothing Then
'                If m.OnCollide.ApplyTo.Count = 0 Then
'                    TestCollision = MoleculeIterateEvent(All, m4, m)
'                Else
'                    TestCollision = MoleculeIterateEvent(m.OnCollide.ApplyTo, m4, m)
'                End If
'            Else
'                TestCollision = MoleculeIterateEvent(All, m4, m)
'            End If
'        Else
'            TestCollision = MoleculeIterateEvent(All, m4, m)
'        End If
'    End If
'
'End Function


Public Sub RenderObject(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)

'    DDevice.SetRenderState D3DRS_ZENABLE, 1
'
'    DDevice.SetRenderState D3DRS_CLIPPING, 1
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'
'    DDevice.SetVertexShader FVF_RENDER
'    DDevice.SetPixelShader PixelShaderDefault
'
''    D3DXMatrixIdentity matWorld
''    D3DXMatrixRotationX matWorld, 0
''    D3DXMatrixRotationY matWorld, 0
''    D3DXMatrixRotationZ matWorld, 0
''
''    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
'
'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
'    DDevice.SetTextureStageState 1, D3DTSS_MAXANISOTROPY, 16
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
'
'    DDevice.SetMaterial LucentMaterial
'    DDevice.SetTexture 0, Nothing
'    DDevice.SetMaterial GenericMaterial
'    DDevice.SetTexture 1, Nothing
'
'    Dim B As Brilliant
'    For Each B In Brilliants
'
'        If Lights(B.LightIndex).Type = D3DLIGHT_DIRECTIONAL Or (B.Enabled And _
'            Distance(MoleculeView.Origin.X, MoleculeView.Origin.Y, MoleculeView.Origin.Z, B.Origin.X, B.Origin.Y, B.Origin.Z) <= (FAR - Lights(B.LightIndex).Range)) Then
'
'            If (B.LightBlink > 0) Or (B.DiffuseRoll <> 0) Then
'                If (B.LightBlink > 0) Then
'                    If (B.LightTimer = 0) Or ((Timer - B.LightTimer) >= B.LightBlink And (B.LightBlink > 0)) Then
'                        B.LightTimer = Timer
'                        B.LightIsOn = Not B.LightIsOn
'                    End If
'                    DDevice.LightEnable (B.LightIndex - 1), B.LightIsOn
'                End If
'                If (B.DiffuseRoll <> 0) Then
'                    If (B.DiffuseTimer = 0) Or ((Timer - B.DiffuseTimer) >= Abs(B.DiffuseRoll) And (B.DiffuseTimer > 0)) Then
'                        B.DiffuseTimer = Timer
'                        If (B.DiffuseRoll > 0) Then
'                            If (B.DIffuseMax > 0 And B.DiffuseNow >= B.DIffuseMax) Or (B.DIffuseMax < 0 And B.DiffuseNow >= -0.01) Then
'                                B.DiffuseRoll = -B.DiffuseRoll
'                            Else
'                                B.DiffuseNow = B.DiffuseNow + 1
'                                Lights(B.LightIndex).Diffuse.r = Lights(B.LightIndex).Diffuse.r + 0.01
'                                Lights(B.LightIndex).Diffuse.g = Lights(B.LightIndex).Diffuse.g + 0.01
'                                Lights(B.LightIndex).Diffuse.B = Lights(B.LightIndex).Diffuse.B + 0.01
'                            End If
'                        Else
'                            If (B.DIffuseMax > 0 And B.DiffuseNow <= 0.01) Or (B.DIffuseMax < 0 And B.DiffuseNow <= B.DIffuseMax) Then
'                                B.DiffuseRoll = -B.DiffuseRoll
'                            Else
'                                B.DiffuseNow = B.DiffuseNow - 1
'                                Lights(B.LightIndex).Diffuse.r = Lights(B.LightIndex).Diffuse.r - 0.01
'                                Lights(B.LightIndex).Diffuse.g = Lights(B.LightIndex).Diffuse.g - 0.01
'                                Lights(B.LightIndex).Diffuse.B = Lights(B.LightIndex).Diffuse.B - 0.01
'                            End If
'                        End If
'
'                    End If
'
'                    DDevice.SetLight (B.LightIndex - 1), Lights(B.LightIndex)
'                    DDevice.LightEnable (B.LightIndex - 1), 1
'                End If
'            Else
'                DDevice.LightEnable B.LightIndex - 1, 1
'            End If
'        Else
'            DDevice.LightEnable B.LightIndex - 1, False
'        End If
'
'    Next
'
'    Static GravityArea As Double
'    Dim AverageMotion As Long
'    Dim MotionFactor As Long
'    Static Motionhandled As Long
'    If Motionhandled <= 0 Then Motionhandled = 1
'    Dim MotionCount As Long
'
'    Dim i As Long
'    Dim dist As Single
'    Dim dist2 As Single
'
'    Dim tmp As D3DVECTOR
'    Dim sow As D3DVECTOR
'
'    Dim act As Motion
'    Dim matTest1 As D3DMATRIX
'    Dim matTest2 As D3DMATRIX
'    Dim orgin1 As D3DVECTOR
'    Dim orgin2 As D3DVECTOR
'
'    D3DXMatrixIdentity matTest1
'    D3DXMatrixIdentity matTest2
'
'    Dim m2 As Variant
'    Dim m4 As Variant
'    Dim m3 As Molecule
'    Dim key As Variant
'    Dim e As Billboard
'    Dim m As Molecule
'    Dim p As Planet
'
'    Dim matMesh As D3DMATRIX
'    Dim matMesh2 As D3DMATRIX
'
'    D3DXMatrixIdentity matMesh
'    DDevice.SetTransform D3DTS_WORLD, matMesh
'
'    GravityArea = 0
'
'    For m4 = 1 To Molecules.Count
'        Set m = Molecules(m4)
'
'        OrientateMolecule matMesh, m, Not (m.MeshIndex = -1)
'
'        If m.Visible Then
'
'            If (m.MeshIndex > 0) Then
'
'                If (Distance(MoleculeView.Origin.X, MoleculeView.Origin.Y, MoleculeView.Origin.Z, m.Origin.X, m.Origin.Y, m.Origin.Z) <= FAR) Then
'
'                    If (Meshes(m.MeshIndex).MaterialCount > 0) And m.MeshIndex = 0 Then
'
'                        For i = 0 To Meshes(m.MeshIndex).MaterialCount - 1
'
'                            Select Case TypeName(Meshes(m.MeshIndex).Textures(i))
'                                Case "Billboard"
'                                    Set e = Meshes(m.MeshIndex).Textures(i)
'
'                                    SetRenderBlends e.Transparent, e.Translucent
'
'                                    If Not (e.Translucent Or e.Transparent) Then
'                                        DDevice.SetMaterial GenericMaterial
'                                        DDevice.SetTexture 0, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                                        DDevice.SetTexture 1, Nothing
'                                    Else
'
'                                        DDevice.SetMaterial LucentMaterial
'                                        DDevice.SetTexture 0, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                                        DDevice.SetMaterial GenericMaterial
'                                        DDevice.SetTexture 1, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                                    End If
'                                    Meshes(m.MeshIndex).mesh.DrawSubset i
'
'                                    Set e = Nothing
'                                Case "Direct3DTexture8", "Unknown"
'                                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'                                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                                    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'                                    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'                                    DDevice.SetMaterial LucentMaterial
'                                    DDevice.SetTexture 0, Meshes(m.MeshIndex).Textures(i)
'                                    DDevice.SetMaterial GenericMaterial
'                                    DDevice.SetTexture 1, Meshes(m.MeshIndex).Textures(i)
'                                    Meshes(m.MeshIndex).mesh.DrawSubset i
'                                Case Else
'                                    Debug.Print TypeName(Meshes(m.MeshIndex).Textures(i))
'                            End Select
'
'
'
'                        Next
'
'                    End If
'                End If
'            End If
'        End If
'
'        If (Not (m.OnInRange Is Nothing)) And (Not (m.Collision = None)) Then
'            If m.OnInRange.ApplyTo.Count = 0 Then
'                MoleculeIterateEvent All, m4, m
'            Else
'                MoleculeIterateEvent m.OnInRange.ApplyTo, m4, m
'            End If
'        ElseIf ((m.Collision And Ranged) = Ranged) Then
'
'        End If
'        If (Not (m.OnOutRange Is Nothing)) And (Not (m.Collision = None)) Then
'            If m.OnOutRange.ApplyTo.Count = 0 Then
'                MoleculeIterateEvent All, m4, m
'            Else
'                MoleculeIterateEvent m.OnOutRange.ApplyTo, m4, m
'            End If
'        ElseIf ((m.Collision And Ranged) = Ranged) Then
'
'        End If
'
'        If m.Motions.Count > 0 Then
'            AverageMotion = AverageMotion + m.Motions.Count
'            MotionFactor = MotionFactor + 1
'            MotionCount = Motionhandled
'
'            Do
'
'                tmp = V(m.Origin)
'                D3DXVec3Add sow, tmp, CalculateMotion(m.Motions.Item(1), Direct)
'                If sow.X <> m.Origin.X Or sow.Y <> m.Origin.Y Or sow.Z <> m.Origin.Z Then
'                    m.Origin.X = sow.X
'                    m.Origin.Y = sow.Y
'                    m.Origin.Z = sow.Z
'                    If CInt(m.Collision) > Through Then
'                        If TestCollision(m4, m) Then
'                            m.Origin.X = tmp.X
'                            m.Origin.Y = tmp.Y
'                            m.Origin.Z = tmp.Z
'
'                            If ((m.Collision And Curbing) = Curbing) Then
'
'                            End If
'                        End If
'                    End If
'                End If
'
'                tmp = V(m.Rotate)
'                D3DXVec3Add sow, tmp, CalculateMotion(m.Motions.Item(1), Rotate)
'                If sow.X <> m.Rotate.X Or sow.Y <> m.Rotate.Y Or sow.Z <> m.Rotate.Z Then
'                    m.Rotate.X = sow.X
'                    m.Rotate.Y = sow.Y
'                    m.Rotate.Z = sow.Z
'                    If CInt(m.Collision) > Through Then
'                        If TestCollision(m4, m) Then
'                            m.Rotate.X = tmp.X
'                            m.Rotate.Y = tmp.Y
'                            m.Rotate.Z = tmp.Z
'
'                            If ((m.Collision And Curbing) = Curbing) Then
'
'                            End If
'                        End If
'                    End If
'                End If
'
'                tmp = V(m.Scaled)
'                D3DXVec3Add sow, tmp, CalculateMotion(m.Motions.Item(1), Scalar)
'                If sow.X <> m.Scaled.X Or sow.Y <> m.Scaled.Y Or sow.Z <> m.Scaled.Z Then
'                    m.Scaled.X = sow.X
'                    m.Scaled.Y = sow.Y
'                    m.Scaled.Z = sow.Z
'                    If CInt(m.Collision) > Through Then
'                        If TestCollision(m4, m) Then
'                            m.Scaled.X = tmp.X
'                            m.Scaled.Y = tmp.Y
'                            m.Scaled.Z = tmp.Z
'
'                            If ((m.Collision And Curbing) = Curbing) Then
'
'                            End If
'                        End If
'                    End If
'                End If
'
'                If m.Motions(1).Reactive > -1 Then
'                    If (Timer - m.Motions(1).Latency) > m.Motions(1).Reactive Then
'                        m.Motions(1).Latency = Timer
'                        m.Motions(1).Emphasis = m.Motions(1).Initials
'                        Set act = m.Motions(1)
'                        key = m.Motions.key(1)
'                        If act.Recount > -1 Then
'                            m.Motions.Remove 1
'                            If act.Recount > 0 Then
'                                act.Recount = act.Recount - 1
'                                AddMotion m, act.Action, act.Axis, act.Emphasis, act.Friction, act.Reactive, act.Recount, key
'                            End If
'                        Else
'                            BackOfTheLine m.Motions
'                        End If
'                        Set act = Nothing
'                    Else
'                        BackOfTheLine m.Motions
'                    End If
'
'                ElseIf (((m.Motions(1).Emphasis = 0) Or (m.Motions(1).Recount = 0)) And _
'                    (Not m.Motions(1).Reactive = -1)) Or ((m.Motions(1).Friction <> 0) And _
'                    (m.Motions(1).Emphasis <= 0)) Then
'                    m.Motions.Remove 1
'                Else
'                    BackOfTheLine m.Motions
'                End If
'
'                MotionCount = MotionCount = 1
'
'            Loop Until MotionCount = 0 Or m.Motions.Count = 0
'
'
''        ElseIf CInt(m.Collision) > 0 Then
''            TestCollision m4, m
'        End If
'        If m.SurfaceArea > GravityArea Then GravityArea = m.SurfaceArea
'
'        If (m.Collision > 0) Then
'
'            If ((m.Collision And Gravity) = Gravity) And (Not ((m.Collision And Liquid) = Liquid)) Then
'
'                If m.SurfaceArea < GravityArea Then AddMotion m, Direct, GravityVector, 1, 0, 0, 0, "Gravity"
'
'            ElseIf ((m.Collision And Liquid) = Liquid) Then
'
'                If m.SurfaceArea < GravityArea Then AddMotion m, Direct, LiquidVector, 2, 0, 0, 0, "Liquid"
'
'            End If
'
'            If ((m.Collision And Freely) = Freely) Then
'                DeleteMotion m, "Gravity"
'                DeleteMotion m, "Liquid"
'
'            End If
'
'        End If
'
'        Set m = Nothing
'
'    Next
'
'
'    If MotionFactor > 0 Then
'        AverageMotion = AverageMotion \ MotionFactor
'
'        If FPSRate < Motionhandled * MotionFactor Or AverageMotion < Motionhandled Then
'            Motionhandled = Motionhandled - 1
'        ElseIf Motionhandled * MotionFactor < FPSRate And AverageMotion > Motionhandled Then
'            Motionhandled = Motionhandled + 1
'        End If
'    End If
'
'    If Motionhandled < 1 Then
'        If FPSRate \ 2 > 1 Then
'            Motionhandled = FPSRate \ 2
'        Else
'            Motionhandled = 1
'        End If
'    End If
'
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'    D3DXMatrixIdentity matWorld
'    DDevice.SetTransform D3DTS_WORLD, matWorld
'
'    For Each e In Billboards
'
'        If (e.Animated > 0) Then
'            If CDbl(Timer - e.AnimateTimer) >= e.Animated Or e.AnimateTimer = 0 Then
'                e.AnimateTimer = GetTimer
'                e.AnimatePoint = e.AnimatePoint + 1
'            End If
'        End If
'
'        If e.Visible And ((e.Form And ThreeDimensions) = ThreeDimensions) Then
'
'            If Not (Faces(e.FaceIndex).VBuffer Is Nothing) Then
'
'                If Distance(MoleculeView.Origin.X, MoleculeView.Origin.Y, MoleculeView.Origin.Z, e.Center.X, e.Center.Y, e.Center.Z) <= FAR Then
'                    If e.Transposing Is Nothing Then
'
'                        SetRenderBlends e.Transparent, e.Translucent
'
'                        If Not (e.Translucent Or e.Transparent) Then
'                            DDevice.SetMaterial GenericMaterial
'                            DDevice.SetTexture 0, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                            DDevice.SetTexture 1, Nothing
'                        Else
'
'                            DDevice.SetMaterial LucentMaterial
'                            DDevice.SetTexture 0, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                            DDevice.SetMaterial GenericMaterial
'                            DDevice.SetTexture 1, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                        End If
'
'                        DDevice.SetStreamSource 0, Faces(e.FaceIndex).VBuffer, Len(Faces(e.FaceIndex).Verticies(0))
'                        DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'
'                    End If
'
'                End If
'            End If
'        End If
'    Next

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
Private Sub BackOfTheLine(ByRef Line As NTNodes10.Collection)
    Dim Key As String
    Dim Obj As Object
    Key = Line.Key(1)
    Set Obj = Line.Item(1)
    Line.Remove 1
    Line.Add Obj, Key
End Sub

Public Sub BeginMirrors(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)

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


Public Sub RenderMirror(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)

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



