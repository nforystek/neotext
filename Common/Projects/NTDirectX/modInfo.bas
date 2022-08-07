Attribute VB_Name = "modInfo"

#Const modInfo = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Private FadeTime As Long
Private FadeText As String

'Private Width As Single '          640.0f
'Private Height As Single '         480.0f
'Private WIDTH_DIV_2 As Single '     (WIDTH*0.5f)
'Private HEIGHT_DIV_2 As Single '    (HEIGHT*0.5f)
'
'Private CircleImgText As Direct3DTexture8
'Private CircleImgVert(0 To 4) As MyScreen
'Public vIntersect As D3DVECTOR


Public Sub CreateInfo()
    
'    Set CircleImgText = LoadTextureRes(LoadResData(109, "CUSTOM")) ' LoadTexture(AppPath & "Base\circle.bmp")
'
'    Width = frmMain.ScaleWidth / Screen.TwipsPerPixelX
'    Height = frmMain.ScaleHeight / Screen.TwipsPerPixelY
'    WIDTH_DIV_2 = Width * 0.5
'    HEIGHT_DIV_2 = Height * 0.5
    
End Sub

Public Sub CleanupInfo()
'    Set CircleImgText = Nothing
End Sub

Public Sub FadeMessage(ByVal txt As String)
    FadeTime = Timer
    txt = Replace(txt, "\n", vbCrLf)
    FadeText = ParseValues(txt)
    AddMessage txt
End Sub

Public Function Row(ByVal num As Long) As Long
    Row = ((TextHeight \ Screen.TwipsPerPixelY) * num) + (2 * num)
End Function

Public Sub RenderInfo(ByRef UserControl As ScratchKad)

    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR

    Dim cnt As Long
    Dim txt As String
    If Not ConsoleVisible Then
    
'        If DebugMode Then
'            DrawText "F1=Hide, E/D=Forward/Back, W/R=Left/Right, Arrows=Direction, SPACE=Jump, TAB=View, F2=Stats, F3=Credits, F5=Culling", 2, 2
'            txt = ""
'            If CullingSetup > 0 Then
'                txt = "Use number keys to set vistype, use F5 to set direction, use F6 to set upvector, use F7 to complete" & vbCrLf & _
'                        "the culling call setup and commit it at location.  Use F8 to clear all the commited culling calls."
'            End If
'            txt = txt & vbCrLf & vbCrLf & lCullCalls & " calls to Culling are setup."
'
'            DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), Row(6)
'        Else
'            If ShowHelp Then
'                DrawText "F1=Hide, E/D=Forward/Back, W/R=Left/Right, Arrows=Direction, SPACE=Jump, TAB=View, F2=Stats, F3=Credits", 2, 2
'            Else
'                DrawText "F1=Help", 2, 2
'            End If
'        End If
'
'        If ShowHelp And ShowStat Then
'
'            DrawText "", 2, Row(1)
'            DrawText "Per World Stats: " & lngFaceCount & " Triangles Total; Culling " & ((lCulledFaces * lCullCalls) \ FPSRate) & " Triangles Ignored", 2, Row(2)
'            DrawText "Per Frame Stats: " & lngTestCalls & " Calls To Collision; Totaling at " & lFacesShown & " Triangles", 2, Row(3)
'
'        End If
'
'        If CheckIdle(4) Or ShowHelp Then
'            If VariableCount > 0 Then
'                For cnt = LBound(Variables) To UBound(Variables)
'                    If LCase(Variables(cnt).Identity) = "idletext" Then
'                        txt = Variables(cnt).Value
'                        Exit For
'                    End If
'                Next
'            End If
'            DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), Row(5)
'        End If
'
'        If ShowHelp Then
'            txt = "TILDA=Console, ESC=" & IIf(TrapMouse, "Exit", "Close")
'        Else
'            txt = "ESC=" & IIf(TrapMouse, "Exit", "Close")
'        End If
'        DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), 2
'
'
'        If ShowHelp And ShowStat Then
'            txt = "Frames Per Second: " & FPSRate
'            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(2)
'
'        End If
'
'        If Perspective = CameraMode Then
'            If Player.CameraIndex > 0 Then
'                txt = "Current Camera View " & Player.CameraIndex
'            Else
'                txt = "Current Camera View NA"
'            End If
'            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(3)
'        End If
'
'        If DrawCount > 0 Then
'
'            For cnt = 1 To DrawCount
'                If Draws(1, cnt) <> "" Then DrawText ParseValues(CStr(Draws(1, cnt))), CSng(Draws(2, cnt)), CSng(Draws(3, cnt))
'            Next
'        End If

    End If
    
    If Not (FadeText = "") Then
        If (Timer - FadeTime) >= 6 Then
            FadeText = ""
        Else
            DrawText FadeText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(FadeText) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - (TextHeight / 2)
        End If
    End If

    DDevice.SetPixelShader PixelShaderDefault

    DDevice.SetVertexShader FVF_SCREEN
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        
'    If ScreenImageCount > 0 Then
'
'        Dim i As Long
'
'        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'        DDevice.SetMaterial GenericMaterial
'
'        For i = 1 To ScreenImageCount
'            If ScreenImages(i).Visible Then
'                If ScreenImages(i).Translucent Then
'                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'                ElseIf ScreenImages(i).BlackAlpha Then
'                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
'                Else
'                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                End If
'
'                DDevice.SetTexture 0, ScreenImages(i).Image
'                DDevice.SetTexture 1, ScreenImages(i).Image
'                DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImages(i).Verticies(0), LenB(ScreenImages(i).Verticies(0))
'
'            End If
'        Next
'
'    End If
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
End Sub





'Public Sub DrawPointer(ByRef UserControl As ScratchKad, ByVal X As Single, ByVal Y As Single)
'    DDevice.SetVertexShader FVF_SCREEN
'
'    DDevice.SetRenderState D3DRS_ZENABLE, False
'    DDevice.SetRenderState D3DRS_LIGHTING, False
'    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
''    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'
'    DDevice.SetMaterial GenericMat
'    If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
'        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
'    End If
'
'    DDevice.SetTexture 0, CircleImgText
'    DDevice.SetTexture 1, CircleImgText
'
'    Dim rec As RECT
'    GetWindowRect UserControl.hWnd, rec
'
'  '  x = x + rec.left
'  '  y = rec.top + InvertNum(y, rec.Bottom - rec.top)
'  '  x = ((frmMain.Width / Screen.TwipsPerPixelX) / 2)
'  '  y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2)
'
'    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
'    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
'    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
'    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
'
'    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'End Sub


'Public Sub RenderInfo(ByRef UserControl As ScratchKad)
'
'
''    DrawPointer UserControl, frmMain.LastX, frmMain.LastY
'
'    MainFont.Begin
'
'
''    Dim vDir As D3DVECTOR
''    Dim vIntersect As D3DVECTOR
''
''    MouseX = Tan(FOV / 2) * (frmMain.LastX / (frmMain.ScaleWidth / 2) - 1) / ASPECT
''    MouseY = Tan(FOV / 2) * (1 - frmMain.LastY / (frmMain.ScaleHeight / 2))
'
''    Dim p1 As D3DVECTOR 'StartPoint on the nearplane
''    Dim p2 As D3DVECTOR 'EndPoint on the farplane
''
''    p1.x = MouseX * NEAR
''    p1.y = MouseY * NEAR
''    p1.Z = NEAR
''
''    p2.x = MouseX * FAR
''    p2.y = MouseY * FAR
''    p2.Z = FAR
'
'
'
'
''    Dim wholeX As Single
''    Dim wholeY As Single
''    Dim unitX As Single
''    Dim unitY As Single
''
'    Dim X As Single
'    Dim Y As Single
'
'    Dim dy As Single
'    Dim dx As Single
'
'
'    Dim r1 As Single
'    Dim r2 As Single
'
'
'    Dim pt As POINTAPI
'    GetCursorPos pt
'
'    Dim rec As RECT
'    Dim rec2 As RECT
'    GetWindowRect UserControl.hWnd, rec
'
'    GetWindowRect UserControl.Parent.hWnd, rec2
'
'    If ((pt.X >= rec.Left) And (pt.X <= rec.Right)) And ((pt.Y >= rec.Top) And (pt.Y <= rec.Bottom)) Then
'
'        'Screen.MousePointer = 99
'
'        Dim MouseX As Single
'        Dim MouseY As Single
'
'        Dim vDir As D3DVECTOR
'        Dim vIntersect As D3DVECTOR
'
'        MouseX = Tan(FOV / 2) * (pt.X / ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - 1) / ASPECT
'        MouseY = Tan(FOV / 2) * (1 - pt.Y / ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2))
'
'        Dim p1 As D3DVECTOR 'StartPoint on the nearplane
'        Dim p2 As D3DVECTOR 'EndPoint on the farplane
'
'        p1.X = MouseX * NEAR
'        p1.Y = MouseY * NEAR
'        p1.z = NEAR
'
'        p2.X = MouseX * FAR
'        p2.Y = MouseY * FAR
'        p2.z = FAR
'
'        'Inverse the view matrix
'        Dim matInverse As D3DMATRIX
'        DDevice.GetTransform D3DTS_VIEW, matView
'
'        D3DXMatrixInverse matInverse, 0, matView
'
'        VectorMatrixMultiply p1, p1, matInverse
'        VectorMatrixMultiply p2, p2, matInverse
'        D3DXVec3Subtract vDir, p2, p1
'
'        'Check if the points hit
'        Dim v1 As D3DVECTOR
'        Dim v2 As D3DVECTOR
'        Dim v3 As D3DVECTOR
'
'        Dim v4 As D3DVECTOR
'        Dim v5 As D3DVECTOR
'        Dim v6 As D3DVECTOR
'
'        Dim pPlane1 As D3DVECTOR4
'
'        Dim cnt As Long
'        v1.X = frmMain.Width '/ Screen.Width)
'        v1.Y = -frmMain.Height ' / Screen.Height)
'        v1.z = 1
'
'        v2.X = -frmMain.Width '/ Screen.Width)
'        v2.Y = -frmMain.Height ' / Screen.Height)
'        v2.z = 1
'
'        v3.X = -frmMain.Width '/ Screen.Width)
'        v3.Y = frmMain.Height '/ Screen.Height)
'        v3.z = 1
'
'        pPlane1 = Create4DPlaneVectorFromPoints(v1, v2, v3)
'
'        Dim c As D3DVECTOR
'        Dim N As D3DVECTOR
'        Dim P As D3DVECTOR
'        Dim V As D3DVECTOR
'
'        Dim hit As Boolean
'        LastMouseSetX = MouseSetX
'        LastMouseSetY = MouseSetY
'
'        MouseSetX = Round(MouseX, 5)
'        MouseSetY = Round(MouseY, 5)
'
'        hit = RayIntersectPlane(pPlane1, p1, vDir, vIntersect)
'
'       ' dx = ((vIntersect.x / Screen.Width) * Screen.TwipsPerPixelX)
'       ' dy = ((vIntersect.y / Screen.Height) * Screen.TwipsPerPixelY)
'
'        LastScreenSetX = ScreenSetX
'        LastScreenSetY = ScreenSetY
'        ScreenSetX = vIntersect.X
'        ScreenSetY = vIntersect.Y
'
'        If hit Then
'
'            'vIntersect.x = (vIntersect.x - rec2.Left)
'           ' vIntersect.y = (-(vIntersect.y + rec2.Top))
'
'           ' x = vIntersect.x
'           ' y = vIntersect.y
'
'
'          '  V = ScreenXYTo3DZ0(x, y)
'
'
'            DrawTextByCoord "A", X, Y
'
'        End If
'
'   ' Else
'   '     Screen.MousePointer = 0
'    End If
'
'   ' Debug.Print hit; X; Y
'
'    MainFont.End
'
'    DDevice.SetVertexShader FVF_SCREEN
'
'    DDevice.SetRenderState D3DRS_ZENABLE, False
'    DDevice.SetRenderState D3DRS_LIGHTING, False
'    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'
'    DDevice.SetMaterial GenericMat
'    If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
'        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
'    End If
'
'  '  If ((pt.x >= rec.Left) And (pt.x <= rec.Right)) And ((pt.y >= rec.Top) And (pt.y <= rec.Bottom)) Then
'
'
'        DDevice.SetTexture 0, CircleImgText
'        DDevice.SetTexture 1, CircleImgText
'
'       ' x = ((frmMain.Width / Screen.TwipsPerPixelX) / 2)
'       ' y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2)
'
'        CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
'        CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
'        CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
'        CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
'
'        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'
' '   End If
'
'
''    MainFont.Begin
''
''    Dim wholeX As Single
''    Dim wholeY As Single
''    Dim unitX As Single
''    Dim unitY As Single
''
''
''    Dim dy As Single
''    Dim dx As Single
''
''    dx = (Player.Location.X / (Screen.Width * Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
''    dy = (-Player.Location.Y / (Screen.Height * Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
'''
'''
''    Dim pt As POINTAPI
''    GetCursorPos pt
''''
''    Dim rec As RECT
''    GetWindowRect UserControl.hWnd, rec
''''
''
''    pt.X = (pt.X + rec.Left)
''    pt.Y = (pt.Y + rec.Top)
''
''    Dim V As D3DVECTOR
''    V = ScreenXYTo3DZ0((pt.X - rec.Right) + rec.Left, (pt.Y - rec.Bottom) + rec.Top)
'''
'''    DrawTextByCoord "A",
'''
'''    MainFont.End
''
''
''
''
''    DDevice.SetVertexShader FVF_SCREEN
''
''    DDevice.SetRenderState D3DRS_ZENABLE, False
''    DDevice.SetRenderState D3DRS_LIGHTING, False
''    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
''    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
''
''    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
''    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
''    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
''    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
''
''    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''
''    DDevice.SetMaterial GenericMat
''    If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
''        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
''    End If
''
''    DDevice.SetTexture 0, CircleImgText
''    DDevice.SetTexture 1, CircleImgText
''
''    Dim X As Single
''    Dim Y As Single
''
''    X = frmMain.LastX + dx
''    Y = frmMain.LastY + dy ' InvertNum(v.y + dy, rec.Bottom - rec.top)
''
''    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
''    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
''    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
''    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
''
''    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'
'
'
'End Sub
'
Public Function MakeScreen(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal tu As Single = 0, Optional ByVal tv As Single = 0) As MyScreen
    MakeScreen.X = X
    MakeScreen.Y = Y
    MakeScreen.Z = Z
    MakeScreen.rhw = 1
    MakeScreen.clr = D3DColorARGB(255, 255, 255, 255)
    MakeScreen.tu = tu
    MakeScreen.tv = tv
End Function

'Public Function IntentJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    'returns the ratio percentage to a full wholetotal or unitmeasures
'    IntentJesus = (PercentUnit / WholeTotal * UnitMeasure / PercentWhole)
'End Function
'Public Function InventJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    InventJesus = Sqr(PercentUnit ^ 2 / WholeTotal ^ 2 * UnitMeasure ^ 2 / PercentWhole ^ 2)
'End Function
'Public Function InvertJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    InvertJesus = Sqr((PercentUnit ^ 3) / (WholeTotal ^ 2) * (UnitMeasure ^ 3) / (PercentWhole ^ 4))
'End Function
'Public Function InnateJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single) As Single
'    InnateJesus = (Sqr((IntentJesus(WholeTotal, UnitMeasure) ^ 3) / (InvertJesus(WholeTotal, UnitMeasure) ^ 2) * _
'                    (InvertJesus(WholeTotal, UnitMeasure) ^ 3) / (IntentJesus(WholeTotal, UnitMeasure) ^ 4)) + _
'                    Sqr((IntentJesus(UnitMeasure, WholeTotal) ^ 3) / (InvertJesus(UnitMeasure, WholeTotal) ^ 2) * _
'                    (InvertJesus(UnitMeasure, WholeTotal) ^ 3) / (IntentJesus(UnitMeasure, WholeTotal) ^ 4))) / 2
'End Function
'Public Function InnertJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single) As Single
'    InnertJesus = Sqr((InventJesus(WholeTotal, UnitMeasure) ^ 3) / (InventJesus(WholeTotal, UnitMeasure) ^ 2) * _
'                    (InventJesus(WholeTotal, UnitMeasure) ^ 3) / (InventJesus(WholeTotal, UnitMeasure) ^ 4))
'End Function


