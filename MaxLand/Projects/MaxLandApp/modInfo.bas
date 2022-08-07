Attribute VB_Name = "modInfo"
#Const modInfo = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private FadeTime As Long
Private FadeText As String

Private PointImgText As Direct3DTexture8
Private PointImgVert(0 To 4) As MyScreen

Private CircleImgText As Direct3DTexture8
Private CircleImgVert(0 To 4) As MyScreen

Public Sub CreateInfo()
    
   ' Set CircleImgText = LoadTexture(AppPath & "Base\circle.bmp")
    
End Sub
Public Sub CleanupInfo()

    'Set CircleImgText = Nothing
    
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

Public Sub RenderInfo()

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
    
        If DebugMode Then
            DrawText "F1=Hide, E/D=Forward/Back, W/R=Left/Right, Arrows=Direction, SPACE=Jump, TAB=View, F2=Stats, F3=Credits, F5=Culling", 2, 2
            txt = ""
            If CullingSetup > 0 Then
                txt = "Use number keys to set vistype, use F5 to set direction, use F6 to set upvector, use F7 to complete" & vbCrLf & _
                        "the culling call setup and commit it at location.  Use F8 to clear all the commited culling calls."
            End If
            txt = txt & vbCrLf & vbCrLf & lCullCalls & " calls to Culling are setup."
        
            DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), Row(6)
        Else
            If ShowHelp Then
                DrawText "F1=Hide, E/D=Forward/Back, W/R=Left/Right, Arrows=Direction, SPACE=Jump, TAB=View, F2=Stats, F3=Credits", 2, 2
            Else
                DrawText "F1=Help", 2, 2
            End If
        End If
                
        If ShowHelp And ShowStat Then
        
            DrawText "", 2, Row(1)
            DrawText "Per World Stats: " & lngFaceCount & " Triangles Total; Culling " & ((lCulledFaces * lCullCalls) \ FPSRate) & " Triangles Ignored", 2, Row(2)
            DrawText "Per Frame Stats: " & lngTestCalls & " Calls To Collision; Totaling at " & lFacesShown & " Triangles", 2, Row(3)

        End If
        
        If CheckIdle(4) Or ShowHelp Then
            If VariableCount > 0 Then
                For cnt = LBound(Variables) To UBound(Variables)
                    If LCase(Variables(cnt).Identity) = "idletext" Then
                        txt = Variables(cnt).Value
                        Exit For
                    End If
                Next
            End If
            DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), Row(5)
        End If
        
        If ShowHelp Then
            txt = "TILDA=Console, ESC=" & IIf(TrapMouse, "Exit", "Close")
        Else
            txt = "ESC=" & IIf(TrapMouse, "Exit", "Close")
        End If
        DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), 2

            
        If ShowHelp And ShowStat Then
            txt = "Frames Per Second: " & FPSRate
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(2)
        
        End If
        
        If Perspective = CameraMode Then
            If Player.CameraIndex > 0 Then
                txt = "Current Camera View " & Player.CameraIndex
            Else
                txt = "Current Camera View NA"
            End If
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(3)
        End If

        If DrawCount > 0 Then

            For cnt = 1 To DrawCount
                If Draws(1, cnt) <> "" Then DrawText ParseValues(CStr(Draws(1, cnt))), CSng(Draws(2, cnt)), CSng(Draws(3, cnt))
            Next
        End If

    End If
    
    If ShowCredits Then
        If VariableCount > 0 Then
            For cnt = LBound(Variables) To UBound(Variables)
                If LCase(Variables(cnt).Identity) = "credittext" Then
                    txt = Replace(Variables(cnt).Value, "\n", vbCrLf)
                    Exit For
                End If
            Next
        End If
                
        DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - ((frmMain.TextHeight(txt) / Screen.TwipsPerPixelY) / 2)
    Else
        If Not (FadeText = "") Then
            If (Timer - FadeTime) >= 6 Then
                FadeText = ""
            Else
                DrawText FadeText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(FadeText) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - (TextHeight / 2)
            End If
        End If
    End If

    DDevice.SetPixelShader PixelShaderDefault

    DDevice.SetVertexShader FVF_SCREEN
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        
    If ScreenImageCount > 0 Then
        
        Dim i As Long

        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
        DDevice.SetMaterial GenericMaterial
    
        For i = 1 To ScreenImageCount
            If ScreenImages(i).Visible Then
                If ScreenImages(i).Translucent Then
                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
                ElseIf ScreenImages(i).blackalpha Then
                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
                Else
                    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                End If
            
                DDevice.SetTexture 0, ScreenImages(i).Image
                DDevice.SetTexture 1, ScreenImages(i).Image
                DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImages(i).Verticies(0), LenB(ScreenImages(i).Verticies(0))

            End If
        Next
        
    End If

        
    If (lngFaceCount > 0) And DebugMode Then
    
        DDevice.SetVertexShader FVF_SCREEN
        DDevice.SetRenderState D3DRS_ZENABLE, False
        DDevice.SetRenderState D3DRS_LIGHTING, False
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        
        DDevice.SetTexture 0, CircleImgText
        DDevice.SetTexture 1, CircleImgText


        For i = 0 To Culling(1, lngFaceCount, sngCamera, sngFaceVis, sngVertexX, sngVertexY, sngVertexZ, sngScreenX, sngScreenY, sngScreenZ, sngZBuffer) - 1
            If sngZBuffer(2, sngZBuffer(3, i)) > 0 Then
              '  DrawText CStr(i), ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngZBuffer(0, sngZBuffer(3, cnt)), ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngZBuffer(1, sngZBuffer(3, cnt))
                DrawText CStr(i), (((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(0, sngFaceVis(5, i))) / (FOVY / AspectRatio), (((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(0, sngFaceVis(5, i))) / (FOVY / AspectRatio) 'sngZBuffer(0, sngZBuffer(3, cnt)), sngZBuffer(1, sngZBuffer(3, cnt))
                
            End If
        Next

        Dim X As Single
        Dim Y As Single
        Dim z As Single
        

        For i = 0 To lngFaceCount - 1

                If sngFaceVis(3, sngFaceVis(5, i)) = 1 Then

                    If sngScreenZ(0, sngFaceVis(5, i)) > 0 Then


                        X = (((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(0, sngFaceVis(5, i))) '/ (FOVY * AspectRatio)
                        Y = (((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(0, sngFaceVis(5, i))) '/ (FOVY * AspectRatio)

                        CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 0, 0)
                        CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 1, 0)
                        CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 0, 1)
                        CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 1, 1)

                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
                        'dx.DrawCircle ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(0, sngFaceVis(5, i)), ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(0, sngFaceVis(5, i)), 2
                    End If
                    If sngScreenZ(1, sngFaceVis(5, i)) > 0 Then
                        X = ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(0, sngFaceVis(5, i)) '/ (FOVY * AspectRatio)
                        Y = ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(0, sngFaceVis(5, i)) '/ (FOVY * AspectRatio)

                        CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 0, 0)
                        CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 1, 0)
                        CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 0, 1)
                        CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 1, 1)

                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
                        'D3D.DrawCircle ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(1, sngFaceVis(5, i)), ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(1, sngFaceVis(5, i)), 2
                    End If
                    If sngScreenZ(2, sngFaceVis(5, i)) > 0 Then
                        X = ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(0, sngFaceVis(5, i)) '/ (FOVY * AspectRatio)
                        Y = ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(0, sngFaceVis(5, i)) '/ (FOVY * AspectRatio)

                        CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 0, 0)
                        CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 1, 0)
                        CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 0, 1)
                        CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 1, 1)

                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))

                        'D3D.DrawCircle ((frmMain.width / Screen.TwipsPerPixelX) / 2) - sngScreenX(2, sngFaceVis(5, i)), ((frmMain.height / Screen.TwipsPerPixelY) / 2) - sngScreenY(2, sngFaceVis(5, i)), 2
                    End If


                End If
        Next


    End If
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
End Sub



