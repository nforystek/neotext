Attribute VB_Name = "modInfo"

#Const modInfo = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public HelpStage As Boolean

Private FadeTime As Long
Private FadeText As String

Private Width As Single '          640.0f
Private Height As Single '         480.0f
Private WIDTH_DIV_2 As Single '     (WIDTH*0.5f)
Private HEIGHT_DIV_2 As Single '    (HEIGHT*0.5f)

Private CircleImgText As Direct3DTexture8
Private CircleImgVert(0 To 4) As TVERTEX1
Public vIntersect As D3DVECTOR


Public Sub CreateInfo()
    
    Set CircleImgText = LoadTexture(AppPath & "Base\circle.bmp")

    Width = frmMain.ScaleWidth / Screen.TwipsPerPixelX
    Height = frmMain.ScaleHeight / Screen.TwipsPerPixelY
    WIDTH_DIV_2 = Width * 0.5
    HEIGHT_DIV_2 = Height * 0.5
    
End Sub

Public Sub CleanupInfo()
    Set CircleImgText = Nothing
End Sub

Public Sub DrawPointer(ByVal X As Single, ByVal Y As Single)
    DDevice.SetVertexShader FVF_SCREEN

    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR

    DDevice.SetMaterial GenericMat
    If DDevice.GetRenderState(D3DRS_AMBIENT) <> RGB(255, 255, 255) Then
        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
    End If

    DDevice.SetTexture 0, CircleImgText
    DDevice.SetTexture 1, CircleImgText
    
    Dim rec As RECT
    GetWindowRect frmStudio.Designer.hwnd, rec
    
  '  x = x + rec.left
  '  y = rec.top + InvertNum(y, rec.Bottom - rec.top)
  '  x = ((frmMain.Width / Screen.TwipsPerPixelX) / 2)
  '  y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2)

    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)

    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
End Sub


Public Sub RenderInfo()

'    MainFont.Begin
'
'    Dim wholeX As Single
'    Dim wholeY As Single
'    Dim unitX As Single
'    Dim unitY As Single
'
'    Dim x As Single
'    Dim y As Single
'
'    Dim dy As Single
'    Dim dx As Single
'
'    dx = (Player.Location.x / Screen.Width) * Screen.TwipsPerPixelX
'    dy = (Player.Location.y / Screen.Height) * Screen.TwipsPerPixelY
'
'
'    Dim pt As POINTAPI
'    GetCursorPos pt
'
'    Dim rec As RECT
'    GetWindowRect frmStudio.Designer.hWNd, rec
'
'    x = (pt.x - rec.left)
'    y = (pt.y - rec.top)
'
'    DrawTextByCoord "A", x + dx, y - dy
'
'    MainFont.End
'
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
'    DDevice.SetTexture 0, CircleImgText
'    DDevice.SetTexture 1, CircleImgText
'
'
'  '  x = ((frmMain.Width / Screen.TwipsPerPixelX) / 2)
'  '  y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2)
'
'    CircleImgVert(0) = MakeScreen(x - 5, y + 5, -1, 1, 1)
'    CircleImgVert(1) = MakeScreen(x + 5, y + 5, -1, 0, 1)
'    CircleImgVert(2) = MakeScreen(x - 5, y - 5, -1, 1, 0)
'    CircleImgVert(3) = MakeScreen(x + 5, y - 5, -1, 0, 0)
'
'    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'
    
    
    
    

''    MainFont.Begin
''
''    Dim wholeX As Single
''    Dim wholeY As Single
''    Dim unitX As Single
''    Dim unitY As Single
''
''
'    Dim dy As Single
'    Dim dx As Single
'
'    dx = (Player.Location.X / (Screen.Width * Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
'    dy = (-Player.Location.Y / (Screen.Height * Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
''
''
'    Dim pt As POINTAPI
'    GetCursorPos pt
'''
'    Dim rec As RECT
'    GetWindowRect frmStudio.Designer.hWNd, rec
'''
'
'    pt.X = (pt.X + rec.left)
'    pt.Y = (pt.Y + rec.top)
'
'    Dim V As D3DVECTOR
'    V = ScreenXYTo3DZ0((pt.X - rec.Right) + rec.left, (pt.Y - rec.Bottom) + rec.top)
''
''    DrawTextByCoord "A",
''
''    MainFont.End
'
'
'
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
'    DDevice.SetTexture 0, CircleImgText
'    DDevice.SetTexture 1, CircleImgText
'
'    Dim X As Single
'    Dim Y As Single
'
'    X = frmStudio.LastX + dx
'    Y = frmStudio.LastY + dy ' InvertNum(v.y + dy, rec.Bottom - rec.top)
'
'    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
'    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
'    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
'    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
'
'    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
    
    

End Sub

Function MakeScreen(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal tu As Single = 0, Optional ByVal tv As Single = 0) As TVERTEX1
    MakeScreen.X = X
    MakeScreen.Y = Y
    MakeScreen.Z = Z
    MakeScreen.RHW = 1
    MakeScreen.Color = D3DColorARGB(255, 255, 255, 255)
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




