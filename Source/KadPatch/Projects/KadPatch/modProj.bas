Attribute VB_Name = "modProj"

Option Explicit

Private Type RenderWall
    Vertex() As TVERTEX2
    Buffer As Direct3DVertexBuffer8
    Skin As Direct3DTexture8
End Type
Private Type ThreadType
    Vertical As RenderWall
    Horizontal As RenderWall
    BackSlash As RenderWall
    ForwardSlash As RenderWall
End Type

Public Const SymbolWidth As Long = 44
Public Const SymbolHeight As Long = 44
    
'input values
Public ThatchHeight As Single
Public ThatchWidth As Single
Public ThatchBlock As Single

'running needed wholes
Public ThatchXUnits As Single
Public ThatchYUnits As Single
Public ThatchScale As Single

Private pProjPath As String
Private pDirty As Boolean

''3d median between aboves
Public ThreadSize As Single
Public DecalWidth As Single
Public DecalHeight As Single
'Public ThatchSize As Single

Public ThreadSqrPerText As Single
Public BlockSqrPerText As Single

Private pThatchIndex As Byte
Private pThatchCount As Byte
Private ThatchSkins() As Direct3DTexture8
Private ThatchVertex() As TVERTEX2
Private ThatchVBuf As Direct3DVertexBuffer8

Private ThreadCount As Long
Private ThreadSkins() As Direct3DTexture8
Private Threadings As ThreadType
Private ThreadVBuf As Direct3DVertexBuffer8

Private SymbolCount As Long
Private SymbolSkins() As Direct3DTexture8
Private SymbolVertex() As TVERTEX2
Private SymbolVBuf As Direct3DVertexBuffer8
Private SymbolKeys As New NTNodes10.Collection

Private PatternGridSKin As Direct3DTexture8

Private PanelVertex() As TVERTEX2
Private PanelVBuf As Direct3DVertexBuffer8

Private FlossCount As Long
Private FlossSkins() As Direct3DTexture8
Private FlossKeys As New NTNodes10.Collection

Private ScreenImgVert(0 To 3) As TVERTEX1

Private MouseSkin As Direct3DTexture8

'mouse location in 3d
Public MouseSetX As Single
Public MouseSetY As Single
Public LastMouseSetX As Single
Public LastMouseSetY As Single

'mouse location in 2d
Public ScreenSetX As Single
Public ScreenSetY As Single
Public LastScreenSetX As Single
Public LastScreenSetY As Single

'row and col mouse is over
Public BlockSetX As Long
Public BlockSetY As Long
Public LastBlockSetX As Long
Public LastBlockSetY As Long

'quadrant detail in block 30x30
Public DetailSetX As Single
Public DetailSetY As Single
Public LastDetailSetX As Single
Public LastDetailSetY As Single

Public ProjGrid() As GridItem
Public Property Get ProjPath() As String
    ProjPath = pProjPath
End Property
Public Property Let ProjPath(ByVal RHS As String)
    pProjPath = RHS
    frmStudio.Caption = "adPatch - [" & pProjPath & "]" & IIf(pDirty, "*", "")
End Property
Public Property Get Dirty() As Boolean
    Dirty = pDirty
End Property
Public Property Let Dirty(ByVal RHS As Boolean)
    pDirty = RHS
    frmStudio.Caption = "adPatch - [" & pProjPath & "]" & IIf(pDirty, "*", "")
End Property
Public Property Get ThatchCount() As Byte
    ThatchCount = pThatchCount
End Property
Private Property Let ThatchCount(ByVal RHS As Byte)
    pThatchCount = RHS
End Property

Public Property Get ThatchIndex() As Byte
    ThatchIndex = pThatchIndex
End Property
Public Property Let ThatchIndex(ByVal RHS As Byte)
    pThatchIndex = RHS
    Dirty = True
End Property

Private Function SymbolToBinary(ByVal Symbol As String) As Boolean()
    Dim cnt As Long

    Dim binary() As Boolean
    
    ReDim binary(1 To SymbolWidth * SymbolHeight) As Boolean
    For cnt = 1 To SymbolWidth * SymbolHeight
        binary(cnt) = CBool("-" & Left(Symbol, 1))
        Symbol = Mid(Symbol, 2)
    Next
    SymbolToBinary = binary

End Function

Private Function SymbolFromBinary(ByRef binary() As Boolean) As String
    Dim cnt As Long
    For cnt = 1 To SymbolWidth * SymbolHeight
        SymbolFromBinary = SymbolFromBinary & Trim(CStr(Abs(CInt(binary(cnt)))))
    Next
End Function

Public Function WriteToDisk() As Boolean
    On Error GoTo faildiskaccess
    'serialization of loaded project to byte array
    Dim clrs As New NTNodes10.Collection
     
    Dim fn As Integer
    fn = FreeFile
    Open ProjPath For Output As #fn
    Close #fn
    
    Open ProjPath For Binary As #fn
    
    Put #fn, 1, ThatchXUnits
    Put #fn, 5, ThatchYUnits
    
    Put #fn, 9, ThatchBlock
    Put #fn, 13, ThatchIndex
    
    Seek #fn, 14
    
    Dim X As Long
    Dim Y As Long
    Dim i As Byte
    Dim c As Long
    
    For X = 1 To ThatchXUnits
        For Y = 1 To ThatchYUnits
            If ProjGrid(X, Y).count > 0 Then
                Put #fn, , ProjGrid(X, Y).count
                For i = 1 To ProjGrid(X, Y).count
                 '   c = ProjGrid(X, Y).Details(i).Color
                 '   c = Val("&" & GetFileTitle(frmStudio.Gallery1.FilePath(c)))
                    If Not clrs.Exists("C" & ProjGrid(X, Y).Details(i).Color) Then clrs.Add ProjGrid(X, Y).Details(i).Color, "C" & ProjGrid(X, Y).Details(i).Color
                    
                    Put #fn, , ProjGrid(X, Y).Details(i).Color
                    Put #fn, , ProjGrid(X, Y).Details(i).Stitch
                Next
            Else
                Put #fn, , CByte(0)
            End If
        Next
    Next
    
    If clrs.count > 0 Then
        Do While clrs.count > 0
            Put #fn, , SymbolToBinary(frmStudio.GetSymbol(clrs(1)))
            clrs.Remove 1
        Loop
    End If
    
    
    Dirty = False
    
faildiskaccess:
    Close #fn
    WriteToDisk = (Err.number = 0)
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function ReadFromDisk() As Boolean
    
    On Error GoTo faildiskaccess
    'serialization of loaded project to byte array
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To ThatchXUnits
        For Y = 1 To ThatchYUnits
            If ProjGrid(X, Y).count > 0 Then
                Erase ProjGrid(X, Y).Details
                ProjGrid(X, Y).count = 0
            End If
        Next
    Next
    Erase ProjGrid
    
    Dim clrs As New NTNodes10.Collection
    
    Dim count As Byte
    Dim i As Byte
    
    Dim tmp As Single
    
    Dim fn As Integer
    fn = FreeFile
'CleanUpProj

    
    Open ProjPath For Binary As #fn
    
    Get #fn, 1, tmp
    ThatchXUnits = tmp
    
    
   ' Width = tmp
    Get #fn, 5, tmp
    ThatchYUnits = tmp
    'Height = tmp
    Get #fn, 9, tmp
    ThatchBlock = tmp
   ' Block = tmp

    ReDim ProjGrid(1 To ThatchXUnits, 1 To ThatchYUnits) As GridItem
    
    Get #fn, 13, i
    ThatchIndex = i

'CreateProj
    Dim total As Long
    Dim cnt As Long

    Seek #fn, 14
    X = 1
    Y = 1
    
    Do
        
        Get #fn, , count
        If count > 0 Then
            ProjGrid(X, Y).count = count
            ReDim Preserve ProjGrid(X, Y).Details(1 To count) As ItemDetail
            i = 0
            Do Until i = count
                i = i + 1
                Get #fn, , total
                
                If Not clrs.Exists("C" & total) Then clrs.Add total, "C" & total
                
                If frmStudio.CreateColor(total, True) Then
                    frmStudio.UpdateGallery
                End If
                
                ProjGrid(X, Y).Details(i).Color = total
                
                Get #fn, , total
                ProjGrid(X, Y).Details(i).Stitch = total
            Loop
        End If
            
        Y = Y + 1
        If Y = ThatchYUnits + 1 Then
            X = X + 1
            If X = ThatchXUnits + 1 Then
                Exit Do
            End If
           Y = 1
        End If
    Loop Until EOF(fn)
    
    If (Not EOF(fn)) And (clrs.count > 0) Then
        If MsgBox("Do you want to import the symbols in the saved file?" & vbCrLf & _
                   "(Note: Yes, will overwrite existing color symbols." & vbCrLf & _
                   "No, erases the files symbols when saving it later.", vbYesNo + vbQuestion) = vbYes Then
        
            Do Until EOF(fn) Or (clrs.count = 0)
                Dim binary(1 To SymbolWidth * SymbolHeight) As Boolean
                Get #fn, , binary
                frmStudio.SetSymbol clrs(1), SymbolFromBinary(binary)
                clrs.Remove 1
            Loop
            
            frmStudio.UpdateGallery
                   
        End If
        
    
    End If
    
    
    
    Dirty = False

    
faildiskaccess:
    Close #fn
    ReadFromDisk = (Err.number = 0)
    If Err Then Err.Clear
    On Error GoTo 0
End Function
Public Function Serialize(ByRef data() As Byte) As Boolean
    On Error GoTo faildiskaccess
    'serialization of loaded project to byte array
    
    
    
    
faildiskaccess:
    Serialize = (Err.number = 0)
    On Error GoTo 0
End Function

Public Function Deserialize(ByRef data() As Byte) As Boolean
    On Error GoTo faildiskaccess
    'deserialization of byte array to project records
    
    
    
    
faildiskaccess:
    Deserialize = (Err.number = 0)
    On Error GoTo 0
End Function

Public Property Get Block() As Single
    Block = ThatchBlock
End Property
Public Property Let Block(ByVal RHS As Single)
    ThatchBlock = RHS
    ThatchXUnits = ThatchWidth \ RHS
    ThatchYUnits = ThatchHeight \ RHS
    ResizeMultidimArray
    GotoCenter
    Dirty = True
End Property

Public Property Get Width() As Single
    Width = ThatchWidth
End Property
Public Property Get Height() As Single
    Height = ThatchHeight
End Property
Public Property Let Width(ByVal RHS As Single)
    ThatchWidth = RHS
    ThatchXUnits = RHS \ ThatchBlock
    ResizeMultidimArray
    GotoCenter
    Dirty = True
End Property
Public Property Let Height(ByVal RHS As Single)
    ThatchHeight = RHS
    ThatchYUnits = RHS \ ThatchBlock
    ResizeMultidimArray
    GotoCenter
    Dirty = True
End Property

Public Sub GotoCenter()
    Player.MoveSpeed = 6
   ' Player.Location.Y = ((BlockHeightY / BlockSqrPerText) * ThatchYUnits) - (frmStudio.Designer.Height / Screen.TwipsPerPixelY)
   ' Player.Location.X = -((BlockWidthX / BlockSqrPerText) * ThatchXUnits) + (frmStudio.Designer.Width / Screen.TwipsPerPixelX)

  ' Player.Location.Y = ((BlockHeightY / BlockSqrPerText) * ThatchYUnits) + ((frmStudio.Designer.Height / Screen.TwipsPerPixelY) / 2)
  '  Player.Location.X = -((BlockWidthX / BlockSqrPerText) * ThatchXUnits) + ((frmStudio.Designer.Width / Screen.TwipsPerPixelX) / 2)
    
    Player.Location.Y = (((BlockHeightY / BlockSqrPerText) * ThatchYUnits) / 2) + ((((frmStudio.Top + frmStudio.Designer.Top) / 2) - (frmStudio.Designer.Height / 2)) / Screen.TwipsPerPixelY)
    Player.Location.X = (((BlockHeightY / BlockSqrPerText) * ThatchXUnits) / 2) + ((((frmStudio.Left + frmStudio.Designer.Left) / 2) - (frmStudio.Designer.Width / 2)) / Screen.TwipsPerPixelX)
    

    Player.Location.Z = 0

    Player.CameraAngle = 0
    Player.CameraPitch = 0
    
    If ThatchYUnits > ThatchXUnits Then
    
        Player.CameraZoom = ThatchYUnits * 80
    Else
        Player.CameraZoom = ThatchXUnits * 80
    End If
End Sub

Public Sub CreateProj()
    
    ResizeMultidimArray
    Set MouseSkin = LoadTextureRes(LoadResData(1, "BMP"))

    Dim PlateSize As Single

    Set Threadings.Vertical.Skin = LoadTexture(AppPath & "Base\Stitchings\Overlay\1_Vertical.bmp")
    ReDim Threadings.Vertical.Vertex(0 To 5) As TVERTEX2
    
    Set Threadings.Horizontal.Skin = LoadTexture(AppPath & "Base\Stitchings\Overlay\2_Horizontal.bmp")
    ReDim Threadings.Horizontal.Vertex(0 To 5) As TVERTEX2
    
    ReDim SymbolVertex(0 To 5) As TVERTEX2
    
    ReDim FlossSkins(0 To 0) As Direct3DTexture8
    ReDim ThatchSkins(0 To 0) As Direct3DTexture8
    Set ThatchSkins(0) = LoadTextureRes(LoadResData(3, "BMP"))
    Set FlossSkins(0) = LoadTextureRes(LoadResData(4, "BMP"))
    Set PatternGridSKin = LoadTextureRes(LoadResData(5, "BMP"))

    Dim bmps As String
    bmps = SearchPath("*.bmp", 1, AppPath & "Base\Stitchings\Overlay", FindAll)
    Do Until bmps = ""
        ThreadCount = ThreadCount + 1
        ReDim Preserve ThreadSkins(0 To ThreadCount) As Direct3DTexture8
        Set ThreadSkins(ThreadCount) = LoadTexture(RemoveNextArg(bmps, vbCrLf))
    Loop
    
    bmps = SearchPath("*.bmp", -1, AppPath & "Base\Stitchings\Mattings", FindAll)
    Do Until bmps = ""
        ThatchCount = ThatchCount + 1
        ReDim Preserve ThatchSkins(0 To ThatchCount) As Direct3DTexture8
        Set ThatchSkins(ThatchCount) = LoadTexture(RemoveNextArg(bmps, vbCrLf))
    Loop
    
    Dim tmpfile As String
    
    bmps = SearchPath("*.bmp", 1, AppPath & "Base\Stitchings\FlossThreads", FindAll)
    Do Until bmps = ""
        FlossCount = FlossCount + 1
        ReDim Preserve FlossSkins(0 To FlossCount) As Direct3DTexture8
        tmpfile = RemoveNextArg(bmps, vbCrLf)
        FlossKeys.Add FlossCount, GetFileTitle(tmpfile)
        Set FlossSkins(FlossCount) = LoadTexture(tmpfile)
    Loop
    

    bmps = SearchPath("*.bmp", 1, AppPath & "Base\Stitchings\LegendKeys", FindAll)
    Do Until bmps = ""
        SymbolCount = SymbolCount + 1
        ReDim Preserve SymbolSkins(0 To SymbolCount) As Direct3DTexture8
        tmpfile = RemoveNextArg(bmps, vbCrLf)
        SymbolKeys.Add SymbolCount, GetFileTitle(tmpfile)
        Set SymbolSkins(SymbolCount) = LoadTexture(tmpfile)
    Loop

    CreateGridPlate SymbolVertex, SymbolVBuf, 1, 1, , (DecalWidth / BlockSqrPerText) * 0.375, (DecalHeight / BlockSqrPerText) * 0.375, 1, 1
    
    CreateGridPlate PanelVertex, PanelVBuf, 1, 1, 1, (DecalWidth / BlockSqrPerText), (DecalHeight / BlockSqrPerText), 1, 1
    CreateGridPlate ThatchVertex, ThatchVBuf, ThatchXUnits, ThatchYUnits, , DecalWidth, DecalHeight, 2, 2
    
    CreateGridPlate Threadings.Vertical.Vertex, Threadings.Vertical.Buffer, 1, 1, , ((DecalWidth / BlockSqrPerText) / ThreadSqrPerText), (DecalHeight / BlockSqrPerText) / 2, 1, 1
    Set Threadings.Vertical.Buffer = DDevice.CreateVertexBuffer(Len(Threadings.Vertical.Vertex(0)) * (UBound(Threadings.Vertical.Vertex) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Threadings.Vertical.Buffer, 0, Len(Threadings.Vertical.Vertex(0)) * (UBound(Threadings.Vertical.Vertex) + 1), 0, Threadings.Vertical.Vertex(0)

    CreateGridPlate Threadings.Horizontal.Vertex, Threadings.Horizontal.Buffer, , , , (DecalWidth / BlockSqrPerText) / 2, ((DecalHeight / BlockSqrPerText) / ThreadSqrPerText), 1, 1
    Set Threadings.Horizontal.Buffer = DDevice.CreateVertexBuffer(Len(Threadings.Horizontal.Vertex(0)) * (UBound(Threadings.Horizontal.Vertex) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Threadings.Horizontal.Buffer, 0, Len(Threadings.Horizontal.Vertex(0)) * (UBound(Threadings.Horizontal.Vertex) + 1), 0, Threadings.Horizontal.Vertex(0)

End Sub

Public Sub CleanUpProj()
    If SymbolCount > 0 Then
        Do While SymbolCount > 0
            Set SymbolSkins(SymbolCount) = Nothing
            ReDim Preserve SymbolSkins(0 To SymbolCount - 1) As Direct3DTexture8
            SymbolCount = SymbolCount - 1
        Loop
        Set SymbolSkins(0) = Nothing
        Erase SymbolSkins
        SymbolCount = 0
    End If
    
    SymbolKeys.Clear
    
    If FlossCount > 0 Then
        Do While FlossCount > 0
            Set FlossSkins(FlossCount) = Nothing
            ReDim Preserve FlossSkins(0 To FlossCount - 1) As Direct3DTexture8
            FlossCount = FlossCount - 1
        Loop
        Set FlossSkins(0) = Nothing
        Erase FlossSkins
        FlossCount = 0
    End If
    
    FlossKeys.Clear
    
    
    If ThatchCount > 0 Then
        Do While ThatchCount > 0
            Set ThatchSkins(ThatchCount) = Nothing
            ReDim Preserve ThatchSkins(0 To ThatchCount - 1) As Direct3DTexture8
            ThatchCount = ThatchCount - 1
        Loop
        Set ThatchSkins(0) = Nothing
        Erase ThatchSkins
        ThatchCount = 0
    End If
    
    Set PatternGridSKin = Nothing
    
    Set Threadings.Vertical.Skin = Nothing
    Set Threadings.Vertical.Buffer = Nothing
    Erase Threadings.Vertical.Vertex
    Set Threadings.Horizontal.Skin = Nothing
    Set Threadings.Horizontal.Buffer = Nothing
    Erase Threadings.Horizontal.Vertex
    Set Threadings.BackSlash.Skin = Nothing
    Set Threadings.BackSlash.Buffer = Nothing
    Erase Threadings.BackSlash.Vertex
    Set Threadings.ForwardSlash.Skin = Nothing
    Set Threadings.ForwardSlash.Buffer = Nothing
    Erase Threadings.ForwardSlash.Vertex
    
    Set ThatchVBuf = Nothing
    Erase ThatchVertex
    
    Set PanelVBuf = Nothing
    Erase PanelVertex
    
End Sub

Private Function BlockWidthX() As Single
    BlockWidthX = ((DecalWidth / BlockSqrPerText) / 2)
End Function
Private Function BlockHeightY() As Single
    BlockHeightY = ((DecalHeight / BlockSqrPerText) / 2)
End Function

Private Function BlockCoordX(ByVal BlockX As Single) As Single
    BlockCoordX = BlockWidthX * BlockX
End Function
Private Function BlockCoordY(ByVal BlockY As Single) As Single
    BlockCoordY = BlockHeightY * BlockY
End Function

Private Property Get ThreadY() As Single
    ThreadY = ((DecalHeight / BlockSqrPerText) / ThreadSqrPerText)
End Property
Private Property Get ThreadX() As Single
    ThreadX = ((DecalWidth / BlockSqrPerText) / ThreadSqrPerText)
End Property

Public Sub RenderView(Optional ByVal TwoDimension As Boolean = False, Optional ByVal HideMOuse As Boolean = False)

    If Not TwoDimension Then

        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_LIGHTING, 1
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        
        DDevice.SetVertexShader FVF_RENDER
    
    Else

        DDevice.SetRenderState D3DRS_ZENABLE, False
        DDevice.SetRenderState D3DRS_LIGHTING, False
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

        DDevice.SetVertexShader FVF_SCREEN
    
    End If

    DDevice.SetPixelShader PixelShaderDefault

    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld

    Dim X As Single
    Dim Y As Single

    Dim cursetX As Single
    Dim cursetY As Single

    Dim matTemp As D3DMATRIX

    Dim offsetx As Single
    Dim offsety As Single
    If Not TwoDimension Then
        offsetx = (BlockCoordX(ThatchXUnits) / 2) - (BlockWidthX * BlockSqrPerText) - BlockWidthX
        offsety = (BlockCoordY(ThatchYUnits) / 2) - (BlockHeightY * BlockSqrPerText) - BlockHeightY
    Else
        offsetx = -24
        offsety = -24
    End If

    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

    If frmStudio.TabStrip1.SelectedItem.Key = "pattern" Then
        DDevice.SetMaterial LucentMat
        DDevice.SetTexture 0, PatternGridSKin
        DDevice.SetMaterial GenericMat
        DDevice.SetTexture 1, PatternGridSKin
    Else
        DDevice.SetMaterial LucentMat
        DDevice.SetTexture 0, ThatchSkins(ThatchIndex)
        DDevice.SetMaterial GenericMat
        DDevice.SetTexture 1, ThatchSkins(ThatchIndex)
    End If

    For Y = (ThatchYUnits - (BlockSqrPerText - 1)) To 0 Step -BlockSqrPerText
        For X = (ThatchXUnits - (BlockSqrPerText - 1)) To 0 Step -BlockSqrPerText

            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + cursetX, -offsety + cursetY, TwoDimension, 256 + 68, 256 + 68
            
            cursetX = cursetX + (BlockWidthX * BlockSqrPerText)
            DDevice.SetTransform D3DTS_WORLD, matWorld
            
            If Not TwoDimension Then
                DDevice.SetStreamSource 0, ThatchVBuf, Len(ThatchVertex(0))
                DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
            Else
                DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
            End If

        Next
        cursetX = 0
        cursetY = cursetY + (BlockHeightY * BlockSqrPerText)
    Next

    If Not TwoDimension Then
        offsetx = (offsetx + ((BlockWidthX * BlockSqrPerText) / (BlockSqrPerText / 2)))
        offsety = (offsety + ((BlockHeightY * BlockSqrPerText) / (BlockSqrPerText / 2)))
    End If
    
    
    
    Dim MouseX As Single
    Dim MouseY As Single

    Dim vDir As D3DVECTOR
    Dim vIntersect As D3DVECTOR

    Dim dViewPort As D3DVIEWPORT8
    DDevice.GetViewport dViewPort

    MouseX = Tan(FOV / 2) * (frmStudio.LastX / ((dViewPort.Width * Screen.TwipsPerPixelX) / 2) - 1) / ASPECT
    MouseY = Tan(FOV / 2) * (1 - frmStudio.LastY / ((dViewPort.Height * Screen.TwipsPerPixelY) / 2))

    Dim p1 As D3DVECTOR 'StartPoint on the nearplane
    Dim p2 As D3DVECTOR 'EndPoint on the farplane

    p1.X = MouseX * NEAR
    p1.Y = MouseY * NEAR
    p1.Z = NEAR

    p2.X = MouseX * FAR
    p2.Y = MouseY * FAR
    p2.Z = FAR

    'Inverse the view matrix
    Dim matInverse As D3DMATRIX
    DDevice.GetTransform D3DTS_VIEW, matView

    D3DXMatrixInverse matInverse, 0, matView

    VectorMatrixMultiply p1, p1, matInverse
    VectorMatrixMultiply p2, p2, matInverse
    D3DXVec3Subtract vDir, p2, p1

    'Check if the points hit
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim v3 As D3DVECTOR

    Dim v4 As D3DVECTOR
    Dim v5 As D3DVECTOR
    Dim v6 As D3DVECTOR

    Dim pPlane1 As D3DVECTOR4

    Dim cnt As Long
    v1.X = ThatchVertex(0).X
    v1.Y = ThatchVertex(0).Y
    v1.Z = ThatchVertex(0).Z

    v2.X = ThatchVertex(1).X
    v2.Y = ThatchVertex(1).Y
    v2.Z = ThatchVertex(1).Z

    v3.X = ThatchVertex(2).X
    v3.Y = ThatchVertex(2).Y
    v3.Z = ThatchVertex(2).Z

    pPlane1 = Create4DPlaneVectorFromPoints(v1, v2, v3)

    Dim c As D3DVECTOR
    Dim N As D3DVECTOR
    Dim P As D3DVECTOR
    Dim V As D3DVECTOR

    Dim hit As Boolean
    LastMouseSetX = MouseSetX
    LastMouseSetY = MouseSetY

    MouseSetX = Round(MouseX, 5)
    MouseSetY = Round(MouseY, 5)

    hit = RayIntersectPlane(pPlane1, p1, vDir, vIntersect)

    LastScreenSetX = ScreenSetX
    LastScreenSetY = ScreenSetY
    ScreenSetX = vIntersect.X
    ScreenSetY = vIntersect.Y

    If hit = True Then
        X = 1
        Y = 1
        hit = False

        If ((vIntersect.X > -offsetx) And (vIntersect.Y > -offsety)) And _
             ((vIntersect.X <= -offsetx + BlockCoordX(ThatchXUnits + 1)) And (vIntersect.Y <= -offsety + BlockCoordY(ThatchYUnits + 1))) Then

            X = vIntersect.X + offsetx
            cursetX = (X Mod BlockWidthX)
            If cursetX < 0 Then cursetX = InvertNum(-cursetX, BlockWidthX) / 2
            offsetx = (X \ BlockWidthX) + IIf(cursetX <> 0, ThreadSize, 0)

            Y = vIntersect.Y + offsety
            cursetY = (Y Mod BlockHeightY)
            If cursetY < 0 Then cursetY = InvertNum(-cursetY, BlockHeightY) / 2
            offsety = (Y \ BlockHeightY) + IIf(cursetY <> 0, ThreadSize, 0)

            hit = True

            If (Not (frmMain.MousePointer = 99)) Then
                frmMain.MousePointer = 99
                frmMain.MouseIcon = LoadPicture(AppPath & "Base\mouse.cur")

            End If
        Else
            If (Not (frmMain.MousePointer = 0)) Then
                frmMain.MousePointer = 0
            End If
            offsetx = 0
            offsety = 0
        End If
    Else
        offsetx = 0
        offsety = 0
    End If

    LastBlockSetX = BlockSetX
    LastBlockSetY = BlockSetY

    LastDetailSetX = DetailSetX
    LastDetailSetY = DetailSetY

    
    If hit And (TrapMouse = 0) Then

        BlockSetX = offsetx
        BlockSetY = offsety
        DetailSetX = cursetX
        DetailSetY = cursetY

        cursetX = ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) + (BlockWidthX * 2)
        cursetY = ((Y \ BlockHeightY) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) + (BlockHeightY * 2)

    Else
        BlockSetX = 0
        BlockSetY = 0
        DetailSetX = 0
        DetailSetY = 0
    End If

    
    Dim i As Byte
    For X = 1 To ThatchXUnits
        For Y = 1 To ThatchYUnits
            'If ProjGrid(X, Y).count > 0 Then
                For i = 1 To ProjGrid(X, Y).count
                    'If ProjGrid(X, Y).Details(i).Stitch > 0 Then
                    
                        DrawStitch X, Y, ProjGrid(X, Y).Details(i).Stitch, ProjGrid(X, Y).Details(i).Color, TwoDimension

                    'End If
                Next
            'End If
        Next
    Next

    If hit And (TrapMouse = 0) Then
        
        If frmStudio.mnuAdvanced.Checked Then
        

            Select Case Abs(frmStudio.VScroll2.Value)
                Case 1 'vertical
                    Select Case ((offsetx - ((offsetx \ BlockWidthX) * BlockWidthX)) \ ThreadX)
                        Case 2
                            DrawHiliteEx cursetX + ThreadX, cursetY - (BlockHeightY / 2)
                        Case Else
                            DrawHiliteEx cursetX - ThreadX, cursetY - (BlockHeightY / 2)
                    End Select

                Case 2 'horizontal
                    DrawHiliteEx cursetX - (BlockWidthX / 2), cursetY
                Case 0, 3, 4, 5 'forwardslash 'backslash 'cross
                    DrawHiliteEx cursetX - (BlockWidthX / 2), cursetY - (BlockHeightY / 2)
            End Select
            

        Else
            Select Case Abs(frmStudio.VScroll2.Value)
                Case 1 'vertical
                    DrawHiliteEx cursetX, cursetY - (BlockHeightY / 2)
                Case 2 'horizontal
                    DrawHiliteEx cursetX - (BlockWidthX / 2), cursetY
                Case 0, 3, 4, 5 'forwardslash 'backslash 'cross
                    DrawHiliteEx cursetX - (BlockWidthX / 2), cursetY - (BlockHeightY / 2)
            End Select
        
        End If
        
        If (Not HideMOuse) And (Not NotFocused) Then

            DDevice.SetTexture 0, MouseSkin
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
            DDevice.SetTexture 1, Nothing
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        End If
    
'        If frmStudio.RemovalTool.Value = 1 Then
'
'            cursetX = (X - ((X \ BlockWidthX) * BlockWidthX))
'            cursetY = (Y - ((Y \ BlockHeightY) * BlockHeightY))
'            If frmStudio.mnuAdvanced.Checked Then
'
'                If cursetX < ThreadX And cursetX > 0 Then
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) + (ThreadX / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2)
'                ElseIf cursetX < ThreadX * 2 And cursetX > ThreadX Then
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) + ((ThreadX / 2) * 3), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2)
'
'                ElseIf cursetY < ThreadY And cursetY > 0 Then
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) + (ThreadY / 2)
'
'                ElseIf cursetY < ThreadY * 2 And cursetY > ThreadY Then
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) + ((ThreadY / 2) * 3)
'
'                Else
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2)
'
'                End If
'
'            Else
'                Select Case Abs(frmStudio.VScroll2.Value)
'                    Case 1 'vertical
'
'                        DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) + (BlockHeightY / 2)
'
'                    Case 2 'horizontal
'
'                        DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) - (BlockWidthX / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2)
'
'                    Case 0, 3, 4, 5 'forwardslash 'backslash 'cross
'                        DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) - (BlockWidthX / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) - (BlockHeightY / 2)
'
'                End Select
'
'            End If
'
'        Else
'            Select Case Abs(frmStudio.VScroll2.Value)
'                Case 1 'vertical
'                    If frmStudio.mnuAdvanced.Checked Then
'
'                    Else
'                        DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) + (BlockHeightY / 2)
'                    End If
'
'                Case 2 'horizontal
'                    If frmStudio.mnuAdvanced.Checked Then
'
'                    Else
'                        DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) - (BlockWidthX / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2)
'                    End If
'
'                Case 0, 3, 4, 5 'forwardslash 'backslash 'cross
'                    DrawHiliteEx ((X \ BlockWidthX) * BlockWidthX) - ((BlockWidthX * ThatchXUnits) / 2) - (BlockWidthX / 2), ((Y \ BlockWidthX) * BlockHeightY) - ((BlockHeightY * ThatchYUnits) / 2) - (BlockHeightY / 2)
'
'            End Select
'        End If


    End If


End Sub
Private Sub SetLocatoins(ByVal locX As Single, ByVal locY As Single, TwoDimension As Boolean, ByVal BlockX As Single, ByVal BlockY As Single, Optional ByVal LineType As Long = 0)

    If Not TwoDimension Then
        D3DXMatrixTranslation matWorld, locX, locY, 0
    Else
        Dim modX As Single
        Dim modY As Single
        
        modX = ((frmMain.Picture1.Width / Screen.TwipsPerPixelX) / ((ThatchXUnits + 0.5) * BlockWidthX))
        modY = ((frmMain.Picture1.Height / Screen.TwipsPerPixelY) / ((ThatchYUnits + 1) * BlockHeightY))
        
        locX = locX * modX
        locY = locY * modY
        BlockX = BlockX * modX
        BlockY = BlockY * modY
        
        Dim totX As Single
        Dim totY As Single
        
        totY = (frmMain.Picture1.Height / Screen.TwipsPerPixelY)
        
        If LineType = 0 Then

            ScreenImgVert(0) = MakeScreen(locX, (locY + BlockY), -1, BlockWidthX * (1 / BlockSqrPerText), BlockHeightY * (1 / BlockSqrPerText))
            ScreenImgVert(1) = MakeScreen((locX + BlockX), (locY + BlockY), -1, 0, BlockHeightY * (1 / BlockSqrPerText))
            ScreenImgVert(2) = MakeScreen(locX, locY, -1, BlockWidthX * (1 / BlockSqrPerText), 0)
            ScreenImgVert(3) = MakeScreen((locX + BlockX), locY, -1, 0, 0)

        ElseIf LineType = 1 Then 'forward
            
            ScreenImgVert(0) = MakeScreen((locX + BlockY + (BlockX / 2)), (locY + BlockY + (BlockX / 2)), -1, 1, 1)
            ScreenImgVert(1) = MakeScreen((locX + BlockX + BlockY + (BlockX / 2)), (locY + BlockY + (BlockX / 2)), -1, 0, 1)
            ScreenImgVert(2) = MakeScreen((locX + (BlockX / 2)), (locY + (BlockX / 2)), -1, 1, 0)
            ScreenImgVert(3) = MakeScreen((locX + BlockX + (BlockX / 2)), (locY + (BlockX / 2)), -1, 0, 0)
            ScreenImgVert(0).Y = (-totY - ScreenImgVert(0).Y) + (totY * 2)
            ScreenImgVert(1).Y = (-totY - ScreenImgVert(1).Y) + (totY * 2)
            ScreenImgVert(2).Y = (-totY - ScreenImgVert(2).Y) + (totY * 2)
            ScreenImgVert(3).Y = (-totY - ScreenImgVert(3).Y) + (totY * 2)
        ElseIf LineType = 2 Then 'black
            ScreenImgVert(0) = MakeScreen((locX + (BlockX / 2)), (locY + BlockY + (BlockX / 2)), -1, 1, 1)
            ScreenImgVert(1) = MakeScreen((locX + BlockX + (BlockX / 2)), (locY + BlockY + (BlockX / 2)), -1, 0, 1)
            ScreenImgVert(2) = MakeScreen((locX + BlockY + (BlockX / 2)), (locY + (BlockX / 2)), -1, 1, 0)
            ScreenImgVert(3) = MakeScreen((locX + BlockX + BlockY + (BlockX / 2)), (locY + (BlockX / 2)), -1, 0, 0)
            ScreenImgVert(0).Y = (-totY - ScreenImgVert(0).Y) + (totY * 2)
            ScreenImgVert(1).Y = (-totY - ScreenImgVert(1).Y) + (totY * 2)
            ScreenImgVert(2).Y = (-totY - ScreenImgVert(2).Y) + (totY * 2)
            ScreenImgVert(3).Y = (-totY - ScreenImgVert(3).Y) + (totY * 2)
        ElseIf LineType = 3 Then 'vertical
            ScreenImgVert(0) = MakeScreen((locX - (BlockY / 2)), (locY + BlockY + (BlockX / 2)), -1, 1, 1)
            ScreenImgVert(1) = MakeScreen((locX + BlockX - (BlockY / 2)), (locY + BlockY + (BlockX / 2)), -1, 0, 1)
            ScreenImgVert(2) = MakeScreen((locX - (BlockY / 2)), (locY + (BlockX / 2)), -1, 1, 0)
            ScreenImgVert(3) = MakeScreen((locX + BlockX - (BlockY / 2)), (locY + (BlockX / 2)), -1, 0, 0)
            ScreenImgVert(0).Y = (-totY - ScreenImgVert(0).Y) + (totY * 2)
            ScreenImgVert(1).Y = (-totY - ScreenImgVert(1).Y) + (totY * 2)
            ScreenImgVert(2).Y = (-totY - ScreenImgVert(2).Y) + (totY * 2)
            ScreenImgVert(3).Y = (-totY - ScreenImgVert(3).Y) + (totY * 2)
        ElseIf LineType = 4 Then 'horizontal
            ScreenImgVert(0) = MakeScreen((locX + (BlockY / 2)), (locY - BlockY - (BlockX / 2)), -1, 1, 1)
            ScreenImgVert(1) = MakeScreen((locX + BlockX + (BlockY / 2)), (locY - BlockY - (BlockX / 2)), -1, 0, 1)
            ScreenImgVert(2) = MakeScreen((locX + (BlockY / 2)), (locY - (BlockX / 2)), -1, 1, 0)
            ScreenImgVert(3) = MakeScreen((locX + BlockX + (BlockY / 2)), (locY - (BlockX / 2)), -1, 0, 0)
            ScreenImgVert(0).Y = (-totY - ScreenImgVert(0).Y) + (totY * 2)
            ScreenImgVert(1).Y = (-totY - ScreenImgVert(1).Y) + (totY * 2)
            ScreenImgVert(2).Y = (-totY - ScreenImgVert(2).Y) + (totY * 2)
            ScreenImgVert(3).Y = (-totY - ScreenImgVert(3).Y) + (totY * 2)
        End If
    End If

End Sub
Public Sub DrawStitch(ByVal BlockX As Single, ByVal BlockY As Single, ByVal ThreadDraw As Long, ByVal ColorDraw As Long, ByVal TwoDimension As Boolean)

    Dim offsetx As Single
    Dim offsety As Single
    If Not TwoDimension Then
        offsetx = (BlockCoordX(ThatchXUnits) / 2) - (BlockWidthX * 2)
        offsety = (BlockCoordY(ThatchYUnits) / 2) - (BlockHeightY * 2)
    End If

'    Dim ThreadY As Single
'    Dim ThreadX As Single
'    ThreadY = ((DecalHeight / BlockSqrPerText) / ThreadSqrPerText)
'    ThreadX = ((DecalWidth / BlockSqrPerText) / ThreadSqrPerText)

    Dim matRot As D3DMATRIX
    Dim matScale As D3DMATRIX
        
    Dim a As Single
    Dim S As Single
    Dim c As Single
    Dim t As Single
    
    If frmStudio.TabStrip1.SelectedItem.Key = "pattern" Then

        D3DXMatrixIdentity matRot
        D3DXMatrixIdentity matScale
        SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY) - (BlockHeightY / 2), TwoDimension, 48, 48
        
        'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY) - (BlockHeightY / 2), 0
        D3DXMatrixScaling matScale, 1, 1, 1
        D3DXMatrixMultiply matWorld, matWorld, matScale
        DDevice.SetTransform D3DTS_WORLD, matWorld
        
        SubDrawSymbol ColorDraw, TwoDimension

    ElseIf frmStudio.mnuAdvanced.Checked Then
        If BitLong(ThreadDraw, StitchBit.LeftEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 1), -offsety + BlockCoordY(BlockY) + (BlockHeightY / 2), TwoDimension, 6, 48, 3
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 1), -offsetY + BlockCoordY(blockY) + (BlockHeightY / 2), 0
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        If BitLong(ThreadDraw, StitchBit.LeftEdgeThick) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - ThreadX, -offsety + BlockCoordY(BlockY) + (BlockHeightY / 2), TwoDimension, 12, 48, 3
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - ThreadX, -offsetY + BlockCoordY(blockY) + (BlockHeightY / 2), 0
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        
        
        If BitLong(ThreadDraw, StitchBit.RightEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX), -offsety + BlockCoordY(BlockY) + (BlockHeightY / 2), TwoDimension, 6, 48, 3
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX), -offsetY + BlockCoordY(blockY) + (BlockHeightY / 2), 0
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        If BitLong(ThreadDraw, StitchBit.RightEdgeThick) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) + ThreadX, -offsety + BlockCoordY(BlockY) + (BlockHeightY / 2), TwoDimension, 12, 48, 3
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) + ThreadX, -offsetY + BlockCoordY(blockY) + (BlockHeightY / 2), 0
    
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        
            
        If BitLong(ThreadDraw, StitchBit.TopEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY) + (ThreadY * 2), TwoDimension, 48, 6, 4
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY) + (ThreadY * 2), 0
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Horizontal, ColorDraw, TwoDimension
        End If
        If BitLong(ThreadDraw, StitchBit.TopEdgeThick) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY) - ThreadY, TwoDimension, 48, 12, 4
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY) - ThreadY, 0
    
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Horizontal, ColorDraw, TwoDimension
        End If
        
        
        If BitLong(ThreadDraw, StitchBit.BottomEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY) + (ThreadY * 1), TwoDimension, 48, 6, 4
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY) + (ThreadY * 1), 0
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Horizontal, ColorDraw, TwoDimension
        End If
        If BitLong(ThreadDraw, StitchBit.BottomEdgeThick) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY), TwoDimension, 48, 12, 4
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY), 0
    
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Horizontal, ColorDraw, TwoDimension
        End If
        
       
        D3DXMatrixIdentity matRot

        D3DXMatrixIdentity matScale

                
        If BitLong(ThreadDraw, StitchBit.ForwardSlashThin) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 4), -offsety + BlockCoordY(BlockY) - (ThreadY * 3), TwoDimension, 8, 48, 1
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 4), -offsetY + BlockCoordY(blockY) - (ThreadY * 3), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, -0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        
        If BitLong(ThreadDraw, StitchBit.ForwardSlashThick) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 3.5), -offsety + BlockCoordY(BlockY) - (ThreadY * 4), TwoDimension, 14, 48, 1
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 3.5), -offsetY + BlockCoordY(blockY) - (ThreadY * 4), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, -0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
    
        If BitLong(ThreadDraw, StitchBit.BackSlashThin) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 3.5), -offsety + BlockCoordY(BlockY) - (ThreadY * 3.5), TwoDimension, 8, 48, 2
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 3.5), -offsetY + BlockCoordY(blockY) - (ThreadY * 3.5), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, 0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
        
        If BitLong(ThreadDraw, StitchBit.BackSlashThick) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 4), -offsety + BlockCoordY(BlockY) - (ThreadY * 4), TwoDimension, 14, 48, 2
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 4), -offsetY + BlockCoordY(blockY) - (ThreadY * 4), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, 0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If
    Else

        If BitLong(ThreadDraw, StitchBit.RightEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX), -offsety + BlockCoordY(BlockY) - (BlockHeightY * 0.5), TwoDimension, 6, 48, 3
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX), -offsetY + BlockCoordY(blockY) - (BlockHeightY * 0.5), 0

            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If

        If BitLong(ThreadDraw, StitchBit.TopEdgeThin) Then
            D3DXMatrixIdentity matWorld
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsety + BlockCoordY(BlockY), TwoDimension, 48, 6, 4
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (BlockWidthX / 2), -offsetY + BlockCoordY(blockY), 0

            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Horizontal, ColorDraw, TwoDimension
        End If


        D3DXMatrixIdentity matRot
        D3DXMatrixIdentity matScale


        If BitLong(ThreadDraw, StitchBit.ForwardSlashThin) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 4), -offsety + BlockCoordY(BlockY) - (ThreadY * 4), TwoDimension, 8, 48, 1
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 4), -offsetY + BlockCoordY(blockY) - (ThreadY * 4), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, -0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If



        If BitLong(ThreadDraw, StitchBit.BackSlashThin) Then
            SetLocatoins -offsetx + BlockCoordX(BlockX) - (ThreadX * 4), -offsety + BlockCoordY(BlockY) - (ThreadY * 4), TwoDimension, 8, 48, 2
            'D3DXMatrixTranslation matWorld, -offsetX + BlockCoordX(BlockX) - (ThreadX * 4), -offsetY + BlockCoordY(blockY) - (ThreadY * 4), 0
            D3DXMatrixRotationYawPitchRoll matRot, 0, 0, 0.75
            D3DXMatrixMultiply matWorld, matRot, matWorld
            D3DXMatrixScaling matScale, 1, 1, 1
            D3DXMatrixMultiply matWorld, matWorld, matScale
            DDevice.SetTransform D3DTS_WORLD, matWorld
            SubDrawStitch Threadings.Vertical, ColorDraw, TwoDimension
        End If

    End If
End Sub
Public Function ColorToHex(ByVal Color As Long) As String
    If 6 - Len(CStr(Hex(Color))) > 0 Then
        ColorToHex = "H" & String(6 - Len(CStr(Hex(Color))), "0") & CStr(Hex(Color))
    Else
        ColorToHex = "H" & CStr(Hex(Color))
    End If
End Function
Private Sub SubDrawStitch(ByRef useWall As RenderWall, ByVal ColorDraw As Long, ByVal TwoDimension As Boolean)
    
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    If Not TwoDimension Then
        DDevice.SetStreamSource 0, useWall.Buffer, Len(useWall.Vertex(0))

    End If
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR

    
    If FlossKeys.Exists(ColorToHex(ColorDraw)) Then
    
       Dim idx As Long
       
       idx = FlossKeys(ColorToHex(ColorDraw))
        
        DDevice.SetTexture 0, useWall.Skin
        If Not TwoDimension Then
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        Else
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
        End If


   '   DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    '   DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

        DDevice.SetTexture 1, FlossSkins(idx)
'        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR + D3DBLEND_SRCALPHA + D3DBLEND_SRCCOLOR
'        If Not TwoDimension Then
'            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'        Else
'            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
'        End If
'
'
'        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA + D3DBLEND_SRCCOLOR
'        If Not TwoDimension Then
'            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'        Else
'            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
'        End If

        DDevice.SetTexture 0, FlossSkins(idx)
        If Not TwoDimension Then
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        Else
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
        End If

    End If
End Sub


Private Sub DrawHilite(ByVal BlockX As Long, ByVal BlockY As Long)

    Dim offsetx As Single
    Dim offsety As Single
    offsetx = (BlockCoordX(ThatchXUnits) / 2)
    offsety = (BlockCoordY(ThatchYUnits) / 2)

    DrawHiliteEx BlockCoordX(BlockX), BlockCoordY(BlockY)

End Sub

Private Sub DrawHiliteEx(ByVal BlockX As Single, ByVal BlockY As Single)

    D3DXMatrixIdentity matWorld
    D3DXMatrixTranslation matWorld, BlockX, BlockY, 0

    DDevice.SetTransform D3DTS_WORLD, matWorld

    If Not DDevice.GetRenderState(D3DRS_ALPHABLENDENABLE) Then DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    If Not DDevice.GetRenderState(D3DRS_ALPHATESTENABLE) Then DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
    DDevice.SetStreamSource 0, PanelVBuf, Len(PanelVertex(0))
    
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVDESTCOLOR And D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCCOLOR And D3DBLEND_DESTALPHA

End Sub



Private Sub SubDrawSymbol(ByVal ColorDraw As Long, ByVal TwoDimension As Boolean)
        
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1

    If Not TwoDimension Then

        DDevice.SetStreamSource 0, SymbolVBuf, Len(SymbolVertex(0))

    End If
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
      '  DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        'DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        
    Dim matID As String
    matID = frmStudio.GetMaterialID(ColorDraw)
    If SymbolKeys.Exists(matID) Then
    
        DDevice.SetTexture 0, SymbolSkins(SymbolKeys(matID))
'
'        If Not TwoDimension Then
'            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'        Else
'            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
'        End If
'
'        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    
        DDevice.SetTexture 1, SymbolSkins(SymbolKeys(matID))
        'DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR + D3DBLEND_SRCALPHA + D3DBLEND_SRCCOLOR
'        If Not TwoDimension Then
'            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'        Else
'            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
'        End If
        
        'DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA + D3DBLEND_SRCCOLOR
        If Not TwoDimension Then
            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        Else
            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
        End If


'        DDevice.SetTexture 0, SymbolSkins(SymbolKeys(matID))
'        If Not TwoDimension Then
'            DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'        Else
'            DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScreenImgVert(0), LenB(ScreenImgVert(0))
'        End If
    End If
    
End Sub

Public Function CountColorUsed(ByVal Color As Long) As Long

    Dim X As Single
    Dim Y As Single
    Dim i As Byte
    For X = 1 To ThatchXUnits
        For Y = 1 To ThatchYUnits
            If ProjGrid(X, Y).count > 0 Then
                For i = 1 To ProjGrid(X, Y).count
                    If ProjGrid(X, Y).Details(i).Stitch > 0 Then
                        If Color = ProjGrid(X, Y).Details(i).Color Then CountColorUsed = CountColorUsed + 1
                    End If
                Next
            End If
        Next
    Next
    
End Function

Public Function RemoveColorUsed(ByVal Color As Long) As Long

    Dim X As Single
    Dim Y As Single
    Dim i As Byte
    For X = 1 To ThatchXUnits
        For Y = 1 To ThatchYUnits
            If ProjGrid(X, Y).count > 0 Then
                For i = 1 To ProjGrid(X, Y).count
                    If ProjGrid(X, Y).Details(i).Stitch > 0 Then
                        If Color = ProjGrid(X, Y).Details(i).Color Then ProjGrid(X, Y).Details(i).Stitch = 0
                    End If
                Next
            End If
        Next
    Next
    modProj.Dirty = True
    
End Function

Public Sub ResizeMultidimArray()
    Dim xcount As Long
    Dim ycount As Long
    Dim N As Long
    Dim xs() As GridItem
    If ThatchXUnits = 0 Or ThatchYUnits = 0 Then
        ThatchBlock = 1
        ThatchXUnits = 100 \ ThatchBlock
        ThatchYUnits = 100 \ ThatchBlock
                
        ThreadSize = 0.2
        BlockSqrPerText = 6
        ThreadSqrPerText = 16

        ReDim ProjGrid(1 To ThatchXUnits, 1 To ThatchYUnits) As GridItem
    Else
    
        xcount = LBound(ProjGrid, 1)
        ycount = LBound(ProjGrid, 2)
        ReDim xs(LBound(ProjGrid, 1) To UBound(ProjGrid, 1), LBound(ProjGrid, 2) To UBound(ProjGrid, 2)) As GridItem
        For xcount = LBound(ProjGrid, 1) To UBound(ProjGrid, 1)
            For ycount = LBound(ProjGrid, 2) To UBound(ProjGrid, 2)
                xs(xcount, ycount) = ProjGrid(xcount, ycount)
            Next
        Next
    
        Erase ProjGrid
        ReDim ProjGrid(1 To ThatchXUnits, 1 To ThatchYUnits) As GridItem

        Dim xcount1 As Long
        Dim ycount1 As Long
    
        xcount1 = xcount
        ycount1 = ycount
        If UBound(ProjGrid, 1) < xcount1 Then xcount1 = UBound(ProjGrid, 1)
        If UBound(ProjGrid, 2) < ycount1 Then ycount1 = UBound(ProjGrid, 2)
        If UBound(xs, 1) < xcount1 Then xcount1 = UBound(xs, 1)
        If UBound(xs, 2) < ycount1 Then ycount1 = UBound(xs, 2)

        For ycount = ycount1 To LBound(ProjGrid, 2) Step -1
            For xcount = xcount1 To LBound(ProjGrid, 1) Step -1
                
                ProjGrid(xcount, ycount) = xs(xcount, ycount)

            Next
        Next
    
        Erase xs
    End If
End Sub

