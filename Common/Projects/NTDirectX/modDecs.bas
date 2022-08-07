Attribute VB_Name = "modDecs"

#Const modDecs = -1
Option Explicit
'TOP DOWN

Option Compare Binary


Public Enum CollisionTypes
    None = 0 'nothing, also hidden objs
    Ranged = 1 'calls oninrange events
    Through = 2 'calls oncollide events
    Freely = 4 'able climb and stayup
    Gravity = 8 'rock gravity with all
    Curbing = 16 'upon move, predicts
    'up small gaps and slopes
    Coupling = 32 'when push comes to
    'shove, couples with other pushes
    Liquid = Freely Or Gravity
    'able climb but also when sitting has lo gravity effect
End Enum

Public Enum CoordinateTypes
    Unspecified = 0
    Absolute = 1
    Relative = 2
End Enum

Public Enum BrilliantTypes
    Directed = 3
    Omni = 1
    Spot = 2
End Enum

Public Enum BillboardTypes
    NotLoaded = 0
    CacheOnly = 1
    HudPanel = 2
    Beacon = 4
End Enum

Public Enum ControllerModes
    Visual = 0 'no mouse conduct
    Hidden = 1 'hidden mouse upon mouse over with focus
    Trapping = 2 'hidden plus the mouse is trappable/untrappable with esc
End Enum

Public Enum PlanetTypes
    Shade = 0 'the space color and fog color, any plane may
    World = 1 'cubic 3d 720 degree panoramic globe atmosphere rendering form
    Plateau = 2 'a single axis value stretch all ways single textured
    Inland = 3 'inland is a plateau with a hole cut out the center
    Island = 4 'opposite inland, the cut is on the plane, and may row/col of h/w
    Screen = 8 'flat face 2d rendering on the screen as a plane
End Enum

Public Enum MotionTypes
    Statue = 0
    Direct = 1
    Rotate = 2
    Scalar = 4
End Enum

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MyScreen
    X As Single
    Y As Single
    z As Single
    rhw As Single
    clr As Long
    tu As Single
    tv As Single
End Type

Public Type MyVertex
    X As Single
    Y As Single
    z As Single
    NX As Single
    NY As Single
    Nz As Single
    tu As Single
    tv As Single
End Type

Public Const FVF_VERTEX_SIZE = 12
Public Const FVF_RENDER_SIZE = 32

Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Const RENDER_MAGFILTER = D3DTEXF_ANISOTROPIC
Public Const RENDER_MINFILTER = D3DTEXF_ANISOTROPIC

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Public Const Transparent As Long = &HFFFF00FF
Public Const MouseSensitivity As Long = 4 'Public Const MouseSensitivity As Long = 1
Public Const MaxDisplacement As Single = 0.05 'Public Const MaxDisplacement As Single = 200
Public Const BeaconSpacing As Single = 40
Public Const BeaconRange As Single = 1000

  'Public Const FadeDistance As Single = 1000000
Public Const SpaceBoundary As Single = 3000
Public Const HoursInOneDay As Single = 24
Public Const LetterPerInch As Single = 10

Public Const MouseZoomInMax As Long = -40000
Public Const MouseZoomOutMax As Long = 40000

Public MaxCameraZoom As Single
Public MinCameraZoom As Single

Public GravityVector As New Point
Public LiquidVector As New Point


Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_F1 = &H70


Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Declare Function CoCreateGuid Lib "ole32" (ByVal pGuid As Long) As Long
Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Const SWW_HPARENT = -8

Public Function ConvertVertexToVector(ByRef V As D3DVERTEX) As D3DVECTOR
    ConvertVertexToVector.X = V.X
    ConvertVertexToVector.Y = V.Y
    ConvertVertexToVector.z = V.z
End Function

Public Function PointInPoly3d(ByRef p As MyVertex, ByRef l() As MyVertex) As Long

    Dim ref1 As Single
    Dim ref2 As Single
    Dim ref3 As Single
    Dim Ret As Single
    Dim f As Long
    f = LBound(l)
    PointInPoly3d = -1
    
    If UBound(l) + IIf(f = 0, 1, 0) > 2 Then
        ref1 = (p.X - l(f).X) * (l(f + 1).Y - l(f).Y) - (p.Y - l(f).Y) * (l(f + 1).X - l(f).X)
        ref2 = (p.Y - l(f).Y) * (l(f + 1).z - l(f).z) - (p.z - l(f).z) * (l(f + 1).Y - l(f).Y)
        ref3 = (p.z - l(f).z) * (l(f + 1).X - l(f).X) - (p.X - l(f).X) * (l(f + 1).z - l(f).z)
   
        Ret = ref1 + ref2 + ref3
        
        Dim i As Long
        For i = f + 1 To UBound(l)
            ref1 = ((p.X - l(f).X) * (l(i).Y - l(f).Y) - (p.Y - l(f).Y) * (l(i).X - l(f).X))
            ref2 = ((p.Y - l(f).Y) * (l(i).z - l(f).z) - (p.z - l(f).z) * (l(i).Y - l(f).Y))
            ref3 = ((p.z - l(f).z) * (l(i).X - l(f).X) - (p.X - l(f).X) * (l(i).z - l(f).z))

            If ((Ret >= 0) Xor ((ref1 + ref2 + ref3) >= 0)) Then
                PointInPoly3d = i
                Exit Function
            End If
            
            Ret = ref1 + ref2 + ref3
        
        Next

    End If
End Function



'Public Sub CreateMesh(ByVal FileName As String, mesh As D3DXMesh, Buffer As D3DXBuffer, MeshMaterials() As D3DMATERIAL8, MeshTextures() As IUnknown, MeshVerticies() As D3DVERTEX, MeshIndicies() As Integer, nMaterials As Long, Optional ByRef SurfaceArea As Single, Optional ByRef Volume As Single)
'    Dim TextureName As String
'
'    Set mesh = D3DX.LoadMeshFromX(FileName, D3DXMESH_DYNAMIC, DDevice, Nothing, Buffer, nMaterials)
'
''    D3DX.CreateMesh
'
'    Dim q As Integer
'
'    If nMaterials > 0 Then
'
'        ReDim MeshMaterials(0 To nMaterials - 1) As D3DMATERIAL8
'        ReDim MeshTextures(0 To nMaterials - 1) As IUnknown
'
'        Dim d As ImageDimensions
'
'        For q = 0 To nMaterials - 1
'
'            D3DX.BufferGetMaterial Buffer, q, MeshMaterials(q)
'
'            TextureName = D3DX.BufferGetTextureName(Buffer, q)
'            If (TextureName <> "") Then
'
'                Set MeshTextures(q) = GetBillboardByFile(GetFilePath(FileName) & "\" & TextureName)
'                If MeshTextures(q) Is Nothing Then
'                    If ImageDimensions(GetFilePath(FileName) & "\" & TextureName, d) Then
'                        Set MeshTextures(q) = D3DX.CreateTextureFromFileEx(DDevice, GetFilePath(FileName) & "\" & TextureName, d.Width, d.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Transparent, ByVal 0, ByVal 0)
'                    Else
'                        Debug.Print "IMAGE ERROR: ImageDimensions - " & GetFilePath(FileName) & "\" & TextureName
'                    End If
'                End If
'            End If
'
'        Next
'    Else
'        ReDim MeshTextures(0 To 0) As IUnknown
'        ReDim MeshMaterials(0 To 0) As D3DMATERIAL8
'    End If
'
'    D3DX.ComputeNormals mesh
'
'    Dim vd As D3DVERTEXBUFFER_DESC
'    mesh.GetVertexBuffer.GetDesc vd
'
'    ReDim MeshVerticies(0 To ((vd.size \ FVF_VERTEX_SIZE) - 1)) As D3DVERTEX
'    D3DVertexBuffer8GetData mesh.GetVertexBuffer, 0, vd.size, 0, MeshVerticies(0)
'
'    Dim ID As D3DINDEXBUFFER_DESC
'    mesh.GetIndexBuffer.GetDesc ID
'
'    ReDim MeshIndicies(0 To ((ID.size \ 2) - 1)) As Integer
'    D3DIndexBuffer8GetData mesh.GetIndexBuffer, 0, ID.size, 0, MeshIndicies(0)
'
'    D3DX.ComputeNormals mesh
'
'    If nMaterials > 0 Then
'
'        Dim l3 As Single
'        Dim l1 As Single
'        Dim l2 As Single
'        Dim l4 As Single
'        Dim l5 As Single
'        Dim l6 As Single
'
'        Dim Index As Long
'        Dim checked As Long
'
'        Index = 0 'start at last point of first triangle where start = 0
'        Do Until checked = mesh.GetNumFaces / 2 'go for amount of faces least 3
'
'             If l1 = 0 Then
'                l1 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                l4 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'            End If
'
'             If l2 = 0 Then
'                l2 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'                l5 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'            End If
'
'             If l3 = 0 Then
'                l3 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                l6 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'            End If
'
'             Index = Index + 1
'
'             If l1 = l2 And l2 = l3 And l1 <> 0 Then
'
'                 l2 = 0
'                 l6 = l4
'                 l4 = l5
'                 l5 = 0
'             Else
'
'                If l1 <> 0 And l2 <> 0 And l3 <> 0 Then
'
'                    SurfaceArea = SurfaceArea + (TriangleAreaByLen(l1, l2, l3) + TriangleAreaByLen(l4, l5, l6))
'
'                    Volume = Volume + (TriangleVolByLen(l1, l2, l3) + TriangleVolByLen(l4, l5, l6))
'
'                    l1 = 0
'                    l2 = 0
'                    l3 = 0
'                    l4 = 0
'                    l5 = 0
'                    l6 = 0
'                    checked = checked + 2
'                End If
'
'                Index = Index + 2
'            End If
'
'        Loop
'
'    End If
'
'    SurfaceArea = (SurfaceArea * 2)
'
'End Sub


Public Function TriangleAreaByLen(ByVal l1 As Single, ByVal l2 As Single, ByVal l3 As Single) As Single
    TriangleAreaByLen = (((((((l1 + l2) - l3) + ((l2 + l3) - l1) + ((l3 + l1) - l2)) * (l1 * l2 * l3)) / (l1 + l2 + l3)) ^ (1 / 2)))
End Function
Public Function TriangleVolByLen(ByVal l1 As Single, ByVal l2 As Single, ByVal l3 As Single) As Single
    TriangleVolByLen = TriangleAreaByLen(l1, l2, l3)
    TriangleVolByLen = ((((TriangleVolByLen ^ (1 / 3)) ^ 2) ^ 3) / 12)
End Function


'Public Function ScreenXYTo3DZ0(ByVal MouseX As Single, ByVal MouseY As Single) As D3DVECTOR
'    Dim Alpha As Single
'
'    Dim vDir As D3DVECTOR
'    Dim vIntersect As D3DVECTOR
'
'    Dim halfwidth As Long
'    Dim halfheight As Long
'
'    halfwidth = Screen.Width / 2
'    halfheight = Screen.Height / 2
'
'    Alpha = FOVY / 2
'
'    'Screen Width is 640x480
'    MouseX = Tan(Alpha) * (frmMain.LastX / halfwidth - 1) / ASPECT
'    MouseY = Tan(Alpha) * (1 - frmMain.LastY / halfheight)
'
'    'Intersect the plane from the eye to the centre of the plane
'
'    Dim p1 As D3DVECTOR 'StartPoint on the nearplane
'    Dim p2 As D3DVECTOR 'EndPoint on the farplane
'
'    p1.X = MouseX * NEAR
'    p1.Y = MouseY * NEAR
'    p1.Z = NEAR
'
'    p2.X = MouseX * FAR
'    p2.Y = MouseY * FAR
'    p2.Z = FAR
'
'    'Inverse the view matrix
'    Dim matInverse As D3DMATRIX
'    DDevice.GetTransform D3DTS_VIEW, matView
'
'    D3DXMatrixInverse matInverse, 0, matView
'
'    VectorMatrixMultiply p1, p1, matInverse
'    VectorMatrixMultiply p2, p2, matInverse
'    D3DXVec3Subtract vDir, p2, p1
'
'    'Check if the points hit
'    Dim v1 As D3DVECTOR
'    Dim v2 As D3DVECTOR
'    Dim v3 As D3DVECTOR
'
'    Dim V4 As D3DVECTOR
'    Dim v5 As D3DVECTOR
'    Dim v6 As D3DVECTOR
'
'    Dim pPlane1 As D3DVECTOR4
'
'    Dim cnt As Long
'    v1.X = -100000
'    v1.Y = -100000
'    v1.Z = Camera.Player.Offset.Z
'
'    v2.X = -100000
'    v2.Y = 100000
'    v2.Z = Camera.Player.Offset.Z
'
'    v3.X = 100000
'    v3.Y = 100000
'    v3.Z = Camera.Player.Offset.Z
'
'    pPlane1 = Create4DPlaneVectorFromPoints(v1, v2, v3)
'
'    Dim c As D3DVECTOR
'    Dim n As D3DVECTOR
'    Dim p As D3DVECTOR
'    Dim V As D3DVECTOR
'
'    Dim hit As Boolean
'   ' vIntersect.X = (vIntersect.X * (ThatchSquare / ThatchScale))
'   ' vIntersect.Y = (vIntersect.Y * (ThatchSquare / ThatchScale))
'
'    If RayIntersectPlane(pPlane1, p1, vDir, vIntersect) Then
'        ScreenXYTo3DZ0 = vIntersect
'    End If
'End Function

'Public Function GetCursorPosEx(ByVal hWnd As Long, ByRef pt As POINTAPI) As Long
'
'    Dim bsize As Long
'    bsize = GetParent(hWnd)
'    hWnd = IIf(bsize > 0, bsize, hWnd)
'    If hWnd<> 0 Then
'        If GetCursorPos(pt) Then
'
'            pt.X = (pt.X * VB.Screen.TwipsPerPixelX) 'convert to screen twips subtract the
'            'left for reletive position coord match from the mouse x coord of designer
'
'            bsize = ((frmStudio.Designer.Width - frmStudio.Designer.ScaleWidth) / 2) 'get a single
'            bsize = (frmStudio.Designer.Left + bsize) ' width of the forms side edge, add it to the
'            pt.X = pt.X - bsize 'forms left on screen subtract that value from the mouse cursor
'            bsize = ((frmStudio.Width - frmStudio.ScaleWidth) / 2) 'do the same with the designer
'            pt.X = pt.X - (frmStudio.Left + bsize) 'subtract that value from the mouse cursor
'
'
'            bsize = (bsize / VB.Screen.TwipsPerPixelX) 'size of one side border in pixel, we are using the width
'            'because of the title bar is alone on the height with two borders similar to width but with title
'            bsize = (bsize * VB.Screen.TwipsPerPixelY) 'convert to pixels of a single size of border into twips of Y axis
'            bsize = (frmStudio.Height - frmStudio.ScaleHeight) - (bsize * 2) 'subtract two of them from the full heights
'            'difference in scalar (internal coordinate system) and it's true value (the external edge coordinate system)
'            'left with just title bar height in twips, between true value and scale coords is broders/title bar sizes
'
'            'same thing we did to the width before, but incorporate the title subtraction to locate absolute designer posotion on screen
'            pt.Y = (pt.Y * VB.Screen.TwipsPerPixelY) - (frmStudio.Top + bsize + (((frmStudio.Height - frmStudio.ScaleHeight) - bsize) / 2) + _
'                    (frmStudio.Designer.Top + ((frmStudio.Designer.Height - frmStudio.Designer.ScaleHeight) / 2)))
'
'            GetCursorPosEx = hWnd
'
'        End If
'    End If
'
'End Function

Function FloatToDWord(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FloatToDWord = l
End Function

Public Function GetUserLoginName() As String

    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = String(255, Chr(0))
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left(sBuffer, lSize)
    End If
    sBuffer = Replace(sBuffer, Chr(0), "")
    GetUserLoginName = sBuffer
    
End Function
Public Function Clamp(ByVal Value As Single, ByVal max As Single, ByVal min As Single) As Single
    If Value > max Then
        Clamp = max
    End If
    If Value < min Then
        Clamp = min
    End If
End Function

Public Function LengthSqr(ByRef V As D3DVECTOR) As Single
    LengthSqr = Sqr(Distance(0, 0, 0, V.X, V.Y, V.z))
End Function

Public Function CreateVertex(X As Single, Y As Single, z As Single, NX As Single, NY As Single, Nz As Single, tu As Single, tv As Single) As MyVertex
    
    With CreateVertex
        .X = X: .Y = Y: .z = z
        .NX = NX: .NY = NY: .Nz = Nz
        .tu = tu: .tv = tv
    End With
    
End Function


'Public Function PointToPlane(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
'    Dim p3 As D3DVECTOR
'    p3 = VectorSubtract(p1, p2)
'    PointToPlane = Sqr(VectorDotProduct(p3, p3))
'End Function



'Public Function InterceptPoint(ByRef l1p1 As D3DVECTOR, ByRef l1p2 As D3DVECTOR, ByRef l2p1 As D3DVECTOR, ByRef l2p2 As D3DVECTOR, ByRef r1 As D3DVECTOR, ByRef r2 As D3DVECTOR) As Boolean
'    ' // Algorithm is ported from the C algorithm of (but third port by Nick to vb)
'    '// Paul Bourke at http://local.wasp.uwa.edu.au/~pbourke/geometry/lineline3d/
'    r1 = VectorSubtract(l1p1, l2p1)
'    r2 = VectorSubtract(l2p2, l2p1)
'    If Not (LengthSqr(r2) < Epsilon) Then
'        Dim p21 As D3DVECTOR
'        p21 = VectorSubtract(l1p2, l1p1)
'        If Not (LengthSqr(p21) < Epsilon) Then
'            Dim d1343 As Single
'            Dim d4321 As Single
'            Dim d1321 As Single
'            Dim d4343 As Single
'            Dim d2121 As Single
'            Dim denom As Single
'            Dim numer As Single
'            d1343 = r1.X * r2.X + r1.Y * r2.Y + r1.Z * r2.Z
'            d4321 = r2.X * p21.X + r2.Y * p21.Y + r2.Z * p21.Z
'            d1321 = r1.X * p21.X + r1.Y * p21.Y + r1.Z * p21.Z
'            d4343 = r2.X * r2.X + r2.Y * r2.Y + r2.Z * r2.Z
'            d2121 = p21.X * p21.X + p21.Y * p21.Y + p21.Z * p21.Z
'            denom = d2121 * d4343 - d4321 * d4321
'            If Not (Abs(denom) < Epsilon) Then
'                numer = d1343 * d4321 - d1321 * d4343
'                numer = numer / denom
'                denom = (d1343 + d4321 * (numer)) / d4343
'                r1.X = (l1p1.X + numer * p21.X)
'                r1.Y = (l1p1.Y + numer * p21.Y)
'                r1.Z = (l1p1.Z + numer * p21.Z)
'                r2.X = (l2p1.X + denom * r2.X)
'                r2.Y = (l2p1.Y + denom * r2.Y)
'                r2.Z = (l2p1.Z + denom * r2.Z)
'                InterceptPoint = True
'            End If
'        End If
'    End If
'End Function
'


Public Function GetNow() As String

    GetNow = CStr(Now)

End Function

Public Function GetTimer() As String

    GetTimer = CStr(CDbl(Timer))
    
End Function

Public Sub Swap(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

'Public Sub CreateGridPlate(ByRef Data() As MyVertex, ByRef Verticies As Direct3DVertexBuffer8, Optional ByVal Col As Long = 1, Optional ByVal row As Long = 1, Optional ByRef Count As Long = -1, Optional ByVal PlateWidth As Single = 1, Optional ByVal PlateHeight As Single = 1, Optional ByVal tv As Single, Optional ByVal tu As Single)
'    'creates a number of square plates n, is of either row*col or count which ever is greater, count can be existing already and elements are added modifying it
'    Dim r As Long
'    Dim N As Long
'    r = row 'backup for later
'    If (row * Col) > Count Then Count = (row * Col)
'    ReDim Preserve Data(0 To Abs(IIf((row * Col) = 0, 1, (row * Col)) * 6) - 1) As MyVertex
'    N = IIf(Count > 0 And Not Count > (row * Col), (row * Col) - Count, Count)
'    For Count = (((N - 1) \ 6) + -CInt(CBool((N - 1) Mod 6 > 0))) * 6 To UBound(Data) Step 6
'        CreateSquare Data, Count, MakeVector((PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), (PlateHeight / 2), 0), _
'            MakeVector((PlateWidth / 2), (PlateHeight / 2), 0), tv, tu
'        row = row - 1
'        If row = 0 Then row = r
'        If row = r Then Col = Col - 1
'    Next
'    Count = Count \ 6
'    Set Verticies = DDevice.CreateVertexBuffer(Len(Data(0)) * (UBound(Data) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Verticies, 0, Len(Data(0)) * (UBound(Data) + 1), 0, Data(0)
'
'End Sub
'
'Public Sub CreateGridPlateEx(ByRef Data() As MyVertex, ByRef Verticies As Direct3DVertexBuffer8, Optional ByVal Col As Long = 1, Optional ByVal row As Long = 1, Optional ByRef Count As Long = -1, Optional ByVal PlateWidth As Single = 1, Optional ByVal PlateHeight As Single = 1, Optional ByVal tv As Single, Optional ByVal tu As Single, Optional ByVal StreamFlag As Long)
'    'creates a number of square plates n, is of either row*col or count which ever is greater, count can be existing already and elements are added modifying it
'    Dim r As Long
'    Dim N As Long
'    r = row 'backup for later
'    If (row * Col) > Count Then Count = (row * Col)
'    ReDim Preserve Data(0 To Abs(IIf((row * Col) = 0, 1, (row * Col)) * 6) - 1) As MyVertex
'    N = IIf(Count > 0 And Not Count > (row * Col), (row * Col) - Count, Count)
'    For Count = (((N - 1) \ 6) + -CInt(CBool((N - 1) Mod 6 > 0))) * 6 To UBound(Data) Step 6
'        CreateSquareEx Data, Count, MakeVector((PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), (PlateHeight / 2), 0), _
'            MakeVector((PlateWidth / 2), (PlateHeight / 2), 0), tv, tu, StreamFlag
'        row = row - 1
'        If row = 0 Then row = r
'        If row = r Then Col = Col - 1
'    Next
'    Count = Count \ 6
'    Set Verticies = DDevice.CreateVertexBuffer(Len(Data(0)) * (UBound(Data) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Verticies, 0, Len(Data(0)) * (UBound(Data) + 1), 0, Data(0)
'
'End Sub

'Public Function CreateSquare(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef P4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Double
'
'    Dim Mesh As D3DXMesh
'    Set Mesh = D3DX.CreateMeshFVF(2, 6, D3DXMESH_DYNAMIC, FVF_RENDER, DDevice)
'
'   ' mesh.Draw
'
'    Dim vn As New Point
'    Dim a As Double
'
'
'    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
'    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
'    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
'    Set vn = TriangleNormal(MakePoint(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
'                            MakePoint(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
'                            MakePoint(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
'    Data(Index + 0).NX = vn.X: Data(Index + 0).NY = vn.Y: Data(Index + 0).Nz = vn.Z
'    Data(Index + 1).NX = vn.X: Data(Index + 1).NY = vn.Y: Data(Index + 1).Nz = vn.Z
'    Data(Index + 2).NX = vn.X: Data(Index + 2).NY = vn.Y: Data(Index + 2).Nz = vn.Z
'
'    CreateSquare = TriangleSurfaceArea(p1, p2, p3) '(Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
'
'
'    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
'    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
'    Data(Index + 5) = CreateVertex(P4.X, P4.Y, P4.Z, 0, 0, 0, 0, 0)
'    vn = TriangleNormal(MakePoint(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
'                            MakePoint(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
'                            MakePoint(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
'    Data(Index + 3).NX = vn.X: Data(Index + 3).NY = vn.Y: Data(Index + 3).Nz = vn.Z
'    Data(Index + 4).NX = vn.X: Data(Index + 4).NY = vn.Y: Data(Index + 4).Nz = vn.Z
'    Data(Index + 5).NX = vn.X: Data(Index + 5).NY = vn.Y: Data(Index + 5).Nz = vn.Z
'
'    CreateSquare = TriangleSurfaceArea(p3, p1, P4) ' (Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
'
'
'End Function

'Public Function CreateCircle(ByRef Data() As MyVertex, ByVal OuterRadii As Single, ByVal Segments As Single, Optional ByVal InnerRadii As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Double
''creates a tiangle list for a circle of segment amount of straight edges arranged to a circle
''of segment amount of triangles, when innerradii is supplied in segment*2 amount of trianlges
'
'    Dim vn As D3DVECTOR
'    Dim start As Integer
'    start = UBound(Data) + 1
'    ReDim Preserve Data(0 To start + ((IIf(InnerRadii > 0, 6, 3) * Segments) - 1) + (IIf(InnerRadii > 0, 6, 3) * 2)) As MyVertex
'
'    Dim i As Long
'    Dim g As Single
'    Dim a As Double
'    Dim l1 As Single
'    Dim l2 As Single
'    Dim l3 As Single
'
'    Dim intX1 As Single
'    Dim intX2 As Single
'    Dim intX3 As Single
'    Dim intX4 As Single
'
'    Dim intY1 As Single
'    Dim intY2 As Single
'    Dim intY3 As Single
'    Dim intY4 As Single
'    Dim dist1 As Single
'    Dim dist2 As Single
'    Dim dist3 As Single
'    Dim dist4 As Single
'
'    For i = -IIf(InnerRadii > 0, 6, 3) To UBound(Data) - start Step IIf(InnerRadii > 0, 6, 3)
'
'        g = (((360 / Segments) * (((i + 1) / IIf(InnerRadii > 0, 6, 3)) - 1)) * RADIAN)
'
'        intX2 = (OuterRadii * Sin(g))
'        intY2 = (-OuterRadii * Cos(g))
'
'        intX3 = (InnerRadii * Sin(g))
'        intY3 = (-InnerRadii * Cos(g))
'
'        If i >= 0 Then  'skip ahead
'
'            If (InnerRadii > 0) Then
'                If (i Mod 12) = 0 Then
'
'                    dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) * -Sin(g)
'                    dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) * -Cos(g)
'
'                    dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) * -Sin(g)
'                    dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) * -Cos(g)
'
'                End If
'
'                Data(start + i + 0) = CreateVertex(intX2, 0, intY2, 0, 0, 0, 0, dist2)
'                Data(start + i + 1) = CreateVertex(intX1, 0, intY1, 0, 0, 0, dist1, dist4)
'                Data(start + i + 2) = CreateVertex(intX4, 0, intY4, 0, 0, 0, dist3, 0)
'                vn = TriangleNormal(MakeVector(Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z), _
'                                    MakeVector(Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z), _
'                                    MakeVector(Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z))
'                Data(start + i + 0).NX = vn.X: Data(start + i + 0).NY = vn.Y: Data(start + i + 0).Nz = vn.Z
'                Data(start + i + 1).NX = vn.X: Data(start + i + 1).NY = vn.Y: Data(start + i + 1).Nz = vn.Z
'                Data(start + i + 2).NX = vn.X: Data(start + i + 2).NY = vn.Y: Data(start + i + 2).Nz = vn.Z
'
'                l1 = Distance(Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z, Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z)
'                l2 = Distance(Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z, Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z)
'                l3 = Distance(Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z, Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z)
'              '  If l1 <> 0 And l2 <> 0 And l3 <> 0 Then CreateCircle = CreateCircle + TriangleSurfaceArea(l1 , l2, l3) '(Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
'
'                Data(start + i + 3) = CreateVertex(intX2, 0, intY2, 0, 0, 0, 0, dist2)
'                Data(start + i + 4) = CreateVertex(intX4, 0, intY4, 0, 0, 0, dist3, 0)
'                Data(start + i + 5) = CreateVertex(intX3, 0, intY3, 0, 0, 0, 0, 0)
'                vn = TriangleNormal(MakeVector(Data(start + i + 3).X, Data(start + i + 3).Y, Data(start + i + 3).Z), _
'                                    MakeVector(Data(start + i + 4).X, Data(start + i + 4).Y, Data(start + i + 4).Z), _
'                                    MakeVector(Data(start + i + 5).X, Data(start + i + 5).Y, Data(start + i + 5).Z))
'                Data(start + i + 3).NX = vn.X: Data(start + i + 3).NY = vn.Y: Data(start + i + 3).Nz = vn.Z
'                Data(start + i + 4).NX = vn.X: Data(start + i + 4).NY = vn.Y: Data(start + i + 4).Nz = vn.Z
'                Data(start + i + 5).NX = vn.X: Data(start + i + 5).NY = vn.Y: Data(start + i + 5).Nz = vn.Z
'
'                l1 = Distance(Data(start + i + 3).X, Data(start + i + 3).Y, Data(start + i + 3).Z, Data(start + i + 4).X, Data(start + i + 4).Y, Data(start + i + 4).Z)
'                l2 = Distance(Data(start + i + 4).X, Data(start + i + 4).Y, Data(start + i + 4).Z, Data(start + i + 5).X, Data(start + i + 5).Y, Data(start + i + 5).Z)
'                l3 = Distance(Data(start + i + 5).X, Data(start + i + 5).Y, Data(start + i + 5).Z, Data(start + i + 3).X, Data(start + i + 3).Y, Data(start + i + 3).Z)
'               ' If l1 <> 0 And l2 <> 0 And l3 <> 0 Then CreateCircle = CreateCircle + TriangleSurfaceArea(l1, l2, l3) '(Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
'
'            Else
'
'                dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * Sin(g)
'                dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * Cos(g)
'
'                Data(start + i + 0) = CreateVertex(intX2, 0, intY2, 0, 0, 0, ((ScaleX / dist1) * (dist1 / 100)), 0)
'                Data(start + i + 1) = CreateVertex(intX1, 0, intY1, 0, 0, 0, 0, ((ScaleY / dist2) * (dist2 / 100)))
'                Data(start + i + 2) = CreateVertex(intX4, 0, intY4, 0, 0, 0, 0, 0)
'                vn = TriangleNormal(MakeVector(Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z), _
'                                    MakeVector(Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z), _
'                                    MakeVector(Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z))
'                Data(start + i + 0).NX = vn.X: Data(start + i + 0).NY = vn.Y: Data(start + i + 0).Nz = vn.Z
'                Data(start + i + 1).NX = vn.X: Data(start + i + 1).NY = vn.Y: Data(start + i + 1).Nz = vn.Z
'                Data(start + i + 2).NX = vn.X: Data(start + i + 2).NY = vn.Y: Data(start + i + 2).Nz = vn.Z
'
'                l1 = Distance(Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z, Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z)
'                l2 = Distance(Data(start + i + 1).X, Data(start + i + 1).Y, Data(start + i + 1).Z, Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z)
'                l3 = Distance(Data(start + i + 2).X, Data(start + i + 2).Y, Data(start + i + 2).Z, Data(start + i + 0).X, Data(start + i + 0).Y, Data(start + i + 0).Z)
'              '  If l1 <> 0 And l2 <> 0 And l3 <> 0 Then CreateCircle = CreateCircle + TriangleSurfaceArea(l1, l2, l3) '(Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
'
'            End If
'
'        End If
'
'        intX1 = intX2
'        intY1 = intY2
'        intX4 = intX3
'        intY4 = intY3
'
'    Next
'
'End Function
''
''Public Function CreateSquareEx(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1, Optional ByVal StreamFlag As Long) As Double
''
''    Static BrushIndex As Long
''    Dim FaceIndex As Long
''    Dim a As Double
''    Dim l1 As Single
''    Dim l2 As Single
''    Dim l3 As Single
''    Dim vn As D3DVECTOR
''
''    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
''    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
''    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
''
''    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
''                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
''                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
''
''    Data(Index + 0).NX = vn.X: Data(Index + 0).NY = vn.Y: Data(Index + 0).Nz = vn.Z
''    Data(Index + 1).NX = vn.X: Data(Index + 1).NY = vn.Y: Data(Index + 1).Nz = vn.Z
''    Data(Index + 2).NX = vn.X: Data(Index + 2).NY = vn.Y: Data(Index + 2).Nz = vn.Z
''
''    AddTriangleFaceToCollision BrushIndex, FaceIndex, vn, MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
''                                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
''                                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z), StreamFlag
''
''
''    l1 = Distance(p1.X, p1.Y, p1.Z, p2.X, p2.Y, p2.Z)
''    l2 = Distance(p2.X, p2.Y, p2.Z, p3.X, p3.Y, p3.Z)
''    l3 = Distance(p3.X, p3.Y, p3.Z, p1.X, p1.Y, p1.Z)
''    CreateSquareEx = (Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
''
''
''    FaceIndex = FaceIndex + 1
''    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
''    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
''    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, 0, 0)
''
''    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
''                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
''                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
''
''    Data(Index + 3).NX = vn.X: Data(Index + 3).NY = vn.Y: Data(Index + 3).Nz = vn.Z
''    Data(Index + 4).NX = vn.X: Data(Index + 4).NY = vn.Y: Data(Index + 4).Nz = vn.Z
''    Data(Index + 5).NX = vn.X: Data(Index + 5).NY = vn.Y: Data(Index + 5).Nz = vn.Z
''
''    AddTriangleFaceToCollision BrushIndex, FaceIndex, vn, MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
''                                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
''                                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z), StreamFlag
''
''    l1 = Distance(p3.X, p3.Y, p3.Z, p4.X, p4.Y, p4.Z)
''    l2 = Distance(p4.X, p4.Y, p4.Z, p1.X, p1.Y, p1.Z)
''    l3 = Distance(p1.X, p1.Y, p1.Z, p3.X, p3.Y, p3.Z)
''    CreateSquareEx = CreateSquareEx + (Sqr((((l1 ^ 2) + (l2 ^ 2) + (l3 ^ 2)) ^ 2) - 2 * ((l1 ^ 4) + (l2 ^ 4) + (l3 ^ 4))) * 0.25)
''
''    BrushIndex = BrushIndex + 1
''End Function
''
''
''
'


