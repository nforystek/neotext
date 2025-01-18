Attribute VB_Name = "modDecs"

#Const modDecs = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public Enum Scalars
    Millimeters = 0
    Inches = 1
    Centimeters = 2
End Enum

'Public Enum Stitches
'    Vertical = 1
'    Horizontal = 2
'    ForwardSlash = 3
'    BackSlash = 4
'End Enum
'
'Public Enum StitchNum
'    Vertical = 1 And 2 And 5 And 6
'    Horizontal = 3 And 4 And 7 And 8
'    ForwardSlash = 9 And 10
'    BackSlash = 11 And 12
'End Enum
'
'Public Enum StitchBit
'    Vertical = Bit01 Or Bit02 Or Bit05 Or Bit06
'    Horizontal = Bit03 Or Bit04 Or Bit07 Or Bit08
'    ForwardSlash = Bit09 Or Bit10
'    BackSlash = Bit11 Or Bit12
'
'    LeftEdgeThin = Bit01
'    LeftEdgeThick = Bit02
'
'    TopEdgeThin = Bit03
'    TopEdgeThick = Bit04
'
'    RightEdgeThin = Bit05
'    RightEdgeThick = Bit06
'
'    BottomEdgeThin = Bit07
'    BottomEdgeThick = Bit08
'
'    ForwardSlashThin = Bit09
'    ForwardSlashThick = Bit10
'
'    BackSlashThin = Bit11
'    BackSlashThick = Bit12
'
'End Enum

Public Enum Stitches
    Vertical = 1
    Horizontal = 2
    ForwardSlash = 3
    BackSlash = 4
End Enum

Public Enum StitchNum
    LeftEdgeThin = 1
    LeftEdgeThick = 2
   
    TopEdgeThin = 3
    TopEdgeThick = 4
    
    RightEdgeThin = 5
    RightEdgeThick = 6
    
    BottomEdgeThin = 7
    BottomEdgeThick = 8
    
    ForwardSlashThin = 9
    ForwardSlashThick = 10
    
    BackSlashThin = 11
    BackSlashThick = 12

End Enum

Public Enum StitchBit
    LeftEdgeThin = Bit01
    LeftEdgeThick = Bit02
    
    TopEdgeThin = Bit03
    TopEdgeThick = Bit04
    
    RightEdgeThin = Bit05
    RightEdgeThick = Bit06
    
    BottomEdgeThin = Bit07
    BottomEdgeThick = Bit08
    
    ForwardSlashThin = Bit09
    ForwardSlashThick = Bit10
    
    BackSlashThin = Bit11
    BackSlashThick = Bit12

End Enum

Public Type ItemDetail
    Stitch As Long
    Color As Long
End Type

Public Type GridItem
    count As Byte
    Details() As ItemDetail
End Type

#If Not modGraphics = -1 Then
Public Type POINTAPI
    X As Long
    Y As Long
End Type
#End If

Public Type TVERTEX1
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    Color As Long
    tu As Single
    tv As Single
End Type

Public Type TVERTEX2
    X    As Single
    Y    As Single
    Z    As Single
    NX   As Single
    NY   As Single
    Nz   As Single
    tu   As Single
    tv   As Single
End Type

Public Type UserType
    Location As D3DVECTOR
    Rotation As Single
    CameraAngle As Single
    CameraPitch As Single
    CameraZoom As Single
    MoveSpeed As Single
    AutoMove As Boolean
End Type

Public Const RENDER_MAGFILTER = D3DTEXF_ANISOTROPIC
Public Const RENDER_MINFILTER = D3DTEXF_ANISOTROPIC

Public Const FVF_VERTEX_SIZE = 12
Public Const FVF_RENDER_SIZE = 32

Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Const PI As Double = 3.14159265358979
Public Const epsilon As Double = 0.999999999999999
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2
Public Const ASPECT As Single = 0.82
Public Const FOV = 0.7853981633975
Public Const NEAR = 0.01
Public Const FAR = 1000000

Public Const FOVY As Single = 0.8 '1.047198



''initial file header
'Public Type GridInfo '18 bytes
'    Height As Integer 'number of blocks in height the canvas is
'    Scalar As Integer 'scalar width ratio to blocks sizes below
'    Blocks As Integer 'proportional metric of scalar height cubed
'    Metric As Byte '0=millimeters, 1=inches, 2=centimeters
'    Thatch As Byte 'index to the rendering material or matting #
'    'below starts related topics of other records not just header
'    GotoYX As Integer 'starting y block of used pattern (ratio)
'    Ratios As Integer 'number count of blocks serial sequenced
'    JumpXY As Integer 'starting x block of used pattern (width)
'
'    'below is only in ram, not part of record
'    HeightAbs As Single
'    WidthAbs As Single
'    Resolute As Single
'
'    'immediate after last grid comes
'    'symbols by numerical ordinal in
'    StyleSheet As Long 'number of symbols
'    '1,1,1,1, until 0, is count 1 then
'    '0,0, until 1 is count 2 and so on
'    'ordinal by refernce, so not dependant
'    'but certianly want to refence linked
'    'basis of the profile and/or artist
'End Type
'
'Public Type RelateInfo '12 bytes 'after header
'    'for kadpatch this is profile and/or artist
'    'identifiers to match symbols and material
'    'of providers or creators to their relate
'    Reserved1 As Long 'pattern hodler business
'    Reserved2 As Long 'artists or pattern lot
'    Reserved3 As Long 'sub collection of above
'                'or subordinate floss suppliers
'End Type
'
''after a single relate record is
''the bult render correlate data
''in sequence until a gapinfo stop
'Public Type BlockInfo '6 bytes
'    Color8bits As Byte
'    Color16bits As Byte
'    Color32bits As Byte
'    Color64bits As Byte
'    StitchIndex As Byte
'    FlossStrands As Byte
'End Type
''a single gapinfo seperates
''any number of blockinfo
'Public Type GapInfo '6 bytes
'     'next x block of used pattern
'    NextXY As Integer
'    NextYX As Integer
'    Scalar As Integer
'End Type
'
'Public Type ProjType
'    Header As GridInfo
'    NativeWork As RelateInfo
'    BlockData() As BlockInfo
'    GapRecord() As GapInfo
'    RTFData() As Byte  'after the symbols
'    'the remaining rest of the file is in
'    'rtf format of styling around patterns
'End Type 'instruction sheets presently made

Public Const Transparent As Long = &HFFFF00FF
Public Const MouseSensitivity As Long = 1
Public Const MouseZoomInMax As Long = -40000
Public Const MouseZoomOutMax As Long = 40000

Public Const MaxDisplacement As Single = 200
Public Const FadeDistance As Single = 1000000

Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Const SWW_HPARENT = -8

Public StitchInfo(1 To 12) As Long

Public Function ScreenXYTo3DZ0(ByVal MouseX As Single, ByVal MouseY As Single) As D3DVECTOR
    Dim Alpha As Single

    Dim vDir As D3DVECTOR
    Dim vIntersect As D3DVECTOR

    Dim halfwidth As Long
    Dim halfheight As Long

    halfwidth = Screen.Width / 2
    halfheight = Screen.Height / 2

    Alpha = FOV / 2

    'Screen Width is 640x480
    MouseX = Tan(Alpha) * (frmStudio.LastX / halfwidth - 1) / ASPECT
    MouseY = Tan(Alpha) * (1 - frmStudio.LastY / halfheight)

    'Intersect the plane from the eye to the centre of the plane

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
    v1.X = -100000
    v1.Y = -100000
    v1.Z = Player.CameraZoom

    v2.X = -100000
    v2.Y = 100000
    v2.Z = Player.CameraZoom

    v3.X = 100000
    v3.Y = 100000
    v3.Z = Player.CameraZoom
    
    pPlane1 = Create4DPlaneVectorFromPoints(v1, v2, v3)

    Dim c As D3DVECTOR
    Dim N As D3DVECTOR
    Dim P As D3DVECTOR
    Dim V As D3DVECTOR
    
    Dim hit As Boolean
   ' vIntersect.X = (vIntersect.X * (ThatchSquare / ThatchScale))
   ' vIntersect.Y = (vIntersect.Y * (ThatchSquare / ThatchScale))
        
    If RayIntersectPlane(pPlane1, p1, vDir, vIntersect) Then
        ScreenXYTo3DZ0 = vIntersect
    End If
End Function

Public Function GetCursorPosEx(ByVal hwnd As Long, ByRef pt As POINTAPI) As Long
    
    Dim bsize As Long
    bsize = GetParent(hwnd)
    hwnd = IIf(bsize > 0, bsize, hwnd)
    If hwnd <> 0 Then
        If GetCursorPos(pt) Then
            
            pt.X = (pt.X * Screen.TwipsPerPixelX) 'convert to screen twips subtract the
            'left for reletive position coord match from the mouse x coord of designer
        
            bsize = ((frmStudio.Designer.Width - frmStudio.Designer.ScaleWidth) / 2) 'get a single
            bsize = (frmStudio.Designer.Left + bsize) ' width of the forms side edge, add it to the
            pt.X = pt.X - bsize 'forms left on screen subtract that value from the mouse cursor
            bsize = ((frmStudio.Width - frmStudio.ScaleWidth) / 2) 'do the same with the designer
            pt.X = pt.X - (frmStudio.Left + bsize) 'subtract that value from the mouse cursor
            
        
            bsize = (bsize / Screen.TwipsPerPixelX) 'size of one side border in pixel, we are using the width
            'because of the title bar is alone on the height with two borders similar to width but with title
            bsize = (bsize * Screen.TwipsPerPixelY) 'convert to pixels of a single size of border into twips of Y axis
            bsize = (frmStudio.Height - frmStudio.ScaleHeight) - (bsize * 2) 'subtract two of them from the full heights
            'difference in scalar (internal coordinate system) and it's true value (the external edge coordinate system)
            'left with just title bar height in twips, between true value and scale coords is broders/title bar sizes
        
            'same thing we did to the width before, but incorporate the title subtraction to locate absolute designer posotion on screen
            pt.Y = (pt.Y * Screen.TwipsPerPixelY) - (frmStudio.Top + bsize + (((frmStudio.Height - frmStudio.ScaleHeight) - bsize) / 2) + _
                    (frmStudio.Designer.Top + ((frmStudio.Designer.Height - frmStudio.Designer.ScaleHeight) / 2)))
            
            GetCursorPosEx = hwnd

        End If
    End If

End Function

Function FloatToDWord(F As Single) As Long
    Dim buf As D3DXBuffer
    Dim L As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, F
    D3DX.BufferGetData buf, 0, 4, 1, L
    FloatToDWord = L
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

Public Function RandomPositive(LowerBound As Long, Upperbound As Long) As Single
    Randomize
    RandomPositive = CSng((Upperbound - LowerBound + 1) * Rnd + LowerBound)
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
End Function

Public Function SquareCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As D3DVECTOR

    SquareCenter.X = (v0.X + v1.X + v2.X + v3.X) / 4
    SquareCenter.Y = (v0.Y + v1.Y + v2.Y + v3.Y) / 4
    SquareCenter.Z = (v0.Z + v1.Z + v2.Z + v3.Z) / 4
    
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = Sqr(((p1x - p2x) * (p1x - p2x)) + ((p1y - p2y) * (p1y - p2y)) + ((p1z - p2z) * (p1z - p2z)))
End Function

Public Function Distance2(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Distance2 = Sqr(((p1.X - p2.X) * (p1.X - p2.X)) + ((p1.Y - p2.Y) * (p1.Y - p2.Y)) + ((p1.Z - p2.Z) * (p1.Z - p2.Z)))
End Function

Public Function CreateVertex(X As Single, Y As Single, Z As Single, NX As Single, NY As Single, Nz As Single, tu As Single, tv As Single) As TVERTEX2
    
    With CreateVertex
        .X = X: .Y = Y: .Z = Z
        .NX = NX: .NY = NY: .Nz = Nz
        .tu = tu: .tv = tv
    End With
    
End Function

Public Function TriangleCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR
    
    TriangleCenter.X = ((v0.X + v1.X + v2.X) / 3)
    TriangleCenter.X = ((v0.Y + v1.Y + v2.Y) / 3)
    TriangleCenter.X = ((v0.Z + v1.Z + v2.Z) / 3)

End Function

Public Function TriangleNormal(ByRef p0 As D3DVECTOR, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As D3DVECTOR
    TriangleNormal = VectorNormalize(VectorCrossProduct(VectorSubtract(p1, p0), VectorSubtract(p2, p0)))
End Function

Private Function VectorNormalize(ByRef V As D3DVECTOR) As D3DVECTOR
    Dim L As Single
    L = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
    If L = 0 Then L = 1
    VectorNormalize.X = (V.X / L)
    VectorNormalize.Y = (V.Y / L)
    VectorNormalize.Z = (V.Z / L)
End Function

Private Function VectorDotProduct(ByRef V As D3DVECTOR, ByRef u As D3DVECTOR) As Single
    VectorDotProduct = (u.X * V.X + u.Y * V.Y + u.Z * V.Z)
End Function

Private Function VectorCrossProduct(ByRef V As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorCrossProduct.X = V.Y * u.Z - V.Z * u.Y
    VectorCrossProduct.Y = V.Z * u.X - V.X * u.Z
    VectorCrossProduct.Z = V.X * u.Y - V.Y * u.X
End Function

Public Function VectorSubtract(ByRef V As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorSubtract.X = V.X - u.X
    VectorSubtract.Y = V.Y - u.Y
    VectorSubtract.Z = V.Z - u.Z
End Function
Public Function VectorAdd(ByRef V As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorAdd.X = V.X + u.X
    VectorAdd.Y = V.Y + u.Y
    VectorAdd.Z = V.Z + u.Z
End Function
Public Function VectorMultiply(ByRef V As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorMultiply.X = V.X * u.X
    VectorMultiply.Y = V.Y * u.Y
    VectorMultiply.Z = V.Z * u.Z
End Function

Public Function PointToPlane(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Dim p3 As D3DVECTOR
    p3 = VectorSubtract(p1, p2)
    PointToPlane = Sqr(VectorDotProduct(p3, p3))
End Function

Public Function MidPoint(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As D3DVECTOR

    With MidPoint
        .X = ((Large(p1.X, p2.X) - Least(p1.X, p2.X)) / 2) '2=number of points
        .Y = ((Large(p1.Y, p2.Y) - Least(p1.Y, p2.Y)) / 2) '2=number of points
        .Z = ((Large(p1.Z, p2.Z) - Least(p1.Z, p2.Z)) / 2) '2=number of points
    End With
    
End Function

Public Function SlopePoint(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As D3DVECTOR

    With SlopePoint
        .X = ((p1.X - p2.X) / 2)  '2=number of points
        .Y = ((p1.Y - p2.Y) / 2)  '2=number of points
        .Z = ((p1.Z - p2.Z) / 2)  '2=number of points
    End With

End Function

Public Function SlopeIntercept(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single

    Dim S As D3DVECTOR
    S = MidPoint(p1, p2)
    SlopeIntercept = Sqr(((S.X + S.Y + S.Z) ^ 2) / 3) '3=number of axii

End Function

Public Function LengthSqr(ByRef V As D3DVECTOR) As Single
    LengthSqr = Sqr(Distance(0, 0, 0, V.X, V.Y, V.Z))
End Function

Public Function InterceptPoint(ByRef l1p1 As D3DVECTOR, ByRef l1p2 As D3DVECTOR, ByRef l2p1 As D3DVECTOR, ByRef l2p2 As D3DVECTOR, ByRef r1 As D3DVECTOR, ByRef r2 As D3DVECTOR) As Boolean
    ' // Algorithm is ported from the C algorithm of (but third port by Nick to vb)
    '// Paul Bourke at http://local.wasp.uwa.edu.au/~pbourke/geometry/lineline3d/
    r1 = VectorSubtract(l1p1, l2p1)
    r2 = VectorSubtract(l2p2, l2p1)
    If Not (LengthSqr(r2) < epsilon) Then
        Dim p21 As D3DVECTOR
        p21 = VectorSubtract(l1p2, l1p1)
        If Not (LengthSqr(p21) < epsilon) Then
            Dim d1343 As Single
            Dim d4321 As Single
            Dim d1321 As Single
            Dim d4343 As Single
            Dim d2121 As Single
            Dim denom As Single
            Dim numer As Single
            d1343 = r1.X * r2.X + r1.Y * r2.Y + r1.Z * r2.Z
            d4321 = r2.X * p21.X + r2.Y * p21.Y + r2.Z * p21.Z
            d1321 = r1.X * p21.X + r1.Y * p21.Y + r1.Z * p21.Z
            d4343 = r2.X * r2.X + r2.Y * r2.Y + r2.Z * r2.Z
            d2121 = p21.X * p21.X + p21.Y * p21.Y + p21.Z * p21.Z
            denom = d2121 * d4343 - d4321 * d4321
            If Not (Abs(denom) < epsilon) Then
                numer = d1343 * d4321 - d1321 * d4343
                numer = numer / denom
                denom = (d1343 + d4321 * (numer)) / d4343
                r1.X = (l1p1.X + numer * p21.X)
                r1.Y = (l1p1.Y + numer * p21.Y)
                r1.Z = (l1p1.Z + numer * p21.Z)
                r2.X = (l2p1.X + denom * r2.X)
                r2.Y = (l2p1.Y + denom * r2.Y)
                r2.Z = (l2p1.Z + denom * r2.Z)
                InterceptPoint = True
            End If
        End If
    End If
End Function

'
'Function TriangleNormal(ByRef v0 As Vector3D, ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Vector3D
'    TriangleNormal = VectorCrossProduct(VectorSubtract(v0, v1), VectorSubtract(v1, v2))
'End Function
'
'Public Function DotProduct(ByRef v As Vector3D, ByRef u As Vector3D) As Single
'    DotProduct = (u.X * v.X + u.Y * v.Y + u.z * v.z)
'End Function
'
'Function VectorCrossProduct(ByRef v As Vector3D, ByRef u As Vector3D) As Vector3D
'    VectorCrossProduct.X = ((v.Y * u.z) - (v.z * u.Y))
'    VectorCrossProduct.Y = ((v.z * u.X) - (v.X * u.z))
'    VectorCrossProduct.z = ((v.X * u.Y) - (v.Y * u.X))
'End Function
'
'Function VectorSubtract(ByRef v As Vector3D, ByRef u As Vector3D) As Vector3D
'    VectorSubtract.X = (v.X - u.X)
'    VectorSubtract.Y = (v.Y - u.Y)
'    VectorSubtract.z = (v.z - u.z)
'End Function
'
'Function VectorNormalize(ByRef v As Vector3D) As Vector3D
'    Dim l As Single
'    l = Sqr(((v.X * v.X) + (v.Y * v.Y) + (v.z * v.z)))
'    If l = 0 Then l = 1
'    VectorNormalize.X = (v.X / l)
'    VectorNormalize.Y = (v.Y / l)
'    VectorNormalize.z = (v.z / l)
'End Function
'
'
'Public Function PointNormal(ByRef v As Vector3D) As Single
'    PointNormal = Sqr(DotProduct(v, v))
'End Function
'
'Function TriangleCenter(ByRef v0 As Vector3D, ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Vector3D
'
'    Dim vR As Vector3D
'
'    vR.X = (v0.X + v1.X + v2.X) / 3
'    vR.Y = (v0.Y + v1.Y + v2.Y) / 3
'    vR.z = (v0.z + v1.z + v2.z) / 3
'
'    TriangleCenter = vR
'
'End Function

Public Function Large(ByVal p1 As Variant, ByVal p2 As Variant, Optional ByVal p3 As Variant = Empty) As Variant
    If TypeName(p3) = "Empty" Then
        If p1 > p2 Then
            Large = p1
        Else
            Large = p2
        End If
    Else
        If p1 > p2 And p1 > p3 Then
            Large = p1
        ElseIf p2 > p1 And p2 > p3 Then
            Large = p2
        Else
            Large = p3
        End If
    End If
End Function

Public Function Least(ByVal p1 As Variant, ByVal p2 As Variant, Optional ByVal p3 As Variant = Empty) As Variant
    If TypeName(p3) = "Empty" Then
        If p1 < p2 Then
            Least = p1
        Else
            Least = p2
        End If
    Else
        If p1 < p2 And p1 < p3 Then
            Least = p1
        ElseIf p2 < p1 And p2 < p3 Then
            Least = p2
        Else
            Least = p3
        End If
    End If
End Function


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


Public Sub CreateGridPlate(ByRef data() As TVERTEX2, ByRef Verticies As Direct3DVertexBuffer8, Optional ByVal col As Long = 1, Optional ByVal row As Long = 1, Optional ByRef count As Long = -1, Optional ByVal PlateWidth As Single = 1, Optional ByVal PlateHeight As Single = 1, Optional ByVal tv As Single, Optional ByVal tu As Single)
    'creates a number of square plates n, is of either row*col or count which ever is greater, count can be existing already and elements are added modifying it
    Dim R As Long
    Dim N As Long
    R = row 'backup for later
    If (row * col) > count Then count = (row * col)
    ReDim Preserve data(0 To Abs(IIf((row * col) = 0, 1, (row * col)) * 6) - 1) As TVERTEX2
    N = IIf(count > 0 And Not count > (row * col), (row * col) - count, count)
    For count = (((N - 1) \ 6) + -CInt(CBool((N - 1) Mod 6 > 0))) * 6 To UBound(data) Step 6
        CreateSquare data, count, MakeVector((PlateWidth / 2), -(PlateHeight / 2), 0), _
            MakeVector(-(PlateWidth / 2), -(PlateHeight / 2), 0), _
            MakeVector(-(PlateWidth / 2), (PlateHeight / 2), 0), _
            MakeVector((PlateWidth / 2), (PlateHeight / 2), 0), tv, tu
        row = row - 1
        If row = 0 Then row = R
        If row = R Then col = col - 1
    Next
    count = count \ 6
    Set Verticies = DDevice.CreateVertexBuffer(Len(data(0)) * (UBound(data) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Verticies, 0, Len(data(0)) * (UBound(data) + 1), 0, data(0)

End Sub

'Public Sub CreateGridPlateEx(ByRef data() As TVERTEX2, ByRef Verticies As Direct3DVertexBuffer8, Optional ByVal col As Long = 1, Optional ByVal row As Long = 1, Optional ByRef count As Long = -1, Optional ByVal PlateWidth As Single = 1, Optional ByVal PlateHeight As Single = 1, Optional ByVal tv As Single, Optional ByVal tu As Single, Optional ByVal StreamFlag As Long)
'    'creates a number of square plates n, is of either row*col or count which ever is greater, count can be existing already and elements are added modifying it
'    Dim R As Long
'    Dim N As Long
'    R = row 'backup for later
'    If (row * col) > count Then count = (row * col)
'    ReDim Preserve data(0 To Abs(IIf((row * col) = 0, 1, (row * col)) * 6) - 1) As TVERTEX2
'    N = IIf(count > 0 And Not count > (row * col), (row * col) - count, count)
'    For count = (((N - 1) \ 6) + -CInt(CBool((N - 1) Mod 6 > 0))) * 6 To UBound(data) Step 6
'        CreateSquareEx data, count, MakeVector((PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), -(PlateHeight / 2), 0), _
'            MakeVector(-(PlateWidth / 2), (PlateHeight / 2), 0), _
'            MakeVector((PlateWidth / 2), (PlateHeight / 2), 0), tv, tu, StreamFlag
'        row = row - 1
'        If row = 0 Then row = R
'        If row = R Then col = col - 1
'    Next
'    count = count \ 6
'    Set Verticies = DDevice.CreateVertexBuffer(Len(data(0)) * (UBound(data) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Verticies, 0, Len(data(0)) * (UBound(data) + 1), 0, data(0)
'
'End Sub
'
Public Sub CreateSquare(ByRef data() As TVERTEX2, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1)

    Dim vn As D3DVECTOR

    data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
    data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    vn = TriangleNormal(MakeVector(data(Index + 0).X, data(Index + 0).Y, data(Index + 0).Z), _
                            MakeVector(data(Index + 1).X, data(Index + 1).Y, data(Index + 1).Z), _
                            MakeVector(data(Index + 2).X, data(Index + 2).Y, data(Index + 2).Z))
    data(Index + 0).NX = vn.X: data(Index + 0).NY = vn.Y: data(Index + 0).Nz = vn.Z
    data(Index + 1).NX = vn.X: data(Index + 1).NY = vn.Y: data(Index + 1).Nz = vn.Z
    data(Index + 2).NX = vn.X: data(Index + 2).NY = vn.Y: data(Index + 2).Nz = vn.Z

    data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, 0, 0)
    vn = TriangleNormal(MakeVector(data(Index + 3).X, data(Index + 3).Y, data(Index + 3).Z), _
                            MakeVector(data(Index + 4).X, data(Index + 4).Y, data(Index + 4).Z), _
                            MakeVector(data(Index + 5).X, data(Index + 5).Y, data(Index + 5).Z))
    data(Index + 3).NX = vn.X: data(Index + 3).NY = vn.Y: data(Index + 3).Nz = vn.Z
    data(Index + 4).NX = vn.X: data(Index + 4).NY = vn.Y: data(Index + 4).Nz = vn.Z
    data(Index + 5).NX = vn.X: data(Index + 5).NY = vn.Y: data(Index + 5).Nz = vn.Z

End Sub
'Public Sub CreateSquareEx(ByRef data() As TVERTEX2, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1, Optional ByVal StreamFlag As Long)
'
'    Static BrushIndex As Long
'    Dim FaceIndex As Long
'
'    Dim vn As D3DVECTOR
'
'    data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
'    data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
'    data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
'
'    vn = TriangleNormal(MakeVector(data(Index + 0).X, data(Index + 0).Y, data(Index + 0).Z), _
'                            MakeVector(data(Index + 1).X, data(Index + 1).Y, data(Index + 1).Z), _
'                            MakeVector(data(Index + 2).X, data(Index + 2).Y, data(Index + 2).Z))
'
'    data(Index + 0).NX = vn.X: data(Index + 0).NY = vn.Y: data(Index + 0).Nz = vn.Z
'    data(Index + 1).NX = vn.X: data(Index + 1).NY = vn.Y: data(Index + 1).Nz = vn.Z
'    data(Index + 2).NX = vn.X: data(Index + 2).NY = vn.Y: data(Index + 2).Nz = vn.Z
'
'    AddTriangleFaceToCollision BrushIndex, FaceIndex, vn, MakeVector(data(Index + 0).X, data(Index + 0).Y, data(Index + 0).Z), _
'                                            MakeVector(data(Index + 1).X, data(Index + 1).Y, data(Index + 1).Z), _
'                                            MakeVector(data(Index + 2).X, data(Index + 2).Y, data(Index + 2).Z), StreamFlag
'    FaceIndex = FaceIndex + 1
'    data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
'    data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
'    data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, 0, 0)
'
'    vn = TriangleNormal(MakeVector(data(Index + 3).X, data(Index + 3).Y, data(Index + 3).Z), _
'                            MakeVector(data(Index + 4).X, data(Index + 4).Y, data(Index + 4).Z), _
'                            MakeVector(data(Index + 5).X, data(Index + 5).Y, data(Index + 5).Z))
'
'    data(Index + 3).NX = vn.X: data(Index + 3).NY = vn.Y: data(Index + 3).Nz = vn.Z
'    data(Index + 4).NX = vn.X: data(Index + 4).NY = vn.Y: data(Index + 4).Nz = vn.Z
'    data(Index + 5).NX = vn.X: data(Index + 5).NY = vn.Y: data(Index + 5).Nz = vn.Z
'
'    AddTriangleFaceToCollision BrushIndex, FaceIndex, vn, MakeVector(data(Index + 3).X, data(Index + 3).Y, data(Index + 3).Z), _
'                                            MakeVector(data(Index + 4).X, data(Index + 4).Y, data(Index + 4).Z), _
'                                            MakeVector(data(Index + 5).X, data(Index + 5).Y, data(Index + 5).Z), StreamFlag
'
'    BrushIndex = BrushIndex + 1
'End Sub




