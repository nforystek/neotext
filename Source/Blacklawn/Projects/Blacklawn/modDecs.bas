#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modDecs"
#Const modDecs = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MyScreen
    X As Single
    Y As Single
    z As Single
    RHW As Single
    Color As Long
    tu As Single
    tv As Single
End Type

Public Type MyVertex
    X    As Single
    Y    As Single
    z    As Single
    nx   As Single
    ny   As Single
    nz   As Single
    tu   As Single
    tv   As Single
End Type

Public Const RENDER_MAGFILTER = D3DTEXF_ANISOTROPIC
Public Const RENDER_MINFILTER = D3DTEXF_ANISOTROPIC

Public Const FVF_VERTEX_SIZE = 12
Public Const FVF_RENDER_SIZE = 32

Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Const Transparent As Long = &HFFFF00FF

Public Const ScaleModelSize As Single = 0.22
Public Const ScaleModelLocX As Single = -2530
Public Const ScaleModelLocY As Single = 10
Public Const ScaleModelLocZ As Single = 1970

Public Const WithInCityLimits As Long = 5000
Public Const MouseSensitivity As Long = 4
Public Const MoveSpeedMax As Single = 10
Public Const MoveSpeedMin As Single = 0.05
Public Const MoveSpeedInc As Single = 0.02
Public Const GroundFriction As Single = 0.05
Public Const GravityVelocity As Single = 25
Public Const MaxGravity As Single = 200
Public Const FadeDistance As Single = 30000
Public Const ZoneDistance As Single = 2500
Public Const BlackBoundary As Single = 1000000
Public Const MaxTalkMsgs As Long = 6

Public Const PartnerMinSpeed As Single = 0.02
Public Const PartnerMaxSpeed As Single = 0.2

Public Const BlacklawnVer As String = "v5"

Public Const Ambient_HI As Long = 255
Public Const Ambient_LO As Long = 64

Public Const PI As Single = 3.14159265359
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2

Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public Function SetAmbientRGB(ByVal r As Single, ByVal G As Single, ByVal b As Single, ByVal Restraint As Integer)
    If Restraint = 0 Then
        If Not DDevice.GetRenderState(D3DRS_AMBIENT) = RGB(r, G, b) Then
            DDevice.SetRenderState D3DRS_AMBIENT, RGB(r, G, b)
        End If
    ElseIf Restraint = 1 Or Restraint = 2 Then
        Dim HI As Long
        Dim Lo As Long
        If Restraint = 1 Then
            HI = 136
            Lo = 32
        ElseIf Restraint = 2 Then
            HI = 64
            Lo = 32
        End If
        If (r - HI) >= Lo Then
            If Not DDevice.GetRenderState(D3DRS_AMBIENT) = RGB(r - HI, G - HI, b - HI) Then
                DDevice.SetRenderState D3DRS_AMBIENT, RGB(r - HI, G - HI, b - HI)
            End If
        Else
            If Not DDevice.GetRenderState(D3DRS_AMBIENT) = RGB(Lo, Lo, Lo) Then
                DDevice.SetRenderState D3DRS_AMBIENT, RGB(Lo, Lo, Lo)
            End If
        End If
    End If
End Function

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
    If sBuffer = "" Then sBuffer = "SHARED"
    GetUserLoginName = sBuffer
    
End Function

Public Function AppPath() As String
    Dim ret As String
    ret = IIf((Right(App.Path, 1) = "\"), App.Path, App.Path & "\")
#If VBIDE = -1 Then
    If Right(ret, 9) = "Projects\" Then
        ret = Replace(ret, "Projects\", "Binary\", , , vbTextCompare)
     End If
#End If
    AppPath = ret
End Function

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(TheParams, TheSeperator) > 0 Then
        NextArg = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
    Else
        NextArg = Trim(TheParams)
    End If
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator) > 0 Then
        RemoveArg = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator)))
    Else
        RemoveArg = ""
    End If
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(TheParams, TheSeperator) > 0 Then
        RemoveNextArg = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
        TheParams = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator)))
    Else
        RemoveNextArg = Trim(TheParams)
        TheParams = ""
    End If
End Function

Public Function RemoveNextArgNoTrim(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(TheParams, TheSeperator) > 0 Then
        RemoveNextArgNoTrim = Left(TheParams, InStr(TheParams, TheSeperator) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator))
    Else
        RemoveNextArgNoTrim = TheParams
        TheParams = ""
    End If
End Function

Public Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    Dim retVal As String
    Dim X As Long
    X = InStr(TheParams, BeginQuote)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + 1, TheParams, EndQuote) > 0) Then
            retVal = Mid(TheParams, X + 1)
            TheParams = Left(TheParams, X - 1) & Mid(retVal, InStr(retVal, EndQuote) + 1)
            retVal = Left(retVal, InStr(retVal, EndQuote) - 1)
        End If
    End If
    RemoveQuotedArg = retVal
End Function

Public Function PathExists(ByVal URL As String, Optional ByVal IsFile As Variant) As Boolean
    If (URL = vbNullString) Then
        PathExists = False
        Exit Function
    ElseIf (Not IsMissing(IsFile)) Then
        If ((GetFilePath(URL) = vbNullString) And IsFile And (Not (URL = vbNullString))) Or ((GetFileName(URL) = vbNullString) And (Not IsFile) And (Not (URL = vbNullString))) Then
            PathExists = False
            Exit Function
        End If
    End If
    If (IsMissing(IsFile)) Then IsFile = False
    If (Len(URL) = 2) And (Mid(URL, 2, 1) = ":") Then
        URL = URL & "\"
    End If
    Dim ret As Long
    On Error Resume Next
    ret = GetAttr(URL)
    If Err.Number = 0 Then
        PathExists = IIf(IsFile, Not CBool(ret And vbDirectory), True)
    Else
        Err.Clear
        PathExists = False
    End If
    On Error GoTo 0
End Function

Public Function ReadFile(ByVal Path As String) As String
    If PathExists(Path, True) Then
        Dim num As Integer
        Dim txt As String
        num = FreeFile
        Open Path For Input Shared As #num
        Close #num
        Open Path For Binary Shared As #num
            txt = String(FileSize(Path), Chr(0))
            Get #num, 1, txt
        Close #num
        ReadFile = txt
    Else
        Err.Raise 53, App.EXEName, "File not found"
    End If
End Function

Public Sub WriteFile(ByVal Path As String, ByVal Text As String)
    If PathExists(Path, True) Then
        Kill Path
    End If
    Dim num As Integer
    num = FreeFile
    Open Path For Output Shared As #num
    Close #num
    num = FreeFile
    Open Path For Binary Shared As #num
        Put #num, 1, Text
    Close #num
End Sub

Public Function GetFilePath(ByVal URL As String) As String
    Dim nFolder As String
    If InStr(URL, "/") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "/") - 1)
        If nFolder = "" Then nFolder = "/"
    ElseIf InStr(URL, "\") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "\") - 1)
        If nFolder = "" Then nFolder = "\"
    Else
        nFolder = ""
    End If
    GetFilePath = nFolder
End Function

Public Function GetFileName(ByVal URL As String) As String
    If InStr(URL, "/") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "/") + 1)
    ElseIf InStr(URL, "\") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "\") + 1)
    Else
        GetFileName = URL
    End If
End Function

Public Function GetFileTitle(ByVal URL As String) As String
    URL = GetFileName(URL)
    If InStrRev(URL, ".") > 0 Then
        URL = Left(URL, InStrRev(URL, ".") - 1)
    End If
    GetFileTitle = URL
End Function

Public Function FileSize(ByVal fName As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fso.GetFile(fName)
    FileSize = f.Size
    Set f = Nothing
    Set fso = Nothing
End Function

Public Function RandomPositive(Lowerbound As Long, Upperbound As Long) As Single
    
    RandomPositive = CSng((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.z = z
End Function

Public Function SquareCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As D3DVECTOR

    SquareCenter.X = (v0.X + v1.X + v2.X + v3.X) / 4
    SquareCenter.Y = (v0.Y + v1.Y + v2.Y + v3.Y) / 4
    SquareCenter.z = (v0.z + v1.z + v2.z + v3.z) / 4
    
End Function

Public Function Distance(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single) As Single
    Distance = Sqr(((p1x - p2x) * (p1x - p2x)) + ((p1y - p2y) * (p1y - p2y)) + ((p1z - p2z) * (p1z - p2z)))
End Function

Public Function CreateVertex(X As Single, Y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As MyVertex
    
    With CreateVertex
        .X = X: .Y = Y: .z = z
        .nx = nx: .ny = ny: .nz = nz
        .tu = tu: .tv = tv
    End With
    
End Function

Public Function TriangleCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR

    Dim vR As D3DVECTOR
 
    vR.X = (v0.X + v1.X + v2.X) / 3
    vR.Y = (v0.Y + v1.Y + v2.Y) / 3
    vR.z = (v0.z + v1.z + v2.z) / 3
    
    TriangleCenter = vR

End Function

Public Function TriangleNormal(ByRef p0 As D3DVECTOR, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As D3DVECTOR
    TriangleNormal = VectorNormalize(VectorCrossProduct(VectorSubtract(p1, p0), VectorSubtract(p2, p0)))
End Function

'Public Function GenerateTriNormals(p0 As D3DVECTOR, p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR
'
'    Dim v01 As D3DVECTOR
'    Dim v02 As D3DVECTOR
'    Dim vNorm As D3DVECTOR
'
'    D3DXVec3Subtract v01, p1, p0
'    D3DXVec3Subtract v02, p2, p0
'
'    D3DXVec3Cross vNorm, v01, v02
'
'    D3DXVec3Normalize vNorm, vNorm
'
'    GenerateTriNormals.X = vNorm.X
'    GenerateTriNormals.Y = vNorm.Y
'    GenerateTriNormals.Z = vNorm.Z
'
'End Function

Public Sub CreateSquare(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.z, 0, 0, 0, 0, ScaleY)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.z, 0, 0, 0, ScaleX, ScaleY)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.z, 0, 0, 0, ScaleX, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.z, 0, 0, 0, 0, ScaleY)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.z, 0, 0, 0, ScaleX, 0)
    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.z, 0, 0, 0, 0, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.z
    
End Sub

Public Sub CreateSquareEx(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As MyVertex, ByRef p2 As MyVertex, ByRef p3 As MyVertex, ByRef p4 As MyVertex)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.z, 0, 0, 0, p2.tu, p2.tv)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.z, 0, 0, 0, p3.tu, p3.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.z, 0, 0, 0, p3.tu, p3.tv)
    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.z, 0, 0, 0, p4.tu, p4.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.z
    
End Sub

Public Function GetNow() As String
    If frmMain.IsPlayback Then
        
        GetNow = DateAdd("s", frmMain.AtTime, StartNow)
        
    Else
        GetNow = CStr(Now)
    End If
End Function

Public Function GetTimer() As String
    If frmMain.IsPlayback Then
        GetTimer = frmMain.AtTime
    Else
        GetTimer = CStr(CDbl(Timer))
    End If
    
End Function

Public Sub Swap(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    Dim pos As Long
    cnt = 0
    pos = InStr(Text, Word)
    Do Until pos = 0
        cnt = cnt + 1
        pos = InStr(pos + 1, Text, Word)
    Loop
    CountWord = cnt
End Function

Public Function GreatestWidth(ByVal Text As String) As Single
    Dim gl As Single
    Dim line As String

    Do
        line = RemoveNextArg(Text, vbCrLf)
        If frmMain.TextWidth(line) > gl Then gl = frmMain.TextWidth(line)
    Loop Until Text = ""
    
    GreatestWidth = gl

End Function

Public Function SortVertices(ByRef PolyX() As Single, ByRef PolyY() As Single, ByRef PolyZ() As Single, ByRef PolyN As Long, ByVal PerFaceN As Long)
    Dim a As D3DVECTOR
    Dim b As D3DVECTOR

    Dim pNormal As D3DVECTOR
    
    Dim vCenter As D3DVECTOR
    
    Dim vNormal As D3DVECTOR
    Dim vOriginDistance As Single

    Dim VertexCount As Long
    Dim vVertex(0 To 2) As D3DVECTOR
    
    Dim angle As Single
    
    Dim smallest As Long
    Dim smallestAngle As Single
    
    Dim n As Long
    Dim m As Long
    
    If PolyN >= 3 Then
    
        'Debug.Print "SortVertices Begin")
        
        pNormal = GetPlaneNormal(MakeVector(PolyX(0), PolyY(0), PolyZ(0)), _
                                MakeVector(PolyX(1), PolyY(1), PolyZ(1)), _
                                MakeVector(PolyX(2), PolyY(2), PolyZ(2)))
                                    
        'Debug.Print "   Calculating Center"
        
        Dim cnt As Long
        For cnt = 0 To PolyN - 1
            vCenter.X = vCenter.X + PolyX(cnt)
            vCenter.Y = vCenter.Y + PolyY(cnt)
            vCenter.z = vCenter.z + PolyZ(cnt)
        Next
        vCenter.X = vCenter.X / PolyN
        vCenter.Y = vCenter.Y / PolyN
        vCenter.z = vCenter.z / PolyN
    
        n = 0
        Do While n <= (PolyN - 1)

                vVertex(0) = MakeVector(PolyX(n), PolyY(n), PolyZ(n))
                a = VectorNormalize(VectorSubtract(vVertex(0), vCenter))
                
                vVertex(1) = vCenter
                vVertex(2) = VectorAdd(vCenter, pNormal)
            
                vNormal = TriangleNormal(vVertex(0), vVertex(1), vVertex(2))
                vOriginDistance = -VectorDotProduct(vNormal, vVertex(0))
                
                smallest = -1
                smallestAngle = -1
                
                m = n + 1
                Do While m <= (PolyN - 1)
                
                        If Not ClassifyPoint(MakeVector(PolyX(m), PolyY(m), PolyZ(m)), vNormal, vOriginDistance) = 2 Then 'not back
                            b = VectorNormalize(VectorSubtract(MakeVector(PolyX(m), PolyY(m), PolyZ(m)), vCenter))
                            
                            angle = VectorDotProduct(a, b)
                            
                            If angle > smallestAngle Then
                                smallestAngle = angle
                                smallest = m
                    
                            End If
                        End If
                    
                        m = m + 1

                Loop
                
                If smallest = -1 Then
                    'Debug.Print "   Degenerate Polygon"
                    'Debug.Print "SortVerticies End"
                    Exit Function
                End If
                
                SwapVertex PolyX, PolyY, PolyZ, n + 1, smallest
                'Debug.Print "   Swapping: " & (n + 1) & " with " & smallest
            
                n = n + PerFaceN

        Loop
        
        a = TriangleNormal(vVertex(0), vVertex(1), vVertex(2))
        b = pNormal
        
        If VectorDotProduct(a, b) < 0 Then
            'Debug.Print "   Reversing Verticies"
            ReverseFaceVertices PolyX, PolyY, PolyX, 3
        End If
        
        pNormal = a
        
        'Debug.Print "SortVerticies End"
    
    Else
        'Debug.Print "   Degenerate Polygon"
        'Debug.Print "SortVerticies End"
    End If
End Function

Public Function GetPlaneNormal(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR

    Dim vector1 As D3DVECTOR
    Dim vector2 As D3DVECTOR
    Dim Normal As D3DVECTOR
    Dim Length As Single
    
    '/*Calculate the Normal*/
    '/*Vector 1*/
    vector1.X = (v0.X - v1.X)
    vector1.Y = (v0.Y - v1.Y)
    vector1.z = (v0.z - v1.z)
    
    '/*Vector 2*/
    vector2.X = (v1.X - v2.X)
    vector2.Y = (v1.Y - v2.Y)
    vector2.z = (v1.z - v2.z)
    
    '/*Apply the Cross Product*/
    Normal.X = (vector1.Y * vector2.z - vector1.z * vector2.Y)
    Normal.Y = (vector1.z * vector2.X - vector1.X * vector2.z)
    Normal.z = (vector1.X * vector2.Y - vector1.Y * vector2.X)
    
    '/*Normalize to a unit vector*/
    Length = Sqr(Normal.X * Normal.X + Normal.Y * Normal.Y + Normal.z * Normal.z)
    
    If Length = 0 Then Length = 1
    
    Normal.X = (Normal.X / Length)
    Normal.Y = (Normal.Y / Length)
    Normal.z = (Normal.z / Length)
    
    GetPlaneNormal = Normal
End Function

Public Function ReverseFaceVertices(ByRef PolyX() As Single, ByRef PolyY() As Single, ByRef PolyZ() As Single, ByVal PolyN As Long)
    Dim cnt As Long
    Dim n As Long
    Dim vt As D3DVECTOR
    
    n = PolyN \ 2
    
    For cnt = 0 To n
        SwapVertex PolyX, PolyY, PolyZ, cnt, (PolyN - 1) - cnt
        'Debug.Print "   Swapping: " & (cnt) & " with " & ((PolyN - 1) - cnt)
    Next
    
End Function

Public Function ClassifyPoint(ByRef Point As D3DVECTOR, ByRef vNormal As D3DVECTOR, ByVal vOriginDistance As Single) As Single
    Const Epsilon = 0.99999
    Dim dtp As Single
    dtp = VectorDotProduct(vNormal, Point) + vOriginDistance
    
    If dtp > Epsilon Then
        ClassifyPoint = 1 'front
    ElseIf dtp < -Epsilon Then
        ClassifyPoint = 2 'back
    Else
        ClassifyPoint = 0 'onplane
    End If
End Function

Public Function SwapVertex(ByRef PolyX() As Single, ByRef PolyY() As Single, ByRef PolyZ() As Single, ByVal n1 As Long, ByVal n2 As Long)
    Dim Swap As Single
    
    Swap = PolyX(n1)
    PolyX(n1) = PolyX(n2)
    PolyX(n2) = Swap

    Swap = PolyY(n1)
    PolyY(n1) = PolyY(n2)
    PolyY(n2) = Swap

    Swap = PolyZ(n1)
    PolyZ(n1) = PolyZ(n2)
    PolyZ(n2) = Swap

End Function

Public Sub SwapVector(ByRef firstValue As D3DVECTOR, ByRef secondValue As D3DVECTOR)
    Dim tmpValue As D3DVECTOR
    tmpValue = firstValue
    firstValue = secondValue
    secondValue = tmpValue
End Sub
Private Function VectorNormalize(ByRef v As D3DVECTOR) As D3DVECTOR
    Dim l As Single
    l = Sqr(v.X * v.X + v.Y * v.Y + v.z * v.z)
    If l = 0 Then l = 1
    VectorNormalize.X = (v.X / l)
    VectorNormalize.Y = (v.Y / l)
    VectorNormalize.z = (v.z / l)
End Function

Private Function VectorDotProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As Single
    VectorDotProduct = (u.X * v.X + u.Y * v.Y + u.z * v.z)
End Function

Private Function VectorCrossProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorCrossProduct.X = v.Y * u.z - v.z * u.Y
    VectorCrossProduct.Y = v.z * u.X - v.X * u.z
    VectorCrossProduct.z = v.X * u.Y - v.Y * u.X
End Function

Private Function VectorSubtract(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorSubtract.X = v.X - u.X
    VectorSubtract.Y = v.Y - u.Y
    VectorSubtract.z = v.z - u.z
End Function
Private Function VectorAdd(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorAdd.X = v.X + u.X
    VectorAdd.Y = v.Y + u.Y
    VectorAdd.z = v.z + u.z
End Function
Private Function VectorMultiply(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorMultiply.X = v.X * u.X
    VectorMultiply.Y = v.Y * u.Y
    VectorMultiply.z = v.z * u.z
End Function

Private Function PointToPlane(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Dim p3 As D3DVECTOR
    p3 = VectorSubtract(p1, p2)
    PointToPlane = Sqr(VectorDotProduct(p3, p3))
End Function

Public Function GreatestX(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the greatext x value of three vectors
    If v0.X > v1.X And v0.X > v2.X Then
        GreatestX = v0.X
    ElseIf v1.X > v0.X And v1.X > v2.X Then
        GreatestX = v1.X
    Else
        GreatestX = v2.X
    End If
End Function

Public Function GreatestY(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the greatext y value of three vectors
    If v0.Y > v1.Y And v0.Y > v2.Y Then
        GreatestY = v0.Y
    ElseIf v1.Y > v0.Y And v1.Y > v2.Y Then
        GreatestY = v1.Y
    Else
        GreatestY = v2.Y
    End If
End Function

Public Function GreatestZ(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the greatext z value of three vectors
    If v0.z > v1.z And v0.z > v2.z Then
        GreatestZ = v0.z
    ElseIf v1.z > v0.z And v1.z > v2.z Then
        GreatestZ = v1.z
    Else
        GreatestZ = v2.z
    End If
End Function

Public Function LeastX(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the least x value of three vectors
    If v0.X < v1.X And v0.X < v2.X Then
        LeastX = v0.X
    ElseIf v1.X < v0.X And v1.X < v2.X Then
        LeastX = v1.X
    Else
        LeastX = v2.X
    End If
End Function

Public Function LeastY(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the least y value of three vectors
    If v0.Y < v1.Y And v0.Y < v2.Y Then
        LeastY = v0.Y
    ElseIf v1.Y < v0.Y And v1.Y < v2.Y Then
        LeastY = v1.Y
    Else
        LeastY = v2.Y
    End If
End Function

Public Function LeastZ(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As Single
    'return the least z value of three vectors
    If v0.z < v1.z And v0.z < v2.z Then
        LeastZ = v0.z
    ElseIf v1.z < v0.z And v1.z < v2.z Then
        LeastZ = v1.z
    Else
        LeastZ = v2.z
    End If
End Function

Public Function PointBehindtriangle(ByRef Point As D3DVECTOR, ByRef Center As D3DVECTOR, ByRef Lengths As D3DVECTOR) As Boolean

' (GreatestX(v1, v2, v3) - LeastX(v1, v2, v3)) / 2, _
'                                (GreatestY(v1, v2, v3) - LeastY(v1, v2, v3)) / 2, _
'                                (GreatestZ(v1, v2, v3) - LeastZ(v1, v2, v3)) / 2, _
'                                Distance(v1, v2), Distance(v2, v3), Distance(v3, v1),

    Dim v As D3DVECTOR
    Dim u As D3DVECTOR
    Dim n As D3DVECTOR
    
    Dim d As Single
    
    u = VectorCrossProduct(Point, Center)
        
    v = VectorSubtract(Point, Center)
    v = VectorSubtract(v, Center)
    v = VectorCrossProduct(v, Center)
    
    n.X = (((u.X + u.Y) * (Lengths.X + Lengths.Y + Lengths.z)) - ((v.X + v.Y) * (Lengths.X + Lengths.Y + Lengths.z)))
    n.Y = (((u.X + u.z) * (Lengths.X + Lengths.Y + Lengths.z)) - ((v.X + v.z) * (Lengths.X + Lengths.Y + Lengths.z)))
    n.X = (((u.Y + u.X) * (Lengths.X + Lengths.Y + Lengths.z)) - ((v.Y + v.X) * (Lengths.X + Lengths.Y + Lengths.z)))

    d = Distance(u.X, u.Y, u.z, v.X, v.Y, v.z)

    n.X = (n.X * u.X) / d
    n.Y = (n.Y * u.Y) / d
    n.z = (n.z * u.z) / d

    'Debug.Print "n: " & n.x & "," & n.y & "," & n.z & "  u: " & u.x & "," & u.y & "," & u.z & "  v: " & v.x & "," & v.y & "," & v.z & "  d: " & d
    
    PointBehindtriangle = n.X + n.Y + n.z > 0
End Function

Private Function TriangleIntersect(ByRef O1 As D3DVECTOR, ByRef O2 As D3DVECTOR, ByRef O3 As D3DVECTOR, ByRef Q1 As D3DVECTOR, ByRef Q2 As D3DVECTOR, ByRef Q3 As D3DVECTOR) As Long
    
    Dim No As D3DVECTOR
    Dim Nq As D3DVECTOR
    Dim Co As D3DVECTOR
    Dim Cq As D3DVECTOR
    Dim Lc As Single
    Dim Lo As Single
    Dim Lq As Single
    Dim Da As D3DVECTOR
    Dim Oi As Single
    Dim Qi As Single
    Dim Jo As D3DVECTOR
    Dim Jq As D3DVECTOR
    Dim K As Integer
    
    No = TriangleNormal(O1, O2, O3)
    Nq = TriangleNormal(Q1, Q2, Q3)
    Co = TriangleCenter(O1, O2, O3)
    Cq = TriangleCenter(Q1, Q2, Q3)
    Lc = Distance(Co.X, Co.Y, Co.z, Cq.X, Cq.Y, Cq.X)
    Lo = Distance(O1.X, O1.Y, O1.z, O2.X, O2.Y, O2.z) + Distance(O2.X, O2.Y, O2.z, O3.X, O3.Y, O3.z) + Distance(O3.X, O3.Y, O3.z, O1.X, O1.Y, O1.z)
    Lq = Distance(Q1.X, Q1.Y, Q1.z, Q2.X, Q2.Y, Q2.z) + Distance(Q2.X, Q2.Y, Q2.z, Q3.X, Q3.Y, Q3.z) + Distance(Q3.X, Q3.Y, Q3.z, Q1.X, Q1.Y, Q1.z)
    
    Da.X = Sqr((((Lo + Lq) * (No.X + Co.X)) + (((Q1.Y + Q2.Y + Q3.Y - O1.Y + O2.Y + O3.Y) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.Y * No.z * Co.Y) + (O1.Y + O2.Y + O3.Y)))) ^ 2)
    Da.Y = Sqr((((Lo + Lq) * (No.Y + Co.Y)) + (((Q1.X + Q2.X + Q3.X - O1.X + O2.X + O3.X) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.X * No.z * Co.X) + (O1.X + O2.X + O3.X)))) ^ 2)
    Da.z = Sqr((((Lo + Lq) * (No.z + Co.z)) + (((Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z) + (Q1.X + Q2.X + Q3.X - O1.X + O2.X + O3.X)) * ((No.X * No.X * Co.z) + (O1.z + O2.z + O3.z)))) ^ 2)

    Oi = Sqr(((Da.X * Da.Y * Da.z) ^ 3 + ((Da.X + Da.Y + Da.z) * (Da.X + Da.Y + Da.z))))
    
    Jo.X = (((Q1.X + Q2.X + Q3.X) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
    Jo.Y = (((Q1.X + Q2.X + Q3.X) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
    Jo.z = (((Q1.X + Q2.X + Q3.X) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
    
    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jo.x) & " " & DebugPrint(Jo.y) & " " & DebugPrint(Jo.z) & " " & DebugPrint(Lo, 10) & " " & DebugPrint(Oi)
 
    Da.X = Sqr((((Lq + Lo) * (Nq.X + Cq.X)) + (((O1.Y + O2.Y + O3.Y - Q1.Y + Q2.Y + Q3.Y) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.Y * Nq.z * Cq.Y) + (Q1.Y + Q2.Y + Q3.Y)))) ^ 2)
    Da.Y = Sqr((((Lq + Lo) * (Nq.Y + Cq.Y)) + (((O1.X + O2.X + O3.X - Q1.X + Q2.X + Q3.X) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.X * Nq.z * Cq.X) + (Q1.X + Q2.X + Q3.X)))) ^ 2)
    Da.z = Sqr((((Lq + Lo) * (Nq.z + Cq.z)) + (((O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z) + (O1.X + O2.X + O3.X - Q1.X + Q2.X + Q3.X)) * ((Nq.X * Nq.X * Cq.z) + (Q1.z + Q2.z + Q3.z)))) ^ 2)

    Qi = Sqr(((Da.X * Da.Y * Da.z) ^ 3 + ((Da.X + Da.Y + Da.z) * (Da.X + Da.Y + Da.z))))

    Jq.X = (((O1.X + O2.X + O3.X) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
    Jq.Y = (((O1.X + O2.X + O3.X) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
    Jq.z = (((O1.X + O2.X + O3.X) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
    
    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jq.x) & " " & DebugPrint(Jq.y) & " " & DebugPrint(Jq.z) & " " & DebugPrint(Lq, 10) & " " & DebugPrint(Qi)
    
    K = (Sqr(Distance(Jo.X, Jo.Y, Jo.z, Jq.X, Jq.Y, Jq.z) * Sqr(((Oi / 2) + (Qi / 2)))))
    
    'Debug.Print "Sect: " & CStr(K)
    
    TriangleIntersect = K
    
End Function


