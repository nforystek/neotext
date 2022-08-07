Attribute VB_Name = "modDecs"
#Const modDecs = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Const WithInCityLimits As Long = 5000
Public Const Transparent As Long = &HFFFF00FF
Public Const MouseSensitivity As Long = 4
Public Const MaxDisplacement As Single = 30
Public Const FadeDistance As Single = 30000

Public Const FrameScale As Single = 246

Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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
    If sBuffer = "" Then sBuffer = "SHARED"
    sBuffer = Replace(sBuffer, Chr(0), "")
    GetUserLoginName = sBuffer
    
End Function

Public Function AppPath() As String
    Dim ret As String
    ret = IIf((Right(App.Path, 1) = "\"), App.Path, App.Path & "\")
#If VBIDE = -1 Then
    If Right(ret, 9) = "Projects\" Then
        ret = Left(ret, Len(ret) - 9) & "Binary\"
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

Public Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    Dim RetVal As String
    Dim X As Long
    X = InStr(TheParams, BeginQuote)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + 1, TheParams, EndQuote) > 0) Then
            RetVal = Mid(TheParams, X + 1)
            TheParams = Left(TheParams, X - 1) & Mid(RetVal, InStr(RetVal, EndQuote) + 1)
            RetVal = Left(RetVal, InStr(RetVal, EndQuote) - 1)
        End If
    End If
    RemoveQuotedArg = RetVal
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

Public Sub WriteFile(ByVal Path As String, ByVal text As String)
    If PathExists(Path, True) Then
        Kill Path
    End If
    Dim num As Integer
    num = FreeFile
    Open Path For Output Shared As #num
    Close #num
    num = FreeFile
    Open Path For Binary Shared As #num
        Put #num, 1, text
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
    Randomize
    RandomPositive = CSng((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
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

Public Function CreateVertex(X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As TVERTEX2
    
    With CreateVertex
        .X = X: .Y = Y: .Z = Z
        .nx = nx: .ny = ny: .nz = nz
        .tu = tu: .tv = tv
    End With
    
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

Public Sub CreateSquare(ByRef Data() As TVERTEX2, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef p4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.Z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.Z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.Z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, 0, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.Z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.Z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.Z
    
End Sub

Public Sub CreateSquareEx(ByRef Data() As TVERTEX2, ByVal Index As Long, ByRef p1 As TVERTEX0, ByRef p2 As TVERTEX0, ByRef p3 As TVERTEX0, ByRef p4 As TVERTEX0)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, p2.tu, p2.tv)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, p3.tu, p3.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.Z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.Z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.Z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, p3.tu, p3.tv)
    Data(Index + 5) = CreateVertex(p4.X, p4.Y, p4.Z, 0, 0, 0, p4.tu, p4.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.Z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.Z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.Z
    
End Sub

Public Function GetNow() As String

    GetNow = CStr(Now)

End Function

Public Function GetTimer() As String

    GetTimer = CStr(CDbl(Timer))

End Function

Public Function GetElapsed() As String

    Dim ret As Single
    ret = Round((Timer - Level.Elapsed), 2)
    
    If InStr(Trim(CStr(ret)), ".") = 0 Then
        GetElapsed = Trim(CStr(ret)) & ".00"
    ElseIf Len(Mid(Trim(CStr(ret)), InStr(Trim(CStr(ret)), ".") + 1)) = 1 Then
        GetElapsed = Trim(CStr(ret)) & "0"
    Else
        GetElapsed = Trim(CStr(ret))
    End If
    
End Function

Public Sub Swap(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function GreatestX(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the greatext x value of three vectors
    If v0.X > v1.X And v0.X > v2.X And v0.X > v3.X Then
        GreatestX = v0.X
    ElseIf v1.X > v0.X And v1.X > v2.X And v1.X > v3.X Then
        GreatestX = v1.X
    ElseIf v2.X > v0.X And v2.X > v1.X And v2.X > v3.X Then
        GreatestX = v2.X
    Else
        GreatestX = v3.X
    End If
End Function

Public Function GreatestY(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the greatext y value of three vectors
    If v0.Y >= v1.Y And v0.Y >= v2.Y And v0.Y >= v3.Y Then
        GreatestY = v0.Y
    ElseIf v1.Y >= v0.Y And v1.Y >= v2.Y And v1.Y >= v3.Y Then
        GreatestY = v1.Y
    ElseIf v2.Y >= v0.Y And v2.Y >= v1.Y And v2.Y >= v3.Y Then
        GreatestY = v2.Y
    Else
        GreatestY = v3.Y
    End If
End Function

Public Function GreatestZ(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the greatext z value of three vectors
    If v0.Z >= v1.Z And v0.Z >= v2.Z And v0.Z >= v3.Z Then
        GreatestZ = v0.Z
    ElseIf v1.Z >= v0.Z And v1.Z >= v2.Z And v1.Z >= v3.Z Then
        GreatestZ = v1.Z
    ElseIf v2.Z >= v0.Z And v2.Z >= v1.Z And v2.Z >= v3.Z Then
        GreatestZ = v2.Z
    Else
        GreatestZ = v3.Z
    End If
End Function

Public Function LeastX(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the least x value of three vectors
    If v0.X <= v1.X And v0.X <= v2.X And v0.X <= v3.X Then
        LeastX = v0.X
    ElseIf v1.X <= v0.X And v1.X <= v2.X And v1.X <= v3.X Then
        LeastX = v1.X
    ElseIf v2.X <= v0.X And v2.X <= v1.X And v2.X <= v3.X Then
        LeastX = v2.X
    Else
        LeastX = v3.X
    End If
End Function

Public Function LeastY(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the least y value of three vectors
    If v0.Y <= v1.Y And v0.Y <= v2.Y And v0.Y <= v3.Y Then
        LeastY = v0.Y
    ElseIf v1.Y <= v0.Y And v1.Y <= v2.Y And v1.Y <= v3.Y Then
        LeastY = v1.Y
    ElseIf v2.Y <= v0.Y And v2.Y <= v1.Y And v2.Y <= v3.Y Then
        LeastY = v2.Y
    Else
        LeastY = v3.Y
    End If
End Function

Public Function LeastZ(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As Single
    'return the least z value of three vectors
    If v0.Z <= v1.Z And v0.Z <= v2.Z And v0.Z <= v3.Z Then
        LeastZ = v0.Z
    ElseIf v1.Z <= v0.Z And v1.Z <= v2.Z And v1.Z <= v3.Z Then
        LeastZ = v1.Z
    ElseIf v2.Z <= v0.Z And v2.Z <= v1.Z And v2.Z <= v3.Z Then
        LeastZ = v2.Z
    Else
        LeastZ = v3.Z
    End If
End Function

Private Function VectorNormalize(ByRef v As D3DVECTOR) As D3DVECTOR
    Dim l As Single
    l = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z)
    If l = 0 Then l = 1
    VectorNormalize.X = (v.X / l)
    VectorNormalize.Y = (v.Y / l)
    VectorNormalize.Z = (v.Z / l)
End Function

Private Function VectorDotProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As Single
    VectorDotProduct = (u.X * v.X + u.Y * v.Y + u.Z * v.Z)
End Function

Private Function VectorCrossProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorCrossProduct.X = v.Y * u.Z - v.Z * u.Y
    VectorCrossProduct.Y = v.Z * u.X - v.X * u.Z
    VectorCrossProduct.Z = v.X * u.Y - v.Y * u.X
End Function

Private Function VectorSubtract(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorSubtract.X = v.X - u.X
    VectorSubtract.Y = v.Y - u.Y
    VectorSubtract.Z = v.Z - u.Z
End Function
Private Function VectorAdd(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorAdd.X = v.X + u.X
    VectorAdd.Y = v.Y + u.Y
    VectorAdd.Z = v.Z + u.Z
End Function
Private Function VectorMultiply(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorMultiply.X = v.X * u.X
    VectorMultiply.Y = v.Y * u.Y
    VectorMultiply.Z = v.Z * u.Z
End Function
