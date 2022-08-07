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
    rhw As Single
    clr As Long
    tu As Single
    tv As Single
End Type

Public Type MyVertex
    X As Single
    Y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Public Const FVF_VERTEX_SIZE = 12
Public Const FVF_RENDER_SIZE = 32
Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1
Public Const Transparent As Long = &HFFFF00FF
Public Const MouseSensitivity As Long = 4
Public Const MaxDisplacement As Single = 0.05
Public Const BeaconSpacing As Single = 40
Public Const BeaconRange As Single = 1000
Public Const FadeDistance As Single = 800
Public Const SpaceBoundary As Single = 3000
Public Const HoursInOneDay As Single = 24
Public Const LetterPerInch As Single = 10

Public MaxCameraZoom As Single
Public MinCameraZoom As Single

Public Const FOVY As Single = 1.047198 '2.3561946
Public Const PI As Single = 3.14159265359
Public Const D90 As Single = PI / 4
Public Const D180 As Single = PI / 2
Public Const D360 As Single = PI
Public Const D720 As Single = PI * 2

Public Const Epsilon = 0.99999

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
    sBuffer = Replace(sBuffer, Chr(0), "")
    If sBuffer = "" Then sBuffer = "SHARED"
    GetUserLoginName = sBuffer
    
End Function

Public Function AppPath() As String

    Dim Ret As String
    Ret = IIf((Right(App.path, 1) = "\"), App.path, App.path & "\")
#If VBIDE = -1 Then
    If Right(Ret, 9) = "Projects\" Then
        Ret = Left(Ret, Len(Ret) - 9) & "Binary\"
     End If
#End If
    AppPath = Ret
End Function

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

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
        Else
            NextArg = Trim(TheParams)
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
        Else
            NextArg = TheParams
        End If
    End If
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator)))
        Else
            RemoveArg = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
        Else
            RemoveArg = ""
        End If
    End If
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function
Public Function NextQuotedArg(ByVal TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    NextQuotedArg = RemoveQuotedArg(TheParams, BeginQuote, EndQuote, Embeded, Compare)
End Function

Public Function RemoveLineArg(ByRef TheParams As Variant, Optional ByVal EndOfLine As String = " ") As String
    If (InStr(TheParams, EndOfLine) > 0) And (InStr(TheParams, vbCrLf) > 1) Then
        If InStr(TheParams, EndOfLine) < InStr(TheParams, vbCrLf) Then
            RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
        Else
            RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
        End If
        
    ElseIf (InStr(TheParams, vbCrLf) = 0) And (InStr(TheParams, EndOfLine) = 0) Then
        RemoveLineArg = TheParams
        TheParams = ""
    ElseIf (InStr(TheParams, vbCrLf) = 0) Then
        RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
    Else
        RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
    End If
End Function

Public Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim retVal As String
    Dim X As Long
    X = InStr(1, TheParams, BeginQuote, Compare)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + Len(BeginQuote), TheParams, EndQuote, Compare) > 0) Then
            If (Not Embeded) Or (EndQuote = BeginQuote) Then
                retVal = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retVal, InStr(1, retVal, EndQuote, Compare) + Len(EndQuote))
                retVal = Left(retVal, InStr(1, retVal, EndQuote, Compare) - 1)
            Else
                Dim l As Long
                Dim Y As Long
                l = 1
                Y = X
                Do Until l = 0
                    If (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare) > 0) And (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare) < InStr(Y + Len(BeginQuote), TheParams, EndQuote, Compare)) Then
                        l = l + 1
                        Y = InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare)
                    ElseIf (InStr(Y + Len(BeginQuote), TheParams, EndQuote, Compare) > 0) Then
                        l = l - 1
                        Y = InStr(Y + Len(EndQuote), TheParams, EndQuote, Compare)
                    Else
                        Y = Len(TheParams)
                        l = 0
                    End If
                Loop
                retVal = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retVal, (Y - X) + Len(EndQuote))
                retVal = Left(retVal, (Y - X) - 1)
            End If
        End If
    End If
    RemoveQuotedArg = retVal
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
    Dim Ret As Long
    On Error Resume Next
    Ret = GetAttr(URL)
    If Err.Number = 0 Then
        PathExists = IIf(IsFile, Not CBool(Ret And vbDirectory), True)
    Else
        Err.Clear
        PathExists = False
    End If
    On Error GoTo 0
End Function

Public Function ReadFile(ByVal path As String) As String
    If PathExists(path, True) Then
        Dim num As Integer
        Dim txt As String
        num = FreeFile
        Open path For Input Shared As #num
        Close #num
        Open path For Binary Shared As #num
            txt = String(FileSize(path), Chr(0))
            Get #num, 1, txt
        Close #num
        ReadFile = txt
    Else
        Err.Raise 53, App.EXEName, "File not found"
    End If
End Function

Public Sub WriteFile(ByVal path As String, ByVal Text As String)
    If PathExists(path, True) Then
        Kill path
    End If
    Dim num As Integer
    num = FreeFile
    Open path For Output Shared As #num
    Close #num
    num = FreeFile
    Open path For Binary Shared As #num
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
    Randomize
    RandomPositive = CSng((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function
Function MakeScreen(ByVal X As Single, ByVal Y As Single, ByVal z As Single, Optional ByVal tu As Single = 0, Optional ByVal tv As Single = 0) As MyScreen
    MakeScreen.X = X
    MakeScreen.Y = Y
    MakeScreen.z = z
    MakeScreen.rhw = 1
    MakeScreen.clr = D3DColorARGB(255, 255, 255, 255)
    MakeScreen.tu = tu
    MakeScreen.tv = tv
End Function
Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.z = z
End Function

Public Function SquareCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef V3 As D3DVECTOR) As D3DVECTOR

    SquareCenter.X = (v0.X + v1.X + v2.X + V3.X) / 4
    SquareCenter.Y = (v0.Y + v1.Y + v2.Y + V3.Y) / 4
    SquareCenter.z = (v0.z + v1.z + v2.z + V3.z) / 4
    
End Function

Public Function Distance(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Distance = Sqr(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.z - p2.z) ^ 2))
    'Distance = Sqr(((p1.X - p2.X) * (p1.X - p2.X)) + ((p1.Y - p2.Y) * (p1.Y - p2.Y)) + ((p1.z - p2.z) * (p1.z - p2.z)))
End Function
'
'
'Public Function PointOfDistanceToPoint(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByVal Distance As Single) As D3DVECTOR
'
'      PointOfDistanceToPoint.z = -(Sqr(-(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) - (Distance ^ 2))) - p1.z)
'      PointOfDistanceToPoint.Y = -(Sqr(-(((p1.X - p2.X) ^ 2) + ((p1.z - p2.z) ^ 2) - (Distance ^ 2))) - p1.Y)
'      PointOfDistanceToPoint.X = -(Sqr(-(((p1.Y - p2.Y) ^ 2) + ((p1.z - p2.z) ^ 2) - (Distance ^ 2))) - p1.X)
'
'    PointOfDistanceToPoint.z = (Sqr(-((((p1.X / (p1.Y / p2.z)) ^ 2) + ((p1.Y / (p1.z / p2.X) ^ 2)) - (Distance ^ 2))) + p2.z)
'    PointOfDistanceToPoint.Y = (Sqr(-((((p1.X / (p1.Y / p2.z)) ^ 2) + ((p1.z / (p1.X / p2.Y) ^ 2)) - (Distance ^ 2))) + p2.Y)
'    PointOfDistanceToPoint.X = (Sqr(-((((p1.Y / (p1.z / p2.X)) ^ 2) + ((p1.z / (p1.X / p2.Y) ^ 2)) - (Distance ^ 2))) + p2.X)
'
'
'End Function
Public Function CreateVertex(X As Single, Y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As MyVertex
    
    With CreateVertex
        .X = X: .Y = Y: .z = z
        .nx = nx: .ny = ny: .nz = nz
        .tu = tu: .tv = tv
    End With
    
End Function
Public Function ConvertVertexToVector(ByRef v As D3DVERTEX) As D3DVECTOR
    ConvertVertexToVector.X = v.X
    ConvertVertexToVector.Y = v.Y
    ConvertVertexToVector.z = v.z
End Function

Public Function VectorNormalize(ByRef v As D3DVECTOR) As D3DVECTOR
    Dim l As Single
    l = Sqr(v.X * v.X + v.Y * v.Y + v.z * v.z)
    If l = 0 Then l = 1
    VectorNormalize.X = (v.X / l)
    VectorNormalize.Y = (v.Y / l)
    VectorNormalize.z = (v.z / l)
End Function

Public Function VectorDotProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As Single
    VectorDotProduct = (u.X * v.X + u.Y * v.Y + u.z * v.z)
End Function

Public Function VectorCrossProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorCrossProduct.X = v.Y * u.z - v.z * u.Y
    VectorCrossProduct.Y = v.z * u.X - v.X * u.z
    VectorCrossProduct.z = v.X * u.Y - v.Y * u.X
End Function

Public Function VectorSubtract(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorSubtract.X = v.X - u.X
    VectorSubtract.Y = v.Y - u.Y
    VectorSubtract.z = v.z - u.z
End Function
Public Function VectorAdd(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorAdd.X = v.X + u.X
    VectorAdd.Y = v.Y + u.Y
    VectorAdd.z = v.z + u.z
End Function
Public Function VectorMultiply(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorMultiply.X = v.X * u.X
    VectorMultiply.Y = v.Y * u.Y
    VectorMultiply.z = v.z * u.z
End Function

Public Function PointToPlane(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Dim p3 As D3DVECTOR
    p3 = VectorSubtract(p1, p2)
    PointToPlane = Sqr(VectorDotProduct(p3, p3))
End Function


'Public Function TriangleNormal(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR
'
'    TriangleNormal.x = Sqr(((v0.Y - v1.Y) * (v0.z - v2.z) - (v0.z - v1.z) * (v0.Y - v2.Y)) ^ 2)
'    TriangleNormal.Y = Sqr(((v0.z - v1.z) * (v0.x - v2.x) - (v0.x - v1.x) * (v0.z - v2.z)) ^ 2)
'    TriangleNormal.z = Sqr(((v0.x - v1.x) * (v0.Y - v2.Y) - (v0.Y - v1.Y) * (v0.x - v2.x)) ^ 2)
'
'End Function

Public Function TriangleNormal(p0 As D3DVECTOR, p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR

    Dim v01 As D3DVECTOR
    Dim v02 As D3DVECTOR
    Dim vNorm As D3DVECTOR

    D3DXVec3Subtract v01, p1, p0
    D3DXVec3Subtract v02, p2, p0

    D3DXVec3Cross vNorm, v01, v02

    D3DXVec3Normalize vNorm, vNorm

    TriangleNormal = vNorm

End Function

Public Function TriangleCenter(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR

    Dim vR As D3DVECTOR
 
    vR.X = (v0.X + v1.X + v2.X) / 3
    vR.Y = (v0.Y + v1.Y + v2.Y) / 3
    vR.z = (v0.z + v1.z + v2.z) / 3
    
    TriangleCenter = vR

End Function

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

Public Function CreateMesh(ByVal FileName As String, Mesh As D3DXMesh, Buffer As D3DXBuffer, Origin As D3DVECTOR, Scaled As D3DVECTOR, MeshMaterials() As D3DMATERIAL8, MeshTextures() As Direct3DTexture8, MeshVerticies() As D3DVERTEX, MeshIndicies() As Integer, nMaterials As Long)
    Dim TextureName As String

    Set Mesh = D3DX.LoadMeshFromX(FileName, D3DXMESH_DYNAMIC, DDevice, Nothing, Buffer, nMaterials)

    If nMaterials > 0 Then
    
        ReDim MeshMaterials(0 To nMaterials - 1) As D3DMATERIAL8
        ReDim MeshTextures(0 To nMaterials - 1) As Direct3DTexture8
    
        Dim d As ImgDimType
        
        Dim q As Integer
        For q = 0 To nMaterials - 1
    
            D3DX.BufferGetMaterial Buffer, q, MeshMaterials(q)
       
            TextureName = D3DX.BufferGetTextureName(Buffer, q)
            If (TextureName <> "") Then
                If ImageDimensions(AppPath & "Models\" & TextureName, d) Then
                    Set MeshTextures(q) = D3DX.CreateTextureFromFileEx(DDevice, AppPath & "Models\" & TextureName, d.width, d.height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Transparent, ByVal 0, ByVal 0)
                Else
                    Debug.Print "IMAGE ERROR: ImageDimensions - " & AppPath & "Models\" & TextureName
                End If
            End If
            
        Next
    Else
        ReDim MeshTextures(0 To 0) As Direct3DTexture8
        ReDim MeshMaterials(0 To 0) As D3DMATERIAL8
    End If
    
    Dim vd As D3DVERTEXBUFFER_DESC
    Mesh.GetVertexBuffer.GetDesc vd

    ReDim MeshVerticies(0 To ((vd.Size \ FVF_VERTEX_SIZE) - 1)) As D3DVERTEX
    D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, vd.Size, 0, MeshVerticies(0)

    Dim id As D3DINDEXBUFFER_DESC
    Mesh.GetIndexBuffer.GetDesc id

    ReDim MeshIndicies(0 To ((id.Size \ 2) - 1)) As Integer
    D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, id.Size, 0, MeshIndicies(0)
    
    D3DX.ComputeNormals Mesh

End Function

Public Function GetNow() As String

    GetNow = CStr(Now)

End Function

Public Function GetTimer() As String

    GetTimer = CStr(CDbl(Timer))
    
End Function

Public Sub SwapSingle(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function ClassifyPoint(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef p As D3DVECTOR) As Single
    Dim dtp As Single
    Dim n As D3DVECTOR
    n = GetPlaneNormal(v0, v1, v2)
    dtp = VectorDotProduct(n, p) + -VectorDotProduct(n, v0)
    
    If dtp > Epsilon Then
        ClassifyPoint = 1 'front
    ElseIf dtp < -Epsilon Then
        ClassifyPoint = 2 'back
    Else
        ClassifyPoint = 0 'onplane
    End If
End Function

'Public Function PointBehindtriangle(ByRef Point As D3DVECTOR, ByRef center As D3DVECTOR, ByRef Lengths As D3DVECTOR) As Boolean
'
'' (GreatestX(v1, v2, v3) - LeastX(v1, v2, v3)) / 2, _
''                                (GreatestY(v1, v2, v3) - LeastY(v1, v2, v3)) / 2, _
''                                (GreatestZ(v1, v2, v3) - LeastZ(v1, v2, v3)) / 2, _
''                                Distance(v1, v2), Distance(v2, v3), Distance(v3, v1),
'
'    Dim v As D3DVECTOR
'    Dim u As D3DVECTOR
'    Dim n As D3DVECTOR
'
'    Dim d As Single
'
'    u = VectorCrossProduct(Point, center)
'
'    v = VectorSubtract(Point, center)
'    v = VectorSubtract(v, center)
'    v = VectorCrossProduct(v, center)
'
'    n.x = (((u.x + u.Y) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.x + v.Y) * (Lengths.x + Lengths.Y + Lengths.z)))
'    n.Y = (((u.x + u.z) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.x + v.z) * (Lengths.x + Lengths.Y + Lengths.z)))
'    n.x = (((u.Y + u.x) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.Y + v.x) * (Lengths.x + Lengths.Y + Lengths.z)))
'
'    d = Distance(u, v)
'
'    n.x = (n.x * u.x) / d
'    n.Y = (n.Y * u.Y) / d
'    n.z = (n.z * u.z) / d
'
'    'Debug.Print "n: " & n.x & "," & n.y & "," & n.z & "  u: " & u.x & "," & u.y & "," & u.z & "  v: " & v.x & "," & v.y & "," & v.z & "  d: " & d
'
'    PointBehindtriangle = n.x + n.Y + n.z > 0
'End Function
'
'Public Function TriangleIntersect(ByRef O1 As D3DVECTOR, ByRef O2 As D3DVECTOR, ByRef O3 As D3DVECTOR, ByRef Q1 As D3DVECTOR, ByRef Q2 As D3DVECTOR, ByRef Q3 As D3DVECTOR) As Long
'
'    Dim No As D3DVECTOR
'    Dim Nq As D3DVECTOR
'    Dim Co As D3DVECTOR
'    Dim Cq As D3DVECTOR
'    Dim Lc As Single
'    Dim Lo As Single
'    Dim Lq As Single
'    Dim Da As D3DVECTOR
'    Dim Oi As Single
'    Dim Qi As Single
'    Dim Jo As D3DVECTOR
'    Dim Jq As D3DVECTOR
'    Dim K As Integer
'
'    No = TriangleNormal(O1, O2, O3)
'    Nq = TriangleNormal(Q1, Q2, Q3)
'    Co = TriangleCenter(O1, O2, O3)
'    Cq = TriangleCenter(Q1, Q2, Q3)
'    Lc = Distance(Co, Cq)
'    Lo = Distance(O1, O2) + Distance(O2, O3) + Distance(O3, O1)
'    Lq = Distance(Q1, Q2) + Distance(Q2, Q3) + Distance(Q3, Q1)
'
'    Da.x = Sqr((((Lo + Lq) * (No.x + Co.x)) + (((Q1.Y + Q2.Y + Q3.Y - O1.Y + O2.Y + O3.Y) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.Y * No.z * Co.Y) + (O1.Y + O2.Y + O3.Y)))) ^ 2)
'    Da.Y = Sqr((((Lo + Lq) * (No.Y + Co.Y)) + (((Q1.x + Q2.x + Q3.x - O1.x + O2.x + O3.x) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.x * No.z * Co.x) + (O1.x + O2.x + O3.x)))) ^ 2)
'    Da.z = Sqr((((Lo + Lq) * (No.z + Co.z)) + (((Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z) + (Q1.x + Q2.x + Q3.x - O1.x + O2.x + O3.x)) * ((No.x * No.x * Co.z) + (O1.z + O2.z + O3.z)))) ^ 2)
'
'    Oi = Sqr(((Da.x * Da.Y * Da.z) ^ 3 + ((Da.x + Da.Y + Da.z) * (Da.x + Da.Y + Da.z))))
'
'    Jo.x = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'    Jo.Y = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'    Jo.z = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'
'    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jo.x) & " " & DebugPrint(Jo.y) & " " & DebugPrint(Jo.z) & " " & DebugPrint(Lo, 10) & " " & DebugPrint(Oi)
'
'    Da.x = Sqr((((Lq + Lo) * (Nq.x + Cq.x)) + (((O1.Y + O2.Y + O3.Y - Q1.Y + Q2.Y + Q3.Y) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.Y * Nq.z * Cq.Y) + (Q1.Y + Q2.Y + Q3.Y)))) ^ 2)
'    Da.Y = Sqr((((Lq + Lo) * (Nq.Y + Cq.Y)) + (((O1.x + O2.x + O3.x - Q1.x + Q2.x + Q3.x) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.x * Nq.z * Cq.x) + (Q1.x + Q2.x + Q3.x)))) ^ 2)
'    Da.z = Sqr((((Lq + Lo) * (Nq.z + Cq.z)) + (((O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z) + (O1.x + O2.x + O3.x - Q1.x + Q2.x + Q3.x)) * ((Nq.x * Nq.x * Cq.z) + (Q1.z + Q2.z + Q3.z)))) ^ 2)
'
'    Qi = Sqr(((Da.x * Da.Y * Da.z) ^ 3 + ((Da.x + Da.Y + Da.z) * (Da.x + Da.Y + Da.z))))
'
'    Jq.x = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'    Jq.Y = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'    Jq.z = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'
'    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jq.x) & " " & DebugPrint(Jq.y) & " " & DebugPrint(Jq.z) & " " & DebugPrint(Lq, 10) & " " & DebugPrint(Qi)
'
'    K = (Sqr(Distance(Jo, Jq) * Sqr(((Oi / 2) + (Qi / 2)))))
'
'    'Debug.Print "Sect: " & CStr(K)
'
'    TriangleIntersect = K
'
'End Function
'



''procedure 3Dto2D (x, y, z, pan, centre, position)
''X = X + position.X
''Y = Y + position.Y
''Z = Z + position.Z
''new.x = x*cos(pan.x) - z*sin(pan.x)
''new.z = x*sin(pan.x) + z*cos(pan.x)
''new.y = y*cos(pan.y) - new.z*sin(pan.y)
''z = new.y*cos(pan.y) - new.z*sin(pan.y)
''x = new.x*cos(pan.z) - new.y*sin(pan.z)
''y = new.x*sin(pan.z) + new.y*cos(pan.z)
''If Z > 0 Then
''    Screen.X = X / Z * Zoom + centre.X
''    Screen.Y = Y / Z * Zoom + centre.Y
''End If

'Public Function ScreenVertex(ByVal VertexIndex As Long, ByRef ScreenX As Single, ByRef ScreenY As Single, ByRef ScreenZ As Single)
'    Dim r As D3DVECTOR 'VRP
'    Dim p As D3DVECTOR 'WCP
'    Dim n As D3DVECTOR
'    Dim up As D3DVECTOR
'    Dim v As D3DVECTOR
'    Dim u As D3DVECTOR
'    Dim dot As Single
'    Dim dif As D3DVECTOR
'
'
'    GetVertex VertexIndex, p
'    FR_Camera.GetPosition grid.Direct3DFrame, r
'    FR_Camera.GetOrientation grid.Direct3DFrame, n, up
'    DX_Main.VectorNormalize n
'
'
'    dot = DX_Main.VectorDotProduct(up, n)
'    v.X = up.X - dot * n.X
'    v.Y = up.Y - dot * n.Y
'    v.z = up.z - dot * n.z
'    DX_Main.VectorNormalize v
'
'
'    DX_Main.VectorCrossProduct u, n, v
'
'
'    DX_Main.VectorSubtract dif, p, r
'
'
'    ScreenX = DX_Main.VectorDotProduct(dif, u) + (ScreenWidth / 2)
'    ScreenY = DX_Main.VectorDotProduct(dif, v) + (ScreenHeight / 2)
'    ScreenZ = DX_Main.VectorDotProduct(dif, n)
'
'
'End Function
'
''Public Property Get ScreenX(ByVal VertexIndex As Integer) As Single
''    On Error Resume Next
''
''    Dim vert As D3DVECTOR
''    Dim vertA As D3DVECTOR
''    Dim vertB As D3DVECTOR
''
''    Dim cVert As D3DVECTOR
''    Dim cRot As D3DVECTOR
''    Dim cData As Single
''
''    FR_Camera.GetRotation FR_Root, cRot, cData
''    FR_Camera.GetPosition FR_Root, cVert
''
''    GetVertex VertexIndex, vert
''
''    'ScreenX = vert.X + (ScreenWidth / 2)
''
''    DX_Main.VectorAdd vertA, cVert, vert
''
''    DX_Main.VectorRotate vertB, vertA, cRot, cData
''
''
''    ScreenX = (vertB.X / vertB.Z * 1000) + (ScreenWidth / 2)    'Zoom + centre.X
''
''    'ScreenX = (vert.X / vert.Z * 100) + (ScreenWidth / 2)  'Zoom + centre.X
''End Property
''Public Property Get ScreenY(ByVal VertexIndex As Integer) As Single
''    On Error Resume Next
''
''    Dim vert As D3DVECTOR
''    Dim vertA As D3DVECTOR
''    Dim vertB As D3DVECTOR
''
''    Dim cVert As D3DVECTOR
''    Dim cRot As D3DVECTOR
''    Dim cData As Single
''
''    FR_Camera.GetRotation FR_Root, cRot, cData
''    FR_Camera.GetPosition FR_Root, cVert
''
''    GetVertex VertexIndex, vert
''
''    'ScreenY = vert.Y + (ScreenWidth / 2)
''
''    DX_Main.VectorAdd vertA, cVert, vert
''
''    DX_Main.VectorRotate vertB, vertA, cRot, cData
''
''
''    ScreenY = (vertB.Y / vertB.Z * 1000) + (ScreenHeight / 2)     'Zoom + centre.X
''
''    'ScreenY = (vert.Y / vert.Z * 100) + (ScreenHeight / 2)   'Zoom + centre.X
''End Property

