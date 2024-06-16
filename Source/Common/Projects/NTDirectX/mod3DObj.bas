Attribute VB_Name = "mod3DObj"
Option Explicit

'############################################################################################################
'Derived Exports ############################################################################################
'############################################################################################################
                                    
Private Declare Function Collision Lib "MaxLandLib.dll" _
                                    (ByVal lngStreamFlagValue As _
                                    Long, ByVal lngTotalTriangles As _
                                    Long, sngTriangleFaceData() As _
                                    Single, sngVertexXAxisData() As _
                                    Single, sngVertexYAxisData() As _
                                    Single, sngVertexZAxisData() As _
                                    Single, ByVal lngTriangleToCheck As _
                                    Long, ByRef lngReturnHitObject As _
                                    Long, ByRef lngReturnHitTriangle As Long) As Boolean

'############################################################################################################
'Variable Declare ###########################################################################################
'############################################################################################################

Public ObjectCount As Long
Public TriangleCount As Long
Public TriangleFace() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'TriangleFace dimension (n,t) where n=0 is x of the face normal
'TriangleFace dimension (n,t) where n=1 is y of the face normal
'TriangleFace dimension (n,t) where n=2 is z of the face normal
'TriangleFace dimension (n,t) where n=3 flag for 1st arg of collision(), culling
'TriangleFace dimension (n,t) where n=4 is the object index
'TriangleFace dimension (n,t) where n=5 is the face index

Public VertexXAxis() As Single
Public VertexYAxis() As Single
Public VertexZAxis() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'VertexXAxis dimension (n,t) where n=0 is X of the first vertex
'VertexXAxis dimension (n,t) where n=1 is X of the second vertex
'VertexXAxis dimension (n,t) where n=2 is X of the third vertex
'VertexYAxis dimension (n,t) where n=0 is Y of the first vertex
'VertexYAxis dimension (n,t) where n=1 is Y of the second vertex
'VertexYAxis dimension (n,t) where n=2 is Y of the third vertex
'VertexZAxis dimension (n,t) where n=0 is Z of the first vertex
'VertexZAxis dimension (n,t) where n=1 is Z of the second vertex
'VertexZAxis dimension (n,t) where n=2 is Z of the third vertex

Public VertexDirectX() As MyVertex
Public ScreenDirectX() As MyScreen

Public Points As Points
'Public Rotates As Orbit 'waiting to be applied rotates
'Public Scalars As Orbit 'waiting to be applied scalars
Public Zero As New Point
Public PlayerGyro As New Point
Public PlanetGyro As New Point
Public Localized As Point

'Public Orbits As Orbits 'collection of in non script accessable for the app for all orbits and implmenets of
'Public Ranges As Ranges 'collection of those that are part of the planet object only needed in global cycle
'Public Points As Points 'cache of all points uniquely, so far this just grows and shouldn't accept change them

Public Sub CleanUpObjs()

    Set Localized = Nothing
    Set Points = Nothing
    
    ObjectCount = 0
    TriangleCount = 0
    Erase TriangleFace
    
    Erase VertexXAxis
    Erase VertexYAxis
    Erase VertexZAxis

    Erase VertexDirectX


End Sub

Public Sub CreateObjs()

    Set Points = New Points
    Set Localized = New Point
         
End Sub

'Public Function SineAngle(ByVal sin0 As Single) As Single
'    SineAngle = Round((InvSin(sin0) * DEGREE + 2), 6)
'End Function

'Public Function CosineAngle(ByVal cos0 As Single) As Single
'    CosineAngle = (90 - Round((InvSin(1 - cos0) * DEGREE + 2), 6))
'End Function
'
'
'Public Sub SinID(ByVal p As Point, ByRef sin0 As Single)
'    If p.y > 0 Then
'        If p.x = 0 Then
'            sin0 = 1
'        ElseIf sin0 < 0 Then
'            sin0 = -sin0
'        End If
'    ElseIf p.y < 0 Then
'        If p.x = 0 Then
'            sin0 = -1
'        ElseIf sin0 > 0 Then
'            sin0 = -sin0
'        End If
'    ElseIf p.x <> 0 Then
'        sin0 = 0
'    End If
'End Sub
'
'Public Sub CosID(ByVal p As Point, ByRef cos0 As Single)
'    If p.x > 0 Then
'        If p.y = 0 Then
'            cos0 = 1
'        ElseIf cos0 < 0 Then
'            cos0 = -cos0
'        End If
'    ElseIf p.x < 0 Then
'        If p.y = 0 Then
'            cos0 = -1
'        ElseIf cos0 > 0 Then
'            cos0 = -cos0
'        End If
'    ElseIf p.y <> 0 Then
'        cos0 = 0
'    End If
'End Sub
'
'Public Sub TanID(ByVal p As Point, ByRef tan0 As Single)
'    If p.x = 0 And p.y <> 0 Then
'        'tan0 = CVErr(0)
'    ElseIf p.y = 0 And p.x <> 0 Then
'        tan0 = 0
'    ElseIf (p.x > 0 And p.y > 0) Or (p.x < 0 And p.y < 0) Then
'        If tan0 < 0 Then tan0 = -tan0
'    ElseIf (p.x < 0 And p.y > 0) Or (p.x > 0 And p.y < 0) Then
'        If tan0 > 0 Then tan0 = -tan0
'    End If
'End Sub
'
'Public Sub SecID(ByVal p As Point, ByRef sec0 As Single)
'    If p.x = 0 Then
'       ' sec0 = CVErr(0)
'    ElseIf p.y = 0 And p.x > 0 Then
'        sec0 = 1
'    ElseIf p.y = 0 And p.x < 0 Then
'        sec0 = -1
'    ElseIf p.x > 0 And p.y <> 0 Then
'        If sec0 < 0 Then sec0 = -sec0
'    ElseIf p.x < 0 And p.y <> 0 Then
'        If sec0 > 0 Then sec0 = -sec0
'    End If
'End Sub
'
'Public Sub CscID(ByVal p As Point, ByRef csc0 As Single)
'    If p.y = 0 Then
'        'csc0 = CVErr(0)
'    ElseIf p.x = 0 And p.y > 0 Then
'        csc0 = 1
'    ElseIf p.x = 0 And p.y < 0 Then
'        csc0 = -1
'    ElseIf p.y > 0 And p.x <> 0 Then
'        If csc0 < 0 Then csc0 = -csc0
'    ElseIf p.y < 0 And p.x <> 0 Then
'        If csc0 > 0 Then csc0 = -csc0
'    End If
'End Sub
'
'Public Sub CotID(ByVal p As Point, ByRef cot0 As Single)
'    If p.y = 0 And p.x <> 0 Then
'        'cot0 = CVErr(0)
'    ElseIf p.x = 0 And p.y <> 0 Then
'        cot0 = 0
'    ElseIf (p.x > 0 And p.y > 0) Or (p.x < 0 And p.y < 0) Then
'        If cot0 < 0 Then cot0 = -cot0
'    ElseIf (p.x < 0 And p.y > 0) Or (p.x > 0 And p.y < 0) Then
'        If cot0 > 0 Then cot0 = -cot0
'    End If
'End Sub
'
'Public Function VectorSecant(ByRef p As Point) As Single
'    VectorSecant = Abs(VectorCosine(p))
'    If VectorSecant <> 0 Then VectorSecant = (1 / VectorSecant)
'    SecID p, VectorSecant
'End Function
'Public Function VectorCosecant(ByRef p As Point) As Single
'    VectorCosecant = Abs(VectorCosine(p))
'    If VectorCosecant <> 0 Then VectorCosecant = (1 / VectorCosecant)
'    CscID p, VectorCosecant
'End Function
'Public Function VectorCotangent(ByRef p As Point) As Single
'    VectorCotangent = Abs(VectorTangent(p))
'    If VectorCotangent <> 0 Then VectorCotangent = (1 / VectorCotangent)
'    CotID p, VectorCotangent
'End Function
'
'Public Function VectorTangent(ByRef p As Point) As Single
'    'returns the z axis angle of the x and y in p
'    If p.x = 0 Then
'        If p.y > 0 Then
'            VectorTangent = Val("1.#IND")
'        ElseIf p.y < 0 Then
'            VectorTangent = 1
'        End If
'    ElseIf (p.y <> 0) Then
'        VectorTangent = Round(Abs(p.y / p.x), 2)
'    End If
'    TanID p, VectorTangent
'End Function
'
'Public Function VectorSine(ByRef p As Point) As Single
'    'returns the z axis angle of the x and y in p
'    If p.x = 0 Then
'        If p.y <> 0 Then
'            VectorSine = Val("0.#IND")
'        End If
'    ElseIf p.y <> 0 Then
'        VectorSine = Round(Abs(p.y / Distance(0, 0, 0, p.x, p.y, 0)), 2)
'    End If
'    SinID p, VectorSine
'End Function
'
'Public Function VectorCosine(ByRef p As Point) As Single
'    'returns the z axis angle of the x and y in p
'    If p.y = 0 Then
'        If p.x <> 0 Then
'            VectorCosine = Val("1.#IND")
'        End If
'    ElseIf p.x <> 0 Then
'        VectorCosine = Round(Abs(p.x / Distance(0, 0, 0, p.x, p.y, 0)), 2)
'    End If
'    CosID p, VectorCosine
'End Function

'Public Function tan0(ByVal sin0 As Single, ByVal cos0 As Single) As Single
'    If cos0 <> 0 Then tan0 = (sin0 / cos0)
'End Function
'Public Function cos0(ByVal sin0 As Single) As Single
'    cos0 = (1 - sin0)
'End Function
'
'Public Function Sec(ByVal x As Single) As Single
'    Sec = Cos(x)
'    If Sec <> 0 Then Sec = 1 / Cos(x)
'End Function
'Public Function Csc(ByVal x As Single) As Single
'    Csc = Sin(x)
'    If Csc <> 0 Then Csc = 1 / Sin(x)
'End Function
'Public Function CTan(ByVal x As Single) As Single
'    CTan = Tan(x)
'    If CTan <> 0 Then CTan = 1 / Tan(x)
'End Function
'Public Function Asin(ByVal x As Single) As Single
'    Asin = Sqr(-x * x + 1)
'    If Asin <> 0 Then Asin = Atn(x / Asin)
'End Function
'
'Public Function Acos(ByVal x As Single) As Single
'    Acos = Sqr(-x * x + 1)
'    If Acos <> 0 Then Acos = Atn(-x / Acos) + 2 * Atn(1)
'End Function
'
'Public Function Asec(ByVal x As Single) As Single
'    Asec = Sqr(x * x - 1)
'    If Asec <> 0 Then Asec = 2 * Atn(1) - Atn(Sign(x) / Asec)
'End Function
'
'Public Function Acsc(ByVal x As Single) As Single
'    Asec = Sqr(x * x - 1)
'    If Asec <> 0 Then Acsc = Atn(Sign(x) / Asec)
'End Function
'
'Public Function Acot(ByVal x As Single) As Single
'    Acot = 2 * Atn(1) - Atn(x)
'End Function
'
'Public Function Sinh(ByVal x As Single) As Single
'    Sinh = (Exp(x) - Exp(-x)) / 2
'End Function
'
'Public Function Cosh(ByVal x As Single) As Single
'    Cosh = (Exp(x) + Exp(-x)) / 2
'End Function
'
'Public Function Tanh(ByVal x As Single) As Single
'    Tanh = (Exp(x) + Exp(-x))
'    If Tanh <> 0 Then Tanh = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
'End Function
'
'
'Public Function Sech(ByVal x As Single) As Single
'    Sech = (Exp(x) + Exp(-x))
'    If Sech <> 0 Then Sech = 2 / (Exp(x) + Exp(-x))
'End Function
'
'
'Public Function Csch(ByVal x As Single) As Single
'    Csch = (Exp(x) - Exp(-x))
'    If Csch <> 0 Then Csch = 2 / (Exp(x) - Exp(-x))
'End Function
'
'
'Public Function Coth(ByVal x As Single) As Single
'    Coth = (Exp(x) - Exp(-x))
'    If Coth <> 0 Then Coth = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
'End Function
'
'
'Public Function Asinh(ByVal x As Single) As Single
'    Asinh = Log(x + Sqr(x * x + 1))
'End Function
'
'
'Public Function Acosh(ByVal x As Single) As Single
'    Acosh = Log(x + Sqr(x * x - 1))
'End Function
'
'
'Public Function Atanh(ByVal x As Single) As Single
'    If (1 - x) <> 0 Then Atanh = Log((1 + x) / (1 - x)) / 2
'End Function
'
'
'Public Function Asech(ByVal x As Single) As Single
'    If x <> 0 Then Asech = Log((Sqr(-x * x + 1) + 1) / x)
'End Function
'
'
'Public Function Acsch(ByVal x As Single) As Single
'    If x <> 0 Then Acsch = Log((Sign(x) * Sqr(x * x + 1) + 1) / x)
'End Function
'
'
'Public Function Acoth(ByVal x As Single) As Single
'    Acoth = Log((x + 1) / (x - 1)) / 2
'End Function
'
'
'Public Function SinDouble(ByVal sin0 As Single, ByVal cos0 As Single) As Single
'    SinDouble = (2 * sin0 * cos0)
'End Function
'Public Function CosDouble(ByVal sin0 As Single, ByVal cos0 As Single) As Single
'    CosDouble = (cos0 ^ 2 - sin0 ^ 2)
'End Function
'Public Function TanDouble(ByVal tan0 As Single) As Single
'    TanDouble = (1 - tan0 ^ 2)
'    If TanDouble <> 0 Then TanDouble = ((2 * tan0) / TanDouble)
'End Function
'
'Public Function SinHalf(ByVal cos0 As Single) As Single
'    SinHalf = (((1 - cos0) / 2) ^ (1 / 2))
'End Function
'Public Function CosHalf(ByVal cos0 As Single) As Single
'    CosHalf = (((1 + cos0) / 2) ^ (1 / 2))
'End Function
'Public Function TanHalf(ByVal sin0 As Single, ByVal cos0 As Single) As Single
'    TanHalf = (1 + cos0)
'    If TanHalf <> 0 Then TanHalf = (sin0 / TanHalf)
'End Function
'
'Public Function SinSum(ByVal sinA As Single, ByVal cosA As Single, ByVal sinB As Single, ByVal cosB As Single) As Single
'    SinSum = (sinA * cosB + cosA * sinB)
'End Function
'Public Function CosSum(ByVal sinA As Single, ByVal cosA As Single, ByVal sinB As Single, ByVal cosB As Single) As Single
'    CosSum = (cosA * cosB - sinA * sinB)
'End Function
'Public Function TanSum(ByVal tanA As Single, ByVal tanB As Single) As Single
'    TanSum = (1 - tanA * tanB)
'    If TanSum <> 0 Then TanSum = ((tanA + tanB) / TanSum)
'End Function
'
'Public Function SinDiff(ByVal sinA As Single, ByVal cosA As Single, ByVal sinB As Single, ByVal cosB As Single) As Single
'    SinDiff = (sinA * cosB - cosA * sinB)
'End Function
'Public Function CosDiff(ByVal sinA As Single, ByVal cosA As Single, ByVal sinB As Single, ByVal cosB As Single) As Single
'    CosDiff = (cosA * cosB + sinA * sinB)
'End Function
'Public Function TanDiff(ByVal tanA As Single, ByVal tanB As Single) As Single
'    TanDiff = (1 + tanA * tanB)
'    If TanDiff <> 0 Then TanDiff = ((tanA - tanB) / TanDiff)
'End Function
'
'Public Function SinTriple(ByVal sin0 As Single) As Single
'    SinTriple = (3 * sin0 - 4 * sin0 ^ 3)
'End Function
'Public Function CosTriple(ByVal cos0 As Single) As Single
'    CosTriple = (4 * cos0 ^ 3 - 3 * cos0)
'End Function
'Public Function TanTriple(ByVal tan0 As Single) As Single
'    TanTriple = (1 - 3 * tan0 ^ 3)
'    If TanTriple <> 0 Then TanTriple = ((3 * tan0 - tan0 ^ 3) / (1 - 3 * tan0 ^ 3))
'End Function

'Public Function InvSin(Number As Single) As Single
'    InvSin = -Number * Number + 1
'    If InvSin > 0 Then
'        InvSin = Sqr(InvSin)
'        If InvSin <> 0 Then InvSin = Atn(Number / InvSin)
'    Else
'        InvSin = 0
'    End If
'End Function
'
'
'Public Function InvCos(Number As Single) As Single
'    InvCos = -Number * Number + 1
'    If InvCos > 0 Then
'        InvCos = Sqr(InvCos)
'        If InvCos <> 0 Then InvCos = Atn(-Number / InvCos) + 2 * Atn(1)
'    Else
'        InvCos = 0
'    End If
'End Function
'
'
'Public Function InvSec(Number As Single) As Single
'    InvSec = Number * Number + 1
'    If InvSec > 0 Then
'        InvSec = Sqr(InvSec)
'        If InvSec <> 0 Then InvSec = Atn(Number / InvSec) + Sgn((Number) - 1) * (2 * Atn(1))
'    Else
'        InvSec = 0
'    End If
'End Function
'
'
'Public Function InvCsc(Number As Single) As Single
'    InvCsc = Number * Number + 1
'    If InvCsc > 0 Then
'        InvCsc = Sqr(InvCsc)
'        If InvCsc <> 0 Then InvCsc = Atn(Number / InvCsc) + (Sgn(Number) - 1) * (2 * Atn(1))
'    Else
'        InvCsc = 0
'    End If
'End Function
'
'
'Public Function InvCot(Number As Single) As Single
'    InvCot = Atn(Number) + 2 * Atn(1)
'End Function
'
'
'Public Function Sec(Number As Single) As Single
'    On Error Resume Next
'    Sec = 1 / Cos(Number * PI / 180)
'End Function
'
'
'Public Function Csc(Number As Single) As Single
'    On Error Resume Next
'    Csc = 1 / Sin(Number * PI / 180)
'End Function
'
'
'Public Function Cot(Number As Single) As Single
'    On Error Resume Next
'    Cot = 1 / Tan(Number * PI / 180)
'End Function
'
'
'Public Function HSin(Number As Single) As Single
'    On Error Resume Next
'    HSin = (Exp(Number) - Exp(-Number)) / 2
'End Function
'
'
'Public Function HCos(Number As Single) As Single
'    On Error Resume Next
'    HCos = (Exp(Number) + Exp(-Number)) / 2
'End Function
'
'
'Public Function HTan(Number As Single) As Single
'    On Error Resume Next
'    HTan = (Exp(Number) - Exp(-Number)) / (Exp(Number) + Exp(-Number))
'End Function
'
'
'Public Function HSec(Number As Single) As Single
'    On Error Resume Next
'    HSec = 2 / (Exp(Number) + Exp(-Number))
'End Function
'
'
'Public Function HCsc(Number As Single) As Single
'    On Error Resume Next
'    HCsc = 2 / (Exp(Number) + Exp(-Number))
'End Function
'
'
'Public Function HCot(Number As Single) As Single
'    On Error Resume Next
'    HCot = (Exp(Number) + Exp(-Number)) / (Exp(Number) - Exp(-Number))
'End Function
'
'
'Public Function InvHSin(Number As Single)
'    On Error Resume Next
'    InvHSin = Log(Number + Sqr(Number * Number + 1))
'End Function
'
'
'Public Function InvHCos(Number As Single) As Single
'    On Error Resume Next
'    InvHCos = Log(Number + Sqr(Number * Number - 1))
'End Function
'
'
'Public Function InvHTan(Number As Single) As Single
'    On Error Resume Next
'    InvHTan = Log((1 + Number) / (1 - Number)) / 2
'End Function
'
'
'Public Function InvHSec(Number As Single) As Single
'    On Error Resume Next
'    InvHSec = Log((Sqr(-Number * Number + 1) + 1) / Number)
'End Function
'
'
'Public Function InvHCsc(Number As Single) As Single
'    On Error Resume Next
'    InvHCsc = Log((Sgn(Number) * Sqr(Number * Number + 1) + 1) / Number)
'End Function
'
'
'Public Function InvHCot(Number As Single) As Single
'    On Error Resume Next
'    InvHCot = Log((Number + 1) / (Number - 1)) / 2
'End Function




'Public Function VectorAxisAngles(ByRef p As Point, Optional ByVal Combined As Boolean = True) As Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Combined Then
'            Dim magnitude As Single
'            Dim heading As Single
'            Dim pitch As Single
'            Dim slope As Single
'            magnitude = Round(Sqr(p.x * p.x + p.y * p.y + p.z * p.z), 6)
'            If magnitude <> 0 Then
'                slope = Round((VectorSlope(MakePoint(0, 0, 0), p) / magnitude), 6)
'                heading = Round(ATan2(p.z, p.x), 6)
'                pitch = Round(ATan2(p.y, Round(Sqr(p.x * p.x + p.z * p.z), 6)), 6)
'                .x = Round((((heading / magnitude) - pitch) * slope), 6)
'                .z = Round((((PI / 2) + (-pitch + (heading / magnitude))) * (1 - slope)), 6)
'                .y = Round(((-heading + (pitch / magnitude)) * (1 - slope)), 6)
'                .y = Round((-(.y + ((.x * slope) / 2) - (.y * 2) - ((.z * slope) / 2))), 6)
'                .x = Round(((PI * 2) - (.x - ((PI / 2) * slope))), 6)
'                .z = Round(((PI * 2) - (.z - ((PI / 2) * slope))), 6)
'            End If
'        Else
''            .x = AngleOfCoord(MakePoint(p.z, p.Y, 0))
''            .Y = AngleOfCoord(MakePoint(p.x, p.z, 0))
''            .z = AngleOfCoord(MakePoint(p.Y, p.x, 0))
'
'            .x = ((360 - (AngleZOfCoordXY(MakePoint(p.z, p.y, p.x)) * DEGREE)) * RADIAN)
'            .y = ((360 - (AngleZOfCoordXY(MakePoint(p.x, p.z, p.y)) * DEGREE)) * RADIAN)
'            .z = -((360 - (AngleZOfCoordXY(MakePoint(p.y, p.x, p.z)) * DEGREE)) * RADIAN)
'
'        End If
'    End With
'End Function

'
'
'Public Function Percent(is_ As Single, of As Single) As Single
'
'    Percent = is_ / of * 100
'End Function


' Return the dot product AB · BC.
' Note that AB · BC = |AB| * |BC| * Cos(theta).
Private Function DotProduct( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal cx As Single, ByVal cy As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BCx As Single
Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BCx = cx - Bx
    BCy = cy - By

    ' Calculate the dot product.
    DotProduct = BAx * BCx + BAy * BCy
End Function

' Return the cross product AB x BC.
' The cross product is a vector perpendicular to AB
' and BC having length |AB| * |BC| * Sin(theta) and
' with direction given by the right-hand rule.
' For two vectors in the X-Y plane, the result is a
' vector with X and Y components 0 so the Z component
' gives the vector's length and direction.
'Public Function CrossProductLength( _
'    ByVal Ax As Single, ByVal Ay As Single, _
'    ByVal Bx As Single, ByVal By As Single, _
'    ByVal Cx As Single, ByVal Cy As Single _
'  ) As Single
'Dim BAx As Single
'Dim BAy As Single
'Dim BCx As Single
'Dim BCy As Single
'
'    ' Get the vectors' coordinates.
'    BAx = Ax - Bx
'    BAy = Ay - By
'    BCx = Cx - Bx
'    BCy = Cy - By
'
'    ' Calculate the Z coordinate of the cross product.
'    CrossProductLength = BAx * BCy - BAy * BCx
'End Function


Public Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, _
    ByVal Bx As Single, ByVal By As Single, ByVal Bz As Single, _
    ByVal cx As Single, ByVal cy As Single, ByVal cz As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BAz As Single
Dim BCx As Single
Dim BCy As Single
Dim BCz As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BAz = Az - Bz
    BCx = cx - Bx
    BCy = cy - By
    BCz = cz - Bz
    
    ' Calculate the Z coordinate of the cross product.
    CrossProductLength = BAx * BCy - BAy * BCz - BAz * BCx
End Function

'Function Atan(X As Single, Y As Single)
'Const PI = 3.14159
'
'Dim angle As Single
'
'    If X = 0 Then
'        angle = 0
'    Else
'        angle = Atn(Y / X)
'        If X < 0 Then angle = PI + angle
'    End If
'
'    Atan = angle
'End Function



''' Return the angle with tangent opp/hyp. The returned
''' value is between PI and -PI.



'Public Function ATan2(ByVal y As Double, ByVal x As Double) As Double
'     Const PI_14 As Double = 0.785398163397448
' Const PI_34 As Double = 2.35619449019234
'
'    'Cheap non-branching workaround for cases where y approaches 0.0
'    Dim absY As Double
'    absY = Abs(y) + 0.0000000001
'
'    If (x >= 0#) Then
'        ATan2 = PI_14 - PI_14 * (x - absY) / (x + absY)
'    Else
'        ATan2 = PI_34 - PI_14 * (x + absY) / (absY - x)
'    End If
'
'    If (y < 0#) Then ATan2 = -ATan2
'
'End Function


'Public Function ATan2(y As Single, x As Single) As Single
'    ' Radians.
'    '
'    On Error GoTo Atan2Error
'    ATan2 = Atn(y / x)
'    If (x < 0) Then If (y < 0) Then ATan2 = ATan2 - PI Else ATan2 = ATan2 + PI
'    Exit Function
'
'Atan2Error:
'    If Abs(y) > Abs(x) Then     ' Must be an overflow.
'        If y > 0 Then ATan2 = PI / 2 Else ATan2 = -PI / 2
'    Else
'        ATan2 = 0           ' Must be an underflow.
'    End If
'    Resume Next
'End Function

'' Return the angle ABC.
'' Return a value between PI and -PI.
'' Note that the value is the opposite of what you might
'' expect because Y coordinates increase downward.
Public Function GetAngle(ByRef p1 As Point, ByRef p2 As Point) As Single
'ByVal Ax As Single, ByVal Ay As _
    'Single, ByVal Bx As Single, ByVal By As Single, ByVal _
    'Cx As Single, ByVal Cy As Single) As Single
Dim dot_product As Single
Dim cross_product As Single

    ' Get the dot product and cross product.
    dot_product = VectorDotProduct(p1, p2) 'dotproduct(p1.x, p1.y, 0, 0, p3.x, p3.y)
    cross_product = DistanceEx(MakePoint(0, 0, 0), VectorCrossProduct(p1, p2)) 'CrossProductLength(p1.x, p1.y, p1.z, 0, 0, 0, p2.x, p2.y, p2.z) 'CrossProductLength(p1.x, p1.y, 0, 0, p3.x, p3.y)

    ' Calculate the angle.
    GetAngle = ATan2(CDbl(cross_product), CDbl(dot_product)) * DEGREE
End Function


'Public Function Sine(ByVal V As Single, Optional ByVal H As Variant) As Single
'    'inputs: v is vertical axis, or angle degree
'    '        h is horizontal axis, or left blank
'    'returns: return is the angle in degrees, or
'    '         a ratio, if only angle v is passed
'    If IsMissing(H) Then
'        If ((V <= 1) And (V > -360)) Then V = (360 - V)
'        If (Abs(V) > 0) Then
'            If (V > 45) Then V = (V - ((V \ 45) * 45))
'            V = ((V / (46 / 100)) / 100)
'            If Abs(V) > 0 Then
'                V = (44 / (1 / V))
'
'                Sine = (V * (180 / PI))
'            End If
'        Else
'            Sine = 0
'        End If
'    Else
'        If ((H = 0 Or V = 0) Or (H = V)) Then
'            Sine = -CBool(Not ((H = V) And Abs(H) > 0))
'        Else
'            Dim d As Single
'            d = Sqr((Abs(H) ^ 2) + (Abs(V) ^ 2))
'            If ((H > 0) And (V > 0)) Then
'                Sine = (Abs(V) / d)
'            ElseIf ((H < 0) And (V > 0)) Then
'                Sine = (Abs(H) / d)
'            ElseIf ((H < 0) And (V < 0)) Then
'                Sine = (Abs(H) / d)
'            ElseIf ((H > 0) And (V < 0)) Then
'                Sine = (Abs(V) / d)
'            End If
'        End If
'        Sine = (Sine * (PI / 4) * (100 / 90))
'    End If
'End Function
'
'Public Function Cosine(ByVal V As Single, Optional ByVal H As Variant) As Single
'    'inputs: v is vertical axis, or angle degree
'    '        h is horizontal axis, or left blank
'    'returns: return is the angle in degrees, or
'    '         a ratio, if only angle v is passed
'    If IsMissing(H) Then
'        If ((V <= 1) And (V > -360)) Then V = (360 - V)
'        If (Abs(V) > 0) Then
'            If (V > 45) Then V = (V - ((V \ 45) * 45))
'            V = ((V / (92 / 100)) / 100)
'            V = (44 / (1 / V))
'            Cosine = (V * (180 / PI))
'        Else
'            Cosine = 1
'        End If
'    Else
'        If ((H = 0 Or V = 0) Or (H = V)) Then
'            Cosine = -CBool(((H = V) And Abs(H) > 0))
'        Else
'            Dim d As Single
'            d = Sqr((Abs(H) ^ 2) + (Abs(V) ^ 2))
'            If ((H > 0) And (V > 0)) Then
'                Cosine = (Abs(H) / d)
'            ElseIf ((H < 0) And (V > 0)) Then
'                Cosine = (Abs(V) / d)
'            ElseIf ((H < 0) And (V < 0)) Then
'                Cosine = (Abs(V) / d)
'            ElseIf ((H > 0) And (V < 0)) Then
'                Cosine = (Abs(H) / d)
'            End If
'        End If
'        Cosine = (Cosine * (PI / 4) * (100 / 90))
'    End If
'End Function
'
'Public Function Tangent(ByVal V As Single, Optional ByVal H As Variant) As Single
'    'inputs: v is vertical axis, or angle degree
'    '        h is horizontal axis, or left blank
'    'returns: return is the angle in degrees, or
'    '         a ratio, if only angle v is passed
'    If IsMissing(H) Then
'        If ((V <= 1) And (V > -360)) Then V = (360 - V)
'        If (Abs(V) > 0) Then
'            If (V > 90) Then V = (V - ((V \ 90) * 90))
'            V = ((V / (46 / 100)) / 50)
'            V = ((44 / (1 / V)) / PI)
'            Tangent = (V * (180 / PI))
'        Else
'            Tangent = 1
'        End If
'    Else
'        If (H = 0 Or V = 0) Then
'            Tangent = 1 + (-CBool(((H = V) And Abs(H) > 0)) / 2)
'        Else
'            If ((H < 0 And V > 0) Or (H < 0 And V < 0)) Then
'                If (Abs(H) > Abs(V)) Then
'                    Tangent = (Abs(V) / Abs(H))
'                Else
'                    Tangent = (Abs(H) / Abs(V))
'                End If
'            ElseIf ((H > 0 And V > 0) Or (H > 0 And V < 0)) Then
'                If Abs(H) > Abs(V) Then
'                    Tangent = (Abs(H) / Abs(V))
'                Else
'                    Tangent = (Abs(V) / Abs(H))
'                End If
'            End If
'            Tangent = (Tangent * (100 / 45))
'        End If
'    End If
'End Function

    Public Function GetAngle2(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        Dim XDiff As Double
        Dim YDiff As Double
        Dim TempAngle As Double

        YDiff = Abs(y2 - y1)

        If x1 = x2 And y1 = y2 Then Exit Function

        If YDiff = 0 And x1 < x2 Then
            GetAngle2 = 0
            Exit Function
        ElseIf YDiff = 0 And x1 > x2 Then
            GetAngle2 = 3.14159265358979
            Exit Function
        End If

        XDiff = Abs(x2 - x1)

        TempAngle = Atn(XDiff / YDiff)

        If y2 > y1 Then TempAngle = 3.14159265358979 - TempAngle
        If x2 < x1 Then TempAngle = -TempAngle
        TempAngle = 1.5707963267949 - TempAngle
        If TempAngle < 0 Then TempAngle = 6.28318530717959 + TempAngle

        GetAngle2 = TempAngle
    End Function


Public Function GetAngle3(ByRef p1 As Point, ByRef p2 As Point) As Single
If p1.X = p2.X Then
    If p1.Y < p2.Y Then
        GetAngle3 = 90
    Else
        GetAngle3 = 270
    End If
    Exit Function
ElseIf p1.Y = p2.Y Then
    If p1.X < p2.X Then
        GetAngle3 = 0
    Else
        GetAngle3 = 180
    End If
    Exit Function
Else
    GetAngle3 = Atn(VectorSlope(p1, p2))
    GetAngle3 = GetAngle3 * 180 / PI
    If GetAngle3 < 0 Then GetAngle3 = GetAngle3 + 360
    '----------Test for direction--------
    If p1.X > p2.X And GetAngle3 <> 180 Then GetAngle3 = GetAngle3 + 180
    If p1.Y > p2.Y And GetAngle3 = 90 Then GetAngle3 = GetAngle3 + 180
    If GetAngle3 > 360 Then GetAngle3 = GetAngle3 - 360
End If
End Function

Public Function Sine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180
    p_dblVal = Val(p_dblVal * dblRadian)
    Sine = Sin(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Sine = 0
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function Cosine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180
    p_dblVal = Val(p_dblVal * dblRadian)
    Cosine = Cos(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Cosine = 0
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function Tangent(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Degree Input Radian Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblRadian As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert degrees to radians,
    'multiply degrees by Pi / 180.
    dblRadian = dblPi / 180

    p_dblVal = Val(p_dblVal * dblRadian)
    Tangent = Tan(p_dblVal)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Tangent = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcSine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblSqr As Single
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
    ' xx Prevent division by Zero error

    If dblSqr = 0 Then
        dblSqr = 1E-30
    End If

    ArcSine = Atn(p_dblVal / dblSqr) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcSine = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcCosine(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblSqr As Single
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    dblSqr = Sqr(-p_dblVal * p_dblVal + 1)
    ' xx Prevent division by Zero error

    If dblSqr = 0 Then
        dblSqr = 1E-30
    End If

    ArcCosine = (Atn(-p_dblVal / dblSqr) + 2 * Atn(1)) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcCosine = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function


Public Function ArcTangent(p_dblVal As Single) As Single

    ' Comments :
    ' Parameters: p_dblVal -
    ' Returns: Double -
    ' Modified :
    '
    ' -------------------------
    'Radian Input Degree Output
    On Error GoTo PROC_ERR
    Dim dblPi As Single
    Dim dblDegree As Single
    ' xx Calculate the value of Pi.
    dblPi = 4 * Atn(1)
    ' xx To convert radians to degrees,
    ' multiply radians by 180/pi.
    dblDegree = 180 / dblPi
    p_dblVal = Val(p_dblVal)
    ArcTangent = Atn(p_dblVal) * dblDegree
PROC_EXIT:
    Exit Function
PROC_ERR:
    ArcTangent = 0
    'MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT
End Function

'Public Sub StrictRotation(ByRef p As Point)
'    Do While p.X >= PI * 2 Or p.X < -PI * 2
'        If p.X >= PI * 2 Then
'            p.X = p.X - PI * 2
'        Else
'            p.X = p.X + PI * 2
'        End If
'    Loop
'    Do While p.Y >= PI * 2 Or p.Y < -PI * 2
'        If p.Y >= PI * 2 Then
'            p.Y = p.Y - PI * 2
'        Else
'            p.Y = p.Y + PI * 2
'        End If
'    Loop
'    Do While p.Z >= PI * 2 Or p.Z < -PI * 2
'        If p.Z >= PI * 2 Then
'            p.Z = p.Z - PI * 2
'        Else
'            p.Z = p.Z + PI * 2
'        End If
'    Loop
'End Sub

'Public Function VectorRotate(ByRef Point As Point, ByRef Angles As Point) As Point
'
'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Dim matMesh As D3DMATRIX
'    Dim matYaw As D3DMATRIX
'    Dim matPitch As D3DMATRIX
'    Dim matRoll As D3DMATRIX
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixIdentity matYaw
'    D3DXMatrixIdentity matPitch
'    D3DXMatrixIdentity matRoll
'
'    D3DXMatrixRotationY matYaw, Angles.x
'    D3DXMatrixMultiply matMesh, matYaw, matMesh
'
'    D3DXMatrixRotationX matPitch, Angles.y
'    D3DXMatrixMultiply matMesh, matPitch, matMesh
'
'    D3DXMatrixRotationZ matRoll, Angles.z
'    D3DXMatrixMultiply matMesh, matRoll, matMesh
'
'    vin = ToVector(Point)
'    D3DXVec3TransformCoord vout, vin, matMesh
'    Set VectorRotate = ToPoint(vout)
'
'End Function

'Private Function CombineRange(ByRef o1 As Range, ByRef o2 As Range) As Range
'    Set CombineRange = New Range
'    With CombineRange
'        .Ranges.X = o1.Ranges.X + o2.Ranges.X
'        .Ranges.Y = o1.Ranges.Y + o2.Ranges.Y
'        .Ranges.Z = o1.Ranges.Z + o2.Ranges.Z
'        If o1.Ranges.W = -1 Or o2.Ranges.W = -1 Then
'            .Ranges.W = -1
'        ElseIf o1.Ranges.W > o2.Ranges.W Then
'            .Ranges.W = o1.Ranges.W
'        Else
'            .Ranges.W = o2.Ranges.W
'        End If
'    End With
'End Function
'
'Private Function CombineRotate(ByRef o1 As Point, ByRef o2 As Point) As Point
'    Set CombineRotate = VectorAddition(o1, o2)
'End Function
'
'Private Function CombineOffset(ByRef o1 As Orbit, ByRef o2 As Orbit) As Orbit
'    Set CombineOffset = New Orbit
'    With CombineOffset
'        .Offset = VectorDeduction(VectorDeduction(VectorAddition(o1.Origin, o1.Offset), VectorAddition(o2.Origin, o2.Offset)), .Origin)
'    End With
'End Function
'Private Function CombineOrigin(ByRef o1 As Point, ByRef o2 As Point) As Point
'    Set CombineOrigin = VectorAddition(o1, o2)
'End Function
'
'Private Function CombineScaled(ByRef o1 As Point, ByRef o2 As Point) As Point
'    Set CombineScaled = VectorAddition(o1, o2)
'End Function
'
'Private Function CombineOrbit(ByRef o1 As Orbit, ByRef o2 As Orbit) As Orbit
'    Set CombineOrbit = New Orbit
'    With CombineOrbit
'        .Origin = VectorAddition(o1.Origin, o2.Origin)
'        .Offset = VectorDeduction(VectorDeduction(VectorAddition(o1.Origin, o1.Offset), VectorAddition(o2.Origin, o2.Offset)), .Origin)
'        .Rotate = VectorAddition(o1.Rotate, o2.Rotate)
'        .Scaled = VectorAddition(o1.Scaled, o2.Scaled)
'        .Ranges.X = o1.Ranges.X + o2.Ranges.X
'        .Ranges.Y = o1.Ranges.Y + o2.Ranges.Y
'        .Ranges.Z = o1.Ranges.Z + o2.Ranges.Z
'        If o1.Ranges.W = -1 Or o2.Ranges.W = -1 Then
'            .Ranges.W = -1
'        ElseIf o1.Ranges.W > o2.Ranges.W Then
'            .Ranges.W = o1.Ranges.W
'        Else
'            .Ranges.W = o2.Ranges.W
'        End If
'    End With
'End Function


'Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing)
''    If Not GlobalGyro Is Nothing Then
''        LocPos VectorRotate(Origin, GlobalGyro), False, ApplyTo 'location is changing the origin to absolute
''    Else
'        LocPos Origin, False, ApplyTo 'location is changing the origin to absolute
''    End If
'End Sub
'Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing)
''    If Not GlobalGyro Is Nothing Then
''        LocPos VectorRotate(Origin, GlobalGyro), True, ApplyTo  'position is changing the origin relative
''    Else
'        LocPos Origin, True, ApplyTo 'position is changing the origin relative
''    End If
'End Sub
'Private Sub LocPos(ByRef Origin As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing)
'    If Origin.X <> 0 Or Origin.Y <> 0 Or Origin.z <> 0 Then
'        Dim o As Orbit
'        Dim m As Molecule
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If p.Ranges.W = -1 Then
'                        LocPos Origin, Relative, p
'                    ElseIf p.Ranges.W > 0 Then
'                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
'                            LocPos Origin, Relative, p
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                If Relative Then
'                    Set ApplyTo.Relative.Origin = VectorAddition(ApplyTo.Relative.Origin, VectorRotateAxis(Origin, ApplyTo.Rotate))
'                Else
'                    Set ApplyTo.Absolute.Origin = VectorDeduction(Origin, ApplyTo.Origin)
'                End If
''                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.W = -1 Then
''                            LocPos Origin, Relative, m
''                        ElseIf ApplyTo.Ranges.W > 0 Then
''                            If ApplyTo.Ranges.W - Distance(m.Origin.x, m.Origin.y, m.Origin.Z, ApplyTo.Origin.x, ApplyTo.Origin.y, ApplyTo.Origin.Z) > 0 Then
''                                LocPos Origin, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        LocPos Origin, Relative, m
''                    Next
''                End If
'        End Select
'    End If
'End Sub
'Public Sub Rotation(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing)
'    RotOri Degrees, False, ApplyTo 'location is changing the origin to absolute
'End Sub
'Public Sub Orientate(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing)
'    RotOri Degrees, True, ApplyTo 'position is changing the origin relative
'End Sub
'Private Sub RotOri(ByRef Degrees As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing)
'    If Degrees.X <> 0 Or Degrees.Y <> 0 Or Degrees.z <> 0 Then
'        Dim m As Molecule
'        Dim o As Point
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If p.Ranges.W = -1 Then
'                        RotOri Degrees, Relative, p
'                    ElseIf p.Ranges.W > 0 Then
'                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
'                            RotOri Degrees, Relative, p
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                If Relative Then
'                    Set ApplyTo.Relative.Rotate = VectorAddition(ApplyTo.Relative.Rotate, Degrees)
'                Else
'                    Set ApplyTo.Absolute.Rotate = VectorDeduction(Degrees, ApplyTo.Rotate)
'                End If
''                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.W = -1 Then
''                            RotOri Degrees, Relative, m
''                        ElseIf ApplyTo.Ranges.W > 0 Then
''                            If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
''                                RotOri Degrees, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        RotOri Degrees, Relative, m
''                    Next
''                End If
'        End Select
'    End If
'End Sub
'
'Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing)
'    ScaExp Ratios, False, ApplyTo 'location is changing the origin to absolute
'End Sub
'Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing)
'    ScaExp Ratios, True, ApplyTo 'position is changing the origin relative
'End Sub
'Private Sub ScaExp(ByRef Scalar As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing)
'    If Abs(Scalar.X) <> 1 Or Abs(Scalar.Y) <> 1 Or Abs(Scalar.z) <> 1 Then
'        Dim m As Molecule
'        Dim o As Orbit
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If p.Ranges.W = -1 Then
'                        ScaExp Scalar, Relative, p
'                    ElseIf p.Ranges.W > 0 Then
'                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
'                            ScaExp Scalar, Relative, p
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                'change all molecules with in the specified planets range
'                If Relative Then
'                    Set ApplyTo.Relative.Scaled = VectorAddition(ApplyTo.Relative.Scaled, Scalar)
'                Else
'                    Set ApplyTo.Absolute.Scaled = VectorDeduction(Scalar, ApplyTo.Scaled)
'                End If
''                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.W = -1 Then
''                            ScaExp Scalar, Relative, m
''                        ElseIf ApplyTo.Ranges.W > 0 Then
''                            If ApplyTo.Ranges.W - Distance(m.Origin.x, m.Origin.y, m.Origin.Z, ApplyTo.Origin.x, ApplyTo.Origin.y, ApplyTo.Origin.Z) > 0 Then
''                                ScaExp Scalar, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        ScaExp Scalar, Relative, m
''                    Next
''                End If
'        End Select
'    End If
'End Sub
''Public Sub Displace(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing)
''    DisBal Offset, False, ApplyTo 'location is changing the origin to absolute
''End Sub
''Public Sub Balanced(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing)
''    DisBal Offset, True, ApplyTo 'position is changing the origin relative
''End Sub
''Private Sub DisBal(ByRef Offset As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing)
''    If Offset.X <> 0 Or Offset.Y <> 0 Or Offset.Z <> 0 Then
''        Dim dist As Single
''        Dim m As Molecule
''        Dim o As Orbit
''        Select Case TypeName(ApplyTo)
''            Case "Nothing"
''                'go retrieve all planets whos range and origin has (0,0,0) with in it
''                'and call change all molucules with in each of those planets as well
''                Dim p As Planet
''                For Each p In Planets
''                    If p.Ranges.W = -1 Then
''                        DisBal Offset, Relative, p
''                    ElseIf p.Ranges.W > 0 Then
''                        dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0)
''                        If p.Ranges.W - dist > 0 Then
''                            DisBal Offset, Relative, p
''                        End If
''                    End If
''                Next
''            Case "Planet"
''                'change all molecules with in the specified planets range
''                If Relative Then
''                    Set ApplyTo.Relative.Offset = VectorAddition(ApplyTo.Relative.Offset, Offset)
''                    Set ApplyTo.Absolute.Offset = VectorNegative(ApplyTo.Absolute.Offset)
''                Else
''                    Set ApplyTo.Absolute.Offset = VectorDeduction(Offset, ApplyTo.Offset)
''                    Set ApplyTo.Relative.Offset = VectorNegative(ApplyTo.Relative.Offset)
''                End If
''                For Each m In Molecules
''                    If ApplyTo.Ranges.W = -1 Then
''                        DisBal Offset, Relative, m
''                    ElseIf ApplyTo.Ranges.W > 0 Then
''                        dist = Distance(m.Origin.X, m.Origin.Y, m.Origin.Z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.Z)
''                        If ApplyTo.Ranges.W - dist > 0 Then
''                            DisBal Offset, Relative, m
''                        End If
''                    End If
''                Next
''            Case "Molecule"
''                If Relative Then
''                    Set ApplyTo.Relative.Offset = VectorAddition(ApplyTo.Relative.Offset, Offset)
''                    Set ApplyTo.Absolute.Offset = VectorNegative(ApplyTo.Absolute.Offset)
''                Else
''                    Set ApplyTo.Absolute.Offset = VectorDeduction(Offset, ApplyTo.Offset)
''                    Set ApplyTo.Relative.Offset = VectorNegative(ApplyTo.Relative.Offset)
''                End If
''                For Each m In ApplyTo.Molecules
''                    DisBal Offset, Relative, m
''                Next
''        End Select
''    End If
''End Sub
'
'Public Sub Begin(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
'
'    Dim m As Molecule
'    Dim p As Planet
'
''
''    For Each p In Planets
''
''        If Not p.Absolute.Rotate.Equals(p.Rotate) Then
''            ApplyRotate p, p.Absolute.Rotate
''        End If
''        If Not p.Absolute.Scaled.Equals(p.Scaled) Then
''            ApplyScaled p, p.Absolute.Scaled
''        End If
''        If Not p.Absolute.Origin.Equals(p.Origin) Then
''            ApplyOrigin p, p.Absolute.Origin
''        End If
''        If Not p.Absolute.Offset.Equals(p.Offset) Then
''            ApplyOffset p, p.Absolute.Offset
''        End If
''
''        If Abs(p.Relative.Scaled.X) <> 1 Or Abs(p.Relative.Scaled.Y) <> 1 Or Abs(p.Relative.Scaled.z) <> 1 Then
''            ApplyScaled p, p.Relative.Scaled
''        End If
''        If p.Relative.Rotate.X <> 0 Or p.Relative.Rotate.Y <> 0 Or p.Relative.Rotate.z <> 0 Then
''            ApplyRotate p, p.Relative.Rotate
''        End If
''        If p.Relative.Origin.X <> 0 Or p.Relative.Origin.Y <> 0 Or p.Relative.Origin.z <> 0 Then
''            ApplyOrigin p, p.Relative.Origin
''        End If
''
''        Set p.Relative = Nothing
''
''    Next
'
'
'    For Each m In Molecules
'
'
'        If Not m.Absolute.Rotate.Equals(m.Rotate) Then
'            ApplyRotate m, m.Absolute.Rotate
'        End If
'        If Not m.Absolute.Scaled.Equals(m.Scaled) Then
'            ApplyScaled m, m.Absolute.Scaled
'        End If
'        If Not m.Absolute.Origin.Equals(m.Origin) Then
'            ApplyOrigin m, m.Absolute.Origin
'        End If
'        If Not m.Absolute.Offset.Equals(m.Offset) Then
'            ApplyOffset m, m.Absolute.Offset
'        End If
'
'        If Abs(m.Relative.Scaled.X) <> 1 Or Abs(m.Relative.Scaled.Y) <> 1 Or Abs(m.Relative.Scaled.z) <> 1 Then
'            ApplyScaled m, m.Relative.Scaled
'        End If
'        If m.Relative.Rotate.X <> 0 Or m.Relative.Rotate.Y <> 0 Or m.Relative.Rotate.z <> 0 Then
'            ApplyRotate m, m.Relative.Rotate
'        End If
'        If m.Relative.Origin.X <> 0 Or m.Relative.Origin.Y <> 0 Or m.Relative.Origin.z <> 0 Then
'            ApplyOrigin m, m.Relative.Origin
'        End If
'
'        Set m.Relative = Nothing
'
'    Next
'
'
'End Sub

'Public Sub Finish(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
'
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
''
''                    DDevice.SetRenderState D3DRS_ZENABLE, 1
''                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'   ' Iterate Planets
'
'
'
''                    DDevice.SetRenderState D3DRS_ZENABLE, 1
''                    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'
'    Iterate Molecules
'
'End Sub
'
'Private Sub Iterate(ByRef col As Object)
'
'    Dim m As Molecule
'
'    For Each m In col
'
'        Render m, m.Origin
'
'    Next
'
'End Sub
'
'
'Private Sub Render(ByRef m As Molecule, ByRef Base As Point)
'
'
'    Dim V As Matter
'
'    Dim mat As D3DMATRIX
'
'    For Each V In m.Volume
'
''        If TypeName(m) = "Planet" Then
''        VertexDirectX((V.TriangleIndex * 3) + 0).X = V.Point1.X '+ m.Origin.X
''        VertexDirectX((V.TriangleIndex * 3) + 0).Y = V.Point1.Y '+ m.Origin.Y
''        VertexDirectX((V.TriangleIndex * 3) + 0).z = V.Point1.z '+ m.Origin.z
''
''        VertexDirectX((V.TriangleIndex * 3) + 1).X = V.Point2.X '+ m.Origin.X
''        VertexDirectX((V.TriangleIndex * 3) + 1).Y = V.Point2.Y '+ m.Origin.Y
''        VertexDirectX((V.TriangleIndex * 3) + 1).z = V.Point2.z '+ m.Origin.z
''
''        VertexDirectX((V.TriangleIndex * 3) + 2).X = V.Point3.X '+ m.Origin.X
''        VertexDirectX((V.TriangleIndex * 3) + 2).Y = V.Point3.Y '+ m.Origin.Y
''        VertexDirectX((V.TriangleIndex * 3) + 2).z = V.Point3.z '+ m.Origin.z
''
''        Else
'
'        VertexDirectX((V.TriangleIndex * 3) + 0).X = V.Point1.X + m.Origin.X
'        VertexDirectX((V.TriangleIndex * 3) + 0).Y = V.Point1.Y + m.Origin.Y
'        VertexDirectX((V.TriangleIndex * 3) + 0).z = V.Point1.z + m.Origin.z
'
'        VertexDirectX((V.TriangleIndex * 3) + 1).X = V.Point2.X + m.Origin.X
'        VertexDirectX((V.TriangleIndex * 3) + 1).Y = V.Point2.Y + m.Origin.Y
'        VertexDirectX((V.TriangleIndex * 3) + 1).z = V.Point2.z + m.Origin.z
'
'        VertexDirectX((V.TriangleIndex * 3) + 2).X = V.Point3.X + m.Origin.X
'        VertexDirectX((V.TriangleIndex * 3) + 2).Y = V.Point3.Y + m.Origin.Y
'        VertexDirectX((V.TriangleIndex * 3) + 2).z = V.Point3.z + m.Origin.z
'''        End If
'
'        VertexDirectX(V.TriangleIndex * 3 + 0).NX = V.Normal.X
'        VertexDirectX(V.TriangleIndex * 3 + 0).NY = V.Normal.Y
'        VertexDirectX(V.TriangleIndex * 3 + 0).Nz = V.Normal.z
'
'        VertexDirectX(V.TriangleIndex * 3 + 1).NX = V.Normal.X
'        VertexDirectX(V.TriangleIndex * 3 + 1).NY = V.Normal.Y
'        VertexDirectX(V.TriangleIndex * 3 + 1).Nz = V.Normal.z
'
'        VertexDirectX(V.TriangleIndex * 3 + 2).NX = V.Normal.X
'        VertexDirectX(V.TriangleIndex * 3 + 2).NY = V.Normal.Y
'        VertexDirectX(V.TriangleIndex * 3 + 2).Nz = V.Normal.z
'
'        VertexXAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).X
'        VertexXAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).X
'        VertexXAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).X
'
'        VertexYAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).Y
'        VertexYAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).Y
'        VertexYAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).Y
'
'        VertexZAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).z
'        VertexZAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).z
'        VertexZAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).z
'
''         If TypeName(m) = "Planet" Then
''
''            ScreenDirectX((V.TriangleIndex * 3) + 0).X = VertexDirectX((V.TriangleIndex * 3) + 0).X
''            ScreenDirectX((V.TriangleIndex * 3) + 0).Y = VertexDirectX((V.TriangleIndex * 3) + 0).z
''            ScreenDirectX((V.TriangleIndex * 3) + 0).z = VertexDirectX((V.TriangleIndex * 3) + 0).Y
''
''            ScreenDirectX((V.TriangleIndex * 3) + 1).X = VertexDirectX((V.TriangleIndex * 3) + 1).X
''            ScreenDirectX((V.TriangleIndex * 3) + 1).Y = VertexDirectX((V.TriangleIndex * 3) + 1).z
''            ScreenDirectX((V.TriangleIndex * 3) + 1).z = VertexDirectX((V.TriangleIndex * 3) + 1).Y
''
''            ScreenDirectX((V.TriangleIndex * 3) + 2).X = VertexDirectX((V.TriangleIndex * 3) + 2).X
''            ScreenDirectX((V.TriangleIndex * 3) + 2).Y = VertexDirectX((V.TriangleIndex * 3) + 2).z
''            ScreenDirectX((V.TriangleIndex * 3) + 2).z = VertexDirectX((V.TriangleIndex * 3) + 2).Y
''
''
''         End If
'
'
'
'         If m.Visible And (Not (TypeName(m) = "Planet")) Then
'
''            If Not Camera.Planet Is Nothing Then
''                D3DXMatrixIdentity mat
''                D3DXMatrixRotationYawPitchRoll mat, m.Origin.X, m.Origin.Y, m.Origin.z
''
''                DDevice.SetTransform D3DTS_WORLD, mat
''
''            End If
'
'
'
'             If Not (V.Translucent Or V.Transparent) Then
'                 DDevice.SetMaterial GenericMaterial
'                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
'                 DDevice.SetTexture 1, Nothing
'             Else
'                 DDevice.SetMaterial LucentMaterial
'                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
'                 DDevice.SetMaterial GenericMaterial
'                 If V.TextureIndex > 0 Then DDevice.SetTexture 1, Files(V.TextureIndex).Data
'             End If
'
'             DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VertexDirectX((V.TriangleIndex * 3)), Len(VertexDirectX(0))
'         End If
'    Next
'
'    Dim m2 As Molecule
'
'    For Each m2 In m.Molecules
'        Render m2, VectorAddition(m.Origin, Base)
'
'    Next
'
'End Sub


'Private Sub ApplyOrigin(ByRef Origin As Point, ByRef ApplyTo As Molecule, Optional ByRef Parent As Molecule)
'    'modifies the actual 3D object's pointsm only once a frame for speed consideration
'
'    If Origin Is Nothing Then Exit Sub
'
'    Static vin As D3DVECTOR
'    Static vout As D3DVECTOR
'    Static matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    Set ApplyTo.Origin = VectorAddition(Origin, ApplyTo.Origin)
'    Set ApplyTo.Absolute.Origin = Origin
'
'    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyOrigin Origin, m, ApplyTo
'    Next
'    stacked = stacked - 1
'End Sub

'Private Static Sub ApplyOffset(ByRef Offset As Point, ByRef ApplyTo As Molecule, Optional ByRef Parent As Molecule)
'    'modifies the actual 3D object's pointsm only once a frame for speed consideration
'    If Offset Is Nothing Then Exit Sub
'
'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Dim matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixTranslation matMesh, -ApplyTo.Offset.X + Offset.X, -ApplyTo.Offset.Y + Offset.Y, -ApplyTo.Offset.z + Offset.z
'
'    Set ApplyTo.Offset = Offset
'    Set ApplyTo.Absolute.Offset = Offset
'
'    Dim V As Matter
'    For Each V In ApplyTo.Volume
'
'        vin.X = V.Point1.X
'        vin.Y = V.Point1.Y
'        vin.z = V.Point1.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point1.X = vout.X
'        V.Point1.Y = vout.Y
'        V.Point1.z = vout.z
'
'        vin.X = V.Point2.X
'        vin.Y = V.Point2.Y
'        vin.z = V.Point2.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point2.X = vout.X
'        V.Point2.Y = vout.Y
'        V.Point2.z = vout.z
'
'        vin.X = V.Point3.X
'        vin.Y = V.Point3.Y
'        vin.z = V.Point3.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point3.X = vout.X
'        V.Point3.Y = vout.Y
'        V.Point3.z = vout.z
'
'        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'    Next
'
'    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyOffset Offset, m, ApplyTo
'    Next
'    stacked = stacked - 1
'
'End Sub
'
'Private Sub ApplyRotate(ByRef Degrees As Point, ByRef ApplyTo As Molecule, Optional ByRef Parent As Molecule)
'    'modifies the actual 3D object's pointsm only once a frame for speed consideration
'    If Degrees Is Nothing Then Exit Sub
'
'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'
'    Dim matYaw As D3DMATRIX
'    Dim matPitch As D3DMATRIX
'    Dim matRoll As D3DMATRIX
'    Dim matPos As D3DMATRIX
'
'    Static matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    If stacked = 0 Then
'
'        D3DXMatrixIdentity matMesh
'        D3DXMatrixIdentity matYaw
'        D3DXMatrixIdentity matPitch
'        D3DXMatrixIdentity matRoll
'
'
'        D3DXMatrixRotationX matYaw, Degrees.X '- ApplyTo.Rotate.X
'        D3DXMatrixMultiply matMesh, matYaw, matMesh
'
'        D3DXMatrixRotationY matPitch, Degrees.Y '- ApplyTo.Rotate.Y
'        D3DXMatrixMultiply matMesh, matPitch, matMesh
'
'        D3DXMatrixRotationZ matRoll, Degrees.z '- ApplyTo.Rotate.z
'        D3DXMatrixMultiply matMesh, matRoll, matMesh
'
'
'        Set ApplyTo.Rotate = Degrees
'        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
'
'    Else
'
'
'        Set ApplyTo.Rotate = AngleAxisAddition(Degrees, ApplyTo.Rotate)
'        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
'
'        Set ApplyTo.Origin = VectorRotateAxis(ApplyTo.Origin, Degrees)
'        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
'
'
'
'
'
'    End If
'
'    If Not TypeName(ApplyTo) = "Planet" Then
'
'        Dim V As Matter
'        For Each V In ApplyTo.Volume
'
'            vin.X = V.Point1.X
'            vin.Y = V.Point1.Y
'            vin.z = V.Point1.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point1.X = vout.X
'            V.Point1.Y = vout.Y
'            V.Point1.z = vout.z
'
'            vin.X = V.Point2.X
'            vin.Y = V.Point2.Y
'            vin.z = V.Point2.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point2.X = vout.X
'            V.Point2.Y = vout.Y
'            V.Point2.z = vout.z
'
'            vin.X = V.Point3.X
'            vin.Y = V.Point3.Y
'            vin.z = V.Point3.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            V.Point3.X = vout.X
'            V.Point3.Y = vout.Y
'            V.Point3.z = vout.z
'
'        Next
'    End If
'
''    If TypeName(ApplyTo) = "Planet" Then
''        Dim V As Matter
''        For Each V In ApplyTo.Volume
''
''            vin.x = V.Point1.x
''            vin.y = V.Point1.y
''            vin.z = V.Point1.z
''            D3DXVec3TransformCoord vout, vin, matMesh
''            V.Point1.x = vout.x
''            V.Point1.y = vout.y
''            V.Point1.z = vout.z
''
''            vin.x = V.Point2.x
''            vin.y = V.Point2.y
''            vin.z = V.Point2.z
''            D3DXVec3TransformCoord vout, vin, matMesh
''            V.Point2.x = vout.x
''            V.Point2.y = vout.y
''            V.Point2.z = vout.z
''
''            vin.x = V.Point3.x
''            vin.y = V.Point3.y
''            vin.z = V.Point3.z
''            D3DXVec3TransformCoord vout, vin, matMesh
''            V.Point3.x = vout.x
''            V.Point3.y = vout.y
''            V.Point3.z = vout.z
''
''        Next
''    End If
'
'
'    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyRotate Degrees, m, ApplyTo
'    Next
'    stacked = stacked - 1
'
'End Sub
'
'Private Static Sub ApplyScaled(ByRef Scalar As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
'    'modifies the actual 3D object's pointsm only once a frame for speed consideration
'    If Scalar Is Nothing Then Exit Sub
'
'    Dim vin As D3DVECTOR
'    Dim vout As D3DVECTOR
'    Static matMesh As D3DMATRIX
'    Static stacked As Integer
'
'    If stacked = 0 Then
'        D3DXMatrixIdentity matMesh
'        D3DXMatrixScaling matMesh, IIf(Scalar.X = 0, 1, Scalar.X), IIf(Scalar.Y = 0, 1, Scalar.Y), IIf(Scalar.z = 0, 1, Scalar.z)
'    End If
'
'    Set ApplyTo.Scaled = Scalar
'    Set ApplyTo.Absolute.Scaled = Scalar
'
'    Dim V As Matter
'    For Each V In ApplyTo.Volume
'
'        vin.X = V.Point1.X
'        vin.Y = V.Point1.Y
'        vin.z = V.Point1.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point1.X = vout.X
'        V.Point1.Y = vout.Y
'        V.Point1.z = vout.z
'
'        vin.X = V.Point2.X
'        vin.Y = V.Point2.Y
'        vin.z = V.Point2.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point2.X = vout.X
'        V.Point2.Y = vout.Y
'        V.Point2.z = vout.z
'
'        vin.X = V.Point3.X
'        vin.Y = V.Point3.Y
'        vin.z = V.Point3.z
'        D3DXVec3TransformCoord vout, vin, matMesh
'        V.Point3.X = vout.X
'        V.Point3.Y = vout.Y
'        V.Point3.z = vout.z
'
'        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
'    Next
'
'    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyScaled Scalar, m, ApplyTo
'    Next
'    stacked = stacked - 1
'End Sub


Private Sub ApplyOrigin(ByRef Origin As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Origin Is Nothing Then Exit Sub


    Static vin As D3DVECTOR
    Static vout As D3DVECTOR
    Static matMesh As D3DMATRIX
    'Static stacked As Integer

'    If Not Parent Is Nothing Then
'        Set ApplyTo.Origin = VectorAddition(ApplyTo.Origin, VectorRotateAxis(Origin, ApplyTo.Rotate))
'    Else
        Set ApplyTo.Origin = Origin
'    End If

    'If stacked = 0 Then
        Set ApplyTo.Absolute.Origin = ApplyTo.Origin

    'End If

'    stacked = stacked + 1
'
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyOrigin VectorAddition(ApplyTo.Origin, Origin), m, ApplyTo
'        Next
'    End If
'
'    stacked = stacked - 1
End Sub


Private Sub ApplyRotate(ByRef Degrees As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Degrees Is Nothing Then Exit Sub

    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Static matMesh As D3DMATRIX
    Static stacked As Integer


    D3DXMatrixIdentity matMesh
    
    If stacked = 0 Then
        
        D3DXMatrixRotationYawPitchRoll matMesh, Degrees.X, Degrees.Y, Degrees.z

        Set ApplyTo.Rotate = Degrees
        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate

    Else
        CommitRotate ApplyTo, Parent
        D3DXMatrixRotationYawPitchRoll matMesh, ApplyTo.Rotate.X, ApplyTo.Rotate.Y, ApplyTo.Rotate.z
    
    End If
    
    If TypeName(ApplyTo) <> "Planet" Then
        Dim V As Matter
        For Each V In ApplyTo.Volume

            vin.X = V.Point1.X
            vin.Y = V.Point1.Y
            vin.z = V.Point1.z
            D3DXVec3TransformCoord vout, vin, matMesh
            V.Point1.X = vout.X
            V.Point1.Y = vout.Y
            V.Point1.z = vout.z

            vin.X = V.Point2.X
            vin.Y = V.Point2.Y
            vin.z = V.Point2.z
            D3DXVec3TransformCoord vout, vin, matMesh
            V.Point2.X = vout.X
            V.Point2.Y = vout.Y
            V.Point2.z = vout.z

            vin.X = V.Point3.X
            vin.Y = V.Point3.Y
            vin.z = V.Point3.z
            D3DXVec3TransformCoord vout, vin, matMesh
            V.Point3.X = vout.X
            V.Point3.Y = vout.Y
            V.Point3.z = vout.z

            'Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)

        Next
    End If

'    If Not Parent Is Nothing Then
'        Set ApplyTo.Rotate = Degrees
'    Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
'    End If
    
    
        

        


      '  D3DXMatrixRotationYawPitchRoll matMesh, -ApplyTo.Rotate.X + Parent.Rotate.X, -ApplyTo.Rotate.Y + Parent.Rotate.Y, -ApplyTo.Rotate.z + Parent.Rotate.z




'        Set ApplyTo.Origin = VectorDeduction(ApplyTo.Origin, VectorRotateAxis(VectorDeduction(ApplyTo.Origin, Parent.Origin), VectorDeduction(Degrees, ApplyTo.Rotate)))
'        Set ApplyTo.Rotate = VectorDeduction(VectorDeduction(Degrees, VectorDeduction(Degrees, ApplyTo.Rotate)), VectorDeduction(Degrees, ApplyTo.Rotate))
'        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
        
        
       ' Set ApplyTo.Origin = VectorDeduction(ApplyTo.Origin, VectorRotateAxis(VectorDeduction(ApplyTo.Origin, Parent.Origin), Degrees))


     ' Set ApplyTo.Origin = VectorDeduction(VectorAddition(ApplyTo.Origin, VectorRotateAxis(VectorRotateAxis(VectorDeduction(Parent.Origin, ApplyTo.Origin), AngleAxisAddition(Parent.Rotate, ApplyTo.Rotate)), Degrees)), Parent.Origin) 'VectorDeduction(ApplyTo.Origin, VectorRotateAxis(ApplyTo.Origin, Degrees))
        
     '   Set ApplyTo.Origin = VectorAddition(VectorRotateAxis(ApplyTo.Origin, VectorNegative(ApplyTo.Rotate)), VectorRotateAxis(ApplyTo.Rotate, Degrees))
        
       '  Set ApplyTo.Origin = VectorRotateAxis(VectorRotateAxis(ApplyTo.Origin, VectorNegative(ApplyTo.Rotate)), Degrees)
        
       ' Set ApplyTo.Absolute.Origin = ApplyTo.Origin


'
        'D3DXMatrixIdentity matMesh
        'D3DXMatrixRotationYawPitchRoll matMesh, -ApplyTo.Rotate.X - Degrees.X - Parent.Rotate.X, -ApplyTo.Rotate.Y - Degrees.X - Parent.Rotate.Y, -ApplyTo.Rotate.z - Degrees.z - Parent.Rotate.z
        
       ' D3DXMatrixIdentity matMesh
       ' D3DXMatrixRotationYawPitchRoll matMesh, Degrees.X, Degrees.Y, Degrees.z


'            vin.X = ApplyTo.Origin.X
'            vin.Y = ApplyTo.Origin.Y
'            vin.z = ApplyTo.Origin.z
'            D3DXVec3TransformCoord vout, vin, matMesh
'            ApplyTo.Origin.X = vout.X
'            ApplyTo.Origin.Y = vout.Y
'            ApplyTo.Origin.z = vout.z

    



    
    stacked = stacked + 1
'    Dim m As Molecule
'    For Each m In RangedMolecules(ApplyTo)
'        ApplyRotate Degrees, m, ApplyTo
'    Next
    
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyRotate Degrees, m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyRotate Degrees, m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyRotate Degrees, m, ApplyTo
'        Next
'    End If
    stacked = stacked - 1

End Sub

Private Static Sub ApplyScaled(ByRef Scalar As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Scalar Is Nothing Then Exit Sub

    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim matMesh As D3DMATRIX
    Static stacked As Integer

    D3DXMatrixIdentity matMesh
    D3DXMatrixScaling matMesh, Scalar.X, Scalar.Y, Scalar.z

    Set ApplyTo.Scaled = VectorAddition(ApplyTo.Scaled, Scalar)

    If stacked = 0 Then
        Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
    End If

    Dim V As Matter
    For Each V In ApplyTo.Volume

        vin.X = V.Point1.X
        vin.Y = V.Point1.Y
        vin.z = V.Point1.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point1.X = vout.X
        V.Point1.Y = vout.Y
        V.Point1.z = vout.z

        vin.X = V.Point2.X
        vin.Y = V.Point2.Y
        vin.z = V.Point2.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point2.X = vout.X
        V.Point2.Y = vout.Y
        V.Point2.z = vout.z

        vin.X = V.Point3.X
        vin.Y = V.Point3.Y
        vin.z = V.Point3.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point3.X = vout.X
        V.Point3.Y = vout.Y
        V.Point3.z = vout.z

        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
    Next

'    stacked = stacked + 1
'
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyScaled VectorAddition(m.Scaled, Scalar), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'        Next
'    End If
'
'    stacked = stacked - 1
End Sub

Private Static Sub ApplyOffset(ByRef Offset As Point, ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    If Offset Is Nothing Then Exit Sub

    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim matMesh As D3DMATRIX

    D3DXMatrixIdentity matMesh
    D3DXMatrixTranslation matMesh, -ApplyTo.Offset.X + Offset.X, -ApplyTo.Offset.Y + Offset.Y, -ApplyTo.Offset.z + Offset.z

    Set ApplyTo.Offset = Offset
    Set ApplyTo.Absolute.Offset = ApplyTo.Offset

    Dim V As Matter
    For Each V In ApplyTo.Volume

        vin.X = V.Point1.X
        vin.Y = V.Point1.Y
        vin.z = V.Point1.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point1.X = vout.X
        V.Point1.Y = vout.Y
        V.Point1.z = vout.z

        vin.X = V.Point2.X
        vin.Y = V.Point2.Y
        vin.z = V.Point2.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point2.X = vout.X
        V.Point2.Y = vout.Y
        V.Point2.z = vout.z

        vin.X = V.Point3.X
        vin.Y = V.Point3.Y
        vin.z = V.Point3.z
        D3DXVec3TransformCoord vout, vin, matMesh
        V.Point3.X = vout.X
        V.Point3.Y = vout.Y
        V.Point3.z = vout.z

        Set V.Normal = TriangleNormal(V.Point1, V.Point2, V.Point3)
    Next

'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyOrigin VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'        Next
'    End If

End Sub

Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
End Sub
Private Sub LocPos(ByRef Origin As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Origin.X <> 0 Or Origin.Y <> 0 Or Origin.z <> 0 Then
        Dim o As Orbit
        Dim m As Molecule
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        LocPos Origin, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            LocPos Origin, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                CommitOrigin ApplyTo, Parent
                If Relative Then
                    If Not Parent Is Nothing Then
                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Parent.Rotate), ApplyTo.Rotate))
                    ElseIf Not Camera.Planet Is Nothing Then
                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Camera.Planet.Rotate), ApplyTo.Rotate))
                    Else
                        Set ApplyTo.Relative.Origin = VectorRotateAxis(Origin, ApplyTo.Rotate)
                    End If
                Else
                    Set ApplyTo.Absolute.Origin = Origin
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    LocPos Origin, Relative, m, ApplyTo
'                Next
        End Select
    End If
End Sub
Public Sub Rotation(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Orientate(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
End Sub
Private Sub RotOri(ByRef Degrees As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Degrees.X <> 0 Or Degrees.Y <> 0 Or Degrees.z <> 0 Then
        Dim m As Molecule
        Dim o As Point
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        RotOri Degrees, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            RotOri Degrees, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                CommitRotate ApplyTo, Parent
                If Relative Then
                    If Not Parent Is Nothing Then
                        Set ApplyTo.Relative.Rotate = AngleAxisDeduction(AngleAxisAddition(AngleAxisDeduction(ApplyTo.Rotate, Parent.Rotate), Degrees), ApplyTo.Rotate)
                    ElseIf Not Camera.Planet Is Nothing Then
                        Set ApplyTo.Relative.Rotate = Degrees
                    Else
                        'Set ApplyTo.Relative.Rotate = AngleAxisDeduction(AngleAxisAddition(ApplyTo.Rotate, Degrees), ApplyTo.Rotate)
                        Set ApplyTo.Relative.Rotate = Degrees
                    End If

          '          Set ApplyTo.Relative.Rotate = Degrees
                Else
                    Set ApplyTo.Absolute.Rotate = Degrees
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    RotOri Degrees, Relative, m, ApplyTo
'                Next

        End Select
    End If
End Sub

Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitScaling(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
End Sub
Private Sub ScaExp(ByRef Scalar As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Abs(Scalar.X) <> 1 Or Abs(Scalar.Y) <> 1 Or Abs(Scalar.z) <> 1 Then
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        ScaExp Scalar, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        If p.Ranges.W - Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0) > 0 Then
                            ScaExp Scalar, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'change all molecules with in the specified planets range
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Scaled = Scalar
                Else
                    Set ApplyTo.Absolute.Scaled = Scalar
                End If
'                For Each m In RangedMolecules(ApplyTo)
'                    ScaExp Scalar, Relative, m, ApplyTo
'                Next

        End Select
    End If
End Sub
Public Sub Displace(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, False, ApplyTo, Parent  'location is changing the origin to absolute
End Sub
Public Sub Balanced(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False
    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True
End Sub
Private Sub DisBal(ByRef Offset As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Offset.X <> 0 Or Offset.Y <> 0 Or Offset.z <> 0 Then
        Dim dist As Single
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If p.Ranges.W = -1 Then
                        DisBal Offset, Relative, p, ApplyTo
                    ElseIf p.Ranges.W > 0 Then
                        dist = Distance(p.Origin.X, p.Origin.Y, p.Origin.z, 0, 0, 0)
                        If p.Ranges.W - dist > 0 Then
                            DisBal Offset, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'change all molecules with in the specified planets range
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Offset = Offset
                Else
                    Set ApplyTo.Absolute.Offset = Offset
                End If

'                For Each m In RangedMolecules(ApplyTo)
'                    DisBal Offset, Relative, m, ApplyTo
'                Next
        End Select
    End If
End Sub

Private Function RangedMolecules(ByRef ApplyTo As Molecule) As NTNodes10.Collection
    Set RangedMolecules = New NTNodes10.Collection

'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.W = -1 Then
'                ApplyScaled Scalar, m, ApplyTo
'            ElseIf ApplyTo.Ranges.W > 0 Then
'                If ApplyTo.Ranges.W - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyScaled Scalar, m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyScaled Scalar, m, ApplyTo
'        Next
'    End If
    Dim m As Molecule
    Dim dist As Single
    For Each m In Molecules
        If ((m.Parent Is Nothing) And (Not TypeName(ApplyTo) = "Planet")) Or (TypeName(ApplyTo) = "Planet") Then
            If ApplyTo.Ranges.W = -1 Then
                RangedMolecules.Add m, m.Key
            ElseIf ApplyTo.Ranges.W > 0 Then
                dist = Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z)
                If ApplyTo.Ranges.W - dist > 0 Then
                    RangedMolecules.Add m, m.Key
                End If
            End If
        End If
    Next
    For Each m In ApplyTo.Molecules
        If Not RangedMolecules.Exists(m.Key) Then RangedMolecules.Add m, m.Key
    Next
End Function

Public Sub CommitRoutine(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal DoRotate As Boolean, ByVal DoScaled As Boolean, ByVal DoOrigin As Boolean, ByVal DoOffset As Boolean, ByVal DoAbsolute As Boolean, ByVal DoRelative As Boolean)
    'partial to committing a 3d objects properties during calls that may not sum, for retaining other properties needing change first and entirety per frame
    Static stacked As Boolean
    If Not stacked Then
        stacked = True

        'any absolute position comes first, pending is a difference from the actual
        If (Not ApplyTo.Absolute.Rotate.Equals(ApplyTo.Rotate)) And ((DoRotate And DoAbsolute) Or ((Not DoRotate) And (Not DoAbsolute))) Then
            ApplyRotate AngleAxisDeduction(ApplyTo.Absolute.Rotate, ApplyTo.Rotate), ApplyTo, Parent
            Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
        End If
        If (Not ApplyTo.Absolute.Scaled.Equals(ApplyTo.Scaled)) And ((DoScaled And DoAbsolute) Or ((Not DoScaled) And (Not DoAbsolute))) Then
            ApplyScaled VectorDeduction(ApplyTo.Absolute.Scaled, ApplyTo.Scaled), ApplyTo, Parent
            Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
        End If
        If (Not ApplyTo.Absolute.Origin.Equals(ApplyTo.Origin)) And ((DoOrigin And DoAbsolute) Or ((Not DoOrigin) And (Not DoAbsolute))) Then
            ApplyOrigin VectorDeduction(ApplyTo.Absolute.Origin, ApplyTo.Origin), ApplyTo, Parent
            Set ApplyTo.Absolute.Origin = ApplyTo.Origin
        End If
        If (Not ApplyTo.Absolute.Offset.Equals(ApplyTo.Offset)) And ((DoOffset And DoAbsolute) Or ((Not DoOffset) And (Not DoAbsolute))) Then
            ApplyOffset VectorDeduction(ApplyTo.Absolute.Offset, ApplyTo.Offset), ApplyTo, Parent
            Set ApplyTo.Absolute.Offset = ApplyTo.Offset
        End If

        'relative positioning comes secondly, pending is there is any value not empty
        If (ApplyTo.Relative.Rotate.X <> 0 Or ApplyTo.Relative.Rotate.Y <> 0 Or ApplyTo.Relative.Rotate.z <> 0) And ((DoRotate And DoRelative) Or ((Not DoRotate) And (Not DoRelative))) Then
            ApplyRotate AngleAxisAddition(ApplyTo.Relative.Rotate, ApplyTo.Rotate), ApplyTo, Parent
            Set ApplyTo.Relative.Rotate = Nothing
        End If
        If (Abs(ApplyTo.Relative.Scaled.X) <> 1 Or Abs(ApplyTo.Relative.Scaled.Y) <> 1 Or Abs(ApplyTo.Relative.Scaled.z) <> 1) And ((DoScaled And DoRelative) Or ((Not DoScaled) And (Not DoRelative))) Then
            ApplyScaled VectorAddition(ApplyTo.Relative.Scaled, ApplyTo.Scaled), ApplyTo, Parent
            Set ApplyTo.Relative.Scaled = Nothing
        End If
        If (ApplyTo.Relative.Origin.X <> 0 Or ApplyTo.Relative.Origin.Y <> 0 Or ApplyTo.Relative.Origin.z <> 0) And ((DoOrigin And DoRelative) Or ((Not DoOrigin) And (Not DoRelative))) Then
            ApplyOrigin VectorAddition(ApplyTo.Relative.Origin, ApplyTo.Origin), ApplyTo, Parent
            Set ApplyTo.Relative.Origin = Nothing
        End If
        If (ApplyTo.Relative.Offset.X <> 0 Or ApplyTo.Relative.Offset.Y <> 0 Or ApplyTo.Relative.Offset.z <> 0) And ((DoOffset And DoRelative) Or ((Not DoOffset) And (Not DoRelative))) Then
            ApplyOffset VectorAddition(ApplyTo.Relative.Offset, ApplyTo.Offset), ApplyTo, Parent
            Set ApplyTo.Relative.Offset = Nothing
        End If

        stacked = False
    End If
End Sub

Public Sub Begin(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    'called once per frame committing changes the last frame has waiting in object properties in entirety
    Dim m As Molecule
    Dim p As Planet

    For Each p In Planets

        CommitRoutine p, Nothing, True, False, False, False, True, False
        CommitRoutine p, Nothing, False, True, False, False, True, False
        CommitRoutine p, Nothing, False, False, True, False, True, False
        CommitRoutine p, Nothing, False, False, False, True, True, False

        CommitRoutine p, Nothing, True, False, False, False, False, True
        CommitRoutine p, Nothing, False, True, False, False, False, True
        CommitRoutine p, Nothing, False, False, True, False, False, True
        CommitRoutine p, Nothing, False, False, False, True, False, True

        Set p.Relative = Nothing

    Next

    For Each m In Molecules

        If m.Parent Is Nothing Then

            CommitRoutine m, Nothing, True, False, False, False, True, False
            CommitRoutine m, Nothing, False, True, False, False, True, False
            CommitRoutine m, Nothing, False, False, True, False, True, False
            CommitRoutine m, Nothing, False, False, False, True, True, False

            CommitRoutine m, Nothing, True, False, False, False, False, True
            CommitRoutine m, Nothing, False, True, False, False, False, True
            CommitRoutine m, Nothing, False, False, True, False, False, True
            CommitRoutine m, Nothing, False, False, False, True, False, True

            Set m.Relative = Nothing

        End If

    Next

'        If Not m.Absolute.Rotate.Equals(m.Rotate) Then
'            ApplyRotate m, m.Absolute.Rotate
'        End If
'        If Not m.Absolute.Scaled.Equals(m.Scaled) Then
'            ApplyScaled m, m.Absolute.Scaled
'        End If
'        If Not m.Absolute.Origin.Equals(m.Origin) Then
'            ApplyOrigin m, m.Absolute.Origin
'        End If
'        If Not m.Absolute.Offset.Equals(m.Offset) Then
'            ApplyOffset m, m.Absolute.Offset
'        End If
'
'        If Abs(m.Relative.Scaled.x) <> 1 Or Abs(m.Relative.Scaled.y) <> 1 Or Abs(m.Relative.Scaled.z) <> 1 Then
'            ApplyScaled m, m.Relative.Scaled
'        End If
'        If m.Relative.Rotate.x <> 0 Or m.Relative.Rotate.y <> 0 Or m.Relative.Rotate.z <> 0 Then
'            ApplyRotate m, m.Relative.Rotate
'        End If
'        If m.Relative.Origin.x <> 0 Or m.Relative.Origin.y <> 0 Or m.Relative.Origin.z <> 0 Then
'            ApplyOrigin m, m.Relative.Origin
'        End If

End Sub

Public Sub Finish(ByRef UserControl As Macroscopic, ByRef MoleculeView As Molecule)
    'called once per frame drawing the objects, with out any of the current frame object
    'properties modifying calls included for latent collision checking rollback

    DDevice.SetRenderState D3DRS_ZENABLE, 1

    DDevice.SetRenderState D3DRS_CLIPPING, 1

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

    DDevice.SetVertexShader FVF_RENDER
    DDevice.SetPixelShader PixelShaderDefault

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

    DDevice.SetMaterial LucentMaterial
    DDevice.SetTexture 0, Nothing
    DDevice.SetMaterial GenericMaterial
    DDevice.SetTexture 1, Nothing
    
    Dim p As Planet
    For Each p In Planets
        Iterate p.Molecules, False
    Next

    Iterate Molecules, True

End Sub

Private Sub Iterate(ByRef col As Object, ByVal NoParentOnly As Boolean)

    Dim m As Molecule
    For Each m In col
        If NoParentOnly Then
            If m.Parent Is Nothing Then
                Render m, Nothing
               
            End If
        Else
            Render m, m.Parent
        End If
    Next

End Sub

Private Sub Render(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    Static stacked As Integer

    Static matMat As D3DMATRIX

    If stacked = 0 Then
        D3DXMatrixIdentity matMat
    End If

    Dim matRoll As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matYaw As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matScale As D3DMATRIX


    
    If Not Parent Is Nothing Then
        D3DXMatrixRotationX matPitch, Parent.Rotate.X
        D3DXMatrixMultiply matMat, matPitch, matMat

        D3DXMatrixRotationY matYaw, Parent.Rotate.Y
        D3DXMatrixMultiply matMat, matYaw, matMat

        D3DXMatrixRotationZ matRoll, Parent.Rotate.z
        D3DXMatrixMultiply matMat, matRoll, matMat
    ElseIf Not Camera.Planet Is Nothing Then
        D3DXMatrixRotationX matPitch, Camera.Planet.Rotate.X
        D3DXMatrixMultiply matMat, matPitch, matMat

        D3DXMatrixRotationY matYaw, Camera.Planet.Rotate.Y
        D3DXMatrixMultiply matMat, matYaw, matMat

        D3DXMatrixRotationZ matRoll, Camera.Planet.Rotate.z
        D3DXMatrixMultiply matMat, matRoll, matMat
    End If

    DDevice.SetTransform D3DTS_WORLD, matMat

    
    

    Dim V As Matter
    For Each V In ApplyTo.Volume
        'update the directx and collision array's then render the object
        VertexDirectX((V.TriangleIndex * 3) + 0).X = V.Point1.X + ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 0).Y = V.Point1.Y + ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 0).z = V.Point1.z + ApplyTo.Origin.z

        VertexDirectX((V.TriangleIndex * 3) + 1).X = V.Point2.X + ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 1).Y = V.Point2.Y + ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 1).z = V.Point2.z + ApplyTo.Origin.z

        VertexDirectX((V.TriangleIndex * 3) + 2).X = V.Point3.X + ApplyTo.Origin.X
        VertexDirectX((V.TriangleIndex * 3) + 2).Y = V.Point3.Y + ApplyTo.Origin.Y
        VertexDirectX((V.TriangleIndex * 3) + 2).z = V.Point3.z + ApplyTo.Origin.z

        If Not Parent Is Nothing Then

            VertexDirectX((V.TriangleIndex * 3) + 0).X = VertexDirectX((V.TriangleIndex * 3) + 0).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 0).Y = VertexDirectX((V.TriangleIndex * 3) + 0).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 0).z = VertexDirectX((V.TriangleIndex * 3) + 0).z + Parent.Origin.z

            VertexDirectX((V.TriangleIndex * 3) + 1).X = VertexDirectX((V.TriangleIndex * 3) + 1).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 1).Y = VertexDirectX((V.TriangleIndex * 3) + 1).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 1).z = VertexDirectX((V.TriangleIndex * 3) + 1).z + Parent.Origin.z

            VertexDirectX((V.TriangleIndex * 3) + 2).X = VertexDirectX((V.TriangleIndex * 3) + 2).X + Parent.Origin.X
            VertexDirectX((V.TriangleIndex * 3) + 2).Y = VertexDirectX((V.TriangleIndex * 3) + 2).Y + Parent.Origin.Y
            VertexDirectX((V.TriangleIndex * 3) + 2).z = VertexDirectX((V.TriangleIndex * 3) + 2).z + Parent.Origin.z

        End If

        VertexDirectX(V.TriangleIndex * 3 + 0).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 0).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 0).Nz = V.Normal.z

        VertexDirectX(V.TriangleIndex * 3 + 1).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 1).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 1).Nz = V.Normal.z

        VertexDirectX(V.TriangleIndex * 3 + 2).NX = V.Normal.X
        VertexDirectX(V.TriangleIndex * 3 + 2).NY = V.Normal.Y
        VertexDirectX(V.TriangleIndex * 3 + 2).Nz = V.Normal.z

        VertexXAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).X
        VertexXAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).X
        VertexXAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).X

        VertexYAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).Y
        VertexYAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).Y
        VertexYAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).Y

        VertexZAxis(0, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 0).z
        VertexZAxis(1, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 1).z
        VertexZAxis(2, V.TriangleIndex) = VertexDirectX(V.TriangleIndex * 3 + 2).z

         If ApplyTo.Visible And (Not (TypeName(ApplyTo) = "Planet")) Then
             If Not (V.Translucent Or V.Transparent) Then
                 DDevice.SetMaterial GenericMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
                 DDevice.SetTexture 1, Nothing
             Else
                 DDevice.SetMaterial LucentMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 0, Files(V.TextureIndex).Data
                 DDevice.SetMaterial GenericMaterial
                 If V.TextureIndex > 0 Then DDevice.SetTexture 1, Files(V.TextureIndex).Data
             End If

             DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VertexDirectX((V.TriangleIndex * 3)), Len(VertexDirectX(0))
         End If
    Next

    stacked = stacked + 1
    Dim m As Molecule

    For Each m In ApplyTo.Molecules
        Render m, ApplyTo
    Next
    stacked = stacked - 1
End Sub

Private Function BuildArrays() As Long
    ReDim Preserve TriangleFace(0 To 5, 0 To TriangleCount) As Single
    ReDim Preserve VertexXAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexYAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexZAxis(0 To 2, 0 To TriangleCount) As Single
    BuildArrays = (((TriangleCount + 1) * 3) - 1)
    ReDim Preserve VertexDirectX(0 To BuildArrays) As MyVertex
    ReDim Preserve ScreenDirectX(0 To BuildArrays) As MyScreen
    BuildArrays = BuildArrays - 2
End Function

Public Function CreateMoleculeFace(ByRef TextureFileName As String, ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, ByRef P4 As Point, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Molecule
    If (((Not (p1.Equals(p2) Or p1.Equals(p3) Or p1.Equals(P4))) And _
        (Not (p3.Equals(p2) Or p3.Equals(p1) Or p3.Equals(P4))) And _
        (Not (p2.Equals(p1) Or p2.Equals(p3) Or p2.Equals(P4))) And _
        (Not (P4.Equals(p2) Or P4.Equals(p3) Or P4.Equals(p1)))) And _
        PathExists(TextureFileName, True)) Then
        
        Dim r As New Molecule
        Set r.Volume = CreateVolumeFace(TextureFileName, p1, p2, p3, P4, ScaleX, ScaleY)
        r.Visible = True
        Set CreateMoleculeFace = r
    End If
End Function
Public Function CreateVolumeFace(ByRef TextureFileName As String, ByRef p1 As Point, ByRef p2 As Point, ByRef p3 As Point, ByRef P4 As Point, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Volume
    If (((Not (p1.Equals(p2) Or p1.Equals(p3) Or p1.Equals(P4))) And _
        (Not (p3.Equals(p2) Or p3.Equals(p1) Or p3.Equals(P4))) And _
        (Not (p2.Equals(p1) Or p2.Equals(p3) Or p2.Equals(P4))) And _
        (Not (P4.Equals(p2) Or P4.Equals(p3) Or P4.Equals(p1)))) And _
        PathExists(TextureFileName, True)) Then
        If ScaleX = 0 Then ScaleX = 1
        If ScaleY = 0 Then ScaleY = 1
        
        Dim vol As New Volume
        Dim m As New Matter
        With m
            .TriangleIndex = TriangleCount
            BuildArrays

            .Index1 = PointCache(p1)
            .Index2 = PointCache(p2)
            .Index3 = PointCache(p3)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .V1 = ScaleY
            .U2 = ScaleX
            .V2 = ScaleY
            .U3 = ScaleX

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.z
            VertexZAxis(1, TriangleCount) = .Point2.z
            VertexZAxis(2, TriangleCount) = .Point3.z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 0

            .ObjectIndex = ObjectCount
            .FaceIndex = 0
            .TextureIndex = GetFileIndex(TextureFileName)
            If TextureFileName <> "" Then
                If Files(.TextureIndex).Data Is Nothing Then
                    Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                    ImageDimensions TextureFileName, Files(.TextureIndex).Size
                End If
            End If
            
            VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
            VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
            ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
            ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
            ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
            ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
            ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
            ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
            ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
            ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
            ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
        End With
        TriangleCount = TriangleCount + 1
        vol.Add m
                
        Set m = New Matter
        With m

            .TriangleIndex = TriangleCount
            BuildArrays

            .Index1 = PointCache(p1)
            .Index2 = PointCache(p3)
            .Index3 = PointCache(P4)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .V1 = ScaleY
            .U2 = ScaleX

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.z
            VertexZAxis(1, TriangleCount) = .Point2.z
            VertexZAxis(2, TriangleCount) = .Point3.z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 1

            .ObjectIndex = ObjectCount
            .FaceIndex = 1
            .TextureIndex = GetFileIndex(TextureFileName)
            If TextureFileName <> "" Then
                If Files(.TextureIndex).Data Is Nothing Then
                    Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                    ImageDimensions TextureFileName, Files(.TextureIndex).Size
                End If
            End If
            
            VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
            VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z

            VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
            VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z

            VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
            VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z

            VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
            VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
            VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

            VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
            VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
            VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
            VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
            VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
            VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
            
            ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
            ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
            ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
            ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
            ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
            ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
            ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
            ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
            ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
            ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
            ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)

            ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
            ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
            ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
            ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
            ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
            ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
        End With
        TriangleCount = TriangleCount + 1
        vol.Add m
        
        ObjectCount = ObjectCount + 1
        
        Set CreateVolumeFace = vol

    End If
End Function

Public Function CreateMoleculeLanding(ByRef TextureFileName As String, ByVal OuterRadii As Single, ByVal RadiiSegments As Single, Optional ByVal InnerRadii As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Molecule
    If OuterRadii <= 0 Then
        Err.Raise 8, , "OuterRadii must be above 0."
    ElseIf InnerRadii < 0 Then
        Err.Raise 8, , "InnerRadii must be 0 or above."
    ElseIf OuterRadii < InnerRadii Then
        Err.Raise 8, , "OuterRadii must be below InnerRadii."
    ElseIf RadiiSegments < 3 Then
        Err.Raise 8, , "RadiiSegments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        
        Dim r As New Molecule
        Set r.Volume = CreateVolumeLanding(TextureFileName, OuterRadii, RadiiSegments, InnerRadii, ScaleX, ScaleY)
        r.Visible = True
        Set CreateMoleculeLanding = r
    End If

End Function


Public Function CreateVolumeLanding(ByRef TextureFileName As String, ByVal OuterRadii As Single, ByVal RadiiSegments As Single, Optional ByVal InnerRadii As Single = 0, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1) As Volume
    If OuterRadii <= 0 Then
        Err.Raise 8, , "OuterRadii must be above 0."
    ElseIf InnerRadii < 0 Then
        Err.Raise 8, , "InnerRadii must be 0 or above."
    ElseIf OuterRadii < InnerRadii Then
        Err.Raise 8, , "OuterRadii must be below InnerRadii."
    ElseIf RadiiSegments < 3 Then
        Err.Raise 8, , "RadiiSegments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        Dim i As Long
        Dim g As Single
        Dim A As Double
        Dim l1 As Single
        Dim l2 As Single
        Dim l3 As Single
    
        Dim intX1 As Single
        Dim intX2 As Single
        Dim intX3 As Single
        Dim intX4 As Single
    
        Dim intY1 As Single
        Dim intY2 As Single
        Dim intY3 As Single
        Dim intY4 As Single
        Dim dist1 As Single
        Dim dist2 As Single
        Dim dist3 As Single
        Dim dist4 As Single

        Dim vol As New Volume
        Dim m As Matter

        
        For i = -IIf(InnerRadii > 0, 6, 3) To ((IIf(InnerRadii > 0, 6, 3) * RadiiSegments) - 1) + (IIf(InnerRadii > 0, 6, 3) * 2) Step IIf(InnerRadii > 0, 6, 3)
    
            g = (((360 / RadiiSegments) * (((i + 1) / IIf(InnerRadii > 0, 6, 3)) - 1)) * RADIAN)
    
            intX2 = (OuterRadii * Sin(g))
            intY2 = (-OuterRadii * Cos(g))
            intX3 = (InnerRadii * Sin(g))
            intY3 = (-InnerRadii * Cos(g))
    
            If i >= 0 Then
            
                If (InnerRadii > 0) Then
                    If (i Mod 12) = 0 Then
    
                        dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
                        dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (ScaleX / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (ScaleY / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
    
                    End If

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .V1 = dist2
                        .U2 = dist1
                        .V2 = dist4
                        .U3 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
            
                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                    
                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX4, 0, intY4))
                        .Index3 = PointCache(MakePoint(intX3, 0, intY3))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .V1 = dist2
                        .U2 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                    
                Else
    
                    dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * IIf(i Mod 4 = 0, 1, Sin(g))
                    dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * IIf(i Mod 4 = 0, 1, Cos(g))

                    Set m = New Matter
                    With m
                        .TriangleIndex = TriangleCount
                        BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .U1 = ((ScaleX / dist1) * (dist1 / 100))
                        .V2 = ((ScaleY / dist2) * (dist2 / 100))
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.z, .Point2.X, .Point2.Y, .Point2.z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.z, .Point3.X, .Point3.Y, .Point3.z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.z, .Point1.X, .Point1.Y, .Point1.z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.z
                        VertexZAxis(1, TriangleCount) = .Point2.z
                        VertexZAxis(2, TriangleCount) = .Point3.z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
                        .TextureIndex = GetFileIndex(TextureFileName)
                        If TextureFileName <> "" Then
                            If Files(.TextureIndex).Data Is Nothing Then
                                Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                                ImageDimensions TextureFileName, Files(.TextureIndex).Size
                            End If
                        End If
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        ScreenDirectX(.TriangleIndex * 3 + 0).Y = .Point1.z
                        ScreenDirectX(.TriangleIndex * 3 + 0).z = .Point1.Y
                        ScreenDirectX(.TriangleIndex * 3 + 0).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 0).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        ScreenDirectX(.TriangleIndex * 3 + 1).Y = .Point2.z
                        ScreenDirectX(.TriangleIndex * 3 + 1).z = .Point2.Y
                        ScreenDirectX(.TriangleIndex * 3 + 1).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 1).clr = D3DColorARGB(255, 255, 255, 255)

                        ScreenDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        ScreenDirectX(.TriangleIndex * 3 + 2).Y = .Point3.z
                        ScreenDirectX(.TriangleIndex * 3 + 2).z = .Point3.Y
                        ScreenDirectX(.TriangleIndex * 3 + 2).rhw = 1
                        ScreenDirectX(.TriangleIndex * 3 + 2).clr = D3DColorARGB(255, 255, 255, 255)
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                        VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                        VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                        VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                        
                        VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                        VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                        VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z

                        VertexDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        VertexDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        VertexDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        VertexDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        VertexDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        VertexDirectX(.TriangleIndex * 3 + 2).tv = .V3
                        
                        ScreenDirectX(.TriangleIndex * 3 + 0).tu = .U1
                        ScreenDirectX(.TriangleIndex * 3 + 0).tv = .V1
                        ScreenDirectX(.TriangleIndex * 3 + 1).tu = .U2
                        ScreenDirectX(.TriangleIndex * 3 + 1).tv = .V2
                        ScreenDirectX(.TriangleIndex * 3 + 2).tu = .U3
                        ScreenDirectX(.TriangleIndex * 3 + 2).tv = .V3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                    vol.Add m
                End If
    
            End If
    
            intX1 = intX2
            intY1 = intY2
            intX4 = intX3
            intY4 = intY3
    
        Next
        
        ObjectCount = ObjectCount + 1

        Set CreateVolumeLanding = vol
    End If
    
End Function

Public Function CreateMoleculeMesh(ByVal DirectXFileName As String) As Molecule
    If PathExists(DirectXFileName, True) Then
    
        Dim r As New Molecule
        Set r.Volume = CreateVolumeMesh(DirectXFileName)
        r.Visible = True
        Set CreateMoleculeMesh = r
    End If
End Function
Public Function CreateVolumeMesh(ByVal DirectXFileName As String) As Volume
    If PathExists(DirectXFileName, True) Then

        Dim MeshVerticies() As D3DVERTEX
        Dim MeshIndicies() As Integer

        Dim nMaterials As Long
        Dim nMatBuffer As D3DXBuffer
                
        Dim Mesh As D3DXMesh
        Set Mesh = D3DX.LoadMeshFromX(DirectXFileName, D3DXMESH_DYNAMIC, DDevice, Nothing, nMatBuffer, nMaterials)

        Dim Index As Long
        Dim cnt As Long

        Dim VD As D3DVERTEXBUFFER_DESC
        Mesh.GetVertexBuffer.GetDesc VD
        ReDim MeshVerticies(0 To 0) As D3DVERTEX
        ReDim MeshVerticies(0 To ((VD.Size \ Len(MeshVerticies(0))) - 1)) As D3DVERTEX
        D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, VD.Size, 0, MeshVerticies(0)

        Dim ID As D3DINDEXBUFFER_DESC
        Mesh.GetIndexBuffer.GetDesc ID
        ReDim MeshIndicies(0 To 0) As Integer
        ReDim MeshIndicies(0 To ((ID.Size \ Len(MeshIndicies(0))) - 1)) As Integer
        D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, ID.Size, 0, MeshIndicies(0)



'        If nMaterials > 0 Then
'
'            Dim P1 As String
'            Dim P2 As String
'            Dim P3 As String
'            Dim P4 As String
'            Dim p5 As String
'            Dim p6 As String
'
'            Dim l3 As Single
'            Dim l1 As Single
'            Dim l2 As Single
'            Dim l4 As Single
'            Dim l5 As Single
'            Dim l6 As Single
'
'            Dim checked As Long
'
'            Index = 0 'start at last point of first triangle where start = 0
'            Do Until checked = Mesh.GetNumFaces  'go for amount of faces least 3
'
'                 If l1 = 0 Then
'                    l1 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P1 = IIf(checked Mod 4 = 0, "1,0", "4,3")
'                    l4 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P4 = IIf(checked Mod 4 = 0, "2,0", "5,3")
'                End If
'
'                 If l2 = 0 Then
'                    l2 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'                    P2 = IIf(checked Mod 4 = 0, "1,2", "4,5")
'                    l5 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    p5 = IIf(checked Mod 4 = 0, "1,0", "4,3")
'                End If
'
'                 If l3 = 0 Then
'                    l3 = Distance(MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z, MeshVerticies(MeshIndicies(Index)).X, MeshVerticies(MeshIndicies(Index)).Y, MeshVerticies(MeshIndicies(Index)).Z)
'                    P3 = IIf(checked Mod 4 = 0, "2,0", "5,3")
'                    l6 = Distance(MeshVerticies(MeshIndicies(Index + 1)).X, MeshVerticies(MeshIndicies(Index + 1)).Y, MeshVerticies(MeshIndicies(Index + 1)).Z, MeshVerticies(MeshIndicies(Index + 2)).X, MeshVerticies(MeshIndicies(Index + 2)).Y, MeshVerticies(MeshIndicies(Index + 2)).Z)
'                    p6 = IIf(checked Mod 4 = 0, "1,2", "4,5")
'                End If
'
'                 Index = Index + 1
'
'                 If l1 = l2 And l2 = l3 And l1 <> 0 Then
'
'                     l2 = 0
'                     l6 = l4
'                     l4 = l5
'                     l5 = 0
'                 Else
'
'                    If l1 <> 0 And l2 <> 0 And l3 <> 0 Then
'
'
'
'
'                        Debug.Print "(" & (Index + CLng(NextArg(P1, ","))) & ", " & (Index + CLng(RemoveArg(P2, ","))) & ", " & (Index + CLng(NextArg(P3, ","))) & ") ";
'                        Debug.Print "(" & (Index + CLng(RemoveArg(P4, ","))) & ", " & (Index + CLng(NextArg(p5, ","))) & ", " & (Index + CLng(NextArg(p6, ","))) & ") ";
'
'                        'SurfaceArea = SurfaceArea + (TriangleAreaByLen(l1, l2, l3) + TriangleAreaByLen(l4, l5, l6))
'
'                        'Volume = Volume + (TriangleVolByLen(l1, l2, l3) + TriangleVolByLen(l4, l5, l6))
'
'                        l1 = 0
'                        l2 = 0
'                        l3 = 0
'                        l4 = 0
'                        l5 = 0
'                        l6 = 0
'                        checked = checked + 2
'                    End If
'
'                    Index = Index + 2
'                End If
'
'            Loop
'
'        End If
'
        'SurfaceArea = (SurfaceArea * 2)





        Dim vol As New Volume
        Dim m As Matter

        Dim Verts() As D3DVERTEX
        Dim chk1 As Integer
        Dim chk2 As Integer
        Dim chk3 As Integer
        Dim chk4 As Integer
        Dim chk5 As Integer
        Dim chk6 As Integer
        chk1 = -1

        Const TrianglePerFace As Long = 2
        Dim TextureFileName As String

        Dim txt As Long
        txt = 1

        Index = 0

        Do While Index <= UBound(MeshIndicies)

            If txt < nMaterials Then
                If D3DX.BufferGetTextureName(nMatBuffer, txt) <> "" Then
                    If PathExists(GetFilePath(DirectXFileName) & "\" & D3DX.BufferGetTextureName(nMatBuffer, txt), True) Then
                        TextureFileName = GetFilePath(DirectXFileName) & "\" & D3DX.BufferGetTextureName(nMatBuffer, txt)
                    End If
                End If
            End If
            txt = txt + 1

            Set m = New Matter
            With m
                .TriangleIndex = TriangleCount
                BuildArrays

                .Index1 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 0)).X, _
                    MeshVerticies(MeshIndicies(Index + 0)).Y, _
                    MeshVerticies(MeshIndicies(Index + 0)).z))

                .Index2 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 1)).X, _
                    MeshVerticies(MeshIndicies(Index + 1)).Y, _
                    MeshVerticies(MeshIndicies(Index + 1)).z))

                .Index3 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 2)).X, _
                    MeshVerticies(MeshIndicies(Index + 2)).Y, _
                    MeshVerticies(MeshIndicies(Index + 2)).z))

                Set .Point1 = Points(.Index1)
                Set .Point2 = Points(.Index2)
                Set .Point3 = Points(.Index3)

                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.z
                VertexZAxis(1, TriangleCount) = .Point2.z
                VertexZAxis(2, TriangleCount) = .Point3.z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.z
                TriangleFace(4, TriangleCount) = ObjectCount
                TriangleFace(5, TriangleCount) = (Index \ 4) + 0

                .ObjectIndex = ObjectCount
                .FaceIndex = (Index \ 4) + 0
                .TextureIndex = GetFileIndex(TextureFileName)
                If TextureFileName <> "" Then
                    If Files(.TextureIndex).Data Is Nothing Then
                        Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                        ImageDimensions TextureFileName, Files(.TextureIndex).Size
                    End If
                End If

                VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 0)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 0)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 1)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 1)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 2).tu = MeshVerticies(MeshIndicies(Index + 2)).tu
                VertexDirectX(.TriangleIndex * 3 + 2).tv = MeshVerticies(MeshIndicies(Index + 2)).tv

            End With
            TriangleCount = TriangleCount + 1
            vol.Add m

            Set m = New Matter
            With m
                .TriangleIndex = TriangleCount
                BuildArrays

                .Index1 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 3)).X, _
                    MeshVerticies(MeshIndicies(Index + 3)).Y, _
                    MeshVerticies(MeshIndicies(Index + 3)).z))

                .Index2 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 4)).X, _
                    MeshVerticies(MeshIndicies(Index + 4)).Y, _
                    MeshVerticies(MeshIndicies(Index + 4)).z))

                .Index3 = PointCache(MakePoint( _
                    MeshVerticies(MeshIndicies(Index + 5)).X, _
                    MeshVerticies(MeshIndicies(Index + 5)).Y, _
                    MeshVerticies(MeshIndicies(Index + 5)).z))

                Set .Point1 = Points(.Index1)
                Set .Point2 = Points(.Index2)
                Set .Point3 = Points(.Index3)

                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

                VertexXAxis(0, TriangleCount) = .Point1.X
                VertexXAxis(1, TriangleCount) = .Point2.X
                VertexXAxis(2, TriangleCount) = .Point3.X

                VertexYAxis(0, TriangleCount) = .Point1.Y
                VertexYAxis(1, TriangleCount) = .Point2.Y
                VertexYAxis(2, TriangleCount) = .Point3.Y

                VertexZAxis(0, TriangleCount) = .Point1.z
                VertexZAxis(1, TriangleCount) = .Point2.z
                VertexZAxis(2, TriangleCount) = .Point3.z

                TriangleFace(0, TriangleCount) = .Normal.X
                TriangleFace(1, TriangleCount) = .Normal.Y
                TriangleFace(2, TriangleCount) = .Normal.z
                TriangleFace(4, TriangleCount) = ObjectCount
                TriangleFace(5, TriangleCount) = (Index \ 4) + 1

                .ObjectIndex = ObjectCount
                .FaceIndex = (Index \ 4) + 1
                .TextureIndex = GetFileIndex(TextureFileName)
                If TextureFileName <> "" Then
                    If Files(.TextureIndex).Data Is Nothing Then
                        Set Files(.TextureIndex).Data = LoadTexture(TextureFileName)
                        ImageDimensions TextureFileName, Files(.TextureIndex).Size
                    End If
                End If

                VertexDirectX(.TriangleIndex * 3 + 0).X = .Point1.X
                VertexDirectX(.TriangleIndex * 3 + 0).Y = .Point1.Y
                VertexDirectX(.TriangleIndex * 3 + 0).z = .Point1.z
                VertexDirectX(.TriangleIndex * 3 + 0).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 0).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 0).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 0).tu = MeshVerticies(MeshIndicies(Index + 3)).tu
                VertexDirectX(.TriangleIndex * 3 + 0).tv = MeshVerticies(MeshIndicies(Index + 3)).tv

                VertexDirectX(.TriangleIndex * 3 + 1).X = .Point2.X
                VertexDirectX(.TriangleIndex * 3 + 1).Y = .Point2.Y
                VertexDirectX(.TriangleIndex * 3 + 1).z = .Point2.z
                VertexDirectX(.TriangleIndex * 3 + 1).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 1).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 1).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 1).tu = MeshVerticies(MeshIndicies(Index + 4)).tu
                VertexDirectX(.TriangleIndex * 3 + 1).tv = MeshVerticies(MeshIndicies(Index + 4)).tv

                VertexDirectX(.TriangleIndex * 3 + 2).X = .Point3.X
                VertexDirectX(.TriangleIndex * 3 + 2).Y = .Point3.Y
                VertexDirectX(.TriangleIndex * 3 + 2).z = .Point3.z
                VertexDirectX(.TriangleIndex * 3 + 2).NX = .Normal.X
                VertexDirectX(.TriangleIndex * 3 + 2).NY = .Normal.Y
                VertexDirectX(.TriangleIndex * 3 + 2).Nz = .Normal.z
                VertexDirectX(.TriangleIndex * 3 + 2).tu = MeshVerticies(MeshIndicies(Index + 5)).tu
                VertexDirectX(.TriangleIndex * 3 + 2).tv = MeshVerticies(MeshIndicies(Index + 5)).tv

            End With

            TriangleCount = TriangleCount + 1
            Index = Index + 6
            vol.Add m

        Loop

        ObjectCount = ObjectCount + 1

        Set CreateVolumeMesh = vol
    End If
End Function


Public Function PointCache(ByRef p As Point) As Long
    Points.Add p
    PointCache = Points.Count
    Exit Function
    If Points.Count > 0 Then
        Dim i As Long
        For i = 1 To Points.Count
            If Points(i).Serialize = p.Serialize Then
                PointCache = i
                Set p = Points(i)
                Exit Function
            End If
        Next
    End If
    Points.Add p, p.Serialize
    PointCache = Points.Count
End Function
