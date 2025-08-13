Attribute VB_Name = "PointRotate"
Option Explicit


'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorAxisAngles(ByRef Point As Point) As Point
    Dim tmp As New Point
    Set VectorAxisAngles = New Point
    With VectorAxisAngles
        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
            Set tmp = Point
            .X = AngleRestrict(AngleOfPlot(MakePoint(tmp.Y, tmp.Z, tmp.X)))
            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
            .Y = AngleRestrict(AngleOfPlot(MakePoint(tmp.Z, tmp.X, tmp.Y)))
            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
            .Z = AngleRestrict(AngleOfPlot(MakePoint(tmp.X, tmp.Y, tmp.Z)))
            Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Z))
            Set tmp = Nothing
        End If
    End With
End Function

'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
    Dim tmp As Point
    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
        Set tmp = VectorRotateX(MakePoint(tmp.Y, tmp.Z, tmp.X), Angles.X)
        Set tmp = VectorRotateY(MakePoint(tmp.Z, tmp.X, tmp.Y), Angles.Y)
    End If
    Set VectorRotateAxis = tmp
    Set tmp = Nothing
End Function


'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As New Point
'    Set VectorRotateAxis = New Point
'    With VectorRotateAxis
'        .Y = (Cos(Angles.X) * Point.Y - Sin(Angles.X) * Point.Z)
'        .Z = (Sin(Angles.X) * Point.Y + Cos(Angles.X) * Point.Z)
'        tmp.X = Point.X
'        tmp.Y = .Y
'        tmp.Z = .Z
'        .X = (Sin(Angles.Y) * tmp.Z + Cos(Angles.Y) * tmp.X)
'        .Z = (Cos(Angles.Y) * tmp.Z - Sin(Angles.Y) * tmp.X)
'        tmp.X = .X
'        .X = (Cos(Angles.Z) * tmp.X - Sin(Angles.Z) * tmp.Y)
'        .Y = (Sin(Angles.Z) * tmp.X + Cos(Angles.Z) * tmp.Y)
'    End With
'End Function




'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################





Public Function VectorRotateX(ByRef Point As Point, ByVal Angle As Single) As Point
    Set VectorRotateX = MakePoint(Point.X, Point.Y, Point.Z)
   ' If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    With VectorRotateX
        .Z = Point.Z * CosPhi - Point.Y * SinPhi
        .Y = Point.Z * SinPhi + Point.Y * CosPhi
        .X = Point.X
    End With
End Function

Public Function VectorRotateY(ByRef Point As Point, ByVal Angle As Single) As Point
    Set VectorRotateY = MakePoint(Point.X, Point.Y, Point.Z)
    'If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(-Angle)
    SinPhi = Sin(-Angle)
    With VectorRotateY
        .X = Point.X * CosPhi - Point.Z * SinPhi
        .Z = Point.X * SinPhi + Point.Z * CosPhi
        .Y = Point.Y
    End With
End Function

Public Function VectorRotateZ(ByRef Point As Point, ByVal Angle As Single) As Point
    Set VectorRotateZ = MakePoint(Point.X, Point.Y, Point.Z)
    'If Round(angle) = 0 Then Exit Function
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    CosPhi = Cos(Angle)
    SinPhi = Sin(Angle)
    With VectorRotateZ
        .X = Point.X * CosPhi - Point.Y * SinPhi
        .Y = Point.X * SinPhi + Point.Y * CosPhi
        .Z = Point.Z
    End With
End Function



'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function
'
'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As New Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = Point
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
'            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
'            .Z = AngleRestrict(AngleOfCoord(MakePoint(tmp.X, tmp.Y, tmp.Z)))
'            Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Z))
'            Set tmp = Nothing
'        End If
'    End With
'End Function

'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As New Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = Point
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
'            Set tmp = Nothing
'        End If
'    End With
'End Function
'
'Public Function VectorAxisAngles(ByRef Point As Point) As Point
'    Dim tmp As Point
'    Set VectorAxisAngles = New Point
'    With VectorAxisAngles
'        If Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0) Then
'            Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'            .X = AngleRestrict(AngleOfCoord(MakePoint(tmp.Y, tmp.Z, tmp.X)))
'            Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.X))
'            .Y = AngleRestrict(AngleOfCoord(MakePoint(tmp.Z, tmp.X, tmp.Y)))
''            Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), AngleInvertRotation(.Y))
''            .Z = AngleRestrict(AngleOfCoord(MakePoint(tmp.X, tmp.Y, tmp.Z)))
'            Set tmp = Nothing
'        End If
'    End With
'End Function




'####################################################################################################
'####################################################################################################
'####################################################################################################
'####################################################################################################

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = MakePoint(Point.X, Point.Y, Point.Z)
'    If (Not (Point.X = 0 And Point.Y = 0 And Point.Z = 0)) And _
'        (Not (Angles.X = 0 And Angles.Y = 0 And Angles.Z = 0)) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateX(MakePoint(tmp.Y, tmp.Z, tmp.X), Angles.X)
'        Set tmp = VectorRotateY(MakePoint(tmp.Z, tmp.X, tmp.Y), Angles.Y)
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'    Dim tmp As Point
'    Set tmp = Point
'    If Abs(Angles.Y) > Abs(Angles.X) And Abs(Angles.Y) > Abs(Angles.Z) And (Angles.Y <> 0) Then
'        Set tmp = VectorRotateY(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Y)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, 0, Angles.Z))
'    ElseIf Abs(Angles.X) > Abs(Angles.Y) And Abs(Angles.X) > Abs(Angles.Z) And (Angles.X <> 0) Then
'        Set tmp = VectorRotateZ(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.X)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(0, Angles.Y, Angles.Z))
'    ElseIf Abs(Angles.Z) > Abs(Angles.Y) And Abs(Angles.Z) > Abs(Angles.X) And (Angles.Z <> 0) Then
'        Set tmp = VectorRotateX(MakePoint(tmp.X, tmp.Y, tmp.Z), Angles.Z)
'        Set tmp = VectorRotateAxis(MakePoint(tmp.X, tmp.Y, tmp.Z), MakePoint(Angles.X, Angles.Y, 0))
'    End If
'    Set VectorRotateAxis = tmp
'    Set tmp = Nothing
'End Function

'Public Function VectorRotateAxis(ByRef Point As Point, ByRef Angles As Point) As Point
'
'    Dim matRoll As D3DMATRIX
'    Dim matYaw As D3DMATRIX
'    Dim matPitch As D3DMATRIX
'    Dim matMat As D3DMATRIX
'
'    D3DXMatrixIdentity matRoll
'    D3DXMatrixIdentity matYaw
'    D3DXMatrixIdentity matPitch
'    D3DXMatrixIdentity matMat
'
'    D3DXMatrixRotationX matPitch, AngleConvertWinToDX3DX(Angles.x)
'    D3DXMatrixMultiply matMat, matPitch, matMat
'
'    D3DXMatrixRotationY matYaw, AngleConvertWinToDX3DY(Angles.y)
'    D3DXMatrixMultiply matMat, matYaw, matMat
'
'    D3DXMatrixRotationZ matRoll, AngleConvertWinToDX3DZ(Angles.Z)
'    D3DXMatrixMultiply matMat, matRoll, matMat
'
'    Dim vout As D3DVECTOR
'    D3DXVec3TransformCoord vout, ToVector(Point), matMat
'
'    Set VectorRotateAxis = New Point
'    With VectorRotateAxis
'        .x = vout.x
'        .y = vout.y
'        .Z = vout.Z
'    End With
'End Function

