Attribute VB_Name = "modCollision"
Option Explicit


Public Function PointBehind3DTriangle(ByVal pointX As Single, ByVal pointY As Single, ByVal pointZ As Single, _
                                ByVal length1 As Single, ByVal length2 As Single, ByVal length3 As Single, _
                                ByVal normalX As Single, ByVal normalY As Single, ByVal normalZ As Single) As Boolean
    PointBehind3DTriangle = ((pointZ * length3 + length2 * pointY + length1 * pointX) - (length3 * normalZ + length1 * normalX + length2 * normalY) <= 0)
End Function


Public Function PointInside2DPolygon(ByVal pX As Single, ByVal pY As Single, polyx() As Single, polyy() As Single, ByVal polyn As Long) As Long

    If (polyn > 2) Then
        Dim ref As Single
        Dim ret As Single
        Dim result As Long

        ref = ((pX - polyx(0)) * (polyy(1) - polyy(0)) - (pY - polyy(0)) * (polyx(1) - polyx(0)))
        ret = ref
        Dim i As Long
        For i = 1 To polyn
            ref = ((pX - polyx(i)) * (polyy(i) - polyy(i - 1)) - (pY - polyy(i)) * (polyx(i) - polyx(i - 1)))
            If ((ret >= 0) And (ref < 0) And (result = 0)) Then
                result = i
            End If
            ret = ref
        Next
        If ((result = 0) Or (result > polyn)) Then
            PointInside2DPolygon = 1 '//todo: this is suppose to return a decimal percent
                                  '                      //of the total polygon points where in is found inside
        Else
            PointInside2DPolygon = 0
        End If
    End If

End Function

