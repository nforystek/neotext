Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Function PaddingLeft(ByVal txt As String, ByVal nLen As Long, Optional ByVal TheChar As String = "0") As String
    If nLen - Len(Trim(txt)) > 0 Then
        PaddingLeft = String(nLen - Len(Trim(txt)), TheChar) & Trim(txt)
    Else
        PaddingLeft = Trim(txt)
    End If
End Function

Public Function PaddingRight(ByVal txt As String, ByVal nLen As Long, Optional ByVal TheChar As String = "0") As String
    If nLen - Len(Trim(txt)) > 0 Then
        PaddingRight = Trim(txt) & String(nLen - Len(Trim(txt)), TheChar)
    Else
        PaddingRight = Trim(txt)
    End If
End Function

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Double
    Randomize
    RandomPositive = CDbl((UpperBound - LowerBound) * Rnd + LowerBound)
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

Public Function ParseNumerical(ByRef inValues As Variant) As Variant
    'parses the next numerical where seperated by commas, returns it
    'not used in this module, rather is for values tostring elsewhere
    If (InStr(inValues, ",") = 0 And InStr(inValues, " ") > 0) Then
        ParseNumerical = RemoveNextArg(inValues, " ")
    Else
        ParseNumerical = RemoveNextArg(inValues, ",")
    End If
    If InStr(ParseNumerical, vbCr) > 0 Then ParseNumerical = RemoveNextArg(ParseNumerical, vbCr)
    If InStr(ParseNumerical, vbLf) > 0 Then ParseNumerical = RemoveNextArg(ParseNumerical, vbLf)
    If IsNumeric(ParseNumerical) Then
        ParseNumerical = CDbl(ParseNumerical)
    ElseIf Trim(ParseNumerical) = "" Then
        ParseNumerical = 0
    Else
        Err.Raise 8, "ParseNumerical", "The input is not recognized as being numerical."
    End If

End Function

Public Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    Dim pos As Long
    cnt = 0
    pos = InStr(Text, Word)
    Do Until pos = 0
        cnt = cnt + 1
        pos = InStr(pos + Len(Word), Text, Word)
    Loop
    CountWord = cnt
End Function


'Public Sub Swap(ByRef Var1 As Variant, ByRef Var2 As Variant, Optional ByRef Var3 As Variant, Optional ByRef Var4 As Variant, Optional ByRef Var5 As Variant, Optional ByRef Var6 As Variant, Optional ByRef Var7 As Variant, Optional ByRef Var8 As Variant, Optional ByRef Var9 As Variant, Optional ByRef Var0 As Variant)
'    If IsMissing(Var3) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var0
'    ElseIf IsMissing(Var4) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var0
'    ElseIf IsMissing(Var5) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var0
'    ElseIf IsMissing(Var6) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var0
'    ElseIf IsMissing(Var7) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var6
'        Var6 = Var0
'    ElseIf IsMissing(Var8) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var6
'        Var6 = Var7
'        Var7 = Var0
'    ElseIf IsMissing(Var9) Then
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var6
'        Var6 = Var7
'        Var7 = Var8
'        Var8 = Var0
'    Else
'        Var0 = Var1
'        Var1 = Var2
'        Var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var6
'        Var6 = Var7
'        Var7 = Var8
'        Var8 = Var9
'        Var9 = Var0
'    End If
'End Sub
'
'Public Function Toggler(ByVal Value As Long) As Long
'    Toggler = ((-CInt(CBool(Value)) + -1) + -CInt(Not CBool(-Value + -1)))
'End Function
'
'Public Function LargeOf(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
'    If IsMissing(V3) Then
'        If V1 > V2 Then
'            LargeOf = V1
'        ElseIf V2 <> 0 Then
'            LargeOf = V2
'        Else
'            LargeOf = V1
'        End If
'    ElseIf IsMissing(V4) Then
'        If V2 > V3 And V2 > V1 Then
'            LargeOf = V2
'        ElseIf V1 > V3 And V1 > V2 Then
'            LargeOf = V1
'        ElseIf V3 <> 0 Then
'            LargeOf = V3
'        Else
'            LargeOf = LargeOf(V1, V2)
'        End If
'    Else
'        If V2 > V3 And V2 > V1 And V2 > V4 Then
'            LargeOf = V2
'        ElseIf V1 > V3 And V1 > V2 And V1 > V4 Then
'            LargeOf = V1
'        ElseIf V3 > V1 And V3 > V2 And V3 > V4 Then
'            LargeOf = V3
'        ElseIf V4 <> 0 Then
'            LargeOf = V4
'        Else
'            LargeOf = LargeOf(V1, V2, V3)
'        End If
'    End If
'End Function
'
'Public Function LeastOf(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
'
'    If IsMissing(V3) Then
'        If (V2 = 0) And (Not (V2 = 0)) Then
'            LeastOf = V1
'        ElseIf (V1 = 0) And (Not (V1 = 0)) Then
'            LeastOf = V2
'        Else
'            If (V1 > V2) Then
'                LeastOf = V2
'            Else
'                LeastOf = V1
'            End If
'        End If
'    ElseIf IsMissing(V4) Then
'        If Not (V1 = 0 And V2 = 0 And V3 = 0) Then
'            If V3 = 0 Then
'                LeastOf = LeastOf(V1, V2)
'            ElseIf V2 = 0 Then
'                LeastOf = LeastOf(V1, V3)
'            ElseIf V1 = 0 Then
'                LeastOf = LeastOf(V3, V2)
'            Else
'                If V2 < V3 And V2 < V1 Then
'                    LeastOf = V2
'                ElseIf V1 < V3 And V1 < V2 Then
'                    LeastOf = V1
'                Else
'                    LeastOf = V3
'                End If
'            End If
'        End If
'    Else
'        If Not (V1 = 0 And V2 = 0 And V3 = 0 And V4 = 0) Then
'            If V3 = 0 Then
'                LeastOf = LeastOf(V1, V2, V4)
'            ElseIf V3 = 0 Then
'                LeastOf = LeastOf(V1, V2, V4)
'            ElseIf V2 = 0 Then
'                LeastOf = LeastOf(V1, V3, V4)
'            ElseIf V1 = 0 Then
'                LeastOf = LeastOf(V3, V2, V4)
'
'            Else
'                If ((V2 < V3) And (V2 < V1) And (V2 < V4)) Then
'                    LeastOf = V2
'                ElseIf ((V1 < V3) And (V1 < V2) And (V1 < V4)) Then
'                    LeastOf = V1
'                ElseIf ((V3 < V1) And (V3 < V2) And (V3 < V4)) Then
'                    LeastOf = V3
'                Else
'                    LeastOf = V4
'                End If
'            End If
'        End If
'    End If
'End Function
'
'
'Public Function AbsoluteFactor(ByVal N As Double) As Double
'    'returns -1 if the n is below zero, returns 1 if n is above zero, and 0 if n is zero
'    AbsoluteFactor = ((-(AbsoluteValue(N - 1) - N) - (-AbsoluteValue(N + 1) + N)) * 0.5)
'End Function
'
'Public Function AbsoluteValue(ByVal N As Double) As Double
'    'same as abs(), returns a number as positive quantified
'    AbsoluteValue = (-((-(N * -1) * N) ^ (1 / 2) * -1))
'End Function
'
'Public Function AbsoluteWhole(ByVal N As Double) As Double
'    'returns only the digits to the left of a decimal in any numerical
'    AbsoluteWhole = (AbsoluteValue(N) - (AbsoluteValue(N) - (AbsoluteValue(N) Mod (AbsoluteValue(N) + 1)))) * AbsoluteFactor(N)
'    'AbsoluteWhole = (N \ 1) 'is also correct
'End Function
'
'Public Function AbsoluteDecimal(ByVal N As Double) As Double
'    'returns only the digits to the right of a decimal in any numerical
'    AbsoluteDecimal = (AbsoluteValue(N) - AbsoluteValue(AbsoluteWhole(N))) * AbsoluteFactor(N)
'End Function
