Attribute VB_Name = "MainModule"
Option Explicit


Public Vertex As New Collection
Public Rotate As New Collection

Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Function RandomPositive(ByVal LowerBound As Double, ByVal UpperBound As Double) As Double
    Randomize
    RandomPositive = Round((((UpperBound - LowerBound) * Rnd) + LowerBound), 0)
End Function

Public Function RandomPoint(ByRef Point As Point, Optional ByVal Min As Double = -1000, Optional ByVal Max As Double = 1000)
    If Point Is Nothing Then Set Point = New Point
    With Point
        .X = CDbl(RandomPositive(Min, (Max / 2)))
        .Y = CDbl(RandomPositive(Min, (Max / 2)))
        .Z = CDbl(RandomPositive(Min, (Max / 2)))
    End With
    Set RandomPoint = Point
End Function
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
