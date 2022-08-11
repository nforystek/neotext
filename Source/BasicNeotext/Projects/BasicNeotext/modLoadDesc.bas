Attribute VB_Name = "modLoadDesc"
Option Explicit

'
Private Const BEGINS = vbLf & vbCr & """"
Private Const MIDDLE1 = """ = noitpircseD_BV."
Private Const MIDDLE2 = " etubirttA" & vbLf & vbCr
Private Const ENDING = vbLf & vbCr & "_ '"
'



Private Function TrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    TrimStrip = LTrimStrip(RTrimStrip(TheStr, TheChar), TheChar)
End Function
Private Function LTrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    Do While Left(TheStr, Len(TheChar)) = TheChar
        TheStr = Mid(TheStr, Len(TheChar) + 1)
    Loop
    LTrimStrip = TheStr
End Function
Private Function RTrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    Do While Right(TheStr, Len(TheChar)) = TheChar
        TheStr = Left(TheStr, Len(TheStr) - Len(TheChar))
    Loop
    RTrimStrip = TheStr
End Function

Private Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) - 1)
    Else
        NextArg = TheParams
    End If
End Function

Private Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
    Else
        RemoveArg = ""
    End If
End Function

Private Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) - 1)
        TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) + Len(TheSeperator))
    Else
        RemoveNextArg = TheParams
        TheParams = ""
    End If
End Function
Private Function NextQuotedArg(ByVal TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    NextQuotedArg = RemoveQuotedArg(TheParams, BeginQuote, EndQuote)
End Function
Private Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    Dim retVal As String
    Dim x As Long
    x = InStr(1, TheParams, BeginQuote, vbTextCompare)
    If (x > 0) And (x < Len(TheParams)) Then
        If (InStr(x + Len(BeginQuote), TheParams, EndQuote, vbTextCompare) > 0) Then
            If True Or (EndQuote = BeginQuote) Then
                retVal = Mid(TheParams, x + Len(BeginQuote))
                TheParams = Left(TheParams, x - 1) & Mid(retVal, InStr(1, retVal, EndQuote, vbTextCompare) + Len(EndQuote))
                retVal = Left(retVal, InStr(1, retVal, EndQuote, vbTextCompare) - 1)
            Else
                Dim l As Long
                Dim Y As Long
                l = 1
                Y = x
                Do Until l = 0
                    If (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare) > 0) And (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare) < InStr(Y + Len(BeginQuote), TheParams, EndQuote, vbTextCompare)) Then
                        l = l + 1
                        Y = InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare)
                    ElseIf (InStr(Y + Len(BeginQuote), TheParams, EndQuote, vbTextCompare) > 0) Then
                        l = l - 1
                        Y = InStr(Y + Len(EndQuote), TheParams, EndQuote, vbTextCompare)
                    Else
                        Y = Len(TheParams)
                        l = 0
                    End If
                Loop
                retVal = Mid(TheParams, x + Len(BeginQuote))
                TheParams = Left(TheParams, x - 1) & Mid(retVal, (Y - x) + Len(EndQuote))
                retVal = Left(retVal, (Y - x) - 1)
            End If
        End If
    End If
    RemoveQuotedArg = retVal
End Function

Private Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , vbTextCompare))
    If cnt > 0 Then CountWord = cnt
End Function
