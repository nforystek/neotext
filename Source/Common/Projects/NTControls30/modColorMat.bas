Attribute VB_Name = "modColorMat"
Option Explicit

Private page() As RangeType
Private clrs() As Long

Public Sub Example()

    If CreateMat(page, ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\Form2.log")) Then

        SetIndexRecord page, 0, 3, 8
        SetIndexRecord page, 2, 4, 8
        SetIndexRecord page, 1, 3, 8
        
        Dim r As RangeType
        r.StartPos = 0
        r.StopPos = 90
        TraverseRange page, r, 0

    End If

End Sub

Public Sub TraverseRange(ByRef mat() As RangeType, ByRef CharRange As RangeType, ByVal CallBack As Long)
   
    Dim startIdx As Long
    Dim stopIdx As Long
    startIdx = GetIndexRecord(mat, CharRange.StartPos)
    stopIdx = GetIndexRecord(mat, CharRange.StopPos)
    If startIdx > -1 And stopIdx >= startIdx Then
        Dim clrIdx As Long
        Dim charCnt As Long
        
        Debug.Print "starting at character " & startIdx
        
        Do
            clrIdx = mat(startIdx).StopPos
            charCnt = mat(startIdx).StartPos
            Do While clrIdx = mat(startIdx).StopPos
                charCnt = charCnt + mat(startIdx).StartPos
                startIdx = startIdx + 1
                If startIdx > stopIdx Then Exit Do
            Loop
            
            Debug.Print "color " & clrIdx & " for " & charCnt & " characters"
            
        Loop Until startIdx >= stopIdx

        'DebugPage GetIndexRecord(mat, CharRange.StartPos), GetIndexRecord(mat, CharRange.StopPos)
    End If

End Sub

Private Sub DebugArry(ByRef ary() As Byte)
    Dim cnt As Long
    For cnt = LBound(ary, 1) To UBound(ary, 1)
        Debug.Print ary(cnt, 0);
    Next
    Debug.Print
End Sub

Private Sub DebugPage(Optional ByVal StartRecord As Long = -1, Optional ByVal StopRecord As Long = -1)
    Dim cnt As Long
    For cnt = IIf(StartRecord = -1, LBound(page), StartRecord) To IIf(StopRecord = -1, UBound(page), StopRecord)
        Debug.Print "Record " & cnt & "; " & page(cnt).StartPos & " " & page(cnt).StopPos
    Next
End Sub

'first record of line is how many characters and how many records
'offset variable of the first record when below zero then no more
'records for line, and the StartPos is number of charcters to color
'of the offset abs(offset) as exact color, the StartPos below/above
'zero determines right to left or left to right, while current is
'color to fill all portion of not color covered by the StartPos and
'when there are records the offsets are relative added to current
'color and the StartPos is determinae next amount of chars to color
Private Function AllWhiteSpace(ByVal txt As String) As Boolean
    'returns true if all of txt is whitespace characters otherwise false if any auditable text exists
    AllWhiteSpace = (Replace(Replace(Replace(Replace(txt, " ", ""), vbTab, ""), vbCr, ""), vbLf, "") = "")
End Function
Private Function NoWhiteSpace(ByVal txt As String) As Boolean
    'returns true if all of txt is auditable characters otherwise false if any whitespace text exists
    NoWhiteSpace = InStr(txt, " ") > 0 Or InStr(txt, vbTab) > 0 Or InStr(txt, vbCr) > 0 Or InStr(txt, vbLf) > 0
End Function

Public Function CreateMat(ByRef mat() As RangeType, ByVal txt As String) As Boolean
    'creates a black and while mat() of the txt returns whether or not mat() is valid created
    If txt <> "" Then
        Dim cnt As Long
        Dim rng As RangeType
        Dim rec As Long
        Dim las As Boolean
        Dim line As String
        Dim lin() As RangeType
        Dim cclr As Long
        Dim lines() As String
        lines = Split(txt & vbCrLf, vbCrLf)
        If Not (UBound(lines) > LBound(lines) And LBound(lines) > -1) Then
            If txt <> "" Then
                Erase lines
                ReDim lines(0 To 0) As String
                lines(0) = txt
            End If
        End If
        If (UBound(lines) >= LBound(lines) And LBound(lines) > -1) Then
            rec = 0
            cclr = IIf(AllWhiteSpace(line), 0, 16777215)
            For cnt = LBound(lines) To UBound(lines)
                line = lines(cnt)
                las = AllWhiteSpace(line) Xor las
                ReDim Preserve lin(0 To rec) As RangeType
                lin(rec).StartPos = Len(lines(cnt))
                If AllWhiteSpace(line) Then
                    lin(rec).StopPos = IIf(las Xor NoWhiteSpace(line), 16777215, 0)
                    lin(rec).StartPos = lin(rec).StartPos
                Else
                    lin(rec).StopPos = cclr * IIf(las Xor NoWhiteSpace(line), IIf(cclr = 0, 1, IIf(cclr < 0, -1, 0)), IIf(cclr = 0, 0, IIf(cclr < 0, 1, -1)))
                    Do Until line = ""
                        las = AllWhiteSpace(Left(line, 1))
                        If (Abs(lin(UBound(lin)).StopPos) <> Abs(CInt(las))) Or (rec = UBound(lin)) Then
                            ReDim Preserve lin(0 To UBound(lin) + 1) As RangeType
                            cclr = cclr * IIf(Not las, IIf(cclr = 0, 1, IIf(cclr < 0, -1, 0)), IIf(cclr = 0, 0, IIf(cclr < 0, 1, -1)))
                            If rec <> UBound(lin) - 1 Then lin(UBound(lin) - 1).StopPos = cclr
                            lin(UBound(lin)).StartPos = ((Abs(lin(UBound(lin)).StartPos) + 1) * IIf(Not (lin(rec).StartPos < 0), 1, -1))
                            lin(UBound(lin)).StopPos = CInt(las)
                        Else
                            lin(UBound(lin)).StartPos = (Abs(lin(UBound(lin)).StartPos) + 1) * IIf(Not (lin(rec).StartPos < 0), 1, -1)
                        End If
                        line = Mid(line, 2)
                    Loop
                    If (UBound(lin) + 1) - (rec + 1) = 2 Then
                        lin(UBound(lin) - 1).StartPos = -lin(UBound(lin) - 1).StartPos * IIf(las, 1, -1)
                        ReDim Preserve lin(LBound(lin) To UBound(lin) - 1) As RangeType
                    Else
                        lin(rec).StopPos = (UBound(lin) + 1) - (rec + 1)
                    End If
                    If rec <> UBound(lin) - 1 Then
                        cclr = cclr * IIf(Not las, IIf(cclr = 0, 1, IIf(cclr < 0, -1, 0)), IIf(cclr = 0, 0, IIf(cclr < 0, 1, -1)))
                        lin(UBound(lin)).StopPos = cclr
                    End If
                End If
                rec = UBound(lin) + 1
                Do While lin(UBound(lin)).StartPos = 0 And lin(UBound(lin)).StopPos = 0 And UBound(lin) > 0
                    ReDim Preserve lin(LBound(lin) To UBound(lin) - 1) As RangeType
                    rec = rec - 1
                Loop
            Next
            CreateMat = True
        End If
        Erase mat
        mat = lin
    Else
        Erase mat
        CreateMat = False
    End If
End Function

Private Function GetIndexRecord(ByRef mat() As RangeType, ByRef charat As Long) As Long
    'returns the index the record of charat number of characters falls in, also changes
    'charat to reflect the how many remaining characters charat truncates the record by
    Do
        If charat > mat(GetIndexRecord).StartPos Then
            charat = charat - mat(GetIndexRecord).StartPos
            GetIndexRecord = GetIndexRecord + mat(GetIndexRecord).StopPos + 1
        Else
            GetIndexRecord = GetIndexRecord + 1
            Do
                charat = charat - mat(GetIndexRecord).StartPos
                GetIndexRecord = GetIndexRecord + 1
            Loop Until charat <= 0
        End If
    Loop Until charat <= 0
    GetIndexRecord = GetIndexRecord - 1
End Function

Private Function SetIndexRecord(ByRef mat() As RangeType, ByVal Offset As Long, ByVal Anylen As Long, ByVal clr As Long)
    'sets a range of characters in mat to clr, starting at offset+1 to anylen amount of characters
    Dim startIdx As Long
    Dim stopIdx As Long
    Dim startOffset As Long
    Dim stopOffset As Long
    Dim beforeOffset As Long
    Dim idx As Long
    
    startOffset = Offset
    stopOffset = Offset + Anylen
    startIdx = GetIndexRecord(mat, startOffset)
    stopIdx = GetIndexRecord(mat, stopOffset)
    
    If (startIdx > 0) Then
        For idx = 1 To startIdx - 1
            beforeOffset = beforeOffset + mat(idx).StartPos
        Next
        beforeOffset = Offset - beforeOffset
    End If
    If (beforeOffset > 0) Then
        ReDim Preserve mat(LBound(mat) To UBound(mat) + 1) As RangeType
        startIdx = startIdx + 1
        For idx = UBound(mat) To startIdx + 1 Step -1
            mat(idx) = mat(idx - 1)
        Next
        mat(startIdx - 1).StartPos = beforeOffset
    End If
    
    If (startIdx = stopIdx) And (startIdx > 0) Then
        mat(startIdx).StartPos = Anylen
        If mat(startIdx).StartPos - (beforeOffset + Anylen) > 0 Then
            startIdx = startIdx + 1
            ReDim Preserve mat(LBound(mat) To UBound(mat) + 1) As RangeType
            For idx = UBound(mat) To startIdx Step -1
                mat(idx) = mat(idx - 1)
            Next
            mat(startIdx).StartPos = mat(startIdx).StartPos - (beforeOffset + Anylen)
        End If
        If stopOffset < 0 Then
            ReDim Preserve mat(LBound(mat) To UBound(mat) + 1) As RangeType
            For idx = UBound(mat) To stopIdx + 1 Step -1
                mat(idx) = mat(idx - 1)
            Next
            mat(stopIdx + 1).StartPos = -stopOffset
        End If
    ElseIf (stopIdx > startIdx) And (startIdx > 0) Then
        mat(startIdx).StartPos = Anylen
        If (stopOffset > 0) Then
            ReDim Preserve mat(LBound(mat) To UBound(mat) + 1) As RangeType
            For idx = UBound(mat) To stopIdx + 1 Step -1
                mat(idx) = mat(idx - 1)
            Next
        End If
        mat(stopIdx + 1).StartPos = -stopOffset
        For idx = startIdx + 1 To UBound(mat) - (stopIdx - startIdx)
            mat(idx) = mat(idx + (stopIdx - startIdx))
        Next
        ReDim Preserve mat(LBound(mat) To UBound(mat) - (stopIdx - startIdx)) As RangeType
    Else
        If stopOffset <> 0 Then
            ReDim Preserve mat(LBound(mat) To UBound(mat) + 1) As RangeType
            For idx = UBound(mat) To stopIdx + 1 Step -1
                mat(idx) = mat(idx - 1)
            Next
            mat(startIdx + 1).StartPos = -stopOffset
        Else
            mat(startIdx).StartPos = (Offset + Anylen) - mat(startIdx).StartPos
        End If
    End If
    mat(startIdx).StopPos = clr
End Function


