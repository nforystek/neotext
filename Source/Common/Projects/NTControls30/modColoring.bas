Attribute VB_Name = "modColoring"
Option Explicit

'first record of line is how many characters and how many records
'offset variable of the first record when below zero then no more
'records for line, and the StartPos is number of charcters to color
'of the offset abs(offset) as exact color, the StartPos below/above
'zero determines right to left or left to right, while current is
'color to fill all portion of not color covered by the StartPos and
'when there are records the offsets are relative added to current
'color and the StartPos is determinae next amound of chars to color
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

Public Function GetIndexRecord(ByRef mat() As RangeType, ByRef charat As Long) As Long
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

Public Sub ModifyMat(ByRef mat() As RangeType, ByVal offSet As Long, ByVal Anylen As Long, ByVal clr As Long)
    Dim startIdx As Long
    Dim stopIdx As Long
    Dim startOffset As Long
    Dim stopOffset As Long
    Dim beforeOffset As Long
    Dim idx As Long

    startOffset = offSet
    stopOffset = offSet + Anylen
    startIdx = GetIndexRecord(mat, startOffset)
    stopIdx = GetIndexRecord(mat, stopOffset)
    
    If (startIdx > 0) Then
        For idx = 1 To startIdx - 1
            beforeOffset = beforeOffset + mat(idx).StartPos
        Next
        beforeOffset = offSet - beforeOffset
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
            mat(startIdx).StartPos = (offSet + Anylen) - mat(startIdx).StartPos
        End If
    End If
    mat(startIdx).StopPos = clr
End Sub



Private Sub PutColor(ByRef Colors() As Long, ByRef CharRange As RangeType, ByVal ColorIndex As Long)
    'simply sets the RangeType to the color given if it is zero
    'or true it does not add a row nor revise outside RangeType
   
  '  FillMemory ByVal VarPtr(Colors(CharRange.StartPos, 0)), CharRange.StopPos - CharRange.StartPos, ColorIndex
    
    Dim cnt As Long
    For cnt = CharRange.StartPos To CharRange.StopPos
        If Colors(cnt, UBound(Colors, 2)) <= 0 Then
            Colors(cnt, UBound(Colors, 2)) = ColorIndex
        End If
    Next
End Sub

Private Sub InitColor(ByRef Colors() As Byte, ByVal BaseColorIndex As Byte, ByVal StartIndex As Long, ByVal CharCount As Long)
    ReDim Colors(StartIndex To CharCount - 1, 0 To 0) As Byte
    FillMemory ByVal VarPtr(Colors(StartIndex, 0)), CharCount, BaseColorIndex
End Sub

Private Sub PushColor(ByRef Colors() As Byte, ByRef CharRange As RangeType, ByVal ColorIndex As Long)
    'this makes a new row of coloring with all true's except the color specified over RangeType
    ReDim Preserve Colors(0 To UBound(Colors, 1), 0 To UBound(Colors, 2) + 1) As Byte
    If CharRange.StartPos > 0 Then
        FillMemory ByVal VarPtr(Colors(0, 0)), CharRange.StartPos, 0
    End If
    FillMemory ByVal VarPtr(Colors(CharRange.StartPos, 0)), CharRange.StopPos - CharRange.StartPos, ColorIndex
    If CharRange.StopPos < UBound(Colors, 1) Then
        FillMemory ByVal VarPtr(Colors(CharRange.StopPos, 0)), UBound(Colors, 1) - (CharRange.StopPos - 1), 0
    End If
End Sub

