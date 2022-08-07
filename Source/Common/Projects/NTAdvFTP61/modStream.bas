Attribute VB_Name = "modStream"
Option Explicit
Public Function HexStream(Stream) As String
    Dim cnt As Integer
    Select Case TypeName(Stream)
        Case "Stream"
            If Stream.Length > 0 Then
                For cnt = 0 To Stream.Length - 1
                    HexStream = HexStream & Padding(2, Hex(Stream.Partial(cnt, 1)(1)), "0") & " "
                Next
                HexStream = Stream.Length & ": " & HexStream
            End If
        Case "String"
            If Stream <> "" Then
                For cnt = 1 To Len(Stream)
                    HexStream = HexStream & Padding(2, Hex(Asc(Mid(Stream, cnt, 1))), "0") & " "
                Next
                HexStream = Len(Stream) & ": " & HexStream
            End If
        Case "Byte()"
            For cnt = LBound(Stream) To UBound(Stream)
                HexStream = HexStream & Padding(2, Hex(Stream(cnt)), "0") & " "
            Next
            HexStream = (UBound(Stream) + IIf(LBound(Stream) = 0, 1, 0)) & ": " & HexStream
    End Select
End Function

Public Function Stream(Optional ByRef OfAny) As Stream
    Set Stream = New Stream
    Select Case TypeName(OfAny)
        Case "Byte"
            Stream.Concat modMemory.Convert(Chr(OfAny))
        Case "Integer"
            Stream.Concat modMemory.Convert(HiByte(CInt(OfAny)))
            Stream.Concat modMemory.Convert(LoByte(CInt(OfAny)))
        Case "String"
            Stream.Concat modMemory.Convert(OfAny)
        Case "Byte()"
            Dim tmp() As Byte
            ReDim tmp(LBound(OfAny) To UBound(OfAny)) As Byte
            Dim idx As Long
            For idx = LBound(OfAny) To UBound(OfAny)
                tmp(idx) = OfAny(idx)
            Next
            Stream.Concat tmp
        Case "Stream"
           ' Set Stream = OfAny.Clone
            Dim tmp2 As Stream
            Set tmp2 = OfAny
            Stream.Clone tmp2
            Set tmp2 = Nothing
    End Select
End Function

Public Function ToString(ByRef InStream) As String
    Select Case TypeName(InStream)
        Case "Stream"
            ToString = modMemory.Convert(InStream.Partial)
        Case "Byte()"
            ToString = modMemory.Convert(InStream)
    End Select
End Function


Public Function ToBytes(ByRef OfAny) As Byte()
    Select Case TypeName(OfAny)
        Case "Stream"
            ToBytes = OfAny.Partial
        Case "String"
            ToBytes = modMemory.Convert(OfAny)
        Case "Byte", "Double", "Integer", "Single", "Long"
            ToBytes = modMemory.Convert(Chr(OfAny))
    End Select
End Function

Public Function InStream(ByRef TheStream As Stream, ByRef SearchTerm As Stream, Optional ByVal Offset As Long = 0) As Long
    If (TheStream.Length - Offset) > 0 And SearchTerm.Length > 0 And (TheStream.Length - Offset) > SearchTerm.Length Then
        Dim idx As Long
        Dim ins() As Byte
        Dim ter() As Byte
        ins = TheStream.Partial
        ter = SearchTerm.Partial
        idx = 1 + Offset
        Do
            'InStream = InStream + 1
            Do While ins(idx) = ter(InStream + 1) And (InStream + 1) < SearchTerm.Length
                InStream = InStream + 1
                idx = idx + 1
            Loop
            If ter(InStream + 1) = ins(idx) And InStream + 1 = SearchTerm.Length Then
                InStream = (idx - InStream)
            Else
                idx = ((idx - InStream) + 1)
                InStream = 0
            End If
        Loop Until InStream <> 0 Or idx = ((TheStream.Length - SearchTerm.Length) + 1)
    End If
End Function
