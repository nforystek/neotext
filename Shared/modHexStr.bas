#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modHexStr"
#Const modHexStr = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function HesEncodeData(ByVal d As String) As String
    Dim s As New Strand
    Dim l As Long
    Dim i As Long
    Dim n As String
    l = Len(d)
    If l > 0 Then
        For i = 1 To l
            n = Hex(Asc(Mid(d, i, 1)))
            If Len(n) < 2 Then
                s.Concat "0" & n
            Else
                s.Concat n
            End If
        Next
    End If
    HesEncodeData = s.GetString
End Function

Public Function HesDecodeData(ByVal d As String) As String
    Dim s As New Strand
    Dim l As Long
    Dim i As Long
    l = Len(d)
    If l > 0 Then
        For i = 1 To l Step 2
            s.Concat Chr(Val("&H" & Mid(d, i, 2)))
        Next
    End If
    HesDecodeData = s.GetString
End Function

Public Function IsHexidecimal(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim c2 As Integer
    Dim retVal As Boolean
    retVal = True
    If (Len(Text) Mod 2) = 0 Then
        If Len(Text) > 0 Then
            For cnt = 1 To Len(Text)
                If (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                    retVal = False
                    Exit For
                End If
            Next
        Else
            retVal = False
        End If
    End If
    IsHexidecimal = retVal And (c2 <= 1)
End Function

Public Function IsNexidecimal(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim c2 As Integer
    Dim retVal As Boolean
    retVal = True
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122) Then
                retVal = False
                Exit For
            End If
        Next
    Else
        retVal = False
    End If
    IsNexidecimal = retVal And (c2 <= 1)
End Function



