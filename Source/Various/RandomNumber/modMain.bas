Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Dim numbersGenerated As String
    Dim randomNumber As Double
    
    Dim cnt As Long
    Do
        randomNumber = Random()
        cnt = cnt + 1
        Debug.Print "Random number: " & randomNumber
        
        If InStr(numbersGenerated, " " & Trim(CStr(randomNumber))) > 0 Then
            Debug.Print "Random number: " & cnt & " repeated"
            Exit Do
        Else
            numbersGenerated = numbersGenerated & " " & Trim(CStr(randomNumber))
        End If
        
    Loop While True
    Debug.Print
End Sub


Public Static Function Random() As Double
    Static toggleSet As Integer
    toggleSet = toggleSet + 1
    Dim ID As String
    Dim UN As String
    Dim CN As String
    Dim HD As String
    Dim SR As String
    Dim MA As String

    ID = (Environ("PROCESSOR_IDENTIFIER"))
    UN = (Environ("USERNAME"))
    CN = (Environ("COMPUTERNAME"))
    HD = (Environ("HOMEDRIVE"))
    SR = SerialNumber(Left(Environ("HOMEDRIVE"), 1))
    MA = MacAddress
    
    Dim ran  As String
    Select Case Abs(toggleSet)
        Case 1
            ran = ID & UN & CN & HD & SR & MA
        Case 2
            ran = UN & CN & HD & SR & MA & ID
        Case 3
            ran = CN & HD & SR & MA & ID & UN
        Case 4
            ran = HD & SR & MA & ID & UN & CN
        Case 5
            ran = SR & MA & ID & UN & CN & HD
        Case 6
            ran = MA & ID & UN & CN & HD & SR
    End Select
    ran = Replace(ran, " ", "")
    
    If Abs(toggleSet) = 6 Then toggleSet = -toggleSet
    
    Dim out As String
    out = Replace(CStr(Timer), ".", "")

    Dim Direction As Boolean
    Dim loc As Byte
    loc = 1
    Do While ran <> ""
        Modify out, loc, Left(ran, 1), Direction
        ran = Mid(ran, 2)
        loc = loc + 1
        If loc > Len(out) Then loc = 1
    Loop
    
    Do While Len(Trim(out)) < 7
        out = "0" & out
    Loop
    If Len(Trim(CStr(CDbl("0." & out)))) < 9 Then
        out = String(9 - Len(Trim(CStr(CDbl("0." & out)))), "0") & out
    End If
    
    Random = CDbl("0." & Trim(CStr(out)))

End Function

Private Function Modify(ByRef Data As String, ByVal CharNumber As Byte, ByVal Modby As String, ByRef Direction As Boolean)
    Dim item As String
    Dim l As String
    Dim r As String
    Dim change As Byte
    If CharNumber > 1 Then
        l = Left(Data, CharNumber - 1)
    End If
    If CharNumber < Len(Data) Then
        r = Mid(Data, CharNumber + 1)
    End If
    item = CByte(Mid(Data, CharNumber, 1))
    change = Asc(Modby)
    Do While change > 0
        If item = 0 And (Not Direction) Then
            Direction = Not Direction
        ElseIf item = 9 And Direction Then
            Direction = Not Direction
        End If
        If Direction And item < 9 Then
            item = item + 1
            change = change - 1
        ElseIf Not Direction And item > 0 Then
            item = item - 1
            change = change - 1
        Else
            Direction = Not Direction
        End If
        If change = 0 Then Direction = Not Direction
    Loop
    Data = l & Trim(CStr(item)) & r
End Function


