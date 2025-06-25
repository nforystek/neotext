Attribute VB_Name = "Module1"
Option Explicit

'Dim dd As Long
'Dim mm As Long
'Dim yy As Long
'Dim h As Long
'Dim m As Long
'Dim s As Long

Public Function UDT(ByVal DateTime As String) As String
    If IsDate(DateTime) Or InStr(DateTime, ":") > 0 Or InStr(DateTime, "/") > 0 Then
        Dim ret As Currency
        On Error Resume Next
        UDT = Year(DateTime)
        ret = ((UDT \ 4) * 28)
        ret = ((UDT - (UDT \ 4)) * 29) + ((UDT - (UDT \ 4)) * (31 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + 31))
        UDT = ret * CLng(24) * CLng(60) * CLng(60)
        If Month(DateTime) >= 2 Then
            If Year(DateTime) Mod 4 = 0 Then
                ret = 31 + 28
            Else
                ret = 31 + 29
            End If
        Else
            ret = 31
        End If
        Select Case Month(DateTime)
            Case 12
                ret = ret + (31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + 31)
            Case 11
                ret = ret + (31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30)
            Case 10
                ret = ret + (31 + 30 + 31 + 30 + 31 + 31 + 30 + 31)
            Case 9
                ret = ret + (31 + 30 + 31 + 30 + 31 + 31 + 30)
            Case 8
                ret = ret + (31 + 30 + 31 + 30 + 31 + 31)
            Case 7
                ret = ret + (31 + 30 + 31 + 30 + 31)
            Case 6
                ret = ret + (31 + 30 + 31 + 30)
            Case 5
                ret = ret + (31 + 30 + 31)
            Case 4
                ret = ret + (31 + 30)
            Case 3
                ret = ret + (31)
        End Select
        UDT = UDT + (ret * CLng(24) * CLng(60) * CLng(60))
        UDT = UDT + (Day(DateTime) * CLng(24) * CLng(60) * CLng(60))
        UDT = UDT + (Hour(DateTime) * 60 * 60)
        UDT = UDT + (Minute(DateTime) * 60)
        UDT = UDT + Second(DateTime)
        If Len(UDT) < 14 Then
            UDT = String(14 - Len(UDT), "0") & UDT
        End If
    Else
        ret = CCur(DateTime)
        

'    Dim datTime As Date
'    Dim dblTime As Double
'    Dim lngTime As Long
'
'    dblTime = ret 'Asc(Left$(ret, 1)) * 256 ^ 3 + Asc(Mid$(ret, 2, 1)) * 256 ^ 2 + Asc(Mid$(ret, 1)) * 256 ^ 1 + Asc(Right$(ret, 1))
'    ret = CCur(dblTime) - 2840140800#
'    datTime = DateAdd("s", CDbl(ret), #1/1/1990#)
    
        Dim dd As Long
        Dim mm As Long
        Dim yy As Long
        Dim h As Long
        Dim m As Long
        Dim s As Long
        yy = 0
        mm = 0
        dd = 0


        Dim inc As Long


        Do Until ret = -1

            If mm = 12 Then
                mm = 0
                yy = yy + 1
            End If
            If mm = 0 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 1 Then
                If yy Mod 4 = 0 Then
                    If ret >= (28 * CLng(24) * CLng(60) * CLng(60)) Then
                        mm = mm + 1
                        ret = ret - (28 * CLng(24) * CLng(60) * CLng(60))
                    Else
                         dd = -1
                    End If
                Else
                    If ret >= (29 * CLng(24) * CLng(60) * CLng(60)) Then
                        mm = mm + 1
                        ret = ret - (29 * CLng(24) * CLng(60) * CLng(60))
                    Else
                         dd = -1
                    End If
                End If
            ElseIf mm = 2 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                     dd = -1
                End If
            ElseIf mm = 3 Then
                If ret >= (30 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (30 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 4 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 5 Then
                If ret >= (30 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (30 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 6 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                     dd = -1
                End If
            ElseIf mm = 7 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 8 Then
                If ret >= (30 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (30 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 9 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            ElseIf mm = 10 Then
                If ret >= (30 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (30 * CLng(24) * CLng(60) * CLng(60))
                Else
                     dd = -1
                End If
            ElseIf mm = 11 Then
                If ret >= (31 * CLng(24) * CLng(60) * CLng(60)) Then
                    mm = mm + 1
                    ret = ret - (31 * CLng(24) * CLng(60) * CLng(60))
                Else
                    dd = -1
                End If
            End If

            If dd = -1 Then
                dd = ret \ (CLng(24) * CLng(60) * CLng(60))
                ret = ret - (dd * CLng(24) * CLng(60) * CLng(60))
                h = ret \ (60 * 60)
                ret = ret - (h * 60 * 60)
                m = ret \ 60
                ret = ret - (m * 60)
                s = ret
                ret = -1
            End If
        Loop

        UDT = h & ":" & m & ":" & s & " " & dd & "/" & mm & "/" & yy

    End If
End Function

'Private Sub Increment(ByVal interval As Integer)
'    Select Case interval
'        Case 1
'            s = s + 1
'        Case 2
'            m = m + 1
'        Case 3
'            h = h + 1
'        Case 4
'            dd = dd + 1
'        Case 5
'            mm = mm + 1
'        Case 6
'            yy = yy + 1
'    End Select
'
'    If s > 59 Then
'        m = m + 1
'        s = 0
'    End If
'    If m > 59 Then
'        h = h + 1
'        m = 0
'        s = 0
'    End If
'    If h > 23 Then
'        dd = dd + 1
'        h = 0
'        m = 0
'        s = 0
'    End If
'
'    Select Case mm
'        Case 1
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 2
'            If dd > IIf(yy Mod 4 = 0, 28, 29) Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 3
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 4
'            If dd > 30 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 5
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 6
'            If dd > 30 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 7
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 8
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 9
'            If dd > 30 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 10
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 11
'            If dd > 30 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'        Case 12
'            If dd > 31 Then
'                dd = 1
'                mm = mm + 1
'                h = 0
'                m = 0
'                s = 0
'            End If
'    End Select
'    If mm > 12 Then
'        mm = 1
'        dd = 1
'        yy = yy + 1
'        h = 0
'        m = 0
'        s = 0
'    End If
'
'End Sub
'
'Private Sub Decrement(ByVal interval As Integer)
'
'    Select Case interval
'        Case 1
'            s = s - 1
'        Case 2
'            m = m - 1
'        Case 3
'            h = h - 1
'        Case 4
'            dd = dd - 1
'        Case 5
'            mm = mm - 1
'        Case 6
'            yy = yy - 1
'    End Select
'
'    If s < 0 Then
'        m = m - 1
'        s = 59
'    End If
'    If m < 0 Then
'        h = h - 1
'        m = 59
'        s = 59
'    End If
'    If h < 0 Then
'        dd = dd - 1
'        h = 23
'        m = 59
'        s = 59
'    End If
'
'    Select Case mm
'        Case 1
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 2
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 3
'            If dd < 1 Then
'                dd = IIf(yy Mod 4 = 0, 28, 29)
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 4
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 5
'            If dd < 1 Then
'                dd = 30
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 6
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 7
'            If dd < 1 Then
'                dd = 30
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 8
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 9
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 10
'            If dd < 1 Then
'                dd = 30
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 11
'            If dd < 1 Then
'                dd = 31
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'        Case 12
'            If dd < 1 Then
'                dd = 30
'                mm = mm - 1
'                h = 23
'                m = 59
'                s = 59
'            End If
'    End Select
'    If mm < 1 Then
'        mm = 12
'        dd = 31
'        yy = yy - 1
'        h = 23
'        m = 59
'        s = 59
'    End If
'
'End Sub


Public Sub Main()

    Debug.Print Now & " " & UDT(UDT(Now))

    Dim entry As String
    
    entry = "4/14/2025 4:24:01 PM"
    Debug.Print entry & " " & UDT(UDT(entry))
    
    entry = "5/23/1940 8:55:01 PM"
    Debug.Print entry & " " & UDT(UDT(entry))
    
'    entry = "4/14/2025 4:24:01 PM"
'    Debug.Print entry & " " & UDT(UDT(entry))
'
'    entry = "4/14/2025 4:24:01 PM"
'    Debug.Print entry & " " & UDT(UDT(entry))
    
    Exit Sub
    
'    'time size 00000000000000
'    'disc size 4608002 (one partition 1st disc)
'    'partition amounts 329143 (for 2nd disc)
'
''    Dim cnt As Double
''    Dim fso As Scripting.FileSystemObject
''    Dim txt As Scripting.TextStream
''    Dim fle As Scripting.File
''    Set fso = New Scripting.FileSystemObject
'
'
'
''    dd = 12
''    mm = 31
''    yy = 1999
''    h = 24
''    m = 60
''    s = 60
''    Increment
''    If fso.FileExists(App.Path & "\Disc1\" & String(14, "0")) Then
''        Set fle = fso.GetFile(App.Path & "\Disc1\" & String(14, "0"))
''        Set txt = fso.OpenTextFile(App.Path & "\Disc1\" & String(14, "0"), ForReading, False) 'Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss"), ForWriting, True)
''        cnt = fle.Size - 14
''        Do
''            txt.Skip LongBound
''            cnt = cnt - LongBound
''        Loop While cnt >= LongBound
''        If cnt > 0 Then txt.Skip cnt
''        If cnt > 0 Then
''            dd = CLng(txt.Read(2))
''            mm = CLng(txt.Read(2))
''            yy = CLng(txt.Read(2))
''            h = CLng(txt.Read(2))
''            m = CLng(txt.Read(2))
''            s = CLng(txt.Read(2))
''        End If
''        Set fle = Nothing
''    Else
''        Set txt = fso.OpenTextFile(App.Path & "\Disc1\" & String(14, "0"), ForWriting, True) 'Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss"), ForWriting, True)
''    End If
''    txt.Close
''    Set txt = Nothing
'
'
'
'
''    cnt = 14
''    Set txt = fso.OpenTextFile(App.Path & "\Disc1\" & String(14, "0"), ForWriting, True)
''    Do
''        txt.Write Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss")
''        Increment
''        cnt = cnt + 14
''    Loop Until cnt >= CDbl(4396483) * CDbl(1024)
''    txt.Close
''    Set txt = Nothing
'
'
'
'    dd = 12
'    mm = 31
'    yy = 1999
'    h = 24
'    m = 60
'    s = 60
'
'
'    dd = Day(Now)
'    mm = Month(Now)
'    yy = Year(Now)
'    h = Hour(Now)
'    m = Minute(Now)
'    s = Second(Now)
'
'
'
'    Do
'        Decrement
'        DoEvents
'        'Debug.Print Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss")
'        Debug.Print Pad(dd) & "-" & Pad(mm) & "-" & yy & "_" & Pad(h) & "-" & Pad(m) & "-" & Pad(s)
'    Loop Until False
'
'
''    Do
''        Set txt = fso.OpenTextFile(App.Path & "\Disc2\" & Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss"), ForWriting, True)
''        txt.Write String(14, "0")
''        txt.Close
''        Set txt = Nothing
''        Increment
''        cnt = cnt + 14
''    Loop Until cnt >= CDbl(4396483) * CDbl(1024)
    
End Sub

Private Function Pad(ByVal Number As Long) As String
    If Len(CStr(Number)) < 2 Then
        Pad = String(2 - Len(CStr(Number)), "0") & Number
    Else
        Pad = Number
    End If
End Function
