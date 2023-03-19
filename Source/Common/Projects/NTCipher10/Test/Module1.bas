Attribute VB_Name = "Module1"
#Const Module1 = -1
Option Explicit
'TOP DOWN
Private cm As New CMultiBits
Private elapsed As Single

Private Sub CheckBreak(ByVal pos As Variant)
    If elapsed = 0 Or Timer - elapsed > 1 Then
        elapsed = Timer
        Open App.path & "\double.txt" For Output As #1
            Print #1, pos
        Close #1
    End If
            
End Sub
Public Sub Test1(ByRef pos As Variant)
    
    Form1.Tag = "Hi/lo byte test"

    Dim num As Integer
    Dim hib As Byte
    Dim lob As Byte
    

        Do Until pos = 32767 Or pos = 0 Or Form1.Tag = "STOP"
            pos = pos + 1
            
            num = pos
            
            'hib = HiByte(num)
            'lob = LoByte(num)
            hib = cm.HiByte(num)
            lob = cm.LoByte(num)
        
            num = cm.MakeWord(lob, hib)
        
'            num = 0
'            LoByte(num) = lob
'            HiByte(num) = hib

            If num <> pos Then
                Form1.Label1.Caption = Form1.Tag & " expected " & pos & " got " & num
                Form1.Tag = "STOP"
                Form1.Visible = True
            Else
                Form1.Label1.Caption = Form1.Tag & " passed for " & pos
            End If
            DoLoop
            CheckBreak pos
        Loop


    Debug.Print "Passed: " & pos
            
        pos = -pos
        
End Sub


Public Sub Test2(ByRef pos As Variant)

    Form1.Tag = "Hi/lo integer test"
    
    Dim num As Long
    Dim hib As Long
    Dim lob As Long

    Do Until pos = 2147483647 Or Form1.Tag = "STOP"
        pos = pos + 1
        
        num = pos
        hib = HiWord(num)
        lob = LoWord(num)
    
        num = 0
        HiWord(num) = hib
        LoWord(num) = lob

        If num <> pos Then
            Form1.Label1.Caption = Form1.Tag & " expected " & pos & " got " & num
            Form1.Tag = "STOP"
            Form1.Visible = True

        Else
            Form1.Label1.Caption = Form1.Tag & " passed for " & pos
        End If
        DoLoop
        CheckBreak pos
    Loop


    Debug.Print "Passed: " & pos
End Sub


Public Sub Test3(ByRef pos As Variant)

    Form1.Tag = "Hi/lo long test"
    
    Dim num As Double
    Dim hib As Long
    Dim lob As Long
    
    Dim sing As Single
    Dim elapse As Single


    Do Until pos = CDec("92233723685477") Or Form1.Tag = "STOP"
        pos = pos + 1
        
        num = pos
        hib = HiLong(num)
        lob = LoLong(num)
    
        num = 0
        HiLong(num) = hib
        LoLong(num) = lob
    
        If num <> pos Then
            Form1.Label1.Caption = Form1.Tag & " expected " & pos & " got " & num
            Form1.Tag = "STOP"
            Form1.Visible = True
        Else
            Form1.Label1.Caption = Form1.Tag & " passed for " & pos
        End If

        DoLoop
        CheckBreak pos
        
    Loop
    

    
End Sub
Public Function Random(Optional ByVal hi As Integer = Num_32767) As Integer
    Random = Round(((((hi - 1) * Rnd) + 1)), 0)
End Function
Public Sub Main()

    Randomize
    
    Dim num As Integer
    num = Random
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    
    b1 = 0
    b2 = 0
    b3 = 0
    b4 = 0
    
    Dim i1 As Integer
    Dim i2 As Integer
    
    Do While True
        DoLoop
        
        If b1 = 255 Then
            If b2 = 255 Then
                If b3 = 255 Then
                    If b4 = 255 Then
                        b4 = 0
                    Else
                        b4 = b4 + 1
                    End If
                    b3 = 0
                Else
                    b3 = b3 + 1
                End If
                b2 = 0
            Else
                b2 = b2 + 1
            End If
            b1 = 0
        Else
            b1 = b1 + 1
        End If
        

    
        i1 = Val("&H" & Padding(2, Hex(b4), "0") & Padding(2, Hex(b1), "0"))
        i2 = Val("&H" & Padding(2, Hex(b3), "0") & Padding(2, Hex(b2), "0"))

        If i1 < 0 Or i2 < 0 Or b1 <> Val("&H" & Right(Padding(4, Hex(i1), "0"), 2)) Or b2 <> Val("&H" & Right(Padding(4, Hex(i2), "0"), 2)) Or _
        b3 <> Val("&H" & Left(Padding(4, Hex(i2), "0"), 2)) Or b4 <> Val("&H" & Left(Padding(4, Hex(i1), "0"), 2)) Then
    
            Stop
            
        End If
        Debug.Print b1; b2; b3; b4; i1; i2
        'Debug.Print Val("&H" & Right(Padding(4, Hex(i1), "0"), 2)); Val("&H" & Right(Padding(4, Hex(i2), "0"), 2));
'        Debug.Print Val("&H" & Left(Padding(4, Hex(i2), "0"), 2)); Val("&H" & Left(Padding(4, Hex(i1), "0"), 2))
    Loop

    End
'    Form1.Show
'
'On Error GoTo extithis:
'
'    Dim pos As Variant
'    Dim tmp As String
'
'    If MsgBox("Start from the beginning?", vbYesNo) = vbNo Then
'
'        Open App.path & "\double.txt" For Input As #1
'            Input #1, tmp
'        Close #1
'        pos = Val(tmp)
'    End If
'
'    If pos <= 32767 Then Test1 pos
'    If pos <= 2147483647 And Form1.Tag <> "STOP" Then Test2 pos
'    If pos <= 92233723685477# And Form1.Tag <> "STOP" Then Test3 pos
'
'    Open App.path & "\double.txt" For Output As #1
'        Print #1, pos
'    Close #1
'
'    If Form1.Tag <> "STOP" Then
'        Form1.Label1.Caption = Form1.Tag & " completed to " & pos
'    ElseIf Not Form1.Visible Then
'        End
'    End If
'
'
'    Exit Sub
'extithis:
'    Open App.path & "\double.txt" For Output As #1
'        Print #1, pos
'    Close #1
'
'     If Form1.Tag <> "STOP" Then
'        Form1.Label1.Caption = Form1.Tag & " errored at " & pos
'    ElseIf Not Form1.Visible Then
'        End
'    End If
    

End Sub
