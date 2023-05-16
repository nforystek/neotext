Attribute VB_Name = "Module1"

#Const modMain = -1

Option Explicit

'TOP DOWN



Option Private Module




Public Sub Main()

    Dim Line1() As Byte

    Dim Line2() As Byte

    Line1 = Convert("Some text to concat to the stream.")

    Line2 = Convert("This is some other sentences text.")




    Line1 = Convert("Some text to con cat to the stream.")





    Dim s1 As New stream

    Dim s2 As New stream

   ' s1.FileName = "C:\Development\Neotext\Common\Projects\NTNodes10\Test\check.txt"



    s1.concat Line1

    DebugPrintStream "Test", s1



    's.Placeat Line2, 5

    s2.concat s1.Partial(5, 16)



    DebugPrintStream "Test", s2



    s2.concat Line2

'

    Debug.Print s1.poll(Asc("e"), 1)
    




    DebugPrintStream "Test", s1





    s1.Pyramid s2, 5, 16



    DebugPrintStream "Test", s1



  '  Debug.Print s1.pop(3)







    Set s1 = Nothing

    Set s2 = Nothing

   'Test2

End Sub



Public Sub Test2()

    Dim Line1() As Byte

    Dim Line2() As Byte

    Line1 = Convert("Some text to concat to the stream.")

    Line2 = Convert("This is some other sentences text.")

    

    Dim s As New stream

    Dim nxt As Long
    Dim tmp As String
    Dim tmp1 As Long
    

    

    ReadyInput

    Do While Not EndOfInput

    

        Randomize

        nxt = (((Rnd * 10) Mod 10) + 1)

    

        'Debug.Print nxt;

    

            If nxt = 1 Then
                Debug.Print "Reset;";
                'reset

                s.Reset

                DebugPrintStream "Reset", s

            ElseIf nxt = 2 Then
                tmp1 = s.Length
                Debug.Print "Size++;";
                'up size

                s.Length = s.Length + 1
                
                Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(s.Length = tmp1 + 1, "True", "False")

                'DebugPrintStream "Size++", s

            ElseIf nxt = 3 Then
                
                'down size

                If s.Length > 0 Then
                    tmp1 = s.Length
                    Debug.Print "Size--;";
                    s.Length = s.Length - 1
    
                    Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(s.Length = tmp1 - 1, "True", "False")
                   ' DebugPrintStream "Size--", s
                End If

            ElseIf nxt = 4 Then
                Debug.Print "Concat;";
                'concat
                tmp = Convert(s.Partial)
                
                s.concat Line1

                'DebugPrintStream "Concat", s
                Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(Convert(s.Partial) = tmp & Convert(Line1), "True", "False")


            ElseIf nxt = 5 Then
                
                'push

                If s.Length \ 4 >= 4 Then
                    tmp1 = (s.Length \ 4)
                    tmp = Convert(s.Partial)
                    
                    Debug.Print "Push(" & (s.Length \ 4) & ");";
                    
                    s.push (s.Length \ 4)

                    Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(s.Length = Len(tmp) - tmp1, "True", "False")


                    'DebugPrintStream "Push", s

                End If

            ElseIf nxt = 6 Then
                
                'pull

                If s.Length \ 4 >= 4 Then
                    tmp1 = (s.Length \ 4)
                    tmp = Convert(s.Partial)
                    Debug.Print "Pull(" & (s.Length \ 4) & ");";

                    s.Pull (s.Length \ 4)

                    Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(s.Length = Len(tmp) - tmp1, "True", "False")
                    'DebugPrintStream "Pull", s

                End If

            ElseIf nxt = 7 Then
                

                'partial

                If s.Length \ 4 >= 4 Then
                    tmp1 = (s.Length \ 4)
                    tmp = Convert(s.Partial)
                    
                    Debug.Print "Partial(" & tmp1 & "," & tmp1 * 2 & ");";
                    

                    Debug.Print "[" & Convert(s.Partial(tmp1, tmp1 * 2)) & "]:" & IIf(Convert(s.Partial(tmp1, tmp1 * 2)) = Mid(tmp, tmp1 + 1, (tmp1 * 2)), "True", "False")


                    'DebugPrintStream "Partial", s

                End If

            ElseIf nxt = 8 Then
                
                'Placeat

                If s.Length \ 4 >= 4 Then
                    tmp1 = (s.Length \ 4)
                    tmp = Left(Convert(Line2), tmp1 * 2)
                    
                    
                    Debug.Print "Placeat(""" & tmp & """," & tmp1 & ");";

                    s.Placeat Convert(tmp), tmp1, (tmp1 * 2)
 
                    Debug.Print "[" & Convert(s.Partial(tmp1, Len(tmp))) & "]:" & IIf(Convert(s.Partial(tmp1, Len(tmp))) = tmp, "True", "False")
 
                   ' DebugPrintStream "Placeat", s

                End If

                

            ElseIf nxt = 9 Then
                tmp1 = (s.Length \ 4)
                tmp = Convert(s.Partial)
                
                    
                Debug.Print "Prepend;";
                'prepend

                s.Prepend Line2

                Debug.Print "[" & Convert(s.Partial) & "]:" & IIf(Convert(s.Partial) = Convert(Line2) & tmp, "True", "False")
               ' DebugPrintStream "Prepend", s
            
            ElseIf nxt = 10 Then
                

                If s.Length \ 4 >= 4 Then

                    tmp1 = s.Length \ 4
                    
                    Debug.Print "Pinch;";
                    
                    tmp = Convert(s.Partial)
                    tmp = Left(tmp, tmp1) & Right(tmp, (s.Length - (tmp1 + (tmp1 * 2))))

                    
                    s.Pinch tmp1, (tmp1 * 2)
       
                    
                    Debug.Print "Pinch; (" & tmp1 & ", " & (tmp1 * 2) & ")?[" & _
                                    Convert(s.Partial) & "]:" & _
                                    IIf(tmp = Convert(s.Partial), _
                                    "True", "False")
                                    
                    'DebugPrintStream "Pinch", s
                    
                    
                End If
                
                
                
            ElseIf nxt = 11 Then
                Debug.Print "HardReset;";
                'hard reset

                Set s = Nothing

                Set s = New stream

                DebugPrintStream "HardReset", s

 

            End If



        

        DoEvents

        InputLoop

    Loop

    

    

    Set s = Nothing

End Sub



Public Sub Test1()

    Dim Line1() As Byte

    Dim Line2() As Byte

    Line1 = Convert("Some text to concat to the stream.")

    Line2 = Convert("This is some other sentences text.")

    

    Debug.Print "Number pad keys: (or Escape to quit testing)"

    Debug.Print "   Add - New instance of the stream object"

    Debug.Print "   Subtract - Set the instance to nothing"

    Debug.Print "   Up - Increment the side of the stream"

    Debug.Print "   Down - Reduce the size of the stream"

    Debug.Print "   Zero - Reset stream size to zero clearing"

    Debug.Print "   One - Concat a sentence of text to stream"

    Debug.Print "   Two - Push the stream right 1/4th a size"

    Debug.Print "   Three - Pull the stream right 1/4th a size"

    Debug.Print "   Four - Partial middle 2/4th's of stream"

    Debug.Print "   Five - Placeat overwrite sentence at begin"

    Debug.Print "   Six - Prepend to the beginning of stream"

    Debug.Print

    

    

       

    Dim tmp As Long

    Dim Sequence As Long

    

    Dim s As stream

    ReadyInput

    Do

                

        InputLoop

        

        If Pressed(VK_SUBTRACT) Then

            'nothing

            If Not s Is Nothing Then

                Set s = Nothing

                Debug.Print "Set object instance to nothing."

            End If

        ElseIf Pressed(VK_ADD) Then

            'new obj

            If s Is Nothing Then

                Set s = New stream

                Debug.Print "Created new object instance."

            End If

        ElseIf Not s Is Nothing Then

            If Pressed(VK_NUMPAD0) Then

                'reset

                s.Reset

                DebugPrintStream "Reset", s

            ElseIf Pressed(VK_UP) Then

                'up size

                s.Length = s.Length + 1

                DebugPrintStream "Size++", s

            ElseIf Pressed(VK_DOWN) Then

                'down size

                If s.Length > 0 Then s.Length = s.Length - 1

                DebugPrintStream "Size--", s

            ElseIf Pressed(VK_NUMPAD1) Then

                'concat

                s.concat Line1

                DebugPrintStream "Concat", s

            ElseIf Pressed(VK_NUMPAD2) Then

                'push

                If s.Length \ 4 >= 4 Then

                    s.push (s.Length \ 4)

                    DebugPrintStream "Push", s

                End If

            ElseIf Pressed(VK_NUMPAD3) Then

                'pull

                If s.Length \ 4 >= 4 Then

                    s.Pull (s.Length \ 4)

                    DebugPrintStream "Pull", s

                End If

            ElseIf Pressed(VK_NUMPAD4) Then

                'partial

                If s.Length \ 4 >= 4 Then

                    

                    Debug.Print "Partial; (" & (s.Length \ 4) & ", " & ((s.Length \ 4) * 2) & ")?[" & _
                                    Convert(s.Partial((s.Length \ 4), ((s.Length \ 4) * 2))) & "]:" & _
                                    IIf((Convert(s.Partial((s.Length \ 4), ((s.Length \ 4) * 2))) = _
                                    Mid(Convert(s.Partial()), (s.Length \ 4) + 1, ((s.Length \ 4) * 2))), _
                                    "Same result as Mid()", "Results incorrect!")

                Else

                    DebugPrintStream "Partial", s

                End If

            ElseIf Pressed(VK_NUMPAD5) Then

                'Placeat

                

                s.Placeat Line2, (s.Length \ 4)  ', ((s.Length \ 4) * 2)



                DebugPrintStream "Placeat", s

            ElseIf Pressed(VK_NUMPAD6) Then

                'prepend

                s.Prepend Line2

                DebugPrintStream "Prepend", s



            ElseIf Pressed(VK_NUMPAD7) Then

                Sequence = -1

            ElseIf Sequence <> 0 Then

                

                If Pressed(VK_NUMPAD7) Then

                    Sequence = 0

                Else

                    Select Case Sequence

                        Case 1, 5, -1

                            Sequence = 1

                            s.Reset

                            DebugPrintStream "Reset", s

                        Case 2, 3, 6, 7

                            s.concat Line1

                            DebugPrintStream "Concat", s

                        Case 4

                            s.push (s.Length \ 4)

                            DebugPrintStream "Push", s

                        Case 8

                            s.Pull (s.Length \ 4)

                            DebugPrintStream "Pull", s

                            Sequence = 0

                    End Select

                    Sequence = Sequence + 1

                End If

            End If

        ElseIf Pressed(VK_UP) Or Pressed(VK_DOWN) Or _
            Pressed(VK_NUMPAD1) Or Pressed(VK_NUMPAD2) Or _
            Pressed(VK_NUMPAD3) Or Pressed(VK_NUMPAD4) Or _
            Pressed(VK_NUMPAD5) Or Pressed(VK_NUMPAD6) Or _
            Pressed(VK_NUMPAD7) Or Pressed(VK_NUMPAD8) Or _
            Pressed(VK_NUMPAD9) Or Pressed(VK_NUMPAD0) Then

            Debug.Print "No object instance set."

        End If



        DoEvents

        

    Loop Until EndOfInput

    Set s = Nothing

    

End Sub



Private Sub DebugPrintStream(ByVal Func As String, ByRef s As stream, Optional Additional As String = "")

    Dim txt As String

    txt = Convert(s.Partial())

    Debug.Print Func & "; (" & s.Address(-4) & ":" & Len(txt) & "(" & LenB(txt) & "b)" & ")=[" & txt & "]" & Additional

End Sub



Public Function ArraySize(InArray, Optional ByVal InBytes As Boolean = False) As Long

On Error GoTo dimerror



    If UBound(InArray) = -1 Or LBound(InArray) = -1 Then

        ArraySize = 0

    Else

        ArraySize = (UBound(InArray) + -CInt(Not CBool(-LBound(InArray)))) * IIf(InBytes, LenB(InArray(LBound(InArray))), 1)

    End If

    

    Exit Function

dimerror:

    Err.Clear

    ArraySize = 0

End Function





Public Function Convert(Info)

    Dim n As Long

    Dim out() As Byte

    Dim Ret As String

    Select Case TypeName(Info)

        Case "String"

            If Len(Info) > 0 Then

                ReDim out(0 To Len(Info) - 1) As Byte

                For n = 0 To Len(Info) - 1

                    out(n) = Asc(Mid(Info, n + 1, 1))

                Next

            End If

            Convert = out

        Case "Byte()"

            If (ArraySize(Info) > 0) Then

                For n = LBound(Info) To UBound(Info)

                    Ret = Ret & Chr(Info(n))

                Next

            End If

            Convert = Ret

    End Select

End Function





