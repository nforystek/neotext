Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN

Option Private Module
Public Sub Main()
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
    
    Dim s As Stream
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
                Set s = New Stream
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
                s.Concat Line1
                DebugPrintStream "Concat", s
            ElseIf Pressed(VK_NUMPAD2) Then
                'push
                If s.Length \ 4 >= 4 Then
                    s.Push (s.Length \ 4)
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
                
                s.Placeat Line2, (s.Length \ 4) ', ((s.Length \ 4) * 2)

                DebugPrintStream "Partial", s
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
                            s.Concat Line1
                            DebugPrintStream "Concat", s
                        Case 4
                            s.Push (s.Length \ 4)
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

Private Sub DebugPrintStream(ByVal Func As String, ByRef s As Stream, Optional Additional As String = "")
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
