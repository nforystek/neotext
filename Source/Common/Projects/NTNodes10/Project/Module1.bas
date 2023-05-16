Attribute VB_Name = "Module1"

#Const Module1 = -1

Option Explicit

'TOP DOWN



Public test As NTNodes10.Strands



Public blanks() As Byte

Public blank() As Byte



Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Public Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (ary() As Any) As Long

 

Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)



Public Function Char(Optional ByVal Value As Byte = 10) As Byte()

    Dim tmp(0 To 0) As Byte

    tmp(0) = Value

    Char = tmp

End Function



Public Sub Main()



'    Dim s As New Stream

'

'    Dim b(1 To 4) As Byte

'    b(1) = Asc("1")

'    b(2) = Asc("2")

'    b(3) = Asc("3")

'    b(3) = Asc("4")

'

'    s.Concat b

'

'    s.Push 1

'    Debug.Print Convert(s.Partial())

'    Set s = Nothing

'

'    End



    ResetObject



    blanks = Convert(ReadFile(App.path & "\License.txt"))

    blank = Convert("End of Agreement")



    Form1.Show

     

End Sub



Public Sub ResetObject()



    If Not test Is Nothing Then

        Set test = Nothing

    End If

    

    Set test = New NTNodes10.Strands

    

    If Form1.Option1.Value Then

        If Form1.Option6.Value Then

            test.Behavior = scopeNormal

        ElseIf Form1.Option4.Value Then

            test.Filename = "(Temp)"

            test.Behavior = scopeNormal

        ElseIf Form1.Option5.Value Then

            test.Filename = App.path & "\Temp.txt"

            test.Behavior = scopeNormal

        End If

    ElseIf Form1.Option2.Value Then

        If Form1.Option6.Value Then

            test.Behavior = ScopeLocale

        ElseIf Form1.Option4.Value Then

            test.Filename = "(Temp)"

            test.Behavior = ScopeLocale

        ElseIf Form1.Option5.Value Then

            test.Filename = App.path & "\Temp.txt"

            test.Behavior = ScopeLocale

        End If

    ElseIf Form1.Option3.Value Then

        If Form1.Option6.Value Then

            test.Behavior = ScopeGlobal

        ElseIf Form1.Option4.Value Then

            test.Filename = "(Temp)"

            test.Behavior = ScopeGlobal

        ElseIf Form1.Option5.Value Then

            test.Filename = App.path & "\Temp.txt"

            test.Behavior = ScopeGlobal

        End If

    End If

    

End Sub



Public Function RandomTest() As Boolean



    Randomize

    RandomTest = PreformAction(RandomPositive(1, 20) - 1)



End Function



Public Function RandomPositive(Lowerbound As Long, Upperbound As Long) As Single

    RandomPositive = CSng((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)

End Function



Public Function PreformAction(ByRef Action As Long) As Boolean

    Dim Full As Boolean

    Dim None As Boolean

    Dim AtFirst As Boolean

    Dim AtFinal As Boolean

    

    Form1.Label1(5).Tag = False

    

    If test.Current > test.count Then test.Current = test.count

    None = (test.count = 0)

    Full = (test.count > 0)

    AtFirst = (test.Current = 1)

    AtFinal = (test.Current = test.count)



    If test.count > 0 Then

        If test.Current > test.count Then test.Current = test.count

        If test.Current < 1 Then test.Current = 1

    Else

        test.Current = 0

    End If
    


    Select Case Action

        Case 0 'Add

            DebugPrint "test.Add " & Convert(blank)

            test.Add blank

            PreformAction = True

        Case 1 'insert

            DebugPrint "test.Insert " & Convert(blank)

            test.Insert blank

            PreformAction = True

        Case 2 'append

            DebugPrint "test.Append " & Convert(blank)

            test.Append blank

            PreformAction = True

        Case 3 'pop

            If Full Then

                DebugPrint "test.pop " & Convert(test.Pop)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 4 'delete

            If Full Then

                DebugPrint "test.Delete " & Convert(test.Delete)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 5 'remove

            If Full Then

                DebugPrint "test.Remove " & Convert(test.Remove)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 6 'forward

            If Not (AtFinal Or (test.count < 2)) Then

                DebugPrint "test.forward " & Convert(test.Forward)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 7 'previous

            If Not (AtFirst Or (test.count < 2)) Then

                DebugPrint "test.previous " & Convert(test.Previous)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 8 'Set Value

            DebugPrint "test.Value = " & Convert(blanks)

            test.Value = blanks

            PreformAction = True

        Case 9 'Get Value

            If Not None Then

                DebugPrint "test.Value " & Convert(test.Value)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 10 'Get First

            If Not None Then

                DebugPrint "test.First " & Convert(test.First)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 11 'Set First

            If Not None Then

                DebugPrint "test.First = " & Convert(blank)

                test.First = blank

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 14 'Get Item

            If Not None Then

                DebugPrint "test.Item " & Convert(test.Item)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 15 'set Item

            If Not None Then

                DebugPrint "test.Item = " & Convert(blank)

                test.Item = blank

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 17 'Get Item(#)

            If Not None Then

                DebugPrint "test.Item(test.Current) " & Convert(test.Item(test.Current))

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 16 'set Item(#)

            If Not None Then

                DebugPrint "test.Item(test.Current) = " & Convert(blank)

                test.Item(test.Current) = blank

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 12 'Get Last

            If Not None Then

                DebugPrint "test.Last " & Convert(test.Last)

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 13 'Set Last

            If Not None Then

                DebugPrint "test.Last = " & Convert(blank)

                test.Last = blank

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 19 '<

            If Not None And Not AtFirst Then

                DebugPrint "test.Current = " & (test.Current - 1)

                test.Current = test.Current - 1

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

        Case 18 '>

            If Not None And Not AtFinal Then

                DebugPrint "test.Current = " & (test.Current + 1)

                test.Current = test.Current + 1

                PreformAction = True

            End If

            Form1.Label1(5).Tag = True

    End Select

    

    Form1.Label1(5).Caption = test.Current

    Form1.Label1(6).Caption = test.count

    

    'If Not test.CRLFCheck Then Err.Raise 8, "Failed CRC Checksum."

End Function



Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal NoTrim As Boolean = False) As String

    If NoTrim Then

        If InStr(TheParams, TheSeperator) > 0 Then

            RemoveNextArg = Left(TheParams, InStr(TheParams, TheSeperator) - 1)

            TheParams = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator))

        Else

            RemoveNextArg = TheParams

            TheParams = ""

        End If

    Else

        If InStr(TheParams, TheSeperator) > 0 Then

            RemoveNextArg = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))

            TheParams = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator)))

        Else

            RemoveNextArg = Trim(TheParams)

            TheParams = ""

        End If

    End If

End Function

Public Sub DebugPrint(ByVal Text As String)



    Open App.path & "\debug.log" For Output As #1

    Print #1, Text

    Close #1

    Debug.Print Text



    Text = Form1.Text1.Text & vbCrLf & Text

    Do While Form1.TextHeight(Text) > Form1.Text1.Height

        RemoveNextArg Text, vbCrLf

    Loop



    Form1.Text1.Text = Text

    Form1.Text1.SelStart = Len(Text)



    DebugTest

End Sub

Public Function InValidLine(ByVal Line As String) As String

    

    Dim str As String

    str = "Open Source License" & vbCrLf

    str = str & "       1] Granting of Licenses." & vbCrLf

    str = str & "       2] Reservation of Copyright." & vbCrLf

    str = str & "       3] Limitation of Liabilities." & vbCrLf

    str = str & "       4] Distribution of Software." & vbCrLf

    str = str & "       5] Restriction of Reserves." & vbCrLf

    str = str & "       6] Ownership of Property." & vbCrLf

    str = str & "       7] General New Underwrite." & vbCrLf

    str = str & "End of Agreement" & vbCrLf

    Dim lin As String

    

    Do

        lin = RemoveNextArg(str, vbCrLf)

        If Line = lin Then

            InValidLine = Line

            Exit Function

        End If

    Loop Until str = ""

    

    If Line = "" Then

        InValidLine = Line

        Exit Function

    End If

    

  '  Debug.Print Replace(Line, vbLf, "\n")

  '  Debug.Print Asc(Right(Line, 1))

 '   Stop

InValidLine = Line

End Function

Public Sub DebugTest()

    Static printText As String



    Dim str As String

    If test.count > 0 Then

        Dim cnt As Long

            str = Convert(test.Value)

            

        'For cnt = 1 To test.Count

        '    str = str & InValidLine(Convert(test.Item(cnt))) & vbCrLf

        'Next

    Else

        printText = str

    End If



    If str <> printText Then

        printText = str

        Form1.Picture1.Cls

        Form1.Picture1.Print printText

        Dim nfile As String



        nfile = test.Filename



        If nfile <> "(None)" And GetFileExt(nfile, True, True) <> "tmp" Then

'            Set test = Nothing

            'test.Reset

            Form1.Text2.Text = Replace(ReadFile(nfile), vbLf, vbCrLf)

        ResetObject

            

            

        End If

    End If

    

    

 '   DoEvents

   ' DoTasks

   DoPower

  '

    DoLoop

   ' Sleep 1



End Sub

