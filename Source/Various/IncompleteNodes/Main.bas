Attribute VB_Name = "Module"
Option Explicit
Option Compare Binary


#If VBIDE = -1 Then
Public sec As Object
#End If

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Addrs As Declares.Addresses


Private Function NewAddr(ByRef Addrs As Addresses, Optional ByRef Addr As Long, Optional ByVal Silent As Boolean = False) As Long
    If Addr < 1 Then
#If VBIDE = -1 Then
        Addr = sec.Alloc(0&, 8)
#Else
        Addr = LocalAlloc(0&, 8)
#End If
        If Addr <= 0 Then
            Err.Raise 17, "AddTo", "Allocate failure."
        Else
            If Not Silent Then DebugLate = DebugLate & "+" & Padding(Addr, 11)
        End If
    ElseIf Addr > 0 And Addr Mod 4 = 0 Then
        If Not Silent Then DebugLate = DebugLate & "=" & Padding(Addr, 11)
    Else
        Err.Raise 17, "AddTo", "Invalid address."
    End If
    NewAddr = Addr
End Function

Private Function DelAddr(ByRef Addrs As Addresses, ByVal Addr As Long, Optional ByVal Silent As Boolean = False) As Long
#If VBIDE = -1 Then
    If sec.Free(Abs(CLng(Addr))) Then
        Err.Raise 17, "DelOf", "Deallocate failure."
    Else
        If Not Silent Then DebugLate = DebugLate & "-" & Padding(Abs(CLng(Addr)), 11)
        DelAddr = Abs(CLng(Addr))
    End If
#Else
    If LocalFree(Abs(CLng(Addr))) Then
        Err.Raise 17, "DelOf", "Deallocate failure."
    Else
        If Not Silent Then DebugLate = DebugLate & "-" & Padding(Abs(CLng(Addr)), 11)
        DelAddr = Abs(CLng(Addr))
    End If
#End If
End Function


Public Function IsEol(ByRef Addrs As Addresses)
    With Addrs
        IsEol = Not (((((Not MPar(Addrs)) And CPar(Addrs)) Xor (MPar(Addrs) And (Not CPar(Addrs)))) And IsInv(Addrs)) Xor _
                ((((Not MPar(Addrs)) And CPar(Addrs)) Xor ((Not MPar(Addrs)) And (Not CPar(Addrs)))) And (Not IsInv(Addrs))))
    End With
End Function

Public Function IsInv(ByRef Addrs As Addresses)
    With Addrs
        IsInv = (.J < .I)
    End With
End Function

Public Function GPar(ByRef Addrs As Addresses) As Boolean
    With Addrs
        'global parity = .J > .I
        'odd even bit of calls to switch that all
        'list modifying functions call via shift
       GPar = (.J > .I)
    End With
End Function

Public Function CPar(ByRef Addrs As Addresses) As Boolean
    With Addrs
        'cycles parity = -CInt(.J >= 0) = 1/0
        'odd even bit shift (inbetween methods
        'add, del, left and right and switch)
        CPar = (.J >= 0) '= CBool(-(.I + .J Mod 2))) And (GPar(Addrs) = (.I >= 0))
    End With
End Function

Public Function MPar(ByRef Addrs As Addresses) As Boolean
    With Addrs
        'method parity = -Cint(.I > 0) = 1/0
        'odd even bit of add/right and del/left
        'method calling, expressing direction
        MPar = (.I > 0)
    End With
End Function


Public Property Get Lowest(ByRef Addrs As Addresses, Optional ByVal Absolute As Boolean = False) As Long
    With Addrs
        Lowest = LeastOf(.G, .H)
    End With
End Property
Public Property Let Lowest(ByRef Addrs As Addresses, Optional ByVal Absolute As Boolean = False, ByVal RHS As Long)
    With Addrs
        If RHS < 0 And .G = Abs(RHS) Then
            .G = LongQuad
        Else
            If .G = 0 Then
                .G = Abs(RHS)
            Else
                .G = LeastOf(.G, Abs(RHS))
            End If
        End If
    End With
End Property

Public Property Get Highest(ByRef Addrs As Addresses, Optional ByVal Absolute As Boolean = False) As Long
    With Addrs
        Highest = LargeOf(.G, .H)
   End With
End Property
Public Property Let Highest(ByRef Addrs As Addresses, Optional ByVal Absolute As Boolean = False, ByVal RHS As Long)
    With Addrs
        If RHS < 0 And .H = Abs(RHS) Then
            .H = 0
        Else
            If .H = 0 Then
                .H = Abs(RHS)
            Else
                .H = LargeOf(.H, Abs(RHS))
            End If
        End If
    End With
End Property

Public Function Total(ByRef Addrs As Addresses) As Long
    With Addrs
        Total = Abs(.J + .I - .J + .I + -.I)
    End With
End Function

Public Function FirstNode(ByRef Addrs As Addresses) As Long
    With Addrs
        If .J > 0 Then Switch Addrs, .I, .B, .A, .D, .C
        
        
  

    End With
End Function
Public Function FinalNode(ByRef Addrs As Addresses) As Long
    With Addrs

        
        
        If .J > 0 Then Switch Addrs, .I, .B, .A, .D, .C
    End With
End Function


Public Sub InsertNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
    With Addrs
        FirstNode Addrs, True
        PushNode Addrs, Addr
    End With
End Sub

Public Sub DeleteNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
    With Addrs
        FinalNode Addrs, True
        KillNode Addrs, Addr
    End With
End Sub

Public Sub AppendNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
    With Addrs
        FinalNode Addrs, True
        PushNode Addrs, Addr
    End With
End Sub

Public Sub RemoveNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
    With Addrs
        FirstNode Addrs, True
        KillNode Addrs, Addr
    End With
End Sub


Public Function NextNode(ByRef Addrs As Addresses, Optional ByVal MoveTo As Boolean = False) As Long
    With Addrs
        If Total(Addrs) = 0 Then Exit Function

        Shift Addrs, MoveTo, True, True

        NextNode = Point(Addrs)
        
        Shift Addrs, MoveTo, True, False
        
        If MoveTo Then
            Shift Addrs, MoveTo, True, True
            
            Shift Addrs, MoveTo, True, False
        End If

    End With
End Function

Public Function PrevNode(ByRef Addrs As Addresses, Optional ByVal MoveTo As Boolean = False) As Long
    With Addrs
        If Total(Addrs) = 0 Then Exit Function
        
        If MoveTo Then
            Shift Addrs, MoveTo, False, True
            Shift Addrs, MoveTo, False, False
        End If
        
        Shift Addrs, MoveTo, True, True
        
        PrevNode = Point(Addrs)

        
        'DebugPrint Addrs
        
'        MoveTo = Not MoveTo
'
        Shift Addrs, MoveTo, True, False
        If MoveTo Then
            Shift Addrs, MoveTo, False, True
            Shift Addrs, MoveTo, False, False
        End If

    End With
End Function


Public Function Point(ByRef Addrs As Addresses) As Long
    With Addrs

        If Not MPar(Addrs) Then
             If .B <> 0 Then
                Point = .B
            Else
                Point = .C
            End If
        Else
            If .A <> 0 Then
                Point = .A
            Else
                Point = .B
            End If
        End If

    End With
End Function
'Public Function GoLeft(ByRef Addrs As Addresses)
''    With Addrs
''        Dim Revert As Boolean
''
''
''        Shift Addrs, Revert, False, True
'''
''        GoLeft = Abs(Point(Addrs))
''
''        Shift Addrs, Revert, False, False
''
''        Switch Addrs, .I, .C, .F, .E, .A
''
''
''       ' DebugPrint Addrs
''    End With
'
'    With Addrs
'        GoLeft = MoveIt(Addrs, CPar(Addrs))
'        MoveIt Addrs, Not (CPar(Addrs) = GPar(Addrs))
'    End With
'End Function
'Public Function GoRight(ByRef Addrs As Addresses)
''    With Addrs
''        Dim Revert As Boolean
''
''        'Switch Addrs, .I, .C, .F, .B, .A
''
''        Shift Addrs, Revert, True, True
''
''        GoRight = Abs(Point(Addrs))
''
''        Shift Addrs, Revert, True, False
''
''
''       ' DebugPrint Addrs
''    End With
'
'    With Addrs
'
'         MoveIt Addrs, CPar(Addrs)
'         GoRight = MoveIt(Addrs, (CPar(Addrs) = GPar(Addrs)))
'    End With
'End Function
'Public Function MoveIt(ByRef Addrs As Addresses, Optional ByVal LeftOrRight As Boolean) As Long
'    With Addrs
'
'        If LeftOrRight Then
'            .J = -.J
'             Swap .A, .B
'        End If
'
'
'        .J = -.J
'        If MPar(Addrs) Then Swap .A, .B
'
'        MoveIt = .A
'
'        .B = -.B
'        .A = -.A
'        .J = ((Abs(.J) - Abs(.I)) * IIf(.J > 0, 1, -1)) + IIf(LeftOrRight Xor (.J < 0), -1, 1)
'
'        If Not MPar(Addrs) Then Swap .A, .B
'
'        If (Not LeftOrRight) Then
'            Swap .A, .B
'            .J = -.J
'        End If
'
'    End With
'End Function



'
'1. G=0, H=0
'
'2.    G=0, F=0
'
'3.        G=0, C=0
'
'4.        G=0, C=0
'
'5.    G=0, E=0
'
'6.    G=0, E=0
'
'7.        G=0, F=0
'
'8. G=0, D=0
'
'9. G=0, D=0
'
'10.    G=0, H=0


'Swap D, B, A
'Swap A, B, D, F, H
'
'G=0, H=0
'
'Swap G, E, H
'Swap G, H, E, C, F
'
'   G=0, F=0
'
'Swap D, B, A
'Swap A, B, D, F, H
'
'       G=0, C=0
'
'Swap G, E, H
'Swap G, H, E, C, F
'
'       G=0, C=0
'
'Swap D, B, A
'Swap A, B, D, F, H
'
'   G=0, E=0
'
'Swap G, E, H
'Swap G, H, E, C, F
'
'   G=0, E=0
'
'Swap D, B, A
'Swap A, B, D, F, H
'
'       G=0, F=0
'
'Swap G, E, H
'Swap G, H, E, C, F
'
'G=0, D=0
'
'Swap D, B, A
'Swap A, B, D, F, H
'
'G=0, D=0
'
'Swap G, E, H
'Swap G, H, E, C, F
'
'   G=0, H=0

Public Function Switch(ByRef Addrs As Addresses, ByRef I As Long, ByRef A As Long, ByRef B As Long, ByRef C As Long, ByRef D As Long) As Long
    With Addrs
        'If -CInt(A = 0) + -CInt(B = 0) + -CInt(C = 0) + -CInt(D = 0) >= 3 Then Exit Function

        I = (-I)

        If -CInt(A = 0) + -CInt(B = 0) + -CInt(C = 0) + -CInt(D = 0) >= 3 Then Exit Function
        Swap .I, I
        Swap I, .J
        Swap .I, .J




        Swap A, B, C
        Swap A, C
        
        If (((((B < Abs(A)) And (A >= 0)) And (D > A)) And ((B > 0) Xor (C >= B))) Or (.I <> I)) Then
            
            Swap A, C
            Swap C, B, A
            Swap B, C, D
            Swap C, D
            
            Swap I, .J
            Swap .I, .J
            Swap I, .I


            Switch Addrs, I, A, D, C, B
 
            .J = .J - I
            Swap .I, .J
            Swap I, .J

          ' Swap .I, .J, I

        Else
            .J = Abs(.J) - Abs(I)
            
'            B = (-B / 2)
'
'            If Abs(D) > Abs(C) Then
'                D = ((((Abs(B) * 3) / 2) + -Abs(D / 2)) * 2)
'            Else
'                D = ((((-(Abs(B) / 2)) + -Abs(D)) * 2) / 3)
'            End If
'            C = -(C + A)
'
'            .J = (Abs(.J) + Abs(.I))
'
'             Swap .A, .H
            
        End If

'        D = (-(D / 2)) + .E
'        If GPar(Addrs) Then
'            B = (-(A / 2)) + (B / 2)
'        Else
'            B = ((A * 3) / 2) - (B / 2)
'        End If

        Swap D, B
        Swap B, C, D
        
       ' C = (C / 2)
        
       ' .H = Abs(.H) - Abs(A)
        
        
        
       ' Swap .E, .F, .H
       If .I = .J Then
            Swap I, .J
            Swap .I, .J
            Swap I, .I


            Switch Addrs, I, C, D, B, A
 
            .J = .J - I
            Swap .I, .J
            Swap I, .J
            
        End If
        If .I <> I Then
            Stop
        End If
 

        

        Swap .I, I
        Swap .I, .J
        Swap I, .J


    End With
End Function


'Swap A, B
'Swap D, E, F
'Swap A, B, C
'Swap A, C
'Pause
'
'Swap A, B, D, E
'Swap A, C
'Pause
'
'Swap A, C
'Swap F, E, D
'Swap C, A, B
'Swap A, F
'Pause
'
'Swap C, F, B, E
'Swap A,D
'Pause


Private Function Shift(ByRef Addrs As Addresses, ByRef Revert As Boolean, ByVal AddOrDel As Boolean, ByVal TopOrBtm As Boolean) As Boolean
    With Addrs
        .I = (-.I)
        
        If TopOrBtm Then
            If (.C > .A) Then
                Swap .C, .A
            Else
                Revert = True
            End If
            Swap .B, .A, .C
        Else
            Revert = True
            Swap .C, .D, .E
        End If
        
        If AddOrDel Then
            Swap .C, .D
        Else
            Swap .F, .A
        End If

        If (((.B < .A) And (.C < .A)) Xor (.C > .B)) And (Not Revert) Then
            Swap .A, .B
        ElseIf (Not AddOrDel) Then
            Swap .A, .C
        End If
        
        If (Not TopOrBtm) Then
            If (.C < .B) Then
                Swap .C, .B
            Else
                Revert = True
            End If
        Else
             Swap .F, .E, .D
        End If

        If Revert And (Not TopOrBtm) Then

            Swap .B, .A, .C
            Swap .C, .B

        End If
 
    End With
End Function


Private Property Get Focus(ByRef Addrs As Addresses, ByVal LeftOrRight As Boolean) As Long
    With Addrs
        If LeftOrRight Then
            Focus = .B
        Else
            Focus = .A
        End If
    End With
End Property
Private Property Let Focus(ByRef Addrs As Addresses, ByVal LeftOrRight As Boolean, ByVal RHS As Long)
    With Addrs
        If LeftOrRight Then
            .B = RHS
        Else
            .A = RHS
        End If
    End With
End Property


'Private Sub PushNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
Public Function PushNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long = 0)

    With Addrs
        Dim Revert As Boolean
        .I = (-.I)
        .J = (-.J)

        Switch Addrs, .I, .C, .F, .B, .A
        
        NewAddr Addrs, Addr
        DebugPrint Addrs
        
        If Total(Addrs) > 0 Then

            If Shift(Addrs, Revert, True, True) Then
                
                
            Else

            End If

        End If



        Register(Addr, 4) = Focus(Addrs, False)
        Focus(Addrs, False) = Addr



        If Shift(Addrs, Revert, True, False) Then

        Else

        End If

        .J = (-(.I - Abs(.J)))
        .I = (.I + IIf(.I > 0, 1, -1))
    End With
    


End Function

'Private Sub KillNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long)
Public Sub KillNode(ByRef Addrs As Addresses, Optional ByRef Addr As Long = 0)

    With Addrs
        If (Total(Addrs) > 0) Then
            Dim Revert As Boolean
            Dim Tot As Long
            Tot = Total(Addrs) * IIf(.J < 0, -1, 1)
            
            .I = (-.I)
            Revert = .J > 0
            .J = Abs(.J)

            If Shift(Addrs, Revert, False, True) Then

            Else

            End If

            Addr = Focus(Addrs, True)
            Focus(Addrs, True) = Register(Addr, 4)

            DelAddr Addrs, Addr
            DebugPrint Addrs

            .J = ((-.J) + IIf(.J > 0, 1, -1))
            .I = (.I + IIf(.I > 0, -1, 1))
            .J = (Abs(.J) * IIf(Revert, -1, 1))
            
            If Total(Addrs) > 0 Then
                
                If Shift(Addrs, Revert, False, False) Then

                Else

                End If
                
                Switch Addrs, .I, .C, .F, .E, .A
            
            Else
                .I = 0
                .J = 0
            End If
    
        End If
    End With

End Sub



Public Sub Main()

    Setup
    
    Form1.Show
    
'    Dim files As String
'    Dim fldr As String
'    Dim inc As Integer
'    Dim txt As String
'    Dim line As String
'    Dim out As String
'    Dim File As String
'    Dim fails As String
'
'    files = SearchPath("*.vbp", , "C:\Development\Basic\Classism", FindAll)
'    inc = 1
'    Do Until files = ""
'
'        fldr = RemoveNextArg(files, vbCrLf)
'        If PathExists(fldr, True) And Not fldr = "C:\Development\Basic\Classism\Project1.vbp" And InStr(fldr, "Repeat") = 0 Then
'            txt = ReadFile(fldr)
'            out = ""
'            Do While txt <> ""
'                line = RemoveNextArg(txt, vbCrLf)
'                Select Case NextArg(line, "=")
'                    Case "ExeName32"
'                        out = out & "ExeName32=""" & StrReverse("t" & RemoveArg(StrReverse(GetFileTitle(fldr)), "t")) & inc & ".exe""" & vbCrLf
'                        out = out & "Path32=""\Development\Basic\Classism\Classism""" & vbCrLf
'                    Case "Reference"
'                        out = out & "Reference=*\G{1A9C764D-90CE-47F5-A0DB-C2BBB9D863AB}#22.0#0#..\..\..\..\Projects\AddrPool\AddrPool.exe#A custom service for matters." & vbCrLf
'                    Case "Path32"
'                    Case "CondComp"
'                        out = out & Replace(Replace(Replace(Replace(line, "VBIDE=-1", "VBIDE=0"), "VBIDE = -1", "VBIDE = 0"), "VBIDE= -1", "VBIDE= 0"), "VBIDE =-1", "VBIDE =0") & vbCrLf
'                    Case "AutoRefresh"
'                        out = out & line & vbCrLf
'                        Exit Do
'                    Case Else
'                        out = out & line & vbCrLf
'                End Select
'            Loop
'            Kill fldr
'            fldr = GetFilePath(fldr) & "\" & StrReverse("t" & RemoveArg(StrReverse(GetFileTitle(fldr)), "t")) & inc & ".vbp"
'            WriteFile GetFilePath(fldr) & "\" & StrReverse("t" & RemoveArg(StrReverse(GetFileTitle(fldr)), "t")) & inc & ".vbp", out
'
'
'            txt = ReadFile(GetFilePath(fldr) & "\Form1.frm")
'            out = ""
'            Do While txt <> ""
'                line = RemoveNextArg(txt, vbCrLf)
'                Select Case NextArg(line, "=")
'                    Case "Caption"
'                        out = out & "Caption = """ & StrReverse("t" & RemoveArg(StrReverse(GetFileTitle(fldr)), "t")) & inc & ".exe""" & vbCrLf
'                        out = out & txt
'                        Exit Do
'
'                    Case Else
'                        out = out & line & vbCrLf
'                End Select
'            Loop
'            WriteFile GetFilePath(fldr) & "\Form1.frm", out
'
'
'
'            CurDir "C:\Development\Basic\Classism\Classism"
'            Debug.Print fldr
'            Dim pid As Long
'           ' pid = RunProcess("C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/open """ & fldr & """", , True)
'            pid = RunProcess("C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/make """ & fldr & """", , True)
'
'
'            inc = inc + 1
'
'
'        End If
'        Do While GetFilePath(NextArg(files, vbCrLf)) = GetFilePath(fldr)
'            Kill RemoveNextArg(files, vbCrLf)
'        Loop
'
'        If InStr(GetFilePath(fldr), GetFileTitle(fldr)) = 0 Then
'
'       '     Name GetFilePath(fldr) As GetFilePath(GetFilePath(fldr)) & "\" & GetFileTitle(fldr)
'        End If
'
'
'    Loop
'
'    WriteFile "C:\Development\Basic\Classism\Classism\Failed Compile.txt", fails
'
'    End

'    Exit Sub
'
'    PushNode Addrs
'    DebugAddrs Addrs
'
'    PushNode Addrs
'    DebugAddrs Addrs
'
''    PushNode Addrs
''    DebugAddrs Addrs
''
''    DebugLate = DebugLate & "Point: " & Point(Addrs) & vbCrLf
''    DebugLate = DebugLate & "NextNode: " & NextNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "NextNode: " & NextNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "NextNode: " & NextNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "Point: " & Point(Addrs) & vbCrLf
''
''    DebugLate = DebugLate & "PrevNode: " & PrevNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "PrevNode: " & PrevNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "PrevNode: " & PrevNode(Addrs) & vbCrLf
''    DebugLate = DebugLate & "Point: " & Point(Addrs) & vbCrLf
'
'    KillNode Addrs
'    DebugAddrs Addrs
'
'    PushNode Addrs
'    DebugAddrs Addrs
'
'    KillNode Addrs
'    DebugAddrs Addrs
'
'    KillNode Addrs
'    DebugAddrs Addrs
'
''    PushNode Addrs
''    DebugAddrs Addrs
''
''    PushNode Addrs
''    DebugAddrs Addrs
''
''    PushNode Addrs
''    DebugAddrs Addrs
''
''
''    KillNode Addrs
''    DebugAddrs Addrs
''
''    KillNode Addrs
''    DebugAddrs Addrs
''
''    KillNode Addrs
''    DebugAddrs Addrs
'
'    DebugFlush

    

End Sub
