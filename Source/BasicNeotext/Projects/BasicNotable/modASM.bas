Attribute VB_Name = "modASM"
Public mFileNumber As Integer

Public Enum vADDRegs
    AX_a = &H5
    AL_a = &H4
    AH_a = &H80
    
    BX_a = &H83
    BL_a = &H80
    BH_a = &H80
    
    CX_a = &H83
    CL_a = &H80
    CH_a = &H80
    
    DX_a = &H83
    DL_a = &H80
    DH_a = &H80
End Enum

Public Enum vCMPRegs
    AX_c = &H3D
    AL_c = &H3C
    AH_c = &H80
    
    BX_c = &H83
    BL_c = &H80
    BH_c = &H80
    
    CX_c = &H83
    CL_c = &H80
    CH_c = &H80
    
    DX_c = &H83
    DL_c = &H80
    DH_c = &H80
End Enum

Public Enum vMOVRegs
    AX_m = &HB8
    AL_m = &HB0
    AH_m = &HB4
    
    BX_m = &HBB
    BL_m = &HB3
    BH_m = &HB7
    
    CX_m = &HB9
    CL_m = &HB1
    CH_m = &HB5
    
    DX_m = &HBA
    DL_m = &HB2
    DH_m = &HB6
End Enum

Public Enum vINCRegs
    AX_i = &H40
    AL_i = &HFE
    AH_i = &HFE
    
    BX_i = &H43
    BL_i = &HFE
    BH_i = &HFE
    
    CX_i = &H41
    CL_i = &HFE
    CH_i = &HFE
    
    DX_i = &H42
    DL_i = &HFE
    DH_i = &HFE
End Enum

Public Enum vPUSHRegs
    AX_pu = &H50
    BX_pu = &H53
    CX_pu = &H51
    DX_pu = &H52
End Enum

Public Enum vPOPRegs
    AX_po = &H58
    BX_po = &H5B
    CX_po = &H59
    DX_po = &H5A
End Enum
Public Function HiByte(WordIn As Integer) As Byte
    If WordIn% And &H8000 Then
        HiByte = &H80 Or ((WordIn% And &H7FFF) \ &HFF)
    Else
        HiByte = WordIn% \ 256
    End If
End Function
Public Function LoByte(WordIn As Integer) As Byte
    LoByte = WordIn% And &HFF&
End Function

Private Function SplitVar(ByRef inline As String) As Variant()
    Dim sp() As String
    Dim va() As Variant
    
    sp = Split(inline, ",")
    ReDim va(LBound(sp) To UBound(sp)) As Variant
    
    Dim cnt As Long
    For cnt = LBound(sp) To UBound(sp)
        va(cnt) = sp(cnt)
    Next
    
    SplitVar = va
End Function

Public Sub ComAsmRun(ByVal FileName As String, ByVal AsmTxt As String)

    If Not PathExists(FileName, True) Then
        MsgBox "Please save the assembly file to compile it to .com with the same title name.", vbInformation
    Else
        
    
        Dim v_Assembler As New clsAssembly
        Dim v_ASM As New clsASM
        
            With v_Assembler
                .FileName = GetFilePath(FileName) & "\" & GetFileTitle(FileName) & ".com"
                
                If .CreateFile = False Then GoTo ContinueAnyWay:
                With v_ASM
                
                On Error GoTo filecompile
                
                Dim linenum As Long
                
                Dim inline As String
                Dim inreg As String
                Do Until AsmTxt = ""
                    inline = RemoveNextArg(AsmTxt, vbCrLf)
                    inline = RemoveNextArg(inline, ";")
                    linenum = linenum + 1
                    Select Case LCase(RemoveNextArg(inline, " "))
                        Case "add"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wADD AX_a, CLng(inline)
                                Case "al"
                                    .wADD AL_a, CLng(inline)
                                Case "ah"
                                    .wADD AH_a, CLng(inline)
                                Case "bx"
                                    .wADD BX_a, CLng(inline)
                                Case "bl"
                                    .wADD BL_a, CLng(inline)
                                Case "bh"
                                    .wADD BH_a, CLng(inline)
                                Case "cx"
                                    .wADD CX_a, CLng(inline)
                                Case "cl"
                                    .wADD CL_a, CLng(inline)
                                Case "ch"
                                    .wADD CH_a, CLng(inline)
                                Case "dx"
                                    .wADD DX_a, CLng(inline)
                                Case "dl"
                                    .wADD DL_a, CLng(inline)
                                Case "dh"
                                    .wADD DH_a, CLng(inline)
                            End Select
                        Case "cmp"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wCMP AX_c, CLng(inline)
                                Case "al"
                                    .wCMP AL_c, CLng(inline)
                                Case "ah"
                                    .wCMP AH_c, CLng(inline)
                                Case "bx"
                                    .wCMP BX_c, CLng(inline)
                                Case "bl"
                                    .wCMP BL_c, CLng(inline)
                                Case "bh"
                                    .wCMP BH_c, CLng(inline)
                                Case "cx"
                                    .wCMP CX_c, CLng(inline)
                                Case "cl"
                                    .wCMP CL_c, CLng(inline)
                                Case "ch"
                                    .wCMP CH_c, CLng(inline)
                                Case "dx"
                                    .wCMP DX_c, CLng(inline)
                                Case "dl"
                                    .wCMP DL_c, CLng(inline)
                                Case "dh"
                                    .wCMP DH_c, CLng(inline)
                            End Select
                        Case "mov"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wMOV AX_m, SplitVar(inline)
                                Case "al"
                                    .wMOV AL_m, SplitVar(inline)
                                Case "ah"
                                    .wMOV AH_m, SplitVar(inline)
                                Case "bx"
                                    .wMOV BX_m, SplitVar(inline)
                                Case "bl"
                                    .wMOV BL_m, SplitVar(inline)
                                Case "bh"
                                    .wMOV BH_m, SplitVar(inline)
                                Case "cx"
                                    .wMOV CX_m, SplitVar(inline)
                                Case "cl"
                                    .wMOV CL_m, SplitVar(inline)
                                Case "ch"
                                    .wMOV CH_m, SplitVar(inline)
                                Case "dx"
                                    .wMOV DX_m, SplitVar(inline)
                                Case "dl"
                                    .wMOV DL_m, SplitVar(inline)
                                Case "dh"
                                    .wMOV DH_m, SplitVar(inline)
                            End Select
                        Case "inc"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wINC AX_i
                                Case "al"
                                    .wINC AL_i
                                Case "ah"
                                    .wINC AH_i
                                Case "bx"
                                    .wINC BX_i
                                Case "bl"
                                    .wINC BL_i
                                Case "bh"
                                    .wINC BH_i
                                Case "cx"
                                    .wINC CX_i
                                Case "cl"
                                    .wINC CL_i
                                Case "ch"
                                    .wINC CH_i
                                Case "dx"
                                    .wINC DX_i
                                Case "dl"
                                    .wINC DL_i
                                Case "dh"
                                    .wINC DH_i
                            End Select
                        Case "pop"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wPOP AX_po
                                Case "bx"
                                    .wPOP BX_po
                                Case "cx"
                                    .wPOP CX_po
                                Case "dx"
                                    .wPOP DX_po
                            End Select
                        Case "psh", "push"
                            Select Case LCase(RemoveNextArg(inline, ","))
                                Case "ax"
                                    .wPUSH AX_pu
                                Case "bx"
                                    .wPUSH BX_pu
                                Case "cx"
                                    .wPUSH CX_pu
                                Case "dx"
                                    .wPUSH DX_pu
                            End Select
                        Case "je", "jme"
                            .wJE CLng(inline)
                        Case "jmp"
                            .wJMP CLng(inline)
                        Case "int"
                            .wINT CLng(inline)
                        Case "ret", "return"
                            .wRET
                        Case "dat", "data", "adddata"
                            .wAddData RemoveQuotedArg(inline, """", """")
                        
                    
                    End Select
                    
                
                
                Loop
                End With
                
                .CloseFile

                
                If MsgBox("Run file?", vbQuestion + vbYesNo, "Run?") = vbYes Then Shell .FileName, vbNormalFocus
                
                
            End With
        GoTo ContinueAnyWay
filecompile:
        Err.Clear

        MsgBox "Highlited is an error compiling the ASM to a COM program.", vbExclamation

        frmMain.txtMain.SelectRow (linenum - 1)
        
ContinueAnyWay:
        Set v_Assembler = Nothing
        Set v_ASM = Nothing

    End If

End Sub
