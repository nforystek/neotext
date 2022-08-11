VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub LoadFileDescCheck(ByVal FileName As String)
    Dim txt As String
    Dim out As String
    Dim user As String
    Dim desc As String
    Dim back As String
    Dim crc1 As Long
    Dim crc2 As Long
    Dim size As Long
    Select Case GetFileExt(FileName, True, True)
        Case "bas", "ctl", "cls", "frm", "dob", "dsr"
            out = ""
            txt = ReadFile(FileName)
            back = txt
            size = Len(txt)
            Do Until txt = ""
                If txt Like "*" & vbCrLf & Chr(65) & "ttribute *." & Chr(86) & "B_Description = ""*""" & vbCrLf & "*" Then
                    out = out & RemoveNextArg(txt, vbCrLf & Chr(65) & "ttribute ", vbTextCompare, False)
                    If Not (InStrRev(out, "' _" & vbCrLf) > InStrRev(out, vbCrLf)) Then
                        user = NextArg(txt, "." & Chr(86) & "B_Description = """, vbTextCompare, False)
                        If InStr(1, StrReverse(NextArg(StrReverse(out), vbLf & vbCr, vbTextCompare, False)), user, vbTextCompare) > 0 Then
                            desc = RemoveArg(NextArg(txt, """" & vbCrLf, vbTextCompare, False), "." & Chr(86) & "B_Description = """, vbTextCompare, False)

                         '   Debug.Print StrReverse(NextArg(StrReverse(out), vbLf & vbCr)) & " ' _"
                             
                            out = RTrim(out) & " ' _" & vbCrLf & desc & vbCrLf & "Attribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
                            If desc <> "" Then
                                crc1 = crc1 + Len(desc)
                                size = size + Len(desc) - 7
                            End If
                            desc = RemoveArg(NextArg(txt, """" & vbCrLf, vbTextCompare, False), """")
                            If desc <> "" Then
                                crc2 = crc2 + Len(desc)
                                size = size - Len(desc)
                            End If
                            RemoveNextArg txt, vbCrLf, vbTextCompare, False
                          '  Debug.Print desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
                         '   Stop
                        Else
                            out = out & vbCrLf & Chr(65) & "ttribute "
                        End If
                    Else
                        out = out & vbCrLf & Chr(65) & "ttribute "
                    End If
                Else
                    out = out & txt
                    txt = ""
                End If
            Loop
            If (out <> back) And ((Len(out) - size) - (Len(out) - Len(back))) = (Len(back) - size) And (crc1 = crc2) Then
                WriteFile FileName, out
            End If
    End Select

End Sub
Private Sub SaveFileDescCheck(ByVal FileName As String)
    Dim txt As String
    Dim out As String
    Dim user As String
    Dim desc As String
    Dim back As String
    Dim crc1 As Long
    Dim crc2 As Long
    Dim size As Long
    
    Select Case GetFileExt(FileName, True, True)
        Case "bas", "ctl", "cls", "frm", "dob", "dsr"
            out = ""
            txt = ReadFile(FileName)
            back = txt
            size = Len(back)
            Do Until txt = ""
                If txt Like "*" & vbCrLf & Chr(65) & "ttribute *." & Chr(86) & "B_Description = ""*""" & vbCrLf & "*" Then
                    out = out & RemoveNextArg(txt, "' _" & vbCrLf, vbTextCompare, False)
                    user = StrReverse(NextArg(RemoveArg(StrReverse(out), "(", vbTextCompare, False), " ", vbTextCompare, False)) ' "." & Chr(86) & "B_Description = """, vbTextCompare, False)
                    
                    If InStr(txt, "Attribute " & user & ".VB_Description = ") > 0 Then

                        desc = NextArg(RemoveArg(txt, vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """, vbTextCompare, False), """" & vbCrLf, vbTextCompare, False)
                        If desc <> "" Then
                            crc2 = crc2 + Len(desc)
                            size = size + Len(desc)
                        End If
                
                        desc = RemoveNextArg(txt, vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """, vbTextCompare, False)
                        If desc <> "" Then
                            crc1 = crc1 + Len(desc)
                            size = size - Len(desc)
                        End If
                        
                        RemoveNextArg txt, """" & vbCrLf, vbTextCompare, False
                        Debug.Print StrReverse(NextArg(StrReverse(out), vbLf & vbCr, vbTextCompare, False)) & " ' _" & vbCrLf & desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf

                                                    
                        out = RTrim(out) & " ' _" & vbCrLf & desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf

                    Else
                        desc = RemoveNextArg(txt, vbCrLf, vbTextCompare, False)
                        If (desc <> "") And (Not (Left(Trim(Replace(desc, vbTab, "")), 1) = "'")) And (Not (Right(Trim(Replace(desc, vbTab, "")), 1) = "_")) Then
                        
                                crc2 = crc2 + Len(desc)
                                If desc <> "" Then
                                    crc1 = crc1 + Len(desc)
                                    size = size - Len(desc)
                                End If
                                
                                Debug.Print StrReverse(NextArg(StrReverse(out), vbLf & vbCr)) & " ' _" & vbCrLf & desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
                                Debug.Print NextArg(txt, vbCrLf)
                                
                                out = RTrim(out) & " ' _" & vbCrLf & desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
                        Else
                            out = out & desc
                        End If
                        'Stop
                    End If
                   
                Else
                    out = out & txt
                    txt = ""
                End If
            Loop

            If (out <> back) And ((Len(out) - size) - (Len(out) - Len(back))) = (Len(back) - size) And (crc1 = crc2) Then
                WriteFile FileName, out
            End If
    End Select

End Sub '
'Private Sub LoadFileDescCheck(ByVal FileName As String)
'    Dim txt As String
'    Dim out As String
'    Dim user As String
'    Dim desc As String
'    Dim back As String
'    Dim crc1 As Long
'    Dim crc2 As Long
'    Dim size As Long
'    Select Case GetFileExt(FileName, True, True)
'        Case "bas", "ctl", "cls", "frm", "dob", "dsr"
'            out = ""
'            txt = ReadFile(FileName)
'            back = txt
'            size = Len(txt)
'            Do Until txt = ""
'                If txt Like "*" & vbCrLf & Chr(65) & "ttribute *." & Chr(86) & "B_Description = ""*""" & vbCrLf & "*" Then
'                    out = out & RemoveNextArg(txt, vbCrLf & Chr(65) & "ttribute ", vbTextCompare, False)
'                    If Not (InStrRev(out, "' _" & vbCrLf) > InStrRev(out, vbCrLf)) Then
'                        user = NextArg(txt, "." & Chr(86) & "B_Description = """, vbTextCompare, False)
'                        If InStr(1, StrReverse(NextArg(StrReverse(out), vbLf & vbCr, vbTextCompare, False)), user, vbTextCompare) > 0 Then
'                            desc = RemoveArg(RemoveNextArg(txt, """" & vbCrLf, vbTextCompare, False), "." & Chr(86) & "B_Description = """, vbTextCompare, False)
'                            out = RTrim(out) & " ' _" & vbCrLf & desc & vbCrLf & "Attribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
'                            If desc <> "" Then
'                                crc1 = crc1 + Len(desc)
'                                size = size + Len(desc) - 7
'                            End If
'                            desc = StrReverse(RemoveNextArg(out, vbLf & vbCr & "_ '", vbTextCompare, False))
'                           ' desc = RemoveArg(RemoveNextArg(txt, """" & vbCrLf, vbTextCompare, False), """")
'                            If desc <> "" Then
'                                crc2 = crc2 + Len(desc)
'                                size = size - Len(desc)
'                            End If
'                           ' RemoveNextArg txt, vbCrLf, vbTextCompare, False
'                        Else
'                            out = out & vbCrLf & Chr(65) & "ttribute "
'                        End If
'                    Else
'                        out = out & vbCrLf & Chr(65) & "ttribute "
'                    End If
'                Else
'                    out = out & txt
'                    txt = ""
'                End If
'            Loop
'            If (out <> back) And ((Len(out) - size) - (Len(out) - Len(back))) = (Len(back) - size) And (crc1 = crc2) Then
'                WriteFile Replace(FileName, "Client", "Client2"), out
'            End If
'    End Select
'
'End Sub
'Private Sub SaveFileDescCheck(ByVal FileName As String)
'    Dim txt As String
'    Dim out As String
'    Dim user As String
'    Dim desc As String
'    Dim back As String
'    Dim crc1 As Long
'    Dim crc2 As Long
'    Dim size As Long
'
'    Select Case GetFileExt(FileName, True, True)
'        Case "bas", "ctl", "cls", "frm", "dob", "dsr"
'            out = ""
'            txt = ReadFile(FileName)
'            back = txt
'            size = Len(back)
'            Do Until txt = ""
'                If txt Like "*" & vbCrLf & Chr(65) & "ttribute *." & Chr(86) & "B_Description = ""*""" & vbCrLf & "*" Then
'                    out = out & RemoveNextArg(txt, vbCrLf & Chr(65) & "ttribute ", vbTextCompare, False)
'                    If Not (InStrRev(out, "' _" & vbCrLf) > InStrRev(out, vbCrLf)) Then
'                        user = NextArg(txt, "." & Chr(86) & "B_Description = """, vbTextCompare, False)
'                        If InStr(1, StrReverse(NextArg(RemoveArg(StrReverse(out), vbLf & vbCr & "_ '", vbTextCompare, False), vbLf & vbCr, vbTextCompare, False)), user, vbTextCompare) > 0 Then
'                            desc = RemoveArg(NextArg(txt, """" & vbCrLf, vbTextCompare, False), """", vbTextCompare, False)
'                            If desc <> "" Then
'                                crc1 = crc1 + Len(desc)
'                                size = size - Len(desc)
'                            End If
'                            out = StrReverse(out)
'                            size = size + Len(desc)
'                            desc = StrReverse(RemoveNextArg(out, vbLf & vbCr & "_ '", vbTextCompare, False))
'                            If desc <> "" Then
'                                crc2 = crc2 + Len(desc)
'                                size = size + Len(desc)
'                            End If
'                            out = StrReverse(out)
'                            out = RTrim(out) & " ' _" & vbCrLf & desc & vbCrLf & Chr(65) & "ttribute " & user & "." & Chr(86) & "B_Description = """ & desc & """" & vbCrLf
'                            RemoveNextArg txt, vbCrLf, vbTextCompare, False
'                        Else
'                            out = out & vbCrLf & Chr(65) & "ttribute "
'                        End If
'                    Else
'                        out = out & vbCrLf & Chr(65) & "ttribute "
'                    End If
'                Else
'                    out = out & txt
'                    txt = ""
'                End If
'            Loop
'            If (out <> back) And ((Len(out) - size) - (Len(out) - Len(back))) = (Len(back) - size) And (crc1 = crc2) Then
'                WriteFile Replace(FileName, "Client", "Client3"), out
'            End If
'    End Select
'
'End Sub



Private Sub Form_Load()
    Kill "C:\Development\Neotext\Common\Projects\NTAdvFTP61\Client.cls"
    FileCopy "C:\Development\Neotext\Common\Projects\Copy of NTAdvFTP61\client.cls", "C:\Development\Neotext\Common\Projects\NTAdvFTP61\client.cls"
    
'    LoadFileDescCheck "C:\Development\Neotext\Common\Projects\NTAdvFTP61\Client.cls"
    SaveFileDescCheck "C:\Development\Neotext\Common\Projects\NTAdvFTP61\Client.cls"
    End
End Sub
