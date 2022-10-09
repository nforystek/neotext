Attribute VB_Name = "modLogSum"
#Const modLogSum = -1

Function UpdateLogFiles()

    If Not PathExists(AppPath & "SumMake.lst", ture) Then WriteFile AppPath & "SumMake.lst", ""
    Dim fso, f, d, s
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set f = fso.GetFolder("C:\Development\Publish\Htdocs\ipub\logs")

    Dim atDateTime As String
    Dim atfile As String
    Dim fname As String
    Dim ftxt As String
    fname = "C:\Development\Publish\Htdocs\ipub\logs\" & Format(Now, "mm-dd-yyyy_hh-mm-ss") & ".txt"
    
    If f.Files.Count > 0 Then

        For Each s In f.Files

            If (GetFileExt(s.Name, True, True) = "txt") And (Not (s.Name = fname)) Then

                If Format(Replace(NextArg(GetFileTitle(s.Name), "_"), "-", "/") & " " & _
                    Replace(RemoveArg(GetFileTitle(s.Name), "_"), "-", ":"), "dd/mm/yyyy hh:mm:ss") > atDateTime Then
                    atDateTime = Format(Replace(NextArg(GetFileTitle(s.Name), "_"), "-", "/") & " " & _
                    Replace(RemoveArg(GetFileTitle(s.Name), "_"), "-", ":"), "dd/mm/yyyy hh:mm:ss")
                    atfile = s.Path
                End If
                ftxt = ftxt & ReadFile(s.Path)

            End If
        Next

    End If

    Dim info As String
    info = "No change log recorded In the source code, this release may be a result of testing" + _
        vbCrLf + "the cycle On a newly formatted installation of other possible media discrepancies."
    If Right(ftxt, Len(info)) = info Then info = ""

    If PathExists(atfile, True) Then atfile = ReadFile(atfile)
    
    atfile = IterateFolder(1, fso.GetFolder("C:\Development\Neotext"), ftxt) & _
             IterateFolder(1, fso.GetFolder("C:\Development\Publish\Htdocs\ipub\apps\binary"), ftxt)
    ftxt = atfile
    atfile = atfile & IterateFolder(2, fso.GetFolder("C:\Development\Neotext"), ftxt)
    ftxt = atfile


    If atfile = "" Then
        atfile = info
    End If
    If atfile <> "" Then

        WriteFile fname, atfile
        Dim curtext As String
        curtext = ReadFile(AppPath & "SumMake.lst")
        
        ftxt = ""
        Do Until atfile = ""
            info = RemoveNextArg(atfile, vbCrLf)
            If InStr(1, info, ".vbp", vbTextCompare) > 0 Then
                info = "C:\Development\" & RemoveQuotedArg(info, """", """") & vbCrLf
                If InStr(1, ftxt, info, vbTextCompare) = 0 And InStr(1, curtext, info, vbTextCompare) = 0 Then
                    ftxt = ftxt & info
                End If
                
            End If
        Loop
        If ftxt <> "" Then WriteFile AppPath & "SumMake.lst", ftxt & curtext
    End If
    
    atfile = atfile & IterateFolder(3, fso.GetFolder("C:\Development\Neotext"), ftxt)
    ftxt = atfile
    atfile = atfile & IterateFolder(4, fso.GetFolder("C:\Development\Publish\Htdocs\ipub\apps\binary"), ftxt)
    If atfile = "" Then
        atfile = info
    End If
    If atfile <> "" Then

        WriteFile fname, atfile
    
    End If
    
    UpdateLogFiles = atfile

End Function
Function Pad(text, number)
    If Len(CStr(text)) < number Then
        Pad = String(number - Len(CStr(text)), " ") & text
    Else
        Pad = text
    End If
End Function
Function IterateFolder(pass, d, old)
    Dim s, f, diz, chk, lin, tmp
    
    For Each s In d.SubFolders
        diz = diz & IterateFolder(pass, s, old)
    Next
    Select Case pass
        Case 1, 2

            If diz <> "" Then
                If InStr(d.Path, "Project1") = 0 And InStr(d.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(d.Path, "\Template") = 0 _
                    And InStr(d.Path, "\ActiveX") = 0 And InStr(d.Path, "\Example") = 0 And InStr(d.Path, "(0)") = 0 And InStr(d.Path, "(1)") = 0 _
                    And InStr(d.Path, "Project3") = 0 And InStr(d.Path, "(2)") = 0 And InStr(d.Path, "(3)") = 0 And InStr(d.Path, "\Groups") = 0 Then

                    If InStr(old, Pad("<DIR>", 20) & " " & Pad(GetFileDate(d.Path), 23) & " """ & Replace(Replace(d.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then
                        diz = Pad("<DIR>", 20) & " " & Pad(GetFileDate(d.Path), 23) & " """ & Replace(Replace(d.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf & diz
                    End If
                End If
            End If
    End Select

    For Each f In d.Files


        Select Case pass
            Case 1

                Select Case GetFileExt(f.Path, True, True)
                    Case "ctl", "bas", "cls", "dsr"
                        If InStr(d.Path, "Project1") = 0 And InStr(f.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(f.Path, "\Template") = 0 _
                            And InStr(d.Path, "\ActiveX") = 0 And InStr(f.Path, "\Example") = 0 And InStr(f.Path, "(0)") = 0 And InStr(f.Path, "(1)") = 0 _
                            And InStr(d.Path, "Project3") = 0 And InStr(f.Path, "(2)") = 0 And InStr(f.Path, "(3)") = 0 And InStr(f.Path, "\Groups") = 0 Then

                            If InStr(old, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then
                                diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                            End If
                        End If

                End Select
            Case 2

                If InStr(d.Path, "Project1") = 0 And InStr(f.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(f.Path, "\Template") = 0 _
                    And InStr(d.Path, "\ActiveX") = 0 And InStr(f.Path, "\Example") = 0 And InStr(f.Path, "(0)") = 0 And InStr(f.Path, "(1)") = 0 _
                    And InStr(d.Path, "Project3") = 0 And InStr(f.Path, "(2)") = 0 And InStr(f.Path, "(3)") = 0 And InStr(f.Path, "\Groups") = 0 Then

                    Select Case GetFileExt(f.Path, True, True)
                        Case "vbp"
                            chk = ReadFile(f.Path)
                            Do Until chk = ""
                                tmp = RemoveNextArg(chk, vbCrLf)
                                
                                lin = Replace(RemoveArg(RemoveArg(tmp, "="), ";"), """", "")
                                If Not PathExists(lin, True) Then

                                    lin = MapFolder(GetFilePath(f.Path), lin)
                                    If Not PathExists(lin, True) Then

                                        lin = NextArg(RemoveArg(RemoveArg(RemoveArg(tmp, "#"), "#"), "#"), "#")
                                        lin = MapFolder(GetFilePath(f.Path), lin)

                                    End If
                                End If
                                

                                Select Case GetFileExt(lin, True, True)
                                    Case "ctl", "bas", "cls", "dsr"

                                    If PathExists(lin, True) Then


                                        If (InStr(old, GetFileTitle(f.Path)) > 0 Or InStr(diz, GetFileTitle(f.Path)) > 0) And (InStr(old, Pad(GetFileSize(lin), 20) & " " & Pad(GetFileDate(lin), 23) & " """ & Replace(Replace(lin, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0 Or _
                                            InStr(diz, Pad(GetFileSize(lin), 20) & " " & Pad(GetFileDate(lin), 23) & " """ & Replace(Replace(lin, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0) Then


                                                If (InStr(diz, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 And _
                                                    InStr(diz, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0) Then

                                                    diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                                                End If
                    
                                        End If
    
                                    End If
                                    Case "exe", "ocx", "dll"
    
                                    If PathExists(lin, True) Then

                                        If (InStr(old, GetFileTitle(lin)) > 0 Or InStr(diz, GetFileTitle(lin)) > 0) Then
                                                    
                                            If InStr(old, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then

                                                If InStr(diz, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then
                                                    diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                                                End If
                                            End If

                                        Else
                                            If InStr(old, Pad("<DIR>", 20) & " " & Pad(GetFileDate(GetFilePath(lin)), 23) & " """ & Replace(Replace(GetFilePath(lin), "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0 Then
                                                
    
                                                If InStr(diz, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0 Then
    
                                                    If InStr(diz, Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then
                                                        diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                                                    End If
                                                End If
                                            End If
                                        End If
    
                                    End If

                                End Select
                            
                            Loop

                    End Select
                End If

            Case 3

                If InStr(d.Path, "Project1") = 0 And InStr(f.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(f.Path, "\Template") = 0 _
                    And InStr(d.Path, "\ActiveX") = 0 And InStr(f.Path, "\Example") = 0 And InStr(f.Path, "(0)") = 0 And InStr(f.Path, "(1)") = 0 _
                    And InStr(d.Path, "Project3") = 0 And InStr(f.Path, "(2)") = 0 And InStr(f.Path, "(3)") = 0 And InStr(f.Path, "\Groups") = 0 Then

                    Select Case GetFileExt(f.Path, True, True)
                        Case "vbp"

                           lin = GetCompileFile(f.Path)

                           If PathExists(lin, True) Then


                               If InStr(old, Pad(GetFileSize(lin), 20) & " " & Pad(GetFileDate(lin), 23) & " """ & Replace(Replace(lin, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 And _
                               InStr(diz, Pad(GetFileSize(lin), 20) & " " & Pad(GetFileDate(lin), 23) & " """ & Replace(Replace(lin, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) = 0 Then


                                   diz = diz & Pad(GetFileSize(lin), 20) & " " & Pad(GetFileDate(lin), 23) & " """ & Replace(Replace(lin, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf

                               End If

                           End If



                        Case "nsi"

                            If InStr(d.Path, "Project1") = 0 And InStr(f.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(f.Path, "\Template") = 0 _
                                And InStr(d.Path, "\ActiveX") = 0 And InStr(f.Path, "\Example") = 0 And InStr(f.Path, "(0)") = 0 And InStr(f.Path, "(1)") = 0 _
                                And InStr(d.Path, "Project3") = 0 And InStr(f.Path, "(2)") = 0 And InStr(f.Path, "(3)") = 0 And InStr(f.Path, "\Groups") = 0 Then
    
                                If InStr(old, Pad("<DIR>", 20) & " " & Pad(GetFileDate(GetFilePath(GetFilePath(f.Path))), 23) & " """ & Replace(Replace(GetFilePath(GetFilePath(f.Path)), "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0 Then
    
                                    diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                                End If
    
                            End If

                    End Select
                End If
            Case 4
                Select Case GetFileExt(f.Path, True, True)
                    Case "exe", "dll", "ocx"
                        If InStr(d.Path, "Project1") = 0 And InStr(f.Path, "Copy of") = 0 And InStr(d.Path, "Test") = 0 And InStr(f.Path, "\Template") = 0 _
                             And InStr(d.Path, "\ActiveX") = 0 And InStr(f.Path, "\Example") = 0 And InStr(f.Path, "(0)") = 0 And InStr(f.Path, "(1)") = 0 _
                             And InStr(d.Path, "Project3") = 0 And InStr(f.Path, "(2)") = 0 And InStr(f.Path, "(3)") = 0 And InStr(f.Path, "\Groups") = 0 Then

                            If InStr(old, Pad("<DIR>", 20) & " " & Pad(GetFileDate(GetFilePath(GetFilePath(f.Path))), 23) & " """ & Replace(Replace(GetFilePath(GetFilePath(f.Path)), "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf) > 0 Then

                                diz = diz & Pad(GetFileSize(f.Path), 20) & " " & Pad(GetFileDate(f.Path), 23) & " """ & Replace(Replace(f.Path, "C:\Development\Neotext", "Neotext"), "C:\Development\Publish\Htdocs\ipub\apps\binary", "PUBLIC") & """" & vbCrLf
                            End If


                        End If

                End Select
        End Select
    Next
    IterateFolder = diz

End Function

Function GetCompileFile(VBPFile)

    Dim inVar
    Dim inVal
    

    inVar = ReadFile(VBPFile)
    If InStr(inVar, "CompatibleEXE32=""") > 0 Then
        RemoveNextArg inVar, "CompatibleEXE32="""
        inVar = RemoveNextArg(inVar, """" & vbCrLf)
        GetCompileFile = MapFolder(GetFilePath(VBPFile), inVar)
    ElseIf InStr(inVar, "ExeName32=""") > 0 Then
        inVal = inVar
        RemoveNextArg inVal, "ExeName32="""
        inVal = RemoveNextArg(inVal, """" & vbCrLf)
        GetCompileFile = MapFolder(GetFilePath(VBPFile), inVal)

    ElseIf InStr(inVar, "Path32=""") > 0 And InStr(inVar, "ExeName32=""") > 0 Then
        inVal = inVar
        RemoveNextArg inVal, "ExeName32="""
        inVal = RemoveNextArg(inVal, """" & vbCrLf)
        RemoveNextArg inVar, "Path32="""
        inVar = RemoveNextArg(inVar, """" & vbCrLf)
    If inVar = "" Then inVar = GetFilePath(VBPFile)
        GetCompileFile = MapFolder(inVar, inVal)
   
    End If


End Function

Private Function MapFolder(RootURL, vURL)
    'concatenates vURL to the RootURL properly by blind path specifications

    
    Do While InStr(vURL, "..\") > 0

    RootURL = GetFilePath(RootURL)
        vURL = Mid(vURL, InStr(vURL, "..\") + 3)

    Loop

    
    Do While InStr(vURL, "../") > 0
    RootURL = GetFilePath(RootURL)
        vURL = Mid(vURL, InStr(vURL, "../") + 3)

    Loop

    If InStr(RootURL, "\") > 0 Then
        RootURL = Replace(RootURL, "/", "\")
        vURL = Replace(vURL, "/", "\")
        If Left(vURL, 1) = "\" And Right(RootURL, 1) = "\" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "\" And Right(RootURL, 1) <> "\" Then
            vURL = RootURL & "\" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "\" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "\"
    ElseIf InStr(RootURL, "/") > 0 Then
        RootURL = Replace(RootURL, "\", "/")
        vURL = Replace(vURL, "\", "/")
        
        If Left(vURL, 1) = "/" And Right(RootURL, 1) = "/" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "/" And Right(RootURL, 1) <> "/" Then
            vURL = RootURL & "/" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "/" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "/"

    End If


    
    MapFolder = vURL

End Function


