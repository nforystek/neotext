Attribute VB_Name = "modGenerate"
Option Explicit
'TOP DOWN

Private Const DLLLIBTYPE = "DLL $1 REBOOT_PROTECTED"
Private Const REGDLLLIBTYPE = "REGDLL $1 REBOOT_PROTECTED"
Private Const REGEXELIBTYPE = "REGEXE $1 REBOOT_PROTECTED"
Private Const TLBLIBTYPE = "TLB $1 REBOOT_PROTECTED"

Private fso As Object

Private Function GetFileEntry(ByVal Folder As String, ByVal File As String, ByVal LibType As String, ByRef ddfText As String, ByRef iniText As String) As String
    Select Case GetFileExt(File)
        Case ".dll", ".ocx", ".oca"
            If GetFileTitle(File) = "maxlandlib" Then
                If Not LibType = DLLLIBTYPE Then LibType = TLBLIBTYPE
            Else
                If Not LibType = DLLLIBTYPE Then LibType = REGDLLLIBTYPE
            End If
        Case ".exe"
            If Not LibType = DLLLIBTYPE Then LibType = REGEXELIBTYPE
            File = ""
        Case ".tlb", ".olb"
            LibType = TLBLIBTYPE
            Exit Function
            
    End Select
    'Dim fvi As FILEVERINFO
    Dim fvi As FILEPROPERTIE
    Dim inc As Boolean
    Dim cnt As Long
    Dim ex As String
    Dim verInfo As String
    verInfo = "0.0.0.0"
   
    If PathExists(Folder & "\Common\Binary\" & GetFileName(File), True) Then
        If Not LibType = DLLLIBTYPE Then LibType = REGDLLLIBTYPE
        If GetFileExt(File) <> ".vbp" Then
            If InStr(1, ddfText, GetFileName(File), vbTextCompare) = 0 Then
                inc = True
                If Program.Excludes.count > 0 Then
                    For cnt = 1 To Program.Excludes.count
                        ex = Program.Excludes(cnt)
                        Select Case RemoveNextArg(ex, " ")
                            Case "*"
                                If LikeCompare(LCase(GetFileName(File)), LCase(ex)) Then inc = False
                            Case "?"
                                If LCase(GetFileName(File)) = LCase(ex) Then inc = False
                        End Select
                        If Not inc Then Exit For
                    Next
                End If
                If inc Then
                    fvi = GetFileInfo(Folder & "\Common\Binary\" & GetFileName(File))
                    'Debug.Print """" & Folder & "\Common\Binary\" & GetFileName(File) & """"
                    ddfText = ddfText & """" & Folder & "\Common\Binary\" & GetFileName(File) & """ """ & GetFileName(File) & """" & vbCrLf
                    iniText = iniText & "Shared=" & GetFileName(File) & "|" & GetFileSize(File) & "|" & GetFileDate(File) & "|" & fvi.FileVersion & vbCrLf
                End If
            End If

        End If

    ElseIf PathExists(Folder & "\Windows\ActiveX\" & GetFileName(File), True) Then
        If GetFileExt(File) <> ".vbp" Then
            If InStr(1, ddfText, GetFileName(File), vbTextCompare) = 0 Then
                inc = True
                If Program.Excludes.count > 0 Then
                    For cnt = 1 To Program.Excludes.count
                        ex = Program.Excludes(cnt)
                        Select Case RemoveNextArg(ex, " ")
                            Case "*"
                                If LikeCompare(LCase(GetFileName(File)), LCase(ex)) Then inc = False
                            Case "?"
                                If LCase(GetFileName(File)) = LCase(ex) Then inc = False
                        End Select
                        If Not inc Then Exit For
                    Next
                End If
                If inc Then
                    fvi = GetFileInfo(Folder & "\Windows\ActiveX\" & GetFileName(File))
                    'Debug.Print """" & Folder & "\Windows\ActiveX\" & GetFileName(File) & """"
                    ddfText = ddfText & """" & Folder & "\Windows\ActiveX\" & GetFileName(File) & """ """ & GetFileName(File) & """" & vbCrLf
                    iniText = iniText & "Shared=" & GetFileName(File) & "|" & GetFileSize(File) & "|" & GetFileDate(File) & "|" & fvi.FileVersion & vbCrLf
                End If
            End If

        End If
        If Not LibType = DLLLIBTYPE Then LibType = REGDLLLIBTYPE

    ElseIf PathExists(Folder & "\Windows\System\" & GetFileName(File), True) Then
        If GetFileExt(File) <> ".vbp" Then
            If InStr(1, ddfText, GetFileName(File), vbTextCompare) = 0 Then
                inc = True
                If Program.Excludes.count > 0 Then
                    For cnt = 1 To Program.Excludes.count
                        ex = Program.Excludes(cnt)
                        Select Case RemoveNextArg(ex, " ")
                            Case "*"
                                If LikeCompare(LCase(GetFileName(File)), LCase(ex)) Then inc = False
                            Case "?"
                                If LCase(GetFileName(File)) = LCase(ex) Then inc = False
                        End Select
                        If Not inc Then Exit For
                    Next
                End If
                If inc Then
                    fvi = GetFileInfo(Folder & "\Windows\System\" & GetFileName(File))
                    'Debug.Print """" & Folder & "\Windows\System\" & GetFileName(File) & """"
                    ddfText = ddfText & """" & Folder & "\Windows\System\" & GetFileName(File) & """ """ & GetFileName(File) & """" & vbCrLf
                    iniText = iniText & "System=" & GetFileName(File) & "|" & GetFileSize(File) & "|" & GetFileDate(File) & "|" & fvi.FileVersion & vbCrLf
                End If
            End If
        End If
        LibType = TLBLIBTYPE

    ElseIf PathExists(SysPath & GetFileName(File), True) Then
        'Debug.Print Folder & "\Windows\System\" & GetFileName(File)
        If Not PathExists(Folder & "\Windows\System\" & GetFileName(File), True) Then
            FileCopy SysPath & GetFileName(File), Folder & "\Windows\System\" & GetFileName(File)
            If GetFileExt(File) <> ".vbp" Then
                If InStr(1, ddfText, GetFileName(File), vbTextCompare) = 0 Then
                    inc = True
                    If Program.Excludes.count > 0 Then
                        For cnt = 1 To Program.Excludes.count
                            ex = Program.Excludes(cnt)
                            Select Case RemoveNextArg(ex, " ")
                                Case "*"
                                    If LikeCompare(LCase(GetFileName(File)), LCase(ex)) Then inc = False
                                Case "?"
                                    If LCase(GetFileName(File)) = LCase(ex) Then inc = False
                            End Select
                            If Not inc Then Exit For
                        Next
                    End If
                    If inc Then
                        fvi = GetFileInfo(Folder & "\Windows\System\" & GetFileName(File))
                        'Debug.Print """" & Folder & "\Windows\System\" & GetFileName(File) & """"
                        ddfText = ddfText & """" & Folder & "\Windows\System\" & GetFileName(File) & """ """ & GetFileName(File) & """" & vbCrLf
                        iniText = iniText & "System=" & GetFileName(File) & "|" & GetFileSize(File) & "|" & GetFileDate(File) & "|" & fvi.FileVersion & vbCrLf
                    End If
                End If

            End If
            LibType = TLBLIBTYPE
        End If

    End If
End Function

Private Function CheckLibrary(ByVal Folder As String, ByVal Path As String, ByVal inVal As String, ByVal LibType As String, ByRef ddfText As String, ByRef iniText As String) As String

    If GetFileName(inVal) = "msvbvm.dll" Then Exit Function
    
    If PathExists(MapFolder(GetFilePath(Path) & "..\Binary\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\Binary\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\Binary\", GetFileName(inVal))
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", GetFileName(inVal))
        Exit Function
    End If
    
    If PathExists(MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", GetFileName(inVal))
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Windows\System\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\..\Windows\System\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\..\Windows\System\", GetFileName(inVal))
        Exit Function
    End If

    If PathExists(MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", GetFileName(inVal))
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", GetFileName(inVal)), True) Then
        GetFileEntry Folder, MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", GetFileName(inVal)), LibType, ddfText, iniText
        CheckLibrary = MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", GetFileName(inVal))
        Exit Function
    End If


        If PathExists(MapFolder("\WINDOWS\SYSTEM32\", inVal), True) Then
            GetFileEntry Folder, MapFolder("\WINDOWS\SYSTEM32\", inVal), LibType, ddfText, iniText
            CheckLibrary = MapFolder("\WINDOWS\SYSTEM32\", inVal)
            Exit Function
        ElseIf PathExists(MapFolder(GetFilePath(Path), inVal), True) Then
            GetFileEntry Folder, MapFolder(GetFilePath(Path), inVal), LibType, ddfText, iniText
            CheckLibrary = MapFolder(GetFilePath(Path), inVal)
            Exit Function
        End If

End Function

Public Function MapFolder(ByVal RootURL As String, ByVal vURL As String) As String
On Error GoTo exitthis

    'concatenates vURL to the RootURL properly by blind path specifications
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
    vURL = Replace(vURL, "\*", "")
    
    Do While InStr(vURL, "..\") > 0
        If InStrRev(Left(vURL, InStr(vURL, "..\") - 5), "\") > 0 Then
            vURL = Left(vURL, InStrRev(Left(vURL, InStr(vURL, "..\") - 5), "\")) & Mid(vURL, InStr(vURL, "..\") + 3)
        End If
    Loop
    
    Do While InStr(vURL, "../") > 0
        If InStrRev(Left(vURL, InStr(vURL, "../") - 5), "/") > 0 Then
            vURL = Left(vURL, InStrRev(Left(vURL, InStr(vURL, "../") - 5), "/")) & Mid(vURL, InStr(vURL, "../") + 3)
        End If
    Loop
    
    MapFolder = vURL
    Exit Function
exitthis:
    Err.Clear
End Function

Private Function ParseSource(ByVal Folder As String, ByVal Path As String, ByVal text As String, ByRef ddfText As String, ByRef iniText As String) As String

    Dim backup As String
    backup = text
    Dim comment As String
    
    Dim inVal As String
    Do Until text = ""
        comment = StrReverse(RemoveNextArg(text, "Lib"))
        If (InStr(comment, "'" & vbLf & vbCr) = 0) Then
            comment = ""
        ElseIf InStr(comment, vbLf & vbCr) > 0 And InStr(comment, vbLf & vbCr) <= InStr(comment, "'" & vbLf & vbCr) Then
            comment = ""
        End If
        
        If text <> "" And comment = "" Then
        
            inVal = RemoveQuotedArg(text, """", """")
            If inVal <> "" And InStr(LCase(inVal), ".dll") = 0 Then inVal = inVal & ".dll"
            CheckLibrary Folder, Path, inVal, DLLLIBTYPE, ddfText, iniText
        End If
    Loop
    
    text = backup
    Do Until text = ""
        comment = RemoveNextArg(text, "CreateObject(")
        If (InStr(comment, "'" & vbLf & vbCr) = 0) Then
            comment = ""
        ElseIf InStr(comment, vbLf & vbCr) > 0 And InStr(comment, vbLf & vbCr) <= InStr(comment, "'" & vbLf & vbCr) Then
            comment = ""
        End If
        
        If text <> "" And comment = "" Then
        
            inVal = RemoveQuotedArg(text, """", """")
            inVal = RemoveNextArg(inVal, ".")
            If inVal <> "" And InStr(LCase(inVal), ".dll") = 0 Then inVal = inVal & ".dll"
            If PathExists(MapFolder("..\Common\Binary", inVal), True) Then
                CheckLibrary Folder, Path, inVal, REGDLLLIBTYPE, ddfText, iniText
            
            End If
        End If
    Loop
End Function

Private Function ParseProject(ByVal Folder As String, ByVal Path As String, ByVal text As String, ByRef ddfText As String, ByRef iniText As String) As String

    Dim inLine As String
    Dim inVar As String
    Dim inVal As String
    Dim nxtChecks As String
    Dim tmpPath As String
    
    Do Until text = ""
        inLine = RemoveNextArg(text, vbCr)
        inVar = RemoveNextArg(inLine, "=")
        inVal = inLine
        inVar = Replace(Replace(inVar, vbCr, ""), vbLf, "")

        Select Case LCase(inVar)
            Case "object", "reference"
           ' Stop
            
                Do Until inVal = ""
                    tmpPath = CheckLibrary(Folder, Path, inVal, REGDLLLIBTYPE, ddfText, iniText)
                    If tmpPath <> "" Then
                        tmpPath = Replace(Replace(Replace(LCase(tmpPath), "binary", "projects"), ".dll", ".vbp"), ".ocx", ".vbp")
                        If PathExists(tmpPath, True) Then

                            If InStr(nxtChecks, tmpPath & vbCrLf) = 0 Then
                                nxtChecks = nxtChecks & tmpPath & vbCrLf

                            End If
                        End If
                        Exit Do
                    End If
                    
                    If LCase(inVar) = "object" Then
                        inVal = StrReverse(inVal)
                        inVar = RemoveNextArg(inVal, "#")
                        inVar = RemoveNextArg(inVar, ";")
                        inVar = StrReverse(inVar)
                        inVal = StrReverse(inVal)
                    Else
                        inVal = StrReverse(inVal)
                        inVar = RemoveNextArg(inVal, "#")
                        inVar = RemoveNextArg(inVar, "#")
                        inVar = StrReverse(inVar)
                        inVal = StrReverse(inVal)
                    End If
               '     Stop
                    tmpPath = CheckLibrary(Folder, Path, inVar, REGDLLLIBTYPE, ddfText, iniText)
                    If tmpPath <> "" Then
                        tmpPath = Replace(Replace(Replace(LCase(tmpPath), "binary", "projects"), ".dll", ".vbp"), ".ocx", ".vbp")
                        If PathExists(tmpPath, True) Then
                            If InStr(nxtChecks, tmpPath & vbCrLf) = 0 Then
                                nxtChecks = nxtChecks & tmpPath & vbCrLf

                            End If
                        End If
                    Else
                        tmpPath = CheckLibrary(Folder, Path, inVal, REGDLLLIBTYPE, ddfText, iniText)
                        If tmpPath <> "" Then
                            tmpPath = Replace(Replace(Replace(LCase(tmpPath), "binary", "projects"), ".dll", ".vbp"), ".ocx", ".vbp")
                            If PathExists(tmpPath, True) Then
                                If InStr(nxtChecks, tmpPath & vbCrLf) = 0 Then
                                    nxtChecks = nxtChecks & tmpPath & vbCrLf
    
                                End If
                            End If
                        End If
                    End If
                Loop

            Case "class", "module", "usercontrol", "form", "designer"
            
                If InStr(inVal, ";") > 0 Then RemoveNextArg inVal, ";"
                inVal = MapFolder(GetFilePath(Path), inVal)
                ParseSource Folder, inVal, ReadFile(inVal), ddfText, iniText
        End Select
    Loop
    
    Do Until nxtChecks = ""
        tmpPath = RemoveNextArg(nxtChecks, vbCrLf)
        ParseProject Folder, tmpPath, ReadFile(tmpPath), ddfText, iniText
    Loop

End Function

Private Function GetReferences(ByVal Folder As String, ByVal f As Object, ByRef ddfText As String, ByRef iniText As String) As String
    On Error GoTo catchexit
    
    Dim s As Object
    Dim e As Object
    
    For Each s In f.SubFolders
        If InStr(s.Path, "\Test") = 0 Then
            GetReferences Folder, s, ddfText, iniText
        End If
    Next
    
    For Each e In f.Files
        
        Select Case GetFileExt(e.Path)
            Case ".bas", ".cls", ".frm", ".ctl"
                ParseSource Folder, e.Path, ReadFile(e.Path), ddfText, iniText
            Case ".vbp"
                ParseProject Folder, e.Path, ReadFile(e.Path), ddfText, iniText
        End Select
    Next
    
    Dim t As Variant
    
    For Each t In Program.System32
        Select Case NextArg(t, " ")
            Case "=*"
                CheckLibrary Folder, AppPath, RemoveArg(t, " "), REGDLLLIBTYPE, ddfText, iniText
            Case "=?"
                CheckLibrary Folder, AppPath, RemoveArg(t, " "), DLLLIBTYPE, ddfText, iniText
            Case "=!"
                CheckLibrary Folder, AppPath, RemoveArg(t, " "), TLBLIBTYPE, ddfText, iniText
        End Select
    Next
    
    
    Exit Function
catchexit:
    Err.Clear
    Resume Next
End Function

Private Function GenerateAppIncludes(ByVal Folder As String, ByVal f As Object, ByRef ddfText As String, ByRef iniText As String) As String
    
    GetReferences Folder, f, ddfText, iniText
    
End Function

Public Sub GenerateSystem32(ByVal Folder As String, ByVal AppName As String, ByRef ddfText As String, ByRef iniText As String)
    Set fso = CreateObject("scripting.FileSystemObject")
    GenerateAppIncludes Folder, fso.GetFolder(Folder & "\" & AppName), ddfText, iniText
    Set fso = Nothing
End Sub
       
    
Public Sub GenerateProgram(ByVal Folder As String, ByVal AppName As String, ByRef ddfText As String, ByRef iniText As String)
    Set fso = CreateObject("scripting.FileSystemObject")
    FolderGetItems Folder, fso.GetFolder(Folder), AppName, ddfText, iniText
    Set fso = Nothing
End Sub

Private Sub FolderGetItems(ByVal Folder As String, ByRef f As Object, ByVal l As String, ByRef ddfText As String, ByRef iniText As String)
    Dim i As Object
    Dim s As Object
    Dim u As Variant
    Dim inc As Boolean
    Dim cnt As Long
    Dim ex As String
    For Each i In f.Files
        inc = True
        If Program.Excludes.count > 0 Then
            For cnt = 1 To Program.Excludes.count
                ex = Program.Excludes(cnt)

                Select Case RemoveNextArg(ex, " ")
                    Case "*"
                        If LikeCompare(LCase(Replace(i.Path, Folder & "\", "", , , vbTextCompare)), LCase(ex)) Then inc = False
                    Case "?"
                        If LCase(Replace(i.Path, Folder & "\", "", , , vbTextCompare)) = LCase(ex) Then inc = False
                End Select

                If Not inc Then Exit For
            Next
        End If
        If inc Then
            ddfText = ddfText & """" & i.Path & """ """ & i.Name & """" & vbCrLf
            For Each u In Program.ProgramFiles
                Select Case LCase(NextArg(u, "|"))

                    Case LCase("! " & Replace(i.Path, Folder & "\", "", , , vbTextCompare)), LCase("-! " & Replace(i.Path, Folder & "\", "", , , vbTextCompare))
                        iniText = iniText & "Current = " & Replace(i.Path, Folder & "\", "", , , vbTextCompare) & "|" & i.Size & "|" & GetFileDate(i.Path) & vbCrLf
                        inc = False
                        Exit For
                    Case LCase("$ " & Replace(i.Path, Folder & "\", "", , , vbTextCompare)), LCase("-$ " & Replace(i.Path, Folder & "\", "", , , vbTextCompare))
                        iniText = iniText & "AllUser = " & Replace(i.Path, Folder & "\", "", , , vbTextCompare) & "|" & i.Size & "|" & GetFileDate(i.Path) & vbCrLf
                        inc = False
                        Exit For
                End Select
            Next
            
            If inc Then
                iniText = iniText & "File = " & Replace(i.Path, Folder & "\", "", , , vbTextCompare) & "|" & i.Size & "|" & GetFileDate(i.Path) & vbCrLf
            End If
        End If
    Next
    For Each s In f.SubFolders
        inc = True
        If Program.Excludes.count > 0 Then
            For cnt = 1 To Program.Excludes.count
                ex = Program.Excludes(cnt)
                Select Case RemoveNextArg(ex, " ")
                    Case "*"
                        If LikeCompare(LCase(Replace(s.Path, Folder & "\", "", , , vbTextCompare)), LCase(ex)) Then inc = False
                    Case "?"
                        If LCase(Replace(s.Path, Folder & "\", "", , , vbTextCompare)) = LCase(ex) Then inc = False
                End Select
                If Not inc Then Exit For
            Next
        End If
        If inc Then
            ddfText = ddfText & ".Set DestinationDir = " & Replace(s.Path, Folder & "\", "", , , vbTextCompare) & vbCrLf
            iniText = iniText & "Folder = " & Replace(s.Path, Folder & "\", "", , , vbTextCompare) & "|<DIR>|" & GetFileDate(s.Path) & vbCrLf
            FolderGetItems Folder, s, l & "\" & s.Name, ddfText, iniText
        End If
    Next
    Set i = Nothing
    Set s = Nothing
End Sub
