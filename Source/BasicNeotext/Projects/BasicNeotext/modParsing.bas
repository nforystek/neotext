Attribute VB_Name = "modParsing"
Option Explicit

'parsing functions set the self properties and return further
'files in a vbcrlf delimited listing to be populated analyize
'exception of the parsecommand, parses out commandline switch

'
'first space delimited word in execs is the action and/or
'param key, (redundant unessisarily also exec's key thus)
'and then param of contains all parameters or text between
'the switch qith quoted arguments removed and put into path
'and marked by always one, to any number of they exists
'because vb6 does not accept multiple same switched per
'execute the environment then we wont see 2 /make in 1
'command line for example, but possible vbn sees only
'quoted parameters and more then one for no switches
'
'Execs(paramkey) = paramkey pathkey1 [pathkey2...]
'Param(paramkey) = <non quoted arguments in the switch
'Paths(pathkey1) = <first quoted argument or nothing>
'[Paths(pathkey2) = <second quoted argument if exists>]
'[...]
'
'paramkey of default is for non / or - switches
'but still found args or quoted arguments exist
'Execs("default") = default default1 [default2...]
'Param("default") = <non quoted arguments in the switch
'Paths("default1") = <first quoted argument or nothing>
'[Paths("default2") = <second quoted argument if exists>]
'[...]
'

Public Sub ParseCommand()
    Dim cmd As String
    Dim line As String
    Dim path As String
    Dim action As String
    Dim arg As String
    Dim cnt As Long
    Dim use As String
    Dim dbg As String
    cmd = Trim(Command)
    If cmd <> "" Then
        Do
            path = ""
            line = ""
            If (Left(cmd, 1) = "/") Then
                use = "/"
                cmd = Mid(cmd, 2)
            ElseIf (Left(cmd, 1) = "-") Then
                use = "-"
                cmd = Mid(cmd, 2)
            End If
            line = RemoveNextArg(cmd, use)
            action = LCase(RemoveNextArg(line, " "))
            If (InStr(line, """") > 0) Then
                arg = RemoveNextArg(line, """")
                line = """" & line
            ElseIf (InStr(line, "'") > 0) Then
                arg = RemoveNextArg(line, "'")
                line = "'" & line
            Else
                arg = line
                line = ""
            End If
            dbg = dbg & "ARGUMENTS(" & action & ":=" & arg & ")"
            If (arg <> "") Then
                Param.Add Trim(arg), action
                arg = ""
            End If
            If line = "" And cmd <> "" And action = "" Then
                line = cmd
                cmd = ""
            End If
            cnt = 0
            Do
                If (InStr(line, """") > 0) Then
                    path = RemoveQuotedArg(line, """", """")
                Else
                    path = RemoveQuotedArg(line, "'", "'")
                End If
                cnt = cnt + 1
                dbg = dbg & " PATHS(" & action & Trim(CStr(cnt)) & ":=" & path & ")"
                If (action <> "") Then
                    Paths.Add path, action & Trim(CStr(cnt))
                    arg = arg & action & Trim(CStr(cnt)) & " "
                Else
                    Paths.Add path, "default" & Trim(CStr(cnt))
                    arg = arg & "default" & Trim(CStr(cnt)) & " "
                End If
            Loop While (InStr(line, """") > 0) Or (InStr(line, "'") > 0)
            If (action <> "") Then
                If (cnt = 0) Then Paths.Add "", action & "1"
                Execs.Add action & " " & Trim(arg), action
            ElseIf Execs.count = 0 And cnt > 0 Then
                If (cnt = 0) Then Paths.Add "", Trim(CStr(cnt))
                Execs.Add "default " & Trim(arg), "default"
            End If
            dbg = "ACTION(" & action & ":=" & Trim(arg) & ") " & dbg
            If (cmd <> "") Then cmd = use & cmd
            Debug.Print dbg
            dbg = ""
        Loop Until (cmd = "")
    End If
End Sub

Public Function ParseGroup(ByRef Self As Project, ByVal URI As String) As String
    
    Self.Location = GetFilePath(URI)
    Self.Contents = ReadFile(URI)
    Dim inText As String
    Dim inLine As String
    inText = Self.Contents
    Do Until inText = ""
        inLine = RemoveNextArg(inText, vbCrLf)
        Select Case LCase(RemoveNextArg(inLine, "="))
            Case "startupproject"
                ParseGroup = MapPaths(Self.Location, inLine) & vbCrLf & ParseGroup
            Case "project"
                ParseGroup = ParseGroup & MapPaths(Self.Location, inLine) & vbCrLf
        End Select
    Loop

End Function

Public Function ParseProject(ByRef Self As Project, ByVal URI As String) As String
    
    Self.Location = GetFilePath(URI)
    
    Self.Contents = ReadFile(URI)
'    If InStr(Self.Contents, "[VERSION 6.0]" & vbCrLf) = 0 Then
'        WriteFile URI, "[VERSION 6.0]" & vbCrLf & Self.Contents
'        Self.Contents = ReadFile(URI)
'    End If
'
    Dim inText As String
    Dim inLine As String
    Dim inPath As String
    Dim inName As String
    
    inText = Self.Contents
    Do Until inText = ""
        inLine = RemoveNextArg(inText, vbCrLf)
        inName = LCase(RemoveNextArg(inLine, "="))
        Select Case inName
            Case "reference", "object"
                Do Until inLine = ""
                    If InStr(NextArg(inLine, "#"), "..\") > 0 Then
                        inLine = NextArg(inLine, "#")
                        Exit Do
                    Else
                        Select Case GetFilePath(NextArg(inLine, "#"))
                            Case "dll", "ocx", "exe", "tlb"
                                inLine = NextArg(inLine, "#")
                                Exit Do
                            Case Else
                                RemoveNextArg inLine, "#"
                        End Select
                    End If
                Loop
                If inLine <> "" Then
                    If Mid(inLine, 2, 1) = "\" And Mid(inLine, 4, 2) = ".." Or Mid(inLine, 5, 1) = ":" And (Not Left(inLine, 1) = ".") Then
                        inLine = Mid(inLine, 3)
                    End If
                    If PathExists(MapPaths(, NextArg(inLine, "#"), Self.Location), True) Then
                        inLine = MapPaths(, NextArg(inLine, "#"), Self.Location)
                    ElseIf PathExists(MapPaths(, NextArg(inLine, "#")), True) Then
                        inLine = MapPaths(, NextArg(inLine, "#"))
                    End If
                    If InStr(1, ParseProject, inLine, vbTextCompare) = 0 Then
                        ParseProject = ParseProject & inLine & vbCrLf
                    End If
                    inLine = ""
                End If
            Case "designer", "module", "class", "userdocument", "form", "relateddoc", "usercontrol", "resfile32"
                If InStr(inLine, ";") > 0 Then inLine = RemoveArg(inLine, ";")
                If PathExists(MapPaths(, inLine, Self.Location), True) Then
                    inLine = MapPaths(, inLine, Self.Location)
                ElseIf PathExists(MapPaths(, inLine), True) Then
                    inLine = MapPaths(, inLine)
                Else
                    inLine = ""
                End If
                If inLine <> "" Then
                    If InStr(1, ParseProject, inLine, vbTextCompare) = 0 Then
                        ParseProject = ParseProject & inLine & vbCrLf
                    End If
                    inLine = ""
                End If
            Case "exename32"
                Self.Compiled = MapPaths(Self.Location, RemoveQuotedArg(inLine, """", """"), inPath)
            Case "path32"
                inPath = Replace(inLine, """", "")
                Self.Compiled = MapPaths(MapPaths(Self.Location, inPath), GetFileName(Self.Compiled))
            Case "compatibleexe32"
                Self.Compiled = MapPaths(Self.Location, RemoveQuotedArg(inLine, """", """"), inPath)
            Case "condcomp"
                Self.CondComp = RemoveQuotedArg(inLine, """", """")
            Case "command32"
                Self.CmdLine = RemoveQuotedArg(inLine, """", """")
            Case "name"
                Self.Name = RemoveQuotedArg(inLine, """", """")
            Case "type"
            Case "iconform"
            Case "startup"
            Case "helpfile"
            Case "title"

            Case "helpcontextid"
            Case "description"
            Case "compatiblemode"
            Case "compcond"
            Case "majorver"
            Case "minorver"
            Case "revisionver"
            Case "autoincrementver"
            Case "serversupportfiles"
            Case "versioncomments"
            Case "versioncompanyname"
            Case "versionlegaltrademarks"
            Case "versionfiledescription"
            Case "versionlegalcopyright"
            Case "versionproductname"
            Case "versioncompatible32"
            Case "compilationtype"
            Case "optimizationtype"
            Case "favorpentiumpro(tm)"
            Case "removeunusedcontrolinfo"
            Case "codeviewdebuginfo"
            Case "noaliasing"
            Case "boundscheck"
            Case "overflowcheck"
            Case "flpointcheck"
            Case "fdivcheck"
            Case "unroundedfp"
            Case "startmode"
            Case "unattended"
            Case "threadingmodel"
            Case "retained"
            Case "threadperobject"
            Case "maxnumberofthreads"
            Case "debugstartupoption"
            Case "useexistingbrowser"
            Case "[ms transaction server]"
            Case "autorefresh"
            Case ""
            Case "[neotext]"

            

        End Select
    Loop

End Function

Public Function ParseNSIScript(ByRef Self As Project, ByVal URI As String, Optional ByRef vars As VBA.Collection) As String
    
    Static runGuid As String
    'when the hell did collections start allowing non alpha numeric only w/o spaces in keys!!!
    'NICE! no dash removal for that wich I been carried over since creatatree in my gen param
    Dim inText As String
    Dim inLine As String
    If runGuid = "" Then
        runGuid = modGuid.GUID
        Self.Location = GetFilePath(URI)
        Self.Contents = ReadFile(URI)
        inText = Self.Contents
    Else
        inText = ReadFile(URI)
    End If
    If vars Is Nothing Then
        Set vars = New VBA.Collection
        vars.Add GetFilePath(Paths("MakeNSIS")), "nsisdir"
        vars.Add "nsisdir ", runGuid
    End If
    Dim Key As String
    Do Until inText = ""
        inLine = RemoveNextArg(inText, vbCrLf)
        Select Case Trim(Replace(LCase(RemoveNextArg(inLine, " ")), vbTab, ""))
            Case "!include"
                If InStr(inLine, """") > 0 Then
                    Key = RemoveQuotedArg(inLine, """", """")
                ElseIf InStr(inLine, "'") > 0 Then
                    Key = RemoveQuotedArg(inLine, "'", "'")
                Else
                    Key = inLine
                End If
                Do
                    If InStr(Key, "${") > 0 Then
                        inLine = Key
                        Key = RemoveNextArg(inLine, "${", vbTextCompare, False)
                        If InStr(vars(runGuid), LCase(NextArg(inLine, "}", vbTextCompare, False)) & " ") > 0 Then
                            Key = Key & vars.Item(LCase(NextArg(inLine, "}", vbTextCompare, False)))
                        End If
                        RemoveNextArg inLine, "}", vbTextCompare, False
                        Key = Key & inLine
                    End If
                Loop Until InStr(Key, "${") = 0
                Key = Replace(Key, "\\", "\")
                If (GetFileExt(Key, True, True) = "nsi" _
                    Or GetFileExt(Key, True, True) = "nsh") Then
                    If PathExists(Key, True) Then
                        'Debug.Print "Include " & Key
                        ChDir GetFilePath(URI)
                        ParseNSIScript = ParseNSIScript & Key & vbCrLf
                        ParseNSIScript Self, Key, vars
                        If Self.Compiled = "" Then
                            If GetFileExt(Key, True, True) = "exe" Then
                                If PathExists(MapPaths(, Key, Self.Location), True) Then
                                    Self.Compiled = MapPaths(, Key, Self.Location)
                                ElseIf PathExists(MapPaths(, Key, GetFilePath(URI)), True) Then
                                    Self.Compiled = MapPaths(, Key, GetFilePath(URI))
                                End If
                            End If
                        End If
                    ElseIf PathExists(MapPaths(, Key, GetFilePath(URI)), True) Then
                        'Debug.Print "Include " & MapPaths(, Key, GetFilePath(URI))
                        ChDir GetFilePath(URI)
                        ParseNSIScript = ParseNSIScript & MapPaths(, Key, GetFilePath(URI)) & vbCrLf
                        ParseNSIScript Self, MapPaths(, Key, GetFilePath(URI)), vars
                        If Self.Compiled = "" Then
                            If GetFileExt(Key, True, True) = "exe" Then
                                If PathExists(MapPaths(, Key, Self.Location), True) Then
                                    Self.Compiled = MapPaths(, Key, Self.Location)
                                ElseIf PathExists(MapPaths(, Key, GetFilePath(URI)), True) Then
                                    Self.Compiled = MapPaths(, Key, GetFilePath(URI))
                                End If
                            End If
                       End If
                    End If
                End If
            Case "!define"
                Key = LCase(RemoveNextArg(inLine, " "))
                If InStr(vars(runGuid), Key & " ") > 0 Then vars.Remove Key
                If InStr(inLine, """") > 0 Then
                    vars.Add RemoveQuotedArg(inLine, """", """"), Key
                ElseIf InStr(inLine, "'") > 0 Then
                    vars.Add RemoveQuotedArg(inLine, "'", "'"), Key
                Else
                    vars.Add inLine, Key
                End If
                'Debug.Print "Var " & Key & ":=" & vars(Key)
                If InStr(vars(runGuid), Key & " ") = 0 Then
                    Key = vars(runGuid) & Key & " "
                    vars.Remove runGuid
                    vars.Add Key, runGuid
                End If
            Case "outfile"
                If InStr(inLine, """") > 0 Then
                    Key = RemoveQuotedArg(inLine, """", """")
                ElseIf InStr(inLine, "'") > 0 Then
                    Key = RemoveQuotedArg(inLine, "'", "'")
                Else
                    Key = inLine
                End If
                Do
                    If InStr(Key, "${") > 0 Then
                        inLine = Key
                        Key = RemoveNextArg(inLine, "${", vbTextCompare, False)
                        If InStr(vars(runGuid), LCase(NextArg(inLine, "}", vbTextCompare, False)) & " ") > 0 Then
                            Key = Key & vars(LCase(NextArg(inLine, "}", vbTextCompare, False)))
                        End If
                        RemoveNextArg inLine, "}", vbTextCompare, False
                        Key = Key & inLine
                    End If
                Loop Until InStr(Key, "${") = 0
                Key = Replace(Key, "\\", "\")
                If Self.Compiled = "" Then
                    If GetFileExt(Key, True, True) = "exe" Then
                        If PathExists(MapPaths(, Key, Self.Location), True) Then
                            Self.Compiled = MapPaths(, Key, Self.Location)
                        ElseIf PathExists(MapPaths(, Key, GetFilePath(URI)), True) Then
                            Self.Compiled = MapPaths(, Key, GetFilePath(URI))
                        End If
                    End If
                End If
        End Select
    Loop

End Function

