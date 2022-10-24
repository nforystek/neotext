VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pIncludes As VBA.Collection
Private pLocation As String
Private pContents As String
Private pCompiled As String
Private pCondComp As String
Private pCmdLine As String
Private pName As String

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal RHS As String)
    pName = RHS
End Property
Public Property Get CmdLine() As String
    CmdLine = pCmdLine
End Property
Friend Property Let CmdLine(ByVal RHS As String)
    pCmdLine = RHS
End Property
Public Property Get Includes() As VBA.Collection
    Set Includes = pIncludes
End Property
Public Property Get Location() As String
    Location = pLocation
End Property
Friend Property Let Location(ByVal RHS As String)
    pLocation = RHS
End Property
Public Property Get Contents() As String
    Contents = pContents
End Property
Friend Property Let Contents(ByVal RHS As String)
    pContents = RHS
End Property
Public Property Get Compiled() As String
    Compiled = pCompiled
End Property
Friend Property Let Compiled(ByVal RHS As String)
    pCompiled = RHS
End Property
Public Property Get CondComp() As String
    If pCondComp <> "" Then
        If GetFileExt(pLocation, True, True) = "nsi" Then
            CondComp = pCondComp
        Else
            CondComp = "/d " & pCondComp
        End If
    End If
End Property
Friend Property Let CondComp(ByVal RHS As String)
    Dim ret As String
    If RHS = "" Then
        ret = ""
    Else
        ret = pCondComp
        Dim invar As String
        Dim inval As String
        Do Until RHS = ""
            If IsNSIS Then
                invar = RemoveNextArg(RHS, " ")
                ret = ret & "/" & Trim(invar) & " "
            Else
                inval = RemoveNextArg(RHS, ":")
                invar = RemoveNextArg(inval, "=")
                If invar <> "" And inval <> "" Then
                    If InStr(ret, UCase(invar) & "=") > 0 Then
                        ret = NextArg(ret, UCase(invar) & "=") & _
                                UCase(invar) & "=" & Trim(invar) & _
                                RemoveArg(RemoveArg(ret, UCase(invar) & "="), ":")
                    Else
                        ret = ret & UCase(invar) & "=" & Trim(inval) & ":"
                    End If
                End If
            End If
        Loop
        If Len(ret) > 0 Then ret = Left(ret, Len(ret) - 1)
    End If
    pCondComp = ret
End Property

Private Property Get IsNSIS() As Boolean
    IsNSIS = (GetFileExt(pLocation, True, True) = "nsi" Or GetFileExt(pLocation, True, True) = "nsh")
End Property

'Public Function NewEnum() As IUnknown
'    Set NewEnum = pIncludes.[_NewEnum]
'End Function

Public Sub Populate(ByVal URI As String)

    On Error GoTo failit
    On Local Error GoTo failit
    
    Reset
    
    If Not URI = "" Then
        
        Dim i As Project
        If PathExists(URI, False) Then
            pLocation = URI
            Dim fldr As String
            Dim fldrs As String
            fldrs = SearchPath("*", 1, URI, True)
            Do Until fldrs = ""
                fldr = RemoveNextArg(fldrs, vbCrLf)
                If (fldr <> URI) Then
                    Set i = New Project
                    pContents = pContents & fldr & vbCrLf
                    i.Populate fldr
                    pIncludes.Add i
                    Set i = Nothing
                End If
            Loop
            Select Case GetFileExt(pLocation, True, True)
                Case "vbp"
        

            End Select
        ElseIf PathExists(URI, True) Then
            Select Case GetFileExt(URI, True, True)
                Case "ocx", "dll", "exe", "tlb", "cab"
                    pCompiled = URI
                    pLocation = GetFilePath(URI)


                Case "vbg"
                    fldrs = ParseGroup(Me, URI)
                    Do Until fldrs = ""
                        fldr = RemoveNextArg(fldrs, vbCrLf)
                        If PathExists(fldr, True) Then
                            Set i = New Project
                            i.Populate fldr
                            If Compiled = "" Then Compiled = i.Compiled
                            pIncludes.Add i
                            Set i = Nothing
                        End If
                    Loop
                Case "bas", "ctl", "frm", "cls", "dsr", "dob", "pag"
                    pLocation = URI
                    pContents = ReadFile(URI)
                    pName = NextArg(RemoveArg(pContents, "Attribute VB_Name = """), """")
                Case "vbp"
                    
                    fldrs = ParseProject(Me, URI)
                    
                    Do Until fldrs = ""
                        fldr = RemoveNextArg(fldrs, vbCrLf)
                        If PathExists(fldr, True) Then
                            Set i = New Project
                            i.Populate fldr
                            pIncludes.Add i
                            Set i = Nothing
                        End If
                    Loop
                    pLocation = URI
                
                    fldrs = SearchPath("*.tmp", 1, GetFilePath(URI), FindAll)
                    Do Until fldrs = ""
                        fldr = RemoveNextArg(fldrs, vbCrLf)
                        If PathExists(fldr, True) Then
                            On Error Resume Next
                            Kill fldr
                            On Error GoTo 0
                        End If
                    Loop
                    fldrs = SearchPath("*.obj", 1, GetFilePath(URI), FindAll)
                    Do Until fldrs = ""
                        fldr = RemoveNextArg(fldrs, vbCrLf)
                        If PathExists(fldr, True) Then
                            On Error Resume Next
                            Kill fldr
                            On Error GoTo 0
                        End If
                    Loop
                    If PathExists(Compiled, True) And LCase(GetFilePath(Compiled)) <> LCase(GetFilePath(URI)) Then
                        fldrs = SearchPath("*.tmp", 1, GetFilePath(Compiled), FindAll)
                        Do Until fldrs = ""
                            fldr = RemoveNextArg(fldrs, vbCrLf)
                            If PathExists(fldr, True) Then
                                On Error Resume Next
                                Kill fldr
                                On Error GoTo 0
                            End If
                        Loop
                        fldrs = SearchPath("*.obj", 1, GetFilePath(Compiled), FindAll)
                        Do Until fldrs = ""
                            fldr = RemoveNextArg(fldrs, vbCrLf)
                            If PathExists(fldr, True) Then
                                On Error Resume Next
                                Kill fldr
                                On Error GoTo 0
                            End If
                        Loop
                    End If
                    
                Case "nsi", "nsh"
                    Dim v As VBA.Collection
                    fldrs = ParseNSIScript(Me, URI, v)
                    Do Until fldrs = ""
                        fldr = RemoveNextArg(fldrs, vbCrLf)
                        If PathExists(fldr, True) Then
                            Set i = New Project
                            ParseNSIScript i, fldr, v
                            pIncludes.Add i
                            Set i = Nothing
                        End If
                    Loop
                    Do Until v.count = 0
                        v.Remove 1
                    Loop
                    Set v = Nothing
                    pLocation = URI
                Case "bat"
                Case Else
                    pLocation = URI
            End Select
        End If
    End If
        
    Cleanup
    
    Dim cc As Variant
    
    For Each cc In Execs
        If (Left(LCase(cc), 4) = "out ") Or (Left(LCase(cc), 7) = "outdir ") Then
            Projs.CmdLine = Trim("/" & cc & " " & Param(cc) & " " & Projs.CmdLine)
        ElseIf (Left(LCase(cc), 2) = "c ") Or (Left(LCase(cc), 4) = "cmd ") Then
            Projs.CmdLine = Trim(Projs.CmdLine & " /" & cc & " " & Param(cc))
        ElseIf (Left(LCase(cc), 1) = "d") Then
            If IsNSIS Then
                Projs.CondComp = cc
            ElseIf Left(LCase(cc), 2) = "d " Then
                Projs.CondComp = Param("d")
            End If
        
        End If
    Next
    
    Exit Sub
failit:
'    If Err.Number = 70 Then
'        Err.Clear
'        DoEvents
'        Resume
'    Else
        Err.Clear
'    End If
End Sub

Public Sub Cleanup()
    
    Dim fldr As String
    Dim fldrs As String
    Select Case GetFileExt(pLocation, True, True)
        Case "vbp"
            If PathExists(pLocation, False) Then
                fldrs = SearchPath("*.tmp", 1, pLocation, FindAll)
                Do Until fldrs = ""
                    fldr = RemoveNextArg(fldrs, vbCrLf)
                    On Error Resume Next
                    Kill fldr
                    On Error GoTo 0
                Loop
            ElseIf PathExists(pLocation, True) Then
                fldrs = SearchPath("*.tmp", 1, GetFilePath(pLocation), FindAll)
                Do Until fldrs = ""
                    fldr = RemoveNextArg(fldrs, vbCrLf)
                    On Error Resume Next
                    Kill fldr
                    On Error GoTo 0
                Loop
            End If
            If PathExists(GetFilePath(pCompiled), False) Then
                fldrs = SearchPath("*.obj", -1, GetFilePath(pCompiled), FindAll)
                Do Until fldrs = ""
                    fldr = RemoveNextArg(fldrs, vbCrLf)
                    On Error Resume Next
                    Kill fldr
                    On Error GoTo 0
                Loop
            End If
    End Select

End Sub

Public Sub Reset()
    Do Until pIncludes.count = 0
        pIncludes.Remove 1
    Loop
    pLocation = ""
    pContents = ""
    pCompiled = ""
End Sub

Private Sub Class_Initialize()
    Set pIncludes = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Cleanup
    Set pIncludes = Nothing
End Sub

