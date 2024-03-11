Attribute VB_Name = "modMain"
Option Explicit

Global Exchange As Exchange
Global Shore As Shore

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long


Private Function Large2Currency(largeInt As LARGE_INTEGER) As Currency


    If (largeInt.lowpart) > 0& Then
        Large2Currency = largeInt.lowpart
    Else
        Large2Currency = CCur(2 ^ 31) + CCur(largeInt.lowpart And &H7FFFFFFF)
    End If

    Large2Currency = Large2Currency + largeInt.highpart * CCur(2 ^ 32)
End Function



Public Static Function ElapsedCounter() As Double
    Static starting As Boolean
    Static m_CounterStart As LARGE_INTEGER
    Static m_crFrequency As Currency
    Dim lResp As Long
    starting = Not starting
    If starting Then
        QueryPerformanceFrequency m_CounterStart
        m_crFrequency = Large2Currency(m_CounterStart)
        lResp = QueryPerformanceCounter(m_CounterStart)
        starting = True
    Else
        Dim m_CounterEnd As LARGE_INTEGER
        lResp = QueryPerformanceCounter(m_CounterEnd)
        Dim crStart As Currency
        Dim crStop As Currency
        crStart = Large2Currency(m_CounterStart)
        crStop = Large2Currency(m_CounterEnd)
        ElapsedCounter = Round(((crStop - crStart) / m_crFrequency) * 1000#, 2) / 1000
    End If
End Function

Public Sub Main()

    
    Set Shore = New Shore
    Shore.Controller.ServiceName = "AcctSvcPool" 'this is how the system refers too
    Shore.Controller.Account = ".\LocalSystem" 'can be a user account and system
    Shore.Controller.Password = "*" 'this is for root accounts else use password
    Shore.Controller.DisplayName = "Account Pooling" 'this is a visual description
    Shore.Controller.Description = "Service for single use global address pooling in managed object orientation."
    Shore.Controller.AutoStart = True 'whether it starts automatically on system
    Shore.Controller.Interactive = False 'if any form will be visible set to true

    Select Case Command
        Case "install"
            Shore.Controller.Install 'this puts it into the service registry
            Set Shore = Nothing 'install should be solo, deinitialize an end
            End
        Case "uninstall"
            Shore.Controller.Uninstall 'remove us from services by ServiceName
            Set Shore = Nothing 'uninstall should be solo, deinitialize an end
            End
        Case Else
            Shore.Controller.StartService 'called to ackknowledge start events

            
            Load frmProcess

    End Select
    
End Sub

Public Sub StopService()
    'drop all objects initialized in the order they come out
    'conflict and then wont need to use End which will fault
    Set Shore = Nothing
End Sub


Public Function PathExists(ByVal URL As String, Optional ByVal IsFile As Variant = Empty) As Boolean
    Dim Ret As Boolean
    
    If Left(URL, 2) = "\\" Then GoTo altcheck
    If Left(LCase(URL), 7) = "file://" Then
        URL = Replace(Mid(URL, 8), "|/", ":/")
    End If
    If (Len(URL) = 2) And (Mid(URL, 2, 1) = ":") Then
        URL = URL & "\"
    End If
        
    On Error GoTo altcheck

    URL = Replace(URL, "/", "\")
    If InStr(Mid(URL, 3), ":") > 0 Or InStr(Mid(URL, 3), "?") > 0 _
        Or InStr(Mid(URL, 3), """") > 0 Or InStr(Mid(URL, 3), "<") > 0 _
         Or InStr(Mid(URL, 3), ">") > 0 Or InStr(Mid(URL, 3), "|") > 0 Then
        PathExists = False
    ElseIf Len(URL) > 2 Then
        If Len(URL) <= 3 And Mid(URL, 2, 1) = ":" Then
            If VBA.TypeName(IsFile) = "Empty" Then
                PathExists = (Dir(URL, vbVolume) <> "") Or (Dir(URL & "\*") <> "")
            Else
                PathExists = ((Dir(URL, vbVolume) <> "") Or (Dir(URL & "\*") <> "")) And (Not IsFile)
            End If
        Else
            Do While Right(URL, 1) = "\"
                URL = Left(URL, Len(URL) - 1)
            Loop
            Dim attr As Long
            Dim chk1 As String
            Do
                If VBA.TypeName(IsFile) = "Empty" Then
                    chk1 = Dir(URL, attr)
                    If chk1 <> "" And Not Ret Then
                        If InStr(URL, "*") > 0 Then
                            Ret = True
                        Else
                            If Len(URL) > Len(chk1) Then
                                Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                            Else
                                Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                            End If
                        End If
                    End If
                    If Not Ret Then
                        chk1 = Dir(URL, attr + vbDirectory)
                        If chk1 <> "" Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                                If Ret Then
                                    If Not (GetAttr(URL) And vbDirectory) = vbDirectory Then
                                        Ret = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If Not IsFile Then
                        chk1 = Dir(URL, attr + vbDirectory)
                        If chk1 <> "" And Not Ret Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                                If Ret Then
                                    If Not (GetAttr(URL) And vbDirectory) = vbDirectory Then
                                        Ret = False
                                    End If
                                End If
                            End If
                        End If
                    Else
                        chk1 = Dir(URL, attr)
                        If chk1 <> "" And Not Ret Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = (LCase(Right(URL, Len(chk1))) = LCase(chk1))
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                            End If
                        End If
                    End If
                End If
                Select Case attr
                    Case vbNormal
                        attr = vbSystem
                    Case vbSystem
                        attr = vbHidden
                    Case vbHidden
                        attr = vbReadOnly
                    Case vbReadOnly
                        attr = vbHidden + vbReadOnly
                    Case vbHidden + vbReadOnly
                        attr = vbHidden + vbSystem
                    Case vbHidden + vbSystem
                        attr = vbHidden + vbSystem + vbReadOnly
                    Case vbHidden + vbSystem + vbReadOnly
                        attr = vbSystem + vbReadOnly
                    Case vbSystem + vbReadOnly
                        attr = vbNormal
                End Select
            Loop Until Ret Or attr = vbNormal
            PathExists = Ret
        End If
    End If

    Exit Function
altcheck:

    Select Case Err.Number
        Case 55, 58, 70
            PathExists = True
        Case Else '53, 52
            Err.Clear
    End Select
'55 File already open
'58 File already exists
'70 Permission denied
'52 Bad file name or number
'53 File not found

    On Error GoTo fixthis:

    If (URL = vbNullString) Then
        PathExists = False
        Exit Function
    ElseIf (Not IsEmpty(IsFile)) Then
        If ((GetFilePath(URL) = vbNullString) And IsFile And (Not (URL = vbNullString))) Or ((GetFileName(URL) = vbNullString) And (Not IsFile) And (Not (URL = vbNullString))) Then
            PathExists = False
            Exit Function
        End If
    End If
    
    On Error GoTo 0
    On Error GoTo -1
    On Error Resume Next
    
    Dim Alt As Integer
    Alt = GetAttr(URL)
    If Err.Number = 0 Then
        If (IsEmpty(IsFile)) Then
            PathExists = True
        Else
            PathExists = IIf(IsFile, Not CBool(((Alt And vbDirectory) = vbDirectory)), CBool(((Alt And vbDirectory) = vbDirectory)))
        End If
        Exit Function
    End If
    
fixthis:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 55, 58, 70
                PathExists = True
            Case Else
                PathExists = False
        End Select
        Err.Clear
    End If
End Function

Public Function SysPath() As String
    Dim winDir As String
    Dim Ret As Long
    winDir = String(45, Chr(0))
    Ret = GetSystemDirectory(winDir, 45)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    SysPath = winDir
End Function

Public Function GetFilePath(ByVal URL As String) As String
    Dim nFolder As String
    If InStr(URL, "/") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "/") - 1)
        If nFolder = "" Then nFolder = "/"
    ElseIf InStr(URL, "\") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "\") - 1)
        If nFolder = "" Then nFolder = "\"
    Else
        nFolder = ""
    End If
    GetFilePath = nFolder
End Function
Public Function GetFileName(ByVal URL As String) As String
    If InStr(URL, "/") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "/") + 1)
    ElseIf InStr(URL, "\") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "\") + 1)
    Else
        GetFileName = URL
    End If
End Function
