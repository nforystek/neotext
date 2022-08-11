Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Enum ColorProperties

    ColorBackground = 1
    ColorForeGround = 2

    BatchInkComment = 3
    BatchInkCommands = 4
    BatchInkFinished = 5
    BatchInkCurrently = 6
    BatchInkIncomming = 7

    JScriptComment = 8
    JScriptStatements = 9
    JScriptOperators = 10
    JScriptVariables = 11
    JScriptValues = 12
    JScriptError = 13

    VBScriptComment = 14
    VBScriptStatements = 15
    VBScriptOperators = 16
    VBScriptVariables = 17
    VBScriptValues = 18
    VBScriptError = 19

    NSISScriptComment = 20
    NSISScriptCommands = 21
    NSISScriptEqualJump = 22
    NSISScriptElseJump = 23
    NSISScriptAboveJump = 24

    AssemblyComment = 25
    AssemblyCommand = 26
    AssemblyNotation = 27
    AssemblyRegister = 28
    AssemblyParameter = 29
    AssemblyError = 30

    
    
End Enum

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public fMain As frmMain

Public Function GetWinDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetWindowsDirectory(winDir, 255)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Trim(Dir(winDir, vbDirectory)) = "" Then winDir = App.path
    If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    GetWinDir = winDir
End Function

Public Function GetWinTempDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetTempPath(255, winDir)
    If (ret <> 16) And (ret <> 34) Then
        winDir = GetWinDir()
        If LCase(Dir(winDir & "TEMP", vbDirectory)) = "" Then
            MkDir winDir + "TEMP"
        End If
        winDir = winDir + "TEMP\"
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWinTempDir = winDir
End Function

Public Function GetTemporaryFile() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetTempFileName(GetWinTempDir, App.Title, 0, winDir)
    If ret = 0 Then
        winDir = GetWinTempDir & "\" & Left(Left(App.Title, 3) & Hex(CLng(Mid(CStr(Rnd), 3))), 14) & ".tmp"
        ret = FreeFile
        Open winDir For Output As #ret
        Close #ret
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
    End If
    GetTemporaryFile = winDir
End Function

Public Function GetUserLoginName() As String

    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        GetUserLoginName = Replace(Left$(sBuffer, lSize), Chr(0), "")
    Else
        GetUserLoginName = "NotableUser"
    End If
    
End Function

Public Function FolderQuoteName83(ByVal Folder As String) As String
    Dim inDir As String
    inDir = NextArg(Folder, "\")
    Folder = RemoveArg(Folder, "\")
    Do Until Folder = ""
        If PathExists(inDir & "\" & NextArg(Folder, "\"), False) Then
            inDir = inDir & "\" & Dir(inDir & "\" & NextArg(Folder, "\"), vbDirectory Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)
        ElseIf PathExists(inDir & "\" & NextArg(Folder, "\"), True) Then
            inDir = inDir & "\" & Dir(inDir & "\" & NextArg(Folder, "\"), vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)
        Else
            inDir = inDir & "\" & NextArg(Folder, "\")
        End If
        Folder = RemoveArg(Folder, "\")
    Loop
    FolderQuoteName83 = """" & inDir & """"
End Function

Public Sub Main()
    
    On Error Resume Next
    Set fMain = New frmMain
    fMain.Show
    If Err Then Err.Clear
    
End Sub

