Attribute VB_Name = "modFolders"
#Const modFolders = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public Enum MatchFlags
    FindAll = -1
    FirstOnly = 0
    ExactMatch = 1
End Enum

Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
                              (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                              pIdl As ITEMIDLIST) As Long

Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pV As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, ByVal _
lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200

Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const ILD_TRANSPARENT = &H1

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long

Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long

Private Function AllShort(ByVal CheckPath As String) As Boolean
    If Len(GetFileTitle(CheckPath)) <= 8 Then
        Do
            CheckPath = GetFilePath(CheckPath)
            If (Not (Len(GetFileName(CheckPath)) <= 8)) Then Exit Function
        Loop Until CheckPath = ""
    Else
        Exit Function
    End If
    AllShort = True
End Function
Public Function GetShortPath(ByVal LongPath As String) As String
    If (Not AllShort(LongPath)) Or (InStr(LongPath, " ") > 0) Then
        Dim sShortFile As String * 67
        Dim lRet As Long
        lRet = GetShortPathName(LongPath, sShortFile, Len(sShortFile))
        LongPath = Left(sShortFile, lRet)
    End If
    GetShortPath = LongPath
End Function
Public Function GetLongPath(ByRef ShortPath As String) As String
    Dim fn As String
    If CountWord(ShortPath, """") >= 2 Then
         fn = RemoveQuotedArg(ShortPath)
    Else
        fn = ShortPath
    End If
    If InStr(fn, "~") > 0 Then
    
        If (InStr(fn, "~") - InStrRev(fn, "\", InStr(fn, "~")) <= 8) And (IsNumeric(NextArg(Mid(fn, InStr(fn, "~") + 1), ".")) Or IsNumeric(NextArg(Mid(fn, InStr(fn, "~") + 1), "\"))) Then
        
            fn = FolderQuoteName83(fn)
            If CountWord(fn, """") = 2 Then
                fn = RemoveQuotedArg(fn)
            End If
            
        End If
    End If
    GetLongPath = fn
End Function


Public Function FolderQuoteName83(ByVal ShortPath As String) As String
  
    Dim inDir As String
    inDir = NextArg(ShortPath, "\")
    ShortPath = RemoveArg(ShortPath, "\")
    Do Until ShortPath = ""
        If PathExists(inDir & "\" & NextArg(ShortPath, "\"), False) Then
            inDir = inDir & "\" & Dir(inDir & "\" & NextArg(ShortPath, "\"), vbDirectory Or vbHidden Or vbSystem)
        ElseIf PathExists(inDir & "\" & NextArg(ShortPath, "\"), True) Then
            inDir = inDir & "\" & Dir(inDir & "\" & NextArg(ShortPath, "\"), vbNormal Or vbHidden Or vbSystem)
        Else
            inDir = inDir & "\" & NextArg(ShortPath, "\")
        End If
        ShortPath = RemoveArg(ShortPath, "\")
    Loop
    FolderQuoteName83 = """" & inDir & """"
End Function

Public Sub RemovePath(ByVal Path As String, Optional ByRef FolderList As String)
    'deletes PATH, and all files and fodlers under PATH
    Dim nxt As String
    On Error Resume Next
    nxt = Dir(Path & "\*", vbDirectory)
    If Err Then
        Err.Clear
        Kill Path
    Else
        Do Until nxt = ""
            If Not nxt = "." And Not nxt = ".." Then
                FolderList = FolderList & Path & "\" & nxt & vbCrLf
            End If
            nxt = Dir
        Loop
    End If
    Do Until FolderList = ""
        RemovePath RemoveNextArg(FolderList, vbCrLf), FolderList
    Loop
    RmDir Path
End Sub

Public Function SearchPath(ByRef FindText As String, Optional ByVal Recursive As Integer = -1, Optional ByVal RootPath As String, Optional ByVal MatchFlag As MatchFlags, Optional ByRef FolderList As String, Optional ByVal Flags As Long = vbDirectory Or vbNormal Or vbSystem Or vbHidden) As String
    If RootPath = "" Then RootPath = Left(CurDir, 2) & "\"
   ' Debug.Print Replace(RootPath & "\", "\\", "\")
    
    Dim nxt As String
    nxt = FindText
    If (nxt <> "") Then
        Do Until (nxt = "")
            If (MatchFlag And ExactMatch) = ExactMatch Then
                If LCase(GetFileName(RootPath)) Like LCase(NextArg(nxt, vbCrLf)) Or LCase(GetFileName(RootPath)) = LCase(NextArg(nxt, vbCrLf)) Then
                    SearchPath = SearchPath & RootPath & vbCrLf
                    If (MatchFlag = FirstOnly) Then FindText = Replace(Replace(FindText, NextArg(nxt, vbCrLf) & vbCrLf, ""), NextArg(nxt, vbCrLf), "")
                End If
            Else
                If InStr(1, LCase(RootPath), LCase(NextArg(nxt, vbCrLf)), vbTextCompare) > 0 Or LCase(RootPath) Like LCase(NextArg(nxt, vbCrLf)) Then
                    SearchPath = SearchPath & RootPath & vbCrLf
                    If (MatchFlag = FirstOnly) Then FindText = Replace(Replace(FindText, NextArg(nxt, vbCrLf) & vbCrLf, ""), NextArg(nxt, vbCrLf), "")
                End If
            End If
            RemoveNextArg nxt, vbCrLf
        Loop
    End If
    If (FindText <> "") Then
        On Error Resume Next
        'If Dir(RootPath, vbDirectory Or vbNormal Or vbSystem Or vbHidden) = "" Then

            nxt = Dir(Replace(RootPath & "\", "\\", "\"), Flags)

            If Err Then
                Err.Clear
            ElseIf Abs(Recursive) > 0 Then
                If Recursive > 0 Then Recursive = Recursive - 1
                Do Until nxt = ""
                    If (Not (nxt = ".")) And (Not (nxt = "..")) Then
                        FolderList = FolderList & RootPath & "\" & nxt & vbCrLf
                    End If
                    nxt = Dir
                Loop
            End If
        'End If
        
        Do Until (FolderList = "") Or (FindText = "")
            nxt = RemoveNextArg(FolderList, vbCrLf)
            SearchPath = SearchPath & SearchPath(FindText, Recursive, nxt, MatchFlag, , Flags)
        Loop

    End If
    SearchPath = Replace(SearchPath, "\\", "\")
End Function


Public Function GetMyDocumentsFolder() As String
    Dim BI As BROWSEINFO
    Dim nFolder As Long
    Dim IDL As ITEMIDLIST
    Dim sPath As String
    With BI
        nFolder = CSIDL_PERSONAL
        If SHGetSpecialFolderLocation(ByVal 0&, ByVal nFolder, IDL) = NOERROR Then
            .pidlRoot = IDL.mkid.cb
        End If
        .pszDisplayName = String$(MAX_PATH, 0)

        .ulFlags = BIF_RETURNONLYFSDIRS

    End With

    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath

    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)

    GetMyDocumentsFolder = sPath
End Function

Public Function GetFavoritesFolder()

    Dim lID As ITEMIDLIST
    Dim sPath As String
            
    If SHGetSpecialFolderLocation(0&, CSIDL_FAVORITES, lID) = NOERROR Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal lID.mkid.cb, ByVal sPath
        If InStr(sPath, vbNullChar) > 0 Then
            sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
        End If
    End If
    sPath = Replace(sPath, vbNullChar, "")
    If Not Right(sPath, 1) = "\" Then sPath = sPath & "\"
    
    GetFavoritesFolder = sPath
    
End Function

Public Function GetSystem32Folder() As String
    Static winDir As String
    If winDir = "" Then
        Dim Ret As Long
        winDir = String(45, Chr(0))
        Ret = GetSystemDirectory(winDir, 45)
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetSystem32Folder = winDir
End Function

Public Function GetWindowsFolder() As String
    Static winDir As String
    If winDir = "" Then
        Dim Ret As Long
        winDir = String(260, Chr(0))
        Ret = GetWindowsDirectory(winDir, 260)
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWindowsFolder = winDir
End Function

Public Function GetWindowsTempFolder(Optional ByVal UseWin As Boolean = False) As String

    Dim winDir As String
    Dim Ret As Long
    winDir = String(260, Chr(0))
    Ret = GetTempPath(260, winDir)
    If (Not ((Ret = 16) And UseWin)) And (Not ((Ret = 34) And Not UseWin)) Then
        If PathExists(GetWindowsFolder() + "TEMP") Then
            winDir = GetWindowsFolder() + "TEMP\"
        Else
            On Error Resume Next
            MkDir GetWindowsFolder() + "TEMP"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            
            If PathExists(GetWindowsFolder() + "TEMP") Then
                winDir = GetWindowsFolder() + "TEMP\"
            Else
                winDir = ""
            End If
        End If
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWindowsTempFolder = winDir
End Function

Public Function GetAllUserProfileFolder() As String

    Dim sAllUser As String
    sAllUser = String(255, Chr(0))
    GetAllUsersProfileDirectory sAllUser, 255
    GetAllUserProfileFolder = Replace(sAllUser, Chr(0), "")
    
End Function
Public Function GetDefaultUserProfileFolder() As String
    
    Dim sDefault As String
    sDefault = String(255, Chr(0))
    GetDefaultUserProfileDirectory sDefault, 255
    GetDefaultUserProfileFolder = Replace(sDefault, Chr(0), "")
    
End Function
Public Function GetCurrentUserProfileFolder() As String

    Dim hToken As Long
    Dim sLibrary As String
    sLibrary = String(255, Chr(0))
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sLibrary, 255
    GetCurrentUserProfileFolder = Replace(sLibrary, Chr(0), "")
    
End Function


'2 C:\Documents and Settings\<User>\Start Menu\Programs
'5 C:\Documents and Settings\<User>\My Documents
'6 C:\Documents and Settings\<User>\Favorites
'7 C:\Documents and Settings\<User>\Start Menu\Programs\Startup
'8 C:\Documents and Settings\<User>\Recent
'9 C:\Documents and Settings\<User>\SendTo
'11 C:\Documents and Settings\<User>\Start Menu
'13 C:\Documents and Settings\<User>\My Documents\My Music
'16 C:\Documents and Settings\<User>\Desktop
'19 C:\Documents and Settings\<User>\NetHood
'20 C:\WINDOWS\Fonts
'21 C:\Documents and Settings\<User>\Templates
'22 C:\Documents and Settings\All Users\Start Menu
'23 C:\Documents and Settings\All Users\Start Menu\Programs
'24 C:\Documents and Settings\All Users\Start Menu\Programs\Startup
'25 C:\Documents and Settings\All Users\Desktop
'26 C:\Documents and Settings\<User>\Application Data
'27 C:\Documents and Settings\<User>\PrintHood
'28 C:\Documents and Settings\<User>\Local Settings\Application Data
'31 C:\Documents and Settings\All Users\Favorites
'32 C:\Documents and Settings\<User>\Local Settings\Temporary Internet Files
'33 C:\Documents and Settings\<User>\Cookies
'34 C:\Documents and Settings\<User>\Local Settings\History
'35 C:\Documents and Settings\All Users\Application Data
'36 C:\WINDOWS
'37 C:\WINDOWS\system32
'38 C:\Program Files
'39 C:\Documents and Settings\<User>\My Documents\My Pictures
'40 C:\Documents and Settings\<User>
'43 C:\Program Files\Common Files
'45 C:\Documents and Settings\All Users\Templates
'46 C:\Documents and Settings\All Users\Documents
'47 C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
'48 C:\Documents and Settings\<User>\Start Menu\Programs\Administrative Tools
'53 C:\Documents and Settings\All Users\Documents\My Music
'54 C:\Documents and Settings\All Users\Documents\My Pictures
'55 C:\Documents and Settings\All Users\Documents\My Videos
'56 C:\WINDOWS\Resources
'59 C:\Documents and Settings\<User>\Local Settings\Application Data\Microsoft\CD Burning
'32782 C:\Documents and Settings\<User>\My Documents\My Videos
'32825 C:\WINDOWS\Resources\0409
'
'Private Function GetFolderValue(tAction As Long) As Long
'
'    Dim wIdx As Integer
'    wIdx = CInt(tAction)
'
'    If wIdx< 2 Then
'        GetFolderValue = 0
'
'    ElseIf wIdx< 12 Then
'        GetFolderValue = wIdx
'
'    Else
'        GetFolderValue = wIdx + 4
'    End If
'
'End Function
'
'Public Function GetFolderByAction(ByVal BrowseAction As Long) As String
'    Dim BI As BROWSEINFO
'    Dim nFolder As Long
'    Dim IDL As ITEMIDLIST
'    Dim sPath As String
'    With BI
'        .hOwner = frm.hwnd
'
'       'nFolder = GetFolderValue(BrowseAction)
'
'        If SHGetSpecialFolderLocation(ByVal 0&, ByVal BrowseAction, IDL) = NOERROR Then
'            .pidlRoot = IDL.mkid.cb
'        End If
'
'        .pszDisplayName = String$(MAX_PATH, 0)
'
'        .lpszTitle = Chr(0)
'
'        .ulFlags = BIF_RETURNONLYFSDIRS
'
'    End With
'
'    sPath = String$(MAX_PATH, 0)
'    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
'
'    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
'
'    GetFolderByAction = sPath
'End Function


Public Function GetStartMenuProgramsFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 2, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetStartMenuProgramsFolder = sPath
End Function

Public Function GetDesktopFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 0, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetDesktopFolder = sPath
End Function

Public Function GetTemporaryFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
    
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 32, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetTemporaryFolder = sPath
End Function

Public Function GetProgramFilesFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
    
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 38, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetProgramFilesFolder = sPath
End Function

Public Function GetCommonFilesFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 43, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetCommonFilesFolder = sPath
End Function
'35
Public Function GetCurrentAppDataFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
    
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 26, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetCurrentAppDataFolder = sPath
End Function

Public Function GetAllUsersAppDataFolder() As String
    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Static sPath As String
    If sPath = "" Then
    
        With BI
            If SHGetSpecialFolderLocation(ByVal 0&, ByVal 35, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
            .pszDisplayName = String$(MAX_PATH, 0)
            .lpszTitle = Chr(0)
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    GetAllUsersAppDataFolder = sPath
End Function

Public Function MapPaths(Optional ByVal inRoot As String = "", Optional ByVal inRelative As String = "", Optional ByVal inReroot As String = "") As String
    Dim outPath As String
    Dim outBlanket As Boolean
    Dim outSplit As Boolean
    
    outPath = Replace(inRelative, """", "")
    If (Not Mid(inRoot, 2, 2) = ":\") And (inRoot = "") Then inRoot = IIf(inReroot <> "", inReroot, inRoot)
    If Not Right(inRoot, 1) = "\" Then inRoot = inRoot & "\.\"
    If (Not Left(inRoot, 1) = "\") And (Not Mid(inRoot, 2, 2) = ":\") Then inRoot = ".\" & inRoot
    If (Mid(inReroot, 2, 2) = ":\") Then
        outBlanket = True
        If (Not (Right(inReroot, 1) = "\")) Then inReroot = inReroot & "\.\"
        If (Not (Left(inReroot, 1) = "\")) And (Not Mid(inReroot, 2, 2) = ":\") Then inReroot = ".\" & inReroot
        outPath = inRoot & outPath
    ElseIf (inReroot <> "") Then
        outSplit = True
        If (Not (Right(inReroot, 1) = "\")) Then inReroot = inReroot & "\.\"
        If (Not (Left(inReroot, 1) = "\")) And (Not Mid(inReroot, 2, 2) = ":\") Then inReroot = ".\" & inReroot
        outPath = inRoot & inReroot & outPath
    Else
        outPath = inRoot & outPath
    End If
    Dim pathRight As String
    Dim pathLeft As String
    Do While (InStr(outPath, ".\") > 0)
        pathLeft = RemoveNextArg(outPath, ".\")
        pathRight = outPath
        outPath = ""
        If (Right(pathLeft, 1) = "\") Then
            pathLeft = Left(pathLeft, Len(pathLeft) - 1)
        ElseIf (Right(pathLeft, 2) = "\.") Then
            pathLeft = GetFilePath(Left(pathLeft, Len(pathLeft) - 2))
        Else
            pathLeft = GetFilePath(pathLeft)
        End If
        If (Len(pathLeft) <= 1) And (inReroot <> "") And outBlanket Then
            pathLeft = GetFilePath(inReroot)
            inReroot = ""
        ElseIf (Right(pathLeft, 2) = "\.") Then
            pathLeft = GetFilePath(Left(pathLeft, Len(pathLeft) - 2))
        ElseIf (Right(pathLeft, 1) = "\.") Then
            pathLeft = GetFilePath(Left(pathLeft, Len(pathLeft) - 2))
        End If
        outPath = pathLeft & "\" & pathRight
    Loop
    If outBlanket Then
        outPath = Replace(outPath, inRoot, inReroot, , , vbTextCompare)
        
        If Left(outPath, 1) = "\" And Mid(inReroot, 2, 1) = ":" Then outPath = Left(inReroot, 2) & outPath
    Else
        If Left(outPath, 1) = "\" And Mid(inRoot, 2, 1) = ":" Then outPath = Left(inRoot, 2) & outPath
    End If
    MapPaths = outPath
End Function

Public Function MakeFolder(ByRef Path As String)
    On Error Resume Next
    If InStr(Path, "\") > 0 Then
        GetAttr Left(Path, InStrRev(Path, "\") - 1)
        If Err.Number = 76 Or Err.Number = 53 Then
            Err.Clear
            MakeFolder = Path
            Path = MakeFolder(Left(Path, InStrRev(Path, "\") - 1))
        Else
            MakeFolder = Path
        End If
    End If
    If Err.Number = 0 Then
        GetAttr MakeFolder
        If Err.Number = 76 Or Err.Number = 53 Then
            Err.Clear
            On Error GoTo -1
            MkDir MakeFolder
        End If
    End If
End Function








