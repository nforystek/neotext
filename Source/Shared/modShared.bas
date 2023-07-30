Attribute VB_Name = "modShared"



#Const modShared = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Const SecurityID = "nt1783405269"

Public Const MaxFileName = "maxftp.exe"
Public Const ServiceFileName = "maxservice.exe"
Public Const DBFileName = "maxftp.mdb"
Public Const MaxIDEFileName = "maxide.exe"
Public Const MaxUtilityFileName = "maxutility.exe"

Public Const MaxServiceName = "MaxFTPService"
Public Const MaxServiceDisplay = "Max-FTP Schedules"

Public Const NeoTextWebSite = "http://www.neotext.org"

Public Const FTPSiteExt = ".mftp"
Public Const MaxDBBackupExt = ".madb"
Public Const MaxProjectExt = ".mprj"
Public Const MaxScriptExt = ".mscr"
Public Const WebSiteExt = ".htm"
Public Const URLSiteExt = ".url"

Public Const ServiceFormCaption = "Max-FTP Schedule Service"
Public Const MaxMainFormCaption = "Max-FTP Application Server"
Public Const MaxIDEFormCaption = "Max-IDE"

Public Const ActiveAppFolder = "ActiveApp\"
Public Const FavoritesFolder = "Favorites\"
Public Const GraphicsFolder = "Graphics\"
Public Const ProjectFolder = "Projects\"
Public Const TemplatesFolder = "Projects\Templates\"

Public Const HelpFolder = "Help\"
Public Const TempFolder = "Temp\"


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

#If (Not modCommon) Then
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const NOERROR = 0
Private Const MAX_PATH = 260
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

Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pIdl As ITEMIDLIST) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long




'Scopes
Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_ENUM_ALL = &HFFFF
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

'The order of images in image list are based off these values
Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEDISPLAYTYPE_FILE = &H4
Private Const RESOURCEDISPLAYTYPE_GROUP = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT = &H7
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
Private Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9
Private Const RESOURCEDISPLAYTYPE_TREE = &HA
Private Const RESOURCEDISPLAYTYPE_NDSCONTAINER = &HB


Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_ALL = &H0
Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const ERROR_NO_MORE_ITEMS = 259

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type
Private Declare Function WNetOpenEnum Lib "mpr" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr" (ByVal hEnum As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long


Public Function DatabaseFilePath() As String
    Static retVal As String
    If retVal = "" Then
        retVal = AppPath & DBFileName
        If PathExists(Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare), True) And _
            (Not PathExists(retVal, True)) Then
            retVal = Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare)
        End If
    End If
    DatabaseFilePath = retVal
End Function

'Public Function HelpFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        retVal = GetProgramFilesFolder & "\Help\"
'    End If
'    HelpFolder = retVal
'End Function
'
'Public Function TempFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        retVal = GetTemporaryFolder & "\"
'    End If
'    TempFolder = retVal
'End Function
'
'Public Function ActiveAppFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        If PathExists(AppPath & "ActiveApp", False) Then
'            retVal = AppPath & "ActiveApp\"
'        Else
'            retVal = GetAllUsersAppDataFolder & "\ActiveApp\"
'        End If
'    End If
'    ActiveAppFolder = retVal
'End Function
'
'Public Function FavoritesFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        If PathExists(AppPath & "Favorites", False) Then
'            retVal = AppPath & "Favorites\"
'        Else
'            retVal = GetAllUsersAppDataFolder & "\Favorites\"
'        End If
'    End If
'    FavoritesFolder = retVal
'End Function
'Public Function GraphicsFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        If PathExists(AppPath & "Graphics", False) Then
'            retVal = AppPath & "Graphics\"
'        Else
'            retVal = GetAllUsersAppDataFolder & "\Graphics\"
'        End If
'    End If
'    GraphicsFolder = retVal
'End Function
'Public Function ProjectFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        If PathExists(AppPath & "Projects", False) Then
'            retVal = AppPath & "Projects\"
'        Else
'            retVal = GetAllUsersAppDataFolder & "\Projects\"
'        End If
'    End If
'    ProjectFolder = retVal
'End Function
'
'Public Function TemplatesFolder() As String
'    Static retVal As String
'    If retVal = "" Then
'        If PathExists(AppPath & "Projects\Templates", False) Then
'            retVal = AppPath & "Projects\Templates\"
'        Else
'            retVal = GetAllUsersAppDataFolder & "\Projects\Templates\"
'        End If
'    End If
'    TemplatesFolder = retVal
'End Function

'Private Type BROWSEINFO
'    hOwner As Long
'    pidlRoot As Long
'    pszDisplayName As String
'    lpszTitle As String
'    ulFlags As Long
'    lpfn As Long
'    lParam As Long
'    iImage As Long
'End Type
'
'Private Const BIF_RETURNONLYFSDIRS = &H1
'Private Const BIF_DONTGOBELOWDOMAIN = &H2
'Private Const BIF_STATUSTEXT = &H4
'Private Const BIF_RETURNFSANCESTORS = &H8
'Private Const BIF_BROWSEFORCOMPUTER = &H1000
'Private Const BIF_BROWSEFORPRINTER = &H2000
'
'Private Const SHGFI_DISPLAYNAME = &H200
'
'Private Const SHGFI_EXETYPE = &H2000
'Private Const SHGFI_SYSICONINDEX = &H4000
'Private Const SHGFI_LARGEICON = &H0
'Private Const SHGFI_SMALLICON = &H1
'Private Const SHGFI_SHELLICONSIZE = &H4
'Private Const SHGFI_TYPENAME = &H400
'
'Private Const ILD_TRANSPARENT = &H1
'
'Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
'
'Private Type SHFILEINFO
'   hIcon As Long
'   iIcon As Long
'   dwAttributes As Long
'   szDisplayName As String * MAX_PATH
'   szTypeName As String * 80
'End Type

'Public Function GetDownloadsFolder() As String
'    Dim BI As BROWSEINFO
'    Dim nFolder As Long
'    Dim IDL As ITEMIDLIST
'    Dim sPath As String
'    With BI
'        nFolder = CSIDL_PERSONAL
'
'        If SHGetSpecialFolderLocation(ByVal 0&, ByVal nFolder, IDL) = NOERROR Then
'            .pidlRoot = IDL.mkid.cb
'        End If
'
'        .pszDisplayName = String$(MAX_PATH, 0)
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
'    GetDownloadsFolder = sPath & "\Downloads"
'End Function

Public Sub FindMicrosoftNetworkNameByMachine(ByVal computername As String, ByRef domainworkgroup As String)
   
    Dim lngresult As Long
    Dim lngenumhwnd As Long
    Dim lngentries As Long
    Dim i As Integer
    Dim strremotename As String
    Dim intseppos As Integer
    Dim netdata(511) As NETRESOURCE
    lngentries = -1

    lngresult = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, ByVal 0, lngenumhwnd)
    If lngresult = 0 And lngenumhwnd <> 0 Then
        DoEvents
        lngresult = WNetEnumResource(lngenumhwnd, lngentries, netdata(0), CLng(Len(netdata(0))) * 512)
        If lngresult = 0 Then
            DoEvents
            For i = 0 To lngentries - 1
                strremotename = ltos(netdata(i).lpRemoteName)
                intseppos = InStrRev(strremotename, "\")
                strremotename = IIf(intseppos > 0, Right(strremotename, Len(strremotename) - intseppos), strremotename)
                If netdata(i).dwUsage And RESOURCEUSAGE_CONTAINER And strremotename = "Microsoft Windows Network" Then
                    SearchNetworkChildren netdata(i), strremotename, computername, domainworkgroup
                    If domainworkgroup <> "" Then Exit For
                End If
            Next i
        ElseIf lngresult = ERROR_NO_MORE_ITEMS Then
        End If
    End If
    lngresult = WNetCloseEnum(lngenumhwnd)

End Sub

Private Sub SearchNetworkChildren(netdata_parent As NETRESOURCE, strremotename_parent As String, ByVal computername As String, ByRef domainworkgroup As String)
    Dim lngDefault As Long

    Dim lngresult As Long
    Dim lngenumhwnd As Long
    Dim lngentries As Long
    Dim i As Integer
    Dim strremotename As String
    Dim intseppos As Integer
    Dim netdata(511) As NETRESOURCE
    lngentries = -1

    lngresult = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, netdata_parent, lngenumhwnd)
    If lngresult = 0 And lngenumhwnd <> 0 Then
        DoEvents
        lngresult = WNetEnumResource(lngenumhwnd, lngentries, netdata(0), CLng(Len(netdata(0))) * 512)
        If lngresult = 0 Then
            DoEvents
            For i = 0 To lngentries - 1
                strremotename = ltos(netdata(i).lpRemoteName)
                intseppos = InStrRev(strremotename, "\")
                strremotename = IIf(intseppos > 0, Right(strremotename, Len(strremotename) - intseppos), strremotename)

                If strremotename = computername Then
                    domainworkgroup = strremotename_parent
                    Exit For
                End If
                If netdata(i).dwUsage And RESOURCEUSAGE_CONTAINER Then
                    SearchNetworkChildren netdata(i), strremotename, computername, domainworkgroup
                    If domainworkgroup <> "" Then Exit For
                    If strremotename_parent = "Microsoft Windows Network" Then lngDefault = lngDefault + 1
                End If
            Next i
            If domainworkgroup = "" And strremotename_parent = "Microsoft Windows Network" Then
            
                Dim objNameSpace As Object
                Dim objDomain As Object
                Set objNameSpace = GetObject("WinNT:")
                
                For Each objDomain In objNameSpace
                    DoEvents
                    domainworkgroup = objDomain.name
                    Exit For
                Next
                
                Set objDomain = Nothing
                Set objNameSpace = Nothing
            
                If domainworkgroup = "" And lngDefault = 1 Then
                    domainworkgroup = strremotename
                End If
                
            End If
            
        End If
    End If
    lngresult = WNetCloseEnum(lngenumhwnd)

End Sub

Public Function GetDomainByMachine(ByVal Machine As String) As String

    Static CacheName As String
    If CacheName = "" Then
       
        FindMicrosoftNetworkNameByMachine Machine, CacheName
        If CacheName = "" Then

            Static showmsgbox As Boolean
            If Not showmsgbox Then
                If Not LCase(App.EXEName) = "maxservice" Then
                    MsgBox "WARNING!!!!  Unable to resolve the domain or workgroup, it is recomended" & vbCrLf & _
                            "that you disable the Network roaming allowance of users data, the domain" & vbCrLf & _
                            "must be aquired to properly share encrypt/decrypt data of Max-FTP users." & vbCrLf _
                            , vbCritical + vbOKOnly, "Warning"
                End If
                showmsgbox = True
            End If

        End If
    End If
    GetDomainByMachine = CacheName

End Function

Private Function ltos(lngh As Long) As String

    Dim strl As String
    strl = Space(lstrlen(lngh))
    lstrcpy strl, lngh
    ltos = strl

End Function

Public Function IsSchedulerInstalled() As Boolean
    IsSchedulerInstalled = PathExists(AppPath & ServiceFileName)
End Function
Public Function IsDocumentationInstalled() As Boolean
    IsDocumentationInstalled = PathExists(AppPath & "Help\index.htm")
End Function
Public Function IsScriptIDEInstalled() As Boolean
    IsScriptIDEInstalled = PathExists(AppPath & MaxIDEFileName)
End Function

Public Function IsAppMaxService() As Boolean
    IsAppMaxService = (Trim(Replace(LCase(App.EXEName), ".exe", "")) = Trim(Replace(LCase(ServiceFileName), ".exe", "")))
End Function

Public Function IsAppMaxFTP() As Boolean
    IsAppMaxFTP = (Trim(Replace(LCase(App.EXEName), ".exe", "")) = Trim(Replace(LCase(MaxFileName), ".exe", "")))
End Function

Public Function IsWindowOpen(ByVal WinCaption As String) As Boolean
    Dim isOpen As Boolean
    isOpen = Not (FindWindow(vbNullString, WinCaption & Chr(0)) = 0)
    If Not isOpen Then isOpen = Not (FindWindow("ThunderFormDC" & Chr(0), WinCaption & Chr(0)) = 0)
    If Not isOpen Then isOpen = Not (FindWindow("ThunderRT6FormDC" & Chr(0), WinCaption & Chr(0)) = 0)
    IsWindowOpen = isOpen
End Function

Public Function IsActiveAppFolder(ByVal Folder As String) As Boolean
    IsActiveAppFolder = (Left(Trim(LCase(Folder)), Len(GetTemporaryFolder & "\" & ActiveAppFolder) - 1) = GetTemporaryFolder & "\" & ActiveAppFolder)
End Function
Public Function MapFolderVariables(ByVal FolderString As String, Optional ByVal FileName As String = "")
    Dim ReturnFolder As String
    ReturnFolder = FolderString
    
    If InStr(LCase(ReturnFolder), "%appfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%appfolder%", Left(AppPath, Len(AppPath) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%favoritesfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%favoritesfolder%", Left(GetWinFavoritesDir, Len(GetWinFavoritesDir) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%wintempfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%wintempfolder%", Left(GetWinTempDir, Len(GetWinTempDir) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%systemfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%systemfolder%", Left(GetWinSysDir, Len(GetWinSysDir) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%systemroot%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%systemroot%", Left(GetWinDir, Len(GetWinDir) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%windowsfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%windowsfolder%", Left(GetWinDir, Len(GetWinDir) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%activeappfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%activeappfolder%", Left(ActiveAppFolder, Len(ActiveAppFolder) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%graphicsfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%graphicsfolder%", Left(GraphicsFolder, Len(GraphicsFolder) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%tempfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%tempfolder%", GetTemporaryFolder, , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%templatesfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%templatesfolder%", Left(TemplatesFolder, Len(TemplatesFolder) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%projectfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%projectfolder%", Left(ProjectFolder, Len(ProjectFolder) - 1), , vbTextCompare)
    End If
    If InStr(LCase(ReturnFolder), "%helpfolder%") > 0 Then
        ReturnFolder = Replace(LCase(ReturnFolder), "%helpfolder%", Left(HelpFolder, Len(HelpFolder) - 1), , vbTextCompare)
    End If

    If FileName <> "" Then
        If Left(FileName, 1) = "\" Then FileName = Mid(FileName, 1)
        If Right(ReturnFolder, 1) = "\" Then ReturnFolder = Left(ReturnFolder, Len(ReturnFolder) - 1)
        MapFolderVariables = ReturnFolder & "\" & FileName
    Else
        MapFolderVariables = ReturnFolder
    End If

End Function

Public Function GetMaxFavoritesDir(ByVal WinFavorites As Boolean) As String
    If Not WinFavorites Then
        GetMaxFavoritesDir = AppPath & FavoritesFolder
    Else
        GetMaxFavoritesDir = GetWinFavoritesDir
    End If
End Function

Public Function GetWinFavoritesDir()

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
    
    GetWinFavoritesDir = sPath
    
End Function

Public Function GetWinDir() As String
    Dim winDir As String
    Dim Ret As Long
    winDir = String(260, Chr(0))
    Ret = GetWindowsDirectory(winDir, 260)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    GetWinDir = winDir
End Function

Public Function GetWinSysDir() As String
    GetWinSysDir = SysPath
End Function

Public Function GetWinTempDir(Optional ByVal UseWin As Boolean = False) As String

    Dim winDir As String
    Dim Ret As Long
    winDir = String(260, Chr(0))
    Ret = GetTempPath(260, winDir)
    If (Not ((Ret = 16) And UseWin)) And (Not ((Ret = 34) And Not UseWin)) Then
        If PathExists(GetWinDir() + "TEMP") Then
            winDir = GetWinDir() + "TEMP\"
        Else
            On Error Resume Next
            MkDir GetWinDir() + "TEMP"
            If Err.number <> 0 Then Err.Clear
            On Error GoTo 0
            
            If PathExists(GetWinDir() + "TEMP") Then
                winDir = GetWinDir() + "TEMP\"
            Else
                winDir = ""
            End If
        End If
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWinTempDir = winDir

End Function

Public Function rsEnd(ByRef rs As ADODB.Recordset) As Boolean
    rsEnd = (rs.EOF Or rs.BOF)
End Function

Public Sub rsClose(ByRef rs As ADODB.Recordset, Optional ByVal SetNothing As Boolean = True)
    If Not rs.State = 0 Then rs.Close
    If SetNothing Then Set rs = Nothing
End Sub





