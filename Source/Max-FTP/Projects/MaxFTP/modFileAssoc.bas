#Const [True] = -1
#Const [False] = 0



Attribute VB_Name = "modFileAssoc"
#Const modFileAssoc = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Enum BrowseActions
    bws_File = 0
    bws_Desktop = 1
    
    bws_ProgramFolder = 2
    bws_Network = 14
    
    bws_ControlPanel = 3
    bws_Printers = 4
    
    bws_MyDocuments = 5
    bws_Favorites = 6
    
    bws_RecyclingBin = 10
    bws_MyComputer = 13
End Enum

Private Const MAX_PATH = 260

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80

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

Private Declare Function ImageList_Draw Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

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
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long

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

Public Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private ShInfo As SHFILEINFO

Private Function GetFolderValue(tAction As BrowseActions) As Long
    
    Dim wIdx As Integer
    wIdx = CInt(tAction)
    
    If wIdx < 2 Then
      GetFolderValue = 0
    
    ElseIf wIdx < 12 Then
      GetFolderValue = wIdx
    
    Else
      GetFolderValue = wIdx + 4
    End If

End Function

Private Function GetReturnType() As Long
  Dim dwRtn As Long
  dwRtn = dwRtn Or BIF_RETURNONLYFSDIRS

  GetReturnType = dwRtn
End Function

Public Function BrowseAction(ByVal pBrowseAction As BrowseActions, ByVal ownerHWnd As Long) As String

    Dim sPath As String
    
    If pBrowseAction = bws_File Then
    
    Else
    
        Dim BI As BROWSEINFO
        Dim nFolder As Long
        Dim IDL As ITEMIDLIST
        Dim pIdl As Long
        Dim SHFI As SHFILEINFO
        With BI
            .hOwner = ownerHWnd
        
            nFolder = GetFolderValue(pBrowseAction)
        
            If SHGetSpecialFolderLocation(ByVal ownerHWnd, ByVal nFolder, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
                End If
        
            .pszDisplayName = String$(MAX_PATH, 0)
        
            .lpszTitle = "Browse for folder"
        
            .ulFlags = GetReturnType()
        
        End With
     
        pIdl = SHBrowseForFolder(BI)
      
        If pIdl = 0 Then
            sPath = ""
        Else
            
            sPath = String$(MAX_PATH, 0)
            SHGetPathFromIDList ByVal pIdl, ByVal sPath
    
            sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    
        End If
        
    End If
    
    BrowseAction = sPath
        
End Function
Public Function OpenAssociatedFile(ByVal FileName As String, ByVal Silent As Boolean) As Boolean
    On Error GoTo catch
    
    Dim retVal As Long
    
    If IsFileExecutable(FileName) Then
        retVal = RunProcess(FileName, "", vbNormalFocus, False)
    Else
        
        Dim FileExt As String
        FileExt = GetFileExt(FileName)
        

        Dim dbFileAssoc As New clsFileAssoc
        If dbFileAssoc.GetWindowsApp(FileExt) Then
            retVal = RunFile(FileName)
            If retVal <= 32 Then retVal = 0
        Else
            Dim AppExec As String
            AppExec = MapFolderVariables(Trim(dbFileAssoc.GetApplicationExe(FileExt)))
            If AppExec = "" Or AppExec = "(windows default)" Or AppExec = "(none)" Then
                retVal = RunFile(FileName)
                If retVal <= 32 Then retVal = 0
            Else
                If InStr(FileName, " ") > 0 Then FileName = Chr(34) & FileName & Chr(34)
                retVal = RunProcess(AppExec, FileName, vbNormalFocus, False)
            End If
        End If
        Set dbFileAssoc = Nothing

    End If
    If retVal = 0 And Not Silent Then
        MsgBox "Unable to open file " & FileName & " or it's associated application.", vbInformation, AppName
    End If
    OpenAssociatedFile = (retVal = 0)
    Exit Function
catch:
    MsgBox "Unable to open file " & FileName & " or it's associated application.", vbInformation, AppName
    Err.Clear
End Function
Public Function LoadAssociation(ByRef tPic As Variant) As String
    Dim testInt As String
    Dim lItem As Object
    For Each lItem In frmMain.imgFiles.ListImages
        If lItem.Key = tPic.Tag Then
            testInt = lItem.Key
            Exit For
        End If
    Next
    If testInt = "" Then Set lItem = Nothing
    If Not lItem Is Nothing Then
        frmMain.imgFiles.ListImages(tPic.Tag).Tag = frmMain.imgFiles.ListImages(tPic.Tag).Tag + 1
        LoadAssociation = testInt
    Else
        Err.Clear
        Dim ImgX As ListImage
        Select Case LCase(TypeName(tPic))
            Case "picturebox"
                Set ImgX = frmMain.imgFiles.ListImages.Add(, tPic.Tag, tPic.Image)
            Case "iimage"
                Set ImgX = frmMain.imgFiles.ListImages.Add(, tPic.Tag, tPic.Picture)
        End Select
        ImgX.Tag = 1
        LoadAssociation = ImgX.Key
    End If
End Function

Public Function RemoveAssociation(ByVal mTag As String) As String

    If mTag > 0 Then
        frmMain.imgFiles.ListImages(mTag).Tag = frmMain.imgFiles.ListImages(mTag).Tag - 1
        If frmMain.imgFiles.ListImages(mTag).Tag = 0 Then
        
            frmMain.imgFiles.ListImages.Remove mTag
        
        End If
    End If

End Function

Private Function GetIcon(ByVal FileName As String) As Long
    Dim hSIcon As Long
    
    hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    If hSIcon = 0 Then
        hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES Or FILE_ATTRIBUTE_NORMAL)
    End If
    GetIcon = hSIcon
    
End Function

Public Sub GetAssociation(ByVal fType As String, ByVal fName As String, ByRef pic As Control)

    Dim hImgSmall As Long
    Dim fKey As String
    
    hImgSmall = 0
    
    If Trim(fType) <> "" Then
        If InStr(LCase(fType), ".exe") > 0 Or InStr(LCase(fType), ".ico") > 0 Then
            If PathExists(fName) Then
                hImgSmall = GetIcon(fName)
                fKey = fName
            Else
                hImgSmall = GetIcon(fType)
                fKey = fType
            End If
        Else
            hImgSmall = GetIcon(fType)
            fKey = fType
        End If
    End If
    
    If hImgSmall = 0 Then
        fName = GetWinSysDir & "shell32.dll"
        hImgSmall = GetIcon(fName)
        fKey = fName
    End If

    Set pic.Picture = LoadPicture("")
    pic.AutoRedraw = True
    ImageList_Draw hImgSmall, ShInfo.iIcon, pic.hDC, 0, 0, ILD_TRANSPARENT
    pic.Refresh

    fKey = Replace(Replace(Replace(Replace(fKey, ".", "d"), "\", ""), ":", ""), "/", "")
        
    pic.Tag = Trim(LCase(fKey))
End Sub
