VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolderBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

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

Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                              pIdl As ITEMIDLIST) As Long
Const NOERROR = 0
Const CSIDL_DESKTOP = &H0
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
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

Const BIF_RETURNONLYFSDIRS = &H1
Const BIF_DONTGOBELOWDOMAIN = &H2
Const BIF_STATUSTEXT = &H4
Const BIF_RETURNFSANCESTORS = &H8
Const BIF_BROWSEFORCOMPUTER = &H1000
Const BIF_BROWSEFORPRINTER = &H2000

Private MousePos As POINTAPI

Private Const OffsetX = 3
Private Const OffsetY = 3

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private BrowseFolder As String

Public Property Get Folder() As String
    Folder = BrowseFolder
End Property

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

Public Function Browse(ByVal pBrowseAction As BrowseActions, ByVal ownerHWnd As Long) As String
On Error Resume Next
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
    
Err.Clear
On Error GoTo 0
BrowseFolder = sPath
Browse = sPath
End Function

