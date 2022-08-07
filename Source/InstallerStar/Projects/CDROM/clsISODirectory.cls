VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISODirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" _
Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function FindNextFile Lib "kernel32" _
Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function GetFileAttributes Lib "kernel32" _
Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String _
) As Long

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
) As Long

Private Const MAX_PATH                  As Long = 260
Private Const MAXDWORD                  As Long = &HFFFF
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY  As Long = &H100

Private Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long
    FTCreationTime              As FILETIME
    FTLastAccessTime            As FILETIME
    FTLastWriteTime             As FILETIME
    nFileSizeHigh               As Long
    nFileSizeLow                As Long
    dwReserved0                 As Long
    dwReserved1                 As Long
    cFileName                   As String * MAX_PATH
    cAlternate                  As String * 14
End Type

Private clsParent               As clsISODirectory

Private clsSubDirs()            As clsISODirectory
Private lngSubDirCnt            As Long

Private strName                 As String
Private strDOSName              As String
Private strJolName              As String

Private clsFiles                As clsISOFiles

Private lngDirNum               As Long
Private lngDirOrderNum          As Long

Private blnNValid               As Boolean

Public Property Get NameValid( _
) As Boolean

    NameValid = blnNValid
End Property

Public Property Let NameValid( _
    ByVal bln As Boolean _
)

    blnNValid = bln
End Property

Public Property Get DirectoryOrderPos( _
) As Long

    DirectoryOrderPos = lngDirOrderNum
End Property

Public Property Let DirectoryOrderPos( _
    ByVal lngNewVal As Long _
)

    lngDirOrderNum = lngNewVal
End Property

Public Property Get DirectoryNumber( _
) As Long

    DirectoryNumber = lngDirNum
End Property

Public Property Let DirectoryNumber( _
    ByVal lngNewVal As Long _
)

    lngDirNum = lngNewVal
End Property

Public Function SubDirExists( _
    ByVal name As String _
) As Boolean

    Dim i   As Long

    For i = 0 To SubDirectoryCount - 1
        If StrComp(SubDirectory(i).name, name, vbTextCompare) = 0 Then
            SubDirExists = True
            Exit Function
        End If
    Next
End Function

Public Property Get FullPath( _
) As String

    Dim i           As Long
    Dim clsISODir   As clsISODirectory
    Dim strPath     As String
    Dim strParts()  As String

    Set clsISODir = Me

    Do While Not clsISODir Is Nothing
        strPath = strPath & "\" & clsISODir.name
        Set clsISODir = clsISODir.Parent
    Loop

    strParts = Split(strPath, "\")

    For i = 0 To UBound(strParts)
        If i > UBound(strParts) / 2 Then Exit For

        strPath = strParts(UBound(strParts) - i)
        strParts(UBound(strParts) - i) = strParts(i)
        strParts(i) = strPath
    Next

    FullPath = Join(strParts, "\")
End Property

Public Sub AddLocalDirectory( _
    ByVal strPath As String, _
    filter() As String, _
    Optional ByVal recursive As Boolean = True _
)

    FindFilesAPI strPath, filter, recursive, Me
End Sub

Public Function AddSubDirectoryByRef( _
    directory As clsISODirectory _
) As clsISODirectory

    ReDim Preserve clsSubDirs(lngSubDirCnt) As clsISODirectory

    Set clsSubDirs(lngSubDirCnt) = directory
    Set clsSubDirs(lngSubDirCnt).Parent = Me

    Set AddSubDirectoryByRef = clsSubDirs(lngSubDirCnt)

    lngSubDirCnt = lngSubDirCnt + 1
End Function

Public Function AddSubDirectory( _
    ByVal name As String _
) As clsISODirectory

    Dim i   As Long

    name = Trim$(name)

    For i = 0 To SubDirectoryCount - 1
        If StrComp(name, SubDirectory(i).name, vbTextCompare) = 0 Then
            Set AddSubDirectory = SubDirectory(i)
            Exit Function
        End If
    Next

    ReDim Preserve clsSubDirs(lngSubDirCnt) As clsISODirectory

    Set clsSubDirs(lngSubDirCnt) = New clsISODirectory
    clsSubDirs(lngSubDirCnt).name = name

    Set clsSubDirs(lngSubDirCnt).Parent = Me

    Set AddSubDirectory = clsSubDirs(lngSubDirCnt)

    lngSubDirCnt = lngSubDirCnt + 1
End Function

Public Sub RemoveSubDirectory( _
    ByVal index As Long, _
    Optional ByVal KeepSubDirs As Boolean _
)

    Dim i   As Long

    If Not KeepSubDirs Then
        For i = 0 To clsSubDirs(index).SubDirectoryCount - 1
            clsSubDirs(index).RemoveSubDirectory i
        Next
    End If

    For i = index + 1 To lngSubDirCnt - 1
        Set clsSubDirs(i - 1) = clsSubDirs(i)
    Next

    lngSubDirCnt = lngSubDirCnt - 1
    If lngSubDirCnt < 0 Then lngSubDirCnt = 0
    ReDim Preserve clsSubDirs(lngSubDirCnt) As clsISODirectory
End Sub

Public Sub Clear()
    Erase clsSubDirs
    lngSubDirCnt = 0
    clsFiles.Clear
End Sub

Public Property Get JolietName( _
) As String

    JolietName = strJolName
End Property

Public Property Let JolietName( _
    ByVal strVal As String _
)

    strJolName = strVal
End Property

Public Property Get DOSName( _
) As String

    DOSName = strDOSName
End Property

Public Property Let DOSName( _
    ByVal strVal As String _
)

    strDOSName = strVal
End Property

Public Property Get name() As String
    name = strName
End Property

Public Property Let name( _
    ByVal strNewName As String _
)

    If strNewName = "" Then Exit Property

    strName = strNewName
End Property

Public Property Get Parent( _
) As clsISODirectory

    Set Parent = clsParent
End Property

Public Property Set Parent( _
    clsNewParent As clsISODirectory _
)

    Set clsParent = clsNewParent
End Property

Public Property Get SubDirectoryCount( _
) As Long

    SubDirectoryCount = lngSubDirCnt
End Property

Public Property Get SubDirectory( _
    ByVal index As Long _
) As clsISODirectory

    Set SubDirectory = clsSubDirs(index)
End Property

Public Property Get Files( _
) As clsISOFiles

    Set Files = clsFiles
End Property

Private Sub Class_Initialize()
    Set clsFiles = New clsISOFiles
End Sub

Private Sub FindFilesAPI( _
    ByVal path As String, _
    filter() As String, _
    ByVal recursive As Boolean, _
    clsDir As clsISODirectory _
)

    Dim hSearch     As Long
    Dim udtFindData As WIN32_FIND_DATA
    Dim lngRet      As Long
    Dim i           As Long
    Dim clsSubDir   As clsISODirectory

    If Not Right$(path, 1) = "\" Then path = path & "\"

    hSearch = FindFirstFile(path & "*.*", udtFindData)
    If hSearch = INVALID_HANDLE Then Exit Sub

    If Not (Left$(udtFindData.cFileName, 1) = "." Or Left$(udtFindData.cFileName, 2) = "..") Then
        For i = LBound(filter) To UBound(filter)
            If StripNulls(udtFindData.cFileName) Like filter(i) Then
                If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If recursive Then
                        Set clsSubDir = clsDir.AddSubDirectory(Trim$(StripNulls(udtFindData.cFileName)))
                        FindFilesAPI path & clsSubDir.name, filter, recursive, clsSubDir
                    End If
                Else
                    clsDir.Files.Add path & Trim$(StripNulls(udtFindData.cFileName))
                End If
                Exit For
            End If
        Next
    End If

    lngRet = 1

    Do
        lngRet = FindNextFile(hSearch, udtFindData)
        If lngRet = 0 Then Exit Do

        If Not (Left$(udtFindData.cFileName, 1) = "." Or Left$(udtFindData.cFileName, 2) = "..") Then
            For i = LBound(filter) To UBound(filter)
                If StripNulls(udtFindData.cFileName) Like filter(i) Then
                    If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        If recursive Then
                            Set clsSubDir = clsDir.AddSubDirectory(Trim$(StripNulls(udtFindData.cFileName)))
                            FindFilesAPI path & clsSubDir.name, filter, recursive, clsSubDir
                        End If
                    Else
                        clsDir.Files.Add path & Trim$(StripNulls(udtFindData.cFileName))
                    End If
                    Exit For
                End If
            Next
        End If
    Loop

    FindClose hSearch
End Sub

Private Function StripNulls( _
    OriginalStr As String _
) As String

    If InStr(OriginalStr, Chr(0)) > 0 Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = OriginalStr
End Function

Private Sub Class_Terminate()
    Set clsFiles = Nothing
    Set clsParent = Nothing
    Erase clsSubDirs
End Sub
