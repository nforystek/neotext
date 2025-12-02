Attribute VB_Name = "modFileSys"












#Const modFileSys = -1
Option Explicit
Option Compare Binary
Option Base 1
Option Private Module
' OpenFile() Flags
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1
Public Const OF_READWRITE = &H2
Public Const OF_SHARE_COMPAT = &H0
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_NONE = &H40
Public Const OF_PARSE = &H100
Public Const OF_DELETE = &H200
Public Const OF_VERIFY = &H400
Public Const OF_CANCEL = &H800
Public Const OF_CREATE = &H1000
Public Const OF_PROMPT = &H2000
Public Const OF_EXIST = &H4000
Public Const OF_REOPEN = &H8000

Public Const OFS_MAXPATHNAME = 128

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

' OpenFile() Structure
Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type



Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

' GetTempFileName() Flags
'
Public Const TF_FORCEDRIVE = &H80

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Public Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Public Declare Function LockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Public Const LOCKFILE_FAIL_IMMEDIATELY = &H1
Public Const LOCKFILE_EXCLUSIVE_LOCK = &H2

Public Declare Function UnlockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function hGetFileType Lib "kernel32" Alias "GetFileType" (ByVal hFile As Long) As Long
Public Declare Function hGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Public Declare Function hWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function hReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function hGetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DataModes_ListingToString = 0
Public Const DataModes_ListingToFile = 1
Public Const DataModes_GetRemoteFile = 2
Public Const DataModes_PutRemoteFile = 3
Public Const DataModes_LocalToLocal = 4
Public Const DataModes_ServerToServer = 5

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Private Property Get GetObject(ByRef lptr As Long) As Object
'    Dim NewObj As Object
'    RtlMoveMemory NewObj, lptr, 4&
'    Set GetObject = NewObj
'    RtlMoveMemory NewObj, 0&, 4&
'End Property

Public Function WholeNumb(num As Double) As Double
    Dim str As String
    str = CStr(num)
    If InStr(num, ".") > 0 Then
        str = Left(str, InStr(str, ".") - 1)
    End If
    WholeNumb = CDbl(str)
End Function
Public Function ZeroNumb(num) As Double
    Dim str As String
    str = CStr(num)
    If InStr(num, ".") > 0 Then
        str = Mid(str, InStr(str, "."))
    Else
        str = "0"
    End If
    ZeroNumb = CDbl(str)
End Function

Public Function ReturnDiv(Dbl1st, Dbl2nd) As Double
    ReturnDiv = CDbl(IIf(Dbl2nd > Dbl1st, 0, CDbl(IIf(Dbl1st < Dbl2nd, 0, WholeNumb(Dbl1st / Dbl2nd)))))
End Function
Public Function ReturnMod(Dbl1st As Double, Dbl2nd As Double) As Double
    ReturnMod = CDbl(IIf(Dbl2nd > Dbl1st, Dbl1st, CDbl(IIf(Dbl1st < Dbl2nd, Dbl1st, CDbl(IIf(WholeNumb(Dbl1st / Dbl2nd) = (Dbl1st / Dbl2nd), 0, Dbl1st - (WholeNumb(Dbl1st / Dbl2nd) * Dbl2nd)))))))
End Function

Public Function MapFileName(ByVal FullName As String) As String
    If InStr(FullName, "/") = 0 And InStr(FullName, "\") = 0 Then
        MapFileName = FullName
    Else
        Dim dirChar As String
        Dim nURL As New URL
        dirChar = nURL.GetDirChar(FullName)
        Set nURL = Nothing
        MapFileName = Mid(FullName, InStrRev(FullName, dirChar) + 1)
    End If
End Function
Public Function MapFolder(ByVal RootFolder As String, ByVal AddFolder As String, ByVal ftpType As URLTypes) As String
    Dim NewFolder As String
    Dim nURL As New URL
    
    If AddFolder = ".." Then
        NewFolder = nURL.GetParentFolder(RootFolder)
    Else
        Dim dirChar As String
        
        dirChar = nURL.GetDirChar(RootFolder)
        If dirChar = "/" Then
            AddFolder = Replace(AddFolder, "\", "/")
        Else
            AddFolder = Replace(AddFolder, "/", "\")
        End If
        
        If ftpType = URLTypes.ftp Then
        
            If Left(AddFolder, 1) = dirChar Then
                NewFolder = AddFolder
            ElseIf Right(RootFolder, 1) = dirChar And Left(AddFolder, 1) = dirChar Then
                NewFolder = RootFolder & Mid(AddFolder, 2)
            ElseIf Right(RootFolder, 1) = dirChar And Left(AddFolder, 1) <> dirChar Then
                NewFolder = RootFolder & AddFolder
            Else
                NewFolder = RootFolder & dirChar & AddFolder
            End If
        
        ElseIf ftpType = URLTypes.File Or ftpType = URLTypes.Remote Then
            
            If Right(RootFolder, 1) = dirChar And Left(AddFolder, 1) = dirChar Then
                NewFolder = RootFolder & Mid(AddFolder, 2)
            ElseIf Right(RootFolder, 1) = dirChar And Left(AddFolder, 1) <> dirChar Then
                NewFolder = RootFolder & AddFolder
            ElseIf Right(RootFolder, 1) <> dirChar And Left(AddFolder, 1) = dirChar Then
                NewFolder = RootFolder & AddFolder
            Else
                NewFolder = RootFolder & dirChar & AddFolder
            End If
        
        End If
        
    End If

    Set nURL = Nothing
    MapFolder = NewFolder
End Function

Public Function GetWinDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetWindowsDirectory(winDir, 255)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Trim(Dir(winDir, vbDirectory)) = "" Then winDir = App.Path
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











