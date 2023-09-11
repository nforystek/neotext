Attribute VB_Name = "modDeclare1"
Option Explicit
' ------------------------------------------------------------------------
'
'    WIN32API.TXT -- Win32 API Declarations for Visual Basic
'
'              Copyright (C) 1994-98 Microsoft Corporation
'
'  This file is required for the Visual Basic 6.0 version of the APILoader.
'  Older versions of this file will not work correctly with the version
'  6.0 APILoader.  This file is backwards compatible with previous releases
'  of the APILoader with the exception that Constants are no longer declared
'  as Global or Public in this file.
'
'  This file contains only the Const, Type,
'  and Public Declare statements for  Win32 APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Public Declare Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Public Declare Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Public Declare Function InterlockedExchange Lib "kernel32" (Target As Long, ByVal Value As Long) As Long
' Loader Routines

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function SetProcessShutdownParameters Lib "kernel32" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetProcessShutdownParameters Lib "kernel32" (lpdwLevel As Long, lpdwFlags As Long) As Long
Public Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Public Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub DebugBreak Lib "kernel32" ()
Public Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Public Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function PulseEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Public Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Public Declare Function LockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function UnlockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Sub SetLastErrorEx Lib "user32" (ByVal dwErrCode As Long, ByVal dwType As Long)
Public Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Public Declare Sub SetDebugErrorLevel Lib "user32" (ByVal dwLevel As Long)
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Public Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Public Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As Long
Public Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Public Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function InitAtomTable Lib "kernel32" (ByVal nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
' User Profile Routines
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'       GetPrivateProfileString, and WritePrivateProfileString can be either
'       a string or NULL.  This is why the argument is defined as "As Any".
'          For example, to pass a string specify   ByVal "wallpaper"
'          To pass NULL specify                    ByVal 0&
'       You can also pass NULL for the lpString argument for WriteProfileString
'       and WritePrivateProfileString

Public Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
Public Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Public Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function GetNamedPipeHandleState Lib "kernel32" Alias "GetNamedPipeHandleStateA" (ByVal hNamedPipe As Long, lpState As Long, lpCurInstances As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long, ByVal lpUserName As String, ByVal nMaxUserNameSize As Long) As Long
Public Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Public Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Public Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Public Declare Sub SetFileApisToOEM Lib "kernel32" ()
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function ClearEventLog Lib "advapi32.dll" Alias "ClearEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Public Declare Function BackupEventLog Lib "advapi32.dll" Alias "BackupEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Public Declare Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Public Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Public Declare Function GetNumberOfEventLogRecords Lib "advapi32.dll" (ByVal hEventLog As Long, NumberOfRecords As Long) As Long
Public Declare Function GetOldestEventLogRecord Lib "advapi32.dll" (ByVal hEventLog As Long, OldestRecord As Long) As Long
Public Declare Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Public Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Public Declare Function OpenBackupEventLog Lib "advapi32.dll" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As Long
Public Declare Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
Public Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long
Public Declare Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As Long, ImpersonationLevel As Integer, DuplicateTokenHandle As Long) As Long
Public Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal handle As Long, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Public Declare Function ImpersonateNamedPipeClient Lib "advapi32.dll" (ByVal hNamedPipe As Long) As Long
Public Declare Function ImpersonateSelf Lib "advapi32.dll" (ImpersonationLevel As Integer) As Long
Public Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Public Declare Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Public Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Public Declare Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function AdjustTokenGroups Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
Public Declare Function PrivilegeCheck Lib "advapi32.dll" (ByVal ClientToken As Long, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As Long) As Long
Public Declare Function AccessCheckAndAuditAlarm Lib "advapi32.dll" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As Long, ByVal pfGenerateOnClose As Long) As Long
Public Declare Function ObjectOpenAuditAlarm Lib "kernel32" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As Long) As Long
Public Declare Function ObjectPrivilegeAuditAlarm Lib "advapi32.dll" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As Long, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Public Declare Function ObjectCloseAuditAlarm Lib "advapi32.dll" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
Public Declare Function PrivilegedServiceAuditAlarm Lib "advapi32.dll" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal ServiceName As String, ByVal ClientToken As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Public Declare Function IsValidSid Lib "advapi32.dll" (pSid As Any) As Long
Public Declare Function EqualSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Public Declare Function EqualPrefixSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Public Declare Function GetSidLengthRequired Lib "advapi32.dll" (ByVal nSubAuthorityCount As Byte) As Long
Public Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Public Declare Sub FreeSid Lib "advapi32.dll" (pSid As Any)
Public Declare Function InitializeSid Lib "advapi32.dll" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
Public Declare Function GetSidIdentifierAuthority Lib "advapi32.dll" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
Public Declare Function GetSidSubAuthority Lib "advapi32.dll" (pSid As Any, ByVal nSubAuthority As Long) As Long
Public Declare Function GetSidSubAuthorityCount Lib "advapi32.dll" (pSid As Any) As Byte
Public Declare Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
Public Declare Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
Public Declare Function AreAllAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Public Declare Function AreAnyAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Public Declare Sub MapGenericMask Lib "advapi32.dll" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
Public Declare Function IsValidAcl Lib "advapi32.dll" (pAcl As ACL) As Long
Public Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Public Declare Function GetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Public Declare Function SetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Public Declare Function AddAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
Public Declare Function DeleteAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
Public Declare Function GetAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
Public Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Public Declare Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Public Declare Function AddAuditAccessAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
Public Declare Function FindFirstFreeAce Lib "advapi32.dll" (pAcl As ACL, pAce As Long) As Long
Public Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Public Declare Function IsValidSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function GetSecurityDescriptorLength Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
Public Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
Public Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
Public Declare Function SetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
Public Declare Function GetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As Long, pSacl As ACL, ByVal lpbSaclDefaulted As Long) As Long
Public Declare Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
Public Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
Public Declare Function SetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
Public Declare Function GetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As Long) As Long
Public Declare Function CreatePrivateObjectSecurity Lib "advapi32.dll" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As Long, GenericMapping As GENERIC_MAPPING) As Long
Public Declare Function SetPrivateObjectSecurity Lib "advapi32.dll" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As Long) As Long
Public Declare Function GetPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
Public Declare Function DestroyPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function MakeSelfRelativeSD Lib "advapi32.dll" (pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwBufferLength As Long) As Long
Public Declare Function MakeAbsoluteSD Lib "advapi32.dll" (pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwAbsoluteSecurityDescriptorSize As Long, pDacl As ACL, lpdwDaclSize As Long, pSacl As ACL, lpdwSaclSize As Long, pOwner As Any, lpdwOwnerSize As Long, pPrimaryGroup As Any, lpdwPrimaryGroupSize As Long) As Long
Public Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Public Declare Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal handle As Long, ByVal SecurityInformation As Long, SecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Public Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Public Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Public Declare Function VirtualLock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Public Declare Function VirtualUnlock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Public Declare Function MapViewOfFileEx Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long, lpBaseAddress As Any) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)
Public Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function UnhandledExceptionFilter Lib "kernel32" (ExceptionInfo As EXCEPTION_POINTERS) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
' COMM declarations

Public Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Public Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Public Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Public Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function TransmitCommChar Lib "kernel32" (ByVal nCid As Long, ByVal cChar As Byte) As Long
Public Declare Function SetCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Public Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Public Declare Function ClearCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Public Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Public Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Public Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Public Declare Function GetCommMask Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long) As Long
Public Declare Function GetCommProperties Lib "kernel32" (ByVal hFile As Long, lpCommProp As COMMPROP) As Long
Public Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Long
Public Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function SetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
Public Declare Function GetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
Public Declare Function PrepareTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
Public Declare Function EraseTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
Public Declare Function CreateTapePartition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
Public Declare Function WriteTapemark Lib "kernel32" (ByVal hDevice As Long, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long
Public Declare Function GetTapeStatus Lib "kernel32" (ByVal hDevice As Long) As Long
Public Declare Function GetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long
Public Declare Function SetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpTapeInformation As Any) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
Public Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SystemTime)
Public Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function SetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
' Routines to convert back and forth
' between system time and file time

Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Public Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Public Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
Public Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Public Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Public Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Public Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Public Declare Function TransactNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Public Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Public Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Public Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Public Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Public Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Public Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
Public Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Public Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
Public Declare Function TlsAlloc Lib "kernel32" () As Long
Public Declare Function TlsGetValue Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Public Declare Function TlsSetValue Lib "kernel32" (ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
Public Declare Function TlsFree Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
Public Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As Long) As Long
Public Declare Function BackupWrite Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Long) As Long
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
Public Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As Long) As Long
Public Declare Function IsBadHugeReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadHugeWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, Sid As Any, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Public Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Public Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
Public Declare Function LookupPrivilegeDisplayName Lib "advapi32.dll" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
Public Declare Function AllocateLocallyUniqueId Lib "advapi32.dll" (Luid As LARGE_INTEGER) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' Performance counter API's

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
' Code Page Dependent APIs

Public Declare Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
Public Declare Function GetACP Lib "kernel32" () As Long
Public Declare Function GetOEMCP Lib "kernel32" () As Long
Public Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Public Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
' Locale Dependent APIs

Public Declare Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Public Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SystemTime, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SystemTime, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Public Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
