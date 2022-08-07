#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modSystem"
#Const modSystem = -1
Option Explicit
'TOP DOWN

Option Private Module
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
    
Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCompName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetMachineName() As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = String(255, Chr(0))
    lSize = Len(sBuffer)
    Call GetCompName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left(sBuffer, lSize)
    End If
    GetMachineName = Replace(sBuffer, Chr(0), "")
End Function

Public Function GetUserLoginName() As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left$(sBuffer, lSize)
    End If
    GetUserLoginName = Replace(sBuffer, Chr(0), "")
End Function


Public Function IsHost64Bit() As Boolean
    Dim handle As Long
    Dim is64Bit As Boolean

    ' Assume initially that this is not a WOW64 process
    is64Bit = False

    ' Then try to prove that wrong by attempting to load the
    ' IsWow64Process function dynamically
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    ' The function exists, so call it
    If handle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If

    ' Return the value
    IsHost64Bit = is64Bit
End Function

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If
End Function


Attribute 
