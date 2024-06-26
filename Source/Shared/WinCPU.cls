VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinCPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'TOP DOWN
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const HKEY_DYN_DATA = &H80000006
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Private HKey As Long, dwDataSize As Long, dwCpuUsage As Byte, dwType As Long
Public Sub Initialize()
    'start the counter by reading the value of the 'StartStat' key
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StartStat", 0, KEY_ALL_ACCESS, HKey) <> ERROR_SUCCESS Then
        Debug.Print "Error while initializing counter"
        Exit Sub
    End If
    dwDataSize = 4 'Length of Long
    RegQueryValueEx HKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    RegCloseKey HKey
    'get current counter's value
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_READ, HKey) <> ERROR_SUCCESS Then
        Debug.Print "Error while opening counter key"
        Exit Sub
    End If
End Sub
Public Function Query() As Long
    dwDataSize = 4 'size of Long
    'Query the counter
    RegQueryValueEx HKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    Query = CLng(dwCpuUsage)
End Function
Public Sub Terminate()
    RegCloseKey HKey
    'stopping the counter
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StopStat", 0, KEY_ALL_ACCESS, HKey) <> ERROR_SUCCESS Then
        Debug.Print "Error while stopping counter"
        Exit Sub
    End If
    dwDataSize = 4 'length of Long
    RegQueryValueEx HKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    RegCloseKey HKey
End Sub

