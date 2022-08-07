VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'this module contains WIN APIs and public functions to read/write/query the registry both local and remote machines.
'each function has an optional parameter, ComputerName. If left out, the function will be applied to the local machine.
'If populated, the function will be applied to that computer name.

''Example for local machine:
'Dim Registry As clsRegistry
'Set Registry = New clsRegistry
'If Registry.SetValue(eHKEY_LOCAL_MACHINE, "Software\Testing", "TestValue", "TestSetting") Then
'    MsgBox "Value was set"
'Else
'    MsgBox "Error setting value"
'End If


''Example for remote machine:
'Dim Registry As clsRegistry
'Set Registry = New clsRegistry
'If Registry.SetValue(eHKEY_LOCAL_MACHINE, "Software\Testing", "TestValue", "TestSetting", "Athlon700") Then
'    MsgBox "Value was set on Athlon700"
'Else
'    MsgBox "Error setting value on Athlon700"
'End If

'note:
'in order to access a registry on a remote machine, the account on the client PC (yours) must have admin rights on the server PC (the target)

'written by zakk hoyt (zakkhoyt@yahoo.com)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const ERROR_SUCCESS = 0
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2&              ' Unicode nul terminated string containing environment variables
Private Const REG_MULTI_SZ = 7&               ' Multiple Unicode strings

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function ArryVarPtr Lib "MSVBVM60.DLL" (var() As Any) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
#If Not modRegistry Then

Public Enum HKEYs
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
#End If
Private Declare Function RegOpenKeyEx Lib "ADVAPI32.DLL" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
Private Declare Function RegConnectRegistry Lib "ADVAPI32.DLL" _
        Alias "RegConnectRegistryA" _
        (ByVal lpMachineName As String, _
        ByVal hKey As Long, _
        phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long) As Long

'APIs to get/set values in the registry
Private Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExAny Lib "ADVAPI32.DLL" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData() As Any, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "ADVAPI32.DLL" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As String, _
        lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "ADVAPI32.DLL" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "ADVAPI32.DLL" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "ADVAPI32.DLL" _
        Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long

Private Declare Function RegEnumKeyEx Lib "ADVAPI32.DLL" _
        Alias "RegEnumKeyExA" _
        (ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpName As String, lpcbName As Long, _
        ByVal lpReserved As Long, ByVal lpClass As String, _
        lpcbClass As Long, _
        lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "ADVAPI32.DLL" _
        Alias "RegEnumValueA" _
        (ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long
        
Private Declare Function RegCreateKeyEx Lib "ADVAPI32.DLL" _
        Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "ADVAPI32.DLL" _
        Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String) As Long
        
        
Private Const REG_FORCE_RESTORE As Long = 8&
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Private Const SE_AUDIT_NAME = "SeAuditPrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Private Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Private Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Private Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Private Const SE_DACL_PRESENT = &H4
Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Private Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Private Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Private Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Private Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Private Const SE_OWNER_DEFAULTED = &H1
Private Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Private Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Private Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Private Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Private Const SE_TCB_NAME = "SeTcbPrivilege"
Private Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"

Private Const REG_OPTION_VOLATILE = 1

Private Type LUID
  lowpart As Long
  highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type

Private Declare Function AdjustTokenPrivileges Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "ADVAPI32.DLL" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "ADVAPI32.DLL" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


Private Function EnablePrivilege(seName As String) As Boolean
  On Error Resume Next
  Dim p_lngRtn As Long
  Dim p_lngToken As Long
  Dim p_lngBufferLen As Long
  Dim p_typLUID As LUID
  Dim p_typTokenPriv As TOKEN_PRIVILEGES
  Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

  p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
  If p_lngRtn = 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  If Err.LastDllError <> 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
  If p_lngRtn = 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  p_typTokenPriv.PrivilegeCount = 1
  p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
  p_typTokenPriv.Privileges.pLuid = p_typLUID
  EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'public wrappers for registry APIs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function SetValue(ByVal PredefinedKey As Long, ByVal KeyName As String, ByVal ValueName As String, ByRef Value As Variant, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim lpType As Long
    Dim lpcbData As Long
    Dim lpData As String
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
    If GetHandle = ERROR_SUCCESS Then
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
        GetHandle = RegQueryValueEx(hKey, ValueName, _
            0, lpType, ByVal lpData, lpcbData)
        
'        Dim ary() As Byte
'        Dim cnt As Long
        If GetHandle <> ERROR_SUCCESS Then
            Select Case TypeName(Value)
                Case "Byte()"

'                    ReDim arv(LBound(Value) To UBound(Value)) As Byte
'                    For cnt = LBound(Value) To UBound(Value)
'                        ary(cnt) = Value(cnt)
'                    Next
                    GetHandle = RegSetValueEx(hKey, ValueName, 0&, REG_BINARY, VarPtr(Value), UBound(Value) + IIf(LBound(Value) = 0, 1, 0))
                    
                Case "String"
                    GetHandle = RegSetValueExString(hKey, ValueName, 0, REG_SZ, ByVal CStr(Value), Len(CStr(Value)) + 1)
                Case "Integer"
                    lpcbData = Value
                    GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_DWORD, lpcbData, 4)
                Case "Long"
                    GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_DWORD, Value, 4)
            End Select
        Else
            Select Case lpType
                 Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                     GetHandle = RegSetValueExString(hKey, ValueName, 0, REG_SZ, ByVal CStr(Value), Len(CStr(Value)) + 1)

                 Case REG_DWORD
                    If TypeName(Value) = "Integer" Then
                        GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_DWORD, lpcbData, 4)
                    Else
                    
                        GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_DWORD, Value, 4)
                    End If
                    
                 Case REG_BINARY
                 
                    Dim ary() As Byte
                    Dim cnt As Long
                    
                    ReDim ary(LBound(Value) To UBound(Value)) As Byte
                    
                    For cnt = LBound(Value) To UBound(Value)
                        ary(cnt) = Value(cnt)
                    Next

                     GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_BINARY, ByVal VarPtr(ary(0)), (UBound(ary) - LBound(ary)) + 1)
                 Case Else
                     Err.Raise "11001", "clsRegistry", "Type mismatch"
             End Select

        End If
        RegCloseKey hKey
    Else
    RegCloseKey hKey
        GoTo ErrHand
        
    End If
    
    
    SetValue = True
    
    Exit Function

ErrHand:
    Err.Raise "11001", "clsRegistry", "SetValue " & Err.Description
End Function
Public Function GetValue(ByVal PredefinedKey As Long, ByVal KeyName As String, ByVal ValueName As String, Optional ByVal DefaultValue As Variant, Optional ComputerName As String) As Variant
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim lpData As String
    Dim lpDataDWORD As Long
    Dim lpcbData As Long
    Dim lpType As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    
    
    If ComputerName = "" Then
        lReturnCode = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        lReturnCode = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
        
    If lReturnCode = ERROR_SUCCESS Then
    
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))

        lReturnCode = RegQueryValueEx(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
    
        If lReturnCode <> ERROR_SUCCESS Then
            lpType = REG_BINARY
            lReturnCode = RegQueryValueEx(hKey, ValueName, 0&, lpType, ByVal 0&, lpcbData)
        End If
        
        If lReturnCode = ERROR_SUCCESS Then
            Select Case lpType
                Case REG_SZ
                    lReturnCode = RegQueryValueExString(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
                
                    If lReturnCode = 0 Then
                        GetValue = Left$(lpData, lpcbData - 1)

                    End If
                    
                Case REG_DWORD
                    lReturnCode = RegQueryValueEx(hKey, ValueName, 0, lpType, ByVal VarPtr(lpDataDWORD), lpcbData)
                    
                    If lReturnCode = 0 Then
                        GetValue = lpDataDWORD

                    End If
                Case REG_BINARY
                    
                    
                    Dim byBuffer() As Byte
                    ReDim byBuffer(0 To lpcbData - 1) As Byte

                    lReturnCode = RegQueryValueEx(hKey, ValueName, 0&, REG_BINARY, byBuffer(0), lpcbData)

                    RegCloseKey hKey

                    ReDim Preserve byBuffer(0 To lpcbData - 1) As Byte
                    Dim ary() As Byte
                    Dim cnt As Long

                    ReDim ary(LBound(byBuffer) To UBound(byBuffer)) As Byte

                    For cnt = LBound(byBuffer) To UBound(byBuffer)
                        ary(cnt) = byBuffer(cnt)
                    Next
                    
                    If lReturnCode = 0 Then GetValue = byBuffer

            End Select

        End If

    End If
    RegCloseKey hKey
    
    If lReturnCode <> 0 Then GetValue = DefaultValue
    
    Exit Function

ErrHand:
    Err.Raise "11002", "clsRegistry", "GetValue"
End Function

Public Function SubKey(ByVal PredefinedKey As Long, ByVal KeyName As String, ByVal Index As Long, Optional ByVal DefaultValue As Variant, Optional ComputerName As String) As String
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim dwIndex As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    SubKey = DefaultValue
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
    If GetHandle = ERROR_SUCCESS Then
        lpcbName = 255: lpName = String$(lpcbName, Chr(0))
                    
        GetHandle = RegEnumKeyEx(hKey, Index, lpName, lpcbName, lpReserved, vbNullString, 0, lpftLastWriteTime)
                
        If GetHandle = ERROR_SUCCESS Then
            SubKey = Left$(lpName, lpcbName)
        End If
            
        
    End If
    RegCloseKey hKey
    
    Exit Function

ErrHand:
    Err.Raise "11003", "clsRegistry", "SubKey"
End Function
Public Function DeleteValue(ByVal PredefinedKey As Long, ByVal KeyName As String, ByVal ValueName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If

    If GetHandle = ERROR_SUCCESS Then
        GetHandle = RegDeleteValue(hKey, ValueName)
        If GetHandle = ERROR_SUCCESS Then
            DeleteValue = True
        Else:
            DeleteValue = False
        End If
    End If
    RegCloseKey hKey
    
    Exit Function
ErrHand:
    Err.Raise "11004", "clsRegistry", "DeleteValue"
End Function

Public Function CreateKey(ByVal PredefinedKey As Long, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim hNewKey As Long
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim GetHandle As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long

    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegCreateKeyEx(PredefinedKey, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, GetHandle)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegCreateKeyEx(lhRemoteRegistry, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, GetHandle)
    End If

    If GetHandle = ERROR_SUCCESS Then
        GetHandle = RegCloseKey(hNewKey)
    
        CreateKey = True
    Else
        CreateKey = False
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11005", "clsRegistry", "CreateKey"
End Function

Public Function DeleteKey(ByVal PredefinedKey As Long, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim hKey As Long
    Dim GetHandle As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If

    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, GetFilePath(KeyName), 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, GetFilePath(KeyName), 0, KEY_ALL_ACCESS, hKey)
    End If
    
    If GetHandle = ERROR_SUCCESS Then
        If ComputerName = "" Then
            GetHandle = RegDeleteKey(hKey, GetFileName(KeyName))
        Else
            lReturnCode = RegConnectRegistry(ComputerName, hKey, lhRemoteRegistry)
            GetHandle = RegDeleteKey(lhRemoteRegistry, GetFileName(KeyName))
        End If
        
        
        RegCloseKey hKey
        DeleteKey = True
    Else
        DeleteKey = False
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11006", "clsRegistry", "DeleteKey"
End Function

Public Function SubKeys(ByVal PredefinedKey As Long, ByVal KeyName As String, ByRef DynArray() As String, Optional ComputerName As String) As Long
    Dim ret As String
    Dim cnt As Long
    ReDim DynArray(0 To 0) As String
    DynArray(0) = ""
    cnt = -1
    Do
        cnt = cnt + 1
        ReDim Preserve DynArray(0 To cnt) As String
        DynArray(cnt) = SubKey(PredefinedKey, KeyName, cnt, ComputerName)
    Loop Until DynArray(cnt) = ""
    If cnt > 0 Then
        ReDim Preserve DynArray(0 To cnt - 1) As String
    End If
    SubKeys = cnt - 1
End Function

Public Function SubValues(ByVal PredefinedKey As Long, ByVal KeyName As String, ByRef DynArray() As String, Optional ComputerName As String) As Long
    Dim ret As String
    Dim cnt As Long
    ReDim DynArray(0 To 0) As String
    DynArray(0) = ""
    cnt = -1
    Do
        cnt = cnt + 1
        ReDim Preserve DynArray(0 To cnt) As String
        DynArray(cnt) = SubValue(PredefinedKey, KeyName, cnt, ComputerName)
    Loop Until DynArray(cnt) = ""
    If cnt > 0 Then
        ReDim Preserve DynArray(0 To cnt - 1) As String
    End If
    SubValues = cnt - 1
End Function

Public Function SubValue(ByVal PredefinedKey As Long, ByVal KeyName As String, ByVal Index As Long, Optional ByVal DefaultValue As Variant, Optional ComputerName As String) As String
    On Error GoTo ErrHand

    Dim GetHandle As Long
    Dim hKey As Long
    Dim dwIndex As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lhRemoteRegistry As Long
    Dim lReturnCode As Long

    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
   
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
            
    If GetHandle = ERROR_SUCCESS Then
        lpcbName = 255: lpName = String$(lpcbName, Chr(0))
        
        GetHandle = RegEnumValue(hKey, Index, lpName, lpcbName, 0, ByVal 0&, ByVal 0&, ByVal 0&)

        If GetHandle = ERROR_SUCCESS Then
            SubValue = Left(lpName, lpcbName)
        Else
            SubValue = ""
        End If

        RegCloseKey hKey
    End If

    Exit Function

ErrHand:
    Err.Raise "11007", "clsRegistry", "SubValue"
End Function

Public Function KeyExists(ByVal PredefinedKey As Long, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand

    Dim hKey As Long
    Dim GetHandle As Long
    Dim lhRemoteRegistry As Long
    Dim lReturnCode As Long
    
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1) & Chr(0)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
   

    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long

    
    If GetHandle = ERROR_SUCCESS Then
    
        lReturnCode = RegQueryValueEx(hKey, "", 0&, lValueType, lBuf, lDataBufSize)
        KeyExists = ((Not ((lReturnCode = 6) Or (GetHandle = 6))) And (lValueType <> 0)) Or (lValueType = 0)

    Else
        KeyExists = False
    End If
    
    RegCloseKey hKey
    Exit Function
ErrHand:
    Err.Raise "11008", "clsRegistry", "KeyExists"
End Function

Private Sub Class_Initialize()
   EnablePrivilege SE_MACHINE_ACCOUNT_NAME
End Sub

Public Sub ExpellKey(ByVal RootKey As Long, ByVal SubKey As String)
    'recursive delete of registry key starting with subkeys then values, then itself by parameters
    Dim Info() As String
    If SubKeys(RootKey, SubKey, Info) > -1 Then
        Dim cnt As Long
        For cnt = LBound(Info) To UBound(Info)
            If Info(cnt) <> "" Then
                ExpellKey RootKey, SubKey & "\" & Info(cnt)
            End If
        Next
    End If
    Erase Info
    If SubValues(RootKey, SubKey, Info) > -1 Then
        For cnt = LBound(Info) To UBound(Info)
            If Info(cnt) <> "" Then
                DeleteValue RootKey, SubKey, Info(cnt)
            End If
        Next
    End If
    DeleteKey RootKey, SubKey
End Sub
