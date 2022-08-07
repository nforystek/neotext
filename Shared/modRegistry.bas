#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modRegistry"
Option Explicit

Option Compare Binary

Public Const KEY_ALL_ACCESS = &H3F

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_EXPAND_SZ = 2
Public Const REG_MULTI_SZ = 7
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub CreateKey(hKey As Long, strPath As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    If lRegResult <> ERROR_SUCCESS Then
    
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
    Dim lRegResult As Long
    
    lRegResult = RegDeleteKey(hKey, strPath)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    
    lRegResult = RegDeleteValue(hCurKey, strValue)
    
    lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        If lValueType = REG_SZ Then
        
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
        
        End If
    
    Else
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    
    If lRegResult <> ERROR_SUCCESS Then
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
    Dim lValueType As Long
    Dim byBuffer() As Byte
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    Dim hCurKey As Long
    
    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetSettingByte = Default
        Else
            GetSettingByte = 0
        End If
    Else
        GetSettingByte = 0
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        If lValueType = REG_BINARY Then
        
            ReDim byBuffer(lDataBufferSize - 1) As Byte
            lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
            
            GetSettingByte = byBuffer
        
        End If
    
    Else
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)

    Dim lRegResult As Long
    Dim hCurKey As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    
    lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function CheckKey(hKey As Long, strPath As String, ByVal strValueName As String) As String ' this function returns if a valuse exists or not

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        CheckKey = "No"
    Else
        CheckKey = "Yes"
    End If
End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
    
    If lRegResult <> ERROR_SUCCESS Then
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long
    
    If Not IsEmpty(Default) Then
        GetSettingLong = Default
    Else
        GetSettingLong = 0
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4
    
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        If lValueType = REG_DWORD Then
            GetSettingLong = lBuffer
        End If
    
    Else
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)

End Function

Public Function GetSettingExpand(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    
    If Not IsEmpty(Default) Then
        GetSettingExpand = Default
    Else
        GetSettingExpand = ""
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        If lValueType = REG_EXPAND_SZ Then
        
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingExpand = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingExpand = strBuffer
            End If
        
        End If
    
    Else
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingExpand(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_EXPAND_SZ, ByVal strdata, Len(strdata))
    
    If lRegResult <> ERROR_SUCCESS Then
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingMulti(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    
    If Not IsEmpty(Default) Then
        GetSettingMulti = Default
    Else
        GetSettingMulti = ""
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        If lValueType = REG_MULTI_SZ Then
        
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingMulti = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingMulti = strBuffer
            End If
        
        End If
    
    Else
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingMulti(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_MULTI_SZ, ByVal strdata, Len(strdata))
    
    If lRegResult <> ERROR_SUCCESS Then
    
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

