Attribute VB_Name = "modRegStrings"
'Option Explicit
'' __________________________________
'' RegStrings Module - RegStrings.bas                     -©Rd-
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' This module creates, reads, writes and deletes Registry keys
'' and their values using the Windows 32-bit Registry API.
''
'' Unlike the internal Registry access methods of VB, it can
'' read and write *any* Registry key with String values.
''
'' The path to store your registry entries should follow the
'' convention of your company name and application name:
''
''     REG_PATH = "Software\Company\AppName"
''
'' Then in your procedure you could code:
''
''     rc = RegSetKey(CURRENT_USER, REG_PATH & "\Settings")
'' _______________
'' Registry Basics
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' Key and subkey names cannot include a backslash (\).
''
'' In the HKEY_CLASSES_ROOT root key, names beginning with a period (.)
'' are reserved for special syntax (filename extensions), but you can
'' include a period within a key name.
''
'' Native data file types must be registered as follows:
''
''     rc = RegSetKey(CLASSES_ROOT, ".tip", "tipfile")
''
'' The name of a subkey must be unique with respect to its parent key.
'' Key names are not localized into other languages, although their
'' values may be.
''
'' Value entries larger than 2048 bytes should be stored as files with
'' just their filenames stored in the registry.
'' _________________________
'' Registry API Declarations
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Declare Function OpenKeyEx Lib "advapi32" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkKeyHandle As Long) As Long
'Private Declare Function CreateKeyEx Lib "advapi32" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkKeyHandle As Long, ByRef lpdwDisposition As Long) As Long
'Private Declare Function QueryKeyInfo Lib "advapi32" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcNumSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcNumValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
'Private Declare Function QueryValueStrEx Lib "advapi32" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpValueBuf As Long, ByRef lpcbBufLen As Long) As Long
'Private Declare Function SetValueStrEx Lib "advapi32" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As Long, ByVal cbValueLen As Long) As Long
'Private Declare Function EnumKeyEx Lib "advapi32" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwKeyIdx As Long, ByVal lpKeyNameBuf As Long, lpcbBufMax As Long, ByVal lpReserved As Long, ByVal lpClass As Long, ByVal lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'Private Declare Function EnumValue Lib "advapi32" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueNameLen As Long, ByVal lpReserved As Long, lpType As Long, lpDataBuf As Any, lpcbDataBufLen As Long) As Long
'Private Declare Function DeleteKey Lib "advapi32" Alias "RegDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
'Private Declare Function DeleteValue Lib "advapi32" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
'Private Declare Function CloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
'' ______________________
'' Set FileTime Structure
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'' ________________________
'' Security Attributes Type
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Type SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type
'' ______________________
'' Reg Root Key Constants
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Enum RegRootHKeys
'    CLASSES_ROOT = &H80000000
'    CURRENT_USER = &H80000001
'    LOCAL_MACHINE = &H80000002
'    ALL_USERS = &H80000003
'    PERFORMANCE_DATA = &H80000004 ' Windows NT only
'    CURRENT_CONFIG = &H80000005
'    DYN_DATA = &H80000006         ' Windows 95 and Windows 98
'End Enum
'' _________________________________
'' Masks for Predefined Access Types
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Const STANDARD_RIGHTS_ALL = &H1F0000  ' STANDARD_RIGHTS_XXX are predefined
'Private Const SPECIFIC_RIGHTS_ALL = &HFFFF&   ' system values used to enforce
'Private Const STANDARD_RIGHTS_READ = &H20000  ' security on system objects
'' _____________________
'' Reg Key Access Rights
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Const KEY_QUERY_VALUE = &H1&          ' Value entries for the key can be read
'Private Const KEY_SET_VALUE = &H2&            ' Value entries for the key can be written
'Private Const KEY_CREATE_SUB_KEY = &H4&       ' Subkeys for the key can be created
'Private Const KEY_ENUMERATE_SUB_KEYS = &H8&   ' All subkeys for the key can be read
'Private Const KEY_NOTIFY = &H10&              ' This flag is irrelevant to device and intermediate drivers, and to other kernel-mode code
'Private Const KEY_CREATE_LINK = &H20&         ' A symbolic link to the key can be created. This flag is irrelvant to device and intermediate drivers
'Private Const READ_CONTROL = &H20000
'Private Const SYNCHRONIZE = &H100000
'Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or READ_CONTROL) And (Not SYNCHRONIZE))
'Private Const KEY_WRITE = (KEY_SET_VALUE Or KEY_CREATE_SUB_KEY)
'Private Const KEY_EXECUTE = KEY_READ
'Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
'              KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or _
'              KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'' _____________________
'' Reg String Data Types
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Private Const REG_SZ = 1&                     ' Unicode nul terminated string
'Private Const REG_EXPAND_SZ = 2&              ' Unicode nul terminated string containing environment variables
'Private Const REG_MULTI_SZ = 7&               ' Multiple Unicode strings
'' ____________
'' Return Codes
'' ¯¯¯¯¯¯¯¯¯¯¯¯
'Private Const ERROR_SUCCESS = 0&
'Private Const ERROR_BADKEY = 2&
'
'' ___________________________________________________________
'' FUNCTION: RegSetKey
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Sub RegSetKey(ByVal hRootKey As RegRootHKeys, PathAndKey As String, Optional DefaultValue As String)
'    ' Just call RegSetValue without a value name
'    RegSetValue hRootKey, PathAndKey, vbNullString, DefaultValue
'End Sub
'
'' ___________________________________________________________
'' FUNCTION: RegSetValue
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Sub RegSetValue(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ValueName As String, ValueData As String)
'    On Error GoTo ExitOut
'
'    Dim Result As Long             ' Return code
'    Dim KeyHandle As Long          ' Handle to Registry key
'    Dim IsNewKey As Long           ' Indicates if new key created
'    Dim tSA As SECURITY_ATTRIBUTES ' Registry Security TYPE
'
'    Const NON_VOLATILE = 0&        ' Key is preserved when system is rebooted
'
'    tSA.lpSecurityDescriptor = 0&  ' Set Security Attributes to defaults...
'    tSA.bInheritHandle = 1&        ' ...
'    tSA.nLength = LenB(tSA)        ' ...
'
'    ' Create/Open //RootKey//PathAndKey
'    Result = CreateKeyEx(hRootKey, StrPtr(PathAndKey), 0&, 0&, NON_VOLATILE, KEY_WRITE, tSA, KeyHandle, IsNewKey)
'    If (Result = ERROR_SUCCESS) Then Else Exit Sub
'
'    ' Create/Modify String value
'    If LenB(ValueData) Then
'        Result = SetValueStrEx(KeyHandle, StrPtr(ValueName), 0&, REG_SZ, StrPtr(ValueData), LenB(ValueData))
'    End If
'
'ExitOut:
'    Result = CloseKey(KeyHandle)   ' Close Registry key
'End Sub
'
'' ___________________________________________________________
'' FUNCTION: RegGetKey
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' Retrieves the unnamed (Default) value for a Registry key.
''
'' If the key exists, but its default value is not set,
'' the function succeeds but returns an empty string.
''
'' Returns: 0 on failure
''          1 on success, default string was retrieved
''         -1 on success, but not string default value
''
''   On success, RetVal is set to the retrieved value.
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Function RegGetKey(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ByRef RetVal As String) As Long
'    ' Just call RegGetValue without a value name
'    RegGetKey = RegGetValue(hRootKey, PathAndKey, vbNullString, RetVal)
'End Function
'
'' ___________________________________________________________
'' FUNCTION: RegGetValue
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' Retrieves the string data for a named value within
'' the specified Registry key.
''
'' If the named value exists, but its data is not set,
'' the function succeeds but does not assign a value.
''
'' Returns: 0 on failure
''          1 on success, value string was retrieved
''         -1 on success, but not string value data
''
''   On success, RetVal is set to the retrieved value.
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Function RegGetValue(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ValueName As String, ByRef RetVal As String) As Long
'    On Error GoTo ExitOut
'
'    Dim Result As Long             ' Return code
'    Dim KeyHandle As Long          ' Handle to Registry key
'    Dim ValueType As Long          ' Data type of Registry value
'    Dim ValueSizeB As Long         ' Length of Registry value
'    Dim ValueBuf As String         ' Registry value buffer
'
'    ' Open Registry key under RootKey
'    Result = OpenKeyEx(hRootKey, StrPtr(PathAndKey), 0&, KEY_READ, KeyHandle)
'
'    ' If the specified key does not exist just fail quietly
'    If (Result = ERROR_SUCCESS) Then Else Exit Function
'
'    ' Get the value type and size in the first call
'    Result = QueryValueStrEx(KeyHandle, StrPtr(ValueName), 0&, ValueType, 0&, ValueSizeB)
'
'    If (Result = ERROR_SUCCESS) Then
'    Else                                 ' Bug fix 26 Aug 2016
'        If ValueName = vbNullString Then
'            ' If the default value data is not set return success
'            If (Result = ERROR_BADKEY) Then RegGetValue = -1
'        End If ' Invalid value name
'        GoTo ExitOut
'    End If
'
'    Select Case ValueType
'        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ ' String
'
'            ' If the value data is not set return success
'            If (ValueSizeB < 2) Then
'                RegGetValue = -1
'                GoTo ExitOut
'            End If
'
'            ValueBuf = String$(ValueSizeB * 0.5, vbNullChar)
'
'            Result = QueryValueStrEx(KeyHandle, StrPtr(ValueName), 0&, ValueType, StrPtr(ValueBuf), ValueSizeB)
'            If (Result = ERROR_SUCCESS) Then Else GoTo ExitOut
'
'            RetVal = LeftB$(ValueBuf, ValueSizeB - 2) ' Remove Null
'            RegGetValue = 1
'
'        Case Else
'            RegGetValue = -1
'    End Select
'ExitOut:
'    Result = CloseKey(KeyHandle)  ' Close Registry key
'End Function
'
'' ___________________________________________________________
'' FUNCTION: RegDeleteKey
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Sub RegDeleteKey(ByVal hRootKey As RegRootHKeys, PathAndKey As String)
'    DeleteKey hRootKey, StrPtr(PathAndKey)
'End Sub
'
'' ___________________________________________________________
'' FUNCTION: RegDeleteValue
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Sub RegDeleteValue(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ValueName As String)
'    On Error GoTo ExitOut
'
'    Dim Result As Long             ' Return code
'    Dim KeyHandle As Long          ' Handle to Registry key
'
'    Result = OpenKeyEx(hRootKey, StrPtr(PathAndKey), 0&, KEY_ALL_ACCESS, KeyHandle)
'    If (Result = ERROR_SUCCESS) Then Else Exit Sub
'
'    Result = DeleteValue(KeyHandle, StrPtr(ValueName))
'ExitOut:
'    Result = CloseKey(KeyHandle)   ' Close Registry key
'End Sub
'
'' ___________________________________________________________
'' FUNCTION: RegGetAllKeys
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' Retrieves all subkeys and their unnamed (Default) values
'' for the specified Registry key.
''
'' The DynArray array must be declared as a dynamic (no size
'' specified) String array, and on success is returned as a
'' zero based two dimensional array.
''
'' If a subkey exists, but its default value is not set,
'' the function succeeds but returns an empty string for
'' that element in the second dimension of the array.
''
'' Returns: Number of subkey(s) found, zero otherwise.
''   On success, DynArray is set to the retrieved subkey
''   names and their corresponding unnamed values.
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Function RegGetAllKeys(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ByRef DynArray() As String) As Long
'    On Error GoTo ExitOut
'
'    Dim Result As Long             ' Return code
'    Dim KeyHandle As Long          ' Handle to Registry key
'    Dim SubKeyBuf As String        ' Temporary storage for Registry subkey
'    Dim ValueBuf As String         ' Temporary storage for Registry key value
'    Dim MaxKeySize As Long         ' Length of longest Registry key name
'    Dim KeySize As Long            ' Length of Registry key/value
'    Dim NumSubKeys As Long         ' Used to ReDim the return array
'    Dim FTime As FILETIME          ' Receives last access info in NT
'    Dim Count As Long              ' Loop counter
'
'    ' Open Registry key under RootKey
'    Result = OpenKeyEx(hRootKey, StrPtr(PathAndKey), 0&, KEY_READ, KeyHandle)
'    If (Result = ERROR_SUCCESS) Then Else Exit Function
'
'    Result = QueryKeyInfo(KeyHandle, 0&, 0&, 0&, NumSubKeys, _
'                          MaxKeySize, 0&, 0&, 0&, 0&, 0&, FTime)
'    If (Result = ERROR_SUCCESS) Then Else GoTo ExitOut
'
'    If (NumSubKeys) Then ' If subkeys exist
'
'        ' ReDim the array used to return subkeys and their values
'        ReDim DynArray(0 To NumSubKeys - 1, 0 To 1) As String
'        MaxKeySize = MaxKeySize + 1   ' Allow for null character
'
'        For Count = 0 To NumSubKeys - 1
'
'            SubKeyBuf = String$(MaxKeySize, vbNullChar)
'            KeySize = MaxKeySize
'
'            Result = EnumKeyEx(KeyHandle, Count, StrPtr(SubKeyBuf), KeySize, 0&, 0&, 0&, FTime)
'            If (Result = ERROR_SUCCESS) Then
'
'                ' Extract SubKey name from Null terminated buffer
'                SubKeyBuf = Left$(SubKeyBuf, KeySize)
'                ValueBuf = vbNullString
'
'                If RegGetValue(KeyHandle, SubKeyBuf, vbNullString, ValueBuf) Then
'                    ' Extract subkey and value from buffers into array
'                    DynArray(Count, 0) = SubKeyBuf
'                    DynArray(Count, 1) = ValueBuf
'                End If
'
'            Else
'                ' Couldn't Enum key, so exit loop
'                Exit For
'            End If
'        Next Count
'    End If
'
'ExitOut:
'    Result = CloseKey(KeyHandle)  ' Close Registry key
'    RegGetAllKeys = Count
'
'End Function
'
'' ___________________________________________________________
'' FUNCTION: RegGetAllValues
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'' Retrieves all named values and string data for the
'' specified Registry key. The passed DynArray array
'' must be a dynamic (no size specified) String array.
''
'' Will also retrieve the unnamed (Default) value if
'' the unnamed (Default) value data has been set. If
'' it does the function returns an empty string for
'' that element in the first dimension of the array
'' and the default value data in the second dimension.
''
'' DynArray is returned as a two dimensional array.
''
'' Returns: Number of value(s) found, zero otherwise.
''   On success, DynArray is set to the retrieved value
''   names and their corresponding string data.
'' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Public Function RegGetAllValues(ByVal hRootKey As RegRootHKeys, PathAndKey As String, ByRef DynArray() As String) As Long
'    On Error GoTo ExitOut
'
'    Dim Result As Long             ' Return code
'    Dim KeyHandle As Long          ' Handle to Registry key
'    Dim ValueType As Long          ' Data type of Registry value
'    Dim ValNameBuf As String       ' Temporary storage for Registry subkey
'    Dim ValDataBuf As String       ' Temporary storage for Registry key value
'    Dim MaxNameSize As Long        ' Length of longest Registry value name
'    Dim NameSize As Long           ' Length of Registry value name
'    Dim NumValues As Long          ' Number of Registry values
'    Dim FTime As FILETIME          ' Receives last access info in NT
'    Dim Count As Long              ' Loop counter
'
'    ' Open Registry key under RootKey
'    Result = OpenKeyEx(hRootKey, StrPtr(PathAndKey), 0&, KEY_READ, KeyHandle)
'    If (Result = ERROR_SUCCESS) Then Else Exit Function
'
'    Result = QueryKeyInfo(KeyHandle, 0&, 0&, 0&, 0&, 0&, 0&, _
'                          NumValues, MaxNameSize, 0&, 0&, FTime)
'    If (Result = ERROR_SUCCESS) Then Else GoTo ExitOut
'
'    If (NumValues) Then
'
'        ' ReDim the array used to return values and their data
'        ReDim DynArray(0 To NumValues - 1, 0 To 1) As String
'        MaxNameSize = MaxNameSize + 1 ' Allow for null character
'
'        For Count = 0 To NumValues - 1
'
'            ValNameBuf = String$(MaxNameSize, vbNullChar)
'            NameSize = MaxNameSize
'
'            Result = EnumValue(KeyHandle, Count, StrPtr(ValNameBuf), NameSize, 0&, ValueType, ByVal 0&, 0&)
'            If (Result = ERROR_SUCCESS) Then
'
'                ' Extract value name from Null terminated buffer
'                ValNameBuf = Left$(ValNameBuf, NameSize)
'                ValDataBuf = vbNullString
'
'                Select Case ValueType
'                    Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ ' String
'                        Result = RegGetValue(hRootKey, PathAndKey, ValNameBuf, ValDataBuf)
'                End Select
'
'                ' Extract value name and data from buffers into array
'                DynArray(Count, 0) = ValNameBuf
'                DynArray(Count, 1) = ValDataBuf
'
'            Else
'                ' Couldn't Enum key, so exit loop
'                Exit For
'            End If
'
'        Next Count
'    End If
'
'ExitOut:
'    Result = CloseKey(KeyHandle)  ' Close Registry key
'    RegGetAllValues = Count
'
'End Function
'
'' ——————————————————————————————————————————————————————————— :›)
'
'
