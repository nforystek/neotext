Attribute VB_Name = "modFileInfo"
#Const modFileInfo = -1

'                      Get File Version Info

' Use error handlers in all procedures that make calls to these functions.

' All effort has been made to eliminate errors. Therefore, these functions
' should operate reliably and without any unexpected runtime exceptions, so
' long as you do not pass invalid arguments. If these functions do receive
' invalid arguments then they will indeed raise errors - to let you know.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public UDT used in call to GetVersionInfo function ****

Public Type FILEVERINFO
    FileVer As String
    ProdVer As String
    FileFlags As String
    FileOS As String
    FileType As String
    FileSubtype As String
    Language As String
    Company As String
    FileDesc As String
    Copyright As String
    ProductName As String
    InternalName As String
    OriginalName As String
'    Comments As String
'    LegalCopyright As String
'    LegalTrademarks As String
'    PrivateBuild As String
'    SpecialBuild As String
End Type

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public UDT used in call to GetVersionInfoStruct function ****
' **** Also used to compare file versions using IsNewerVersion  ****

Public Type FIXEDFILEINFO ' VS_FIXEDFILEINFO variant
    Signature As Long
    StructVerl As Integer     '  e.g. = &h0000 = 0
    StructVerh As Integer     '  e.g. = &h0042 = .42
    ' There is data in the first 8 bytes, but is for
    ' Windows internal use and should be ignored
    FileVerPart2 As Integer   '  e.g. = &h0003 = 3
    FileVerPart1 As Integer   '  e.g. = &h0075 = .75
    FileVerPart4 As Integer   '  e.g. = &h0000 = 0
    FileVerPart3 As Integer   '  e.g. = &h0031 = .31
    ProdVerPart2 As Integer   '  e.g. = &h0003 = 3
    ProdVerPart1 As Integer   '  e.g. = &h0010 = .1
    ProdVerPart4 As Integer   '  e.g. = &h0000 = 0
    ProdVerPart3 As Integer   '  e.g. = &h0031 = .31
    FileFlagsMask As Long     '  = &h3F for version "0.42"        - VersionFileFlags
    FileFlags As Long         '  e.g. VFF_DEBUG Or VFF_PRERELEASE - VersionFileFlags
    FileOS As Long            '  e.g. VOS_DOS_WINDOWS16           - VersionOperatingSystemTypes
    FileType As Long          '  e.g. VFT_DRIVER
    FileSubtype As Long       '  e.g. VFT2_DRV_KEYBOARD           - VersionFileSubTypes
    ' I've never seen any data in the following two dwords
    FileDateMS As Long        '  e.g. 0                           - DateHighPart
    FileDateLS As Long        '  e.g. 0                           - DateLowPart
End Type

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lpszFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lpszFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VersionQueryValue Lib "Version" Alias "VerQueryValueA" (lpBlock As Any, ByVal lpSubBlock As String, lpBufPtr As Long, lBufLen As Long) As Long

Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

Private Declare Function ShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBufLen As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVerInfo As OSVERSIONINFO) As Long

Private Declare Function pStrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal lpszDest As String, ByVal lpSrc As Long) As Long
Private Declare Function pStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal length As Long)

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MAXSHORT = 128
Private Const MAXLONG = 260

Private Type OSVERSIONINFO
    dwSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * MAXSHORT
End Type

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ----- VS_VERSION.dwFileFlags -----

Private Const ffi_SIGNATURE = &HFEEF04BD
Private Const ffi_STRUCVERSION = &H10000
Private Const ffi_FILEFLAGSMASK = &H3F&

' ----- VS_VERSION.FileFlags -----

Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

' ----- VS_VERSION.FileOS -----

Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000

Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

' ----- VS_VERSION.FileType -----

Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

' ----- VS_VERSION.FileSubtype for VFT_WINDOWS_DRV -----

Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public Get-Version-Info Function ****

Public Function GetVersionInfo(sFileSpec As String, fvi As FILEVERINFO) As Boolean
    If Len(sFileSpec) = 0 Then Err.Raise 5
    If Len(Dir$(sFileSpec)) = 0 Then Err.Raise 53

    If (IsWin9x) Then If (Not ValidLen(sFileSpec)) Then GoTo Fail

    Dim ffi As FIXEDFILEINFO
    Dim lRet         As Long
    Dim lDummy       As Long
    Dim ayBlock()    As Byte
    Dim ayBuf()      As Byte
    Dim lBufferLen   As Long
    Dim lVerPointer  As Long
    Dim lLangId      As Long
    Dim sTemp        As String
    Dim sLangCharset As String

    On Error GoTo Fail

    '**** Get size ****
    lBufferLen = GetFileVersionInfoSize(sFileSpec, lDummy)
    If (lBufferLen = 0) Then GoTo Fail ' No Version Info available

    '**** Create Version Info struct ****
    ReDim ayBlock(lBufferLen) As Byte
    lRet = GetFileVersionInfo(sFileSpec, 0&, lBufferLen, ayBlock(0))
    If (lRet = 0) Then GoTo Fail ' GetFileVersionInfo failed

    ' Note - ayBlock cannot be accessed directly, use VerQueryValue
    lRet = VersionQueryValue(ayBlock(0), "\", lVerPointer, lBufferLen)
    If (lRet = 0) Then GoTo Fail ' VerQueryValue failed

    '**** Store info to ffi struct ****
    CopyMemory ffi, ByVal lVerPointer, Len(ffi)

    '**** Determine File Version number ****
    fvi.FileVer = Format$(ffi.FileVerPart1) & "." & Format$(ffi.FileVerPart2) & "." & _
                  Format$(ffi.FileVerPart3) & "." & Format$(ffi.FileVerPart4)

    '**** Determine Product Version number ****
    fvi.ProdVer = Format$(ffi.ProdVerPart1) & "." & Format$(ffi.ProdVerPart2) & "." & _
                  Format$(ffi.ProdVerPart3) & "." & Format$(ffi.ProdVerPart4)

    '**** Determine Boolean attributes of File ****
    fvi.FileFlags = vbNullString
    If ffi.FileFlags And VS_FF_DEBUG Then fvi.FileFlags = "Debug "
    If ffi.FileFlags And VS_FF_PRERELEASE Then fvi.FileFlags = fvi.FileFlags & "PreRel "
    If ffi.FileFlags And VS_FF_PATCHED Then fvi.FileFlags = fvi.FileFlags & "Patched "
    If ffi.FileFlags And VS_FF_PRIVATEBUILD Then fvi.FileFlags = fvi.FileFlags & "Private "
    If ffi.FileFlags And VS_FF_INFOINFERRED Then fvi.FileFlags = fvi.FileFlags & "Info "
    If ffi.FileFlags And VS_FF_SPECIALBUILD Then fvi.FileFlags = fvi.FileFlags & "Special "
    If ffi.FileFlags And VFT2_UNKNOWN Then fvi.FileFlags = fvi.FileFlags + "Unknown "

    '**** Determine OS for which file was designed ****
    Select Case ffi.FileOS
        Case VOS_DOS_WINDOWS16: fvi.FileOS = "DOS-Win16"
        Case VOS_DOS_WINDOWS32: fvi.FileOS = "DOS-Win32"
        Case VOS_OS216_PM16:    fvi.FileOS = "OS/2-16 PM-16"
        Case VOS_OS232_PM32:    fvi.FileOS = "OS/2-16 PM-32"
        Case VOS_NT_WINDOWS32:  fvi.FileOS = "NT-Win32"
        Case Else:              fvi.FileOS = "Unknown"
    End Select

    Select Case ffi.FileType
        Case VFT_APP:        fvi.FileType = "App"
        Case VFT_DLL:        fvi.FileType = "DLL"
        Case VFT_DRV:        fvi.FileType = "Driver"
            Select Case ffi.FileSubtype
                Case VFT2_DRV_PRINTER:     fvi.FileSubtype = "Printer drv"
                Case VFT2_DRV_KEYBOARD:    fvi.FileSubtype = "Keyboard drv"
                Case VFT2_DRV_LANGUAGE:    fvi.FileSubtype = "Language drv"
                Case VFT2_DRV_DISPLAY:     fvi.FileSubtype = "Display drv"
                Case VFT2_DRV_MOUSE:       fvi.FileSubtype = "Mouse drv"
                Case VFT2_DRV_NETWORK:     fvi.FileSubtype = "Network drv"
                Case VFT2_DRV_SYSTEM:      fvi.FileSubtype = "System drv"
                Case VFT2_DRV_INSTALLABLE: fvi.FileSubtype = "Installable"
                Case VFT2_DRV_SOUND:       fvi.FileSubtype = "Sound drv"
                Case VFT2_DRV_COMM:        fvi.FileSubtype = "Comm drv"
                Case VFT2_UNKNOWN:         fvi.FileSubtype = "Unknown"
            End Select
        Case VFT_FONT:       fvi.FileType = "Font"
            Select Case ffi.FileSubtype
                Case VFT_FONT_RASTER:      fvi.FileSubtype = "Raster Font"
                Case VFT_FONT_VECTOR:      fvi.FileSubtype = "Vector Font"
                Case VFT_FONT_TRUETYPE:    fvi.FileSubtype = "TrueType Font"
            End Select
        Case VFT_VXD:        fvi.FileType = "VxD"
        Case VFT_STATIC_LIB: fvi.FileType = "Lib"
        Case Else:           fvi.FileType = "Unknown"
    End Select

    lRet = VersionQueryValue(ayBlock(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
    If (lRet = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen = 0) Then GoTo Fail ' Specified sub-block does not exist

    ReDim ayBuf(lBufferLen) As Byte
    CopyMemory ayBuf(0), ByVal lVerPointer, lBufferLen

    'lVerPointer is a pointer to four bytes of Hex number, the
    'first two bytes are the language id, and the last two bytes
    'are the code page.

    CopyMemory lLangId, ayBuf(0), 2
    sTemp = String$(MAXLONG, vbNullChar)
    lRet = VerLanguageName(lLangId, sTemp, MAXLONG)

    fvi.Language = Left$(sTemp, InStr(sTemp, vbNullChar) - 1)

    'However, sLangCharset needs a string of 4 hex digits, the
    'first two characters correspond to the language id and the
    'last two characters correspond to the code page id.

    sLangCharset = Hex$(ayBuf(2) + ayBuf(3) * &H100 + ayBuf(0) * &H10000 + ayBuf(1) * &H1000000)

    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    '----04E4 = 1252 = Codepage for Windows:Multilingual
    Do While Len(sLangCharset) < 8
        sLangCharset = "0" & sLangCharset
    Loop

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\CompanyName", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.Company = PointerToStr(lVerPointer)
    End If

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\FileDescription", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.FileDesc = PointerToStr(lVerPointer)
    End If

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\LegalCopyright", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.Copyright = PointerToStr(lVerPointer)
    End If

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\ProductName", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail
    ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.ProductName = PointerToStr(lVerPointer)
    End If

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\InternalName", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.InternalName = PointerToStr(lVerPointer)
    End If

    rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\OriginalFilename", lVerPointer, lBufferLen)
    If (rc = 0) Then GoTo Fail ' VerQueryValue failed

    If (lBufferLen <> 0) Then ' If specified sub-block exists
        fvi.OriginalName = PointerToStr(lVerPointer)
    End If

'rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\Comments", lVerPointer, lBufferLen)
'rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\LegalCopyright", lVerPointer, lBufferLen)
'rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\LegalTrademarks", lVerPointer, lBufferLen)
'rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\PrivateBuild", lVerPointer, lBufferLen)
'rc = VersionQueryValue(ayBlock(0), "\StringFileInfo\" & sLangCharset & "\SpecialBuild", lVerPointer, lBufferLen)
'If (rc = 0) Then Exit Function ' VerQueryValue failed
'If (lBufferLen <> 0) Then ' If specified sub-block exists
'    fvi.Comments = PointerToStr(lVerPointer)
'End If

    GetVersionInfo = True
Fail:
    If fvi.FileVer = "" Then fvi.FileVer = "0.0.0.0"
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public Get-Version-Info-Structure Function ****

' Returns False if no version information is availabe.

Public Function GetVersionInfoStruct(sFileSpec As String, ffi As FIXEDFILEINFO) As Boolean
    If Len(sFileSpec) = 0 Then Err.Raise 5
    If Len(Dir$(sFileSpec)) = 0 Then Err.Raise 53

    If (IsWin9x) Then If (Not ValidLen(sFileSpec)) Then Exit Function

    Dim lRet         As Long
    Dim lDummy       As Long
    Dim ayBlock()    As Byte
    Dim lBufferLen   As Long
    Dim lVerPointer  As Long

    On Error GoTo Fail

    '**** Get size ****
    lBufferLen = GetFileVersionInfoSize(sFileSpec, lDummy)
    If (lBufferLen = 0) Then Exit Function ' No Version Info available

    '**** Create Version Info struct ****
    ReDim ayBlock(lBufferLen) As Byte
    lRet = GetFileVersionInfo(sFileSpec, 0&, lBufferLen, ayBlock(0))
    If (lRet = 0) Then Exit Function ' GetFileVersionInfo failed

    ' Note - ayBlock cannot be accessed directly, use VerQueryValue
    lRet = VersionQueryValue(ayBlock(0), "\", lVerPointer, lBufferLen)
    If (lRet = 0) Then Exit Function ' VerQueryValue failed

    '**** Store info to ffi struct ****
    CopyMemory ffi, ByVal lVerPointer, Len(ffi)

    GetVersionInfoStruct = True
Fail:
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public Is-Newer-Version Function ****

' Compares two file version structures returned by GetVersionInfoStruct
' and determines whether the source file version is newer (greater) than
' the destination file version.

' Important - this function will raise error 13 if the files are not the
' same FileOS and FileType. This function tests these file attributes to
' assert that these files 'could be' the same file.

' It is the callers responsibility to determine with greater accuracy
' that these two files are indeed the same file. This can be easily
' achieved by calling the GetVersionInfo function and testing various
' file attributes such as OriginalName, InternalName, FileOS, FileType,
' Language and Company.

' If these attributes match but the FileVer string does not you can use
' this function in conjunction with GetVersionInfoStruct to determine
' which file is newer.

Public Function IsNewerVersion(SrcVer As FIXEDFILEINFO, ThanDestVer As FIXEDFILEINFO) As Boolean
   With SrcVer
      If .FileOS <> ThanDestVer.FileOS Then Err.Raise 13
      If .FileType <> ThanDestVer.FileType Then Err.Raise 13

      If .FileVerPart1 > ThanDestVer.FileVerPart1 Then GoTo Newer
      If .FileVerPart1 < ThanDestVer.FileVerPart1 Then Exit Function

      If .FileVerPart2 > ThanDestVer.FileVerPart2 Then GoTo Newer
      If .FileVerPart2 < ThanDestVer.FileVerPart2 Then Exit Function

      If .FileVerPart3 > ThanDestVer.FileVerPart3 Then GoTo Newer
      If .FileVerPart3 < ThanDestVer.FileVerPart3 Then Exit Function

      If .FileVerPart4 > ThanDestVer.FileVerPart4 Then GoTo Newer
   End With
   Exit Function
Newer:
   IsNewerVersion = True
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Private Support Functions ****

Private Function PointerToStr(ByVal pStr As Long) As String
    Dim lLen As Long, sTemp As String
    lLen = pStrLen(pStr)
    sTemp = String$(lLen + 1, vbNullChar)
    pStrToStr sTemp, pStr
    PointerToStr = Left$(sTemp, lLen) 'B.Mc
End Function

Private Function ValidLen(sLongPath As String) As Boolean
    Dim rc As Long, sPath As String
    sPath = String$(MAXLONG, vbNullChar)
    rc = ShortPathName(sLongPath, sPath, MAXLONG)
    If (rc) Then ValidLen = (InStr(sPath, vbNullChar) <= MAXSHORT)
End Function

Private Function IsWin9x() As Boolean
    Dim osvi As OSVERSIONINFO: osvi.dwSize = Len(osvi)
    If (GetVersionEx(osvi)) Then IsWin9x = (osvi.dwPlatformId And &H1&)
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' GetFileVersionInfoSize - Returns size of version information in bytes.

' Determines whether the OS can obtain version information about a specified file.
' If so, it returns the size, in bytes, required to recieve that information.

' lpszFilename - Pointer to null-terminated filename string. The short path form
'                of filename must be less than 126 characters for Win95/98/Me.
' lpdwHandle   - Pointer to variable that the function sets to zero.

' If the function fails, the return value is zero. On error, call Err.LastDllError.

' Call the GetFileVersionInfoSize function before calling the GetFileVersionInfo
' function. The size returned by GetFileVersionInfoSize indicates the buffer size
' required for the version information returned by GetFileVersionInfo.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' GetFileVersionInfo - Reads version information into buffer.

' Creates a version information structure for the specified file, and returns
' it in the lpData member.

' lpszFilename - Pointer to null-terminated filename string. The short path form
'                of filename must be less than 126 characters for Win95/98/Me.
' dwHandle     - This parameter is ignored.
' dwLen        - Specifies size in bytes of the buffer pointed to by lpData. Should be
'                equal to or greater than the value returned by GetFileVersionInfoSize.
' lpData       - Pointer to buffer to receive file-version info. You can use this value
'                in a subsequent call to VerQueryValue to retrieve data from the buffer.

' If the buffer pointed to by lpData is not large enough, the function truncates
' the file's version information to the size of the buffer (actually, to the size
' specified by the dwLen member).

' The file version information is always in Unicode format.

' If the function succeeds, the return value is nonzero. If the function fails,
' the return value is zero. To get extended error information, call Err.LastDllError.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' VerQueryValue - Returns selected version information from version-info resource.

' Returns selected version information from the specified version-info resource.

' lpBlock    - Pointer to the buffer containing the version-information resource
'              returned in the lpData member of the GetFileVersionInfo function.
' lpSubBlock - Address of value to retrieve. Pointer to a zero-terminated string
'              specifying which version-information value to retrieve. The string
'              must consist of names separated by backslashes (\) and it must have
'              one of the following forms:
'              \                                         - Specifies the root block. The function retrieves a pointer to the FIXEDFILEINFO structure for the version-information resource.
'              \VarFileInfo\Translation                  - Specifies the translation array in a Var variable information structure. The function retrieves a pointer to an array of language and code page identifiers. An application can use these identifiers to access a language-specific StringTable structure in the version-information resource.
'              \StringFileInfo\lang-codepage\string-name - Specifies a value in a language-specific StringTable structure. The lang-codepage name is a concatenation of a language and code page identifier pair found as a DWORD in the translation array for the resource. Here the lang-codepage name must be specified as a hexadecimal string. The string-name name must be one of the predefined strings described in the following Remarks section. The function retrieves a string value specific to the language and code page indicated.
' lpBufPtr   - Address of buffer for version value pointer. Pointer to a variable
'              that receives a pointer to the requested version information in the
'              buffer pointed to by lpBlock. The memory pointed to by lpBufPtr is
'              freed when the associated lpBlock memory is freed.
' lBufLen    - Pointer to a buffer that receives the length, in characters, of the
'              version-information value.

' If the specified version-information structure exists, and version information is
' available, the return value is nonzero. If the specified name does not exist or the
' specified resource is not valid, the return value is zero.

' If no value is available for the specified version-information name, the address
' of the length buffer is zero.

' Remarks
' The Win32 API contains the following predefined version information Unicode strings:
'     CompanyName
'     FileDescription
'     FileVersion
'     InternalName
'     LegalCopyright
'     OriginalFilename
'     ProductName
'     ProductVersion
'     Comments
'     LegalTrademarks
'     PrivateBuild
'     SpecialBuild

' The following example shows how to retrieve the FileDescription string-value from
' a block of version information, if the language is U.S. English and the code page
' is Windows Multilingual:

' rc = VerQueryValue(pBlock, "\StringFileInfo\040904E4\FileDescription", lpBuffer, dwBytes)

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' VerLanguageName - Retrieves a description string for the language associated
'                   with a specified binary Microsoft language identifier.

' wLang  - Specifies the binary Microsoft language identifier. For a complete list
'          of the language identifiers supported by Win32, see Language Identifiers.
'          If the identifier is unknown, the szLang parameter points to a default
'          string "Language Neutral".
' szLang - Pointer to the buffer to receive the null-terminated string representing
'          the language specified by the wLang parameter.
' nSize  - Indicates the size of the buffer, in characters, pointed to by szLang.

' If the return value is less than or equal to the buffer size, the return value is
' the size, in characters, of the string returned in the buffer.

' Note - this value does not include the terminating null character.

' If the return value is greater than the buffer size, the return value is the size
' of the buffer required to hold the entire string. The string is truncated to the
' length of the existing buffer.

' If an error occurs, the return value is zero. Unknown language identifiers do not
' produce errors.

' Typically, an installation program uses this function to translate a language
' identifier returned by the VerQueryValue function. The text string may be used
' in a dialog box that asks the user how to proceed in the event of a language
' conflict.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


