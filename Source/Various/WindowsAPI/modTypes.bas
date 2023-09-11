Attribute VB_Name = "modTypes"
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

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type POINTL
        X As Long
        Y As Long
End Type
Public Type Size
        cx As Long
        cy As Long
End Type
Public Type POINTS
        X  As Integer
        Y  As Integer
End Type
Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public Type SID_IDENTIFIER_AUTHORITY
        Value(6) As Byte
End Type
Public Type SID_AND_ATTRIBUTES
        Sid As Long
        Attributes As Long
End Type
Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SystemTime
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Public Type COMMPROP
        wPacketLength As Integer
        wPacketVersion As Integer
        dwServiceMask As Long
        dwReserved1 As Long
        dwMaxTxQueue As Long
        dwMaxRxQueue As Long
        dwMaxBaud As Long
        dwProvSubPublic Type As Long
        dwProvCapabilities As Long
        dwSettableParams As Long
        dwSettableBaud As Long
        wSettableData As Integer
        wSettableStopParity As Integer
        dwCurrentTxQueue As Long
        dwCurrentRxQueue As Long
        dwProvSpec1 As Long
        dwProvSpec2 As Long
        wcProvChar(1) As Integer
End Type
'Public Type COMSTAT
'        fCtsHold As Long
'        fDsrHold As Long
'        fRlsdHold As Long
'        fXoffHold As Long
'        fXoffSent As Long
'        fEof As Long
'        fTxim As Long
'        fReserved As Long
'        cbInQue As Long
'        cbOutQue As Long
'End Type

Public Type COMSTAT
        fBitFields As Long 'See Comment in Win32API.Txt
        cbInQue As Long
        cbOutQue As Long
End Type
'Public Type DCB
'        DCBlength As Long
'        BaudRate As Long
'        fBinary As Long
'        fParity As Long
'        fOutxCtsFlow As Long
'        fOutxDsrFlow As Long
'        fDtrControl As Long
'        fDsrSensitivity As Long
'        fTXContinueOnXoff As Long
'        fOutX As Long
'        fInX As Long
'        fErrorChar As Long
'        fNull As Long
'        fRtsControl As Long
'        fAbortOnError As Long
'        fDummy2 As Long
'        wReserved As Integer
'        XonLim As Integer
'        XoffLim As Integer
'        ByteSize As Byte
'        Parity As Byte
'        StopBits As Byte
'        XonChar As Byte
'        XoffChar As Byte
'        ErrorChar As Byte
'        EofChar As Byte
'        EvtChar As Byte
'End Type

Public Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long 'See Comments in Win32API.Txt
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer 'Reserved; Do Not Use
End Type
' The fourteen actual DCB bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName             Bit #     Description
' -----------------     -----     ------------------------------
' fBinary                 1       binary mode, no EOF check
' fParity                 2       enable parity checking
' fOutxCtsFlow            3       CTS output flow control
' fOutxDsrFlow            4       DSR output flow control
' fDtrControl             5       DTR flow control Public Type (2 bits)
' fDsrSensitivity         7       DSR sensitivity
' fTXContinueOnXoff       8       XOFF continues Tx
' fOutX                   9       XON/XOFF out flow control
' fInX                   10       XON/XOFF in flow control
' fErrorChar             11       enable error replacement
' fNull                  12       enable null stripping
' fRtsControl            13       RTS flow control (2 bits)
' fAbortOnError          15       abort reads/writes on error
' fDummy2                16       reserved

Public Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type
Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorPublic Type As Long
        dwAllocationGranularity As Long
        dwReserved As Long
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
'   Define the generic mapping array.  This is used to denote the
'   mapping of each generic access right to a specific access mask.

Public Type GENERIC_MAPPING
        GenericRead As Long
        GenericWrite As Long
        GenericExecute As Long
        GenericAll As Long
End Type
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                         LUID_AND_ATTRIBUTES                         //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'

Public Type Luid
        lowpart As Long
        highpart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
        pLuid As Luid
        Attributes As Long
End Type
Public Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
' typedef ACL *PACL;
'  end_ntddk
'   The structure of an ACE is a common ace header followed by ace type
'   specific data.  Pictorally the structure of the common ace header is
'   as follows:
'   AcePublic Type denotes the Public Type of the ace, there are some predefined ace
'   types
'
'   AceSize is the size, in bytes, of ace.
'
'   AceFlags are the Ace flags for audit and inheritance, defined Integerly.

Public Type ACE_HEADER
        AcePublic Type As Byte
        AceFlags As Byte
        AceSize As Long
End Type
'
'   We'll define the structure of the predefined ACE types.  Pictorally
'   the structure of the predefined ACE's is as follows:
'   Mask is the access mask associated with the ACE.  This is either the
'   access allowed, access denied, audit, or alarm mask.
'
'   Sid is the Sid associated with the ACE.
'
'   The following are the four predefined ACE types.
'   Examine the AcePublic Type field in the Header to determine
'   which structure is appropriate to use for casting.

Public Type ACCESS_ALLOWED_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type
Public Type ACCESS_DENIED_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type
Public Type SYSTEM_AUDIT_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type
Public Type SYSTEM_ALARM_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type
'
'   This record is returned/sent if the user is requesting/setting the
'   AclRevisionInformation
'

Public Type ACL_REVISION_INFORMATION
        AclRevision As Long
End Type
'
'   This record is returned if the user is requesting AclSizeInformation
'

Public Type ACL_SIZE_INFORMATION
        AceCount As Long
        AclBytesInUse As Long
        AclBytesFree As Long
End Type
'
'   Where:
'
'       SE_OWNER_DEFAULTED - This boolean flag, when set, indicates that the
'           SID pointed to by the Owner field was provided by a
'           defaulting mechanism rather than explicitly provided by the
'           original provider of the security descriptor.  This may
'           affect the treatment of the SID with respect to inheritence
'           of an owner.
'
'       SE_GROUP_DEFAULTED - This boolean flag, when set, indicates that the
'           SID in the Group field was provided by a defaulting mechanism
'           rather than explicitly provided by the original provider of
'           the security descriptor.  This may affect the treatment of
'           the SID with respect to inheritence of a primary group.
'
'       SE_DACL_PRESENT - This boolean flag, when set, indicates that the
'           security descriptor contains a discretionary ACL.  If this
'           flag is set and the Dacl field of the SECURITY_DESCRIPTOR is
'           null, then a null ACL is explicitly being specified.
'
'       SE_DACL_DEFAULTED - This boolean flag, when set, indicates that the
'           ACL pointed to by the Dacl field was provided by a defaulting
'           mechanism rather than explicitly provided by the original
'           provider of the security descriptor.  This may affect the
'           treatment of the ACL with respect to inheritence of an ACL.
'           This flag is ignored if the DaclPresent flag is not set.
'
'       SE_SACL_PRESENT - This boolean flag, when set,  indicates that the
'           security descriptor contains a system ACL pointed to by the
'           Sacl field.  If this flag is set and the Sacl field of the
'           SECURITY_DESCRIPTOR is null, then an empty (but present)
'           ACL is being specified.
'
'       SE_SACL_DEFAULTED - This boolean flag, when set, indicates that the
'           ACL pointed to by the Sacl field was provided by a defaulting
'           mechanism rather than explicitly provided by the original
'           provider of the security descriptor.  This may affect the
'           treatment of the ACL with respect to inheritence of an ACL.
'           This flag is ignored if the SaclPresent flag is not set.
'
'       SE_SELF_RELATIVE - This boolean flag, when set, indicates that the
'           security descriptor is in self-relative form.  In this form,
'           all fields of the security descriptor are contiguous in memory
'           and all pointer fields are expressed as offsets from the
'           beginning of the security descriptor.  This form is useful
'           for treating security descriptors as opaque data structures
'           for transmission in communication protocol or for storage on
'           secondary media.
'
'
'
'  In general, this data structure should be treated opaquely to ensure future
'  compatibility.
'
'

Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
'
'   Privilege Set - This is defined for a privilege set of one.
'                   If more than one privilege is needed, then this structure
'                   will need to be allocated with more space.
'
'   Note: don't change this structure without fixing the INITIAL_PRIVILEGE_SET
'   structure (defined in se.h)
'

Public Type PRIVILEGE_SET
        PrivilegeCount As Long
        Control As Long
        Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Public Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type
Public Type CREATE_THREAD_DEBUG_INFO
        hThread As Long
        lpThreadLocalBase As Long
        lpStartAddress As Long
End Type
Public Type CREATE_PROCESS_DEBUG_INFO
        hFile As Long
        hProcess As Long
        hThread As Long
        lpBaseOfImage As Long
        dwDebugInfoFileOffset As Long
        nDebugInfoSize As Long
        lpThreadLocalBase As Long
        lpStartAddress As Long
        lpImageName As Long
        fUnicode As Integer
End Type
Public Type EXIT_THREAD_DEBUG_INFO
        dwExitCode As Long
End Type
Public Type EXIT_PROCESS_DEBUG_INFO
        dwExitCode As Long
End Type
Public Type LOAD_DLL_DEBUG_INFO
        hFile As Long
        lpBaseOfDll As Long
        dwDebugInfoFileOffset As Long
        nDebugInfoSize As Long
        lpImageName As Long
        fUnicode As Integer
End Type
Public Type UNLOAD_DLL_DEBUG_INFO
        lpBaseOfDll As Long
End Type
Public Type OUTPUT_DEBUG_STRING_INFO
        lpDebugStringData As String
        fUnicode As Integer
        nDebugStringLength As Integer
End Type
Public Type RIP_INFO
        dwError As Long
        dwPublic Type As Long
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
Public Type CRITICAL_SECTION
        dummy As Long
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
Public Type MEMORY_BASIC_INFORMATION
     BaseAddress As Long
     AllocationBase As Long
     AllocationProtect As Long
     RegionSize As Long
     State As Long
     Protect As Long
     lPublic Type As Long
End Type
Public Type EVENTLOGRECORD
     Length As Long     '  Length of full record
     Reserved As Long     '  Used by the service
     RecordNumber As Long     '  Absolute record number
     TimeGenerated As Long     '  Seconds since 1-1-1970
     TimeWritten As Long     'Seconds since 1-1-1970
     EventID As Long
     EventPublic Type As Integer
     NumStrings As Integer
     EventCategory As Integer
     ReservedFlags As Integer     '  For use with paired events (auditing)
     ClosingRecordNumber As Long     'For use with paired events (auditing)
     StringOffset As Long     '  Offset from beginning of record
     UserSidLength As Long
     UserSidOffset As Long
     DataLength As Long
     DataOffset As Long     '  Offset from beginning of record
End Type
Public Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type
Public Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Public Type CONTEXT
        FltF0 As Double
        FltF1 As Double
        FltF2 As Double
        FltF3 As Double
        FltF4 As Double
        FltF5 As Double
        FltF6 As Double
        FltF7 As Double
        FltF8 As Double
        FltF9 As Double
        FltF10 As Double
        FltF11 As Double
        FltF12 As Double
        FltF13 As Double
        FltF14 As Double
        FltF15 As Double
        FltF16 As Double
        FltF17 As Double
        FltF18 As Double
        FltF19 As Double
        FltF20 As Double
        FltF21 As Double
        FltF22 As Double
        FltF23 As Double
        FltF24 As Double
        FltF25 As Double
        FltF26 As Double
        FltF27 As Double
        FltF28 As Double
        FltF29 As Double
        FltF30 As Double
        FltF31 As Double

        IntV0 As Double
        IntT0 As Double
        IntT1 As Double
        IntT2 As Double
        IntT3 As Double
        IntT4 As Double
        IntT5 As Double
        IntT6 As Double
        IntT7 As Double
        IntS0 As Double
        IntS1 As Double
        IntS2 As Double
        IntS3 As Double
        IntS4 As Double
        IntS5 As Double
        IntFp As Double
        IntA0 As Double
        IntA1 As Double
        IntA2 As Double
        IntA3 As Double
        IntA4 As Double
        IntA5 As Double
        IntT8 As Double
        IntT9 As Double
        IntT10 As Double
        IntT11 As Double
        IntRa As Double
        IntT12 As Double
        IntAt As Double
        IntGp As Double
        IntSp As Double
        IntZero As Double

        Fpcr As Double
        SoftFpcr As Double

        Fir As Double
        Psr As Long

        ContextFlags As Long
        Fill(4) As Long
End Type
Public Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
Public Type LDT_BYTES  ' Defined for use in LDT_ENTRY Type
        BaseMid As Byte
        Flags1 As Byte
        Flags2 As Byte
        BaseHi As Byte
End Type
Public Type LDT_ENTRY
        LimitLow As Integer
        BaseLow As Integer
        HighWord As Long        ' Can use LDT_BYTES Type
End Type
Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SystemTime
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SystemTime
        DaylightBias As Long
End Type
' Stream ID type

Public Type WIN32_STREAM_ID
        dwStreamID As Long
        dwStreamAttributes As Long
        dwStreamSizeLow As Long
        dwStreamSizeHigh As Long
        dwStreamNameSize As Long
        cStreamName As Byte
End Type
Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
' *************************************************************************** Typedefs
' *
' * Define all types for the NLS component here.
' \***************************************************************************/
'
'  *  CP Info.
'  */

Public Type CPINFO
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type
Public Type NUMBERFMT
        NumDigits As Long                 '  number of decimal digits
        LeadingZero As Long '  if leading zero in decimal fields
        Grouping As Long '  group size left of decimal
        lpDecimalSep As String              '  ptr to decimal separator string
        lpThousandSep As String             '  ptr to thousand separator string
        NegativeOrder As Long '  negative number ordering
End Type
'
'  *  Currency format.
'  */

Public Type CURRENCYFMT
        NumDigits As Long '  number of decimal digits
        LeadingZero As Long '  if leading zero in decimal fields
        Grouping As Long '  group size left of decimal
        lpDecimalSep As String              '  ptr to decimal separator string
        lpThousandSep As String             '  ptr to thousand separator string
        NegativeOrder As Long '  negative currency ordering
        PositiveOrder As Long '  positive currency ordering
        lpCurrencySymbol As String          '  ptr to currency symbol string
End Type
' The following section contains the Public data structures, data types,
' and procedures exported by the NT console subsystem.

Public Type COORD
        X As Integer
        Y As Integer
End Type
Public Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type
Public Type KEY_EVENT_RECORD
        bKeyDown As Long
        wRepeatCount As Integer
        wVirtualKeyCode As Integer
        wVirtualScanCode As Integer
        uChar As Byte
        dwControlKeyState As Long
End Type
Public Type MOUSE_EVENT_RECORD
        dwMousePosition As COORD
        dwButtonState As Long
        dwControlKeyState As Long
        dwEventFlags As Long
End Type
Public Type WINDOW_BUFFER_SIZE_RECORD
        dwSize As COORD
End Type
Public Type MENU_EVENT_RECORD
        dwCommandId As Long
End Type
Public Type FOCUS_EVENT_RECORD
        bSetFocus As Long
End Type
Public Type CHAR_INFO
        Char As Integer
        Attributes As Integer
End Type
Public Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type
Public Type CONSOLE_CURSOR_INFO
        dwSize As Long
        bVisible As Long
End Type
Public Type xform
        eM11 As Double
        eM12 As Double
        eM21 As Double
        eM22 As Double
        eDx As Double
        eDy As Double
End Type
' Bitmap Header Definition

Public Type BITMAP '14 bytes
        bmPublic Type As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Public Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
' structures for defining DIBs

Public Type BITMAPCOREHEADER '12 bytes
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type
Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Public Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As RGBTRIPLE
End Type
Public Type BITMAPFILEHEADER
        bfPublic Type As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
' Clipboard Metafile Picture Structure

Public Type HANDLETABLE
        objectHandle(1) As Long
End Type
Public Type METARECORD
        rdSize As Long
        rdFunction As Integer
        rdParm(1) As Integer
End Type
Public Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        hMF As Long
End Type
Public Type METAHEADER
        mtPublic Type As Integer
        mtHeaderSize As Integer
        mtVersion As Integer
        mtSize As Long
        mtNoObjects As Integer
        mtMaxRecord As Long
        mtNoParameters As Integer
End Type
Public Type ENHMETARECORD
        iPublic Type As Long
        nSize As Long
        dParm(1) As Long
End Type
Public Type SIZEL
    cx As Long
    cy As Long
End Type
Public Type ENHMETAHEADER
        iPublic Type As Long
        nSize As Long
        rclBounds As RECTL
        rclFrame As RECTL
        dSignature As Long
        nVersion As Long
        nBytes As Long
        nRecords As Long
        nHandles As Integer
        sReserved As Integer
        nDescription As Long
        offDescription As Long
        nPalEntries As Long
        szlDevice As SIZEL
        szlMillimeters As SIZEL
End Type
Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
' Structure passed to FONTENUMPROC
' NOTE: NEWTEXTMETRIC is the same as TEXTMETRIC plus 4 new fields

Public Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type
' GDI Logical Objects:

Public Type PELARRAY
        paXCount As Long
        paYCount As Long
        paXExt As Long
        paYExt As Long
        paRGBs As Integer
End Type
' Logical Brush (or Pattern)

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
' Logical Pen

Public Type LOGPEN
        lopnStyle As Long
        lopnWidth As POINTAPI
        lopnColor As Long
End Type
Public Type EXTLOGPEN
        elpPenStyle As Long
        elpWidth As Long
        elpBrushStyle As Long
        elpColor As Long
        elpHatch As Long
        elpNumEntries As Long
        elpStyleEntry(1) As Long
End Type
Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type
' Logical Palette

Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(1) As PALETTEENTRY
End Type
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Public Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type
Public Type ENUMLOGFONT
        elfLogFont As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
End Type
Public Type PANOSE
        ulculture As Long
        bFamilyPublic Type As Byte
        bSerifStyle As Byte
        bWeight As Byte
        bProportion As Byte
        bContrast As Byte
        bStrokeVariation As Byte
        bArmStyle As Byte
        bLetterform As Byte
        bMidline As Byte
        bXHeight As Byte
End Type
Public Type EXTLOGFONT
        elfLogFont  As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
        elfVersion As Long
        elfStyleSize As Long
        elfMatch As Long
        elfReserved As Long
        elfVendorId(ELF_VENDOR_SIZE) As Byte
        elfCulture As Long
        elfPanose As PANOSE
End Type
Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
Public Type RGNDATAHEADER
        dwSize As Long
        iPublic Type As Long
        nCount As Long
        nRgnSize As Long
        rcBound As RECT
End Type
Public Type RgnData
        rdh As RGNDATAHEADER
        Buffer As Byte
End Type
Public Type ABC
        abcA As Long
        abcB As Long
        abcC As Long
End Type
Public Type ABCFLOAT
        abcfA As Double
        abcfB As Double
        abcfC As Double
End Type
Public Type OUTLINETEXTMETRIC
        otmSize As Long
        otmTextMetrics As TEXTMETRIC
        otmFiller As Byte
        otmPanoseNumber As PANOSE
        otmfsSelection As Long
        otmfsPublic Type As Long
        otmsCharSlopeRise As Long
        otmsCharSlopeRun As Long
        otmItalicAngle As Long
        otmEMSquare As Long
        otmAscent As Long
        otmDescent As Long
        otmLineGap As Long
        otmsCapEmHeight As Long
        otmsXHeight As Long
        otmrcFontBox As RECT
        otmMacAscent As Long
        otmMacDescent As Long
        otmMacLineGap As Long
        otmusMinimumPPEM As Long
        otmptSubscriptSize As POINTAPI
        otmptSubscriptOffset As POINTAPI
        otmptSuperscriptSize As POINTAPI
        otmptSuperscriptOffset As POINTAPI
        otmsStrikeoutSize As Long
        otmsStrikeoutPosition As Long
        otmsUnderscorePosition As Long
        otmsUnderscoreSize As Long
        otmpFamilyName As String
        otmpFaceName As String
        otmpStyleName As String
        otmpFullName As String
End Type
Public Type POLYTEXT
        X As Long
        Y As Long
        n As Long
        lpStr As String
        uiFlags As Long
        rcl As RECT
        pdx As Long
End Type
Public Type FIXED
        fract As Integer
        Value As Integer
End Type
Public Type MAT2
        eM11 As FIXED
        eM12 As FIXED
        eM21 As FIXED
        eM22 As FIXED
End Type
Public Type GLYPHMETRICS
        gmBlackBoxX As Long
        gmBlackBoxY As Long
        gmptGlyphOrigin As POINTAPI
        gmCellIncX As Integer
        gmCellIncY As Integer
End Type
Public Type POINTFX
        X As FIXED
        Y As FIXED
End Type
Public Type TTPOLYCURVE
        wPublic Type As Integer
        cpfx As Integer
        apfx As POINTFX
End Type
Public Type TTPOLYGONHEADER
        cb As Long
        dwPublic Type As Long
        pfxStart As POINTFX
End Type
Public Type RASTERIZER_STATUS
        nSize As Integer
        wFlags As Integer
        nLanguageID As Integer
End Type
Public Type ColorAdjustment
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type
Public Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
End Type
Public Type KERNINGPAIR
        wFirst As Integer
        wSecond As Integer
        iKernAmount As Long
End Type
Public Type emr
        iPublic Type As Long
        nSize As Long
End Type
Public Type emrtext
        ptlReference As POINTL
        nchars As Long
        offString As Long
        fOptions As Long
        rcl As RECTL
        offDx As Long
End Type
Public Type EMRABORTPATH
        pEmr As emr
End Type
Public Type EMRBEGINPATH
        pEmr As emr
End Type
Public Type EMRENDPATH
        pEmr As emr
End Type
Public Type EMRCLOSEFIGURE
        pEmr As emr
End Type
Public Type EMRFLATTENPATH
        pEmr As emr
End Type
Public Type EMRWIDENPATH
        pEmr As emr
End Type
Public Type EMRSETMETARGN
        pEmr As emr
End Type
Public Type EMREMRSAVEDC
        pEmr As emr
End Type
Public Type EMRREALIZEPALETTE
        pEmr As emr
End Type
Public Type EMRSELECTCLIPPATH
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETBKMODE
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETMAPMODE
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETPOLYFILLMODE
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETROP2
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETSTRETCHBLTMODE
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETTEXTALIGN
        pEmr As emr
        iMode As Long
End Type
Public Type EMRSETMITERLIMIT
        pEmr As emr
        eMiterLimit As Double
End Type
Public Type EMRRESTOREDC
        pEmr As emr
        iRelative As Long
End Type
Public Type EMRSETARCDIRECTION
        pEmr As emr
        iArcDirection As Long
End Type
Public Type EMRSETMAPPERFLAGS
        pEmr As emr
        dwFlags As Long
End Type
Public Type EMRSETTEXTCOLOR
        pEmr As emr
        crColor As Long
End Type
Public Type EMRSETBKCOLOR
        pEmr As emr
        crColor As Long
End Type
Public Type EMRSELECTOBJECT
        pEmr As emr
        ihObject As Long
End Type
Public Type EMRDELETEOBJECT
        pEmr As emr
        ihObject As Long
End Type
Public Type EMRSELECTPALETTE
        pEmr As emr
        ihPal As Long
End Type
Public Type EMRRESIZEPALETTE
        pEmr As emr
        ihPal As Long
        cEntries As Long
End Type
Public Type EMRSETPALETTEENTRIES
        pEmr As emr
        ihPal As Long
        iStart As Long
        cEntries As Long
        aPalEntries(1) As PALETTEENTRY
End Type
Public Type EMRSETCOLORADJUSTMENT
        pEmr As emr
        ColorAdjustment As ColorAdjustment
End Type
Public Type EMRGDICOMMENT
        pEmr As emr
        cbData As Long
        Data(1) As Integer
End Type
Public Type EMREOF
        pEmr As emr
        nPalEntries As Long
        offPalEntries As Long
        nSizeLast As Long
End Type
Public Type EMRLINETO
        pEmr As emr
        ptl As POINTL
End Type
Public Type EMRMOVETOEX
        pEmr As emr
        ptl As POINTL
End Type
Public Type EMROFFSETCLIPRGN
        pEmr As emr
        ptlOffset As POINTL
End Type
Public Type EMRFILLPATH
        pEmr As emr
        rclBounds As RECTL
End Type
Public Type EMRSTROKEANDFILLPATH
        pEmr As emr
        rclBounds As RECTL
End Type
Public Type EMRSTROKEPATH
        pEmr As emr
        rclBounds As RECTL
End Type
Public Type EMREXCLUDECLIPRECT
        pEmr As emr
        rclClip As RECTL
End Type
Public Type EMRINTERSECTCLIPRECT
        pEmr As emr
        rclClip As RECTL
End Type
Public Type EMRSETVIEWPORTORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type
Public Type EMRSETWINDOWORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type
Public Type EMRSETBRUSHORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type
Public Type EMRSETVIEWPORTEXTEX
        pEmr As emr
        szlExtent As SIZEL
End Type
Public Type EMRSETWINDOWEXTEX
        pEmr As emr
        szlExtent As SIZEL
End Type
Public Type EMRSCALEVIEWPORTEXTEX
        pEmr As emr
        xNum As Long
        xDenom As Long
        yNum As Long
        yDemon As Long
End Type
Public Type EMRSCALEWINDOWEXTEX
        pEmr As emr
        xNum As Long
        xDenom As Long
        yNum As Long
        yDemon As Long
End Type
Public Type EMRSETWORLDTRANSFORM
        pEmr As emr
        xform As xform
End Type
Public Type EMRMODIFYWORLDTRANSFORM
        pEmr As emr
        xform As xform
        iMode As Long
End Type
Public Type EMRSETPIXELV
        pEmr As emr
        ptlPixel As POINTL
        crColor As Long
End Type
Public Type EMREXTFLOODFILL
        pEmr As emr
        ptlStart As POINTL
        crColor As Long
        iMode As Long
End Type
Public Type EMRELLIPSE
        pEmr As emr
        rclBox As RECTL
End Type
Public Type EMRRECTANGLE
        pEmr As emr
        rclBox As RECTL
End Type
Public Type EMRROUNDRECT
        pEmr As emr
        rclBox As RECTL
        szlCorner As SIZEL
End Type
Public Type EMRARC
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type
Public Type EMRARCTO
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type
Public Type EMRCHORD
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type
Public Type EMRPIE
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type
Public Type EMRANGLEARC
        pEmr As emr
        ptlCenter As POINTL
        nRadius As Long
        eStartAngle As Double
        eSweepAngle As Double
End Type
Public Type EMRPOLYLINE
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYBEZIER
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYGON
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYBEZIERTO
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYLINE16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type
Public Type EMRPOLYBEZIER16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type
Public Type EMRPOLYGON16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type
Public Type EMRPLOYBEZIERTO16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type
Public Type EMRPOLYLINETO16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type
Public Type EMRPOLYDRAW
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
        abTypes(1) As Integer
End Type
Public Type EMRPOLYDRAW16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
        abTypes(1) As Integer
End Type
Public Type EMRPOLYPOLYLINE
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cptl As Long
        aPolyCounts(1) As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYPOLYGON
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cptl As Long
        aPolyCounts(1) As Long
        aptl(1) As POINTL
End Type
Public Type EMRPOLYPOLYLINE16
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cpts As Long
        aPolyCounts(1) As Long
        apts(1) As POINTS
End Type
Public Type EMRPOLYPOLYGON16
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cpts As Long
        aPolyCounts(1) As Long
        apts(1) As POINTS
End Type
Public Type EMRINVERTRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        RgnData(1) As Integer
End Type
Public Type EMRPAINTRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        RgnData(1) As Integer
End Type
Public Type EMRFILLRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        ihBrush As Long
        RgnData(1) As Integer
End Type
Public Type EMRFRAMERGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        ihBrush As Long
        szlStroke As SIZEL
        RgnData(1) As Integer
End Type
Public Type EMREXTSELECTCLIPRGN
        pEmr As emr
        cbRgnData As Long
        iMode As Long
        RgnData(1) As Integer
End Type
Public Type EMREXTTEXTOUT
        pEmr As emr
        rclBounds As RECTL
        iGraphicsMode As Long
        exScale As Double
        eyScale As Double
        emrtext As emrtext
End Type
Public Type EMRBITBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
End Type
Public Type EMRSTRETCHBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        cxSrc As Long
        cySrc As Long
End Type
Public Type EMRMASKBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc2 As Long
        cyDest2 As Long
        dwRop2 As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        xMask As Long
        yMask As Long
        iUsageMask As Long
        offBmiMask As Long
        cbBmiMask As Long
        offBitsMask As Long
        cbBitsMask As Long
End Type
Public Type EMRPLGBLT
        pEmr As emr
        rclBounds As RECTL
        aptlDest(3) As POINTL
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        xMask As Long
        yMask As Long
        iUsageMask As Long
        offBmiMask As Long
        cbBmiMask As Long
        offBitsMask As Long
        cbBitsMask As Long
End Type
Public Type EMRSETDIBITSTODEVICE
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        iUsageSrc As Long
        iStartScan As Long
        cScans As Long
End Type
Public Type EMRSTRETCHDIBITS
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        iUsageSrc As Long
        dwRop As Long
        cxDest As Long
        cyDest As Long
End Type
Public Type EMREXTCREATEFONTINDIRECT
        pEmr As emr
        ihFont As Long
        elfw As EXTLOGFONT
End Type
Public Type EMRCREATEPALETTE
        pEmr As emr
        ihPal As Long
        lgpl As LOGPALETTE
End Type
Public Type EMRCREATEPEN
        pEmr As emr
        ihPen As Long
        lopn As LOGPEN
End Type
Public Type EMREXTCREATEPEN
        pEmr As emr
        ihPen As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
        elp As EXTLOGPEN
End Type
Public Type EMRCREATEBRUSHINDIRECT
        pEmr As emr
        ihBrush As Long
        lb As LOGBRUSH
End Type
Public Type EMRCREATEMONOBRUSH
        pEmr As emr
        ihBrush As Long
        iUsage As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
End Type
Public Type EMRCREATEDIBPATTERNBRUSHPT
        pEmr As emr
        ihBursh As Long
        iUsage As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
End Type
Public Type BITMAPV4HEADER
        bV4Size As Long
        bV4Width As Long
        bV4Height As Long
        bV4Planes As Integer
        bV4BitCount As Integer
        bV4V4Compression As Long
        bV4SizeImage As Long
        bV4XPelsPerMeter As Long
        bV4YPelsPerMeter As Long
        bV4ClrUsed As Long
        bV4ClrImportant As Long
        bV4RedMask As Long
        bV4GreenMask As Long
        bV4BlueMask As Long
        bV4AlphaMask As Long
        bV4CSPublic Type As Long
        bV4Endpoints As Long
        bV4GammaRed As Long
        bV4GammaGreen As Long
        bV4GammaBlue As Long
End Type
Public Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type
Public Type CHARSETINFO
        ciCharset As Long
        ciACP As Long
        fs As FONTSIGNATURE
End Type
Public Type LOCALESIGNATURE
        lsUsb(4) As Long
        lsCsbDefault(2) As Long
        lsCsbSupported(2) As Long
End Type
Public Type NEWTEXTMETRICEX
        ntmTm As NEWTEXTMETRIC
        ntmFontSig As FONTSIGNATURE
End Type
Public Type ENUMLOGFONTEX
        elfLogFont As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
        elfScript(LF_FACESIZE) As Byte
End Type
Public Type GCP_RESULTS
        lStructSize As Long
        lpOutString As String
        lpOrder As Long
        lpDx As Long
        lpCaretPos As Long
        lpClass As String
        lpGlyphs As String
        nGlyphs As Long
        nMaxFit As Long
End Type
Public Type CIEXYZ
        ciexyzX As Long
        ciexyzY As Long
        ciexyzZ As Long
End Type
Public Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type
Public Type LOGCOLORSPACE
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSPublic Type As Long
    lcsIntent As Long
    lcsEndPoints As CIEXYZTRIPLE
    lcsGammaRed As Long
    lcsGammaGreen As Long
    lcsGammaBlue As Long
    lcsFileName As String * MAX_PATH
End Type
Public Type EMRSELECTCOLORSPACE
        pEmr As emr
        ihCS As Long               '  ColorSpace handle index
End Type
Public Type EMRCREATECOLORSPACE
        pEmr As emr
        ihCS As Long        '  ColorSpace handle index
        lcs As LOGCOLORSPACE
End Type
' HCBT_ACTIVATE structure pointed to by lParam

Public Type CBTACTIVATESTRUCT
        fMouse As Long
        hWndActive As Long
End Type
' Message Structure used in Journaling

Public Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hWnd As Long
End Type
Public Type CWPSTRUCT
        lParam As Long
        wParam As Long
        message As Long
        hWnd As Long
End Type
Public Type DEBUGHOOKINFO
        hModuleHook As Long
        Reserved As Long
        lParam As Long
        wParam As Long
        code As Long
End Type
Public Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hWnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type
' WM_WINDOWPOSCHANGING/CHANGED struct pointed to by lParam

Public Type WINDOWPOS
        hWnd As Long
        hWndInsertAfter As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        Flags As Long
End Type
Public Type ACCEL
        fVirt As Byte
        key As Integer
        cmd As Integer
End Type
Public Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type
Public Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hwndParent As Long
        cy As Long
        cx As Long
        Y As Long
        X As Long
        style As Long
        lpszName As String
        lpszClass As String
        ExStyle As Long
End Type
' HCBT_CREATEWND parameters pointed to by lParam

Public Type CBT_CREATEWND
        lpcs As CREATESTRUCT
        hWndInsertAfter As Long
End Type
Public Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
' MEASUREITEMSTRUCT for ownerdraw

Public Type MEASUREITEMSTRUCT
        CtlPublic Type As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type
' DRAWITEMSTRUCT for ownerdraw

Public Type DRAWITEMSTRUCT
        CtlPublic Type As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type
' DELETEITEMSTRUCT for ownerdraw

Public Type DELETEITEMSTRUCT
        CtlPublic Type As Long
        CtlID As Long
        itemID As Long
        hwndItem As Long
        itemData As Long
End Type
' COMPAREITEMSTRUCT for ownerdraw sorting

Public Type COMPAREITEMSTRUCT
        CtlPublic Type As Long
        CtlID As Long
        hwndItem As Long
        itemID1 As Long
        itemData1 As Long
        itemID2 As Long
        itemData2 As Long
End Type
Public Type WNDCLASS
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Type DLGTEMPLATE
        style As Long
        dwExtendedStyle As Long
        cdit As Integer
        X As Integer
        Y As Integer
        cx As Integer
        cy As Integer
End Type
Public Type DLGITEMTEMPLATE
        style As Long
        dwExtendedStyle As Long
        X As Integer
        Y As Integer
        cx As Integer
        cy As Integer
        id As Integer
End Type
' Menu item resource format

Public Type MENUITEMTEMPLATEHEADER
        versionNumber As Integer
        offset As Integer
End Type
Public Type MENUITEMTEMPLATE
        mtOption As Integer
        mtID As Integer
        mtString As Byte
End Type
Public Type ICONINFO
        fIcon As Long
        xHotspot As Long
        yHotspot As Long
        hbmMask As Long
        hbmColor As Long
End Type
Public Type MDICREATESTRUCT
        szClass As String
        szTitle As String
        hOwner As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        style As Long
        lParam As Long
End Type
Public Type CLIENTCREATESTRUCT
        hWindowMenu As Long
        idFirstChild As Long
End Type
'  Help engine section.

Public Type MULTIKEYHELP
        mkSize As Long
        mkKeylist As Byte
        szKeyphrase As String * 253 ' Array length is arbitrary; may be changed
End Type
Public Type HELPWININFO
        wStructSize As Long
        X As Long
        Y As Long
        dx As Long
        dy As Long
        wMax As Long
        rgchMember As String * 2
End Type
' *****************************************************************************                                                                             *
' * dde.h -       Dynamic Data Exchange structures and definitions              *
' *                                                                             *
' * Copyright (c) 1993-1995, Microsoft Corp.        All rights reserved              *
' *                                                                             *
' \*****************************************************************************/
' ----------------------------------------------------------------------------
'        DDEACK structure
'
'         Structure of wStatus (LOWORD(lParam)) in WM_DDE_ACK message
'        sent in response to a WM_DDE_DATA, WM_DDE_REQUEST, WM_DDE_POKE,
'        WM_DDE_ADVISE, or WM_DDE_UNADVISE message.
'
' ----------------------------------------------------------------------------*/

Public Type DDEACK
        bAppReturnCode As Integer
        Reserved As Integer
        fbusy As Integer
        fAck As Integer
End Type
' ----------------------------------------------------------------------------
'        DDEADVISE structure
'
'         WM_DDE_ADVISE parameter structure for hOptions (LOWORD(lParam))
'
' ----------------------------------------------------------------------------*/

Public Type DDEADVISE
        Reserved As Integer
        fDeferUpd As Integer
        fAckReq As Integer
        cfFormat As Integer
End Type
' ----------------------------------------------------------------------------
'        DDEDATA structure
'
'        WM_DDE_DATA parameter structure for hData (LOWORD(lParam)).
'        The actual size of this structure depends on the size of
'        the Value array.
'
' ----------------------------------------------------------------------------*/

Public Type DDEDATA
        unused As Integer
        fresponse As Integer
        fRelease As Integer
        Reserved As Integer
        fAckReq As Integer
        cfFormat As Integer
        Value(1) As Byte
End Type
' ----------------------------------------------------------------------------
'         DDEPOKE structure
'
'         WM_DDE_POKE parameter structure for hData (LOWORD(lParam)).
'        The actual size of this structure depends on the size of
'        the Value array.
'
' ----------------------------------------------------------------------------*/

Public Type DDEPOKE
        unused As Integer
        fRelease As Integer
        fReserved As Integer
        cfFormat As Integer
        Value(1) As Byte
End Type
' ----------------------------------------------------------------------------
' The following typedef's were used in previous versions of the Windows SDK.
' They are still valid.  The above typedef's define exactly the same structures
' as those below.  The above typedef names are recommended, however, as they
' are more meaningful.
' Note that the DDEPOKE structure typedef'ed in earlier versions of DDE.H did
' not correctly define the bit positions.
' ----------------------------------------------------------------------------*/

Public Type DDELN
        unused As Integer
        fRelease As Integer
        fDeferUpd As Integer
        fAckReq As Integer
        cfFormat As Integer
End Type
Public Type DDEUP
        unused As Integer
        fAck As Integer
        fRelease As Integer
        fReserved As Integer
        fAckReq As Integer
        cfFormat As Integer
        rgb(1) As Byte
End Type
Public Type HSZPAIR
        hszSvc As Long
        hszTopic As Long
End Type
'//
'// Quality Of Service
'//

Public Type SECURITY_QUALITY_OF_SERVICE
    Length As Long
    ImpersonationLevel As Integer
    ContextTrackingMode As Integer
    EffectiveOnly As Long
End Type
Public Type CONVCONTEXT
        cb As Long
        wFlags As Long
        wCountryID As Long
        iCodePage As Long
        dwLangID As Long
        dwSecurity As Long
        qos As SECURITY_QUALITY_OF_SERVICE
End Type
Public Type CONVINFO
        cb As Long
        hUser As Long
        hConvPartner As Long
        hszSvcPartner As Long
        hszServiceReq As Long
        hszTopic As Long
        hszItem As Long
        wFmt As Long
        wPublic Type As Long
        wStatus As Long
        wConvst As Long
        wLastError As Long
        hConvList As Long
        ConvCtxt As CONVCONTEXT
        hWnd As Long
        hwndPartner As Long
End Type
Public Type DDEML_MSG_HOOK_DATA    '  new for NT
        uiLo As Long  '  unpacked lo and hi parts of lParam
        uiHi As Long
        cbData As Long   '  amount of data in message, if any. May be > than 32 bytes.
        Data(8) As Long  '  data peeking by DDESPY is limited to 32 bytes.
End Type
Public Type MONMSGSTRUCT
        cb As Long
        hwndTo As Long
        dwTime As Long
        htask As Long
        wMsg As Long
        wParam As Long
        lParam As Long
        dmhd As DDEML_MSG_HOOK_DATA       '  new for NT
End Type
Public Type MONCBSTRUCT
        cb As Long
        dwTime As Long
        htask As Long
        dwRet As Long
        wPublic Type As Long
        wFmt As Long
        hConv As Long
        hsz1 As Long
        hsz2 As Long
        hData As Long
        dwData1 As Long
        dwData2 As Long
        cc As CONVCONTEXT                 '  new for NT for XTYP_CONNECT callbacks
        cbData As Long                  '  new for NT for data peeking
        Data(8) As Long                 '  new for NT for data peeking
End Type
Public Type MONHSZSTRUCT
        cb As Long
        fsAction As Long '  MH_ value
        dwTime As Long
        hsz As Long
        htask As Long
        str As Byte
End Type
Public Type MONERRSTRUCT
        cb As Long
        wLastError As Long
        dwTime As Long
        htask As Long
End Type
Public Type MONLINKSTRUCT
        cb As Long
        dwTime As Long
        htask As Long
        fEstablished As Long
        fNoData As Long
        hszSvc As Long
        hszTopic As Long
        hszItem As Long
        wFmt As Long
        fServer As Long
        hConvServer As Long
        hConvClient As Long
End Type
Public Type MONCONVSTRUCT
        cb As Long
        fConnect As Long
        dwTime As Long
        htask As Long
        hszSvc As Long
        hszTopic As Long
        hConvClient As Long        '  Globally unique value != apps local hConv
        hConvServer As Long        '  Globally unique value != apps local hConv
End Type
Public Type smpte
        hour As Byte
        min As Byte
        sec As Byte
        frame As Byte
        fps As Byte
        dummy As Byte
        pad(2) As Byte
End Type
Public Type midi
        songptrpos As Long
End Type
Public Type MMTIME
        wPublic Type As Long
        u As Long
End Type
Public Type MIDIEVENT
        dwDeltaTime As Long          '  Ticks since last event
        dwStreamID As Long           '  Reserved; must be zero
        dwEvent As Long              '  Event Public Type and parameters
        dwParms(1) As Long           '  Parameters if this is a long event
End Type
Public Type MIDISTRMBUFFVER
        dwVersion As Long                  '  Stream buffer format version
        dwMid As Long                      '  Manufacturer ID as defined in MMREG.H
        dwOEMVersion As Long               '  Manufacturer version for custom ext
End Type
Public Type MIDIPROPTIMEDIV
        cbStruct As Long
        dwTimeDiv As Long
End Type
Public Type MIDIPROPTEMPO
        cbStruct As Long
        dwTempo As Long
End Type
Public Type MIXERCAPS
        wMid As Integer                   '  manufacturer id
        wPid As Integer                   '  product id
        vDriverVersion As Long            '  version of the driver
        szPname As String * MAXPNAMELEN   '  product name
        fdwSupport As Long             '  misc. support bits
        cDestinations As Long          '  count of destinations
End Type
Public Type Target    ' for use in MIXERLINE and others (embedded structure)
        
        dwPublic Type As Long                 '  MIXERLINE_TARGETTYPE_xxxx
        dwDeviceID As Long             '  target device ID of device type
        wMid As Integer                   '  of target device
        wPid As Integer                   '       "
        vDriverVersion As Long            '       "
        szPname As String * MAXPNAMELEN
End Type
Public Type MIXERLINE
        cbStruct As Long               '  size of MIXERLINE structure
        dwDestination As Long          '  zero based destination index
        dwSource As Long               '  zero based source index (if source)
        dwLineID As Long               '  unique line id for mixer device
        fdwLine As Long                '  state/information about line
        dwUser As Long                 '  driver specific information
        dwComponentPublic Type As Long        '  component Public Type line connects to
        cChannels As Long              '  number of channels line supports
        cConnections As Long           '  number of connections (possible)
        cControls As Long              '  number of controls at this line
        szShortName As String * MIXER_SHORT_NAME_CHARS
        szName As String * MIXER_LONG_NAME_CHARS
        tTarget As Target
End Type
'   MIXERCONTROL

Public Type MIXERCONTROL
        cbStruct As Long           '  size in Byte of MIXERCONTROL
        dwControlID As Long        '  unique control id for mixer device
        dwControlPublic Type As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
        fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
        cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
        szShortName As String * MIXER_SHORT_NAME_CHARS
        szName As String * MIXER_LONG_NAME_CHARS
        Bounds(1 To 6) As Long     '  Longest member of the Bounds union
        Metrics(1 To 6) As Long    '  Longest member of the Metrics union
End Type
'
'   MIXERLINECONTROLS
'

Public Type MIXERLINECONTROLS
        cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
        dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                                             '  MIXER_GETLINECONTROLSF_ONEBYID or
        dwControl As Long  '  MIXER_GETLINECONTROLSF_ONEBYTYPE
        cControls As Long      '  count of controls pmxctrl points to
        cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
        pamxctrl As MIXERCONTROL       '  pointer to first MIXERCONTROL array
End Type
Public Type MIXERCONTROLDETAILS
        cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
        dwControlID As Long    '  control id to get/set details on
        cChannels As Long      '  number of channels in paDetails array
        item As Long                           ' hwndOwner or cMultipleItems
        cbDetails As Long      '  size of _one_ details_XX struct
        paDetails As Long      '  pointer to array of details_XX structs
End Type
'   MIXER_GETCONTROLDETAILSF_LISTTEXT

Public Type MIXERCONTROLDETAILS_LISTTEXT
        dwParam1 As Long
        dwParam2 As Long
        szName As String * MIXER_LONG_NAME_CHARS
End Type
'   MIXER_GETCONTROLDETAILSF_VALUE

Public Type MIXERCONTROLDETAILS_BOOLEAN
        fValue As Long
End Type
Public Type MIXERCONTROLDETAILS_SIGNED
        lValue As Long
End Type
Public Type MIXERCONTROLDETAILS_UNSIGNED
        dwValue As Long
End Type
Public Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type
Public Type DRVCONFIGINFO
        dwDCISize As Long
        lpszDCISectionName As String
        lpszDCIAliasName As String
        dnDevNode As Long
End Type
Public Type WAVEHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        dwLoops As Long
        lpNext As Long
        Reserved As Long
End Type
Public Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type
Public Type WAVEINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
End Type
Public Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type
Public Type PCMWAVEFORMAT
        wf As WAVEFORMAT
        wBitsPerSample As Integer
End Type
Public Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type
Public Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
End Type
Public Type MIDIHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        lpNext As Long
        Reserved As Long
End Type
Public Type AUXCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        dwSupport As Long
End Type
Public Type TIMECAPS
        wPeriodMin As Long
        wPeriodMax As Long
End Type
Public Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Integer
        wXmax As Integer
        wYmin As Integer
        wYmax As Integer
        wZmin As Integer
        wZmax As Integer
        wNumButtons As Integer
        wPeriodMin As Integer
        wPeriodMax As Integer
End Type
Public Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type
Public Type MMIOINFO
        dwFlags As Long
        fccIOProc As Long
        pIOProc As Long
        wErrorRet As Long
        htask As Long
        cchBuffer As Long
        pchBuffer As String
        pchNext As String
        pchEndRead As String
        pchEndWrite As String
        lBufOffset As Long
        lDiskOffset As Long
        adwInfo(4) As Long
        dwReserved1 As Long
        dwReserved2 As Long
        hmmio As Long
End Type
Public Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccPublic Type As Long
    dwDataOffset As Long
    dwFlags As Long
End Type
Public Type MCI_GENERIC_PARMS
        dwCallback As Long
End Type
Public Type MCI_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDevicePublic Type As String
        lpstrElementName As String
        lpstrAlias As String
End Type
Public Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_SEEK_PARMS
        dwCallback As Long
        dwTo As Long
End Type
Public Type MCI_STATUS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwItem As Long
        dwTrack As Integer
End Type
Public Type MCI_INFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
End Type
Public Type MCI_GETDEVCAPS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwIten As Long
End Type
Public Type MCI_SYSINFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
        dwNumber As Long
        wDevicePublic Type As Long
End Type
Public Type MCI_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
End Type
Public Type MCI_BREAK_PARMS
        dwCallback As Long
        nVirtKey As Long
        hwndBreak As Long
End Type
Public Type MCI_SOUND_PARMS
        dwCallback As Long
        lpstrSoundName As String
End Type
Public Type MCI_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
End Type
Public Type MCI_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
End Type
Public Type MCI_RECORD_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_VD_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type
Public Type MCI_VD_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type
Public Type MCI_VD_ESCAPE_PARMS
        dwCallback As Long
        lpstrCommand As String
End Type
Public Type MCI_WAVE_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDevicePublic Type As String
        lpstrElementName As String
        lpstrAlias As String
        dwBufferSeconds As Long
End Type
Public Type MCI_WAVE_DELETE_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_WAVE_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        wInput As Long
        wOutput As Long
        wFormatTag As Integer
        wReserved2 As Integer
        nChannels As Integer
        wReserved3 As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wReserved4 As Integer
        wBitsPerSample As Integer
        wReserved5 As Integer
End Type
Public Type MCI_SEQ_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        dwTempo As Long
        dwPort As Long
        dwSlave As Long
        dwMaster As Long
        dwOffset As Long
End Type
Public Type MCI_ANIM_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDevicePublic Type As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hwndParent As Long
End Type
Public Type MCI_ANIM_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type
Public Type MCI_ANIM_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type
Public Type MCI_ANIM_WINDOW_PARMS
        dwCallback As Long
        hWnd As Long
        nCmdShow As Long
        lpstrText As String
End Type
Public Type MCI_ANIM_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type
Public Type MCI_ANIM_UPDATE_PARMS
        dwCallback As Long
        rc As RECT
        hdc As Long
End Type
Public Type MCI_OVLY_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDevicePublic Type As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hwndParent As Long
End Type
Public Type MCI_OVLY_WINDOW_PARMS
        dwCallback As Long
        hWnd As Long
        nCmdShow As Long
        lpstrText As String
End Type
Public Type MCI_OVLY_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type
Public Type MCI_OVLY_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type
Public Type MCI_OVLY_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type
' -------------
' Print APIs
' -------------

Public Type PRINTER_INFO_1
        Flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type
Public Type PRINTER_INFO_2
        pServerName As String
        pPrinterName As String
        pShareName As String
        pPortName As String
        pDriverName As String
        pComment As String
        pLocation As String
        pDevmode As DEVMODE
        pSepFile As String
        pPrintProcessor As String
        pDataPublic Type As String
        pParameters As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Attributes As Long
        Priority As Long
        DefaultPriority As Long
        StartTime As Long
        UntilTime As Long
        Status As Long
        cJobs As Long
        AveragePPM As Long
End Type
Public Type PRINTER_INFO_3
        pSecurityDescriptor As SECURITY_DESCRIPTOR
End Type
Public Type JOB_INFO_1
        JobId As Long
        pPrinterName As String
        pMachineName As String
        pUserName As String
        pDocument As String
        pDataPublic Type As String
        pStatus As String
        Status As Long
        Priority As Long
        Position As Long
        TotalPages As Long
        PagesPrinted As Long
        Submitted As SystemTime
End Type
Public Type JOB_INFO_2
        JobId As Long
        pPrinterName As String
        pMachineName As String
        pUserName As String
        pDocument As String
        pNotifyName As String
        pDataPublic Type As String
        pPrintProcessor As String
        pParameters As String
        pDriverName As String
        pDevmode As DEVMODE
        pStatus As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Status As Long
        Priority As Long
        Position As Long
        StartTime As Long
        UntilTime As Long
        TotalPages As Long
        Size As Long
        Submitted As SystemTime
        time As Long
        PagesPrinted As Long
End Type
Public Type ADDJOB_INFO_1
        Path As String
        JobId As Long
End Type
Public Type DRIVER_INFO_1
        pName As String
End Type
Public Type DRIVER_INFO_2
        cVersion As Long
        pName As String
        pEnvironment As String
        pDriverPath As String
        pDataFile As String
        pConfigFile As String
End Type
Public Type DOC_INFO_1
        pDocName As String
        pOutputFile As String
        pDataPublic Type As String
End Type
Public Type FORM_INFO_1
        pName As String
        Size As SIZEL
        ImageableArea As RECTL
End Type
Public Type PRINTPROCESSOR_INFO_1
        pName As String
End Type
Public Type PORT_INFO_1
        pName As String
End Type
Public Type MONITOR_INFO_1
        pName As String
End Type
Public Type MONITOR_INFO_2
        pName As String
        pEnvironment As String
        pDLLName As String
End Type
Public Type DATATYPES_INFO_1
        pName As String
End Type
Public Type PRINTER_DEFAULTS
        pDataPublic Type As String
        pDevmode As DEVMODE
        DesiredAccess As Long
End Type
Public Type PRINTER_INFO_4
        pPrinterName As String
        pServerName As String
        Attributes As Long
End Type
Public Type PRINTER_INFO_5
        pPrinterName As String
        pPortName As String
        Attributes As Long
        DeviceNotSelectedTimeout As Long
        TransmissionRetryTimeout As Long
End Type
Public Type DRIVER_INFO_3
        cVersion As Long
        pName As String                    '  QMS 810
        pEnvironment As String             '  Win32 x86
        pDriverPath As String              '  c:\drivers\pscript.dll
        pDataFile As String                '  c:\drivers\QMS810.PPD
        pConfigFile As String              '  c:\drivers\PSCRPTUI.DLL
        pHelpFile As String                '  c:\drivers\PSCRPTUI.HLP
        pDependentFiles As String          '
        pMonitorName As String             '  "PJL monitor"
        pDefaultDataPublic Type As String         '  "EMF"
End Type
Public Type DOC_INFO_2
        pDocName As String
        pOutputFile As String
        pDataPublic Type As String
        dwMode As Long
        JobId As Long
End Type
Public Type PORT_INFO_2
        pPortName As String
        pMonitorName As String
        pDescription As String
        fPortPublic Type As Long
        Reserved As Long
End Type
Public Type PROVIDOR_INFO_1
        pName As String
        pEnvironment As String
        pDLLName As String
End Type
Public Type NCB
        ncb_command As Integer
        ncb_retcode As Integer
        ncb_lsn As Integer
        ncb_num As Integer
        ncb_buffer As String
        ncb_length As Integer
        ncb_callname As String * NCBNAMSZ
        ncb_name As String * NCBNAMSZ
        ncb_rto As Integer
        ncb_sto As Integer
        ncb_post As Long
        ncb_lana_num As Integer
        ncb_cmd_cplt As Integer
        ncb_reserve(10) As Byte ' Reserved, must be 0
        ncb_event As Long
End Type
Public Type ADAPTER_STATUS
        adapter_address As String * 6
        rev_major As Integer
        reserved0 As Integer
        adapter_Public Type As Integer
        rev_minor As Integer
        duration As Integer
        frmr_recv As Integer
        frmr_xmit As Integer
        iframe_recv_err As Integer
        xmit_aborts As Integer
        xmit_success As Long
        recv_success As Long
        iframe_xmit_err As Integer
        recv_buff_unavail As Integer
        t1_timeouts As Integer
        ti_timeouts As Integer
        Reserved1 As Long
        free_ncbs As Integer
        max_cfg_ncbs As Integer
        max_ncbs As Integer
        xmit_buf_unavail As Integer
        max_dgram_size As Integer
        pending_sess As Integer
        max_cfg_sess As Integer
        max_sess As Integer
        max_sess_pkt_size As Integer
        name_count As Integer
End Type
Public Type NAME_BUFFER
        Name  As String * NCBNAMSZ
        name_num As Integer
        name_flags As Integer
End Type
Public Type SESSION_HEADER
        sess_name As Integer
        num_sess As Integer
        rcv_dg_outstanding As Integer
        rcv_any_outstanding As Integer
End Type
Public Type SESSION_BUFFER
        lsn As Integer
        State As Integer
        local_name As String * NCBNAMSZ
        remote_name As String * NCBNAMSZ
        rcvs_outstanding As Integer
        sends_outstanding As Integer
End Type
Public Type LANA_ENUM
        Length As Integer
        lana(MAX_LANA) As Integer
End Type
Public Type FIND_NAME_HEADER
        node_count As Integer
        Reserved As Integer
        unique_group As Integer
End Type
Public Type FIND_NAME_BUFFER
        Length As Integer
        access_control As Integer
        frame_control As Integer
        destination_addr(6) As Integer
        source_addr(6) As Integer
        routing_info(18) As Integer
End Type
Public Type ACTION_HEADER
        transport_id As Long
        action_code As Integer
        Reserved As Integer
End Type
Public Type USER_INFO_3
   ' Level 0 starts here
   Name As Long
   ' Level 1 starts here
   Password As Long
   PasswordAge As Long
   Privilege As Long
   HomeDir As Long
   Comment As Long
   Flags As Long
   ScriptPath As Long
   ' Level 2 starts here
   AuthFlags As Long
   FullName As Long
   UserComment As Long
   Parms As Long
   Workstations As Long
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours As Long
   BadPwCount As Long
   NumLogons As Long
   LogonServer As Long
   CountryCode As Long
   CodePage As Long
   ' Level 3 starts here
   UserID As Long
   PrimaryGroupID As Long
   Profile As Long
   HomeDirDrive As Long
   PasswordExpired As Long
End Type
Public Type GROUP_INFO_2
   Name As Long
   Comment As Long
   GroupID As Long
   Attributes As Long
End Type
Public Type LOCALGROUP_MEMBERS_INFO_0
   pSid As Long
End Type
Public Type LOCALGROUP_MEMBERS_INFO_1
   'Level 0 Starts Here
   pSid As Long
   'Level 1 Starts Here
   eUsage As g_netSID_NAME_USE
   psName As Long
End Type
Public Type WKSTA_INFO_102
   wki102_platform_id As Long
   wki102_computername As Long
   wki102_langroup As Long
   wki102_ver_major As Long
   wki102_ver_minor As Long
   wki102_lanroot As Long
   wki102_logged_on_users As Long
End Type
Public Type WKSTA_USER_INFO_1
   wkui1_username As Long
   wkui1_logon_domain As Long
   wkui1_oth_domains As Long
   wkui1_logon_server As Long
End Type
Public Type NETRESOURCE
    dwScope As Long
    dwPublic Type As Long
    dwDisplayPublic Type As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type
Public Type CRGB
        bRed As Byte
        bGreen As Byte
        bBlue As Byte
        bExtra As Byte
End Type
Public Type SERVICE_STATUS
        dwServicePublic Type As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type
Public Type ENUM_SERVICE_STATUS
        lpServiceName As String
        lpDisplayName As String
        ServiceStatus As SERVICE_STATUS
End Type
Public Type QUERY_SERVICE_LOCK_STATUS
        fIsLocked As Long
        lpLockOwner As String
        dwLockDuration As Long
End Type
Public Type QUERY_SERVICE_CONFIG
        dwServicePublic Type As Long
        dwStartPublic Type As Long
        dwErrorControl As Long
        lpBinaryPathName As String
        lpLoadOrderGroup As String
        dwTagId As Long
        lpDependencies As String
        lpServiceStartName As String
        lpDisplayName As String
End Type
Public Type SERVICE_TABLE_ENTRY
        lpServiceName As String
        lpServiceProc As Long
End Type
Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Public Type PERF_DATA_BLOCK
        Signature As String * 4
        LittleEndian As Long
        Version As Long
        Revision As Long
        TotalByteLength As Long
        HeaderLength As Long
        NumObjectTypes As Long
        DefaultObject As Long
        SystemTime As SystemTime
        PerfTime As LARGE_INTEGER
        PerfFreq As LARGE_INTEGER
        PerTime100nSec As LARGE_INTEGER
        SystemNameLength As Long
        SystemNameOffset As Long
End Type
Public Type PERF_OBJECT_TYPE
        TotalByteLength As Long
        DefinitionLength As Long
        HeaderLength As Long
        ObjectNameTitleIndex As Long
        ObjectNameTitle As String
        ObjectHelpTitleIndex As Long
        ObjectHelpTitle As String
        DetailLevel As Long
        NumCounters As Long
        DefaultCounter As Long
        NumInstances As Long
        CodePage As Long
        PerfTime As LARGE_INTEGER
        PerfFreq As LARGE_INTEGER
End Type
Public Type PERF_COUNTER_DEFINITION
        ByteLength As Long
        CounterNameTitleIndex As Long
        CounterNameTitle As String
        CounterHelpTitleIndex As Long
        CounterHelpTitle As String
        DefaultScale As Long
        DetailLevel As Long
        CounterPublic Type As Long
        CounterSize As Long
        CounterOffset As Long
End Type
Public Type PERF_INSTANCE_DEFINITION
        ByteLength As Long
        ParentObjectTitleIndex As Long
        ParentObjectInstance As Long
        UniqueID As Long
        NameOffset As Long
        NameLength As Long
End Type
Public Type PERF_COUNTER_BLOCK
        ByteLength As Long
End Type
Public Type COMPOSITIONFORM
        dwStyle As Long
        ptCurrentPos As POINTAPI
        rcArea As RECT
End Type
Public Type CANDIDATEFORM
        dwIndex As Long
        dwStyle As Long
        ptCurrentPos As POINTAPI
        rcArea As RECT
End Type
Public Type CANDIDATELIST
        dwSize As Long
        dwStyle As Long
        dwCount As Long
        dwSelection As Long
        dwPageStart As Long
        dwPageSize As Long
        dwOffset(1) As Long
End Type
Public Type STYLEBUF
        dwStyle As Long
        szDescription As String * STYLE_DESCRIPTION_SIZE
End Type
' ***********************************************************************
' *                                                                       *
' *   mcx.h -- This module defines the 32-Bit Windows MCX APIs            *
' *                                                                       *
' *   Copyright (c) 1990-1995, Microsoft Corp. All rights reserved.       *
' *                                                                       *
' ************************************************************************/

Public Type MODEMDEVCAPS
        dwActualSize As Long
        dwRequiredSize As Long
        dwDevSpecificOffset As Long
        dwDevSpecificSize As Long

    '  product and version identification
        dwModemProviderVersion As Long
        dwModemManufacturerOffset As Long
        dwModemManufacturerSize As Long
        dwModemModelOffset As Long
        dwModemModelSize As Long
        dwModemVersionOffset As Long
        dwModemVersionSize As Long

    '  local option capabilities
        dwDialOptions As Long          '  bitmap of supported values
        dwCallSetupFailTimer As Long   '  maximum in seconds
        dwInactivityTimeout As Long    '  maximum in seconds
        dwSpeakerVolume As Long        '  bitmap of supported values
        dwSpeakerMode As Long          '  bitmap of supported values
        dwModemOptions As Long         '  bitmap of supported values
        dwMaxDTERate As Long           '  maximum value in bit/s
        dwMaxDCERate As Long           '  maximum value in bit/s

    '  Variable portion for proprietary expansion
        abVariablePortion(1) As Byte
End Type
Public Type MODEMSETTINGS
        dwActualSize As Long
        dwRequiredSize As Long
        dwDevSpecificOffset As Long
        dwDevSpecificSize As Long

    '  static local options (read/write)
        dwCallSetupFailTimer As Long       '  seconds
        dwInactivityTimeout As Long        '  seconds
        dwSpeakerVolume As Long            '  level
        dwSpeakerMode As Long              '  mode
        dwPreferredModemOptions As Long    '  bitmap
    
    '  negotiated options (read only) for current or last call
        dwNegotiatedModemOptions As Long   '  bitmap
        dwNegotiatedDCERate As Long        '  bit/s

    '  Variable portion for proprietary expansion
        abVariablePortion(1) As Byte
End Type
Public Type DRAGINFO
        uSize As Long                 '  init with sizeof(DRAGINFO)
        pt As POINTAPI
        fNC As Long
        lpFileList As String
        grfKeyState As Long
End Type
Public Type APPBARDATA
        cbSize As Long
        hWnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type
'  no POF_ flags currently defined
'  implicit parameters are:
'       if pFrom or pTo are unqualified names the current directories are
'       taken from the global current drive/directory settings managed
'       by Get/SetCurrentDrive/Directory
'
'       the global confirmation settings

Public Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
Public Type SHNAMEMAPPING
        pszOldPath As String
        pszNewPath As String
        cchOldPath As Long
        cchNewPath As Long
End Type
Public Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hWnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type
' //  End ShellExecuteEx and family
' // Tray notification definitions

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
' // End Tray Notification Icons
' // Begin SHGetFileInfo
'  * The SHGetFileInfo API provides an easy way to get attributes
'  * for a file given a pathname.
'  *
'  *   PARAMETERS
'  *
'  *     pszPath              file name to get info about
'  *     dwFileAttributes     file attribs, only used with SHGFI_USEFILEATTRIBUTES
'  *     psfi                 place to return file info
'  *     cbFileInfo           size of structure
'  *     uFlags               flags
'  *
'  *   RETURN
'  *     TRUE if things worked
'  */

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: Public Type name
End Type
'  ----- Types and structures -----

Public Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFilePublic Type As Long             '  e.g. VFT_DRIVER
        dwFileSubPublic Type As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type
' ***********************************************************************
' *                                                                       *
' *   winbase.h -- This module defines the 32-Bit Windows Base APIs       *
' *                                                                       *
' *   Copyright (c) 1990-1995, Microsoft Corp. All rights reserved.       *
' *                                                                       *
' ************************************************************************/

Public Type ICONMETRICS
        cbSize As Long
        iHorzSpacing As Long
        iVertSpacing As Long
        iTitleWrap As Long
        lfFont As LOGFONT
End Type
Public Type HELPINFO
        cbSize As Long
        iContextPublic Type As Long
        iCtrlId As Long
        hItemHandle As Long
        dwContextId As Long
        MousePos As POINTAPI
End Type
Public Type ANIMATIONINFO
        cbSize As Long
        iMinAnimate As Long
End Type
Public Type MINIMIZEDMETRICS
        cbSize As Long
        iWidth As Long
        iHorzGap As Long
        iVertGap As Long
        iArrange As Long
        lfFont As LOGFONT
End Type
'  Performance counter API's

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Public Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type
' *   commdlg.h -- This module defines the 32-Bit Common Dialog APIs      *

Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Public Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type
Public Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Type FINDREPLACE
        lStructSize As Long        '  size of this struct 0x20
        hwndOwner As Long          '  handle to owner's window
        hInstance As Long          '  instance handle of.EXE that
                                    '    contains cust. dlg. template
        Flags As Long              '  one or more of the FR_??
        lpstrFindWhat As String      '  ptr. to search string
        lpstrReplaceWith As String   '  ptr. to replace string
        wFindWhatLen As Integer       '  size of find buffer
        wReplaceWithLen As Integer    '  size of replace buffer
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long            '  ptr. to hook fn. or NULL
        lpTemplateName As String     '  custom template name
End Type
Public Type ChooseFont
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As Long
        iPointSize As Long         '  10 * size in points of selected font
        Flags As Long              '  enum. Public Type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontPublic Type As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Public Type PrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type
Public Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type
Public Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        Flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type
Public Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubPublic Type As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type
Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelPublic Type As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerPublic Type As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type
Public Type DRAWTEXTPARAMS
        cbSize As Long
        iTabLength As Long
        iLeftMargin As Long
        iRightMargin As Long
        uiLengthDrawn As Long
End Type
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fPublic Type As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Type SCROLLINFO
        cbSize As Long
        fMask As Long
        nMin As Long
        nMax As Long
        nPage As Long
        nPos As Long
        nTrackPos As Long
End Type
Public Type MSGBOXPARAMS
        cbSize As Long
        hwndOwner As Long
        hInstance As Long
        lpszText As String
        lpszCaption As String
        dwStyle As Long
        lpszIcon As String
        dwContextHelpId As Long
        lpfnMsgBoxCallback As Long
        dwLanguageId As Long
End Type
Public Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type
Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type
