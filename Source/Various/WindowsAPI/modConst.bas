Attribute VB_Name = "modConst"
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
' Type definitions for Windows' basic types.

Public Const ANYSIZE_ARRAY = 1
Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Public Const SID_REVISION = (1)                         '  Current revision level
Public Const SID_MAX_SUB_AUTHORITIES = (15)
Public Const SID_RECOMMENDED_SUB_AUTHORITIES = (1)    ' Will change to around 6 in a future release.
Public Const SidTypeUser = 1
Public Const SidTypeGroup = 2
Public Const SidTypeDomain = 3
Public Const SidTypeAlias = 4
Public Const SidTypeWellKnownGroup = 5
Public Const SidTypeDeletedAccount = 6
Public Const SidTypeInvalid = 7
Public Const SidTypeUnknown = 8
' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  Universal well-known SIDs                                               //
'                                                                          //
'      Null SID              S-1-0-0                                       //
'      World                 S-1-1-0                                       //
'      Local                 S-1-2-0                                       //
'      Creator Owner ID      S-1-3-0                                       //
'      Creator Group ID      S-1-3-1                                       //
'                                                                          //
'      (Non-unique IDs)      S-1-4                                         //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////

Public Const SECURITY_NULL_RID = &H0
Public Const SECURITY_WORLD_RID = &H0
Public Const SECURITY_LOCAL_RID = &H0
Public Const SECURITY_CREATOR_OWNER_RID = &H0
Public Const SECURITY_CREATOR_GROUP_RID = &H1
' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  NT well-known SIDs                                                      //
'                                                                          //
'      NT Authority          S-1-5                                         //
'      Dialup                S-1-5-1                                       //
'                                                                          //
'      Network               S-1-5-2                                       //
'      Batch                 S-1-5-3                                       //
'      Interactive           S-1-5-4                                       //
'      Service               S-1-5-6                                       //
'      AnonymousLogon        S-1-5-7       (aka null logon session)        //
'                                                                          //
'      (Logon IDs)           S-1-5-5-X-Y                                   //
'                                                                          //
'      (NT non-unique IDs)   S-1-5-0x15-...                                //
'                                                                          //
'      (Built-in domain)     s-1-5-0x20                                    //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////

Public Const SECURITY_DIALUP_RID = &H1
Public Const SECURITY_NETWORK_RID = &H2
Public Const SECURITY_BATCH_RID = &H3
Public Const SECURITY_INTERACTIVE_RID = &H4
Public Const SECURITY_SERVICE_RID = &H6
Public Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Public Const SECURITY_LOGON_IDS_RID = &H5
Public Const SECURITY_LOCAL_SYSTEM_RID = &H12
Public Const SECURITY_NT_NON_UNIQUE = &H15
Public Const SECURITY_BUILTIN_DOMAIN_RID = &H20
' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  well-known domain relative sub-authority values (RIDs)...               //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////

Public Const DOMAIN_USER_RID_ADMIN = &H1F4
Public Const DOMAIN_USER_RID_GUEST = &H1F5
Public Const DOMAIN_GROUP_RID_ADMINS = &H200
Public Const DOMAIN_GROUP_RID_USERS = &H201
Public Const DOMAIN_GROUP_RID_GUESTS = &H202
Public Const DOMAIN_ALIAS_RID_ADMINS = &H220
Public Const DOMAIN_ALIAS_RID_USERS = &H221
Public Const DOMAIN_ALIAS_RID_GUESTS = &H222
Public Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Public Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Public Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Public Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Public Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Public Const DOMAIN_ALIAS_RID_REPLICATOR = &H228
'  Allocate the System Luid.  The first 1000 LUIDs are reserved.
'  Use #999 here0x3E7 = 999)
'  end_ntifs
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                           User and Group related SID attributes     //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'  Group attributes

Public Const SE_GROUP_MANDATORY = &H1
Public Const SE_GROUP_ENABLED_BY_DEFAULT = &H2
Public Const SE_GROUP_ENABLED = &H4
Public Const SE_GROUP_OWNER = &H8
Public Const SE_GROUP_LOGON_ID = &HC0000000
'  User attributes
'  (None yet defined.)
' ----------------
'  Kernel Section
' ----------------

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5
' Define the dwOpenMode values for CreateNamedPipe

Public Const PIPE_ACCESS_INBOUND = &H1
Public Const PIPE_ACCESS_OUTBOUND = &H2
Public Const PIPE_ACCESS_DUPLEX = &H3
' Define the Named Pipe End flags for GetNamedPipeInfo

Public Const PIPE_CLIENT_END = &H0
Public Const PIPE_SERVER_END = &H1
' Define the dwPipeMode values for CreateNamedPipe

Public Const PIPE_WAIT = &H0
Public Const PIPE_NOWAIT = &H1
Public Const PIPE_READMODE_BYTE = &H0
Public Const PIPE_READMODE_MESSAGE = &H2
Public Const PIPE_TYPE_BYTE = &H0
Public Const PIPE_TYPE_MESSAGE = &H4
' Define the well known values for CreateNamedPipe nMaxInstances

Public Const PIPE_UNLIMITED_INSTANCES = 255
' Define the Security Quality of Service bits to be passed
'  into CreateFile

Public Const SECURITY_CONTEXT_TRACKING = &H40000
Public Const SECURITY_EFFECTIVE_ONLY = &H80000
Public Const SECURITY_SQOS_PRESENT = &H100000
Public Const SECURITY_VALID_SQOS_FLAGS = &H1F0000
'  Serial provider type.

Public Const SP_SERIALCOMM = &H1&
'  Provider SubTypes

Public Const PST_UNSPECIFIED = &H0&
Public Const PST_RS232 = &H1&
Public Const PST_PARALLELPORT = &H2&
Public Const PST_RS422 = &H3&
Public Const PST_RS423 = &H4&
Public Const PST_RS449 = &H5&
Public Const PST_FAX = &H21&
Public Const PST_SCANNER = &H22&
Public Const PST_NETWORK_BRIDGE = &H100&
Public Const PST_LAT = &H101&
Public Const PST_TCPIP_TELNET = &H102&
Public Const PST_X25 = &H103&
'  Provider capabilities flags.

Public Const PCF_DTRDSR = &H1&
Public Const PCF_RTSCTS = &H2&
Public Const PCF_RLSD = &H4&
Public Const PCF_PARITY_CHECK = &H8&
Public Const PCF_XONXOFF = &H10&
Public Const PCF_SETXCHAR = &H20&
Public Const PCF_TOTALTIMEOUTS = &H40&
Public Const PCF_INTTIMEOUTS = &H80&
Public Const PCF_SPECIALCHARS = &H100&
Public Const PCF_16BITMODE = &H200&
'  Comm provider settable parameters.

Public Const SP_PARITY = &H1&
Public Const SP_BAUD = &H2&
Public Const SP_DATABITS = &H4&
Public Const SP_STOPBITS = &H8&
Public Const SP_HANDSHAKING = &H10&
Public Const SP_PARITY_CHECK = &H20&
Public Const SP_RLSD = &H40&
'  Settable baud rates in the provider.

Public Const BAUD_075 = &H1&
Public Const BAUD_110 = &H2&
Public Const BAUD_134_5 = &H4&
Public Const BAUD_150 = &H8&
Public Const BAUD_300 = &H10&
Public Const BAUD_600 = &H20&
Public Const BAUD_1200 = &H40&
Public Const BAUD_1800 = &H80&
Public Const BAUD_2400 = &H100&
Public Const BAUD_4800 = &H200&
Public Const BAUD_7200 = &H400&
Public Const BAUD_9600 = &H800&
Public Const BAUD_14400 = &H1000&
Public Const BAUD_19200 = &H2000&
Public Const BAUD_38400 = &H4000&
Public Const BAUD_56K = &H8000&
Public Const BAUD_128K = &H10000
Public Const BAUD_115200 = &H20000
Public Const BAUD_57600 = &H40000
Public Const BAUD_USER = &H10000000
'  Settable Data Bits

Public Const DATABITS_5 = &H1&
Public Const DATABITS_6 = &H2&
Public Const DATABITS_7 = &H4&
Public Const DATABITS_8 = &H8&
Public Const DATABITS_16 = &H10&
Public Const DATABITS_16X = &H20&
'  Settable Stop and Parity bits.

Public Const STOPBITS_10 = &H1&
Public Const STOPBITS_15 = &H2&
Public Const STOPBITS_20 = &H4&
Public Const PARITY_NONE = &H100&
Public Const PARITY_ODD = &H200&
Public Const PARITY_EVEN = &H400&
Public Const PARITY_MARK = &H800&
Public Const PARITY_SPACE = &H1000&
' The eight actual COMSTAT bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName     Bit #     Description
' ---------     -----     ---------------------------
' fCtsHold        1       Tx waiting for CTS signal
' fDsrHold        2       Tx waiting for DSR signal
' fRlsdHold       3       Tx waiting for RLSD signal
' fXoffHold       4       Tx waiting, XOFF char rec'd
' fXoffSent       5       Tx waiting, XOFF char sent
' fEof            6       EOF character sent
' fTxim           7       character waiting for Tx
' fReserved       8       reserved (25 bits)
'  DTR Control Flow Values.

Public Const DTR_CONTROL_DISABLE = &H0
Public Const DTR_CONTROL_ENABLE = &H1
Public Const DTR_CONTROL_HANDSHAKE = &H2
'  RTS Control Flow Values

Public Const RTS_CONTROL_DISABLE = &H0
Public Const RTS_CONTROL_ENABLE = &H1
Public Const RTS_CONTROL_HANDSHAKE = &H2
Public Const RTS_CONTROL_TOGGLE = &H3
' Global Memory Flags

Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
' Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)

Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_LOCKCOUNT = &HFF
' Local Memory Flags

Public Const LMEM_FIXED = &H0
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10
Public Const LMEM_NODISCARD = &H20
Public Const LMEM_ZEROINIT = &H40
Public Const LMEM_MODIFY = &H80
Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_VALID_FLAGS = &HF72
Public Const LMEM_INVALID_HANDLE = &H8000
Public Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Public Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)
Public Const NONZEROLHND = (LMEM_MOVEABLE)
Public Const NONZEROLPTR = (LMEM_FIXED)
' Flags returned by LocalFlags (in addition to LMEM_DISCARDABLE)

Public Const LMEM_DISCARDED = &H4000
Public Const LMEM_LOCKCOUNT = &HFF
'  dwCreationFlag values

Public Const DEBUG_PROCESS = &H1
Public Const DEBUG_ONLY_THIS_PROCESS = &H2
Public Const CREATE_SUSPENDED = &H4
Public Const DETACHED_PROCESS = &H8
Public Const CREATE_NEW_CONSOLE = &H10
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const CREATE_NEW_PROCESS_GROUP = &H200
Public Const CREATE_NO_WINDOW = &H8000000
Public Const PROFILE_USER = &H10000000
Public Const PROFILE_KERNEL = &H20000000
Public Const PROFILE_SERVER = &H40000000
Public Const MAXLONG = &H7FFFFFFF
Public Const THREAD_BASE_PRIORITY_MIN = -2
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_BASE_PRIORITY_LOWRT = 15
Public Const THREAD_BASE_PRIORITY_IDLE = -15
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_ERROR_RETURN = (MAXLONG)
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
' ++ BUILD Version: 0093     Increment this if a change has global effects
' Copyright (c) 1990-1995  Microsoft Corporation
' Module Name:
'     winnt.h
' Abstract:
'     This module defines the 32-Bit Windows types and constants that are
'     defined by NT, but exposed through the Win32 API.
' Revision History:

Public Const APPLICATION_ERROR_MASK = &H20000000
Public Const ERROR_SEVERITY_SUCCESS = &H0
Public Const ERROR_SEVERITY_INFORMATIONAL = &H40000000
Public Const ERROR_SEVERITY_WARNING = &H80000000
Public Const ERROR_SEVERITY_ERROR = &HC0000000
Public Const MINCHAR = &H80
Public Const MAXCHAR = &H7F
Public Const MINSHORT = &H8000
Public Const MAXSHORT = &H7FFF
Public Const MINLONG = &H80000000
Public Const MAXByte = &HFF
Public Const MAXWORD = &HFFFF
Public Const MAXDWORD = &HFFFF
'
'  Calculate the byte offset of a field in a structure of type type.
'  *  Language IDs.
'  *
'  *  The following two combinations of primary language ID and
'  *  sublanguage ID have special semantics:
'  *
'  *    Primary Language ID   Sublanguage ID      Result
'  *    -------------------   ---------------     ------------------------
'  *    LANG_NEUTRAL          SUBLANG_NEUTRAL     Language neutral
'  *    LANG_NEUTRAL          SUBLANG_DEFAULT     User default language
'  *    LANG_NEUTRAL          SUBLANG_SYS_DEFAULT System default language
'  */
'
'  *  Primary language IDs.
'  */

Public Const LANG_NEUTRAL = &H0
Public Const LANG_BULGARIAN = &H2
Public Const LANG_CHINESE = &H4
Public Const LANG_CROATIAN = &H1A
Public Const LANG_CZECH = &H5
Public Const LANG_DANISH = &H6
Public Const LANG_DUTCH = &H13
Public Const LANG_ENGLISH = &H9
Public Const LANG_FINNISH = &HB
Public Const LANG_FRENCH = &HC
Public Const LANG_GERMAN = &H7
Public Const LANG_GREEK = &H8
Public Const LANG_HUNGARIAN = &HE
Public Const LANG_ICELANDIC = &HF
Public Const LANG_ITALIAN = &H10
Public Const LANG_JAPANESE = &H11
Public Const LANG_KOREAN = &H12
Public Const LANG_NORWEGIAN = &H14
Public Const LANG_POLISH = &H15
Public Const LANG_PORTUGUESE = &H16
Public Const LANG_ROMANIAN = &H18
Public Const LANG_RUSSIAN = &H19
Public Const LANG_SLOVAK = &H1B
Public Const LANG_SLOVENIAN = &H24
Public Const LANG_SPANISH = &HA
Public Const LANG_SWEDISH = &H1D
Public Const LANG_TURKISH = &H1F
'
'  *  Sublanguage IDs.
'  *
'  *  The name immediately following SUBLANG_ dictates which primary
'  *  language ID that sublanguage ID can be combined with to form a
'  *  valid language ID.
'  */

Public Const SUBLANG_NEUTRAL = &H0                       '  language neutral
Public Const SUBLANG_DEFAULT = &H1                       '  user default
Public Const SUBLANG_SYS_DEFAULT = &H2                   '  system default
Public Const SUBLANG_CHINESE_TRADITIONAL = &H1           '  Chinese (Taiwan)
Public Const SUBLANG_CHINESE_SIMPLIFIED = &H2            '  Chinese (PR China)
Public Const SUBLANG_CHINESE_HONGKONG = &H3              '  Chinese (Hong Kong)
Public Const SUBLANG_CHINESE_SINGAPORE = &H4             '  Chinese (Singapore)
Public Const SUBLANG_DUTCH = &H1                         '  Dutch
Public Const SUBLANG_DUTCH_BELGIAN = &H2                 '  Dutch (Belgian)
Public Const SUBLANG_ENGLISH_US = &H1                    '  English (USA)
Public Const SUBLANG_ENGLISH_UK = &H2                    '  English (UK)
Public Const SUBLANG_ENGLISH_AUS = &H3                   '  English (Australian)
Public Const SUBLANG_ENGLISH_CAN = &H4                   '  English (Canadian)
Public Const SUBLANG_ENGLISH_NZ = &H5                    '  English (New Zealand)
Public Const SUBLANG_ENGLISH_EIRE = &H6                  '  English (Irish)
Public Const SUBLANG_FRENCH = &H1                        '  French
Public Const SUBLANG_FRENCH_BELGIAN = &H2                '  French (Belgian)
Public Const SUBLANG_FRENCH_CANADIAN = &H3               '  French (Canadian)
Public Const SUBLANG_FRENCH_SWISS = &H4                  '  French (Swiss)
Public Const SUBLANG_GERMAN = &H1                        '  German
Public Const SUBLANG_GERMAN_SWISS = &H2                  '  German (Swiss)
Public Const SUBLANG_GERMAN_AUSTRIAN = &H3               '  German (Austrian)
Public Const SUBLANG_ITALIAN = &H1                       '  Italian
Public Const SUBLANG_ITALIAN_SWISS = &H2                 '  Italian (Swiss)
Public Const SUBLANG_NORWEGIAN_BOKMAL = &H1              '  Norwegian (Bokma
Public Const SUBLANG_NORWEGIAN_NYNORSK = &H2             '  Norwegian (Nynorsk)
Public Const SUBLANG_PORTUGUESE = &H2                    '  Portuguese
Public Const SUBLANG_PORTUGUESE_BRAZILIAN = &H1          '  Portuguese (Brazilian)
Public Const SUBLANG_SPANISH = &H1                       '  Spanish (Castilian)
Public Const SUBLANG_SPANISH_MEXICAN = &H2               '  Spanish (Mexican)
Public Const SUBLANG_SPANISH_MODERN = &H3                '  Spanish (Modern)
'
'  *  Sorting IDs.
'  *
'  */

Public Const SORT_DEFAULT = &H0                          '  sorting default
Public Const SORT_JAPANESE_XJIS = &H0                    '  Japanese0xJIS order
Public Const SORT_JAPANESE_UNICODE = &H1                 '  Japanese Unicode order
Public Const SORT_CHINESE_BIG5 = &H0                     '  Chinese BIG5 order
Public Const SORT_CHINESE_UNICODE = &H1                  '  Chinese Unicode order
Public Const SORT_KOREAN_KSC = &H0                       '  Korean KSC order
Public Const SORT_KOREAN_UNICODE = &H1                   '  Korean Unicode order
'  The FILE_READ_DATA and FILE_WRITE_DATA constants are also defined in
'  devioctl.h as FILE_READ_ACCESS and FILE_WRITE_ACCESS. The values for these
'  constants *MUST* always be in sync.
'  The values are redefined in devioctl.h because they must be available to
'  both DOS and NT.
'

Public Const FILE_READ_DATA = (&H1)                     '  file pipe
Public Const FILE_LIST_DIRECTORY = (&H1)                '  directory
Public Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Public Const FILE_ADD_FILE = (&H2)                      '  directory
Public Const FILE_APPEND_DATA = (&H4)                   '  file
Public Const FILE_ADD_SUBDIRECTORY = (&H4)              '  directory
Public Const FILE_CREATE_PIPE_INSTANCE = (&H4)          '  named pipe
Public Const FILE_READ_EA = (&H8)                       '  file directory
Public Const FILE_READ_PROPERTIES = FILE_READ_EA
Public Const FILE_WRITE_EA = (&H10)                     '  file directory
Public Const FILE_WRITE_PROPERTIES = FILE_WRITE_EA
Public Const FILE_EXECUTE = (&H20)                      '  file
Public Const FILE_TRAVERSE = (&H20)                     '  directory
Public Const FILE_DELETE_CHILD = (&H40)                 '  directory
Public Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Public Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all
Public Const FILE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF)
Public Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)
Public Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)
Public Const FILE_GENERIC_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Public Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Public Const MAILSLOT_NO_MESSAGE = (-1)
Public Const MAILSLOT_WAIT_FOREVER = (-1)
Public Const FILE_CASE_SENSITIVE_SEARCH = &H1
Public Const FILE_CASE_PRESERVED_NAMES = &H2
Public Const FILE_UNICODE_ON_DISK = &H4
Public Const FILE_PERSISTENT_ACLS = &H8
Public Const FILE_FILE_COMPRESSION = &H10
Public Const FILE_VOLUME_IS_COMPRESSED = &H8000
Public Const IO_COMPLETION_MODIFY_STATE = &H2
Public Const IO_COMPLETION_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)
Public Const DUPLICATE_CLOSE_SOURCE = &H1
Public Const DUPLICATE_SAME_ACCESS = &H2
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              ACCESS MASK                            //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'
'   Define the access mask as a longword sized structure divided up as
'   follows:
'       typedef struct _ACCESS_MASK {
'           WORD   SpecificRights;
'           Byte  StandardRights;
'           Byte  AccessSystemAcl : 1;
'           Byte  Reserved : 3;
'           Byte  GenericAll : 1;
'           Byte  GenericExecute : 1;
'           Byte  GenericWrite : 1;
'           Byte  GenericRead : 1;
'       } ACCESS_MASK;
'       typedef ACCESS_MASK *PACCESS_MASK;
'
'   but to make life simple for programmer's we'll allow them to specify
'   a desired access mask by simply OR'ing together mulitple single rights
'   and treat an access mask as a DWORD.  For example
'
'       DesiredAccess = DELETE  Or  READ_CONTROL
'
'   So we'll Public Declare ACCESS_MASK as DWORD
'
'  begin_ntddk begin_nthal begin_ntifs
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              ACCESS TYPES                           //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'  begin_ntddk begin_nthal begin_ntifs
'
'   The following are masks for the predefined standard access types
'  AccessSystemAcl access type

Public Const ACCESS_SYSTEM_SECURITY = &H1000000
'  MaximumAllowed access type

Public Const MAXIMUM_ALLOWED = &H2000000
'   These are the generic rights.

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                          ACL  and  ACE                              //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'
'   Define an ACL and the ACE format.  The structure of an ACL header
'   followed by one or more ACEs.  Pictorally the structure of an ACL header
'   is as follows:
'
'   The current AclRevision is defined to be ACL_REVISION.
'
'   AclSize is the size, in bytes, allocated for the ACL.  This includes
'   the ACL header, ACES, and remaining free space in the buffer.
'
'   AceCount is the number of ACES in the ACL.
'
'  begin_ntddk begin_ntifs
'  This is the *current* ACL revision

Public Const ACL_REVISION = (2)
'  This is the history of ACL revisions.  Add a new one whenever
'  ACL_REVISION is updated

Public Const ACL_REVISION1 = (1)
Public Const ACL_REVISION2 = (2)
'
'   The following are the predefined ace types that go into the AceType
'   field of an Ace header.

Public Const ACCESS_ALLOWED_ACE_TYPE = &H0
Public Const ACCESS_DENIED_ACE_TYPE = &H1
Public Const SYSTEM_AUDIT_ACE_TYPE = &H2
Public Const SYSTEM_ALARM_ACE_TYPE = &H3
'   The following are the inherit flags that go into the AceFlags field
'   of an Ace header.

Public Const OBJECT_INHERIT_ACE = &H1
Public Const CONTAINER_INHERIT_ACE = &H2
Public Const NO_PROPAGATE_INHERIT_ACE = &H4
Public Const INHERIT_ONLY_ACE = &H8
Public Const VALID_INHERIT_FLAGS = &HF
'   The following are the currently defined ACE flags that go into the
'   AceFlags field of an ACE header.  Each ACE type has its own set of
'   AceFlags.
'
'   SUCCESSFUL_ACCESS_ACE_FLAG - used only with system audit and alarm ACE
'   types to indicate that a message is generated for successful accesses.
'
'   FAILED_ACCESS_ACE_FLAG - used only with system audit and alarm ACE types
'   to indicate that a message is generated for failed accesses.
'   SYSTEM_AUDIT and SYSTEM_ALARM AceFlags
'
'   These control the signaling of audit and alarms for success or failure.

Public Const SUCCESSFUL_ACCESS_ACE_FLAG = &H40
Public Const FAILED_ACCESS_ACE_FLAG = &H80
'   The following declarations are used for setting and querying information
'   about and ACL.  First are the various information classes available to
'   the user.
'

Public Const AclRevisionInformation = 1
Public Const AclSizeInformation = 2
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              SECURITY_DESCRIPTOR                    //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'
'   Define the Security Descriptor and related data types.
'   This is an opaque data structure.
'
'  begin_ntddk begin_ntifs
'
'  Current security descriptor revision value
'

Public Const SECURITY_DESCRIPTOR_REVISION = (1)
Public Const SECURITY_DESCRIPTOR_REVISION1 = (1)
'  end_ntddk
'
'  Minimum length, in bytes, needed to build a security descriptor
'  (NOTE: This must manually be kept consistent with the)
'  (sizeof(SECURITY_DESCRIPTOR)                         )
'

Public Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Public Const SE_OWNER_DEFAULTED = &H1
Public Const SE_GROUP_DEFAULTED = &H2
Public Const SE_DACL_PRESENT = &H4
Public Const SE_DACL_DEFAULTED = &H8
Public Const SE_SACL_PRESENT = &H10
Public Const SE_SACL_DEFAULTED = &H20
Public Const SE_SELF_RELATIVE = &H8000
'  Where:
'
'      Revision - Contains the revision level of the security
'          descriptor.  This allows this structure to be passed between
'          systems or stored on disk even though it is expected to
'          change in the future.
'
'      Control - A set of flags which qualify the meaning of the
'          security descriptor or individual fields of the security
'          descriptor.
'
'      Owner - is a pointer to an SID representing an object's owner.
'          If this field is null, then no owner SID is present in the
'          security descriptor.  If the security descriptor is in
'          self-relative form, then this field contains an offset to
'          the SID, rather than a pointer.
'
'      Group - is a pointer to an SID representing an object's primary
'          group.  If this field is null, then no primary group SID is
'          present in the security descriptor.  If the security descriptor
'          is in self-relative form, then this field contains an offset to
'          the SID, rather than a pointer.
'
'      Sacl - is a pointer to a system ACL.  This field value is only
'          valid if the DaclPresent control flag is set.  If the
'          SaclPresent flag is set and this field is null, then a null
'          ACL  is specified.  If the security descriptor is in
'          self-relative form, then this field contains an offset to
'          the ACL, rather than a pointer.
'
'      Dacl - is a pointer to a discretionary ACL.  This field value is
'          only valid if the DaclPresent control flag is set.  If the
'          DaclPresent flag is set and this field is null, then a null
'          ACL (unconditionally granting access) is specified.  If the
'          security descriptor is in self-relative form, then this field
'          contains an offset to the ACL, rather than a pointer.
'
' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                Privilege Related Data Structures                    //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'  Privilege attributes
'

Public Const SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000
'
'  Privilege Set Control flags
'

Public Const PRIVILEGE_SET_ALL_NECESSARY = (1)
'//////////////////////////////////////////////////////////////////////
'                                                                     //
'                NT Defined Privileges                                //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

Public Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Public Const SE_TCB_NAME = "SeTcbPrivilege"
Public Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP_NAME = "SeBackupPrivilege"
Public Const SE_RESTORE_NAME = "SeRestorePrivilege"
Public Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"
Public Const SE_AUDIT_NAME = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
' //////////////////////////////////////////////////////////////////
'                                                                 //
'            Security Quality Of Service                          //
'                                                                 //
'                                                                 //
' //////////////////////////////////////////////////////////////////
'  begin_ntddk begin_nthal begin_ntifs
'
'  Impersonation Level
'
'  Impersonation level is represented by a pair of bits in Windows.
'  If a new impersonation level is added or lowest value is changed from
'  0 to something else, fix the Windows CreateFile call.
'

Public Const SecurityAnonymous = 1
Public Const SecurityIdentification = 2
'//////////////////////////////////////////////////////////////////////
'                                                                     //
'                Registry API Constants                                //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
' Reg Data Types...

Public Const REG_NONE = 0                       ' No value type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Public Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Public Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Public Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8

Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore

Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
' Reg Create Type Values...

'Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
'Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
'Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
'Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
'Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
' Reg Key Security Options
' Public Const READ_CONTROL = &H20000

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
'Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
' end winnt.txt
' Debug APIs

Public Const EXCEPTION_DEBUG_EVENT = 1
Public Const CREATE_THREAD_DEBUG_EVENT = 2
Public Const CREATE_PROCESS_DEBUG_EVENT = 3
Public Const EXIT_THREAD_DEBUG_EVENT = 4
Public Const EXIT_PROCESS_DEBUG_EVENT = 5
Public Const LOAD_DLL_DEBUG_EVENT = 6
Public Const UNLOAD_DLL_DEBUG_EVENT = 7
Public Const OUTPUT_DEBUG_STRING_EVENT = 8
Public Const RIP_EVENT = 9
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
' GetDriveType return values

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
Public Const FILE_TYPE_UNKNOWN = &H0
Public Const FILE_TYPE_DISK = &H1
Public Const FILE_TYPE_CHAR = &H2
Public Const FILE_TYPE_PIPE = &H3
Public Const FILE_TYPE_REMOTE = &H8000
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&
Public Const NOPARITY = 0
Public Const ODDPARITY = 1
Public Const EVENPARITY = 2
Public Const MARKPARITY = 3
Public Const SPACEPARITY = 4
Public Const ONESTOPBIT = 0
Public Const ONE5STOPBITS = 1
Public Const TWOSTOPBITS = 2
Public Const IGNORE = 0 '  Ignore signal
Public Const INFINITE = &HFFFF      '  Infinite timeout
' Comm Baud Rate indices

Public Const CBR_110 = 110
Public Const CBR_300 = 300
Public Const CBR_600 = 600
Public Const CBR_1200 = 1200
Public Const CBR_2400 = 2400
Public Const CBR_4800 = 4800
Public Const CBR_9600 = 9600
Public Const CBR_14400 = 14400
Public Const CBR_19200 = 19200
Public Const CBR_38400 = 38400
Public Const CBR_56000 = 56000
Public Const CBR_57600 = 57600
Public Const CBR_115200 = 115200
Public Const CBR_128000 = 128000
Public Const CBR_256000 = 256000
' Error Flags

Public Const CE_RXOVER = &H1                '  Receive Queue overflow
Public Const CE_OVERRUN = &H2               '  Receive Overrun Error
Public Const CE_RXPARITY = &H4              '  Receive Parity Error
Public Const CE_FRAME = &H8                 '  Receive Framing error
Public Const CE_BREAK = &H10                '  Break Detected
Public Const CE_TXFULL = &H100              '  TX Queue is full
Public Const CE_PTO = &H200                 '  LPTx Timeout
Public Const CE_IOE = &H400                 '  LPTx I/O Error
Public Const CE_DNS = &H800                 '  LPTx Device not selected
Public Const CE_OOP = &H1000                '  LPTx Out-Of-Paper
Public Const CE_MODE = &H8000               '  Requested mode unsupported
Public Const IE_BADID = (-1)                '  Invalid or unsupported id
Public Const IE_OPEN = (-2)                 '  Device Already Open
Public Const IE_NOPEN = (-3)                '  Device Not Open
Public Const IE_MEMORY = (-4)               '  Unable to allocate queues
Public Const IE_DEFAULT = (-5)              '  Error in default parameters
Public Const IE_HARDWARE = (-10)            '  Hardware Not Present
Public Const IE_BYTESIZE = (-11)            '  Illegal Byte Size
Public Const IE_BAUDRATE = (-12)            '  Unsupported BaudRate
' Events

Public Const EV_RXCHAR = &H1                '  Any Character received
Public Const EV_RXFLAG = &H2                '  Received certain character
Public Const EV_TXEMPTY = &H4               '  Transmitt Queue Empty
Public Const EV_CTS = &H8                   '  CTS changed state
Public Const EV_DSR = &H10                  '  DSR changed state
Public Const EV_RLSD = &H20                 '  RLSD changed state
Public Const EV_BREAK = &H40                '  BREAK received
Public Const EV_ERR = &H80                  '  Line status error occurred
Public Const EV_RING = &H100                '  Ring signal detected
Public Const EV_PERR = &H200                '  Printer error occured
Public Const EV_RX80FULL = &H400            '  Receive buffer is 80 percent full
Public Const EV_EVENT1 = &H800              '  Provider specific event 1
Public Const EV_EVENT2 = &H1000             '  Provider specific event 2
' Escape Functions

Public Const SETXOFF = 1  '  Simulate XOFF received
Public Const SETXON = 2 '  Simulate XON received
Public Const SETRTS = 3 '  Set RTS high
Public Const CLRRTS = 4 '  Set RTS low
Public Const SETDTR = 5 '  Set DTR high
Public Const CLRDTR = 6 '  Set DTR low
Public Const RESETDEV = 7       '  Reset device if possible
Public Const SETBREAK = 8  'Set the device break line
Public Const CLRBREAK = 9 ' Clear the device break line
'  PURGE function flags.

Public Const PURGE_TXABORT = &H1     '  Kill the pending/current writes to the comm port.
Public Const PURGE_RXABORT = &H2     '  Kill the pending/current reads to the comm port.
Public Const PURGE_TXCLEAR = &H4     '  Kill the transmit queue if there.
Public Const PURGE_RXCLEAR = &H8     '  Kill the typeahead buffer if there.
Public Const LPTx = &H80        '  Set if ID is for LPT device
'  Modem Status Flags

Public Const MS_CTS_ON = &H10&
Public Const MS_DSR_ON = &H20&
Public Const MS_RING_ON = &H40&
Public Const MS_RLSD_ON = &H80&
' WaitSoundState() Constants

Public Const S_QUEUEEMPTY = 0
Public Const S_THRESHOLD = 1
Public Const S_ALLTHRESHOLD = 2
' Accent Modes

Public Const S_NORMAL = 0
Public Const S_LEGATO = 1
Public Const S_STACCATO = 2
' SetSoundNoise() Sources

Public Const S_PERIOD512 = 0    '  Freq = N/512 high pitch, less coarse hiss
Public Const S_PERIOD1024 = 1   '  Freq = N/1024
Public Const S_PERIOD2048 = 2   '  Freq = N/2048 low pitch, more coarse hiss
Public Const S_PERIODVOICE = 3  '  Source is frequency from voice channel (3)
Public Const S_WHITE512 = 4     '  Freq = N/512 high pitch, less coarse hiss
Public Const S_WHITE1024 = 5    '  Freq = N/1024
Public Const S_WHITE2048 = 6    '  Freq = N/2048 low pitch, more coarse hiss
Public Const S_WHITEVOICE = 7   '  Source is frequency from voice channel (3)
Public Const S_SERDVNA = (-1)   '  Device not available
Public Const S_SEROFM = (-2)    '  Out of memory
Public Const S_SERMACT = (-3)   '  Music active
Public Const S_SERQFUL = (-4)   '  Queue full
Public Const S_SERBDNT = (-5)   '  Invalid note
Public Const S_SERDLN = (-6)    '  Invalid note length
Public Const S_SERDCC = (-7)    '  Invalid note count
Public Const S_SERDTP = (-8)    '  Invalid tempo
Public Const S_SERDVL = (-9)    '  Invalid volume
Public Const S_SERDMD = (-10)   '  Invalid mode
Public Const S_SERDSH = (-11)   '  Invalid shape
Public Const S_SERDPT = (-12)   '  Invalid pitch
Public Const S_SERDFQ = (-13)   '  Invalid frequency
Public Const S_SERDDR = (-14)   '  Invalid duration
Public Const S_SERDSR = (-15)   '  Invalid source
Public Const S_SERDST = (-16)   '  Invalid state
Public Const NMPWAIT_WAIT_FOREVER = &HFFFF
Public Const NMPWAIT_NOWAIT = &H1
Public Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Public Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Public Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Public Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Public Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_COPY = SECTION_QUERY
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Public Const FILE_MAP_READ = SECTION_MAP_READ
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS
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
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
Public Const PROCESSOR_ARCHITECTURE_INTEL = 0
Public Const PROCESSOR_ARCHITECTURE_MIPS = 1
Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Public Const PROCESSOR_ARCHITECTURE_PPC = 3
Public Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF
' Flags for DrawFrameControl

Public Const DFC_CAPTION = 1
Public Const DFC_MENU = 2
Public Const DFC_SCROLL = 3
Public Const DFC_BUTTON = 4
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_MENUARROW = &H0
Public Const DFCS_MENUCHECK = &H1
Public Const DFCS_MENUBULLET = &H2
Public Const DFCS_MENUARROWRIGHT = &H4
Public Const DFCS_SCROLLUP = &H0
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_SCROLLLEFT = &H2
Public Const DFCS_SCROLLRIGHT = &H3
Public Const DFCS_SCROLLCOMBOBOX = &H5
Public Const DFCS_SCROLLSIZEGRIP = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_INACTIVE = &H100
Public Const DFCS_PUSHED = &H200
Public Const DFCS_CHECKED = &H400
Public Const DFCS_ADJUSTRECT = &H2000
Public Const DFCS_FLAT = &H4000
Public Const DFCS_MONO = &H8000
Public Const DONT_RESOLVE_DLL_REFERENCES = &H1
' GetTempFileName() Flags
'

Public Const TF_FORCEDRIVE = &H80
Public Const LOCKFILE_FAIL_IMMEDIATELY = &H1
Public Const LOCKFILE_EXCLUSIVE_LOCK = &H2
Public Const LNOTIFY_OUTOFMEM = 0
Public Const LNOTIFY_MOVE = 1
Public Const LNOTIFY_DISCARD = 2
Public Const SLE_ERROR = &H1
Public Const SLE_MINORERROR = &H2
Public Const SLE_WARNING = &H3
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000
' Predefined Resource Types

Public Const RT_CURSOR = 1&
Public Const RT_BITMAP = 2&
Public Const RT_ICON = 3&
Public Const RT_MENU = 4&
Public Const RT_DIALOG = 5&
Public Const RT_STRING = 6&
Public Const RT_FONTDIR = 7&
Public Const RT_FONT = 8&
Public Const RT_ACCELERATOR = 9&
Public Const RT_RCDATA = 10&
Public Const DDD_RAW_TARGET_PATH = &H1
Public Const DDD_REMOVE_DEFINITION = &H2
Public Const DDD_EXACT_MATCH_ON_REMOVE = &H4
Public Const MAX_PATH = 260
Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const MOVEFILE_COPY_ALLOWED = &H2
Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
' Security APIs

Public Const TokenUser = 1
Public Const TokenGroups = 2
Public Const TokenPrivileges = 3
Public Const TokenOwner = 4
Public Const TokenPrimaryGroup = 5
Public Const TokenDefaultDacl = 6
Public Const TokenSource = 7
Public Const TokenType = 8
Public Const TokenImpersonationLevel = 9
Public Const TokenStatistics = 10
Public Const GET_TAPE_MEDIA_INFORMATION = 0
Public Const GET_TAPE_DRIVE_INFORMATION = 1
Public Const SET_TAPE_MEDIA_INFORMATION = 0
Public Const SET_TAPE_DRIVE_INFORMATION = 1
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const TLS_OUT_OF_INDEXES = &HFFFF
' Stream IDs

Public Const BACKUP_DATA = &H1
Public Const BACKUP_EA_DATA = &H2
Public Const BACKUP_SECURITY_DATA = &H3
Public Const BACKUP_ALTERNATE_DATA = &H4
Public Const BACKUP_LINK = &H5
'   Stream Attributes

Public Const STREAM_MODIFIED_WHEN_READ = &H1
Public Const STREAM_CONTAINS_SECURITY = &H2
'  Dual Mode API below this line. Dual Mode Types also included.

Public Const STARTF_USESHOWWINDOW = &H1
Public Const STARTF_USESIZE = &H2
Public Const STARTF_USEPOSITION = &H4
Public Const STARTF_USECOUNTCHARS = &H8
Public Const STARTF_USEFILLATTRIBUTE = &H10
Public Const STARTF_RUNFULLSCREEN = &H20        '  ignored for non-x86 platforms
Public Const STARTF_FORCEONFEEDBACK = &H40
Public Const STARTF_FORCEOFFFEEDBACK = &H80
Public Const STARTF_USESTDHANDLES = &H100
Public Const SHUTDOWN_NORETRY = &H1
'  Abnormal termination codes

Public Const TC_NORMAL = 0
Public Const TC_HARDERR = 1
Public Const TC_GP_TRAP = 2
Public Const TC_SIGNAL = 3
' Procedure declarations, constant definitions, and macros
' for the NLS component
' String Length Maximums

Public Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.
' MBCS and Unicode Translation Flags.

Public Const MB_PRECOMPOSED = &H1         '  use precomposed chars
Public Const MB_COMPOSITE = &H2         '  use composite chars
Public Const MB_USEGLYPHCHARS = &H4         '  use glyph chars, not ctrl chars
Public Const WC_DEFAULTCHECK = &H100       '  check for default char
Public Const WC_COMPOSITECHECK = &H200       '  convert composite to precomposed
Public Const WC_DISCARDNS = &H10        '  discard non-spacing chars
Public Const WC_SEPCHARS = &H20        '  generate separate chars
Public Const WC_DEFAULTCHAR = &H40        '  replace w/ default char
' Character Type Flags.

Public Const CT_CTYPE1 = &H1         '  ctype 1 information
Public Const CT_CTYPE2 = &H2         '  ctype 2 information
Public Const CT_CTYPE3 = &H4         '  ctype 3 information
' CType 1 Flag Bits.

Public Const C1_UPPER = &H1     '  upper case
Public Const C1_LOWER = &H2     '  lower case
Public Const C1_DIGIT = &H4     '  decimal digits
Public Const C1_SPACE = &H8     '  spacing characters
Public Const C1_PUNCT = &H10    '  punctuation characters
Public Const C1_CNTRL = &H20    '  control characters
Public Const C1_BLANK = &H40    '  blank characters
Public Const C1_XDIGIT = &H80    '  other digits
Public Const C1_ALPHA = &H100   '  any letter
' CType 2 Flag Bits.

Public Const C2_LEFTTORIGHT = &H1     '  left to right
Public Const C2_RIGHTTOLEFT = &H2     '  right to left
Public Const C2_EUROPENUMBER = &H3     '  European number, digit
Public Const C2_EUROPESEPARATOR = &H4     '  European numeric separator
Public Const C2_EUROPETERMINATOR = &H5     '  European numeric terminator
Public Const C2_ARABICNUMBER = &H6     '  Arabic number
Public Const C2_COMMONSEPARATOR = &H7     '  common numeric separator
Public Const C2_BLOCKSEPARATOR = &H8     '  block separator
Public Const C2_SEGMENTSEPARATOR = &H9     '  segment separator
Public Const C2_WHITESPACE = &HA     '  white space
Public Const C2_OTHERNEUTRAL = &HB     '  other neutrals
Public Const C2_NOTAPPLICABLE = &H0     '  no implicit directionality
' CType 3 Flag Bits.

Public Const C3_NONSPACING = &H1     '  nonspacing character
Public Const C3_DIACRITIC = &H2     '  diacritic mark
Public Const C3_VOWELMARK = &H4     '  vowel mark
Public Const C3_SYMBOL = &H8     '  symbols
Public Const C3_NOTAPPLICABLE = &H0     '  ctype 3 is not applicable
' String Flags.

Public Const NORM_IGNORECASE = &H1         '  ignore case
Public Const NORM_IGNORENONSPACE = &H2         '  ignore nonspacing chars
Public Const NORM_IGNORESYMBOLS = &H4         '  ignore symbols
' Locale Independent Mapping Flags.

Public Const MAP_FOLDCZONE = &H10        '  fold compatibility zone chars
Public Const MAP_PRECOMPOSED = &H20        '  convert to precomposed chars
Public Const MAP_COMPOSITE = &H40        '  convert to composite chars
Public Const MAP_FOLDDIGITS = &H80        '  all digits to ASCII 0-9
' Locale Dependent Mapping Flags.

Public Const LCMAP_LOWERCASE = &H100       '  lower case letters
Public Const LCMAP_UPPERCASE = &H200       '  upper case letters
Public Const LCMAP_SORTKEY = &H400       '  WC sort key (normalize)
Public Const LCMAP_BYTEREV = &H800       '  byte reversal
' Sorting Flags.

Public Const SORT_STRINGSORT = &H1000      '  use string sort method
' Code Page Default Values.

Public Const CP_ACP = 0  '  default to ANSI code page
Public Const CP_OEMCP = 1  '  default to OEM  code page
' Country Codes.

Public Const CTRY_DEFAULT = 0
Public Const CTRY_AUSTRALIA = 61  '  Australia
Public Const CTRY_AUSTRIA = 43  '  Austria
Public Const CTRY_BELGIUM = 32  '  Belgium
Public Const CTRY_BRAZIL = 55  '  Brazil
Public Const CTRY_CANADA = 2  '  Canada
Public Const CTRY_DENMARK = 45  '  Denmark
Public Const CTRY_FINLAND = 358  '  Finland
Public Const CTRY_FRANCE = 33  '  France
Public Const CTRY_GERMANY = 49  '  Germany
Public Const CTRY_ICELAND = 354  '  Iceland
Public Const CTRY_IRELAND = 353  '  Ireland
Public Const CTRY_ITALY = 39  '  Italy
Public Const CTRY_JAPAN = 81  '  Japan
Public Const CTRY_MEXICO = 52  '  Mexico
Public Const CTRY_NETHERLANDS = 31  '  Netherlands
Public Const CTRY_NEW_ZEALAND = 64  '  New Zealand
Public Const CTRY_NORWAY = 47  '  Norway
Public Const CTRY_PORTUGAL = 351  '  Portugal
Public Const CTRY_PRCHINA = 86  '  PR China
Public Const CTRY_SOUTH_KOREA = 82  '  South Korea
Public Const CTRY_SPAIN = 34  '  Spain
Public Const CTRY_SWEDEN = 46  '  Sweden
Public Const CTRY_SWITZERLAND = 41  '  Switzerland
Public Const CTRY_TAIWAN = 886  '  Taiwan
Public Const CTRY_UNITED_KINGDOM = 44  '  United Kingdom
Public Const CTRY_UNITED_STATES = 1  '  United States
' Locale Types.
' These types are used for the GetLocaleInfoW NLS API routine.
' LOCALE_NOUSEROVERRIDE is also used in GetTimeFormatW and GetDateFormatW.

Public Const LOCALE_NOUSEROVERRIDE = &H80000000  '  do not use user overrides
Public Const LOCALE_ILANGUAGE = &H1         '  language id
Public Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Public Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Public Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Public Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Public Const LOCALE_ICOUNTRY = &H5         '  country code
Public Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Public Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Public Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Public Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Public Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Public Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page
Public Const LOCALE_SLIST = &HC         '  list item separator
Public Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Public Const LOCALE_STHOUSAND = &HF         '  thousand separator
Public Const LOCALE_SGROUPING = &H10        '  digit grouping
Public Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Public Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Public Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
Public Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Public Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Public Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Public Const LOCALE_INEGCURR = &H1C        '  negative currency mode
Public Const LOCALE_SDATE = &H1D        '  date separator
Public Const LOCALE_STIME = &H1E        '  time separator
Public Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Public Const LOCALE_IDATE = &H21        '  short date format ordering
Public Const LOCALE_ILDATE = &H22        '  long date format ordering
Public Const LOCALE_ITIME = &H23        '  time format specifier
Public Const LOCALE_ICENTURY = &H24        '  century format specifier
Public Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Public Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Public Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Public Const LOCALE_S1159 = &H28        '  AM designator
Public Const LOCALE_S2359 = &H29        '  PM designator
Public Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Public Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Public Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Public Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Public Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Public Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Public Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Public Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Public Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Public Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Public Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Public Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Public Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Public Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Public Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Public Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Public Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Public Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Public Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Public Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Public Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Public Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Public Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Public Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Public Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Public Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Public Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Public Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Public Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Public Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Public Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Public Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Public Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt
' Time Flags for GetTimeFormatW.

Public Const TIME_NOMINUTESORSECONDS = &H1         '  do not use minutes or seconds
Public Const TIME_NOSECONDS = &H2         '  do not use seconds
Public Const TIME_NOTIMEMARKER = &H4         '  do not use time marker
Public Const TIME_FORCE24HOURFORMAT = &H8         '  always use 24 hour format
' Date Flags for GetDateFormatW.

Public Const DATE_SHORTDATE = &H1         '  use short date picture
Public Const DATE_LONGDATE = &H2         '  use long date picture
' *************************************************************************
' *                                                                         *
' * winnls.h -- NLS procedure declarations, constant definitions and macros *
' *                                                                         *
' * Copyright (c) 1991-1995, Microsoft Corp. All rights reserved.           *
' *                                                                         *
' **************************************************************************/
' *  Calendar Types.
'  *
'  *  These types are used for the GetALTCalendarInfoW NLS API routine.
'  */

Public Const MAX_DEFAULTCHAR = 2
Public Const CAL_ICALINTVALUE = &H1                     '  calendar type
Public Const CAL_SCALNAME = &H2                         '  native name of calendar
Public Const CAL_IYEAROFFSETRANGE = &H3                 '  starting years of eras
Public Const CAL_SERASTRING = &H4                       '  era name for IYearOffsetRanges
Public Const CAL_SSHORTDATE = &H5                       '  Integer date format string
Public Const CAL_SLONGDATE = &H6                        '  long date format string
Public Const CAL_SDAYNAME1 = &H7                        '  native name for Monday
Public Const CAL_SDAYNAME2 = &H8                        '  native name for Tuesday
Public Const CAL_SDAYNAME3 = &H9                        '  native name for Wednesday
Public Const CAL_SDAYNAME4 = &HA                        '  native name for Thursday
Public Const CAL_SDAYNAME5 = &HB                        '  native name for Friday
Public Const CAL_SDAYNAME6 = &HC                        '  native name for Saturday
Public Const CAL_SDAYNAME7 = &HD                        '  native name for Sunday
Public Const CAL_SABBREVDAYNAME1 = &HE                  '  abbreviated name for Monday
Public Const CAL_SABBREVDAYNAME2 = &HF                  '  abbreviated name for Tuesday
Public Const CAL_SABBREVDAYNAME3 = &H10                 '  abbreviated name for Wednesday
Public Const CAL_SABBREVDAYNAME4 = &H11                 '  abbreviated name for Thursday
Public Const CAL_SABBREVDAYNAME5 = &H12                 '  abbreviated name for Friday
Public Const CAL_SABBREVDAYNAME6 = &H13                 '  abbreviated name for Saturday
Public Const CAL_SABBREVDAYNAME7 = &H14                 '  abbreviated name for Sunday
Public Const CAL_SMONTHNAME1 = &H15                     '  native name for January
Public Const CAL_SMONTHNAME2 = &H16                     '  native name for February
Public Const CAL_SMONTHNAME3 = &H17                     '  native name for March
Public Const CAL_SMONTHNAME4 = &H18                     '  native name for April
Public Const CAL_SMONTHNAME5 = &H19                     '  native name for May
Public Const CAL_SMONTHNAME6 = &H1A                     '  native name for June
Public Const CAL_SMONTHNAME7 = &H1B                     '  native name for July
Public Const CAL_SMONTHNAME8 = &H1C                     '  native name for August
Public Const CAL_SMONTHNAME9 = &H1D                     '  native name for September
Public Const CAL_SMONTHNAME10 = &H1E                    '  native name for October
Public Const CAL_SMONTHNAME11 = &H1F                    '  native name for November
Public Const CAL_SMONTHNAME12 = &H20                    '  native name for December
Public Const CAL_SMONTHNAME13 = &H21                    '  native name for 13th month (if any)
Public Const CAL_SABBREVMONTHNAME1 = &H22               '  abbreviated name for January
Public Const CAL_SABBREVMONTHNAME2 = &H23               '  abbreviated name for February
Public Const CAL_SABBREVMONTHNAME3 = &H24               '  abbreviated name for March
Public Const CAL_SABBREVMONTHNAME4 = &H25               '  abbreviated name for April
Public Const CAL_SABBREVMONTHNAME5 = &H26               '  abbreviated name for May
Public Const CAL_SABBREVMONTHNAME6 = &H27               '  abbreviated name for June
Public Const CAL_SABBREVMONTHNAME7 = &H28               '  abbreviated name for July
Public Const CAL_SABBREVMONTHNAME8 = &H29               '  abbreviated name for August
Public Const CAL_SABBREVMONTHNAME9 = &H2A               '  abbreviated name for September
Public Const CAL_SABBREVMONTHNAME10 = &H2B              '  abbreviated name for October
Public Const CAL_SABBREVMONTHNAME11 = &H2C              '  abbreviated name for November
Public Const CAL_SABBREVMONTHNAME12 = &H2D              '  abbreviated name for December
Public Const CAL_SABBREVMONTHNAME13 = &H2E              '  abbreviated name for 13th month (if any)
'
'  *  Calendar Enumeration Value.
'  */

Public Const ENUM_ALL_CALENDARS = &HFFFF                '  enumerate all calendars
'
'  *  Calendar ID Values.
'  */

Public Const CAL_GREGORIAN = 1                 '  Gregorian (localized) calendar
Public Const CAL_GREGORIAN_US = 2              '  Gregorian (U.S.) calendar
Public Const CAL_JAPAN = 3                     '  Japanese Emperor Era calendar
Public Const CAL_TAIWAN = 4                    '  Taiwan Region Era calendar
Public Const CAL_KOREA = 5                     '  Korean Tangun Era calendar
'  ControlKeyState flags

Public Const RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
Public Const LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
Public Const RIGHT_CTRL_PRESSED = &H4     '  the right ctrl key is pressed.
Public Const LEFT_CTRL_PRESSED = &H8     '  the left ctrl key is pressed.
Public Const SHIFT_PRESSED = &H10    '  the shift key is pressed.
Public Const NUMLOCK_ON = &H20    '  the numlock light is on.
Public Const SCROLLLOCK_ON = &H40    '  the scrolllock light is on.
Public Const CAPSLOCK_ON = &H80    '  the capslock light is on.
Public Const ENHANCED_KEY = &H100   '  the key is enhanced.
'  ButtonState flags

Public Const FROM_LEFT_1ST_BUTTON_PRESSED = &H1
Public Const RIGHTMOST_BUTTON_PRESSED = &H2
Public Const FROM_LEFT_2ND_BUTTON_PRESSED = &H4
Public Const FROM_LEFT_3RD_BUTTON_PRESSED = &H8
Public Const FROM_LEFT_4TH_BUTTON_PRESSED = &H10
'  EventFlags

Public Const MOUSE_MOVED = &H1
Public Const DOUBLE_CLICK = &H2
'   EventType flags:

Public Const KEY_EVENT = &H1     '  Event contains key event record
Public Const mouse_eventC = &H2     '  Event contains mouse event record
Public Const WINDOW_BUFFER_SIZE_EVENT = &H4     '  Event contains window change event record
Public Const MENU_EVENT = &H8     '  Event contains menu event record
Public Const FOCUS_EVENT = &H10    '  event contains focus change
'  Attributes flags:

Public Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Public Const FOREGROUND_GREEN = &H2     '  text color contains green.
Public Const FOREGROUND_RED = &H4     '  text color contains red.
Public Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Public Const BACKGROUND_BLUE = &H10    '  background color contains blue.
Public Const BACKGROUND_GREEN = &H20    '  background color contains green.
Public Const BACKGROUND_RED = &H40    '  background color contains red.
Public Const BACKGROUND_INTENSITY = &H80    '  background color is intensified.
Public Const CTRL_C_EVENT = 0
Public Const CTRL_BREAK_EVENT = 1
Public Const CTRL_CLOSE_EVENT = 2
'  3 is reserved!
'  4 is reserved!

Public Const CTRL_LOGOFF_EVENT = 5
Public Const CTRL_SHUTDOWN_EVENT = 6
' Input Mode flags:

Public Const ENABLE_PROCESSED_INPUT = &H1
Public Const ENABLE_LINE_INPUT = &H2
Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_WINDOW_INPUT = &H8
Public Const ENABLE_MOUSE_INPUT = &H10
' Output Mode flags:

Public Const ENABLE_PROCESSED_OUTPUT = &H1
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Public Const CONSOLE_TEXTMODE_BUFFER = 1
' -------------
'  GDI Section
' -------------
' Binary raster ops

Public Const R2_BLACK = 1       '   0
Public Const R2_NOTMERGEPEN = 2 '  DPon
Public Const R2_MASKNOTPEN = 3  '  DPna
Public Const R2_NOTCOPYPEN = 4  '  PN
Public Const R2_MASKPENNOT = 5  '  PDna
Public Const R2_NOT = 6 '  Dn
Public Const R2_XORPEN = 7      '  DPx
Public Const R2_NOTMASKPEN = 8  '  DPan
Public Const R2_MASKPEN = 9     '  DPa
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Const R2_NOP = 11        '  D
Public Const R2_MERGENOTPEN = 12        '  DPno
Public Const R2_COPYPEN = 13    '  P
Public Const R2_MERGEPENNOT = 14        '  PDno
Public Const R2_MERGEPEN = 15   '  DPo
Public Const R2_WHITE = 16      '   1
Public Const R2_LAST = 16
'  Ternary raster operations

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Public Const GDI_ERROR = &HFFFF
Public Const HGDI_ERROR = &HFFFF
' Region Flags

Public Const ERRORAPI = 0
Public Const NULLREGION = 1
Public Const SIMPLEREGION = 2
Public Const COMPLEXREGION = 3
' CombineRgn() Styles

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const RGN_MIN = RGN_AND
Public Const RGN_MAX = RGN_COPY
' StretchBlt() Modes

Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
' PolyFill() Modes

Public Const ALTERNATE = 1
Public Const WINDING = 2
Public Const POLYFILL_LAST = 2
' Text Alignment Options

Public Const TA_NOUPDATECP = 0
Public Const TA_UPDATECP = 1
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Public Const VTA_BASELINE = TA_BASELINE
Public Const VTA_LEFT = TA_BOTTOM
Public Const VTA_RIGHT = TA_TOP
Public Const VTA_CENTER = TA_CENTER
Public Const VTA_BOTTOM = TA_RIGHT
Public Const VTA_TOP = TA_LEFT
Public Const ETO_GRAYED = 1
Public Const ETO_OPAQUE = 2
Public Const ETO_CLIPPED = 4
Public Const ASPECT_FILTERING = &H1
Public Const DCB_RESET = &H1
Public Const DCB_ACCUMULATE = &H2
Public Const DCB_DIRTY = DCB_ACCUMULATE
Public Const DCB_SET = (DCB_RESET Or DCB_ACCUMULATE)
Public Const DCB_ENABLE = &H4
Public Const DCB_DISABLE = &H8
' Metafile Functions

Public Const META_SETBKCOLOR = &H201
Public Const META_SETBKMODE = &H102
Public Const META_SETMAPMODE = &H103
Public Const META_SETROP2 = &H104
Public Const META_SETRELABS = &H105
Public Const META_SETPOLYFILLMODE = &H106
Public Const META_SETSTRETCHBLTMODE = &H107
Public Const META_SETTEXTCHAREXTRA = &H108
Public Const META_SETTEXTCOLOR = &H209
Public Const META_SETTEXTJUSTIFICATION = &H20A
Public Const META_SETWINDOWORG = &H20B
Public Const META_SETWINDOWEXT = &H20C
Public Const META_SETVIEWPORTORG = &H20D
Public Const META_SETVIEWPORTEXT = &H20E
Public Const META_OFFSETWINDOWORG = &H20F
Public Const META_SCALEWINDOWEXT = &H410
Public Const META_OFFSETVIEWPORTORG = &H211
Public Const META_SCALEVIEWPORTEXT = &H412
Public Const META_LINETO = &H213
Public Const META_MOVETO = &H214
Public Const META_EXCLUDECLIPRECT = &H415
Public Const META_INTERSECTCLIPRECT = &H416
Public Const META_ARC = &H817
Public Const META_ELLIPSE = &H418
Public Const META_FLOODFILL = &H419
Public Const META_PIE = &H81A
Public Const META_RECTANGLE = &H41B
Public Const META_ROUNDRECT = &H61C
Public Const META_PATBLT = &H61D
Public Const META_SAVEDC = &H1E
Public Const META_SETPIXEL = &H41F
Public Const META_OFFSETCLIPRGN = &H220
Public Const META_TEXTOUT = &H521
Public Const META_BITBLT = &H922
Public Const META_STRETCHBLT = &HB23
Public Const META_POLYGON = &H324
Public Const META_POLYLINE = &H325
Public Const META_ESCAPE = &H626
Public Const META_RESTOREDC = &H127
Public Const META_FILLREGION = &H228
Public Const META_FRAMEREGION = &H429
Public Const META_INVERTREGION = &H12A
Public Const META_PAINTREGION = &H12B
Public Const META_SELECTCLIPREGION = &H12C
Public Const META_SELECTOBJECT = &H12D
Public Const META_SETTEXTALIGN = &H12E
Public Const META_CHORD = &H830
Public Const META_SETMAPPERFLAGS = &H231
Public Const META_EXTTEXTOUT = &HA32
Public Const META_SETDIBTODEV = &HD33
Public Const META_SELECTPALETTE = &H234
Public Const META_REALIZEPALETTE = &H35
Public Const META_ANIMATEPALETTE = &H436
Public Const META_SETPALENTRIES = &H37
Public Const META_POLYPOLYGON = &H538
Public Const META_RESIZEPALETTE = &H139
Public Const META_DIBBITBLT = &H940
Public Const META_DIBSTRETCHBLT = &HB41
Public Const META_DIBCREATEPATTERNBRUSH = &H142
Public Const META_STRETCHDIB = &HF43
Public Const META_EXTFLOODFILL = &H548
Public Const META_DELETEOBJECT = &H1F0
Public Const META_CREATEPALETTE = &HF7
Public Const META_CREATEPATTERNBRUSH = &H1F9
Public Const META_CREATEPENINDIRECT = &H2FA
Public Const META_CREATEFONTINDIRECT = &H2FB
Public Const META_CREATEBRUSHINDIRECT = &H2FC
Public Const META_CREATEREGION = &H6FF
' GDI Escapes

Public Const NEWFRAME = 1
Public Const AbortDocC = 2
Public Const NEXTBAND = 3
Public Const SETCOLORTABLE = 4
Public Const GETCOLORTABLE = 5
Public Const FLUSHOUTPUT = 6
Public Const DRAFTMODE = 7
Public Const QUERYESCSUPPORT = 8
Public Const SETABORTPROC = 9
Public Const StartDocC = 10
Public Const EndDocC = 11
Public Const GETPHYSPAGESIZE = 12
Public Const GETPRINTINGOFFSET = 13
Public Const GETSCALINGFACTOR = 14
Public Const MFCOMMENT = 15
Public Const GETPENWIDTH = 16
Public Const SETCOPYCOUNT = 17
Public Const SELECTPAPERSOURCE = 18
Public Const DEVICEDATA = 19
Public Const PASSTHROUGH = 19
Public Const GETTECHNOLGY = 20
Public Const GETTECHNOLOGY = 20
Public Const SETLINECAP = 21
Public Const SETLINEJOIN = 22
Public Const SetMiterLimitC = 23
Public Const BANDINFO = 24
Public Const DRAWPATTERNRECT = 25
Public Const GETVECTORPENSIZE = 26
Public Const GETVECTORBRUSHSIZE = 27
Public Const ENABLEDUPLEX = 28
Public Const GETSETPAPERBINS = 29
Public Const GETSETPRINTORIENT = 30
Public Const ENUMPAPERBINS = 31
Public Const SETDIBSCALING = 32
Public Const EPSPRINTING = 33
Public Const ENUMPAPERMETRICS = 34
Public Const GETSETPAPERMETRICS = 35
Public Const POSTSCRIPT_DATA = 37
Public Const POSTSCRIPT_IGNORE = 38
Public Const MOUSETRAILS = 39
Public Const GETDEVICEUNITS = 42
Public Const GETEXTENDEDTEXTMETRICS = 256
Public Const GETEXTENTTABLE = 257
Public Const GETPAIRKERNTABLE = 258
Public Const GETTRACKKERNTABLE = 259
Public Const ExtTextOutC = 512
Public Const GETFACENAME = 513
Public Const DOWNLOADFACE = 514
Public Const ENABLERELATIVEWIDTHS = 768
Public Const ENABLEPAIRKERNING = 769
Public Const SETKERNTRACK = 770
Public Const SETALLJUSTVALUES = 771
Public Const SETCHARSET = 772
Public Const StretchBltC = 2048
Public Const GETSETSCREENPARAMS = 3072
Public Const BEGIN_PATH = 4096
Public Const CLIP_TO_PATH = 4097
Public Const END_PATH = 4098
Public Const EXT_DEVICE_CAPS = 4099
Public Const RESTORE_CTM = 4100
Public Const SAVE_CTM = 4101
Public Const SET_ARC_DIRECTION = 4102
Public Const SET_BACKGROUND_COLOR = 4103
Public Const SET_POLY_MODE = 4104
Public Const SET_SCREEN_ANGLE = 4105
Public Const SET_SPREAD = 4106
Public Const TRANSFORM_CTM = 4107
Public Const SET_CLIP_BOX = 4108
Public Const SET_BOUNDS = 4109
Public Const SET_MIRROR_MODE = 4110
Public Const OPENCHANNEL = 4110
Public Const DOWNLOADHEADER = 4111
Public Const CLOSECHANNEL = 4112
Public Const POSTSCRIPT_PASSTHROUGH = 4115
Public Const ENCAPSULATED_POSTSCRIPT = 4116
' Spooler Error Codes

Public Const SP_NOTREPORTED = &H4000
Public Const SP_ERROR = (-1)
Public Const SP_APPABORT = (-2)
Public Const SP_USERABORT = (-3)
Public Const SP_OUTOFDISK = (-4)
Public Const SP_OUTOFMEMORY = (-5)
Public Const PR_JOBSTATUS = &H0
'  Object Definitions for EnumObjects()

Public Const OBJ_PEN = 1
Public Const OBJ_BRUSH = 2
Public Const OBJ_DC = 3
Public Const OBJ_METADC = 4
Public Const OBJ_PAL = 5
Public Const OBJ_FONT = 6
Public Const OBJ_BITMAP = 7
Public Const OBJ_REGION = 8
Public Const OBJ_METAFILE = 9
Public Const OBJ_MEMDC = 10
Public Const OBJ_EXTPEN = 11
Public Const OBJ_ENHMETADC = 12
Public Const OBJ_ENHMETAFILE = 13
'  xform stuff

Public Const MWT_IDENTITY = 1
Public Const MWT_LEFTMULTIPLY = 2
Public Const MWT_RIGHTMULTIPLY = 3
Public Const MWT_MIN = MWT_IDENTITY
Public Const MWT_MAX = MWT_RIGHTMULTIPLY
' constants for the biCompression field

Public Const BI_RGB = 0&
Public Const BI_RLE8 = 1&
Public Const BI_RLE4 = 2&
Public Const BI_bitfields = 3&
' ntmFlags field flags

Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&
'  tmPitchAndFamily flags

Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
' Logical Font

Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_PRECIS = 4
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_OUTLINE_PRECIS = 8
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_MASK = &HF
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_TT_ALWAYS = 32
Public Const CLIP_EMBEDDED = 128
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2
Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
' Font Families
'

Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.
' Times Roman, Century Schoolbook, etc.

Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.
' Helvetica, Swiss, etc.

Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.
' Pica, Elite, Courier, etc.

Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.
' Font Weights

Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_ULTRALIGHT = FW_EXTRALIGHT
Public Const FW_REGULAR = FW_NORMAL
Public Const FW_DEMIBOLD = FW_SEMIBOLD
Public Const FW_ULTRABOLD = FW_EXTRABOLD
Public Const FW_BLACK = FW_HEAVY
Public Const PANOSE_COUNT = 10
Public Const PAN_FAMILYTYPE_INDEX = 0
Public Const PAN_SERIFSTYLE_INDEX = 1
Public Const PAN_WEIGHT_INDEX = 2
Public Const PAN_PROPORTION_INDEX = 3
Public Const PAN_CONTRAST_INDEX = 4
Public Const PAN_STROKEVARIATION_INDEX = 5
Public Const PAN_ARMSTYLE_INDEX = 6
Public Const PAN_LETTERFORM_INDEX = 7
Public Const PAN_MIDLINE_INDEX = 8
Public Const PAN_XHEIGHT_INDEX = 9
Public Const PAN_CULTURE_LATIN = 0
Public Const PAN_ANY = 0  '  Any
Public Const PAN_NO_FIT = 1  '  No Fit
Public Const PAN_FAMILY_TEXT_DISPLAY = 2  '  Text and Display
Public Const PAN_FAMILY_SCRIPT = 3  '  Script
Public Const PAN_FAMILY_DECORATIVE = 4  '  Decorative
Public Const PAN_FAMILY_PICTORIAL = 5  '  Pictorial
Public Const PAN_SERIF_COVE = 2  '  Cove
Public Const PAN_SERIF_OBTUSE_COVE = 3  '  Obtuse Cove
Public Const PAN_SERIF_SQUARE_COVE = 4  '  Square Cove
Public Const PAN_SERIF_OBTUSE_SQUARE_COVE = 5  '  Obtuse Square Cove
Public Const PAN_SERIF_SQUARE = 6  '  Square
Public Const PAN_SERIF_THIN = 7  '  Thin
Public Const PAN_SERIF_BONE = 8  '  Bone
Public Const PAN_SERIF_EXAGGERATED = 9  '  Exaggerated
Public Const PAN_SERIF_TRIANGLE = 10  '  Triangle
Public Const PAN_SERIF_NORMAL_SANS = 11  '  Normal Sans
Public Const PAN_SERIF_OBTUSE_SANS = 12  '  Obtuse Sans
Public Const PAN_SERIF_PERP_SANS = 13  '  Prep Sans
Public Const PAN_SERIF_FLARED = 14  '  Flared
Public Const PAN_SERIF_ROUNDED = 15  '  Rounded
Public Const PAN_WEIGHT_VERY_LIGHT = 2  '  Very Light
Public Const PAN_WEIGHT_LIGHT = 3  '  Light
Public Const PAN_WEIGHT_THIN = 4  '  Thin
Public Const PAN_WEIGHT_BOOK = 5  '  Book
Public Const PAN_WEIGHT_MEDIUM = 6  '  Medium
Public Const PAN_WEIGHT_DEMI = 7  '  Demi
Public Const PAN_WEIGHT_BOLD = 8  '  Bold
Public Const PAN_WEIGHT_HEAVY = 9  '  Heavy
Public Const PAN_WEIGHT_BLACK = 10  '  Black
Public Const PAN_WEIGHT_NORD = 11  '  Nord
Public Const PAN_PROP_OLD_STYLE = 2  '  Old Style
Public Const PAN_PROP_MODERN = 3  '  Modern
Public Const PAN_PROP_EVEN_WIDTH = 4  '  Even Width
Public Const PAN_PROP_EXPANDED = 5  '  Expanded
Public Const PAN_PROP_CONDENSED = 6  '  Condensed
Public Const PAN_PROP_VERY_EXPANDED = 7  '  Very Expanded
Public Const PAN_PROP_VERY_CONDENSED = 8  '  Very Condensed
Public Const PAN_PROP_MONOSPACED = 9  '  Monospaced
Public Const PAN_CONTRAST_NONE = 2  '  None
Public Const PAN_CONTRAST_VERY_LOW = 3  '  Very Low
Public Const PAN_CONTRAST_LOW = 4  '  Low
Public Const PAN_CONTRAST_MEDIUM_LOW = 5  '  Medium Low
Public Const PAN_CONTRAST_MEDIUM = 6  '  Medium
Public Const PAN_CONTRAST_MEDIUM_HIGH = 7  '  Mediim High
Public Const PAN_CONTRAST_HIGH = 8  '  High
Public Const PAN_CONTRAST_VERY_HIGH = 9  '  Very High
Public Const PAN_STROKE_GRADUAL_DIAG = 2  '  Gradual/Diagonal
Public Const PAN_STROKE_GRADUAL_TRAN = 3  '  Gradual/Transitional
Public Const PAN_STROKE_GRADUAL_VERT = 4  '  Gradual/Vertical
Public Const PAN_STROKE_GRADUAL_HORZ = 5  '  Gradual/Horizontal
Public Const PAN_STROKE_RAPID_VERT = 6  '  Rapid/Vertical
Public Const PAN_STROKE_RAPID_HORZ = 7  '  Rapid/Horizontal
Public Const PAN_STROKE_INSTANT_VERT = 8  '  Instant/Vertical
Public Const PAN_STRAIGHT_ARMS_HORZ = 2  '  Straight Arms/Horizontal
Public Const PAN_STRAIGHT_ARMS_WEDGE = 3  '  Straight Arms/Wedge
Public Const PAN_STRAIGHT_ARMS_VERT = 4  '  Straight Arms/Vertical
Public Const PAN_STRAIGHT_ARMS_SINGLE_SERIF = 5 '  Straight Arms/Single-Serif
Public Const PAN_STRAIGHT_ARMS_DOUBLE_SERIF = 6 '  Straight Arms/Double-Serif
Public Const PAN_BENT_ARMS_HORZ = 7  '  Non-Straight Arms/Horizontal
Public Const PAN_BENT_ARMS_WEDGE = 8  '  Non-Straight Arms/Wedge
Public Const PAN_BENT_ARMS_VERT = 9  '  Non-Straight Arms/Vertical
Public Const PAN_BENT_ARMS_SINGLE_SERIF = 10  '  Non-Straight Arms/Single-Serif
Public Const PAN_BENT_ARMS_DOUBLE_SERIF = 11  '  Non-Straight Arms/Double-Serif
Public Const PAN_LETT_NORMAL_CONTACT = 2  '  Normal/Contact
Public Const PAN_LETT_NORMAL_WEIGHTED = 3  '  Normal/Weighted
Public Const PAN_LETT_NORMAL_BOXED = 4  '  Normal/Boxed
Public Const PAN_LETT_NORMAL_FLATTENED = 5  '  Normal/Flattened
Public Const PAN_LETT_NORMAL_ROUNDED = 6  '  Normal/Rounded
Public Const PAN_LETT_NORMAL_OFF_CENTER = 7  '  Normal/Off Center
Public Const PAN_LETT_NORMAL_SQUARE = 8  '  Normal/Square
Public Const PAN_LETT_OBLIQUE_CONTACT = 9  '  Oblique/Contact
Public Const PAN_LETT_OBLIQUE_WEIGHTED = 10  '  Oblique/Weighted
Public Const PAN_LETT_OBLIQUE_BOXED = 11  '  Oblique/Boxed
Public Const PAN_LETT_OBLIQUE_FLATTENED = 12  '  Oblique/Flattened
Public Const PAN_LETT_OBLIQUE_ROUNDED = 13  '  Oblique/Rounded
Public Const PAN_LETT_OBLIQUE_OFF_CENTER = 14  '  Oblique/Off Center
Public Const PAN_LETT_OBLIQUE_SQUARE = 15  '  Oblique/Square
Public Const PAN_MIDLINE_STANDARD_TRIMMED = 2  '  Standard/Trimmed
Public Const PAN_MIDLINE_STANDARD_POINTED = 3  '  Standard/Pointed
Public Const PAN_MIDLINE_STANDARD_SERIFED = 4  '  Standard/Serifed
Public Const PAN_MIDLINE_HIGH_TRIMMED = 5  '  High/Trimmed
Public Const PAN_MIDLINE_HIGH_POINTED = 6  '  High/Pointed
Public Const PAN_MIDLINE_HIGH_SERIFED = 7  '  High/Serifed
Public Const PAN_MIDLINE_CONSTANT_TRIMMED = 8  '  Constant/Trimmed
Public Const PAN_MIDLINE_CONSTANT_POINTED = 9  '  Constant/Pointed
Public Const PAN_MIDLINE_CONSTANT_SERIFED = 10  '  Constant/Serifed
Public Const PAN_MIDLINE_LOW_TRIMMED = 11  '  Low/Trimmed
Public Const PAN_MIDLINE_LOW_POINTED = 12  '  Low/Pointed
Public Const PAN_MIDLINE_LOW_SERIFED = 13  '  Low/Serifed
Public Const PAN_XHEIGHT_CONSTANT_SMALL = 2  '  Constant/Small
Public Const PAN_XHEIGHT_CONSTANT_STD = 3  '  Constant/Standard
Public Const PAN_XHEIGHT_CONSTANT_LARGE = 4  '  Constant/Large
Public Const PAN_XHEIGHT_DUCKING_SMALL = 5  '  Ducking/Small
Public Const PAN_XHEIGHT_DUCKING_STD = 6  '  Ducking/Standard
Public Const PAN_XHEIGHT_DUCKING_LARGE = 7  '  Ducking/Large
Public Const ELF_VENDOR_SIZE = 4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0
'  EnumFonts Masks

Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4
' palette entry flags

Public Const PC_RESERVED = &H1  '  palette index used for animation
Public Const PC_EXPLICIT = &H2  '  palette index is explicit to device
Public Const PC_NOCOLLAPSE = &H4        '  do not match color to system palette
' Background Modes

Public Const TRANSPARENT = 1
Public Const OPAQUE = 2
Public Const BKMODE_LAST = 2
'  Graphics Modes

Public Const GM_COMPATIBLE = 1
Public Const GM_ADVANCED = 2
Public Const GM_LAST = 2
'  PolyDraw and GetPath point types

Public Const PT_CLOSEFIGURE = &H1
Public Const PT_LINETO = &H2
Public Const PT_BEZIERTO = &H4
Public Const PT_MOVETO = &H6
'  Mapping Modes

Public Const MM_TEXT = 1
Public Const MM_LOMETRIC = 2
Public Const MM_HIMETRIC = 3
Public Const MM_LOENGLISH = 4
Public Const MM_HIENGLISH = 5
Public Const MM_TWIPS = 6
Public Const MM_ISOTROPIC = 7
Public Const MM_ANISOTROPIC = 8
'  Min and Max Mapping Mode values

Public Const MM_MIN = MM_TEXT
Public Const MM_MAX = MM_ANISOTROPIC
Public Const MM_MAX_FIXEDSCALE = MM_TWIPS
' Coordinate Modes

Public Const ABSOLUTE = 1
Public Const RELATIVE = 2
' Stock Logical Objects

Public Const WHITE_BRUSH = 0
Public Const LTGRAY_BRUSH = 1
Public Const GRAY_BRUSH = 2
Public Const DKGRAY_BRUSH = 3
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH
Public Const WHITE_PEN = 6
Public Const BLACK_PEN = 7
Public Const NULL_PEN = 8
Public Const OEM_FIXED_FONT = 10
Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12
Public Const SYSTEM_FONT = 13
Public Const DEVICE_DEFAULT_FONT = 14
Public Const DEFAULT_PALETTE = 15
Public Const SYSTEM_FIXED_FONT = 16
Public Const STOCK_LAST = 16
Public Const CLR_INVALID = &HFFFF
' Brush Styles

Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8
'  Hatch Styles

Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL1 = 6
Public Const HS_BDIAGONAL1 = 7
Public Const HS_SOLID = 8
Public Const HS_DENSE1 = 9
Public Const HS_DENSE2 = 10
Public Const HS_DENSE3 = 11
Public Const HS_DENSE4 = 12
Public Const HS_DENSE5 = 13
Public Const HS_DENSE6 = 14
Public Const HS_DENSE7 = 15
Public Const HS_DENSE8 = 16
Public Const HS_NOSHADE = 17
Public Const HS_HALFTONE = 18
Public Const HS_SOLIDCLR = 19
Public Const HS_DITHEREDCLR = 20
Public Const HS_SOLIDTEXTCLR = 21
Public Const HS_DITHEREDTEXTCLR = 22
Public Const HS_SOLIDBKCLR = 23
Public Const HS_DITHEREDBKCLR = 24
Public Const HS_API_MAX = 25
'  Pen Styles

Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const PS_USERSTYLE = 7
Public Const PS_ALTERNATE = 8
Public Const PS_STYLE_MASK = &HF
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00
Public Const PS_JOIN_ROUND = &H0
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_MASK = &HF000
Public Const PS_COSMETIC = &H0
Public Const PS_GEOMETRIC = &H10000
Public Const PS_TYPE_MASK = &HF0000
Public Const AD_COUNTERCLOCKWISE = 1
Public Const AD_CLOCKWISE = 2
'  Device Parameters for GetDeviceCaps()

Public Const DRIVERVERSION = 0      '  Device driver version
Public Const TECHNOLOGY = 2         '  Device classification
Public Const HORZSIZE = 4           '  Horizontal size in millimeters
Public Const VERTSIZE = 6           '  Vertical size in millimeters
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const BITSPIXEL = 12         '  Number of bits per pixel
Public Const PLANES = 14            '  Number of planes
Public Const NUMBRUSHES = 16        '  Number of brushes the device has
Public Const NUMPENS = 18           '  Number of pens the device has
Public Const NUMMARKERS = 20        '  Number of markers the device has
Public Const NUMFONTS = 22          '  Number of fonts the device has
Public Const NUMCOLORS = 24         '  Number of colors the device supports
Public Const PDEVICESIZE = 26       '  Size required for device descriptor
Public Const CURVECAPS = 28         '  Curve capabilities
Public Const LINECAPS = 30          '  Line capabilities
Public Const POLYGONALCAPS = 32     '  Polygonal capabilities
Public Const TEXTCAPS = 34          '  Text capabilities
Public Const CLIPCAPS = 36          '  Clipping capabilities
Public Const RASTERCAPS = 38        '  Bitblt capabilities
Public Const ASPECTX = 40           '  Length of the X leg
Public Const ASPECTY = 42           '  Length of the Y leg
Public Const ASPECTXY = 44          '  Length of the hypotenuse
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Const SIZEPALETTE = 104      '  Number of entries in physical palette
Public Const NUMRESERVED = 106      '  Number of reserved entries in palette
Public Const COLORRES = 108         '  Actual color resolution
'  Printing related DeviceCaps. These replace the appropriate Escapes

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Public Const SCALINGFACTORX = 114 '  Scaling factor x
Public Const SCALINGFACTORY = 115 '  Scaling factor y
'  Device Capability Masks:
'  Device Technologies

Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_DISPFILE = 6            '  Display-file
'  Curve Capabilities

Public Const CC_NONE = 0                '  Curves not supported
Public Const CC_CIRCLES = 1             '  Can do circles
Public Const CC_PIE = 2                 '  Can do pie wedges
Public Const CC_CHORD = 4               '  Can do chord arcs
Public Const CC_ELLIPSES = 8            '  Can do ellipese
Public Const CC_WIDE = 16               '  Can do wide lines
Public Const CC_STYLED = 32             '  Can do styled lines
Public Const CC_WIDESTYLED = 64         '  Can do wide styled lines
Public Const CC_INTERIORS = 128 '  Can do interiors
Public Const CC_ROUNDRECT = 256 '
'  Line Capabilities

Public Const LC_NONE = 0                '  Lines not supported
Public Const LC_POLYLINE = 2            '  Can do polylines
Public Const LC_MARKER = 4              '  Can do markers
Public Const LC_POLYMARKER = 8          '  Can do polymarkers
Public Const LC_WIDE = 16               '  Can do wide lines
Public Const LC_STYLED = 32             '  Can do styled lines
Public Const LC_WIDESTYLED = 64         '  Can do wide styled lines
Public Const LC_INTERIORS = 128 '  Can do interiors
'  Polygonal Capabilities

Public Const PC_NONE = 0                '  Polygonals not supported
Public Const PC_POLYGON = 1             '  Can do polygons
Public Const PC_RECTANGLE = 2           '  Can do rectangles
Public Const PC_WINDPOLYGON = 4         '  Can do winding polygons
Public Const PC_TRAPEZOID = 4           '  Can do trapezoids
Public Const PC_SCANLINE = 8            '  Can do scanlines
Public Const PC_WIDE = 16               '  Can do wide borders
Public Const PC_STYLED = 32             '  Can do styled borders
Public Const PC_WIDESTYLED = 64         '  Can do wide styled borders
Public Const PC_INTERIORS = 128 '  Can do interiors
'  Polygonal Capabilities

Public Const CP_NONE = 0                '  No clipping of output
Public Const CP_RECTANGLE = 1           '  Output clipped to rects
Public Const CP_REGION = 2              '
'  Text Capabilities

Public Const TC_OP_CHARACTER = &H1              '  Can do OutputPrecision   CHARACTER
Public Const TC_OP_STROKE = &H2                 '  Can do OutputPrecision   STROKE
Public Const TC_CP_STROKE = &H4                 '  Can do ClipPrecision     STROKE
Public Const TC_CR_90 = &H8                     '  Can do CharRotAbility    90
Public Const TC_CR_ANY = &H10                   '  Can do CharRotAbility    ANY
Public Const TC_SF_X_YINDEP = &H20              '  Can do ScaleFreedom      X_YINDEPENDENT
Public Const TC_SA_DOUBLE = &H40                '  Can do ScaleAbility      DOUBLE
Public Const TC_SA_INTEGER = &H80               '  Can do ScaleAbility      INTEGER
Public Const TC_SA_CONTIN = &H100               '  Can do ScaleAbility      CONTINUOUS
Public Const TC_EA_DOUBLE = &H200               '  Can do EmboldenAbility   DOUBLE
Public Const TC_IA_ABLE = &H400                 '  Can do ItalisizeAbility  ABLE
Public Const TC_UA_ABLE = &H800                 '  Can do UnderlineAbility  ABLE
Public Const TC_SO_ABLE = &H1000                '  Can do StrikeOutAbility  ABLE
Public Const TC_RA_ABLE = &H2000                '  Can do RasterFontAble    ABLE
Public Const TC_VA_ABLE = &H4000                '  Can do VectorFontAble    ABLE
Public Const TC_RESERVED = &H8000
Public Const TC_SCROLLBLT = &H10000             '  do text scroll with blt
'  Raster Capabilities

Public Const RC_NONE = 0
Public Const RC_BITBLT = 1                  '  Can do standard BLT.
Public Const RC_BANDING = 2                 '  Device requires banding support
Public Const RC_SCALING = 4                 '  Device requires scaling support
Public Const RC_BITMAP64 = 8                '  Device can support >64K bitmap
Public Const RC_GDI20_OUTPUT = &H10             '  has 2.0 output calls
Public Const RC_GDI20_STATE = &H20
Public Const RC_SAVEBITMAP = &H40
Public Const RC_DI_BITMAP = &H80                '  supports DIB to memory
Public Const RC_PALETTE = &H100                 '  supports a palette
Public Const RC_DIBTODEV = &H200                '  supports DIBitsToDevice
Public Const RC_BIGFONT = &H400                 '  supports >64K fonts
Public Const RC_STRETCHBLT = &H800              '  supports StretchBlt
Public Const RC_FLOODFILL = &H1000              '  supports FloodFill
Public Const RC_STRETCHDIB = &H2000             '  supports StretchDIBits
Public Const RC_OP_DX_OUTPUT = &H4000
Public Const RC_DEVBITS = &H8000
' DIB color table identifiers

Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_PAL_INDICES = 2 '  No color table indices into surf palette
Public Const DIB_PAL_PHYSINDICES = 2 '  No color table indices into surf palette
Public Const DIB_PAL_LOGINDICES = 4 '  No color table indices into DC palette
' constants for Get/SetSystemPaletteUse()

Public Const SYSPAL_ERROR = 0
Public Const SYSPAL_STATIC = 1
Public Const SYSPAL_NOSTATIC = 2
' constants for CreateDIBitmap

Public Const CBM_CREATEDIB = &H2      '  create DIB bitmap
Public Const CBM_INIT = &H4           '  initialize bitmap
' ExtFloodFill style flags

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1
'  size of a device name string

Public Const CCHDEVICENAME = 32
'  size of a form name string

Public Const CCHFORMNAME = 32
' current version of specification

Public Const DM_SPECVERSION = &H320
' field selection bits

Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&
Public Const DM_YRESOLUTION = &H2000&
Public Const DM_TTOPTION = &H4000&
Public Const DM_COLLATE As Long = &H8000
Public Const DM_FORMNAME As Long = &H10000
'  orientation selections

Public Const DMORIENT_PORTRAIT = 1
Public Const DMORIENT_LANDSCAPE = 2
'  paper selections

Public Const DMPAPER_LETTER = 1
Public Const DMPAPER_FIRST = DMPAPER_LETTER
               '  Letter 8 1/2 x 11 in

Public Const DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
Public Const DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
Public Const DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
Public Const DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
Public Const DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
Public Const DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
Public Const DMPAPER_A3 = 8                     '  A3 297 x 420 mm
Public Const DMPAPER_A4 = 9                     '  A4 210 x 297 mm
Public Const DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
Public Const DMPAPER_A5 = 11                    '  A5 148 x 210 mm
Public Const DMPAPER_B4 = 12                    '  B4 250 x 354
Public Const DMPAPER_B5 = 13                    '  B5 182 x 257 mm
Public Const DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
Public Const DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
Public Const DMPAPER_10X14 = 16                 '  10x14 in
Public Const DMPAPER_11X17 = 17                 '  11x17 in
Public Const DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
Public Const DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
Public Const DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
Public Const DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
Public Const DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
Public Const DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
Public Const DMPAPER_CSHEET = 24                '  C size sheet
Public Const DMPAPER_DSHEET = 25                '  D size sheet
Public Const DMPAPER_ESHEET = 26                '  E size sheet
Public Const DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
Public Const DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
Public Const DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
Public Const DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
Public Const DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
Public Const DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
Public Const DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
Public Const DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
Public Const DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
Public Const DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
Public Const DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
Public Const DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
Public Const DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in
Public Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN
Public Const DMPAPER_USER = 256
'  bin selections

Public Const DMBIN_UPPER = 1
Public Const DMBIN_FIRST = DMBIN_UPPER
Public Const DMBIN_ONLYONE = 1
Public Const DMBIN_LOWER = 2
Public Const DMBIN_MIDDLE = 3
Public Const DMBIN_MANUAL = 4
Public Const DMBIN_ENVELOPE = 5
Public Const DMBIN_ENVMANUAL = 6
Public Const DMBIN_AUTO = 7
Public Const DMBIN_TRACTOR = 8
Public Const DMBIN_SMALLFMT = 9
Public Const DMBIN_LARGEFMT = 10
Public Const DMBIN_LARGECAPACITY = 11
Public Const DMBIN_CASSETTE = 14
Public Const DMBIN_LAST = DMBIN_CASSETTE
Public Const DMBIN_USER = 256               '  device specific bins start here
'  print qualities

Public Const DMRES_DRAFT = (-1)
Public Const DMRES_LOW = (-2)
Public Const DMRES_MEDIUM = (-3)
Public Const DMRES_HIGH = (-4)
'  color enable/disable for color printers

Public Const DMCOLOR_MONOCHROME = 1
Public Const DMCOLOR_COLOR = 2
'  duplex enable

Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_VERTICAL = 2
Public Const DMDUP_HORIZONTAL = 3
'  TrueType options

Public Const DMTT_BITMAP = 1            '  print TT fonts as graphics
Public Const DMTT_DOWNLOAD = 2          '  download TT fonts as soft fonts
Public Const DMTT_SUBDEV = 3            '  substitute device fonts for TT fonts
'  Collation selections

Public Const DMCOLLATE_FALSE = 0
Public Const DMCOLLATE_TRUE = 1
'  DEVMODE dmDisplayFlags flags

Public Const DM_GRAYSCALE = &H1
Public Const DM_INTERLACED = &H2
'  GetRegionData/ExtCreateRegion

Public Const RDH_RECTANGLES = 1
' GetGlyphOutline constants

Public Const GGO_METRICS = 0
Public Const GGO_BITMAP = 1
Public Const GGO_NATIVE = 2
Public Const TT_POLYGON_TYPE = 24
Public Const TT_PRIM_LINE = 1
Public Const TT_PRIM_QSPLINE = 2
' bits defined in wFlags of RASTERIZER_STATUS

Public Const TT_AVAILABLE = &H1
Public Const TT_ENABLED = &H2
'  mode selections for the device mode function

Public Const DM_UPDATE = 1
Public Const DM_COPY = 2
Public Const DM_PROMPT = 4
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_OUT_DEFAULT = DM_UPDATE
'  device capabilities indices

Public Const DC_FIELDS = 1
Public Const DC_PAPERS = 2
Public Const DC_PAPERSIZE = 3
Public Const DC_MINEXTENT = 4
Public Const DC_MAXEXTENT = 5
Public Const DC_BINS = 6
Public Const DC_DUPLEX = 7
Public Const DC_SIZE = 8
Public Const DC_EXTRA = 9
Public Const DC_VERSION = 10
Public Const DC_DRIVER = 11
Public Const DC_BINNAMES = 12
Public Const DC_ENUMRESOLUTIONS = 13
Public Const DC_FILEDEPENDENCIES = 14
Public Const DC_TRUETYPE = 15
Public Const DC_PAPERNAMES = 16
Public Const DC_ORIENTATION = 17
Public Const DC_COPIES = 18
'  bit fields of the return value (DWORD) for DC_TRUETYPE

Public Const DCTT_BITMAP = &H1&
Public Const DCTT_DOWNLOAD = &H2&
Public Const DCTT_SUBDEV = &H4&
'  Flags value for COLORADJUSTMENT

Public Const CA_NEGATIVE = &H1
Public Const CA_LOG_FILTER = &H2
'  IlluminantIndex values

Public Const ILLUMINANT_DEVICE_DEFAULT = 0
Public Const ILLUMINANT_A = 1
Public Const ILLUMINANT_B = 2
Public Const ILLUMINANT_C = 3
Public Const ILLUMINANT_D50 = 4
Public Const ILLUMINANT_D55 = 5
Public Const ILLUMINANT_D65 = 6
Public Const ILLUMINANT_D75 = 7
Public Const ILLUMINANT_F2 = 8
Public Const ILLUMINANT_MAX_INDEX = ILLUMINANT_F2
Public Const ILLUMINANT_TUNGSTEN = ILLUMINANT_A
Public Const ILLUMINANT_DAYLIGHT = ILLUMINANT_C
Public Const ILLUMINANT_FLUORESCENT = ILLUMINANT_F2
Public Const ILLUMINANT_NTSC = ILLUMINANT_C
'  Min and max for RedGamma, GreenGamma, BlueGamma

Public Const RGB_GAMMA_MIN = 2500 'words
Public Const RGB_GAMMA_MAX = 65000
'  Min and max for ReferenceBlack and ReferenceWhite

Public Const REFERENCE_WHITE_MIN = 6000 'words
Public Const REFERENCE_WHITE_MAX = 10000
Public Const REFERENCE_BLACK_MIN = 0
Public Const REFERENCE_BLACK_MAX = 4000
'  Min and max for Contrast, Brightness, Colorfulness, RedGreenTint

Public Const COLOR_ADJ_MIN = -100 'shorts
Public Const COLOR_ADJ_MAX = 100
Public Const FONTMAPPER_MAX = 10
' Enhanced metafile constants

Public Const ENHMETA_SIGNATURE = &H464D4520
'  Stock object flag used in the object handle
' index in the enhanced metafile records.
'  E.g. The object handle index (META_STOCK_OBJECT Or BLACK_BRUSH)
'  represents the stock object BLACK_BRUSH.

Public Const ENHMETA_STOCK_OBJECT = &H80000000
'  Enhanced metafile record types.

Public Const EMR_HEADER = 1
Public Const EMR_POLYBEZIER = 2
Public Const EMR_POLYGON = 3
Public Const EMR_POLYLINE = 4
Public Const EMR_POLYBEZIERTO = 5
Public Const EMR_POLYLINETO = 6
Public Const EMR_POLYPOLYLINE = 7
Public Const EMR_POLYPOLYGON = 8
Public Const EMR_SETWINDOWEXTEX = 9
Public Const EMR_SETWINDOWORGEX = 10
Public Const EMR_SETVIEWPORTEXTEX = 11
Public Const EMR_SETVIEWPORTORGEX = 12
Public Const EMR_SETBRUSHORGEX = 13
Public Const EMR_EOF = 14
Public Const EMR_SETPIXELV = 15
Public Const EMR_SETMAPPERFLAGS = 16
Public Const EMR_SETMAPMODE = 17
Public Const EMR_SETBKMODE = 18
Public Const EMR_SETPOLYFILLMODE = 19
Public Const EMR_SETROP2 = 20
Public Const EMR_SETSTRETCHBLTMODE = 21
Public Const EMR_SETTEXTALIGN = 22
Public Const EMR_SETCOLORADJUSTMENT = 23
Public Const EMR_SETTEXTCOLOR = 24
Public Const EMR_SETBKCOLOR = 25
Public Const EMR_OFFSETCLIPRGN = 26
Public Const EMR_MOVETOEX = 27
Public Const EMR_SETMETARGN = 28
Public Const EMR_EXCLUDECLIPRECT = 29
Public Const EMR_INTERSECTCLIPRECT = 30
Public Const EMR_SCALEVIEWPORTEXTEX = 31
Public Const EMR_SCALEWINDOWEXTEX = 32
Public Const EMR_SAVEDC = 33
Public Const EMR_RESTOREDC = 34
Public Const EMR_SETWORLDTRANSFORM = 35
Public Const EMR_MODIFYWORLDTRANSFORM = 36
Public Const EMR_SELECTOBJECT = 37
Public Const EMR_CREATEPEN = 38
Public Const EMR_CREATEBRUSHINDIRECT = 39
Public Const EMR_DELETEOBJECT = 40
Public Const EMR_ANGLEARC = 41
Public Const EMR_ELLIPSE = 42
Public Const EMR_RECTANGLE = 43
Public Const EMR_ROUNDRECT = 44
Public Const EMR_ARC = 45
Public Const EMR_CHORD = 46
Public Const EMR_PIE = 47
Public Const EMR_SELECTPALETTE = 48
Public Const EMR_CREATEPALETTE = 49
Public Const EMR_SETPALETTEENTRIES = 50
Public Const EMR_RESIZEPALETTE = 51
Public Const EMR_REALIZEPALETTE = 52
Public Const EMR_EXTFLOODFILL = 53
Public Const EMR_LINETO = 54
Public Const EMR_ARCTO = 55
Public Const EMR_POLYDRAW = 56
Public Const EMR_SETARCDIRECTION = 57
Public Const EMR_SETMITERLIMIT = 58
Public Const EMR_BEGINPATH = 59
Public Const EMR_ENDPATH = 60
Public Const EMR_CLOSEFIGURE = 61
Public Const EMR_FILLPATH = 62
Public Const EMR_STROKEANDFILLPATH = 63
Public Const EMR_STROKEPATH = 64
Public Const EMR_FLATTENPATH = 65
Public Const EMR_WIDENPATH = 66
Public Const EMR_SELECTCLIPPATH = 67
Public Const EMR_ABORTPATH = 68
Public Const EMR_GDICOMMENT = 70
Public Const EMR_FILLRGN = 71
Public Const EMR_FRAMERGN = 72
Public Const EMR_INVERTRGN = 73
Public Const EMR_PAINTRGN = 74
Public Const EMR_EXTSELECTCLIPRGN = 75
Public Const EMR_BITBLT = 76
Public Const EMR_STRETCHBLT = 77
Public Const EMR_MASKBLT = 78
Public Const EMR_PLGBLT = 79
Public Const EMR_SETDIBITSTODEVICE = 80
Public Const EMR_STRETCHDIBITS = 81
Public Const EMR_EXTCREATEFONTINDIRECTW = 82
Public Const EMR_EXTTEXTOUTA = 83
Public Const EMR_EXTTEXTOUTW = 84
Public Const EMR_POLYBEZIER16 = 85
Public Const EMR_POLYGON16 = 86
Public Const EMR_POLYLINE16 = 87
Public Const EMR_POLYBEZIERTO16 = 88
Public Const EMR_POLYLINETO16 = 89
Public Const EMR_POLYPOLYLINE16 = 90
Public Const EMR_POLYPOLYGON16 = 91
Public Const EMR_POLYDRAW16 = 92
Public Const EMR_CREATEMONOBRUSH = 93
Public Const EMR_CREATEDIBPATTERNBRUSHPT = 94
Public Const EMR_EXTCREATEPEN = 95
Public Const EMR_POLYTEXTOUTA = 96
Public Const EMR_POLYTEXTOUTW = 97
Public Const EMR_MIN = 1
Public Const EMR_MAX = 97
' new wingdi
' *************************************************************************
' *                                                                         *
' * wingdi.h -- GDI procedure declarations, constant definitions and macros *
' *                                                                         *
' * Copyright (c) 1985-1995, Microsoft Corp. All rights reserved.           *
' *                                                                         *
' **************************************************************************/
'  StretchBlt() Modes

Public Const STRETCH_ANDSCANS = 1
Public Const STRETCH_ORSCANS = 2
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_HALFTONE = 4
Public Const TCI_SRCCHARSET = 1
Public Const TCI_SRCCODEPAGE = 2
Public Const TCI_SRCFONTSIG = 3
Public Const MONO_FONT = 8
Public Const JOHAB_CHARSET = 130
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const THAI_CHARSET = 222
Public Const EASTEUROPE_CHARSET = 238
Public Const RUSSIAN_CHARSET = 204
Public Const MAC_CHARSET = 77
Public Const BALTIC_CHARSET = 186
Public Const FS_LATIN1 = &H1&
Public Const FS_LATIN2 = &H2&
Public Const FS_CYRILLIC = &H4&
Public Const FS_GREEK = &H8&
Public Const FS_TURKISH = &H10&
Public Const FS_HEBREW = &H20&
Public Const FS_ARABIC = &H40&
Public Const FS_BALTIC = &H80&
Public Const FS_THAI = &H10000
Public Const FS_JISJAPAN = &H20000
Public Const FS_CHINESESIMP = &H40000
Public Const FS_WANSUNG = &H80000
Public Const FS_CHINESETRAD = &H100000
Public Const FS_JOHAB = &H200000
Public Const FS_SYMBOL = &H80000000
Public Const DEFAULT_GUI_FONT = 17
'  current version of specification

Public Const DM_RESERVED1 = &H800000
Public Const DM_RESERVED2 = &H1000000
Public Const DM_ICMMETHOD = &H2000000
Public Const DM_ICMINTENT = &H4000000
Public Const DM_MEDIATYPE = &H8000000
Public Const DM_DITHERTYPE = &H10000000
Public Const DMPAPER_ISO_B4 = 42                '  B4 (ISO) 250 x 353 mm
Public Const DMPAPER_JAPANESE_POSTCARD = 43     '  Japanese Postcard 100 x 148 mm
Public Const DMPAPER_9X11 = 44                  '  9 x 11 in
Public Const DMPAPER_10X11 = 45                 '  10 x 11 in
Public Const DMPAPER_15X11 = 46                 '  15 x 11 in
Public Const DMPAPER_ENV_INVITE = 47            '  Envelope Invite 220 x 220 mm
Public Const DMPAPER_RESERVED_48 = 48           '  RESERVED--DO NOT USE
Public Const DMPAPER_RESERVED_49 = 49           '  RESERVED--DO NOT USE
Public Const DMPAPER_LETTER_EXTRA = 50              '  Letter Extra 9 \275 x 12 in
Public Const DMPAPER_LEGAL_EXTRA = 51               '  Legal Extra 9 \275 x 15 in
Public Const DMPAPER_TABLOID_EXTRA = 52              '  Tabloid Extra 11.69 x 18 in
Public Const DMPAPER_A4_EXTRA = 53                   '  A4 Extra 9.27 x 12.69 in
Public Const DMPAPER_LETTER_TRANSVERSE = 54     '  Letter Transverse 8 \275 x 11 in
Public Const DMPAPER_A4_TRANSVERSE = 55         '  A4 Transverse 210 x 297 mm
Public Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56 '  Letter Extra Transverse 9\275 x 12 in
Public Const DMPAPER_A_PLUS = 57                '  SuperA/SuperA/A4 227 x 356 mm
Public Const DMPAPER_B_PLUS = 58                '  SuperB/SuperB/A3 305 x 487 mm
Public Const DMPAPER_LETTER_PLUS = 59           '  Letter Plus 8.5 x 12.69 in
Public Const DMPAPER_A4_PLUS = 60               '  A4 Plus 210 x 330 mm
Public Const DMPAPER_A5_TRANSVERSE = 61         '  A5 Transverse 148 x 210 mm
Public Const DMPAPER_B5_TRANSVERSE = 62         '  B5 (JIS) Transverse 182 x 257 mm
Public Const DMPAPER_A3_EXTRA = 63              '  A3 Extra 322 x 445 mm
Public Const DMPAPER_A5_EXTRA = 64              '  A5 Extra 174 x 235 mm
Public Const DMPAPER_B5_EXTRA = 65              '  B5 (ISO) Extra 201 x 276 mm
Public Const DMPAPER_A2 = 66                    '  A2 420 x 594 mm
Public Const DMPAPER_A3_TRANSVERSE = 67         '  A3 Transverse 297 x 420 mm
Public Const DMPAPER_A3_EXTRA_TRANSVERSE = 68   '  A3 Extra Transverse 322 x 445 mm
Public Const DMTT_DOWNLOAD_OUTLINE = 4 '  download TT fonts as outline soft fonts
'  ICM methods

Public Const DMICMMETHOD_NONE = 1       '  ICM disabled
Public Const DMICMMETHOD_SYSTEM = 2     '  ICM handled by system
Public Const DMICMMETHOD_DRIVER = 3     '  ICM handled by driver
Public Const DMICMMETHOD_DEVICE = 4     '  ICM handled by device
Public Const DMICMMETHOD_USER = 256     '  Device-specific methods start here
'  ICM Intents

Public Const DMICM_SATURATE = 1         '  Maximize color saturation
Public Const DMICM_CONTRAST = 2         '  Maximize color contrast
Public Const DMICM_COLORMETRIC = 3      '  Use specific color metric
Public Const DMICM_USER = 256           '  Device-specific intents start here
'  Media types

Public Const DMMEDIA_STANDARD = 1         '  Standard paper
Public Const DMMEDIA_GLOSSY = 2           '  Glossy paper
Public Const DMMEDIA_TRANSPARENCY = 3     '  Transparency
Public Const DMMEDIA_USER = 256           '  Device-specific media start here
'  Dither types

Public Const DMDITHER_NONE = 1          '  No dithering
Public Const DMDITHER_COARSE = 2        '  Dither with a coarse brush
Public Const DMDITHER_FINE = 3          '  Dither with a fine brush
Public Const DMDITHER_LINEART = 4       '  LineArt dithering
Public Const DMDITHER_GRAYSCALE = 5     '  Device does grayscaling
Public Const DMDITHER_USER = 256        '  Device-specific dithers start here
Public Const GGO_GRAY2_BITMAP = 4
Public Const GGO_GRAY4_BITMAP = 5
Public Const GGO_GRAY8_BITMAP = 6
Public Const GGO_GLYPH_INDEX = &H80
Public Const GCP_DBCS = &H1
Public Const GCP_REORDER = &H2
Public Const GCP_USEKERNING = &H8
Public Const GCP_GLYPHSHAPE = &H10
Public Const GCP_LIGATE = &H20
Public Const GCP_DIACRITIC = &H100
Public Const GCP_KASHIDA = &H400
Public Const GCP_ERROR = &H8000
Public Const FLI_MASK = &H103B
Public Const GCP_JUSTIFY = &H10000
Public Const GCP_NODIACRITICS = &H20000
Public Const FLI_GLYPHS = &H40000
Public Const GCP_CLASSIN = &H80000
Public Const GCP_MAXEXTENT = &H100000
Public Const GCP_JUSTIFYIN = &H200000
Public Const GCP_DISPLAYZWG = &H400000
Public Const GCP_SYMSWAPOFF = &H800000
Public Const GCP_NUMERICOVERRIDE = &H1000000
Public Const GCP_NEUTRALOVERRIDE = &H2000000
Public Const GCP_NUMERICSLATIN = &H4000000
Public Const GCP_NUMERICSLOCAL = &H8000000
Public Const GCPCLASS_LATIN = 1
Public Const GCPCLASS_HEBREW = 2
Public Const GCPCLASS_ARABIC = 2
Public Const GCPCLASS_NEUTRAL = 3
Public Const GCPCLASS_LOCALNUMBER = 4
Public Const GCPCLASS_LATINNUMBER = 5
Public Const GCPCLASS_LATINNUMERICTERMINATOR = 6
Public Const GCPCLASS_LATINNUMERICSEPARATOR = 7
Public Const GCPCLASS_NUMERICSEPARATOR = 8
Public Const GCPCLASS_PREBOUNDRTL = &H80
Public Const GCPCLASS_PREBOUNDLTR = &H40
Public Const DC_BINADJUST = 19
Public Const DC_EMF_COMPLIANT = 20
Public Const DC_DATATYPE_PRODUCED = 21
Public Const DC_COLLATE = 22
Public Const DCTT_DOWNLOAD_OUTLINE = &H8&
'  return values for DC_BINADJUST

Public Const DCBA_FACEUPNONE = &H0
Public Const DCBA_FACEUPCENTER = &H1
Public Const DCBA_FACEUPLEFT = &H2
Public Const DCBA_FACEUPRIGHT = &H3
Public Const DCBA_FACEDOWNNONE = &H100
Public Const DCBA_FACEDOWNCENTER = &H101
Public Const DCBA_FACEDOWNLEFT = &H102
Public Const DCBA_FACEDOWNRIGHT = &H103
Public Const ICM_OFF = 1
Public Const ICM_ON = 2
Public Const ICM_QUERY = 3
Public Const EMR_SETICMMODE = 98
Public Const EMR_CREATECOLORSPACE = 99
Public Const EMR_SETCOLORSPACE = 100
Public Const EMR_DELETECOLORSPACE = 101
' --------------
'  USER Section
' --------------
' Scroll Bar Constants

Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3
' Scroll Bar Commands

Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1
Public Const SB_PAGEUP = 2
Public Const SB_PAGELEFT = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGERIGHT = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_LEFT = 6
Public Const SB_BOTTOM = 7
Public Const SB_RIGHT = 7
Public Const SB_ENDSCROLL = 8
' ShowWindow() Commands

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
' Old ShowWindow() Commands

Public Const HIDE_WINDOW = 0
Public Const SHOW_OPENWINDOW = 1
Public Const SHOW_ICONWINDOW = 2
Public Const SHOW_FULLSCREEN = 3
Public Const SHOW_OPENNOACTIVATE = 4
' Identifiers for the WM_SHOWWINDOW message

Public Const SW_PARENTCLOSING = 1
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTOPENING = 3
Public Const SW_OTHERUNZOOM = 4
' WM_KEYUP/DOWN/CHAR HIWORD(lParam) flags

Public Const KF_EXTENDED = &H100
Public Const KF_DLGMODE = &H800
Public Const KF_MENUMODE = &H1000
Public Const KF_ALTDOWN = &H2000
Public Const KF_REPEAT = &H4000
Public Const KF_UP = &H8000
' Virtual Keys, Standard Set

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F
' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /

Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE
' SetWindowsHook() codes

Public Const WH_MIN = (-1)
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_MOUSE = 7
Public Const WH_HARDWARE = 8
Public Const WH_DEBUG = 9
Public Const WH_SHELL = 10
Public Const WH_FOREGROUNDIDLE = 11
Public Const WH_MAX = 11
' Hook Codes

Public Const HC_ACTION = 0
Public Const HC_GETNEXT = 1
Public Const HC_SKIP = 2
Public Const HC_NOREMOVE = 3
Public Const HC_NOREM = HC_NOREMOVE
Public Const HC_SYSMODALON = 4
Public Const HC_SYSMODALOFF = 5
' CBT Hook Codes

Public Const HCBT_MOVESIZE = 0
Public Const HCBT_MINMAX = 1
Public Const HCBT_QS = 2
Public Const HCBT_CREATEWND = 3
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_ACTIVATE = 5
Public Const HCBT_CLICKSKIPPED = 6
Public Const HCBT_KEYSKIPPED = 7
Public Const HCBT_SYSCOMMAND = 8
Public Const HCBT_SETFOCUS = 9
' WH_MSGFILTER Filter Proc Codes

Public Const MSGF_DIALOGBOX = 0
Public Const MSGF_MESSAGEBOX = 1
Public Const MSGF_MENU = 2
Public Const MSGF_MOVE = 3
Public Const MSGF_SIZE = 4
Public Const MSGF_SCROLLBAR = 5
Public Const MSGF_NEXTWINDOW = 6
Public Const MSGF_MAINLOOP = 8
Public Const MSGF_MAX = 8
Public Const MSGF_USER = 4096
Public Const HSHELL_WINDOWCREATED = 1
Public Const HSHELL_WINDOWDESTROYED = 2
Public Const HSHELL_ACTIVATESHELLWINDOW = 3
' Keyboard Layout API

Public Const HKL_PREV = 0
Public Const HKL_NEXT = 1
Public Const KLF_ACTIVATE = &H1
Public Const KLF_SUBSTITUTE_OK = &H2
Public Const KLF_UNLOADPREVIOUS = &H4
Public Const KLF_REORDER = &H8
' Size of KeyboardLayoutName (number of characters), including nul terminator

Public Const KL_NAMELENGTH = 9
' Desktop-specific access flags

Public Const DESKTOP_READOBJECTS = &H1&
Public Const DESKTOP_CREATEWINDOW = &H2&
Public Const DESKTOP_CREATEMENU = &H4&
Public Const DESKTOP_HOOKCONTROL = &H8&
Public Const DESKTOP_JOURNALRECORD = &H10&
Public Const DESKTOP_JOURNALPLAYBACK = &H20&
Public Const DESKTOP_ENUMERATE = &H40&
Public Const DESKTOP_WRITEOBJECTS = &H80&
' Windowstation-specific access flags

Public Const WINSTA_ENUMDESKTOPS = &H1&
Public Const WINSTA_READATTRIBUTES = &H2&
Public Const WINSTA_ACCESSCLIPBOARD = &H4&
Public Const WINSTA_CREATEDESKTOP = &H8&
Public Const WINSTA_WRITEATTRIBUTES = &H10&
Public Const WINSTA_ACCESSPUBLICATOMS = &H20&
Public Const WINSTA_EXITWINDOWS = &H40&
Public Const WINSTA_ENUMERATE = &H100&
Public Const WINSTA_READSCREEN = &H200&
' Message structure
' Window field offsets for GetWindowLong() and GetWindowWord()

Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
' Class field offsets for GetClassLong() and GetClassWord()

Public Const GCL_MENUNAME = (-8)
Public Const GCL_HBRBACKGROUND = (-10)
Public Const GCL_HCURSOR = (-12)
Public Const GCL_HICON = (-14)
Public Const GCL_HMODULE = (-16)
Public Const GCL_CBWNDEXTRA = (-18)
Public Const GCL_CBCLSEXTRA = (-20)
Public Const GCL_WNDPROC = (-24)
Public Const GCL_STYLE = (-26)
Public Const GCW_ATOM = (-32)
' Window Messages

Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41
Public Const WM_OTHERWINDOWCREATED = &H42               '  no longer suported
Public Const WM_OTHERWINDOWDESTROYED = &H43             '  no longer suported
Public Const WM_COMMNOTIFY = &H44                       '  no longer suported
' notifications passed in low word of lParam on WM_COMMNOTIFY messages

Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48
'
'  wParam for WM_POWER window message and DRV_POWER driver notification

Public Const PWR_OK = 1
Public Const PWR_FAIL = (-1)
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_SUSPENDRESUME = 2
Public Const PWR_CRITICALRESUME = 3
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
' NOTE: All Message Numbers below 0x0400 are RESERVED.
' Private Window Messages Start Here:

Public Const WM_USER = &H400
' WM_SYNCTASK Commands

Public Const ST_BEGINSWP = 0
Public Const ST_ENDSWP = 1
' WM_NCHITTEST and MOUSEHOOKSTRUCT Mouse Position Codes

Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
'  SendMessageTimeout values

Public Const SMTO_NORMAL = &H0
Public Const SMTO_BLOCK = &H1
Public Const SMTO_ABORTIFHUNG = &H2
' WM_MOUSEACTIVATE Return Codes

Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4
' WM_SIZE message wParam values

Public Const SIZE_RESTORED = 0
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_MAXIMIZED = 2
Public Const SIZE_MAXSHOW = 3
Public Const SIZE_MAXHIDE = 4
' Obsolete constant names

Public Const SIZENORMAL = SIZE_RESTORED
Public Const SIZEICONIC = SIZE_MINIMIZED
Public Const SIZEFULLSCREEN = SIZE_MAXIMIZED
Public Const SIZEZOOMSHOW = SIZE_MAXSHOW
Public Const SIZEZOOMHIDE = SIZE_MAXHIDE
' WM_NCCALCSIZE return flags

Public Const WVR_ALIGNTOP = &H10
Public Const WVR_ALIGNLEFT = &H20
Public Const WVR_ALIGNBOTTOM = &H40
Public Const WVR_ALIGNRIGHT = &H80
Public Const WVR_HREDRAW = &H100
Public Const WVR_VREDRAW = &H200
Public Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
Public Const WVR_VALIDRECTS = &H400
' Key State Masks for Mouse Messages

Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_CONTROL = &H8
Public Const MK_MBUTTON = &H10
' Window Styles

Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
'
'   Common Window Styles
'  /

Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
' Extended Window Styles

Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
' Class styles

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000
' Predefined Clipboard Formats

Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_SYLK = 4
Public Const CF_DIF = 5
Public Const CF_TIFF = 6
Public Const CF_OEMTEXT = 7
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9
Public Const CF_PENDATA = 10
Public Const CF_RIFF = 11
Public Const CF_WAVE = 12
Public Const CF_UNICODETEXT = 13
Public Const CF_ENHMETAFILE = 14
Public Const CF_OWNERDISPLAY = &H80
Public Const CF_DSPTEXT = &H81
Public Const CF_DSPBITMAP = &H82
Public Const CF_DSPMETAFILEPICT = &H83
Public Const CF_DSPENHMETAFILE = &H8E
' "Private" formats don't get GlobalFree()'d

Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF
' "GDIOBJ" formats do get DeleteObject()'d

Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF
'  Defines for the fVirt field of the Accelerator table structure.

Public Const FVIRTKEY = True          '  Assumed to be == TRUE
Public Const FNOINVERT = &H2
Public Const FSHIFT = &H4
Public Const FCONTROL = &H8
Public Const FALT = &H10
Public Const WPF_SETMINPOSITION = &H1
Public Const WPF_RESTORETOMAXIMIZED = &H2
' Owner draw control types

Public Const ODT_MENU = 1
Public Const ODT_LISTBOX = 2
Public Const ODT_COMBOBOX = 3
Public Const ODT_BUTTON = 4
' Owner draw actions

Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_SELECT = &H2
Public Const ODA_FOCUS = &H4
' Owner draw state

Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10
' PeekMessage() Options

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const IDHOT_SNAPWINDOW = (-1)    '  SHIFT-PRINTSCRN
Public Const IDHOT_SNAPDESKTOP = (-2)    '  PRINTSCRN
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const READAPI = 0        '  Flags for _lopen
Public Const WRITEAPI = 1
Public Const READ_WRITE = 2
' Special HWND value for use with PostMessage and SendMessage

Public Const HWND_BROADCAST = &HFFFF&
Public Const CW_USEDEFAULT = &H80000000
Public Const HWND_DESKTOP = 0
' SetWindowPos Flags

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const DLGWINDOWEXTRA = 30        '  Window extra bytes needed for private dialog classes
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
' GetQueueStatus flags

Public Const QS_KEY = &H1
Public Const QS_MOUSEMOVE = &H2
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_POSTMESSAGE = &H8
Public Const QS_TIMER = &H10
Public Const QS_PAINT = &H20
Public Const QS_SENDMESSAGE = &H40
Public Const QS_HOTKEY = &H80
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
' GetSystemMetrics() codes

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CMETRICS = 44
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
' Flags for TrackPopupMenu

Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_RIGHTALIGN = &H8&
' DrawText() Format Flags

Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DCX_WINDOW = &H1&
Public Const DCX_CACHE = &H2&
Public Const DCX_NORESETATTRS = &H4&
Public Const DCX_CLIPCHILDREN = &H8&
Public Const DCX_CLIPSIBLINGS = &H10&
Public Const DCX_PARENTCLIP = &H20&
Public Const DCX_EXCLUDERGN = &H40&
Public Const DCX_INTERSECTRGN = &H80&
Public Const DCX_EXCLUDEUPDATE = &H100&
Public Const DCX_INTERSECTUPDATE = &H200&
Public Const DCX_LOCKWINDOWUPDATE = &H400&
Public Const DCX_NORECOMPUTE = &H100000
Public Const DCX_VALIDATE = &H200000
Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ERASE = &H4
Public Const RDW_VALIDATE = &H8
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_NOERASE = &H20
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_NOFRAME = &H800
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_INVALIDATE = &H2
Public Const SW_ERASE = &H4
' EnableScrollBar() flags

Public Const ESB_ENABLE_BOTH = &H0
Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_DISABLE_LEFT = &H1
Public Const ESB_DISABLE_RIGHT = &H2
Public Const ESB_DISABLE_UP = &H1
Public Const ESB_DISABLE_DOWN = &H2
Public Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Public Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT
' MessageBox() Flags

Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_YESNOCANCEL = &H3&
Public Const MB_YESNO = &H4&
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&
Public Const MB_APPLMODAL = &H0&
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_NOFOCUS = &H8000&
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_TYPEMASK = &HF&
Public Const MB_ICONMASK = &HF0&
Public Const MB_DEFMASK = &HF00&
Public Const MB_MODEMASK = &H3000&
Public Const MB_MISCMASK = &HC000&
' Color Types

Public Const CTLCOLOR_MSGBOX = 0
Public Const CTLCOLOR_EDIT = 1
Public Const CTLCOLOR_LISTBOX = 2
Public Const CTLCOLOR_BTN = 3
Public Const CTLCOLOR_DLG = 4
Public Const CTLCOLOR_SCROLLBAR = 5
Public Const CTLCOLOR_STATIC = 6
Public Const CTLCOLOR_MAX = 8   '  three bits max
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
' GetWindow() Constants

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5
' Menu flags for Add/Check/EnableMenuItem()

Public Const MF_INSERT = &H0&
Public Const MF_CHANGE = &H80&
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_SEPARATOR = &H800&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_UNHILITE = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_SYSMENU = &H2000&
Public Const MF_HELP = &H4000&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_END = &H80
' System Menu Command Values

Public Const SC_SIZE = &HF000&
Public Const SC_MOVE = &HF010&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_NEXTWINDOW = &HF040&
Public Const SC_PREVWINDOW = &HF050&
Public Const SC_CLOSE = &HF060&
Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&
Public Const SC_MOUSEMENU = &HF090&
Public Const SC_KEYMENU = &HF100&
Public Const SC_ARRANGE = &HF110&
Public Const SC_RESTORE = &HF120&
Public Const SC_TASKLIST = &HF130&
Public Const SC_SCREENSAVE = &HF140&
Public Const SC_HOTKEY = &HF150&
' Obsolete names

Public Const SC_ICON = SC_MINIMIZE
Public Const SC_ZOOM = SC_MAXIMIZE
' Standard Cursor IDs

Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
' OEM Resource Ordinal Numbers

Public Const OBM_CLOSE = 32754
Public Const OBM_UPARROW = 32753
Public Const OBM_DNARROW = 32752
Public Const OBM_RGARROW = 32751
Public Const OBM_LFARROW = 32750
Public Const OBM_REDUCE = 32749
Public Const OBM_ZOOM = 32748
Public Const OBM_RESTORE = 32747
Public Const OBM_REDUCED = 32746
Public Const OBM_ZOOMD = 32745
Public Const OBM_RESTORED = 32744
Public Const OBM_UPARROWD = 32743
Public Const OBM_DNARROWD = 32742
Public Const OBM_RGARROWD = 32741
Public Const OBM_LFARROWD = 32740
Public Const OBM_MNARROW = 32739
Public Const OBM_COMBO = 32738
Public Const OBM_UPARROWI = 32737
Public Const OBM_DNARROWI = 32736
Public Const OBM_RGARROWI = 32735
Public Const OBM_LFARROWI = 32734
Public Const OBM_OLD_CLOSE = 32767
Public Const OBM_SIZE = 32766
Public Const OBM_OLD_UPARROW = 32765
Public Const OBM_OLD_DNARROW = 32764
Public Const OBM_OLD_RGARROW = 32763
Public Const OBM_OLD_LFARROW = 32762
Public Const OBM_BTSIZE = 32761
Public Const OBM_CHECK = 32760
Public Const OBM_CHECKBOXES = 32759
Public Const OBM_BTNCORNERS = 32758
Public Const OBM_OLD_REDUCE = 32757
Public Const OBM_OLD_ZOOM = 32756
Public Const OBM_OLD_RESTORE = 32755
Public Const OCR_NORMAL = 32512
Public Const OCR_IBEAM = 32513
Public Const OCR_WAIT = 32514
Public Const OCR_CROSS = 32515
Public Const OCR_UP = 32516
Public Const OCR_SIZE = 32640
Public Const OCR_ICON = 32641
Public Const OCR_SIZENWSE = 32642
Public Const OCR_SIZENESW = 32643
Public Const OCR_SIZEWE = 32644
Public Const OCR_SIZENS = 32645
Public Const OCR_SIZEALL = 32646
Public Const OCR_ICOCUR = 32647
Public Const OCR_NO = 32648 ' not in win3.1
Public Const OIC_SAMPLE = 32512
Public Const OIC_HAND = 32513
Public Const OIC_QUES = 32514
Public Const OIC_BANG = 32515
Public Const OIC_NOTE = 32516
Public Const ORD_LANGDRIVER = 1 '  The ordinal number for the entry point of
                                '  language drivers.
' Standard Icon IDs

Public Const IDI_APPLICATION = 32512&
Public Const IDI_HAND = 32513&
Public Const IDI_QUESTION = 32514&
Public Const IDI_EXCLAMATION = 32515&
Public Const IDI_ASTERISK = 32516&
' Dialog Box Command IDs

Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
' Control Manager Structures and Definitions
' Edit Control Styles

Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&
' Edit Control Notification Codes

Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const EN_CHANGE = &H300
Public Const EN_UPDATE = &H400
Public Const EN_ERRSPACE = &H500
Public Const EN_MAXTEXT = &H501
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602
' Edit Control Messages

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETRECT = &HB2
Public Const EM_SETRECT = &HB3
Public Const EM_SETRECTNP = &HB4
Public Const EM_SCROLL = &HB5
Public Const EM_LINESCROLL = &HB6
Public Const EM_SCROLLCARET = &HB7
Public Const EM_GETMODIFY = &HB8
Public Const EM_SETMODIFY = &HB9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETHANDLE = &HBC
Public Const EM_GETHANDLE = &HBD
Public Const EM_GETTHUMB = &HBE
Public Const EM_LINELENGTH = &HC1
Public Const EM_REPLACESEL = &HC2
Public Const EM_GETLINE = &HC4
Public Const EM_LIMITTEXT = &HC5
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_FMTLINES = &HC8
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_GETPASSWORDCHAR = &HD2
' EDITWORDBREAKPROC code values

Public Const WB_LEFT = 0
Public Const WB_RIGHT = 1
Public Const WB_ISDELIMITER = 2
' Button Control Styles

Public Const BS_PUSHBUTTON = &H0&
Public Const BS_DEFPUSHBUTTON = &H1&
Public Const BS_CHECKBOX = &H2&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BS_RADIOBUTTON = &H4&
Public Const BS_3STATE = &H5&
Public Const BS_AUTO3STATE = &H6&
Public Const BS_GROUPBOX = &H7&
Public Const BS_USERBUTTON = &H8&
Public Const BS_AUTORADIOBUTTON = &H9&
Public Const BS_OWNERDRAW = &HB&
Public Const BS_LEFTTEXT = &H20&
' User Button Notification Codes

Public Const BN_CLICKED = 0
Public Const BN_PAINT = 1
Public Const BN_HILITE = 2
Public Const BN_UNHILITE = 3
Public Const BN_DISABLE = 4
Public Const BN_DOUBLECLICKED = 5
' Button Control Messages

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const BM_GETSTATE = &HF2
Public Const BM_SETSTATE = &HF3
Public Const BM_SETSTYLE = &HF4
' Static Control Constants

Public Const SS_LEFT = &H0&
Public Const SS_CENTER = &H1&
Public Const SS_RIGHT = &H2&
Public Const SS_ICON = &H3&
Public Const SS_BLACKRECT = &H4&
Public Const SS_GRAYRECT = &H5&
Public Const SS_WHITERECT = &H6&
Public Const SS_BLACKFRAME = &H7&
Public Const SS_GRAYFRAME = &H8&
Public Const SS_WHITEFRAME = &H9&
Public Const SS_USERITEM = &HA&
Public Const SS_SIMPLE = &HB&
Public Const SS_LEFTNOWORDWRAP = &HC&
Public Const SS_NOPREFIX = &H80           '  Don't do "&" character translation
' Static Control Mesages

Public Const STM_SETICON = &H170
Public Const STM_GETICON = &H171
Public Const STM_MSGMAX = &H172
Public Const WC_DIALOG = 8002&
'  Get/SetWindowWord/Long offsets for use with WC_DIALOG windows

Public Const DWL_MSGRESULT = 0
Public Const DWL_DLGPROC = 4
Public Const DWL_USER = 8
' DlgDirList, DlgDirListComboBox flags values

Public Const DDL_READWRITE = &H0
Public Const DDL_READONLY = &H1
Public Const DDL_HIDDEN = &H2
Public Const DDL_SYSTEM = &H4
Public Const DDL_DIRECTORY = &H10
Public Const DDL_ARCHIVE = &H20
Public Const DDL_POSTMSGS = &H2000
Public Const DDL_DRIVES = &H4000
Public Const DDL_EXCLUSIVE = &H8000
' Dialog Styles

Public Const DS_ABSALIGN = &H1&
Public Const DS_SYSMODAL = &H2&
Public Const DS_LOCALEDIT = &H20          '  Edit items get Local storage.
Public Const DS_SETFONT = &H40            '  User specified font for Dlg controls
Public Const DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION
Public Const DS_NOIDLEMSG = &H100         '  WM_ENTERIDLE message will not be sent
Public Const DS_SETFOREGROUND = &H200     '  not in win3.1
Public Const DM_GETDEFID = WM_USER + 0
Public Const DM_SETDEFID = WM_USER + 1
Public Const DC_HASDEFID = &H534      '0x534B
' Dialog Codes

Public Const DLGC_WANTARROWS = &H1              '  Control wants arrow keys
Public Const DLGC_WANTTAB = &H2                 '  Control wants tab keys
Public Const DLGC_WANTALLKEYS = &H4             '  Control wants all keys
Public Const DLGC_WANTMESSAGE = &H4             '  Pass message to control
Public Const DLGC_HASSETSEL = &H8               '  Understands EM_SETSEL message
Public Const DLGC_DEFPUSHBUTTON = &H10          '  Default pushbutton
Public Const DLGC_UNDEFPUSHBUTTON = &H20        '  Non-default pushbutton
Public Const DLGC_RADIOBUTTON = &H40            '  Radio button
Public Const DLGC_WANTCHARS = &H80              '  Want WM_CHAR messages
Public Const DLGC_STATIC = &H100                '  Static item: don't include
Public Const DLGC_BUTTON = &H2000               '  Button item: can be checked
Public Const LB_CTLCODE = 0&
' Listbox Return Values

Public Const LB_OKAY = 0
Public Const LB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)
' The idStaticPath parameter to DlgDirList can have the following values
' ORed if the list box should show other details of the files along with
' the name of the files;
' all other details also will be returned
' Listbox Notification Codes

Public Const LBN_ERRSPACE = (-2)
Public Const LBN_SELCHANGE = 1
Public Const LBN_DBLCLK = 2
Public Const LBN_SELCANCEL = 3
Public Const LBN_SETFOCUS = 4
Public Const LBN_KILLFOCUS = 5
' Listbox messages

Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_DELETESTRING = &H182
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_RESETCONTENT = &H184
Public Const LB_SETSEL = &H185
Public Const LB_SETCURSEL = &H186
Public Const LB_GETSEL = &H187
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const LB_SELECTSTRING = &H18C
Public Const LB_DIR = &H18D
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_SETTABSTOPS = &H192
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETCOLUMNWIDTH = &H195
Public Const LB_ADDFILE = &H196
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETITEMDATA = &H19A
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SETANCHORINDEX = &H19C
Public Const LB_GETANCHORINDEX = &H19D
Public Const LB_SETCARETINDEX = &H19E
Public Const LB_GETCARETINDEX = &H19F
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SETLOCALE = &H1A5
Public Const LB_GETLOCALE = &H1A6
Public Const LB_SETCOUNT = &H1A7
Public Const LB_MSGMAX = &H1A8
' Listbox Styles

Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_NODATA = &H2000&
Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)
' Combo Box return Values

Public Const CB_OKAY = 0
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)
' Combo Box Notification Codes

Public Const CBN_ERRSPACE = (-1)
Public Const CBN_SELCHANGE = 1
Public Const CBN_DBLCLK = 2
Public Const CBN_SETFOCUS = 3
Public Const CBN_KILLFOCUS = 4
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_DROPDOWN = 7
Public Const CBN_CLOSEUP = 8
Public Const CBN_SELENDOK = 9
Public Const CBN_SELENDCANCEL = 10
' Combo Box styles

Public Const CBS_SIMPLE = &H1&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_SORT = &H100&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_DISABLENOSCROLL = &H800&
' Combo Box messages

Public Const CB_GETEDITSEL = &H140
Public Const CB_LIMITTEXT = &H141
Public Const CB_SETEDITSEL = &H142
Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_INSERTSTRING = &H14A
Public Const CB_RESETCONTENT = &H14B
Public Const CB_FINDSTRING = &H14C
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETITEMDATA = &H151
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SETLOCALE = &H159
Public Const CB_GETLOCALE = &H15A
Public Const CB_MSGMAX = &H15B
' Scroll Bar Styles

Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&
'  Scroll bar messages

Public Const SBM_SETPOS = &HE0 ' not in win3.1
Public Const SBM_GETPOS = &HE1 ' not in win3.1
Public Const SBM_SETRANGE = &HE2 ' not in win3.1
Public Const SBM_SETRANGEREDRAW = &HE6 ' not in win3.1
Public Const SBM_GETRANGE = &HE3 ' not in win3.1
Public Const SBM_ENABLE_ARROWS = &HE4 ' not in win3.1
Public Const MDIS_ALLCHILDSTYLES = &H1
' wParam values for WM_MDITILE and WM_MDICASCADE messages.

Public Const MDITILE_VERTICAL = &H0
Public Const MDITILE_HORIZONTAL = &H1
Public Const MDITILE_SKIPDISABLED = &H2
' Commands to pass WinHelp()

Public Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Public Const HELP_QUIT = &H2             '  Terminate help
Public Const HELP_INDEX = &H3            '  Display index
Public Const HELP_CONTENTS = &H3&
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Public Const HELP_COMMAND = &H102&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_SETWINPOS = &H203&
' Parameter for SystemParametersInfo()

Public Const SPI_GETBEEP = 1
Public Const SPI_SETBEEP = 2
Public Const SPI_GETMOUSE = 3
Public Const SPI_SETMOUSE = 4
Public Const SPI_GETBORDER = 5
Public Const SPI_SETBORDER = 6
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_LANGDRIVER = 12
Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_ICONVERTICALSPACING = 24
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_SETFASTTASKSWITCH = 36
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_GETMINIMIZEDMETRICS = 43
Public Const SPI_SETMINIMIZEDMETRICS = 44
Public Const SPI_GETICONMETRICS = 45
Public Const SPI_SETICONMETRICS = 46
Public Const SPI_SETWORKAREA = 47
Public Const SPI_GETWORKAREA = 48
Public Const SPI_SETPENWINDOWS = 49
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_SETFILTERKEYS = 51
Public Const SPI_GETTOGGLEKEYS = 52
Public Const SPI_SETTOGGLEKEYS = 53
Public Const SPI_GETMOUSEKEYS = 54
Public Const SPI_SETMOUSEKEYS = 55
Public Const SPI_GETSHOWSOUNDS = 56
Public Const SPI_SETSHOWSOUNDS = 57
Public Const SPI_GETSTICKYKEYS = 58
Public Const SPI_SETSTICKYKEYS = 59
Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_SETACCESSTIMEOUT = 61
Public Const SPI_GETSERIALKEYS = 62
Public Const SPI_SETSERIALKEYS = 63
Public Const SPI_GETSOUNDSENTRY = 64
Public Const SPI_SETSOUNDSENTRY = 65
Public Const SPI_GETHIGHCONTRAST = 66
Public Const SPI_SETHIGHCONTRAST = 67
Public Const SPI_GETKEYBOARDPREF = 68
Public Const SPI_SETKEYBOARDPREF = 69
Public Const SPI_GETSCREENREADER = 70
Public Const SPI_SETSCREENREADER = 71
Public Const SPI_GETANIMATION = 72
Public Const SPI_SETANIMATION = 73
Public Const SPI_GETFONTSMOOTHING = 74
Public Const SPI_SETFONTSMOOTHING = 75
Public Const SPI_SETDRAGWIDTH = 76
Public Const SPI_SETDRAGHEIGHT = 77
Public Const SPI_SETHANDHELD = 78
Public Const SPI_GETLOWPOWERTIMEOUT = 79
Public Const SPI_GETPOWEROFFTIMEOUT = 80
Public Const SPI_SETLOWPOWERTIMEOUT = 81
Public Const SPI_SETPOWEROFFTIMEOUT = 82
Public Const SPI_GETLOWPOWERACTIVE = 83
Public Const SPI_GETPOWEROFFACTIVE = 84
Public Const SPI_SETLOWPOWERACTIVE = 85
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_SETCURSORS = 87
Public Const SPI_SETICONS = 88
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_SETDEFAULTINPUTLANG = 90
Public Const SPI_SETLANGTOGGLE = 91
Public Const SPI_GETWINDOWSEXTENSION = 92
Public Const SPI_SETMOUSETRAILS = 93
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_SCREENSAVERRUNNING = 97
' SystemParametersInfo flags

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
'  DDE window messages

Public Const WM_DDE_FIRST = &H3E0
Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
'  conversation states (usState)

Public Const XST_NULL = 0  '  quiescent states
Public Const XST_INCOMPLETE = 1
Public Const XST_CONNECTED = 2
Public Const XST_INIT1 = 3  '  mid-initiation states
Public Const XST_INIT2 = 4
Public Const XST_REQSENT = 5  '  active conversation states
Public Const XST_DATARCVD = 6
Public Const XST_POKESENT = 7
Public Const XST_POKEACKRCVD = 8
Public Const XST_EXECSENT = 9
Public Const XST_EXECACKRCVD = 10
Public Const XST_ADVSENT = 11
Public Const XST_UNADVSENT = 12
Public Const XST_ADVACKRCVD = 13
Public Const XST_UNADVACKRCVD = 14
Public Const XST_ADVDATASENT = 15
Public Const XST_ADVDATAACKRCVD = 16
'  used in LOWORD(dwData1) of XTYP_ADVREQ callbacks...

Public Const CADV_LATEACK = &HFFFF
'  conversation status bits (fsStatus)

Public Const ST_CONNECTED = &H1
Public Const ST_ADVISE = &H2
Public Const ST_ISLOCAL = &H4
Public Const ST_BLOCKED = &H8
Public Const ST_CLIENT = &H10
Public Const ST_TERMINATED = &H20
Public Const ST_INLIST = &H40
Public Const ST_BLOCKNEXT = &H80
Public Const ST_ISSELF = &H100
'  DDE constants for wStatus field

Public Const DDE_FACK = &H8000
Public Const DDE_FBUSY = &H4000
Public Const DDE_FDEFERUPD = &H4000
Public Const DDE_FACKREQ = &H8000
Public Const DDE_FRELEASE = &H2000
Public Const DDE_FREQUESTED = &H1000
Public Const DDE_FAPPSTATUS = &HFF
Public Const DDE_FNOTPROCESSED = &H0
Public Const DDE_FACKRESERVED = (Not (DDE_FACK Or DDE_FBUSY Or DDE_FAPPSTATUS))
Public Const DDE_FADVRESERVED = (Not (DDE_FACKREQ Or DDE_FDEFERUPD))
Public Const DDE_FDATRESERVED = (Not (DDE_FACKREQ Or DDE_FRELEASE Or DDE_FREQUESTED))
Public Const DDE_FPOKRESERVED = (Not (DDE_FRELEASE))
'  message filter hook types

Public Const MSGF_DDEMGR = &H8001
'  codepage constants

Public Const CP_WINANSI = 1004  '  default codepage for windows old DDE convs.
Public Const CP_WINUNICODE = 1200
'  transaction types

Public Const XTYPF_NOBLOCK = &H2     '  CBR_BLOCK will not work
Public Const XTYPF_NODATA = &H4     '  DDE_FDEFERUPD
Public Const XTYPF_ACKREQ = &H8     '  DDE_FACKREQ
Public Const XCLASS_MASK = &HFC00
Public Const XCLASS_BOOL = &H1000
Public Const XCLASS_DATA = &H2000
Public Const XCLASS_FLAGS = &H4000
Public Const XCLASS_NOTIFICATION = &H8000
Public Const XTYP_ERROR = (&H0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Public Const XTYP_ADVDATA = (&H10 Or XCLASS_FLAGS)
Public Const XTYP_ADVREQ = (&H20 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Public Const XTYP_ADVSTART = (&H30 Or XCLASS_BOOL)
Public Const XTYP_ADVSTOP = (&H40 Or XCLASS_NOTIFICATION)
Public Const XTYP_EXECUTE = (&H50 Or XCLASS_FLAGS)
Public Const XTYP_CONNECT = (&H60 Or XCLASS_BOOL Or XTYPF_NOBLOCK)
Public Const XTYP_CONNECT_CONFIRM = (&H70 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Public Const XTYP_XACT_COMPLETE = (&H80 Or XCLASS_NOTIFICATION)
Public Const XTYP_POKE = (&H90 Or XCLASS_FLAGS)
Public Const XTYP_REGISTER = (&HA0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Public Const XTYP_REQUEST = (&HB0 Or XCLASS_DATA)
Public Const XTYP_DISCONNECT = (&HC0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Public Const XTYP_UNREGISTER = (&HD0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Public Const XTYP_WILDCONNECT = (&HE0 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Public Const XTYP_MASK = &HF0
Public Const XTYP_SHIFT = 4  '  shift to turn XTYP_ into an index
'  Timeout constants

Public Const TIMEOUT_ASYNC = &HFFFF
'  Transaction ID constants

Public Const QID_SYNC = &HFFFF
' Public strings used in DDE

Public Const SZDDESYS_TOPIC = "System"
Public Const SZDDESYS_ITEM_TOPICS = "Topics"
Public Const SZDDESYS_ITEM_SYSITEMS = "SysItems"
Public Const SZDDESYS_ITEM_RTNMSG = "ReturnMessage"
Public Const SZDDESYS_ITEM_STATUS = "Status"
Public Const SZDDESYS_ITEM_FORMATS = "Formats"
Public Const SZDDESYS_ITEM_HELP = "Help"
Public Const SZDDE_ITEM_ITEMLIST = "TopicItemList"
Public Const CBR_BLOCK = &HFFFF
' Callback filter flags for use with standard apps.

Public Const CBF_FAIL_SELFCONNECTIONS = &H1000
Public Const CBF_FAIL_CONNECTIONS = &H2000
Public Const CBF_FAIL_ADVISES = &H4000
Public Const CBF_FAIL_EXECUTES = &H8000
Public Const CBF_FAIL_POKES = &H10000
Public Const CBF_FAIL_REQUESTS = &H20000
Public Const CBF_FAIL_ALLSVRXACTIONS = &H3F000
Public Const CBF_SKIP_CONNECT_CONFIRMS = &H40000
Public Const CBF_SKIP_REGISTRATIONS = &H80000
Public Const CBF_SKIP_UNREGISTRATIONS = &H100000
Public Const CBF_SKIP_DISCONNECTS = &H200000
Public Const CBF_SKIP_ALLNOTIFICATIONS = &H3C0000
' Application command flags

Public Const APPCMD_CLIENTONLY = &H10&
Public Const APPCMD_FILTERINITS = &H20&
Public Const APPCMD_MASK = &HFF0&
' Application classification flags

Public Const APPCLASS_STANDARD = &H0&
Public Const APPCLASS_MASK = &HF&
Public Const EC_ENABLEALL = 0
Public Const EC_ENABLEONE = ST_BLOCKNEXT
Public Const EC_DISABLE = ST_BLOCKED
Public Const EC_QUERYWAITING = 2
Public Const DNS_REGISTER = &H1
Public Const DNS_UNREGISTER = &H2
Public Const DNS_FILTERON = &H4
Public Const DNS_FILTEROFF = &H8
Public Const HDATA_APPOWNED = &H1
Public Const DMLERR_NO_ERROR = 0                           '  must be 0
Public Const DMLERR_FIRST = &H4000
Public Const DMLERR_ADVACKTIMEOUT = &H4000
Public Const DMLERR_BUSY = &H4001
Public Const DMLERR_DATAACKTIMEOUT = &H4002
Public Const DMLERR_DLL_NOT_INITIALIZED = &H4003
Public Const DMLERR_DLL_USAGE = &H4004
Public Const DMLERR_EXECACKTIMEOUT = &H4005
Public Const DMLERR_INVALIDPARAMETER = &H4006
Public Const DMLERR_LOW_MEMORY = &H4007
Public Const DMLERR_MEMORY_ERROR = &H4008
Public Const DMLERR_NOTPROCESSED = &H4009
Public Const DMLERR_NO_CONV_ESTABLISHED = &H400A
Public Const DMLERR_POKEACKTIMEOUT = &H400B
Public Const DMLERR_POSTMSG_FAILED = &H400C
Public Const DMLERR_REENTRANCY = &H400D
Public Const DMLERR_SERVER_DIED = &H400E
Public Const DMLERR_SYS_ERROR = &H400F
Public Const DMLERR_UNADVACKTIMEOUT = &H4010
Public Const DMLERR_UNFOUND_QUEUE_ID = &H4011
Public Const DMLERR_LAST = &H4011
Public Const MH_CREATE = 1
Public Const MH_KEEP = 2
Public Const MH_DELETE = 3
Public Const MH_CLEANUP = 4
Public Const MAX_MONITORS = 4
Public Const APPCLASS_MONITOR = &H1&
Public Const XTYP_MONITOR = (&HF0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
' Callback filter flags for use with MONITOR apps - 0 implies no monitor callbacks

Public Const MF_HSZ_INFO = &H1000000
Public Const MF_SENDMSGS = &H2000000
Public Const MF_POSTMSGS = &H4000000
Public Const MF_CALLBACKS = &H8000000
Public Const MF_ERRORS = &H10000000
Public Const MF_LINKS = &H20000000
Public Const MF_CONV = &H40000000
Public Const MF_MASK = &HFF000000
' -----------------------------------------
' Win32 API error code definitions
' -----------------------------------------
' This section contains the error code definitions for the Win32 API functions.
' NO_ERROR

Public Const NO_ERROR = 0 '  dderror
' The configuration registry database operation completed successfully.

Public Const ERROR_SUCCESS = 0&
'   Incorrect function.

Public Const ERROR_INVALID_FUNCTION = 1 '  dderror
'   The system cannot find the file specified.

Public Const ERROR_FILE_NOT_FOUND = 2&
'   The system cannot find the path specified.

Public Const ERROR_PATH_NOT_FOUND = 3&
'   The system cannot open the file.

Public Const ERROR_TOO_MANY_OPEN_FILES = 4&
'   Access is denied.

Public Const ERROR_ACCESS_DENIED = 5&
'   The handle is invalid.

Public Const ERROR_INVALID_HANDLE = 6&
'   The storage control blocks were destroyed.

Public Const ERROR_ARENA_TRASHED = 7&
'   Not enough storage is available to process this command.

Public Const ERROR_NOT_ENOUGH_MEMORY = 8 '  dderror
'   The storage control block address is invalid.

Public Const ERROR_INVALID_BLOCK = 9&
'   The environment is incorrect.

Public Const ERROR_BAD_ENVIRONMENT = 10&
'   An attempt was made to load a program with an
'   incorrect format.

Public Const ERROR_BAD_FORMAT = 11&
'   The access code is invalid.

Public Const ERROR_INVALID_ACCESS = 12&
'   The data is invalid.

Public Const ERROR_INVALID_DATA = 13&
'   Not enough storage is available to complete this operation.

Public Const ERROR_OUTOFMEMORY = 14&
'   The system cannot find the drive specified.

Public Const ERROR_INVALID_DRIVE = 15&
'   The directory cannot be removed.

Public Const ERROR_CURRENT_DIRECTORY = 16&
'   The system cannot move the file
'   to a different disk drive.

Public Const ERROR_NOT_SAME_DEVICE = 17&
'   There are no more files.

Public Const ERROR_NO_MORE_FILES = 18&
'   The media is write protected.

Public Const ERROR_WRITE_PROTECT = 19&
'   The system cannot find the device specified.

Public Const ERROR_BAD_UNIT = 20&
'   The device is not ready.

Public Const ERROR_NOT_READY = 21&
'   The device does not recognize the command.

Public Const ERROR_BAD_COMMAND = 22&
'   Data error (cyclic redundancy check)

Public Const ERROR_CRC = 23&
'   The program issued a command but the
'   command length is incorrect.

Public Const ERROR_BAD_LENGTH = 24&
'   The drive cannot locate a specific
'   area or track on the disk.

Public Const ERROR_SEEK = 25&
'   The specified disk or diskette cannot be accessed.

Public Const ERROR_NOT_DOS_DISK = 26&
'   The drive cannot find the sector requested.

Public Const ERROR_SECTOR_NOT_FOUND = 27&
'   The printer is out of paper.

Public Const ERROR_OUT_OF_PAPER = 28&
'   The system cannot write to the specified device.

Public Const ERROR_WRITE_FAULT = 29&
'   The system cannot read from the specified device.

Public Const ERROR_READ_FAULT = 30&
'   A device attached to the system is not functioning.

Public Const ERROR_GEN_FAILURE = 31&
'   The process cannot access the file because
'   it is being used by another process.

Public Const ERROR_SHARING_VIOLATION = 32&
'   The process cannot access the file because
'   another process has locked a portion of the file.

Public Const ERROR_LOCK_VIOLATION = 33&
'   The wrong diskette is in the drive.
'   Insert %2 (Volume Serial Number: %3)
'   into drive %1.

Public Const ERROR_WRONG_DISK = 34&
'   Too many files opened for sharing.

Public Const ERROR_SHARING_BUFFER_EXCEEDED = 36&
'   Reached end of file.

Public Const ERROR_HANDLE_EOF = 38&
'   The disk is full.

Public Const ERROR_HANDLE_DISK_FULL = 39&
'   The network request is not supported.

Public Const ERROR_NOT_SUPPORTED = 50&
'   The remote computer is not available.

Public Const ERROR_REM_NOT_LIST = 51&
'   A duplicate name exists on the network.

Public Const ERROR_DUP_NAME = 52&
'   The network path was not found.

Public Const ERROR_BAD_NETPATH = 53&
'   The network is busy.

Public Const ERROR_NETWORK_BUSY = 54&
'   The specified network resource or device is no longer
'   available.

Public Const ERROR_DEV_NOT_EXIST = 55 '  dderror
'   The network BIOS command limit has been reached.

Public Const ERROR_TOO_MANY_CMDS = 56&
'   A network adapter hardware error occurred.

Public Const ERROR_ADAP_HDW_ERR = 57&
'   The specified server cannot perform the requested
'   operation.

Public Const ERROR_BAD_NET_RESP = 58&
'   An unexpected network error occurred.

Public Const ERROR_UNEXP_NET_ERR = 59&
'   The remote adapter is not compatible.

Public Const ERROR_BAD_REM_ADAP = 60&
'   The printer queue is full.

Public Const ERROR_PRINTQ_FULL = 61&
'   Space to store the file waiting to be printed is
'   not available on the server.

Public Const ERROR_NO_SPOOL_SPACE = 62&
'   Your file waiting to be printed was deleted.

Public Const ERROR_PRINT_CANCELLED = 63&
'   The specified network name is no longer available.

Public Const ERROR_NETNAME_DELETED = 64&
'   Network access is denied.

Public Const ERROR_NETWORK_ACCESS_DENIED = 65&
'   The network resource type is not correct.

Public Const ERROR_BAD_DEV_TYPE = 66&
'   The network name cannot be found.

Public Const ERROR_BAD_NET_NAME = 67&
'   The name limit for the local computer network
'   adapter card was exceeded.

Public Const ERROR_TOO_MANY_NAMES = 68&
'   The network BIOS session limit was exceeded.

Public Const ERROR_TOO_MANY_SESS = 69&
'   The remote server has been paused or is in the
'   process of being started.

Public Const ERROR_SHARING_PAUSED = 70&
'   The network request was not accepted.

Public Const ERROR_REQ_NOT_ACCEP = 71&
'   The specified printer or disk device has been paused.

Public Const ERROR_REDIR_PAUSED = 72&
'   The file exists.

Public Const ERROR_FILE_EXISTS = 80&
'   The directory or file cannot be created.

Public Const ERROR_CANNOT_MAKE = 82&
'   Fail on INT 24

Public Const ERROR_FAIL_I24 = 83&
'   Storage to process this request is not available.

Public Const ERROR_OUT_OF_STRUCTURES = 84&
'   The local device name is already in use.

Public Const ERROR_ALREADY_ASSIGNED = 85&
'   The specified network password is not correct.

Public Const ERROR_INVALID_PASSWORD = 86&
'   The parameter is incorrect.

Public Const ERROR_INVALID_PARAMETER = 87 '  dderror
'   A write fault occurred on the network.

Public Const ERROR_NET_WRITE_FAULT = 88&
'   The system cannot start another process at
'   this time.

Public Const ERROR_NO_PROC_SLOTS = 89&
'   Cannot create another system semaphore.

Public Const ERROR_TOO_MANY_SEMAPHORES = 100&
'   The exclusive semaphore is owned by another process.

Public Const ERROR_EXCL_SEM_ALREADY_OWNED = 101&
'   The semaphore is set and cannot be closed.

Public Const ERROR_SEM_IS_SET = 102&
'   The semaphore cannot be set again.

Public Const ERROR_TOO_MANY_SEM_REQUESTS = 103&
'   Cannot request exclusive semaphores at interrupt time.

Public Const ERROR_INVALID_AT_INTERRUPT_TIME = 104&
'   The previous ownership of this semaphore has ended.

Public Const ERROR_SEM_OWNER_DIED = 105&
'   Insert the diskette for drive %1.

Public Const ERROR_SEM_USER_LIMIT = 106&
'   Program stopped because alternate diskette was not inserted.

Public Const ERROR_DISK_CHANGE = 107&
'   The disk is in use or locked by
'   another process.

Public Const ERROR_DRIVE_LOCKED = 108&
'   The pipe has been ended.

Public Const ERROR_BROKEN_PIPE = 109&
'   The system cannot open the
'   device or file specified.

Public Const ERROR_OPEN_FAILED = 110&
'   The file name is too long.

Public Const ERROR_BUFFER_OVERFLOW = 111&
'   There is not enough space on the disk.

Public Const ERROR_DISK_FULL = 112&
'   No more internal file identifiers available.

Public Const ERROR_NO_MORE_SEARCH_HANDLES = 113&
'   The target internal file identifier is incorrect.

Public Const ERROR_INVALID_TARGET_HANDLE = 114&
'   The IOCTL call made by the application program is
'   not correct.

Public Const ERROR_INVALID_CATEGORY = 117&
'   The verify-on-write switch parameter value is not
'   correct.

Public Const ERROR_INVALID_VERIFY_SWITCH = 118&
'   The system does not support the command requested.

Public Const ERROR_BAD_DRIVER_LEVEL = 119&
'   This function is only valid in Windows NT mode.

Public Const ERROR_CALL_NOT_IMPLEMENTED = 120&
'   The semaphore timeout period has expired.

Public Const ERROR_SEM_TIMEOUT = 121&
'   The data area passed to a system call is too
'   small.

Public Const ERROR_INSUFFICIENT_BUFFER = 122 '  dderror
'   The filename, directory name, or volume label syntax is incorrect.

Public Const ERROR_INVALID_NAME = 123&
'   The system call level is not correct.

Public Const ERROR_INVALID_LEVEL = 124&
'   The disk has no volume label.

Public Const ERROR_NO_VOLUME_LABEL = 125&
'   The specified module could not be found.

Public Const ERROR_MOD_NOT_FOUND = 126&
'   The specified procedure could not be found.

Public Const ERROR_PROC_NOT_FOUND = 127&
'   There are no child processes to wait for.

Public Const ERROR_WAIT_NO_CHILDREN = 128&
'   The %1 application cannot be run in Windows NT mode.

Public Const ERROR_CHILD_NOT_COMPLETE = 129&
'   Attempt to use a file handle to an open disk partition for an
'   operation other than raw disk I/O.

Public Const ERROR_DIRECT_ACCESS_HANDLE = 130&
'   An attempt was made to move the file pointer before the beginning of the file.

Public Const ERROR_NEGATIVE_SEEK = 131&
'   The file pointer cannot be set on the specified device or file.

Public Const ERROR_SEEK_ON_DEVICE = 132&
'   A JOIN or SUBST command
'   cannot be used for a drive that
'   contains previously joined drives.

Public Const ERROR_IS_JOIN_TARGET = 133&
'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been joined.

Public Const ERROR_IS_JOINED = 134&
'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been substituted.

Public Const ERROR_IS_SUBSTED = 135&
'   The system tried to delete
'   the JOIN of a drive that is not joined.

Public Const ERROR_NOT_JOINED = 136&
'   The system tried to delete the
'   substitution of a drive that is not substituted.

Public Const ERROR_NOT_SUBSTED = 137&
'   The system tried to join a drive
'   to a directory on a joined drive.

Public Const ERROR_JOIN_TO_JOIN = 138&
'   The system tried to substitute a
'   drive to a directory on a substituted drive.

Public Const ERROR_SUBST_TO_SUBST = 139&
'   The system tried to join a drive to
'   a directory on a substituted drive.

Public Const ERROR_JOIN_TO_SUBST = 140&
'   The system tried to SUBST a drive
'   to a directory on a joined drive.

Public Const ERROR_SUBST_TO_JOIN = 141&
'   The system cannot perform a JOIN or SUBST at this time.

Public Const ERROR_BUSY_DRIVE = 142&
'   The system cannot join or substitute a
'   drive to or for a directory on the same drive.

Public Const ERROR_SAME_DRIVE = 143&
'   The directory is not a subdirectory of the root directory.

Public Const ERROR_DIR_NOT_ROOT = 144&
'   The directory is not empty.

Public Const ERROR_DIR_NOT_EMPTY = 145&
'   The path specified is being used in
'   a substitute.

Public Const ERROR_IS_SUBST_PATH = 146&
'   Not enough resources are available to
'   process this command.

Public Const ERROR_IS_JOIN_PATH = 147&
'   The path specified cannot be used at this time.

Public Const ERROR_PATH_BUSY = 148&
'   An attempt was made to join
'   or substitute a drive for which a directory
'   on the drive is the target of a previous
'   substitute.

Public Const ERROR_IS_SUBST_TARGET = 149&
'   System trace information was not specified in your
'   CONFIG.SYS file, or tracing is disallowed.

Public Const ERROR_SYSTEM_TRACE = 150&
'   The number of specified semaphore events for
'   DosMuxSemWait is not correct.

Public Const ERROR_INVALID_EVENT_COUNT = 151&
'   DosMuxSemWait did not execute; too many semaphores
'   are already set.

Public Const ERROR_TOO_MANY_MUXWAITERS = 152&
'   The DosMuxSemWait list is not correct.

Public Const ERROR_INVALID_LIST_FORMAT = 153&
'   The volume label you entered exceeds the
'   11 character limit.  The first 11 characters were written
'   to disk.  Any characters that exceeded the 11 character limit
'   were automatically deleted.

Public Const ERROR_LABEL_TOO_LONG = 154&
'   Cannot create another thread.

Public Const ERROR_TOO_MANY_TCBS = 155&
'   The recipient process has refused the signal.

Public Const ERROR_SIGNAL_REFUSED = 156&
'   The segment is already discarded and cannot be locked.

Public Const ERROR_DISCARDED = 157&
'   The segment is already unlocked.

Public Const ERROR_NOT_LOCKED = 158&
'   The address for the thread ID is not correct.

Public Const ERROR_BAD_THREADID_ADDR = 159&
'   The argument string passed to DosExecPgm is not correct.

Public Const ERROR_BAD_ARGUMENTS = 160&
'   The specified path is invalid.

Public Const ERROR_BAD_PATHNAME = 161&
'   A signal is already pending.

Public Const ERROR_SIGNAL_PENDING = 162&
'   No more threads can be created in the system.

Public Const ERROR_MAX_THRDS_REACHED = 164&
'   Unable to lock a region of a file.

Public Const ERROR_LOCK_FAILED = 167&
'   The requested resource is in use.

Public Const ERROR_BUSY = 170&
'   A lock request was not outstanding for the supplied cancel region.

Public Const ERROR_CANCEL_VIOLATION = 173&
'   The file system does not support atomic changes to the lock type.

Public Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED = 174&
'   The system detected a segment number that was not correct.

Public Const ERROR_INVALID_SEGMENT_NUMBER = 180&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_ORDINAL = 182&
'   Cannot create a file when that file already exists.

Public Const ERROR_ALREADY_EXISTS = 183&
'   The flag passed is not correct.

Public Const ERROR_INVALID_FLAG_NUMBER = 186&
'   The specified system semaphore name was not found.

Public Const ERROR_SEM_NOT_FOUND = 187&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_STARTING_CODESEG = 188&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_STACKSEG = 189&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_MODULETYPE = 190&
'   Cannot run %1 in Windows NT mode.

Public Const ERROR_INVALID_EXE_SIGNATURE = 191&
'   The operating system cannot run %1.

Public Const ERROR_EXE_MARKED_INVALID = 192&
'   %1 is not a valid Windows NT application.

Public Const ERROR_BAD_EXE_FORMAT = 193&
'   The operating system cannot run %1.

Public Const ERROR_ITERATED_DATA_EXCEEDS_64k = 194&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_MINALLOCSIZE = 195&
'   The operating system cannot run this
'   application program.

Public Const ERROR_DYNLINK_FROM_INVALID_RING = 196&
'   The operating system is not presently
'   configured to run this application.

Public Const ERROR_IOPL_NOT_ENABLED = 197&
'   The operating system cannot run %1.

Public Const ERROR_INVALID_SEGDPL = 198&
'   The operating system cannot run this
'   application program.

Public Const ERROR_AUTODATASEG_EXCEEDS_64k = 199&
'   The code segment cannot be greater than or equal to 64KB.

Public Const ERROR_RING2SEG_MUST_BE_MOVABLE = 200&
'   The operating system cannot run %1.

Public Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM = 201&
'   The operating system cannot run %1.

Public Const ERROR_INFLOOP_IN_RELOC_CHAIN = 202&
'   The system could not find the environment
'   option that was entered.

Public Const ERROR_ENVVAR_NOT_FOUND = 203&
'   No process in the command subtree has a
'   signal handler.

Public Const ERROR_NO_SIGNAL_SENT = 205&
'   The filename or extension is too long.

Public Const ERROR_FILENAME_EXCED_RANGE = 206&
'   The ring 2 stack is in use.

Public Const ERROR_RING2_STACK_IN_USE = 207&
'   The Global filename characters,  or ?, are entered
'   incorrectly or too many Global filename characters are specified.

Public Const ERROR_META_EXPANSION_TOO_LONG = 208&
'   The signal being posted is not correct.

Public Const ERROR_INVALID_SIGNAL_NUMBER = 209&
'   The signal handler cannot be set.

Public Const ERROR_THREAD_1_INACTIVE = 210&
'   The segment is locked and cannot be reallocated.

Public Const ERROR_LOCKED = 212&
'   Too many dynamic link modules are attached to this
'   program or dynamic link module.

Public Const ERROR_TOO_MANY_MODULES = 214&
'   Can't nest calls to LoadModule.

Public Const ERROR_NESTING_NOT_ALLOWED = 215&
'   The pipe state is invalid.

Public Const ERROR_BAD_PIPE = 230&
'   All pipe instances are busy.

Public Const ERROR_PIPE_BUSY = 231&
'   The pipe is being closed.

Public Const ERROR_NO_DATA = 232&
'   No process is on the other end of the pipe.

Public Const ERROR_PIPE_NOT_CONNECTED = 233&
'   More data is available.

Public Const ERROR_MORE_DATA = 234 '  dderror
'   The session was cancelled.

Public Const ERROR_VC_DISCONNECTED = 240&
'   The specified extended attribute name was invalid.

Public Const ERROR_INVALID_EA_NAME = 254&
'   The extended attributes are inconsistent.

Public Const ERROR_EA_LIST_INCONSISTENT = 255&
'   No more data is available.

Public Const ERROR_NO_MORE_ITEMS = 259&
'   The Copy API cannot be used.

Public Const ERROR_CANNOT_COPY = 266&
'   The directory name is invalid.

Public Const ERROR_DIRECTORY = 267&
'   The extended attributes did not fit in the buffer.

Public Const ERROR_EAS_DIDNT_FIT = 275&
'   The extended attribute file on the mounted file system is corrupt.

Public Const ERROR_EA_FILE_CORRUPT = 276&
'   The extended attribute table file is full.

Public Const ERROR_EA_TABLE_FULL = 277&
'   The specified extended attribute handle is invalid.

Public Const ERROR_INVALID_EA_HANDLE = 278&
'   The mounted file system does not support extended attributes.

Public Const ERROR_EAS_NOT_SUPPORTED = 282&
'   Attempt to release mutex not owned by caller.

Public Const ERROR_NOT_OWNER = 288&
'   Too many posts were made to a semaphore.

Public Const ERROR_TOO_MANY_POSTS = 298&
'   The system cannot find message for message number 0x%1
'   in message file for %2.

Public Const ERROR_MR_MID_NOT_FOUND = 317&
'   Attempt to access invalid address.

Public Const ERROR_INVALID_ADDRESS = 487&
'   Arithmetic result exceeded 32 bits.

Public Const ERROR_ARITHMETIC_OVERFLOW = 534&
'   There is a process on other end of the pipe.

Public Const ERROR_PIPE_CONNECTED = 535&
'   Waiting for a process to open the other end of the pipe.

Public Const ERROR_PIPE_LISTENING = 536&
'   Access to the extended attribute was denied.

Public Const ERROR_EA_ACCESS_DENIED = 994&
'   The I/O operation has been aborted because of either a thread exit
'   or an application request.

Public Const ERROR_OPERATION_ABORTED = 995&
'   Overlapped I/O event is not in a signalled state.

Public Const ERROR_IO_INCOMPLETE = 996&
'   Overlapped I/O operation is in progress.

Public Const ERROR_IO_PENDING = 997 '  dderror
'   Invalid access to memory location.

Public Const ERROR_NOACCESS = 998&
'   Error performing inpage operation.

Public Const ERROR_SWAPERROR = 999&
'   Recursion too deep, stack overflowed.

Public Const ERROR_STACK_OVERFLOW = 1001&
'   The window cannot act on the sent message.

Public Const ERROR_INVALID_MESSAGE = 1002&
'   Cannot complete this function.

Public Const ERROR_CAN_NOT_COMPLETE = 1003&
'   Invalid flags.

Public Const ERROR_INVALID_FLAGS = 1004&
'   The volume does not contain a recognized file system.
'   Please make sure that all required file system drivers are loaded and that the
'   volume is not corrupt.

Public Const ERROR_UNRECOGNIZED_VOLUME = 1005&
'   The volume for a file has been externally altered such that the
'   opened file is no longer valid.

Public Const ERROR_FILE_INVALID = 1006&
'   The requested operation cannot be performed in full-screen mode.

Public Const ERROR_FULLSCREEN_MODE = 1007&
'   An attempt was made to reference a token that does not exist.

Public Const ERROR_NO_TOKEN = 1008&
'   The configuration registry database is corrupt.

Public Const ERROR_BADDB = 1009&
'   The configuration registry key is invalid.

Public Const ERROR_BADKEY = 1010&
'   The configuration registry key could not be opened.

Public Const ERROR_CANTOPEN = 1011&
'   The configuration registry key could not be read.

Public Const ERROR_CANTREAD = 1012&
'   The configuration registry key could not be written.

Public Const ERROR_CANTWRITE = 1013&
'   One of the files in the Registry database had to be recovered
'   by use of a log or alternate copy.  The recovery was successful.

Public Const ERROR_REGISTRY_RECOVERED = 1014&
'   The Registry is corrupt. The structure of one of the files that contains
'   Registry data is corrupt, or the system's image of the file in memory
'   is corrupt, or the file could not be recovered because the alternate
'   copy or log was absent or corrupt.

Public Const ERROR_REGISTRY_CORRUPT = 1015&
'   An I/O operation initiated by the Registry failed unrecoverably.
'   The Registry could not read in, or write out, or flush, one of the files
'   that contain the system's image of the Registry.

Public Const ERROR_REGISTRY_IO_FAILED = 1016&
'   The system has attempted to load or restore a file into the Registry, but the
'   specified file is not in a Registry file format.

Public Const ERROR_NOT_REGISTRY_FILE = 1017&
'   Illegal operation attempted on a Registry key which has been marked for deletion.

Public Const ERROR_KEY_DELETED = 1018&
'   System could not allocate the required space in a Registry log.

Public Const ERROR_NO_LOG_SPACE = 1019&
'   Cannot create a symbolic link in a Registry key that already
'   has subkeys or values.

Public Const ERROR_KEY_HAS_CHILDREN = 1020&
'   Cannot create a stable subkey under a volatile parent key.

Public Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
'   A notify change request is being completed and the information
'   is not being returned in the caller's buffer. The caller now
'   needs to enumerate the files to find the changes.

Public Const ERROR_NOTIFY_ENUM_DIR = 1022&
'   A stop control has been sent to a service which other running services
'   are dependent on.

Public Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051&
'   The requested control is not valid for this service

Public Const ERROR_INVALID_SERVICE_CONTROL = 1052&
'   The service did not respond to the start or control request in a timely
'   fashion.

Public Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053&
'   A thread could not be created for the service.

Public Const ERROR_SERVICE_NO_THREAD = 1054&
'   The service database is locked.

Public Const ERROR_SERVICE_DATABASE_LOCKED = 1055&
'   An instance of the service is already running.

Public Const ERROR_SERVICE_ALREADY_RUNNING = 1056&
'   The account name is invalid or does not exist.

Public Const ERROR_INVALID_SERVICE_ACCOUNT = 1057&
'   The specified service is disabled and cannot be started.

Public Const ERROR_SERVICE_DISABLED = 1058&
'   Circular service dependency was specified.

Public Const ERROR_CIRCULAR_DEPENDENCY = 1059&
'   The specified service does not exist as an installed service.

Public Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&
'   The service cannot accept control messages at this time.

Public Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL = 1061&
'   The service has not been started.

Public Const ERROR_SERVICE_NOT_ACTIVE = 1062&
'   The service process could not connect to the service controller.

Public Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063&
'   An exception occurred in the service when handling the control request.

Public Const ERROR_EXCEPTION_IN_SERVICE = 1064&
'   The database specified does not exist.

Public Const ERROR_DATABASE_DOES_NOT_EXIST = 1065&
'   The service has returned a service-specific error code.

Public Const ERROR_SERVICE_SPECIFIC_ERROR = 1066&
'   The process terminated unexpectedly.

Public Const ERROR_PROCESS_ABORTED = 1067&
'   The dependency service or group failed to start.

Public Const ERROR_SERVICE_DEPENDENCY_FAIL = 1068&
'   The service did not start due to a logon failure.

Public Const ERROR_SERVICE_LOGON_FAILED = 1069&
'   After starting, the service hung in a start-pending state.

Public Const ERROR_SERVICE_START_HANG = 1070&
'   The specified service database lock is invalid.

Public Const ERROR_INVALID_SERVICE_LOCK = 1071&
'   The specified service has been marked for deletion.

Public Const ERROR_SERVICE_MARKED_FOR_DELETE = 1072&
'   The specified service already exists.

Public Const ERROR_SERVICE_EXISTS = 1073&
'   The system is currently running with the last-known-good configuration.

Public Const ERROR_ALREADY_RUNNING_LKG = 1074&
'   The dependency service does not exist or has been marked for
'   deletion.

Public Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075&
'   The current boot has already been accepted for use as the
'   last-known-good control set.

Public Const ERROR_BOOT_ALREADY_ACCEPTED = 1076&
'   No attempts to start the service have been made since the last boot.

Public Const ERROR_SERVICE_NEVER_STARTED = 1077&
'   The name is already in use as either a service name or a service display
'   name.

Public Const ERROR_DUPLICATE_SERVICE_NAME = 1078&
'   The physical end of the tape has been reached.

Public Const ERROR_END_OF_MEDIA = 1100&
'   A tape access reached a filemark.

Public Const ERROR_FILEMARK_DETECTED = 1101&
'   Beginning of tape or partition was encountered.

Public Const ERROR_BEGINNING_OF_MEDIA = 1102&
'   A tape access reached the end of a set of files.

Public Const ERROR_SETMARK_DETECTED = 1103&
'   No more data is on the tape.

Public Const ERROR_NO_DATA_DETECTED = 1104&
'   Tape could not be partitioned.

Public Const ERROR_PARTITION_FAILURE = 1105&
'   When accessing a new tape of a multivolume partition, the current
'   blocksize is incorrect.

Public Const ERROR_INVALID_BLOCK_LENGTH = 1106&
'   Tape partition information could not be found when loading a tape.

Public Const ERROR_DEVICE_NOT_PARTITIONED = 1107&
'   Unable to lock the media eject mechanism.

Public Const ERROR_UNABLE_TO_LOCK_MEDIA = 1108&
'   Unable to unload the media.

Public Const ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109&
'   Media in drive may have changed.

Public Const ERROR_MEDIA_CHANGED = 1110&
'   The I/O bus was reset.

Public Const ERROR_BUS_RESET = 1111&
'   No media in drive.

Public Const ERROR_NO_MEDIA_IN_DRIVE = 1112&
'   No mapping for the Unicode character exists in the target multi-byte code page.

Public Const ERROR_NO_UNICODE_TRANSLATION = 1113&
'   A dynamic link library (DLL) initialization routine failed.

Public Const ERROR_DLL_INIT_FAILED = 1114&
'   A system shutdown is in progress.

Public Const ERROR_SHUTDOWN_IN_PROGRESS = 1115&
'   Unable to abort the system shutdown because no shutdown was in progress.

Public Const ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116&
'   The request could not be performed because of an I/O device error.

Public Const ERROR_IO_DEVICE = 1117&
'   No serial device was successfully initialized.  The serial driver will unload.

Public Const ERROR_SERIAL_NO_DEVICE = 1118&
'   Unable to open a device that was sharing an interrupt request (IRQ)
'   with other devices. At least one other device that uses that IRQ
'   was already opened.

Public Const ERROR_IRQ_BUSY = 1119&
'   A serial I/O operation was completed by another write to the serial port.
'   (The IOCTL_SERIAL_XOFF_COUNTER reached zero.)

Public Const ERROR_MORE_WRITES = 1120&
'   A serial I/O operation completed because the time-out period expired.
'   (The IOCTL_SERIAL_XOFF_COUNTER did not reach zero.)

Public Const ERROR_COUNTER_TIMEOUT = 1121&
'   No ID address mark was found on the floppy disk.

Public Const ERROR_FLOPPY_ID_MARK_NOT_FOUND = 1122&
'   Mismatch between the floppy disk sector ID field and the floppy disk
'   controller track address.

Public Const ERROR_FLOPPY_WRONG_CYLINDER = 1123&
'   The floppy disk controller reported an error that is not recognized
'   by the floppy disk driver.

Public Const ERROR_FLOPPY_UNKNOWN_ERROR = 1124&
'   The floppy disk controller returned inconsistent results in its registers.

Public Const ERROR_FLOPPY_BAD_REGISTERS = 1125&
'   While accessing the hard disk, a recalibrate operation failed, even after retries.

Public Const ERROR_DISK_RECALIBRATE_FAILED = 1126&
'   While accessing the hard disk, a disk operation failed even after retries.

Public Const ERROR_DISK_OPERATION_FAILED = 1127&
'   While accessing the hard disk, a disk controller reset was needed, but
'   even that failed.

Public Const ERROR_DISK_RESET_FAILED = 1128&
'   Physical end of tape encountered.

Public Const ERROR_EOM_OVERFLOW = 1129&
'   Not enough server storage is available to process this command.

Public Const ERROR_NOT_ENOUGH_SERVER_MEMORY = 1130&
'   A potential deadlock condition has been detected.

Public Const ERROR_POSSIBLE_DEADLOCK = 1131&
'   The base address or the file offset specified does not have the proper
'   alignment.

Public Const ERROR_MAPPED_ALIGNMENT = 1132&
' NEW for Win32

Public Const ERROR_INVALID_PIXEL_FORMAT = 2000
Public Const ERROR_BAD_DRIVER = 2001
Public Const ERROR_INVALID_WINDOW_STYLE = 2002
Public Const ERROR_METAFILE_NOT_SUPPORTED = 2003
Public Const ERROR_TRANSFORM_NOT_SUPPORTED = 2004
Public Const ERROR_CLIPPING_NOT_SUPPORTED = 2005
Public Const ERROR_UNKNOWN_PRINT_MONITOR = 3000
Public Const ERROR_PRINTER_DRIVER_IN_USE = 3001
Public Const ERROR_SPOOL_FILE_NOT_FOUND = 3002
Public Const ERROR_SPL_NO_STARTDOC = 3003
Public Const ERROR_SPL_NO_ADDJOB = 3004
Public Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED = 3005
Public Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED = 3006
Public Const ERROR_WINS_INTERNAL = 4000
Public Const ERROR_CAN_NOT_DEL_LOCAL_WINS = 4001
Public Const ERROR_STATIC_INIT = 4002
Public Const ERROR_INC_BACKUP = 4003
Public Const ERROR_FULL_BACKUP = 4004
Public Const ERROR_REC_NON_EXISTENT = 4005
Public Const ERROR_RPL_NOT_ALLOWED = 4006
Public Const SEVERITY_SUCCESS = 0
Public Const SEVERITY_ERROR = 1
Public Const FACILITY_NT_BIT = &H10000000
Public Const NOERROR = 0
Public Const E_UNEXPECTED = &H8000FFFF
Public Const E_NOTIMPL = &H80004001
Public Const E_OUTOFMEMORY = &H8007000E
Public Const E_INVALIDARG = &H80070057
Public Const E_NOINTERFACE = &H80004002
Public Const E_POINTER = &H80004003
Public Const E_HANDLE = &H80070006
Public Const E_ABORT = &H80004004
Public Const E_FAIL = &H80004005
Public Const E_ACCESSDENIED = &H80070005
Public Const CO_E_INIT_TLS = &H80004006
Public Const CO_E_INIT_SHARED_ALLOCATOR = &H80004007
Public Const CO_E_INIT_MEMORY_ALLOCATOR = &H80004008
Public Const CO_E_INIT_CLASS_CACHE = &H80004009
Public Const CO_E_INIT_RPC_CHANNEL = &H8000400A
Public Const CO_E_INIT_TLS_SET_CHANNEL_CONTROL = &H8000400B
Public Const CO_E_INIT_TLS_CHANNEL_CONTROL = &H8000400C
Public Const CO_E_INIT_UNACCEPTED_USER_ALLOCATOR = &H8000400D
Public Const CO_E_INIT_SCM_MUTEX_EXISTS = &H8000400E
Public Const CO_E_INIT_SCM_FILE_MAPPING_EXISTS = &H8000400F
Public Const CO_E_INIT_SCM_MAP_VIEW_OF_FILE = &H80004010
Public Const CO_E_INIT_SCM_EXEC_FAILURE = &H80004011
Public Const CO_E_INIT_ONLY_SINGLE_THREADED = &H80004012
Public Const S_OK = &H0
Public Const S_FALSE = &H1
Public Const OLE_E_FIRST = &H80040000
Public Const OLE_E_LAST = &H800400FF
Public Const OLE_S_FIRST = &H40000
Public Const OLE_S_LAST = &H400FF
Public Const OLE_E_OLEVERB = &H80040000
Public Const OLE_E_ADVF = &H80040001
Public Const OLE_E_ENUM_NOMORE = &H80040002
Public Const OLE_E_ADVISENOTSUPPORTED = &H80040003
Public Const OLE_E_NOCONNECTION = &H80040004
Public Const OLE_E_NOTRUNNING = &H80040005
Public Const OLE_E_NOCACHE = &H80040006
Public Const OLE_E_BLANK = &H80040007
Public Const OLE_E_CLASSDIFF = &H80040008
Public Const OLE_E_CANT_GETMONIKER = &H80040009
Public Const OLE_E_CANT_BINDTOSOURCE = &H8004000A
Public Const OLE_E_STATIC = &H8004000B
Public Const OLE_E_PROMPTSAVECANCELLED = &H8004000C
Public Const OLE_E_INVALIDRECT = &H8004000D
Public Const OLE_E_WRONGCOMPOBJ = &H8004000E
Public Const OLE_E_INVALIDHWND = &H8004000F
Public Const OLE_E_NOT_INPLACEACTIVE = &H80040010
Public Const OLE_E_CANTCONVERT = &H80040011
Public Const OLE_E_NOSTORAGE = &H80040012
Public Const DV_E_FORMATETC = &H80040064
Public Const DV_E_DVTARGETDEVICE = &H80040065
Public Const DV_E_STGMEDIUM = &H80040066
Public Const DV_E_STATDATA = &H80040067
Public Const DV_E_LINDEX = &H80040068
Public Const DV_E_TYMED = &H80040069
Public Const DV_E_CLIPFORMAT = &H8004006A
Public Const DV_E_DVASPECT = &H8004006B
Public Const DV_E_DVTARGETDEVICE_SIZE = &H8004006C
Public Const DV_E_NOIVIEWOBJECT = &H8004006D
Public Const DRAGDROP_E_FIRST = &H80040100
Public Const DRAGDROP_E_LAST = &H8004010F
Public Const DRAGDROP_S_FIRST = &H40100
Public Const DRAGDROP_S_LAST = &H4010F
Public Const DRAGDROP_E_NOTREGISTERED = &H80040100
Public Const DRAGDROP_E_ALREADYREGISTERED = &H80040101
Public Const DRAGDROP_E_INVALIDHWND = &H80040102
Public Const CLASSFACTORY_E_FIRST = &H80040110
Public Const CLASSFACTORY_E_LAST = &H8004011F
Public Const CLASSFACTORY_S_FIRST = &H40110
Public Const CLASSFACTORY_S_LAST = &H4011F
Public Const CLASS_E_NOAGGREGATION = &H80040110
Public Const CLASS_E_CLASSNOTAVAILABLE = &H80040111
Public Const MARSHAL_E_FIRST = &H80040120
Public Const MARSHAL_E_LAST = &H8004012F
Public Const MARSHAL_S_FIRST = &H40120
Public Const MARSHAL_S_LAST = &H4012F
Public Const DATA_E_FIRST = &H80040130
Public Const DATA_E_LAST = &H8004013F
Public Const DATA_S_FIRST = &H40130
Public Const DATA_S_LAST = &H4013F
Public Const VIEW_E_FIRST = &H80040140
Public Const VIEW_E_LAST = &H8004014F
Public Const VIEW_S_FIRST = &H40140
Public Const VIEW_S_LAST = &H4014F
Public Const VIEW_E_DRAW = &H80040140
Public Const REGDB_E_FIRST = &H80040150
Public Const REGDB_E_LAST = &H8004015F
Public Const REGDB_S_FIRST = &H40150
Public Const REGDB_S_LAST = &H4015F
Public Const REGDB_E_READREGDB = &H80040150
Public Const REGDB_E_WRITEREGDB = &H80040151
Public Const REGDB_E_KEYMISSING = &H80040152
Public Const REGDB_E_INVALIDVALUE = &H80040153
Public Const REGDB_E_CLASSNOTREG = &H80040154
Public Const REGDB_E_IIDNOTREG = &H80040155
Public Const CACHE_E_FIRST = &H80040170
Public Const CACHE_E_LAST = &H8004017F
Public Const CACHE_S_FIRST = &H40170
Public Const CACHE_S_LAST = &H4017F
Public Const CACHE_E_NOCACHE_UPDATED = &H80040170
Public Const OLEOBJ_E_FIRST = &H80040180
Public Const OLEOBJ_E_LAST = &H8004018F
Public Const OLEOBJ_S_FIRST = &H40180
Public Const OLEOBJ_S_LAST = &H4018F
Public Const OLEOBJ_E_NOVERBS = &H80040180
Public Const OLEOBJ_E_INVALIDVERB = &H80040181
Public Const CLIENTSITE_E_FIRST = &H80040190
Public Const CLIENTSITE_E_LAST = &H8004019F
Public Const CLIENTSITE_S_FIRST = &H40190
Public Const CLIENTSITE_S_LAST = &H4019F
Public Const INPLACE_E_NOTUNDOABLE = &H800401A0
Public Const INPLACE_E_NOTOOLSPACE = &H800401A1
Public Const INPLACE_E_FIRST = &H800401A0
Public Const INPLACE_E_LAST = &H800401AF
Public Const INPLACE_S_FIRST = &H401A0
Public Const INPLACE_S_LAST = &H401AF
Public Const ENUM_E_FIRST = &H800401B0
Public Const ENUM_E_LAST = &H800401BF
Public Const ENUM_S_FIRST = &H401B0
Public Const ENUM_S_LAST = &H401BF
Public Const CONVERT10_E_FIRST = &H800401C0
Public Const CONVERT10_E_LAST = &H800401CF
Public Const CONVERT10_S_FIRST = &H401C0
Public Const CONVERT10_S_LAST = &H401CF
Public Const CONVERT10_E_OLESTREAM_GET = &H800401C0
Public Const CONVERT10_E_OLESTREAM_PUT = &H800401C1
Public Const CONVERT10_E_OLESTREAM_FMT = &H800401C2
Public Const CONVERT10_E_OLESTREAM_BITMAP_TO_DIB = &H800401C3
Public Const CONVERT10_E_STG_FMT = &H800401C4
Public Const CONVERT10_E_STG_NO_STD_STREAM = &H800401C5
Public Const CONVERT10_E_STG_DIB_TO_BITMAP = &H800401C6
Public Const CLIPBRD_E_FIRST = &H800401D0
Public Const CLIPBRD_E_LAST = &H800401DF
Public Const CLIPBRD_S_FIRST = &H401D0
Public Const CLIPBRD_S_LAST = &H401DF
Public Const CLIPBRD_E_CANT_OPEN = &H800401D0
Public Const CLIPBRD_E_CANT_EMPTY = &H800401D1
Public Const CLIPBRD_E_CANT_SET = &H800401D2
Public Const CLIPBRD_E_BAD_DATA = &H800401D3
Public Const CLIPBRD_E_CANT_CLOSE = &H800401D4
Public Const MK_E_FIRST = &H800401E0
Public Const MK_E_LAST = &H800401EF
Public Const MK_S_FIRST = &H401E0
Public Const MK_S_LAST = &H401EF
Public Const MK_E_CONNECTMANUALLY = &H800401E0
Public Const MK_E_EXCEEDEDDEADLINE = &H800401E1
Public Const MK_E_NEEDGENERIC = &H800401E2
Public Const MK_E_UNAVAILABLE = &H800401E3
Public Const MK_E_SYNTAX = &H800401E4
Public Const MK_E_NOOBJECT = &H800401E5
Public Const MK_E_INVALIDEXTENSION = &H800401E6
Public Const MK_E_INTERMEDIATEINTERFACENOTSUPPORTED = &H800401E7
Public Const MK_E_NOTBINDABLE = &H800401E8
Public Const MK_E_NOTBOUND = &H800401E9
Public Const MK_E_CANTOPENFILE = &H800401EA
Public Const MK_E_MUSTBOTHERUSER = &H800401EB
Public Const MK_E_NOINVERSE = &H800401EC
Public Const MK_E_NOSTORAGE = &H800401ED
Public Const MK_E_NOPREFIX = &H800401EE
Public Const MK_E_ENUMERATION_FAILED = &H800401EF
Public Const CO_E_FIRST = &H800401F0
Public Const CO_E_LAST = &H800401FF
Public Const CO_S_FIRST = &H401F0
Public Const CO_S_LAST = &H401FF
Public Const CO_E_NOTINITIALIZED = &H800401F0
Public Const CO_E_ALREADYINITIALIZED = &H800401F1
Public Const CO_E_CANTDETERMINECLASS = &H800401F2
Public Const CO_E_CLASSSTRING = &H800401F3
Public Const CO_E_IIDSTRING = &H800401F4
Public Const CO_E_APPNOTFOUND = &H800401F5
Public Const CO_E_APPSINGLEUSE = &H800401F6
Public Const CO_E_ERRORINAPP = &H800401F7
Public Const CO_E_DLLNOTFOUND = &H800401F8
Public Const CO_E_ERRORINDLL = &H800401F9
Public Const CO_E_WRONGOSFORAPP = &H800401FA
Public Const CO_E_OBJNOTREG = &H800401FB
Public Const CO_E_OBJISREG = &H800401FC
Public Const CO_E_OBJNOTCONNECTED = &H800401FD
Public Const CO_E_APPDIDNTREG = &H800401FE
Public Const CO_E_RELEASED = &H800401FF
Public Const OLE_S_USEREG = &H40000
Public Const OLE_S_STATIC = &H40001
Public Const OLE_S_MAC_CLIPFORMAT = &H40002
Public Const DRAGDROP_S_DROP = &H40100
Public Const DRAGDROP_S_CANCEL = &H40101
Public Const DRAGDROP_S_USEDEFAULTCURSORS = &H40102
Public Const DATA_S_SAMEFORMATETC = &H40130
Public Const VIEW_S_ALREADY_FROZEN = &H40140
Public Const CACHE_S_FORMATETC_NOTSUPPORTED = &H40170
Public Const CACHE_S_SAMECACHE = &H40171
Public Const CACHE_S_SOMECACHES_NOTUPDATED = &H40172
Public Const OLEOBJ_S_INVALIDVERB = &H40180
Public Const OLEOBJ_S_CANNOT_DOVERB_NOW = &H40181
Public Const OLEOBJ_S_INVALIDHWND = &H40182
Public Const INPLACE_S_TRUNCATED = &H401A0
Public Const CONVERT10_S_NO_PRESENTATION = &H401C0
Public Const MK_S_REDUCED_TO_SELF = &H401E2
Public Const MK_S_ME = &H401E4
Public Const MK_S_HIM = &H401E5
Public Const MK_S_US = &H401E6
Public Const MK_S_MONIKERALREADYREGISTERED = &H401E7
Public Const CO_E_CLASS_CREATE_FAILED = &H80080001
Public Const CO_E_SCM_ERROR = &H80080002
Public Const CO_E_SCM_RPC_FAILURE = &H80080003
Public Const CO_E_BAD_PATH = &H80080004
Public Const CO_E_SERVER_EXEC_FAILURE = &H80080005
Public Const CO_E_OBJSRV_RPC_FAILURE = &H80080006
Public Const MK_E_NO_NORMALIZED = &H80080007
Public Const CO_E_SERVER_STOPPING = &H80080008
Public Const MEM_E_INVALID_ROOT = &H80080009
Public Const MEM_E_INVALID_LINK = &H80080010
Public Const MEM_E_INVALID_SIZE = &H80080011
Public Const DISP_E_UNKNOWNINTERFACE = &H80020001
Public Const DISP_E_MEMBERNOTFOUND = &H80020003
Public Const DISP_E_PARAMNOTFOUND = &H80020004
Public Const DISP_E_TYPEMISMATCH = &H80020005
Public Const DISP_E_UNKNOWNNAME = &H80020006
Public Const DISP_E_NONAMEDARGS = &H80020007
Public Const DISP_E_BADVARTYPE = &H80020008
Public Const DISP_E_EXCEPTION = &H80020009
Public Const DISP_E_OVERFLOW = &H8002000A
Public Const DISP_E_BADINDEX = &H8002000B
Public Const DISP_E_UNKNOWNLCID = &H8002000C
Public Const DISP_E_ARRAYISLOCKED = &H8002000D
Public Const DISP_E_BADPARAMCOUNT = &H8002000E
Public Const DISP_E_PARAMNOTOPTIONAL = &H8002000F
Public Const DISP_E_BADCALLEE = &H80020010
Public Const DISP_E_NOTACOLLECTION = &H80020011
Public Const TYPE_E_BUFFERTOOSMALL = &H80028016
Public Const TYPE_E_INVDATAREAD = &H80028018
Public Const TYPE_E_UNSUPFORMAT = &H80028019
Public Const TYPE_E_REGISTRYACCESS = &H8002801C
Public Const TYPE_E_LIBNOTREGISTERED = &H8002801D
Public Const TYPE_E_UNDEFINEDTYPE = &H80028027
Public Const TYPE_E_QUALIFIEDNAMEDISALLOWED = &H80028028
Public Const TYPE_E_INVALIDSTATE = &H80028029
Public Const TYPE_E_WRONGTYPEKIND = &H8002802A
Public Const TYPE_E_ELEMENTNOTFOUND = &H8002802B
Public Const TYPE_E_AMBIGUOUSNAME = &H8002802C
Public Const TYPE_E_NAMECONFLICT = &H8002802D
Public Const TYPE_E_UNKNOWNLCID = &H8002802E
Public Const TYPE_E_DLLFUNCTIONNOTFOUND = &H8002802F
Public Const TYPE_E_BADMODULEKIND = &H800288BD
Public Const TYPE_E_SIZETOOBIG = &H800288C5
Public Const TYPE_E_DUPLICATEID = &H800288C6
Public Const TYPE_E_INVALIDID = &H800288CF
Public Const TYPE_E_TYPEMISMATCH = &H80028CA0
Public Const TYPE_E_OUTOFBOUNDS = &H80028CA1
Public Const TYPE_E_IOERROR = &H80028CA2
Public Const TYPE_E_CANTCREATETMPFILE = &H80028CA3
Public Const TYPE_E_CANTLOADLIBRARY = &H80029C4A
Public Const TYPE_E_INCONSISTENTPROPFUNCS = &H80029C83
Public Const TYPE_E_CIRCULARTYPE = &H80029C84
Public Const STG_E_INVALIDFUNCTION = &H80030001
Public Const STG_E_FILENOTFOUND = &H80030002
Public Const STG_E_PATHNOTFOUND = &H80030003
Public Const STG_E_TOOMANYOPENFILES = &H80030004
Public Const STG_E_ACCESSDENIED = &H80030005
Public Const STG_E_INVALIDHANDLE = &H80030006
Public Const STG_E_INSUFFICIENTMEMORY = &H80030008
Public Const STG_E_INVALIDPOINTER = &H80030009
Public Const STG_E_NOMOREFILES = &H80030012
Public Const STG_E_DISKISWRITEPROTECTED = &H80030013
Public Const STG_E_SEEKERROR = &H80030019
Public Const STG_E_WRITEFAULT = &H8003001D
Public Const STG_E_READFAULT = &H8003001E
Public Const STG_E_SHAREVIOLATION = &H80030020
Public Const STG_E_LOCKVIOLATION = &H80030021
Public Const STG_E_FILEALREADYEXISTS = &H80030050
Public Const STG_E_INVALIDPARAMETER = &H80030057
Public Const STG_E_MEDIUMFULL = &H80030070
Public Const STG_E_ABNORMALAPIEXIT = &H800300FA
Public Const STG_E_INVALIDHEADER = &H800300FB
Public Const STG_E_INVALIDNAME = &H800300FC
Public Const STG_E_UNKNOWN = &H800300FD
Public Const STG_E_UNIMPLEMENTEDFUNCTION = &H800300FE
Public Const STG_E_INVALIDFLAG = &H800300FF
Public Const STG_E_INUSE = &H80030100
Public Const STG_E_NOTCURRENT = &H80030101
Public Const STG_E_REVERTED = &H80030102
Public Const STG_E_CANTSAVE = &H80030103
Public Const STG_E_OLDFORMAT = &H80030104
Public Const STG_E_OLDDLL = &H80030105
Public Const STG_E_SHAREREQUIRED = &H80030106
Public Const STG_E_NOTFILEBASEDSTORAGE = &H80030107
Public Const STG_E_EXTANTMARSHALLINGS = &H80030108
Public Const STG_S_CONVERTED = &H30200
Public Const RPC_E_CALL_REJECTED = &H80010001
Public Const RPC_E_CALL_CANCELED = &H80010002
Public Const RPC_E_CANTPOST_INSENDCALL = &H80010003
Public Const RPC_E_CANTCALLOUT_INASYNCCALL = &H80010004
Public Const RPC_E_CANTCALLOUT_INEXTERNALCALL = &H80010005
Public Const RPC_E_CONNECTION_TERMINATED = &H80010006
Public Const RPC_E_SERVER_DIED = &H80010007
Public Const RPC_E_CLIENT_DIED = &H80010008
Public Const RPC_E_INVALID_DATAPACKET = &H80010009
Public Const RPC_E_CANTTRANSMIT_CALL = &H8001000A
Public Const RPC_E_CLIENT_CANTMARSHAL_DATA = &H8001000B
Public Const RPC_E_CLIENT_CANTUNMARSHAL_DATA = &H8001000C
Public Const RPC_E_SERVER_CANTMARSHAL_DATA = &H8001000D
Public Const RPC_E_SERVER_CANTUNMARSHAL_DATA = &H8001000E
Public Const RPC_E_INVALID_DATA = &H8001000F
Public Const RPC_E_INVALID_PARAMETER = &H80010010
Public Const RPC_E_CANTCALLOUT_AGAIN = &H80010011
Public Const RPC_E_SERVER_DIED_DNE = &H80010012
Public Const RPC_E_SYS_CALL_FAILED = &H80010100
Public Const RPC_E_OUT_OF_RESOURCES = &H80010101
Public Const RPC_E_ATTEMPTED_MULTITHREAD = &H80010102
Public Const RPC_E_NOT_REGISTERED = &H80010103
Public Const RPC_E_FAULT = &H80010104
Public Const RPC_E_SERVERFAULT = &H80010105
Public Const RPC_E_CHANGED_MODE = &H80010106
Public Const RPC_E_INVALIDMETHOD = &H80010107
Public Const RPC_E_DISCONNECTED = &H80010108
Public Const RPC_E_RETRY = &H80010109
Public Const RPC_E_SERVERCALL_RETRYLATER = &H8001010A
Public Const RPC_E_SERVERCALL_REJECTED = &H8001010B
Public Const RPC_E_INVALID_CALLDATA = &H8001010C
Public Const RPC_E_CANTCALLOUT_ININPUTSYNCCALL = &H8001010D
Public Const RPC_E_WRONG_THREAD = &H8001010E
Public Const RPC_E_THREAD_NOT_INIT = &H8001010F
Public Const RPC_E_UNEXPECTED = &H8001FFFF
' /////////////////////////
'                        //
'  Winnet32 Status Codes //
'                        //
' /////////////////////////
'   The specified username is invalid.

Public Const ERROR_BAD_USERNAME = 2202&
'   This network connection does not exist.

Public Const ERROR_NOT_CONNECTED = 2250&
'   This network connection has files open or requests pending.

Public Const ERROR_OPEN_FILES = 2401&
'   The device is in use by an active process and cannot be disconnected.

Public Const ERROR_DEVICE_IN_USE = 2404&
'   The specified device name is invalid.

Public Const ERROR_BAD_DEVICE = 1200&
'   The device is not currently connected but it is a remembered connection.

Public Const ERROR_CONNECTION_UNAVAIL = 1201&
'   An attempt was made to remember a device that had previously been remembered.

Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
'   No network provider accepted the given network path.

Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&
'   The specified network provider name is invalid.

Public Const ERROR_BAD_PROVIDER = 1204&
'   Unable to open the network connection profile.

Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
'   The network connection profile is corrupt.

Public Const ERROR_BAD_PROFILE = 1206&
'   Cannot enumerate a non-container.

Public Const ERROR_NOT_CONTAINER = 1207&
'   An extended error has occurred.

Public Const ERROR_EXTENDED_ERROR = 1208&
'   The format of the specified group name is invalid.

Public Const ERROR_INVALID_GROUPNAME = 1209&
'   The format of the specified computer name is invalid.

Public Const ERROR_INVALID_COMPUTERNAME = 1210&
'   The format of the specified event name is invalid.

Public Const ERROR_INVALID_EVENTNAME = 1211&
'   The format of the specified domain name is invalid.

Public Const ERROR_INVALID_DOMAINNAME = 1212&
'   The format of the specified service name is invalid.

Public Const ERROR_INVALID_SERVICENAME = 1213&
'   The format of the specified network name is invalid.

Public Const ERROR_INVALID_NETNAME = 1214&
'   The format of the specified share name is invalid.

Public Const ERROR_INVALID_SHARENAME = 1215&
'   The format of the specified password is invalid.

Public Const ERROR_INVALID_PASSWORDNAME = 1216&
'   The format of the specified message name is invalid.

Public Const ERROR_INVALID_MESSAGENAME = 1217&
'   The format of the specified message destination is invalid.

Public Const ERROR_INVALID_MESSAGEDEST = 1218&
'   The credentials supplied conflict with an existing set of credentials.

Public Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&
'   An attempt was made to establish a session to a Lan Manager server, but there
'   are already too many sessions established to that server.

Public Const ERROR_REMOTE_SESSION_LIMIT_EXCEEDED = 1220&
'   The workgroup or domain name is already in use by another computer on the
'   network.

Public Const ERROR_DUP_DOMAINNAME = 1221&
'   The network is not present or not started.

Public Const ERROR_NO_NETWORK = 1222&
' /////////////////////////
'                        //
'  Security Status Codes //
'                        //
' /////////////////////////
'   Not all privileges referenced are assigned to the caller.

Public Const ERROR_NOT_ALL_ASSIGNED = 1300&
'   Some mapping between account names and security IDs was not done.

Public Const ERROR_SOME_NOT_MAPPED = 1301&
'   No system quota limits are specifically set for this account.

Public Const ERROR_NO_QUOTAS_FOR_ACCOUNT = 1302&
'   No encryption key is available.  A well-known encryption key was returned.

Public Const ERROR_LOCAL_USER_SESSION_KEY = 1303&
'   The NT password is too complex to be converted to a LAN Manager
'   password.  The LAN Manager password returned is a NULL string.

Public Const ERROR_NULL_LM_PASSWORD = 1304&
'   The revision level is unknown.

Public Const ERROR_UNKNOWN_REVISION = 1305&
'   Indicates two revision levels are incompatible.

Public Const ERROR_REVISION_MISMATCH = 1306&
'   This security ID may not be assigned as the owner of this object.

Public Const ERROR_INVALID_OWNER = 1307&
'   This security ID may not be assigned as the primary group of an object.

Public Const ERROR_INVALID_PRIMARY_GROUP = 1308&
'   An attempt has been made to operate on an impersonation token
'   by a thread that is not currently impersonating a client.

Public Const ERROR_NO_IMPERSONATION_TOKEN = 1309&
'   The group may not be disabled.

Public Const ERROR_CANT_DISABLE_MANDATORY = 1310&
'   There are currently no logon servers available to service the logon
'   request.

Public Const ERROR_NO_LOGON_SERVERS = 1311&
'    A specified logon session does not exist.  It may already have
'    been terminated.

Public Const ERROR_NO_SUCH_LOGON_SESSION = 1312&
'    A specified privilege does not exist.

Public Const ERROR_NO_SUCH_PRIVILEGE = 1313&
'    A required privilege is not held by the client.

Public Const ERROR_PRIVILEGE_NOT_HELD = 1314&
'   The name provided is not a properly formed account name.

Public Const ERROR_INVALID_ACCOUNT_NAME = 1315&
'   The specified user already exists.

Public Const ERROR_USER_EXISTS = 1316&
'   The specified user does not exist.

Public Const ERROR_NO_SUCH_USER = 1317&
'   The specified group already exists.

Public Const ERROR_GROUP_EXISTS = 1318&
'   The specified group does not exist.

Public Const ERROR_NO_SUCH_GROUP = 1319&
'   Either the specified user account is already a member of the specified
'   group, or the specified group cannot be deleted because it contains
'   a member.

Public Const ERROR_MEMBER_IN_GROUP = 1320&
'   The specified user account is not a member of the specified group account.

Public Const ERROR_MEMBER_NOT_IN_GROUP = 1321&
'   The last remaining administration account cannot be disabled
'   or deleted.

Public Const ERROR_LAST_ADMIN = 1322&
'   Unable to update the password.  The value provided as the current
'   password is incorrect.

Public Const ERROR_WRONG_PASSWORD = 1323&
'   Unable to update the password.  The value provided for the new password
'   contains values that are not allowed in passwords.

Public Const ERROR_ILL_FORMED_PASSWORD = 1324&
'   Unable to update the password because a password update rule has been
'   violated.

Public Const ERROR_PASSWORD_RESTRICTION = 1325&
'   Logon failure: unknown user name or bad password.

Public Const ERROR_LOGON_FAILURE = 1326&
'   Logon failure: user account restriction.

Public Const ERROR_ACCOUNT_RESTRICTION = 1327&
'   Logon failure: account logon time restriction violation.

Public Const ERROR_INVALID_LOGON_HOURS = 1328&
'   Logon failure: user not allowed to log on to this computer.

Public Const ERROR_INVALID_WORKSTATION = 1329&
'   Logon failure: the specified account password has expired.

Public Const ERROR_PASSWORD_EXPIRED = 1330&
'   Logon failure: account currently disabled.

Public Const ERROR_ACCOUNT_DISABLED = 1331&
'   No mapping between account names and security IDs was done.

Public Const ERROR_NONE_MAPPED = 1332&
'   Too many local user identifiers (LUIDs) were requested at one time.

Public Const ERROR_TOO_MANY_LUIDS_REQUESTED = 1333&
'   No more local user identifiers (LUIDs) are available.

Public Const ERROR_LUIDS_EXHAUSTED = 1334&
'   The subauthority part of a security ID is invalid for this particular use.

Public Const ERROR_INVALID_SUB_AUTHORITY = 1335&
'   The access control list (ACL) structure is invalid.

Public Const ERROR_INVALID_ACL = 1336&
'   The security ID structure is invalid.

Public Const ERROR_INVALID_SID = 1337&
'   The security descriptor structure is invalid.

Public Const ERROR_INVALID_SECURITY_DESCR = 1338&
'   The inherited access control list (ACL) or access control entry (ACE)
'   could not be built.

Public Const ERROR_BAD_INHERITANCE_ACL = 1340&
'   The server is currently disabled.

Public Const ERROR_SERVER_DISABLED = 1341&
'   The server is currently enabled.

Public Const ERROR_SERVER_NOT_DISABLED = 1342&
'   The value provided was an invalid value for an identifier authority.

Public Const ERROR_INVALID_ID_AUTHORITY = 1343&
'   No more memory is available for security information updates.

Public Const ERROR_ALLOTTED_SPACE_EXCEEDED = 1344&
'   The specified attributes are invalid, or incompatible with the
'   attributes for the group as a whole.

Public Const ERROR_INVALID_GROUP_ATTRIBUTES = 1345&
'   Either a required impersonation level was not provided, or the
'   provided impersonation level is invalid.

Public Const ERROR_BAD_IMPERSONATION_LEVEL = 1346&
'   Cannot open an anonymous level security token.

Public Const ERROR_CANT_OPEN_ANONYMOUS = 1347&
'   The validation information class requested was invalid.

Public Const ERROR_BAD_VALIDATION_CLASS = 1348&
'   The type of the token is inappropriate for its attempted use.

Public Const ERROR_BAD_TOKEN_TYPE = 1349&
'   Unable to perform a security operation on an object
'   which has no associated security.

Public Const ERROR_NO_SECURITY_ON_OBJECT = 1350&
'   Indicates a Windows NT Advanced Server could not be contacted or that
'   objects within the domain are protected such that necessary
'   information could not be retrieved.

Public Const ERROR_CANT_ACCESS_DOMAIN_INFO = 1351&
'   The security account manager (SAM) or local security
'   authority (LSA) server was in the wrong state to perform
'   the security operation.

Public Const ERROR_INVALID_SERVER_STATE = 1352&
'   The domain was in the wrong state to perform the security operation.

Public Const ERROR_INVALID_DOMAIN_STATE = 1353&
'   This operation is only allowed for the Primary Domain Controller of the domain.

Public Const ERROR_INVALID_DOMAIN_ROLE = 1354&
'   The specified domain did not exist.

Public Const ERROR_NO_SUCH_DOMAIN = 1355&
'   The specified domain already exists.

Public Const ERROR_DOMAIN_EXISTS = 1356&
'   An attempt was made to exceed the limit on the number of domains per server.

Public Const ERROR_DOMAIN_LIMIT_EXCEEDED = 1357&
'   Unable to complete the requested operation because of either a
'   catastrophic media failure or a data structure corruption on the disk.

Public Const ERROR_INTERNAL_DB_CORRUPTION = 1358&
'   The security account database contains an internal inconsistency.

Public Const ERROR_INTERNAL_ERROR = 1359&
'   Generic access types were contained in an access mask which should
'   already be mapped to non-generic types.

Public Const ERROR_GENERIC_NOT_MAPPED = 1360&
'   A security descriptor is not in the right format (absolute or self-relative).

Public Const ERROR_BAD_DESCRIPTOR_FORMAT = 1361&
'   The requested action is restricted for use by logon processes
'   only.  The calling process has not registered as a logon process.

Public Const ERROR_NOT_LOGON_PROCESS = 1362&
'   Cannot start a new logon session with an ID that is already in use.

Public Const ERROR_LOGON_SESSION_EXISTS = 1363&
'   A specified authentication package is unknown.

Public Const ERROR_NO_SUCH_PACKAGE = 1364&
'   The logon session is not in a state that is consistent with the
'   requested operation.

Public Const ERROR_BAD_LOGON_SESSION_STATE = 1365&
'   The logon session ID is already in use.

Public Const ERROR_LOGON_SESSION_COLLISION = 1366&
'   A logon request contained an invalid logon type value.

Public Const ERROR_INVALID_LOGON_TYPE = 1367&
'   Unable to impersonate via a named pipe until data has been read
'   from that pipe.

Public Const ERROR_CANNOT_IMPERSONATE = 1368&
'   The transaction state of a Registry subtree is incompatible with the
'   requested operation.

Public Const ERROR_RXACT_INVALID_STATE = 1369&
'   An internal security database corruption has been encountered.

Public Const ERROR_RXACT_COMMIT_FAILURE = 1370&
'   Cannot perform this operation on built-in accounts.

Public Const ERROR_SPECIAL_ACCOUNT = 1371&
'   Cannot perform this operation on this built-in special group.

Public Const ERROR_SPECIAL_GROUP = 1372&
'   Cannot perform this operation on this built-in special user.

Public Const ERROR_SPECIAL_USER = 1373&
'   The user cannot be removed from a group because the group
'   is currently the user's primary group.

Public Const ERROR_MEMBERS_PRIMARY_GROUP = 1374&
'   The token is already in use as a primary token.

Public Const ERROR_TOKEN_ALREADY_IN_USE = 1375&
'   The specified local group does not exist.

Public Const ERROR_NO_SUCH_ALIAS = 1376&
'   The specified account name is not a member of the local group.

Public Const ERROR_MEMBER_NOT_IN_ALIAS = 1377&
'   The specified account name is already a member of the local group.

Public Const ERROR_MEMBER_IN_ALIAS = 1378&
'   The specified local group already exists.

Public Const ERROR_ALIAS_EXISTS = 1379&
'   Logon failure: the user has not been granted the requested
'   logon type at this computer.

Public Const ERROR_LOGON_NOT_GRANTED = 1380&
'   The maximum number of secrets that may be stored in a single system has been
'   exceeded.

Public Const ERROR_TOO_MANY_SECRETS = 1381&
'   The length of a secret exceeds the maximum length allowed.

Public Const ERROR_SECRET_TOO_LONG = 1382&
'   The local security authority database contains an internal inconsistency.

Public Const ERROR_INTERNAL_DB_ERROR = 1383&
'   During a logon attempt, the user's security context accumulated too many
'   security IDs.

Public Const ERROR_TOO_MANY_CONTEXT_IDS = 1384&
'   Logon failure: the user has not been granted the requested logon type
'   at this computer.

Public Const ERROR_LOGON_TYPE_NOT_GRANTED = 1385&
'   A cross-encrypted password is necessary to change a user password.

Public Const ERROR_NT_CROSS_ENCRYPTION_REQUIRED = 1386&
'   A new member could not be added to a local group because the member does
'   not exist.

Public Const ERROR_NO_SUCH_MEMBER = 1387&
'   A new member could not be added to a local group because the member has the
'   wrong account type.

Public Const ERROR_INVALID_MEMBER = 1388&
'   Too many security IDs have been specified.

Public Const ERROR_TOO_MANY_SIDS = 1389&
'   A cross-encrypted password is necessary to change this user password.

Public Const ERROR_LM_CROSS_ENCRYPTION_REQUIRED = 1390&
'   Indicates an ACL contains no inheritable components

Public Const ERROR_NO_INHERITANCE = 1391&
'   The file or directory is corrupt and non-readable.

Public Const ERROR_FILE_CORRUPT = 1392&
'   The disk structure is corrupt and non-readable.

Public Const ERROR_DISK_CORRUPT = 1393&
'   There is no user session key for the specified logon session.

Public Const ERROR_NO_USER_SESSION_KEY = 1394&
'  End of security error codes
' /////////////////////////
'                        //
'  WinUser Error Codes   //
'                        //
' /////////////////////////
'   Invalid window handle.

Public Const ERROR_INVALID_WINDOW_HANDLE = 1400&
'   Invalid menu handle.

Public Const ERROR_INVALID_MENU_HANDLE = 1401&
'   Invalid cursor handle.

Public Const ERROR_INVALID_CURSOR_HANDLE = 1402&
'   Invalid accelerator table handle.

Public Const ERROR_INVALID_ACCEL_HANDLE = 1403&
'   Invalid hook handle.

Public Const ERROR_INVALID_HOOK_HANDLE = 1404&
'   Invalid handle to a multiple-window position structure.

Public Const ERROR_INVALID_DWP_HANDLE = 1405&
'   Cannot create a top-level child window.

Public Const ERROR_TLW_WITH_WSCHILD = 1406&
'   Cannot find window class.

Public Const ERROR_CANNOT_FIND_WND_CLASS = 1407&
'   Invalid window, belongs to other thread.

Public Const ERROR_WINDOW_OF_OTHER_THREAD = 1408&
'   Hot key is already registered.

Public Const ERROR_HOTKEY_ALREADY_REGISTERED = 1409&
'   Class already exists.

Public Const ERROR_CLASS_ALREADY_EXISTS = 1410&
'   Class does not exist.

Public Const ERROR_CLASS_DOES_NOT_EXIST = 1411&
'   Class still has open windows.

Public Const ERROR_CLASS_HAS_WINDOWS = 1412&
'   Invalid index.

Public Const ERROR_INVALID_INDEX = 1413&
'   Invalid icon handle.

Public Const ERROR_INVALID_ICON_HANDLE = 1414&
'   Using private DIALOG window words.

Public Const ERROR_PRIVATE_DIALOG_INDEX = 1415&
'   The listbox identifier was not found.

Public Const ERROR_LISTBOX_ID_NOT_FOUND = 1416&
'   No wildcards were found.

Public Const ERROR_NO_WILDCARD_CHARACTERS = 1417&
'   Thread does not have a clipboard open.

Public Const ERROR_CLIPBOARD_NOT_OPEN = 1418&
'   Hot key is not registered.

Public Const ERROR_HOTKEY_NOT_REGISTERED = 1419&
'   The window is not a valid dialog window.

Public Const ERROR_WINDOW_NOT_DIALOG = 1420&
'   Control ID not found.

Public Const ERROR_CONTROL_ID_NOT_FOUND = 1421&
'   Invalid message for a combo box because it does not have an edit control.

Public Const ERROR_INVALID_COMBOBOX_MESSAGE = 1422&
'   The window is not a combo box.

Public Const ERROR_WINDOW_NOT_COMBOBOX = 1423&
'   Height must be less than 256.

Public Const ERROR_INVALID_EDIT_HEIGHT = 1424&
'   Invalid device context (DC) handle.

Public Const ERROR_DC_NOT_FOUND = 1425&
'   Invalid hook procedure type.

Public Const ERROR_INVALID_HOOK_FILTER = 1426&
'   Invalid hook procedure.

Public Const ERROR_INVALID_FILTER_PROC = 1427&
'   Cannot set non-local hook without a module handle.

Public Const ERROR_HOOK_NEEDS_HMOD = 1428&
'   This hook procedure can only be set Globally.
'

Public Const ERROR_PUBLIC_ONLY_HOOK = 1429&
'   The journal hook procedure is already installed.

Public Const ERROR_JOURNAL_HOOK_SET = 1430&
'   The hook procedure is not installed.

Public Const ERROR_HOOK_NOT_INSTALLED = 1431&
'   Invalid message for single-selection listbox.

Public Const ERROR_INVALID_LB_MESSAGE = 1432&
'   LB_SETCOUNT sent to non-lazy listbox.

Public Const ERROR_SETCOUNT_ON_BAD_LB = 1433&
'   This list box does not support tab stops.

Public Const ERROR_LB_WITHOUT_TABSTOPS = 1434&
'   Cannot destroy object created by another thread.

Public Const ERROR_DESTROY_OBJECT_OF_OTHER_THREAD = 1435&
'   Child windows cannot have menus.

Public Const ERROR_CHILD_WINDOW_MENU = 1436&
'   The window does not have a system menu.

Public Const ERROR_NO_SYSTEM_MENU = 1437&
'   Invalid message box style.

Public Const ERROR_INVALID_MSGBOX_STYLE = 1438&
'   Invalid system-wide (SPI_) parameter.

Public Const ERROR_INVALID_SPI_VALUE = 1439&
'   Screen already locked.

Public Const ERROR_SCREEN_ALREADY_LOCKED = 1440&
'   All handles to windows in a multiple-window position structure must
'   have the same parent.

Public Const ERROR_HWNDS_HAVE_DIFF_PARENT = 1441&
'   The window is not a child window.

Public Const ERROR_NOT_CHILD_WINDOW = 1442&
'   Invalid GW_ command.

Public Const ERROR_INVALID_GW_COMMAND = 1443&
'   Invalid thread identifier.

Public Const ERROR_INVALID_THREAD_ID = 1444&
'   Cannot process a message from a window that is not a multiple document
'   interface (MDI) window.

Public Const ERROR_NON_MDICHILD_WINDOW = 1445&
'   Popup menu already active.

Public Const ERROR_POPUP_ALREADY_ACTIVE = 1446&
'   The window does not have scroll bars.

Public Const ERROR_NO_SCROLLBARS = 1447&
'   Scroll bar range cannot be greater than 0x7FFF.

Public Const ERROR_INVALID_SCROLLBAR_RANGE = 1448&
'   Cannot show or remove the window in the way specified.

Public Const ERROR_INVALID_SHOWWIN_COMMAND = 1449&
'  End of WinUser error codes
' /////////////////////////
'                        //
'  Eventlog Status Codes //
'                        //
' /////////////////////////
'   The event log file is corrupt.

Public Const ERROR_EVENTLOG_FILE_CORRUPT = 1500&
'   No event log file could be opened, so the event logging service did not start.

Public Const ERROR_EVENTLOG_CANT_START = 1501&
'   The event log file is full.

Public Const ERROR_LOG_FILE_FULL = 1502&
'   The event log file has changed between reads.

Public Const ERROR_EVENTLOG_FILE_CHANGED = 1503&
'  End of eventlog error codes
' /////////////////////////
'                        //
'    RPC Status Codes    //
'                        //
' /////////////////////////
'   The string binding is invalid.

Public Const RPC_S_INVALID_STRING_BINDING = 1700&
'   The binding handle is not the correct type.

Public Const RPC_S_WRONG_KIND_OF_BINDING = 1701&
'   The binding handle is invalid.

Public Const RPC_S_INVALID_BINDING = 1702&
'   The RPC protocol sequence is not supported.

Public Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&
'   The RPC protocol sequence is invalid.

Public Const RPC_S_INVALID_RPC_PROTSEQ = 1704&
'   The string universal unique identifier (UUID) is invalid.

Public Const RPC_S_INVALID_STRING_UUID = 1705&
'   The endpoint format is invalid.

Public Const RPC_S_INVALID_ENDPOINT_FORMAT = 1706&
'   The network address is invalid.

Public Const RPC_S_INVALID_NET_ADDR = 1707&
'   No endpoint was found.

Public Const RPC_S_NO_ENDPOINT_FOUND = 1708&
'   The timeout value is invalid.

Public Const RPC_S_INVALID_TIMEOUT = 1709&
'   The object universal unique identifier (UUID) was not found.

Public Const RPC_S_OBJECT_NOT_FOUND = 1710&
'   The object universal unique identifier (UUID) has already been registered.

Public Const RPC_S_ALREADY_REGISTERED = 1711&
'   The type universal unique identifier (UUID) has already been registered.

Public Const RPC_S_TYPE_ALREADY_REGISTERED = 1712&
'   The RPC server is already listening.

Public Const RPC_S_ALREADY_LISTENING = 1713&
'   No protocol sequences have been registered.

Public Const RPC_S_NO_PROTSEQS_REGISTERED = 1714&
'   The RPC server is not listening.

Public Const RPC_S_NOT_LISTENING = 1715&
'   The manager type is unknown.

Public Const RPC_S_UNKNOWN_MGR_TYPE = 1716&
'   The interface is unknown.

Public Const RPC_S_UNKNOWN_IF = 1717&
'   There are no bindings.

Public Const RPC_S_NO_BINDINGS = 1718&
'   There are no protocol sequences.

Public Const RPC_S_NO_PROTSEQS = 1719&
'   The endpoint cannot be created.

Public Const RPC_S_CANT_CREATE_ENDPOINT = 1720&
'   Not enough resources are available to complete this operation.

Public Const RPC_S_OUT_OF_RESOURCES = 1721&
'   The RPC server is unavailable.

Public Const RPC_S_SERVER_UNAVAILABLE = 1722&
'   The RPC server is too busy to complete this operation.

Public Const RPC_S_SERVER_TOO_BUSY = 1723&
'   The network options are invalid.

Public Const RPC_S_INVALID_NETWORK_OPTIONS = 1724&
'   There is not a remote procedure call active in this thread.

Public Const RPC_S_NO_CALL_ACTIVE = 1725&
'   The remote procedure call failed.

Public Const RPC_S_CALL_FAILED = 1726&
'   The remote procedure call failed and did not execute.

Public Const RPC_S_CALL_FAILED_DNE = 1727&
'   A remote procedure call (RPC) protocol error occurred.

Public Const RPC_S_PROTOCOL_ERROR = 1728&
'   The transfer syntax is not supported by the RPC server.

Public Const RPC_S_UNSUPPORTED_TRANS_SYN = 1730&
'   The universal unique identifier (UUID) type is not supported.

Public Const RPC_S_UNSUPPORTED_TYPE = 1732&
'   The tag is invalid.

Public Const RPC_S_INVALID_TAG = 1733&
'   The array bounds are invalid.

Public Const RPC_S_INVALID_BOUND = 1734&
'   The binding does not contain an entry name.

Public Const RPC_S_NO_ENTRY_NAME = 1735&
'   The name syntax is invalid.

Public Const RPC_S_INVALID_NAME_SYNTAX = 1736&
'   The name syntax is not supported.

Public Const RPC_S_UNSUPPORTED_NAME_SYNTAX = 1737&
'   No network address is available to use to construct a universal
'   unique identifier (UUID).

Public Const RPC_S_UUID_NO_ADDRESS = 1739&
'   The endpoint is a duplicate.

Public Const RPC_S_DUPLICATE_ENDPOINT = 1740&
'   The authentication type is unknown.

Public Const RPC_S_UNKNOWN_AUTHN_TYPE = 1741&
'   The maximum number of calls is too small.

Public Const RPC_S_MAX_CALLS_TOO_SMALL = 1742&
'   The string is too long.

Public Const RPC_S_STRING_TOO_LONG = 1743&
'   The RPC protocol sequence was not found.

Public Const RPC_S_PROTSEQ_NOT_FOUND = 1744&
'   The procedure number is out of range.

Public Const RPC_S_PROCNUM_OUT_OF_RANGE = 1745&
'   The binding does not contain any authentication information.

Public Const RPC_S_BINDING_HAS_NO_AUTH = 1746&
'   The authentication service is unknown.

Public Const RPC_S_UNKNOWN_AUTHN_SERVICE = 1747&
'   The authentication level is unknown.

Public Const RPC_S_UNKNOWN_AUTHN_LEVEL = 1748&
'   The security context is invalid.

Public Const RPC_S_INVALID_AUTH_IDENTITY = 1749&
'   The authorization service is unknown.

Public Const RPC_S_UNKNOWN_AUTHZ_SERVICE = 1750&
'   The entry is invalid.

Public Const EPT_S_INVALID_ENTRY = 1751&
'   The server endpoint cannot perform the operation.

Public Const EPT_S_CANT_PERFORM_OP = 1752&
'   There are no more endpoints available from the endpoint mapper.

Public Const EPT_S_NOT_REGISTERED = 1753&
'   No interfaces have been exported.

Public Const RPC_S_NOTHING_TO_EXPORT = 1754&
'   The entry name is incomplete.

Public Const RPC_S_INCOMPLETE_NAME = 1755&
'   The version option is invalid.

Public Const RPC_S_INVALID_VERS_OPTION = 1756&
'   There are no more members.

Public Const RPC_S_NO_MORE_MEMBERS = 1757&
'   There is nothing to unexport.

Public Const RPC_S_NOT_ALL_OBJS_UNEXPORTED = 1758&
'   The interface was not found.

Public Const RPC_S_INTERFACE_NOT_FOUND = 1759&
'   The entry already exists.

Public Const RPC_S_ENTRY_ALREADY_EXISTS = 1760&
'   The entry is not found.

Public Const RPC_S_ENTRY_NOT_FOUND = 1761&
'   The name service is unavailable.

Public Const RPC_S_NAME_SERVICE_UNAVAILABLE = 1762&
'   The network address family is invalid.

Public Const RPC_S_INVALID_NAF_ID = 1763&
'   The requested operation is not supported.

Public Const RPC_S_CANNOT_SUPPORT = 1764&
'   No security context is available to allow impersonation.

Public Const RPC_S_NO_CONTEXT_AVAILABLE = 1765&
'   An internal error occurred in a remote procedure call (RPC).

Public Const RPC_S_INTERNAL_ERROR = 1766&
'   The RPC server attempted an integer division by zero.'

Public Const RPC_S_ZERO_DIVIDE = 1767&
'   An addressing error occurred in the RPC server.

Public Const RPC_S_ADDRESS_ERROR = 1768&
'   A floating-point operation at the RPC server caused a division by zero.

Public Const RPC_S_FP_DIV_ZERO = 1769&
'   A floating-point underflow occurred at the RPC server.

Public Const RPC_S_FP_UNDERFLOW = 1770&
'   A floating-point overflow occurred at the RPC server.

Public Const RPC_S_FP_OVERFLOW = 1771&
'   The list of RPC servers available for the binding of auto handles
'   has been exhausted.

Public Const RPC_X_NO_MORE_ENTRIES = 1772&
'   Unable to open the character translation table file.

Public Const RPC_X_SS_CHAR_TRANS_OPEN_FAIL = 1773&
'   The file containing the character translation table has fewer than
'   512 bytes.

Public Const RPC_X_SS_CHAR_TRANS_SHORT_FILE = 1774&
'   A null context handle was passed from the client to the host during
'   a remote procedure call.

Public Const RPC_X_SS_IN_NULL_CONTEXT = 1775&
'   The context handle changed during a remote procedure call.

Public Const RPC_X_SS_CONTEXT_DAMAGED = 1777&
'   The binding handles passed to a remote procedure call do not match.

Public Const RPC_X_SS_HANDLES_MISMATCH = 1778&
'   The stub is unable to get the remote procedure call handle.

Public Const RPC_X_SS_CANNOT_GET_CALL_HANDLE = 1779&
'   A null reference pointer was passed to the stub.

Public Const RPC_X_NULL_REF_POINTER = 1780&
'   The enumeration value is out of range.

Public Const RPC_X_ENUM_VALUE_OUT_OF_RANGE = 1781&
'   The byte count is too small.

Public Const RPC_X_BYTE_COUNT_TOO_SMALL = 1782&
'   The stub received bad data.

Public Const RPC_X_BAD_STUB_DATA = 1783&
'   The supplied user buffer is not valid for the requested operation.

Public Const ERROR_INVALID_USER_BUFFER = 1784&
'   The disk media is not recognized.  It may not be formatted.

Public Const ERROR_UNRECOGNIZED_MEDIA = 1785&
'   The workstation does not have a trust secret.

Public Const ERROR_NO_TRUST_LSA_SECRET = 1786&
'   The SAM database on the Windows NT Advanced Server does not have a computer
'   account for this workstation trust relationship.

Public Const ERROR_NO_TRUST_SAM_ACCOUNT = 1787&
'   The trust relationship between the primary domain and the trusted
'   domain failed.

Public Const ERROR_TRUSTED_DOMAIN_FAILURE = 1788&
'   The trust relationship between this workstation and the primary
'   domain failed.

Public Const ERROR_TRUSTED_RELATIONSHIP_FAILURE = 1789&
'   The network logon failed.

Public Const ERROR_TRUST_FAILURE = 1790&
'   A remote procedure call is already in progress for this thread.

Public Const RPC_S_CALL_IN_PROGRESS = 1791&
'   An attempt was made to logon, but the network logon service was not started.

Public Const ERROR_NETLOGON_NOT_STARTED = 1792&
'   The user's account has expired.

Public Const ERROR_ACCOUNT_EXPIRED = 1793&
'   The redirector is in use and cannot be unloaded.

Public Const ERROR_REDIRECTOR_HAS_OPEN_HANDLES = 1794&
'   The specified printer driver is already installed.

Public Const ERROR_PRINTER_DRIVER_ALREADY_INSTALLED = 1795&
'   The specified port is unknown.

Public Const ERROR_UNKNOWN_PORT = 1796&
'   The printer driver is unknown.

Public Const ERROR_UNKNOWN_PRINTER_DRIVER = 1797&
'   The print processor is unknown.
'

Public Const ERROR_UNKNOWN_PRINTPROCESSOR = 1798&
'   The specified separator file is invalid.

Public Const ERROR_INVALID_SEPARATOR_FILE = 1799&
'   The specified priority is invalid.

Public Const ERROR_INVALID_PRIORITY = 1800&
'   The printer name is invalid.

Public Const ERROR_INVALID_PRINTER_NAME = 1801&
'   The printer already exists.

Public Const ERROR_PRINTER_ALREADY_EXISTS = 1802&
'   The printer command is invalid.

Public Const ERROR_INVALID_PRINTER_COMMAND = 1803&
'   The specified datatype is invalid.

Public Const ERROR_INVALID_DATATYPE = 1804&
'   The Environment specified is invalid.

Public Const ERROR_INVALID_ENVIRONMENT = 1805&
'   There are no more bindings.

Public Const RPC_S_NO_MORE_BINDINGS = 1806&
'   The account used is an interdomain trust account.  Use your Global user account or local user account to access this server.

Public Const ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT = 1807&
'   The account used is a Computer Account.  Use your Global user account or local user account to access this server.

Public Const ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT = 1808&
'   The account used is an server trust account.  Use your Global user account or local user account to access this server.

Public Const ERROR_NOLOGON_SERVER_TRUST_ACCOUNT = 1809&
'   The name or security ID (SID) of the domain specified is inconsistent
'   with the trust information for that domain.

Public Const ERROR_DOMAIN_TRUST_INCONSISTENT = 1810&
'   The server is in use and cannot be unloaded.

Public Const ERROR_SERVER_HAS_OPEN_HANDLES = 1811&
'   The specified image file did not contain a resource section.

Public Const ERROR_RESOURCE_DATA_NOT_FOUND = 1812&
'   The specified resource type can not be found in the image file.

Public Const ERROR_RESOURCE_TYPE_NOT_FOUND = 1813&
'   The specified resource name can not be found in the image file.

Public Const ERROR_RESOURCE_NAME_NOT_FOUND = 1814&
'   The specified resource language ID cannot be found in the image file.

Public Const ERROR_RESOURCE_LANG_NOT_FOUND = 1815&
'   Not enough quota is available to process this command.

Public Const ERROR_NOT_ENOUGH_QUOTA = 1816&
'   The group member was not found.

Public Const RPC_S_GROUP_MEMBER_NOT_FOUND = 1898&
'   The endpoint mapper database could not be created.

Public Const EPT_S_CANT_CREATE = 1899&
'   The object universal unique identifier (UUID) is the nil UUID.

Public Const RPC_S_INVALID_OBJECT = 1900&
'   The specified time is invalid.

Public Const ERROR_INVALID_TIME = 1901&
'   The specified Form name is invalid.

Public Const ERROR_INVALID_FORM_NAME = 1902&
'   The specified Form size is invalid

Public Const ERROR_INVALID_FORM_SIZE = 1903&
'   The specified Printer handle is already being waited on

Public Const ERROR_ALREADY_WAITING = 1904&
'   The specified Printer has been deleted

Public Const ERROR_PRINTER_DELETED = 1905&
'   The state of the Printer is invalid

Public Const ERROR_INVALID_PRINTER_STATE = 1906&
'   The list of servers for this workgroup is not currently available

Public Const ERROR_NO_BROWSER_SERVERS_FOUND = 6118&
' -------------------------
'  MMSystem Section
' -------------------------
' This section defines all the support for Multimedia applications
'  general constants

Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Public Const MAXERRORLENGTH = 128  '  max error text length (including final NULL)
'  values for wType field in MMTIME struct

Public Const TIME_MS = &H1     '  time in Milliseconds
Public Const TIME_SAMPLES = &H2     '  number of wave samples
Public Const TIME_BYTES = &H4     '  current byte offset
Public Const TIME_SMPTE = &H8     '  SMPTE time
Public Const TIME_MIDI = &H10    '  MIDI time
'  Multimedia Window Messages

Public Const MM_JOY1MOVE = &H3A0  '  joystick
Public Const MM_JOY2MOVE = &H3A1
Public Const MM_JOY1ZMOVE = &H3A2
Public Const MM_JOY2ZMOVE = &H3A3
Public Const MM_JOY1BUTTONDOWN = &H3B5
Public Const MM_JOY2BUTTONDOWN = &H3B6
Public Const MM_JOY1BUTTONUP = &H3B7
Public Const MM_JOY2BUTTONUP = &H3B8
Public Const MM_MCINOTIFY = &H3B9  '  MCI
Public Const MM_MCISYSTEM_STRING = &H3CA
Public Const MM_WOM_OPEN = &H3BB  '  waveform output
Public Const MM_WOM_CLOSE = &H3BC
Public Const MM_WOM_DONE = &H3BD
Public Const MM_WIM_OPEN = &H3BE  '  waveform input
Public Const MM_WIM_CLOSE = &H3BF
Public Const MM_WIM_DATA = &H3C0
Public Const MM_MIM_OPEN = &H3C1  '  MIDI input
Public Const MM_MIM_CLOSE = &H3C2
Public Const MM_MIM_DATA = &H3C3
Public Const MM_MIM_LONGDATA = &H3C4
Public Const MM_MIM_ERROR = &H3C5
Public Const MM_MIM_LONGERROR = &H3C6
Public Const MM_MOM_OPEN = &H3C7  '  MIDI output
Public Const MM_MOM_CLOSE = &H3C8
Public Const MM_MOM_DONE = &H3C9
' String resource number bases (internal use)

Public Const MMSYSERR_BASE = 0
Public Const WAVERR_BASE = 32
Public Const MIDIERR_BASE = 64
Public Const TIMERR_BASE = 96   '  was 128, changed to match Win 31 Sonic
Public Const JOYERR_BASE = 160
Public Const MCIERR_BASE = 256
Public Const MCI_STRING_OFFSET = 512  '  if this number is changed you MUST
                                    '  alter the MCI_DEVTYPE_... list below

Public Const MCI_VD_OFFSET = 1024
Public Const MCI_CD_OFFSET = 1088
Public Const MCI_WAVE_OFFSET = 1152
Public Const MCI_SEQ_OFFSET = 1216
' General error return values

Public Const MMSYSERR_NOERROR = 0  '  no error
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)  '  unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '  device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)  '  driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)  '  device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '  device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)  '  memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  '  function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)  '  error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10) '  invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11) '  invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12) '  handle being used
                                                   '  simultaneously on another
                                                   '  thread (eg callback)

Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) '  "Specified alias not found in WIN.INI
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13) '  last error in range
Public Const MM_MOM_POSITIONCB = &H3CA              '  Callback for MEVT_POSITIONCB
Public Const MM_MCISIGNAL = &H3CB
Public Const MM_MIM_MOREDATA = &H3CC                '  MIM_DONE w/ pending events
Public Const MIDICAPS_STREAM = &H8               '  driver supports midiStreamOut directly
'  Type codes which go in the high byte of the event DWORD of a stream buffer
'  Type codes 00-7F contain parameters within the low 24 bits
'  Type codes 80-FF contain a length of their parameter in the low 24
'  bits, followed by their parameter data in the buffer. The event
'  DWORD contains the exact byte length; the parm data itself must be
'  padded to be an even multiple of 4 Byte long.
'

Public Const MEVT_F_SHORT = &H0&
Public Const MEVT_F_LONG = &H80000000
Public Const MEVT_F_CALLBACK = &H40000000
Public Const MIDISTRM_ERROR = -2
'
'  Structures and defines for midiStreamProperty
'

Public Const MIDIPROP_SET = &H80000000
Public Const MIDIPROP_GET = &H40000000
'  These are intentionally both non-zero so the app cannot accidentally
'  leave the operation off and happen to appear to work due to default
'  action.

Public Const MIDIPROP_TIMEDIV = &H1&
Public Const MIDIPROP_TEMPO = &H2&
'  MIDI function prototypes *
' ***************************************************************************
'                             Mixer Support
' **************************************************************************

Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_LONG_NAME_CHARS = 64
'
'   MMRESULT error return values specific to the mixer API
'
'

Public Const MIXERR_BASE = 1024
Public Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Public Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Public Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Public Const MIXERR_LASTERROR = (MIXERR_BASE + 2)
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_AUX = &H50000000
'   MIXERLINE.fdwLine

Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000
'   MIXERLINE.dwComponentType
'   component types for destinations and sources

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
'
'   MIXERLINE.Target.dwType
'
'

Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_AUX = 5
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
'
'   MIXERCONTROL.fdwControl

Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
'   MIXERCONTROL_CONTROLTYPE_xxx building block defines

Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000 '  in 10ths
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000 '  in 10ths
'
'   Commonly used control types for specifying MIXERCONTROL.dwControlType
'

Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
'  constants used with JOYINFOEX

Public Const JOY_BUTTON5 = &H10&
Public Const JOY_BUTTON6 = &H20&
Public Const JOY_BUTTON7 = &H40&
Public Const JOY_BUTTON8 = &H80&
Public Const JOY_BUTTON9 = &H100&
Public Const JOY_BUTTON10 = &H200&
Public Const JOY_BUTTON11 = &H400&
Public Const JOY_BUTTON12 = &H800&
Public Const JOY_BUTTON13 = &H1000&
Public Const JOY_BUTTON14 = &H2000&
Public Const JOY_BUTTON15 = &H4000&
Public Const JOY_BUTTON16 = &H8000&
Public Const JOY_BUTTON17 = &H10000
Public Const JOY_BUTTON18 = &H20000
Public Const JOY_BUTTON19 = &H40000
Public Const JOY_BUTTON20 = &H80000
Public Const JOY_BUTTON21 = &H100000
Public Const JOY_BUTTON22 = &H200000
Public Const JOY_BUTTON23 = &H400000
Public Const JOY_BUTTON24 = &H800000
Public Const JOY_BUTTON25 = &H1000000
Public Const JOY_BUTTON26 = &H2000000
Public Const JOY_BUTTON27 = &H4000000
Public Const JOY_BUTTON28 = &H8000000
Public Const JOY_BUTTON29 = &H10000000
Public Const JOY_BUTTON30 = &H20000000
Public Const JOY_BUTTON31 = &H40000000
Public Const JOY_BUTTON32 = &H80000000
'  constants used with JOYINFOEX structure

Public Const JOY_POVCENTERED = -1
Public Const JOY_POVFORWARD = 0
Public Const JOY_POVRIGHT = 9000
Public Const JOY_POVBACKWARD = 18000
Public Const JOY_POVLEFT = 27000
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10                             '  axis 5
Public Const JOY_RETURNV = &H20                             '  axis 6
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNRAWDATA = &H100&
Public Const JOY_RETURNPOVCTS = &H200&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const JOY_CAL_READALWAYS = &H10000
Public Const JOY_CAL_READXYONLY = &H20000
Public Const JOY_CAL_READ3 = &H40000
Public Const JOY_CAL_READ4 = &H80000
Public Const JOY_CAL_READXONLY = &H100000
Public Const JOY_CAL_READYONLY = &H200000
Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READRONLY = &H2000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000
Public Const WAVE_FORMAT_QUERY = &H1
Public Const SND_PURGE = &H40               '  purge non-static events for task
Public Const SND_APPLICATION = &H80         '  look for application specific association
Public Const WAVE_MAPPED = &H4
Public Const WAVE_FORMAT_DIRECT = &H8
Public Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Public Const MIM_MOREDATA = MM_MIM_MOREDATA
Public Const MOM_POSITIONCB = MM_MOM_POSITIONCB
'  flags for dwFlags parm of midiInOpen()

Public Const MIDI_IO_STATUS = &H20&
' Installable driver support
' Driver messages

Public Const DRV_LOAD = &H1
Public Const DRV_ENABLE = &H2
Public Const DRV_OPEN = &H3
Public Const DRV_CLOSE = &H4
Public Const DRV_DISABLE = &H5
Public Const DRV_FREE = &H6
Public Const DRV_CONFIGURE = &H7
Public Const DRV_QUERYCONFIGURE = &H8
Public Const DRV_INSTALL = &H9
Public Const DRV_REMOVE = &HA
Public Const DRV_EXITSESSION = &HB
Public Const DRV_POWER = &HF
Public Const DRV_RESERVED = &H800
Public Const DRV_USER = &H4000
' Supported return values for DRV_CONFIGURE message

Public Const DRVCNF_CANCEL = &H0
Public Const DRVCNF_OK = &H1
Public Const DRVCNF_RESTART = &H2
'  return values from DriverProc() function

Public Const DRV_CANCEL = DRVCNF_CANCEL
Public Const DRV_OK = DRVCNF_OK
Public Const DRV_RESTART = DRVCNF_RESTART
Public Const DRV_MCI_FIRST = DRV_RESERVED
Public Const DRV_MCI_LAST = DRV_RESERVED + &HFFF
' Driver callback support
'  flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and
'  midiOutOpen() to specify the type of the dwCallback parameter.

Public Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Public Const CALLBACK_NULL = &H0        '  no callback
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Public Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC
'  manufacturer IDs

Public Const MM_MICROSOFT = 1  '  Microsoft Corp.
'  product IDs

Public Const MM_MIDI_MAPPER = 1  '  MIDI Mapper
Public Const MM_WAVE_MAPPER = 2  '  Wave Mapper
Public Const MM_SNDBLST_MIDIOUT = 3  '  Sound Blaster MIDI output port
Public Const MM_SNDBLST_MIDIIN = 4  '  Sound Blaster MIDI input port
Public Const MM_SNDBLST_SYNTH = 5  '  Sound Blaster internal synthesizer
Public Const MM_SNDBLST_WAVEOUT = 6  '  Sound Blaster waveform output
Public Const MM_SNDBLST_WAVEIN = 7  '  Sound Blaster waveform input
Public Const MM_ADLIB = 9  '  Ad Lib-compatible synthesizer
Public Const MM_MPU401_MIDIOUT = 10  '  MPU401-compatible MIDI output port
Public Const MM_MPU401_MIDIIN = 11  '  MPU401-compatible MIDI input port
Public Const MM_PC_JOYSTICK = 12  '  Joystick adapter
'  flag values for uFlags parameter

Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
                                    '  this range will raise an error

Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007
'  waveform audio error return values

Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)    '  unsupported wave format
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  header not prepared
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)    '  device is synchronous
Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)    '  last error in range
'  wave callback messages

Public Const WOM_OPEN = MM_WOM_OPEN
Public Const WOM_CLOSE = MM_WOM_CLOSE
Public Const WOM_DONE = MM_WOM_DONE
Public Const WIM_OPEN = MM_WIM_OPEN
Public Const WIM_CLOSE = MM_WIM_CLOSE
Public Const WIM_DATA = MM_WIM_DATA
'  device ID for wave device mapper

Public Const WAVE_MAPPER = -1&
'  flags for dwFlags parameter in waveOutOpen() and waveInOpen()

Public Const WAVE_ALLOWSYNC = &H2
Public Const WAVE_VALID = &H3         '  ;Internal
'  flags for dwFlags field of WAVEHDR

Public Const WHDR_DONE = &H1         '  done bit
Public Const WHDR_PREPARED = &H2         '  set if this header has been prepared
Public Const WHDR_BEGINLOOP = &H4         '  loop start block
Public Const WHDR_ENDLOOP = &H8         '  loop end block
Public Const WHDR_INQUEUE = &H10        '  reserved for driver
Public Const WHDR_VALID = &H1F        '  valid flags      / ;Internal /
'  flags for dwSupport field of WAVEOUTCAPS

Public Const WAVECAPS_PITCH = &H1         '  supports pitch control
Public Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control
Public Const WAVECAPS_VOLUME = &H4         '  supports volume control
Public Const WAVECAPS_LRVOLUME = &H8         '  separate left-right volume control
Public Const WAVECAPS_SYNC = &H10
'  defines for dwFormat field of WAVEINCAPS and WAVEOUTCAPS

Public Const WAVE_INVALIDFORMAT = &H0              '  invalid format
Public Const WAVE_FORMAT_1M08 = &H1              '  11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1S08 = &H2              '  11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1M16 = &H4              '  11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S16 = &H8              '  11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10             '  22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2S08 = &H20             '  22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2M16 = &H40             '  22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S16 = &H80             '  22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100            '  44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4S08 = &H200            '  44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4M16 = &H400            '  44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S16 = &H800            '  44.1   kHz, Stereo, 16-bit
'  flags for wFormatTag field of WAVEFORMAT

Public Const WAVE_FORMAT_PCM = 1  '  Needed in resource files so outside #ifndef RC_INVOKED
'  MIDI error return values

Public Const MIDIERR_UNPREPARED = (MIDIERR_BASE + 0)   '  header not prepared
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)   '  still something playing
Public Const MIDIERR_NOMAP = (MIDIERR_BASE + 2)   '  no current map
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)   '  hardware is still busy
Public Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)   '  port no longer connected
Public Const MIDIERR_INVALIDSETUP = (MIDIERR_BASE + 5)   '  invalid setup
Public Const MIDIERR_LASTERROR = (MIDIERR_BASE + 5)   '  last error in range
'  MIDI callback messages

Public Const MIM_OPEN = MM_MIM_OPEN
Public Const MIM_CLOSE = MM_MIM_CLOSE
Public Const MIM_DATA = MM_MIM_DATA
Public Const MIM_LONGDATA = MM_MIM_LONGDATA
Public Const MIM_ERROR = MM_MIM_ERROR
Public Const MIM_LONGERROR = MM_MIM_LONGERROR
Public Const MOM_OPEN = MM_MOM_OPEN
Public Const MOM_CLOSE = MM_MOM_CLOSE
Public Const MOM_DONE = MM_MOM_DONE
'  device ID for MIDI mapper

Public Const MIDIMAPPER = (-1)  '  Cannot be cast to DWORD as RC complains
Public Const MIDI_MAPPER = -1&
'  flags for wFlags parm of midiOutCachePatches(), midiOutCacheDrumPatches()

Public Const MIDI_CACHE_ALL = 1
Public Const MIDI_CACHE_BESTFIT = 2
Public Const MIDI_CACHE_QUERY = 3
Public Const MIDI_UNCACHE = 4
Public Const MIDI_CACHE_VALID = (MIDI_CACHE_ALL Or MIDI_CACHE_BESTFIT Or MIDI_CACHE_QUERY Or MIDI_UNCACHE)  '  ;Internal
'  flags for wTechnology field of MIDIOUTCAPS structure

Public Const MOD_MIDIPORT = 1  '  output port
Public Const MOD_SYNTH = 2  '  generic internal synth
Public Const MOD_SQSYNTH = 3  '  square wave internal synth
Public Const MOD_FMSYNTH = 4  '  FM internal synth
Public Const MOD_MAPPER = 5  '  MIDI mapper
'  flags for dwSupport field of MIDIOUTCAPS

Public Const MIDICAPS_VOLUME = &H1         '  supports volume control
Public Const MIDICAPS_LRVOLUME = &H2         '  separate left-right volume control
Public Const MIDICAPS_CACHE = &H4
'  flags for dwFlags field of MIDIHDR structure

Public Const MHDR_DONE = &H1         '  done bit
Public Const MHDR_PREPARED = &H2         '  set if header prepared
Public Const MHDR_INQUEUE = &H4         '  reserved for driver
Public Const MHDR_VALID = &H7         '  valid flags / ;Internal /
'  device ID for aux device mapper

Public Const AUX_MAPPER = -1&
'  flags for wTechnology field in AUXCAPS structure

Public Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Public Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks
'  flags for dwSupport field in AUXCAPS structure

Public Const AUXCAPS_VOLUME = &H1         '  supports volume control
Public Const AUXCAPS_LRVOLUME = &H2         '  separate left-right volume control
'  timer error return values

Public Const TIMERR_NOERROR = (0)  '  no error
Public Const TIMERR_NOCANDO = (TIMERR_BASE + 1) '  request not completed
Public Const TIMERR_STRUCT = (TIMERR_BASE + 33) '  time struct size
'  flags for wFlags parameter of timeSetEvent() function

Public Const TIME_ONESHOT = 0  '  program timer for single event
Public Const TIME_PERIODIC = 1  '  program for continuous periodic event
'  joystick error return values

Public Const JOYERR_NOERROR = (0)  '  no error
Public Const JOYERR_PARMS = (JOYERR_BASE + 5) '  bad parameters
Public Const JOYERR_NOCANDO = (JOYERR_BASE + 6) '  request not completed
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7) '  joystick is unplugged
'  constants used with JOYINFO structure and MM_JOY messages

Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOY_BUTTON1CHG = &H100
Public Const JOY_BUTTON2CHG = &H200
Public Const JOY_BUTTON3CHG = &H400
Public Const JOY_BUTTON4CHG = &H800
'  joystick ID constants

Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
'  MMIO error return values

Public Const MMIOERR_BASE = 256
Public Const MMIOERR_FILENOTFOUND = (MMIOERR_BASE + 1)  '  file not found
Public Const MMIOERR_OUTOFMEMORY = (MMIOERR_BASE + 2)  '  out of memory
Public Const MMIOERR_CANNOTOPEN = (MMIOERR_BASE + 3)  '  cannot open
Public Const MMIOERR_CANNOTCLOSE = (MMIOERR_BASE + 4)  '  cannot close
Public Const MMIOERR_CANNOTREAD = (MMIOERR_BASE + 5)  '  cannot read
Public Const MMIOERR_CANNOTWRITE = (MMIOERR_BASE + 6) '  cannot write
Public Const MMIOERR_CANNOTSEEK = (MMIOERR_BASE + 7)  '  cannot seek
Public Const MMIOERR_CANNOTEXPAND = (MMIOERR_BASE + 8)  '  cannot expand file
Public Const MMIOERR_CHUNKNOTFOUND = (MMIOERR_BASE + 9)  '  chunk not found
Public Const MMIOERR_UNBUFFERED = (MMIOERR_BASE + 10) '  file is unbuffered
'  MMIO constants

Public Const CFSEPCHAR = "+"  '  compound file name separator char.
Public Const MMIO_RWMODE = &H3         '  mask to get bits used for opening
                                        '  file for reading/writing/both

Public Const MMIO_SHAREMODE = &H70        '  file sharing mode number
'  constants for dwFlags field of MMIOINFO

Public Const MMIO_CREATE = &H1000      '  create new file (or truncate file)
Public Const MMIO_PARSE = &H100       '  parse new file returning path
Public Const MMIO_DELETE = &H200       '  create new file (or truncate file)
Public Const MMIO_EXIST = &H4000      '  checks for existence of file
Public Const MMIO_ALLOCBUF = &H10000     '  mmioOpen() should allocate a buffer
Public Const MMIO_GETTEMP = &H20000     '  mmioOpen() should retrieve temp name
Public Const MMIO_DIRTY = &H10000000  '  I/O buffer is dirty
'  MMIO_DIRTY is also used in the <dwFlags> field of MMCKINFO structure

Public Const MMIO_OPEN_VALID = &H3FFFF     '  valid flags for mmioOpen / ;Internal /
'  read/write mode numbers (bit field MMIO_RWMODE)

Public Const MMIO_READ = &H0         '  open file for reading only
Public Const MMIO_WRITE = &H1         '  open file for writing only
Public Const MMIO_READWRITE = &H2         '  open file for reading and writing
'  share mode numbers (bit field MMIO_SHAREMODE)

Public Const MMIO_COMPAT = &H0         '  compatibility mode
Public Const MMIO_EXCLUSIVE = &H10        '  exclusive-access mode
Public Const MMIO_DENYWRITE = &H20        '  deny writing to other processes
Public Const MMIO_DENYREAD = &H30        '  deny reading to other processes
Public Const MMIO_DENYNONE = &H40        '  deny nothing to other processes
'  flags for other functions

Public Const MMIO_FHOPEN = &H10    '  mmioClose(): keep file handle open
Public Const MMIO_EMPTYBUF = &H10    '  mmioFlush(): empty the I/O buffer
Public Const MMIO_TOUPPER = &H10    '  mmioStringToFOURCC(): cvt. to u-case
Public Const MMIO_INSTALLPROC = &H10000     '  mmioInstallIOProc(): install MMIOProc
Public Const MMIO_PUBLICPROC = &H10000000  '  mmioInstallIOProc: install Globally
Public Const MMIO_UNICODEPROC = &H1000000   '  mmioInstallIOProc(): Unicode MMIOProc
Public Const MMIO_REMOVEPROC = &H20000     '  mmioInstallIOProc(): remove MMIOProc
Public Const MMIO_FINDPROC = &H40000     '  mmioInstallIOProc(): find an MMIOProc
Public Const MMIO_FINDCHUNK = &H10    '  mmioDescend(): find a chunk by ID
Public Const MMIO_FINDRIFF = &H20    '  mmioDescend(): find a LIST chunk
Public Const MMIO_FINDLIST = &H40    '  mmioDescend(): find a RIFF chunk
Public Const MMIO_CREATERIFF = &H20    '  mmioCreateChunk(): make a LIST chunk
Public Const MMIO_CREATELIST = &H40    '  mmioCreateChunk(): make a RIFF chunk
Public Const MMIO_VALIDPROC = &H11070000  '  valid for mmioInstallIOProc / ;Internal /
'  message numbers for MMIOPROC I/O procedure functions

Public Const MMIOM_READ = MMIO_READ  '  read (must equal MMIO_READ!)
Public Const MMIOM_WRITE = MMIO_WRITE  '  write (must equal MMIO_WRITE!)
Public Const MMIOM_SEEK = 2  '  seek to a new position in file
Public Const MMIOM_OPEN = 3  '  open file
Public Const MMIOM_CLOSE = 4  '  close file
Public Const MMIOM_WRITEFLUSH = 5  '  write and flush
Public Const MMIOM_RENAME = 6  '  rename specified file
Public Const MMIOM_USER = &H8000  '  beginning of user-defined messages
'  flags for mmioSeek()

Public Const SEEK_SET = 0  '  seek to an absolute position
Public Const SEEK_CUR = 1  '  seek relative to current position
Public Const SEEK_END = 2  '  seek relative to end of file
'  other constants

Public Const MMIO_DEFAULTBUFFER = 8192  '  default buffer size
'   MCI error return values

Public Const MCIERR_INVALID_DEVICE_ID = (MCIERR_BASE + 1)
Public Const MCIERR_UNRECOGNIZED_KEYWORD = (MCIERR_BASE + 3)
Public Const MCIERR_UNRECOGNIZED_COMMAND = (MCIERR_BASE + 5)
Public Const MCIERR_HARDWARE = (MCIERR_BASE + 6)
Public Const MCIERR_INVALID_DEVICE_NAME = (MCIERR_BASE + 7)
Public Const MCIERR_OUT_OF_MEMORY = (MCIERR_BASE + 8)
Public Const MCIERR_DEVICE_OPEN = (MCIERR_BASE + 9)
Public Const MCIERR_CANNOT_LOAD_DRIVER = (MCIERR_BASE + 10)
Public Const MCIERR_MISSING_COMMAND_STRING = (MCIERR_BASE + 11)
Public Const MCIERR_PARAM_OVERFLOW = (MCIERR_BASE + 12)
Public Const MCIERR_MISSING_STRING_ARGUMENT = (MCIERR_BASE + 13)
Public Const MCIERR_BAD_INTEGER = (MCIERR_BASE + 14)
Public Const MCIERR_PARSER_INTERNAL = (MCIERR_BASE + 15)
Public Const MCIERR_DRIVER_INTERNAL = (MCIERR_BASE + 16)
Public Const MCIERR_MISSING_PARAMETER = (MCIERR_BASE + 17)
Public Const MCIERR_UNSUPPORTED_FUNCTION = (MCIERR_BASE + 18)
Public Const MCIERR_FILE_NOT_FOUND = (MCIERR_BASE + 19)
Public Const MCIERR_DEVICE_NOT_READY = (MCIERR_BASE + 20)
Public Const MCIERR_INTERNAL = (MCIERR_BASE + 21)
Public Const MCIERR_DRIVER = (MCIERR_BASE + 22)
Public Const MCIERR_CANNOT_USE_ALL = (MCIERR_BASE + 23)
Public Const MCIERR_MULTIPLE = (MCIERR_BASE + 24)
Public Const MCIERR_EXTENSION_NOT_FOUND = (MCIERR_BASE + 25)
Public Const MCIERR_OUTOFRANGE = (MCIERR_BASE + 26)
Public Const MCIERR_FLAGS_NOT_COMPATIBLE = (MCIERR_BASE + 28)
Public Const MCIERR_FILE_NOT_SAVED = (MCIERR_BASE + 30)
Public Const MCIERR_DEVICE_TYPE_REQUIRED = (MCIERR_BASE + 31)
Public Const MCIERR_DEVICE_LOCKED = (MCIERR_BASE + 32)
Public Const MCIERR_DUPLICATE_ALIAS = (MCIERR_BASE + 33)
Public Const MCIERR_BAD_CONSTANT = (MCIERR_BASE + 34)
Public Const MCIERR_MUST_USE_SHAREABLE = (MCIERR_BASE + 35)
Public Const MCIERR_MISSING_DEVICE_NAME = (MCIERR_BASE + 36)
Public Const MCIERR_BAD_TIME_FORMAT = (MCIERR_BASE + 37)
Public Const MCIERR_NO_CLOSING_QUOTE = (MCIERR_BASE + 38)
Public Const MCIERR_DUPLICATE_FLAGS = (MCIERR_BASE + 39)
Public Const MCIERR_INVALID_FILE = (MCIERR_BASE + 40)
Public Const MCIERR_NULL_PARAMETER_BLOCK = (MCIERR_BASE + 41)
Public Const MCIERR_UNNAMED_RESOURCE = (MCIERR_BASE + 42)
Public Const MCIERR_NEW_REQUIRES_ALIAS = (MCIERR_BASE + 43)
Public Const MCIERR_NOTIFY_ON_AUTO_OPEN = (MCIERR_BASE + 44)
Public Const MCIERR_NO_ELEMENT_ALLOWED = (MCIERR_BASE + 45)
Public Const MCIERR_NONAPPLICABLE_FUNCTION = (MCIERR_BASE + 46)
Public Const MCIERR_ILLEGAL_FOR_AUTO_OPEN = (MCIERR_BASE + 47)
Public Const MCIERR_FILENAME_REQUIRED = (MCIERR_BASE + 48)
Public Const MCIERR_EXTRA_CHARACTERS = (MCIERR_BASE + 49)
Public Const MCIERR_DEVICE_NOT_INSTALLED = (MCIERR_BASE + 50)
Public Const MCIERR_GET_CD = (MCIERR_BASE + 51)
Public Const MCIERR_SET_CD = (MCIERR_BASE + 52)
Public Const MCIERR_SET_DRIVE = (MCIERR_BASE + 53)
Public Const MCIERR_DEVICE_LENGTH = (MCIERR_BASE + 54)
Public Const MCIERR_DEVICE_ORD_LENGTH = (MCIERR_BASE + 55)
Public Const MCIERR_NO_INTEGER = (MCIERR_BASE + 56)
Public Const MCIERR_WAVE_OUTPUTSINUSE = (MCIERR_BASE + 64)
Public Const MCIERR_WAVE_SETOUTPUTINUSE = (MCIERR_BASE + 65)
Public Const MCIERR_WAVE_INPUTSINUSE = (MCIERR_BASE + 66)
Public Const MCIERR_WAVE_SETINPUTINUSE = (MCIERR_BASE + 67)
Public Const MCIERR_WAVE_OUTPUTUNSPECIFIED = (MCIERR_BASE + 68)
Public Const MCIERR_WAVE_INPUTUNSPECIFIED = (MCIERR_BASE + 69)
Public Const MCIERR_WAVE_OUTPUTSUNSUITABLE = (MCIERR_BASE + 70)
Public Const MCIERR_WAVE_SETOUTPUTUNSUITABLE = (MCIERR_BASE + 71)
Public Const MCIERR_WAVE_INPUTSUNSUITABLE = (MCIERR_BASE + 72)
Public Const MCIERR_WAVE_SETINPUTUNSUITABLE = (MCIERR_BASE + 73)
Public Const MCIERR_SEQ_DIV_INCOMPATIBLE = (MCIERR_BASE + 80)
Public Const MCIERR_SEQ_PORT_INUSE = (MCIERR_BASE + 81)
Public Const MCIERR_SEQ_PORT_NONEXISTENT = (MCIERR_BASE + 82)
Public Const MCIERR_SEQ_PORT_MAPNODEVICE = (MCIERR_BASE + 83)
Public Const MCIERR_SEQ_PORT_MISCERROR = (MCIERR_BASE + 84)
Public Const MCIERR_SEQ_TIMER = (MCIERR_BASE + 85)
Public Const MCIERR_SEQ_PORTUNSPECIFIED = (MCIERR_BASE + 86)
Public Const MCIERR_SEQ_NOMIDIPRESENT = (MCIERR_BASE + 87)
Public Const MCIERR_NO_WINDOW = (MCIERR_BASE + 90)
Public Const MCIERR_CREATEWINDOW = (MCIERR_BASE + 91)
Public Const MCIERR_FILE_READ = (MCIERR_BASE + 92)
Public Const MCIERR_FILE_WRITE = (MCIERR_BASE + 93)
'  All custom device driver errors must be >= this value

Public Const MCIERR_CUSTOM_DRIVER_BASE = (MCIERR_BASE + 256)
'  Message numbers must be in the range between MCI_FIRST and MCI_LAST

Public Const MCI_FIRST = &H800
'  Messages 0x801 and 0x802 are reserved

Public Const MCI_OPEN = &H803
Public Const MCI_CLOSE = &H804
Public Const MCI_ESCAPE = &H805
Public Const MCI_PLAY = &H806
Public Const MCI_SEEK = &H807
Public Const MCI_STOP = &H808
Public Const MCI_PAUSE = &H809
Public Const MCI_INFO = &H80A
Public Const MCI_GETDEVCAPS = &H80B
Public Const MCI_SPIN = &H80C
Public Const MCI_SET = &H80D
Public Const MCI_STEP = &H80E
Public Const MCI_RECORD = &H80F
Public Const MCI_SYSINFO = &H810
Public Const MCI_BREAK = &H811
Public Const MCI_SOUND = &H812
Public Const MCI_SAVE = &H813
Public Const MCI_STATUS = &H814
Public Const MCI_CUE = &H830
Public Const MCI_REALIZE = &H840
Public Const MCI_WINDOW = &H841
Public Const MCI_PUT = &H842
Public Const MCI_WHERE = &H843
Public Const MCI_FREEZE = &H844
Public Const MCI_UNFREEZE = &H845
Public Const MCI_LOAD = &H850
Public Const MCI_CUT = &H851
Public Const MCI_COPY = &H852
Public Const MCI_PASTE = &H853
Public Const MCI_UPDATE = &H854
Public Const MCI_RESUME = &H855
Public Const MCI_DELETE = &H856
Public Const MCI_LAST = &HFFF
'  the next 0x400 message ID's are reserved for custom drivers
'  all custom MCI command messages must be >= than this value

Public Const MCI_USER_MESSAGES = (&H400 + MCI_FIRST)
Public Const MCI_ALL_DEVICE_ID = -1   '  Matches all MCI devices
'  constants for predefined MCI device types

Public Const MCI_DEVTYPE_VCR = 513
Public Const MCI_DEVTYPE_VIDEODISC = 514
Public Const MCI_DEVTYPE_OVERLAY = 515
Public Const MCI_DEVTYPE_CD_AUDIO = 516
Public Const MCI_DEVTYPE_DAT = 517
Public Const MCI_DEVTYPE_SCANNER = 518
Public Const MCI_DEVTYPE_ANIMATION = 519
Public Const MCI_DEVTYPE_DIGITAL_VIDEO = 520
Public Const MCI_DEVTYPE_OTHER = 521
Public Const MCI_DEVTYPE_WAVEFORM_AUDIO = 522
Public Const MCI_DEVTYPE_SEQUENCER = 523
Public Const MCI_DEVTYPE_FIRST = MCI_DEVTYPE_VCR
Public Const MCI_DEVTYPE_LAST = MCI_DEVTYPE_SEQUENCER
Public Const MCI_DEVTYPE_FIRST_USER = &H1000
'  return values for 'status mode' command

Public Const MCI_MODE_NOT_READY = (MCI_STRING_OFFSET + 12)
Public Const MCI_MODE_STOP = (MCI_STRING_OFFSET + 13)
Public Const MCI_MODE_PLAY = (MCI_STRING_OFFSET + 14)
Public Const MCI_MODE_RECORD = (MCI_STRING_OFFSET + 15)
Public Const MCI_MODE_SEEK = (MCI_STRING_OFFSET + 16)
Public Const MCI_MODE_PAUSE = (MCI_STRING_OFFSET + 17)
Public Const MCI_MODE_OPEN = (MCI_STRING_OFFSET + 18)
'  constants used in 'set time format' and 'status time format' commands

Public Const MCI_FORMAT_MILLISECONDS = 0
Public Const MCI_FORMAT_HMS = 1
Public Const MCI_FORMAT_MSF = 2
Public Const MCI_FORMAT_FRAMES = 3
Public Const MCI_FORMAT_SMPTE_24 = 4
Public Const MCI_FORMAT_SMPTE_25 = 5
Public Const MCI_FORMAT_SMPTE_30 = 6
Public Const MCI_FORMAT_SMPTE_30DROP = 7
Public Const MCI_FORMAT_BYTES = 8
Public Const MCI_FORMAT_SAMPLES = 9
Public Const MCI_FORMAT_TMSF = 10
'  Flags for wParam of the MM_MCINOTIFY message

Public Const MCI_NOTIFY_SUCCESSFUL = &H1
Public Const MCI_NOTIFY_SUPERSEDED = &H2
Public Const MCI_NOTIFY_ABORTED = &H4
Public Const MCI_NOTIFY_FAILURE = &H8
'  common flags for dwFlags parameter of MCI command messages

Public Const MCI_NOTIFY = &H1&
Public Const MCI_WAIT = &H2&
Public Const MCI_FROM = &H4&
Public Const MCI_TO = &H8&
Public Const MCI_TRACK = &H10&
'  flags for dwFlags parameter of MCI_OPEN command message

Public Const MCI_OPEN_SHAREABLE = &H100&
Public Const MCI_OPEN_ELEMENT = &H200&
Public Const MCI_OPEN_ALIAS = &H400&
Public Const MCI_OPEN_ELEMENT_ID = &H800&
Public Const MCI_OPEN_TYPE_ID = &H1000&
Public Const MCI_OPEN_TYPE = &H2000&
'  flags for dwFlags parameter of MCI_SEEK command message

Public Const MCI_SEEK_TO_START = &H100&
Public Const MCI_SEEK_TO_END = &H200&
'  flags for dwFlags parameter of MCI_STATUS command message

Public Const MCI_STATUS_ITEM = &H100&
Public Const MCI_STATUS_START = &H200&
'  flags for dwItem field of the MCI_STATUS_PARMS parameter block

Public Const MCI_STATUS_LENGTH = &H1&
Public Const MCI_STATUS_POSITION = &H2&
Public Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Public Const MCI_STATUS_MODE = &H4&
Public Const MCI_STATUS_MEDIA_PRESENT = &H5&
Public Const MCI_STATUS_TIME_FORMAT = &H6&
Public Const MCI_STATUS_READY = &H7&
Public Const MCI_STATUS_CURRENT_TRACK = &H8&
'  flags for dwFlags parameter of MCI_INFO command message

Public Const MCI_INFO_PRODUCT = &H100&
Public Const MCI_INFO_FILE = &H200&
'  flags for dwFlags parameter of MCI_GETDEVCAPS command message

Public Const MCI_GETDEVCAPS_ITEM = &H100&
'  flags for dwItem field of the MCI_GETDEVCAPS_PARMS parameter block

Public Const MCI_GETDEVCAPS_CAN_RECORD = &H1&
Public Const MCI_GETDEVCAPS_HAS_AUDIO = &H2&
Public Const MCI_GETDEVCAPS_HAS_VIDEO = &H3&
Public Const MCI_GETDEVCAPS_DEVICE_TYPE = &H4&
Public Const MCI_GETDEVCAPS_USES_FILES = &H5&
Public Const MCI_GETDEVCAPS_COMPOUND_DEVICE = &H6&
Public Const MCI_GETDEVCAPS_CAN_EJECT = &H7&
Public Const MCI_GETDEVCAPS_CAN_PLAY = &H8&
Public Const MCI_GETDEVCAPS_CAN_SAVE = &H9&
'  flags for dwFlags parameter of MCI_SYSINFO command message

Public Const MCI_SYSINFO_QUANTITY = &H100&
Public Const MCI_SYSINFO_OPEN = &H200&
Public Const MCI_SYSINFO_NAME = &H400&
Public Const MCI_SYSINFO_INSTALLNAME = &H800&
'  flags for dwFlags parameter of MCI_SET command message

Public Const MCI_SET_DOOR_OPEN = &H100&
Public Const MCI_SET_DOOR_CLOSED = &H200&
Public Const MCI_SET_TIME_FORMAT = &H400&
Public Const MCI_SET_AUDIO = &H800&
Public Const MCI_SET_VIDEO = &H1000&
Public Const MCI_SET_ON = &H2000&
Public Const MCI_SET_OFF = &H4000&
'  flags for dwAudio field of MCI_SET_PARMS or MCI_SEQ_SET_PARMS

Public Const MCI_SET_AUDIO_ALL = &H4001&
Public Const MCI_SET_AUDIO_LEFT = &H4002&
Public Const MCI_SET_AUDIO_RIGHT = &H4003&
'  flags for dwFlags parameter of MCI_BREAK command message

Public Const MCI_BREAK_KEY = &H100&
Public Const MCI_BREAK_HWND = &H200&
Public Const MCI_BREAK_OFF = &H400&
'  flags for dwFlags parameter of MCI_RECORD command message

Public Const MCI_RECORD_INSERT = &H100&
Public Const MCI_RECORD_OVERWRITE = &H200&
'  flags for dwFlags parameter of MCI_SOUND command message

Public Const MCI_SOUND_NAME = &H100&
'  flags for dwFlags parameter of MCI_SAVE command message

Public Const MCI_SAVE_FILE = &H100&
'  flags for dwFlags parameter of MCI_LOAD command message

Public Const MCI_LOAD_FILE = &H100&
Public Const MCI_VD_MODE_PARK = (MCI_VD_OFFSET + 1)
'  return ID's for videodisc MCI_GETDEVCAPS command
'  flag for dwReturn field of MCI_STATUS_PARMS
'  MCI_STATUS command, (dwItem == MCI_VD_STATUS_MEDIA_TYPE)

Public Const MCI_VD_MEDIA_CLV = (MCI_VD_OFFSET + 2)
Public Const MCI_VD_MEDIA_CAV = (MCI_VD_OFFSET + 3)
Public Const MCI_VD_MEDIA_OTHER = (MCI_VD_OFFSET + 4)
Public Const MCI_VD_FORMAT_TRACK = &H4001
'  flags for dwFlags parameter of MCI_PLAY command message

Public Const MCI_VD_PLAY_REVERSE = &H10000
Public Const MCI_VD_PLAY_FAST = &H20000
Public Const MCI_VD_PLAY_SPEED = &H40000
Public Const MCI_VD_PLAY_SCAN = &H80000
Public Const MCI_VD_PLAY_SLOW = &H100000
'  flag for dwFlags parameter of MCI_SEEK command message

Public Const MCI_VD_SEEK_REVERSE = &H10000
'  flags for dwItem field of MCI_STATUS_PARMS parameter block

Public Const MCI_VD_STATUS_SPEED = &H4002&
Public Const MCI_VD_STATUS_FORWARD = &H4003&
Public Const MCI_VD_STATUS_MEDIA_TYPE = &H4004&
Public Const MCI_VD_STATUS_SIDE = &H4005&
Public Const MCI_VD_STATUS_DISC_SIZE = &H4006&
'  flags for dwFlags parameter of MCI_GETDEVCAPS command message

Public Const MCI_VD_GETDEVCAPS_CLV = &H10000
Public Const MCI_VD_GETDEVCAPS_CAV = &H20000
Public Const MCI_VD_SPIN_UP = &H10000
Public Const MCI_VD_SPIN_DOWN = &H20000
'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block

Public Const MCI_VD_GETDEVCAPS_CAN_REVERSE = &H4002&
Public Const MCI_VD_GETDEVCAPS_FAST_RATE = &H4003&
Public Const MCI_VD_GETDEVCAPS_SLOW_RATE = &H4004&
Public Const MCI_VD_GETDEVCAPS_NORMAL_RATE = &H4005&
'  flags for the dwFlags parameter of MCI_STEP command message

Public Const MCI_VD_STEP_FRAMES = &H10000
Public Const MCI_VD_STEP_REVERSE = &H20000
'  flag for the MCI_ESCAPE command message

Public Const MCI_VD_ESCAPE_STRING = &H100&
Public Const MCI_WAVE_PCM = (MCI_WAVE_OFFSET + 0)
Public Const MCI_WAVE_MAPPER = (MCI_WAVE_OFFSET + 1)
'  flags for the dwFlags parameter of MCI_OPEN command message

Public Const MCI_WAVE_OPEN_BUFFER = &H10000
'  flags for the dwFlags parameter of MCI_SET command message

Public Const MCI_WAVE_SET_FORMATTAG = &H10000
Public Const MCI_WAVE_SET_CHANNELS = &H20000
Public Const MCI_WAVE_SET_SAMPLESPERSEC = &H40000
Public Const MCI_WAVE_SET_AVGBYTESPERSEC = &H80000
Public Const MCI_WAVE_SET_BLOCKALIGN = &H100000
Public Const MCI_WAVE_SET_BITSPERSAMPLE = &H200000
'  flags for the dwFlags parameter of MCI_STATUS, MCI_SET command messages

Public Const MCI_WAVE_INPUT = &H400000
Public Const MCI_WAVE_OUTPUT = &H800000
'  flags for the dwItem field of MCI_STATUS_PARMS parameter block

Public Const MCI_WAVE_STATUS_FORMATTAG = &H4001&
Public Const MCI_WAVE_STATUS_CHANNELS = &H4002&
Public Const MCI_WAVE_STATUS_SAMPLESPERSEC = &H4003&
Public Const MCI_WAVE_STATUS_AVGBYTESPERSEC = &H4004&
Public Const MCI_WAVE_STATUS_BLOCKALIGN = &H4005&
Public Const MCI_WAVE_STATUS_BITSPERSAMPLE = &H4006&
Public Const MCI_WAVE_STATUS_LEVEL = &H4007&
'  flags for the dwFlags parameter of MCI_SET command message

Public Const MCI_WAVE_SET_ANYINPUT = &H4000000
Public Const MCI_WAVE_SET_ANYOUTPUT = &H8000000
'  flags for the dwFlags parameter of MCI_GETDEVCAPS command message

Public Const MCI_WAVE_GETDEVCAPS_INPUTS = &H4001&
Public Const MCI_WAVE_GETDEVCAPS_OUTPUTS = &H4002&
'  flags for the dwReturn field of MCI_STATUS_PARMS parameter block
'  MCI_STATUS command, (dwItem == MCI_SEQ_STATUS_DIVTYPE)

Public Const MCI_SEQ_DIV_PPQN = (0 + MCI_SEQ_OFFSET)
Public Const MCI_SEQ_DIV_SMPTE_24 = (1 + MCI_SEQ_OFFSET)
Public Const MCI_SEQ_DIV_SMPTE_25 = (2 + MCI_SEQ_OFFSET)
Public Const MCI_SEQ_DIV_SMPTE_30DROP = (3 + MCI_SEQ_OFFSET)
Public Const MCI_SEQ_DIV_SMPTE_30 = (4 + MCI_SEQ_OFFSET)
'  flags for the dwMaster field of MCI_SEQ_SET_PARMS parameter block
'  MCI_SET command, (dwFlags == MCI_SEQ_SET_MASTER)

Public Const MCI_SEQ_FORMAT_SONGPTR = &H4001
Public Const MCI_SEQ_FILE = &H4002
Public Const MCI_SEQ_MIDI = &H4003
Public Const MCI_SEQ_SMPTE = &H4004
Public Const MCI_SEQ_NONE = 65533
Public Const MCI_SEQ_MAPPER = 65535
'  flags for the dwItem field of MCI_STATUS_PARMS parameter block

Public Const MCI_SEQ_STATUS_TEMPO = &H4002&
Public Const MCI_SEQ_STATUS_PORT = &H4003&
Public Const MCI_SEQ_STATUS_SLAVE = &H4007&
Public Const MCI_SEQ_STATUS_MASTER = &H4008&
Public Const MCI_SEQ_STATUS_OFFSET = &H4009&
Public Const MCI_SEQ_STATUS_DIVTYPE = &H400A&
'  flags for the dwFlags parameter of MCI_SET command message

Public Const MCI_SEQ_SET_TEMPO = &H10000
Public Const MCI_SEQ_SET_PORT = &H20000
Public Const MCI_SEQ_SET_SLAVE = &H40000
Public Const MCI_SEQ_SET_MASTER = &H80000
Public Const MCI_SEQ_SET_OFFSET = &H1000000
'  flags for dwFlags parameter of MCI_OPEN command message

Public Const MCI_ANIM_OPEN_WS = &H10000
Public Const MCI_ANIM_OPEN_PARENT = &H20000
Public Const MCI_ANIM_OPEN_NOSTATIC = &H40000
'  flags for dwFlags parameter of MCI_PLAY command message

Public Const MCI_ANIM_PLAY_SPEED = &H10000
Public Const MCI_ANIM_PLAY_REVERSE = &H20000
Public Const MCI_ANIM_PLAY_FAST = &H40000
Public Const MCI_ANIM_PLAY_SLOW = &H80000
Public Const MCI_ANIM_PLAY_SCAN = &H100000
'  flags for dwFlags parameter of MCI_STEP command message

Public Const MCI_ANIM_STEP_REVERSE = &H10000
Public Const MCI_ANIM_STEP_FRAMES = &H20000
'  flags for dwItem field of MCI_STATUS_PARMS parameter block

Public Const MCI_ANIM_STATUS_SPEED = &H4001&
Public Const MCI_ANIM_STATUS_FORWARD = &H4002&
Public Const MCI_ANIM_STATUS_HWND = &H4003&
Public Const MCI_ANIM_STATUS_HPAL = &H4004&
Public Const MCI_ANIM_STATUS_STRETCH = &H4005&
'  flags for the dwFlags parameter of MCI_INFO command message

Public Const MCI_ANIM_INFO_TEXT = &H10000
'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block

Public Const MCI_ANIM_GETDEVCAPS_CAN_REVERSE = &H4001&
Public Const MCI_ANIM_GETDEVCAPS_FAST_RATE = &H4002&
Public Const MCI_ANIM_GETDEVCAPS_SLOW_RATE = &H4003&
Public Const MCI_ANIM_GETDEVCAPS_NORMAL_RATE = &H4004&
Public Const MCI_ANIM_GETDEVCAPS_PALETTES = &H4006&
Public Const MCI_ANIM_GETDEVCAPS_CAN_STRETCH = &H4007&
Public Const MCI_ANIM_GETDEVCAPS_MAX_WINDOWS = &H4008&
'  flags for the MCI_REALIZE command message

Public Const MCI_ANIM_REALIZE_NORM = &H10000
Public Const MCI_ANIM_REALIZE_BKGD = &H20000
'  flags for dwFlags parameter of MCI_WINDOW command message

Public Const MCI_ANIM_WINDOW_HWND = &H10000
Public Const MCI_ANIM_WINDOW_STATE = &H40000
Public Const MCI_ANIM_WINDOW_TEXT = &H80000
Public Const MCI_ANIM_WINDOW_ENABLE_STRETCH = &H100000
Public Const MCI_ANIM_WINDOW_DISABLE_STRETCH = &H200000
'  flags for hWnd field of MCI_ANIM_WINDOW_PARMS parameter block
'  MCI_WINDOW command message, (dwFlags == MCI_ANIM_WINDOW_HWND)

Public Const MCI_ANIM_WINDOW_DEFAULT = &H0&
'  flags for dwFlags parameter of MCI_PUT command message

Public Const MCI_ANIM_RECT = &H10000
Public Const MCI_ANIM_PUT_SOURCE = &H20000      '  also  MCI_WHERE
Public Const MCI_ANIM_PUT_DESTINATION = &H40000      '  also  MCI_WHERE
'  flags for dwFlags parameter of MCI_WHERE command message

Public Const MCI_ANIM_WHERE_SOURCE = &H20000
Public Const MCI_ANIM_WHERE_DESTINATION = &H40000
'  flags for dwFlags parameter of MCI_UPDATE command message

Public Const MCI_ANIM_UPDATE_HDC = &H20000
'  flags for dwFlags parameter of MCI_OPEN command message

Public Const MCI_OVLY_OPEN_WS = &H10000
Public Const MCI_OVLY_OPEN_PARENT = &H20000
'  flags for dwFlags parameter of MCI_STATUS command message

Public Const MCI_OVLY_STATUS_HWND = &H4001&
Public Const MCI_OVLY_STATUS_STRETCH = &H4002&
'  flags for dwFlags parameter of MCI_INFO command message

Public Const MCI_OVLY_INFO_TEXT = &H10000
'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block

Public Const MCI_OVLY_GETDEVCAPS_CAN_STRETCH = &H4001&
Public Const MCI_OVLY_GETDEVCAPS_CAN_FREEZE = &H4002&
Public Const MCI_OVLY_GETDEVCAPS_MAX_WINDOWS = &H4003&
'  flags for dwFlags parameter of MCI_WINDOW command message

Public Const MCI_OVLY_WINDOW_HWND = &H10000
Public Const MCI_OVLY_WINDOW_STATE = &H40000
Public Const MCI_OVLY_WINDOW_TEXT = &H80000
Public Const MCI_OVLY_WINDOW_ENABLE_STRETCH = &H100000
Public Const MCI_OVLY_WINDOW_DISABLE_STRETCH = &H200000
'  flags for hWnd parameter of MCI_OVLY_WINDOW_PARMS parameter block

Public Const MCI_OVLY_WINDOW_DEFAULT = &H0&
'  flags for dwFlags parameter of MCI_PUT command message

Public Const MCI_OVLY_RECT = &H10000
Public Const MCI_OVLY_PUT_SOURCE = &H20000
Public Const MCI_OVLY_PUT_DESTINATION = &H40000
Public Const MCI_OVLY_PUT_FRAME = &H80000
Public Const MCI_OVLY_PUT_VIDEO = &H100000
'  flags for dwFlags parameter of MCI_WHERE command message

Public Const MCI_OVLY_WHERE_SOURCE = &H20000
Public Const MCI_OVLY_WHERE_DESTINATION = &H40000
Public Const MCI_OVLY_WHERE_FRAME = &H80000
Public Const MCI_OVLY_WHERE_VIDEO = &H100000
Public Const CAPS1 = 94              '  other caps
Public Const C1_TRANSPARENT = &H1     '  new raster cap
Public Const NEWTRANSPARENT = 3  '  use with SetBkMode()
Public Const QUERYROPSUPPORT = 40  '  use to determine ROP support
Public Const SELECTDIB = 41  '  DIB.DRV select dib escape
' ----------------
' shell association database management functions
' -----------------
'  error values for ShellExecute() beyond the regular WinExec() codes

Public Const SE_ERR_SHARE = 26
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_NOASSOC = 31
Public Const PRINTER_CONTROL_PAUSE = 1
Public Const PRINTER_CONTROL_RESUME = 2
Public Const PRINTER_CONTROL_PURGE = 3
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_WARMING_UP = &H10000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_ATTRIBUTE_QUEUED = &H1
Public Const PRINTER_ATTRIBUTE_DIRECT = &H2
Public Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Public Const PRINTER_ATTRIBUTE_SHARED = &H8
Public Const PRINTER_ATTRIBUTE_NETWORK = &H10
Public Const PRINTER_ATTRIBUTE_HIDDEN = &H20
Public Const PRINTER_ATTRIBUTE_LOCAL = &H40
Public Const NO_PRIORITY = 0
Public Const MAX_PRIORITY = 99
Public Const MIN_PRIORITY = 1
Public Const DEF_PRIORITY = 1
Public Const JOB_CONTROL_PAUSE = 1
Public Const JOB_CONTROL_RESUME = 2
Public Const JOB_CONTROL_CANCEL = 3
Public Const JOB_CONTROL_RESTART = 4
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_POSITION_UNSPECIFIED = 0
Public Const FORM_BUILTIN = &H1
Public Const PRINTER_CONTROL_SET_STATUS = 4
Public Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
Public Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800
Public Const JOB_CONTROL_DELETE = 5
Public Const JOB_STATUS_USER_INTERVENTION = &H10000
Public Const DI_CHANNEL = 1                  '  start direct read/write channel,
Public Const DI_READ_SPOOL_JOB = 3
Public Const PORT_TYPE_WRITE = &H1
Public Const PORT_TYPE_READ = &H2
Public Const PORT_TYPE_REDIRECTED = &H4
Public Const PORT_TYPE_NET_ATTACHED = &H8
Public Const PRINTER_ENUM_DEFAULT = &H1
Public Const PRINTER_ENUM_LOCAL = &H2
Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_FAVORITE = &H4
Public Const PRINTER_ENUM_NAME = &H8
Public Const PRINTER_ENUM_REMOTE = &H10
Public Const PRINTER_ENUM_SHARED = &H20
Public Const PRINTER_ENUM_NETWORK = &H40
Public Const PRINTER_ENUM_EXPAND = &H4000
Public Const PRINTER_ENUM_CONTAINER = &H8000
Public Const PRINTER_ENUM_ICONMASK = &HFF0000
Public Const PRINTER_ENUM_ICON1 = &H10000
Public Const PRINTER_ENUM_ICON2 = &H20000
Public Const PRINTER_ENUM_ICON3 = &H40000
Public Const PRINTER_ENUM_ICON4 = &H80000
Public Const PRINTER_ENUM_ICON5 = &H100000
Public Const PRINTER_ENUM_ICON6 = &H200000
Public Const PRINTER_ENUM_ICON7 = &H400000
Public Const PRINTER_ENUM_ICON8 = &H800000
Public Const PRINTER_CHANGE_ADD_PRINTER = &H1
Public Const PRINTER_CHANGE_SET_PRINTER = &H2
Public Const PRINTER_CHANGE_DELETE_PRINTER = &H4
Public Const PRINTER_CHANGE_PRINTER = &HFF
Public Const PRINTER_CHANGE_ADD_JOB = &H100
Public Const PRINTER_CHANGE_SET_JOB = &H200
Public Const PRINTER_CHANGE_DELETE_JOB = &H400
Public Const PRINTER_CHANGE_WRITE_JOB = &H800
Public Const PRINTER_CHANGE_JOB = &HFF00
Public Const PRINTER_CHANGE_ADD_FORM = &H10000
Public Const PRINTER_CHANGE_SET_FORM = &H20000
Public Const PRINTER_CHANGE_DELETE_FORM = &H40000
Public Const PRINTER_CHANGE_FORM = &H70000
Public Const PRINTER_CHANGE_ADD_PORT = &H100000
Public Const PRINTER_CHANGE_CONFIGURE_PORT = &H200000
Public Const PRINTER_CHANGE_DELETE_PORT = &H400000
Public Const PRINTER_CHANGE_PORT = &H700000
Public Const PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
Public Const PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
Public Const PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
Public Const PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
Public Const PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
Public Const PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
Public Const PRINTER_CHANGE_TIMEOUT = &H80000000
Public Const PRINTER_CHANGE_ALL = &H7777FFFF
Public Const PRINTER_ERROR_INFORMATION = &H80000000
Public Const PRINTER_ERROR_WARNING = &H40000000
Public Const PRINTER_ERROR_SEVERE = &H20000000
Public Const PRINTER_ERROR_OUTOFPAPER = &H1
Public Const PRINTER_ERROR_JAM = &H2
Public Const PRINTER_ERROR_OUTOFTONER = &H4
Public Const SERVER_ACCESS_ADMINISTER = &H1
Public Const SERVER_ACCESS_ENUMERATE = &H2
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const JOB_ACCESS_ADMINISTER = &H10
' Access rights for print servers

Public Const SERVER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Public Const SERVER_READ = (STANDARD_RIGHTS_READ Or SERVER_ACCESS_ENUMERATE)
Public Const SERVER_WRITE = (STANDARD_RIGHTS_WRITE Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Public Const SERVER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or SERVER_ACCESS_ENUMERATE)
' Access rights for printers

Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Public Const PRINTER_READ = (STANDARD_RIGHTS_READ Or PRINTER_ACCESS_USE)
Public Const PRINTER_WRITE = (STANDARD_RIGHTS_WRITE Or PRINTER_ACCESS_USE)
Public Const PRINTER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or PRINTER_ACCESS_USE)
' Access rights for jobs

Public Const JOB_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or JOB_ACCESS_ADMINISTER)
Public Const JOB_READ = (STANDARD_RIGHTS_READ Or JOB_ACCESS_ADMINISTER)
Public Const JOB_WRITE = (STANDARD_RIGHTS_WRITE Or JOB_ACCESS_ADMINISTER)
Public Const JOB_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or JOB_ACCESS_ADMINISTER)
'  Windows Network support
'  RESOURCE ENUMERATION

Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_UNKNOWN = &HFFFF
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
Public Const RESOURCEUSAGE_RESERVED = &H80000000
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEDISPLAYTYPE_FILE = &H4
Public Const RESOURCEDISPLAYTYPE_GROUP = &H5
Public Const CONNECT_UPDATE_PROFILE = &H1
' Status Codes
' This section is provided for backward compatibility.  Use of the ERROR_
' codes is preferred.  The WN_ error codes may not be available in future
' releases.
' General

Public Const WN_SUCCESS = NO_ERROR
Public Const WN_NOT_SUPPORTED = ERROR_NOT_SUPPORTED
Public Const WN_NET_ERROR = ERROR_UNEXP_NET_ERR
Public Const WN_MORE_DATA = ERROR_MORE_DATA
Public Const WN_BAD_POINTER = ERROR_INVALID_ADDRESS
Public Const WN_BAD_VALUE = ERROR_INVALID_PARAMETER
Public Const WN_BAD_PASSWORD = ERROR_INVALID_PASSWORD
Public Const WN_ACCESS_DENIED = ERROR_ACCESS_DENIED
Public Const WN_FUNCTION_BUSY = ERROR_BUSY
Public Const WN_WINDOWS_ERROR = ERROR_UNEXP_NET_ERR
Public Const WN_BAD_USER = ERROR_BAD_USERNAME
Public Const WN_OUT_OF_MEMORY = ERROR_NOT_ENOUGH_MEMORY
Public Const WN_NO_NETWORK = ERROR_NO_NETWORK
Public Const WN_EXTENDED_ERROR = ERROR_EXTENDED_ERROR
' Connection

Public Const WN_NOT_CONNECTED = ERROR_NOT_CONNECTED
Public Const WN_OPEN_FILES = ERROR_OPEN_FILES
Public Const WN_DEVICE_IN_USE = ERROR_DEVICE_IN_USE
Public Const WN_BAD_NETNAME = ERROR_BAD_NET_NAME
Public Const WN_BAD_LOCALNAME = ERROR_BAD_DEVICE
Public Const WN_ALREADY_CONNECTED = ERROR_ALREADY_ASSIGNED
Public Const WN_DEVICE_ERROR = ERROR_GEN_FAILURE
Public Const WN_CONNECTION_CLOSED = ERROR_CONNECTION_UNAVAIL
Public Const WN_NO_NET_OR_BAD_PATH = ERROR_NO_NET_OR_BAD_PATH
Public Const WN_BAD_PROVIDER = ERROR_BAD_PROVIDER
Public Const WN_CANNOT_OPEN_PROFILE = ERROR_CANNOT_OPEN_PROFILE
Public Const WN_BAD_PROFILE = ERROR_BAD_PROFILE
' Enumeration

Public Const WN_BAD_HANDLE = ERROR_INVALID_HANDLE
Public Const WN_NO_MORE_ENTRIES = ERROR_NO_MORE_ITEMS
Public Const WN_NOT_CONTAINER = ERROR_NOT_CONTAINER
Public Const WN_NO_ERROR = NO_ERROR
' This section contains the definitions
' for portable NetBIOS 3.0 support.

Public Const NCBNAMSZ = 16  '  absolute length of a net name
Public Const MAX_LANA = 254  '  lana's in range 0 to MAX_LANA
' values for name_flags bits.

Public Const NAME_FLAGS_MASK = &H87
Public Const GROUP_NAME = &H80
Public Const UNIQUE_NAME = &H0
Public Const REGISTERING = &H0
Public Const REGISTERED = &H4
Public Const DEREGISTERED = &H5
Public Const DUPLICATE = &H6
Public Const DUPLICATE_DEREG = &H7
' Values for state

Public Const LISTEN_OUTSTANDING = &H1
Public Const CALL_PENDING = &H2
Public Const SESSION_ESTABLISHED = &H3
Public Const HANGUP_PENDING = &H4
Public Const HANGUP_COMPLETE = &H5
Public Const SESSION_ABORTED = &H6
' Values for transport_id

Public Const ALL_TRANSPORTS = "M\0\0\0"
Public Const MS_NBF = "MNBF"
' NCB Command codes

Public Const NCBCALL = &H10  '  NCB CALL
Public Const NCBLISTEN = &H11  '  NCB LISTEN
Public Const NCBHANGUP = &H12  '  NCB HANG UP
Public Const NCBSEND = &H14  '  NCB SEND
Public Const NCBRECV = &H15  '  NCB RECEIVE
Public Const NCBRECVANY = &H16  '  NCB RECEIVE ANY
Public Const NCBCHAINSEND = &H17  '  NCB CHAIN SEND
Public Const NCBDGSEND = &H20  '  NCB SEND DATAGRAM
Public Const NCBDGRECV = &H21  '  NCB RECEIVE DATAGRAM
Public Const NCBDGSENDBC = &H22  '  NCB SEND BROADCAST DATAGRAM
Public Const NCBDGRECVBC = &H23  '  NCB RECEIVE BROADCAST DATAGRAM
Public Const NCBADDNAME = &H30  '  NCB ADD NAME
Public Const NCBDELNAME = &H31  '  NCB DELETE NAME
Public Const NCBRESET = &H32  '  NCB RESET
Public Const NCBASTAT = &H33  '  NCB ADAPTER STATUS
Public Const NCBSSTAT = &H34  '  NCB SESSION STATUS
Public Const NCBCANCEL = &H35  '  NCB CANCEL
Public Const NCBADDGRNAME = &H36  '  NCB ADD GROUP NAME
Public Const NCBENUM = &H37  '  NCB ENUMERATE LANA NUMBERS
Public Const NCBUNLINK = &H70  '  NCB UNLINK
Public Const NCBSENDNA = &H71  '  NCB SEND NO ACK
Public Const NCBCHAINSENDNA = &H72  '  NCB CHAIN SEND NO ACK
Public Const NCBLANSTALERT = &H73  '  NCB LAN STATUS ALERT
Public Const NCBACTION = &H77  '  NCB ACTION
Public Const NCBFINDNAME = &H78  '  NCB FIND NAME
Public Const NCBTRACE = &H79  '  NCB TRACE
Public Const ASYNCH = &H80  '  high bit set == asynchronous
' NCB Return codes

Public Const NRC_GOODRET = &H0   '  good return
                                '  also returned when ASYNCH request accepted

Public Const NRC_BUFLEN = &H1   '  illegal buffer length
Public Const NRC_ILLCMD = &H3   '  illegal command
Public Const NRC_CMDTMO = &H5   '  command timed out
Public Const NRC_INCOMP = &H6   '  message incomplete, issue another command
Public Const NRC_BADDR = &H7   '  illegal buffer address
Public Const NRC_SNUMOUT = &H8   '  session number out of range
Public Const NRC_NORES = &H9   '  no resource available
Public Const NRC_SCLOSED = &HA   '  session closed
Public Const NRC_CMDCAN = &HB   '  command cancelled
Public Const NRC_DUPNAME = &HD   '  duplicate name
Public Const NRC_NAMTFUL = &HE   '  name table full
Public Const NRC_ACTSES = &HF   '  no deletions, name has active sessions
Public Const NRC_LOCTFUL = &H11  '  local session table full
Public Const NRC_REMTFUL = &H12  '  remote session table full
Public Const NRC_ILLNN = &H13  '  illegal name number
Public Const NRC_NOCALL = &H14  '  no callname
Public Const NRC_NOWILD = &H15  '  cannot put  in NCB_NAME
Public Const NRC_INUSE = &H16  '  name in use on remote adapter
Public Const NRC_NAMERR = &H17  '  name deleted
Public Const NRC_SABORT = &H18  '  session ended abnormally
Public Const NRC_NAMCONF = &H19  '  name conflict detected
Public Const NRC_IFBUSY = &H21  '  interface busy, IRET before retrying
Public Const NRC_TOOMANY = &H22  '  too many commands outstanding, retry later
Public Const NRC_BRIDGE = &H23  '  ncb_lana_num field invalid
Public Const NRC_CANOCCR = &H24  '  command completed while cancel occurring
Public Const NRC_CANCEL = &H26  '  command not valid to cancel
Public Const NRC_DUPENV = &H30  '  name defined by anther local process
Public Const NRC_ENVNOTDEF = &H34  '  environment undefined. RESET required
Public Const NRC_OSRESNOTAV = &H35  '  required OS resources exhausted
Public Const NRC_MAXAPPS = &H36  '  max number of applications exceeded
Public Const NRC_NOSAPS = &H37  '  no saps available for netbios
Public Const NRC_NORESOURCES = &H38  '  requested resources are not available
Public Const NRC_INVADDRESS = &H39  '  invalid ncb address or length > segment
Public Const NRC_INVDDID = &H3B  '  invalid NCB DDID
Public Const NRC_LOCKFAIL = &H3C  '  lock of user area failed
Public Const NRC_OPENERR = &H3F  '  NETBIOS not loaded
Public Const NRC_SYSTEM = &H40  '  system error
Public Const NRC_PENDING = &HFF  '  asynchronous command is not yet finished
Public Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Public Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Public Const FILTER_PROXY_ACCOUNT As Long = &H4&
Public Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Public Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Public Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&
Public Const TIMEQ_FOREVER = -1&             '((unsigned long) -1L)
Public Const USER_MAXSTORAGE_UNLIMITED = -1& '((unsigned long) -1L)
Public Const USER_NO_LOGOFF = -1&            '((unsigned long) -1L)
Public Const UNITS_PER_DAY = 24
Public Const UNITS_PER_WEEK = UNITS_PER_DAY * 7
Public Const USER_PRIV_MASK = 3
Public Const USER_PRIV_GUEST = 0
Public Const USER_PRIV_USER = 1
Public Const USER_PRIV_ADMIN = 2
Public Const UNLEN = 256         ' Maximum username length
Public Const GNLEN = UNLEN       ' Maximum groupname length
Public Const CNLEN = 15          ' Maximum computer name length
Public Const PWLEN = 256         ' Maximum password length
Public Const LM20_PWLEN = 14     ' LM 2.0 Maximum password length
Public Const MAXCOMMENTSZ = 256  ' Multipurpose comment length
'Public Const LG_INCLUDE_INDIRECT As Long = &H1&
Public Const UF_SCRIPT = &H1
Public Const UF_ACCOUNTDISABLE = &H2
Public Const UF_HOMEDIR_REQUIRED = &H8
Public Const UF_LOCKOUT = &H10
Public Const UF_PASSWD_NOTREQD = &H20
Public Const UF_PASSWD_CANT_CHANGE = &H40
Public Const LG_INCLUDE_INDIRECT As Long = &H1&
Public Const NERR_Success As Long = 0&
Public Const NERR_BASE = 2100
Public Const NERR_InvalidComputer = (NERR_BASE + 251)
Public Const NERR_NotPrimary = (NERR_BASE + 126)
Public Const NERR_GroupExists = (NERR_BASE + 123)
Public Const NERR_UserExists = (NERR_BASE + 124)
Public Const NERR_PasswordTooShort = (NERR_BASE + 145)
'Public Const RESOURCE_CONNECTED As Long = &H1&
Public Const RESOURCE_GLOBALNET As Long = &H2&
'Public Const RESOURCE_REMEMBERED As Long = &H3&
Public Const RESOURCE_ENUM_ALL As Long = &HFFFF
'Public Const RESOURCEDISPLAYTYPE_DOMAIN As Long = &H1&
'Public Const RESOURCEDISPLAYTYPE_FILE As Long = &H4&
'Public Const RESOURCEDISPLAYTYPE_GENERIC As Long = &H0&
'Public Const RESOURCEDISPLAYTYPE_GROUP As Long = &H5&
'Public Const RESOURCEDISPLAYTYPE_SERVER As Long = &H2&
'Public Const RESOURCEDISPLAYTYPE_SHARE As Long = &H3&
'Public Const RESOURCETYPE_ANY As Long = &H0&
'Public Const RESOURCETYPE_DISK As Long = &H1&
'Public Const RESOURCETYPE_PRINT As Long = &H2&
'Public Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Public Const RESOURCEUSAGE_ALL As Long = &H0&
'Public Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
'Public Const RESOURCEUSAGE_CONTAINER As Long = &H2&
'Public Const RESOURCEUSAGE_RESERVED As Long = &H80000000
' Legal values for expression in except().

Public Const EXCEPTION_EXECUTE_HANDLER = 1
Public Const EXCEPTION_CONTINUE_SEARCH = 0
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
' UI dialog constants and types
' ----Constants--------------------------------------------------------------

Public Const ctlFirst = &H400
Public Const ctlLast = &H4FF
    '  Push buttons

Public Const psh1 = &H400
Public Const psh2 = &H401
Public Const psh3 = &H402
Public Const psh4 = &H403
Public Const psh5 = &H404
Public Const psh6 = &H405
Public Const psh7 = &H406
Public Const psh8 = &H407
Public Const psh9 = &H408
Public Const psh10 = &H409
Public Const psh11 = &H40A
Public Const psh12 = &H40B
Public Const psh13 = &H40C
Public Const psh14 = &H40D
Public Const psh15 = &H40E
Public Const pshHelp = psh15
Public Const psh16 = &H40F
    '  Checkboxes

Public Const chx1 = &H410
Public Const chx2 = &H411
Public Const chx3 = &H412
Public Const chx4 = &H413
Public Const chx5 = &H414
Public Const chx6 = &H415
Public Const chx7 = &H416
Public Const chx8 = &H417
Public Const chx9 = &H418
Public Const chx10 = &H419
Public Const chx11 = &H41A
Public Const chx12 = &H41B
Public Const chx13 = &H41C
Public Const chx14 = &H41D
Public Const chx15 = &H41E
Public Const chx16 = &H41D
    '  Radio buttons

Public Const rad1 = &H420
Public Const rad2 = &H421
Public Const rad3 = &H422
Public Const rad4 = &H423
Public Const rad5 = &H424
Public Const rad6 = &H425
Public Const rad7 = &H426
Public Const rad8 = &H427
Public Const rad9 = &H428
Public Const rad10 = &H429
Public Const rad11 = &H42A
Public Const rad12 = &H42B
Public Const rad13 = &H42C
Public Const rad14 = &H42D
Public Const rad15 = &H42E
Public Const rad16 = &H42F
    '  Groups, frames, rectangles, and icons

Public Const grp1 = &H430
Public Const grp2 = &H431
Public Const grp3 = &H432
Public Const grp4 = &H433
Public Const frm1 = &H434
Public Const frm2 = &H435
Public Const frm3 = &H436
Public Const frm4 = &H437
Public Const rct1 = &H438
Public Const rct2 = &H439
Public Const rct3 = &H43A
Public Const rct4 = &H43B
Public Const ico1 = &H43C
Public Const ico2 = &H43D
Public Const ico3 = &H43E
Public Const ico4 = &H43F
    '  Static text

Public Const stc1 = &H440
Public Const stc2 = &H441
Public Const stc3 = &H442
Public Const stc4 = &H443
Public Const stc5 = &H444
Public Const stc6 = &H445
Public Const stc7 = &H446
Public Const stc8 = &H447
Public Const stc9 = &H448
Public Const stc10 = &H449
Public Const stc11 = &H44A
Public Const stc12 = &H44B
Public Const stc13 = &H44C
Public Const stc14 = &H44D
Public Const stc15 = &H44E
Public Const stc16 = &H44F
Public Const stc17 = &H450
Public Const stc18 = &H451
Public Const stc19 = &H452
Public Const stc20 = &H453
Public Const stc21 = &H454
Public Const stc22 = &H455
Public Const stc23 = &H456
Public Const stc24 = &H457
Public Const stc25 = &H458
Public Const stc26 = &H459
Public Const stc27 = &H45A
Public Const stc28 = &H45B
Public Const stc29 = &H45C
Public Const stc30 = &H45D
Public Const stc31 = &H45E
Public Const stc32 = &H45F
    '  Listboxes

Public Const lst1 = &H460
Public Const lst2 = &H461
Public Const lst3 = &H462
Public Const lst4 = &H463
Public Const lst5 = &H464
Public Const lst6 = &H465
Public Const lst7 = &H466
Public Const lst8 = &H467
Public Const lst9 = &H468
Public Const lst10 = &H469
Public Const lst11 = &H46A
Public Const lst12 = &H46B
Public Const lst13 = &H46C
Public Const lst14 = &H46D
Public Const lst15 = &H46E
Public Const lst16 = &H46F
    '  Combo boxes

Public Const cmb1 = &H470
Public Const cmb2 = &H471
Public Const cmb3 = &H472
Public Const cmb4 = &H473
Public Const cmb5 = &H474
Public Const cmb6 = &H475
Public Const cmb7 = &H476
Public Const cmb8 = &H477
Public Const cmb9 = &H478
Public Const cmb10 = &H479
Public Const cmb11 = &H47A
Public Const cmb12 = &H47B
Public Const cmb13 = &H47C
Public Const cmb14 = &H47D
Public Const cmb15 = &H47E
Public Const cmb16 = &H47F
    '  Edit controls

Public Const edt1 = &H480
Public Const edt2 = &H481
Public Const edt3 = &H482
Public Const edt4 = &H483
Public Const edt5 = &H484
Public Const edt6 = &H485
Public Const edt7 = &H486
Public Const edt8 = &H487
Public Const edt9 = &H488
Public Const edt10 = &H489
Public Const edt11 = &H48A
Public Const edt12 = &H48B
Public Const edt13 = &H48C
Public Const edt14 = &H48D
Public Const edt15 = &H48E
Public Const edt16 = &H48F
    '  Scroll bars

Public Const scr1 = &H490
Public Const scr2 = &H491
Public Const scr3 = &H492
Public Const scr4 = &H493
Public Const scr5 = &H494
Public Const scr6 = &H495
Public Const scr7 = &H496
Public Const scr8 = &H497
Public Const FILEOPENORD = 1536
Public Const MULTIFILEOPENORD = 1537
Public Const PRINTDLGORD = 1538
Public Const PRNSETUPDLGORD = 1539
Public Const FINDDLGORD = 1540
Public Const REPLACEDLGORD = 1541
Public Const FONTDLGORD = 1542
Public Const FORMATDLGORD31 = 1543
Public Const FORMATDLGORD30 = 1544
' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
' Service database names

Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
Public Const SERVICES_FAILED_DATABASE = "ServicesFailed"
' Value to indicate no change to an optional parameter

Public Const SERVICE_NO_CHANGE = &HFFFF
' Service State -- for Enum Requests (Bit Mask)

Public Const SERVICE_ACTIVE = &H1
Public Const SERVICE_INACTIVE = &H2
Public Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)
' Controls

Public Const SERVICE_CONTROL_STOP = &H1
Public Const SERVICE_CONTROL_PAUSE = &H2
Public Const SERVICE_CONTROL_CONTINUE = &H3
Public Const SERVICE_CONTROL_INTERROGATE = &H4
Public Const SERVICE_CONTROL_SHUTDOWN = &H5
' Service State -- for CurrentState

Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7
' Controls Accepted  (Bit Mask)

Public Const SERVICE_ACCEPT_STOP = &H1
Public Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Public Const SERVICE_ACCEPT_SHUTDOWN = &H4
' Service Control Manager object specific access types

Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
' Service object specific access type

Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
' ++ BUILD Version: 0010    '  Increment this if a change has global effects
' Copyright (c) 1995  Microsoft Corporation
' Module Name:
'     winsvc.h
' Abstract:
'     Header file for the Service Control Manager
' Environment:
'     User Mode - Win32
' --*/
'
'  Constants
'  Character to designate that a name is a group
'

Public Const SC_GROUP_IDENTIFIER = "+"
' Section for Performance Monitor data

Public Const PERF_DATA_VERSION = 1
Public Const PERF_DATA_REVISION = 1
Public Const PERF_NO_INSTANCES = -1  '  no instances
' The counter type is the "or" of the following values as described below
'
' select one of the following to indicate the counter's data size

Public Const PERF_SIZE_DWORD = &H0
Public Const PERF_SIZE_LARGE = &H100
Public Const PERF_SIZE_ZERO = &H200       '  for Zero Length fields
Public Const PERF_SIZE_VARIABLE_LEN = &H300       '  length is in CounterLength field of Counter Definition struct
' select one of the following values to indicate the counter field usage

Public Const PERF_TYPE_NUMBER = &H0         '  a number (not a counter)
Public Const PERF_TYPE_COUNTER = &H400       '  an increasing numeric value
Public Const PERF_TYPE_TEXT = &H800       '  a text field
Public Const PERF_TYPE_ZERO = &HC00       '  displays a zero
' If the PERF_TYPE_NUMBER field was selected, then select one of the
' following to describe the Number

Public Const PERF_NUMBER_HEX = &H0         '  display as HEX value
Public Const PERF_NUMBER_DECIMAL = &H10000     '  display as a decimal integer
Public Const PERF_NUMBER_DEC_1000 = &H20000     '  display as a decimal/1000
'
' If the PERF_TYPE_COUNTER value was selected then select one of the
' following to indicate the type of counter

Public Const PERF_COUNTER_VALUE = &H0         '  display counter value
Public Const PERF_COUNTER_RATE = &H10000     '  divide ctr / delta time
Public Const PERF_COUNTER_FRACTION = &H20000     '  divide ctr / base
Public Const PERF_COUNTER_BASE = &H30000     '  base value used in fractions
Public Const PERF_COUNTER_ELAPSED = &H40000     '  subtract counter from current time
Public Const PERF_COUNTER_QUEUELEN = &H50000     '  Use Queuelen processing func.
Public Const PERF_COUNTER_HISTOGRAM = &H60000     '  Counter begins or ends a histogram
' If the PERF_TYPE_TEXT value was selected, then select one of the
' following to indicate the type of TEXT data.

Public Const PERF_TEXT_UNICODE = &H0         '  type of text in text field
Public Const PERF_TEXT_ASCII = &H10000     '  ASCII using the CodePage field
' Timer SubTypes

Public Const PERF_TIMER_TICK = &H0         '  use system perf. freq for base
Public Const PERF_TIMER_100NS = &H100000    '  use 100 NS timer time base units
Public Const PERF_OBJECT_TIMER = &H200000    '  use the object timer freq
' Any types that have calculations performed can use one or more of
' the following calculation modification flags listed here

Public Const PERF_DELTA_COUNTER = &H400000    '  compute difference first
Public Const PERF_DELTA_BASE = &H800000    '  compute base diff as well
Public Const PERF_INVERSE_COUNTER = &H1000000   '  show as 1.00-value (assumes:
Public Const PERF_MULTI_COUNTER = &H2000000   '  sum of multiple instances
' Select one of the following values to indicate the display suffix (if any)

Public Const PERF_DISPLAY_NO_SUFFIX = &H0         '  no suffix
Public Const PERF_DISPLAY_PER_SEC = &H10000000  '  "/sec"
Public Const PERF_DISPLAY_PERCENT = &H20000000  '  "%"
Public Const PERF_DISPLAY_SECONDS = &H30000000  '  "secs"
Public Const PERF_DISPLAY_NOSHOW = &H40000000  '  value is not displayed
' Predefined counter types
' 32-bit Counter.  Divide delta by delta time.  Display suffix: "/sec"

Public Const PERF_COUNTER_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)
' 64-bit Timer.  Divide delta by delta time.  Display suffix: "%"

Public Const PERF_COUNTER_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)
' Queue Length Space-Time Product. Divide delta by delta time. No Display Suffix.

Public Const PERF_COUNTER_QUEUELEN_TYPE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_QUEUELEN Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)
' 64-bit Counter.  Divide delta by delta time. Display Suffix: "/sec"

Public Const PERF_COUNTER_BULK_COUNT = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)
' Indicates the counter is not a  counter but rather Unicode text Display as text.

Public Const PERF_COUNTER_TEXT = (PERF_SIZE_VARIABLE_LEN Or PERF_TYPE_TEXT Or PERF_TEXT_UNICODE Or PERF_DISPLAY_NO_SUFFIX)
' Indicates the data is a counter  which should not be
' time averaged on display (such as an error counter on a serial line)
' Display as is.  No Display Suffix.

Public Const PERF_COUNTER_RAWCOUNT = (PERF_SIZE_DWORD Or PERF_TYPE_NUMBER Or PERF_NUMBER_DECIMAL Or PERF_DISPLAY_NO_SUFFIX)
' A count which is either 1 or 0 on each sampling interrupt (% busy)
' Divide delta by delta base. Display Suffix: "%"

Public Const PERF_SAMPLE_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DELTA_COUNTER Or PERF_DELTA_BASE Or PERF_DISPLAY_PERCENT)
' A count which is sampled on each sampling interrupt (queue length)
' Divide delta by delta time. No Display Suffix.

Public Const PERF_SAMPLE_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)
' A label: no data is associated with this counter (it has 0 length)
' Do not display.

Public Const PERF_COUNTER_NODATA = (PERF_SIZE_ZERO Or PERF_DISPLAY_NOSHOW)
' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100 - delta divided by delta time.  Display suffix: "%"

Public Const PERF_COUNTER_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
' The divisor for a sample, used with the previous counter to form a
' sampled %.  You must check for >0 before dividing by this!  This
' counter will directly follow the  numerator counter.  It should not
' be displayed to the user.

Public Const PERF_SAMPLE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H1)         '  for compatibility with pre-beta versions
' A timer which, when divided by an average base, produces a time
' in seconds which is the average time of some operation.  This
' timer times total operations, and  the base is the number of opera-
' tions.  Display Suffix: "sec"

Public Const PERF_AVERAGE_TIMER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_SECONDS)
' Used as the denominator in the computation of time or count
' averages.  Must directly follow the numerator counter.  Not dis-
' played to the user.

Public Const PERF_AVERAGE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H2)         '  for compatibility with pre-beta versions
' A bulk count which, when divided (typically) by the number of
' operations, gives (typically) the number of bytes per operation.
' No Display Suffix.

Public Const PERF_AVERAGE_BULK = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_NOSHOW)
' 64-bit Timer in 100 nsec units. Display delta divided by
' delta time.  Display suffix: "%"

Public Const PERF_100NSEC_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)
' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100 - delta divided by delta time.  Display suffix: "%"

Public Const PERF_100NSEC_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
' 64-bit Timer.  Divide delta by delta time.  Display suffix: "%"
' Timer for multiple instances, so result can exceed 100%.

Public Const PERF_COUNTER_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_TIMER_TICK Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)
' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100  _MULTI_BASE - delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result
' can exceed 100%.  Followed by a counter of type _MULTI_BASE.

Public Const PERF_COUNTER_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_MULTI_COUNTER Or PERF_TIMER_TICK Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
' Number of instances to which the preceding _MULTI_..._INV counter
' applies.  Used as a factor to get the percentage.

Public Const PERF_COUNTER_MULTI_BASE = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_MULTI_COUNTER Or PERF_DISPLAY_NOSHOW)
' 64-bit Timer in 100 nsec units. Display delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result can exceed 100%.

Public Const PERF_100NSEC_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)
' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100  _MULTI_BASE - delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result
' can exceed 100%.  Followed by a counter of type _MULTI_BASE.

Public Const PERF_100NSEC_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
' Indicates the data is a fraction of the following counter  which
' should not be time averaged on display (such as free space over
' total space.) Display as is.  Display the quotient as "%".

Public Const PERF_RAW_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_PERCENT)
' Indicates the data is a base for the preceding counter which should
' not be time averaged on display (such as free space over total space.)

Public Const PERF_RAW_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H3)         '  for compatibility with pre-beta versions
' The data collected in this counter is actually the start time of the
' item being measured. For display, this data is subtracted from the
' sample time to yield the elapsed time as the difference between the two.
' In the definition below, the PerfTime field of the Object contains
' the sample time as indicated by the PERF_OBJECT_TIMER bit and the
' difference is scaled by the PerfFreq of the Object to convert the time
' units into seconds.

Public Const PERF_ELAPSED_TIME = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_ELAPSED Or PERF_OBJECT_TIMER Or PERF_DISPLAY_SECONDS)
' The following counter type can be used with the preceding types to
' define a range of values to be displayed in a histogram.

Public Const PERF_COUNTER_HISTOGRAM_TYPE = &H80000000  ' Counter begins or ends a histogram
' The following are used to determine the level of detail associated
' with the counter.  The user will be setting the level of detail
' that should be displayed at any given time.

Public Const PERF_DETAIL_NOVICE = 100 '  The uninformed can understand it
Public Const PERF_DETAIL_ADVANCED = 200 '  For the advanced user
Public Const PERF_DETAIL_EXPERT = 300 '  For the expert user
Public Const PERF_DETAIL_WIZARD = 400 '  For the system designer
Public Const PERF_NO_UNIQUE_ID = -1
Public Const CDERR_DIALOGFAILURE = &HFFFF
Public Const CDERR_GENERALCODES = &H0
Public Const CDERR_STRUCTSIZE = &H1
Public Const CDERR_INITIALIZATION = &H2
Public Const CDERR_NOTEMPLATE = &H3
Public Const CDERR_NOHINSTANCE = &H4
Public Const CDERR_LOADSTRFAILURE = &H5
Public Const CDERR_FINDRESFAILURE = &H6
Public Const CDERR_LOADRESFAILURE = &H7
Public Const CDERR_LOCKRESFAILURE = &H8
Public Const CDERR_MEMALLOCFAILURE = &H9
Public Const CDERR_MEMLOCKFAILURE = &HA
Public Const CDERR_NOHOOK = &HB
Public Const CDERR_REGISTERMSGFAIL = &HC
Public Const PDERR_PRINTERCODES = &H1000
Public Const PDERR_SETUPFAILURE = &H1001
Public Const PDERR_PARSEFAILURE = &H1002
Public Const PDERR_RETDEFFAILURE = &H1003
Public Const PDERR_LOADDRVFAILURE = &H1004
Public Const PDERR_GETDEVMODEFAIL = &H1005
Public Const PDERR_INITFAILURE = &H1006
Public Const PDERR_NODEVICES = &H1007
Public Const PDERR_NODEFAULTPRN = &H1008
Public Const PDERR_DNDMMISMATCH = &H1009
Public Const PDERR_CREATEICFAILURE = &H100A
Public Const PDERR_PRINTERNOTFOUND = &H100B
Public Const PDERR_DEFAULTDIFFERENT = &H100C
Public Const CFERR_CHOOSEFONTCODES = &H2000
Public Const CFERR_NOFONTS = &H2001
Public Const CFERR_MAXLESSTHANMIN = &H2002
Public Const FNERR_FILENAMECODES = &H3000
Public Const FNERR_SUBCLASSFAILURE = &H3001
Public Const FNERR_INVALIDFILENAME = &H3002
Public Const FNERR_BUFFERTOOSMALL = &H3003
Public Const FRERR_FINDREPLACECODES = &H4000
Public Const FRERR_BUFFERLENGTHZERO = &H4001
Public Const CCERR_CHOOSECOLORCODES = &H5000
' Public interface to LZEXP?.LIB
'  LZEXPAND error return codes

Public Const LZERROR_BADINHANDLE = (-1)  '  invalid input handle
Public Const LZERROR_BADOUTHANDLE = (-2) '  invalid output handle
Public Const LZERROR_READ = (-3)         '  corrupt compressed file format
Public Const LZERROR_WRITE = (-4)        '  out of space for output file
Public Const LZERROR_PUBLICLOC = (-5)    '  insufficient memory for LZFile struct
Public Const LZERROR_GLOBLOCK = (-6)     '  bad Global handle
Public Const LZERROR_BADVALUE = (-7)     '  input parameter out of range
Public Const LZERROR_UNKNOWNALG = (-8)   '  compression algorithm not recognized
' ********************************************************************
'       IMM.H - Input Method Manager definitions
'
'       Copyright (c) 1993-1995  Microsoft Corporation
' ********************************************************************

Public Const VK_PROCESSKEY = &HE5
Public Const STYLE_DESCRIPTION_SIZE = 32
'  the IME related messages

Public Const WM_CONVERTREQUESTEX = &H108
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
'  wParam for WM_IME_CONTROL

Public Const IMC_GETCANDIDATEPOS = &H7
Public Const IMC_SETCANDIDATEPOS = &H8
Public Const IMC_GETCOMPOSITIONFONT = &H9
Public Const IMC_SETCOMPOSITIONFONT = &HA
Public Const IMC_GETCOMPOSITIONWINDOW = &HB
Public Const IMC_SETCOMPOSITIONWINDOW = &HC
Public Const IMC_GETSTATUSWINDOWPOS = &HF
Public Const IMC_SETSTATUSWINDOWPOS = &H10
Public Const IMC_CLOSESTATUSWINDOW = &H21
Public Const IMC_OPENSTATUSWINDOW = &H22
'  wParam for WM_IME_CONTROL to the soft keyboard
'  dwAction for ImmNotifyIME

Public Const NI_OPENCANDIDATE = &H10
Public Const NI_CLOSECANDIDATE = &H11
Public Const NI_SELECTCANDIDATESTR = &H12
Public Const NI_CHANGECANDIDATELIST = &H13
Public Const NI_FINALIZECONVERSIONRESULT = &H14
Public Const NI_COMPOSITIONSTR = &H15
Public Const NI_SETCANDIDATE_PAGESTART = &H16
Public Const NI_SETCANDIDATE_PAGESIZE = &H17
'  lParam for WM_IME_SETCONTEXT

Public Const ISC_SHOWUICANDIDATEWINDOW = &H1
Public Const ISC_SHOWUICOMPOSITIONWINDOW = &H80000000
Public Const ISC_SHOWUIGUIDELINE = &H40000000
Public Const ISC_SHOWUIALLCANDIDATEWINDOW = &HF
Public Const ISC_SHOWUIALL = &HC000000F
'  dwIndex for ImmNotifyIME/NI_COMPOSITIONSTR

Public Const CPS_COMPLETE = &H1
Public Const CPS_CONVERT = &H2
Public Const CPS_REVERT = &H3
Public Const CPS_CANCEL = &H4
'  Windows for Simplified Chinese Edition hot key ID from 0x10 - 0x2F

Public Const IME_CHOTKEY_IME_NONIME_TOGGLE = &H10
Public Const IME_CHOTKEY_SHAPE_TOGGLE = &H11
Public Const IME_CHOTKEY_SYMBOL_TOGGLE = &H12
'  Windows for Japanese Edition hot key ID from 0x30 - 0x4F

Public Const IME_JHOTKEY_CLOSE_OPEN = &H30
'  Windows for Korean Edition hot key ID from 0x50 - 0x6F

Public Const IME_KHOTKEY_SHAPE_TOGGLE = &H50
Public Const IME_KHOTKEY_HANJACONVERT = &H51
Public Const IME_KHOTKEY_ENGLISH = &H52
'  Windows for Tranditional Chinese Edition hot key ID from 0x70 - 0x8F

Public Const IME_THOTKEY_IME_NONIME_TOGGLE = &H70
Public Const IME_THOTKEY_SHAPE_TOGGLE = &H71
Public Const IME_THOTKEY_SYMBOL_TOGGLE = &H72
'  direct switch hot key ID from 0x100 - 0x11F

Public Const IME_HOTKEY_DSWITCH_FIRST = &H100
Public Const IME_HOTKEY_DSWITCH_LAST = &H11F
'  IME private hot key from 0x200 - 0x21F

Public Const IME_ITHOTKEY_RESEND_RESULTSTR = &H200
Public Const IME_ITHOTKEY_PREVIOUS_COMPOSITION = &H201
Public Const IME_ITHOTKEY_UISTYLE_TOGGLE = &H202
'  parameter of ImmGetCompositionString

Public Const GCS_COMPREADSTR = &H1
Public Const GCS_COMPREADATTR = &H2
Public Const GCS_COMPREADCLAUSE = &H4
Public Const GCS_COMPSTR = &H8
Public Const GCS_COMPATTR = &H10
Public Const GCS_COMPCLAUSE = &H20
Public Const GCS_CURSORPOS = &H80
Public Const GCS_DELTASTART = &H100
Public Const GCS_RESULTREADSTR = &H200
Public Const GCS_RESULTREADCLAUSE = &H400
Public Const GCS_RESULTSTR = &H800
Public Const GCS_RESULTCLAUSE = &H1000
'  style bit flags for WM_IME_COMPOSITION

Public Const CS_INSERTCHAR = &H2000
Public Const CS_NOMOVECARET = &H4000
'  bits of fdwInit of INPUTCONTEXT
'  IME property bits

Public Const IME_PROP_AT_CARET = &H10000
Public Const IME_PROP_SPECIAL_UI = &H20000
Public Const IME_PROP_CANDLIST_START_FROM_1 = &H40000
Public Const IME_PROP_UNICODE = &H80000
'  IME UICapability bits

Public Const UI_CAP_2700 = &H1
Public Const UI_CAP_ROT90 = &H2
Public Const UI_CAP_ROTANY = &H4
'  ImmSetCompositionString Capability bits

Public Const SCS_CAP_COMPSTR = &H1
Public Const SCS_CAP_MAKEREAD = &H2
'  IME WM_IME_SELECT inheritance Capability bits

Public Const SELECT_CAP_CONVERSION = &H1
Public Const SELECT_CAP_SENTENCE = &H2
'  ID for deIndex of ImmGetGuideLine

Public Const GGL_LEVEL = &H1
Public Const GGL_INDEX = &H2
Public Const GGL_STRING = &H3
Public Const GGL_PRIVATE = &H4
'  ID for dwLevel of GUIDELINE Structure

Public Const GL_LEVEL_NOGUIDELINE = &H0
Public Const GL_LEVEL_FATAL = &H1
Public Const GL_LEVEL_ERROR = &H2
Public Const GL_LEVEL_WARNING = &H3
Public Const GL_LEVEL_INFORMATION = &H4
'  ID for dwIndex of GUIDELINE Structure

Public Const GL_ID_UNKNOWN = &H0
Public Const GL_ID_NOMODULE = &H1
Public Const GL_ID_NODICTIONARY = &H10
Public Const GL_ID_CANNOTSAVE = &H11
Public Const GL_ID_NOCONVERT = &H20
Public Const GL_ID_TYPINGERROR = &H21
Public Const GL_ID_TOOMANYSTROKE = &H22
Public Const GL_ID_READINGCONFLICT = &H23
Public Const GL_ID_INPUTREADING = &H24
Public Const GL_ID_INPUTRADICAL = &H25
Public Const GL_ID_INPUTCODE = &H26
Public Const GL_ID_INPUTSYMBOL = &H27
Public Const GL_ID_CHOOSECANDIDATE = &H28
Public Const GL_ID_REVERSECONVERSION = &H29
Public Const GL_ID_PRIVATE_FIRST = &H8000
Public Const GL_ID_PRIVATE_LAST = &HFFFF
'  ID for dwIndex of ImmGetProperty

Public Const IGP_PROPERTY = &H4
Public Const IGP_CONVERSION = &H8
Public Const IGP_SENTENCE = &HC
Public Const IGP_UI = &H10
Public Const IGP_SETCOMPSTR = &H14
Public Const IGP_SELECT = &H18
'  dwIndex for ImmSetCompositionString API

Public Const SCS_SETSTR = (GCS_COMPREADSTR Or GCS_COMPSTR)
Public Const SCS_CHANGEATTR = (GCS_COMPREADATTR Or GCS_COMPATTR)
Public Const SCS_CHANGECLAUSE = (GCS_COMPREADCLAUSE Or GCS_COMPCLAUSE)
'  attribute for COMPOSITIONSTRING Structure

Public Const ATTR_INPUT = &H0
Public Const ATTR_TARGET_CONVERTED = &H1
Public Const ATTR_CONVERTED = &H2
Public Const ATTR_TARGET_NOTCONVERTED = &H3
Public Const ATTR_INPUT_ERROR = &H4
'  bit field for IMC_SETCOMPOSITIONWINDOW, IMC_SETCANDIDATEWINDOW

Public Const CFS_DEFAULT = &H0
Public Const CFS_RECT = &H1
Public Const CFS_POINT = &H2
Public Const CFS_SCREEN = &H4
Public Const CFS_FORCE_POSITION = &H20
Public Const CFS_CANDIDATEPOS = &H40
Public Const CFS_EXCLUDE = &H80
'  conversion direction for ImmGetConversionList

Public Const GCL_CONVERSION = &H1
Public Const GCL_REVERSECONVERSION = &H2
Public Const GCL_REVERSE_LENGTH = &H3
'  bit field for conversion mode

Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_CHINESE = IME_CMODE_NATIVE
Public Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Public Const IME_CMODE_JAPANESE = IME_CMODE_NATIVE
Public Const IME_CMODE_KATAKANA = &H2                   '  only effect under IME_CMODE_NATIVE
Public Const IME_CMODE_LANGUAGE = &H3
Public Const IME_CMODE_FULLSHAPE = &H8
Public Const IME_CMODE_ROMAN = &H10
Public Const IME_CMODE_CHARCODE = &H20
Public Const IME_CMODE_HANJACONVERT = &H40
Public Const IME_CMODE_SOFTKBD = &H80
Public Const IME_CMODE_NOCONVERSION = &H100
Public Const IME_CMODE_EUDC = &H200
Public Const IME_CMODE_SYMBOL = &H400
Public Const IME_SMODE_NONE = &H0
Public Const IME_SMODE_PLAURALCLAUSE = &H1
Public Const IME_SMODE_SINGLECONVERT = &H2
Public Const IME_SMODE_AUTOMATIC = &H4
Public Const IME_SMODE_PHRASEPREDICT = &H8
'  style of candidate

Public Const IME_CAND_UNKNOWN = &H0
Public Const IME_CAND_READ = &H1
Public Const IME_CAND_CODE = &H2
Public Const IME_CAND_MEANING = &H3
Public Const IME_CAND_RADICAL = &H4
Public Const IME_CAND_STROKE = &H5
'  wParam of report message WM_IME_NOTIFY

Public Const IMN_CLOSESTATUSWINDOW = &H1
Public Const IMN_OPENSTATUSWINDOW = &H2
Public Const IMN_CHANGECANDIDATE = &H3
Public Const IMN_CLOSECANDIDATE = &H4
Public Const IMN_OPENCANDIDATE = &H5
Public Const IMN_SETCONVERSIONMODE = &H6
Public Const IMN_SETSENTENCEMODE = &H7
Public Const IMN_SETOPENSTATUS = &H8
Public Const IMN_SETCANDIDATEPOS = &H9
Public Const IMN_SETCOMPOSITIONFONT = &HA
Public Const IMN_SETCOMPOSITIONWINDOW = &HB
Public Const IMN_SETSTATUSWINDOWPOS = &HC
Public Const IMN_GUIDELINE = &HD
Public Const IMN_PRIVATE = &HE
'  error code of ImmGetCompositionString

Public Const IMM_ERROR_NODATA = (-1)
Public Const IMM_ERROR_GENERAL = (-2)
'  dialog mode of ImmConfigureIME

Public Const IME_CONFIG_GENERAL = 1
Public Const IME_CONFIG_REGISTERWORD = 2
Public Const IME_CONFIG_SELECTDICTIONARY = 3
'  dialog mode of ImmEscape

Public Const IME_ESC_QUERY_SUPPORT = &H3
Public Const IME_ESC_RESERVED_FIRST = &H4
Public Const IME_ESC_RESERVED_LAST = &H7FF
Public Const IME_ESC_PRIVATE_FIRST = &H800
Public Const IME_ESC_PRIVATE_LAST = &HFFF
Public Const IME_ESC_SEQUENCE_TO_INTERNAL = &H1001
Public Const IME_ESC_GET_EUDC_DICTIONARY = &H1003
Public Const IME_ESC_SET_EUDC_DICTIONARY = &H1004
Public Const IME_ESC_MAX_KEY = &H1005
Public Const IME_ESC_IME_NAME = &H1006
Public Const IME_ESC_SYNC_HOTKEY = &H1007
Public Const IME_ESC_HANJA_MODE = &H1008
'  style of word registration

Public Const IME_REGWORD_STYLE_EUDC = &H1
Public Const IME_REGWORD_STYLE_USER_FIRST = &H80000000
Public Const IME_REGWORD_STYLE_USER_LAST = &HFFFF
'  type of soft keyboard
'  for Windows Tranditional Chinese Edition

Public Const SOFTKEYBOARD_TYPE_T1 = &H1
'  for Windows Simplified Chinese Edition

Public Const SOFTKEYBOARD_TYPE_C1 = &H2
'  Dial Options

Public Const DIALOPTION_BILLING = &H40          '  Supports wait for bong "$"
Public Const DIALOPTION_QUIET = &H80            '  Supports wait for quiet "@"
Public Const DIALOPTION_DIALTONE = &H100        '  Supports wait for dial tone "W"
'  SpeakerVolume for MODEMDEVCAPS

Public Const MDMVOLFLAG_LOW = &H1
Public Const MDMVOLFLAG_MEDIUM = &H2
Public Const MDMVOLFLAG_HIGH = &H4
'  SpeakerVolume for MODEMSETTINGS

Public Const MDMVOL_LOW = &H0
Public Const MDMVOL_MEDIUM = &H1
Public Const MDMVOL_HIGH = &H2
'  SpeakerMode for MODEMDEVCAPS

Public Const MDMSPKRFLAG_OFF = &H1
Public Const MDMSPKRFLAG_DIAL = &H2
Public Const MDMSPKRFLAG_ON = &H4
Public Const MDMSPKRFLAG_CALLSETUP = &H8
'  SpeakerMode for MODEMSETTINGS

Public Const MDMSPKR_OFF = &H0
Public Const MDMSPKR_DIAL = &H1
Public Const MDMSPKR_ON = &H2
Public Const MDMSPKR_CALLSETUP = &H3
'  Modem Options

Public Const MDM_COMPRESSION = &H1
Public Const MDM_ERROR_CONTROL = &H2
Public Const MDM_FORCED_EC = &H4
Public Const MDM_CELLULAR = &H8
Public Const MDM_FLOWCONTROL_HARD = &H10
Public Const MDM_FLOWCONTROL_SOFT = &H20
Public Const MDM_CCITT_OVERRIDE = &H40
Public Const MDM_SPEED_ADJUST = &H80
Public Const MDM_TONE_DIAL = &H100
Public Const MDM_BLIND_DIAL = &H200
Public Const MDM_V23_OVERRIDE = &H400
' // AppBar stuff

Public Const ABM_NEW = &H0
Public Const ABM_REMOVE = &H1
Public Const ABM_QUERYPOS = &H2
Public Const ABM_SETPOS = &H3
Public Const ABM_GETSTATE = &H4
Public Const ABM_GETTASKBARPOS = &H5
Public Const ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
Public Const ABM_GETAUTOHIDEBAR = &H7
Public Const ABM_SETAUTOHIDEBAR = &H8          '  this can fail at any time.  MUST check the result
                                        '  lParam = TRUE/FALSE  Set/Unset
                                        '  uEdge = what edge

Public Const ABM_WINDOWPOSCHANGED = &H9
'  these are put in the wparam of callback messages

Public Const ABN_STATECHANGE = &H0
Public Const ABN_POSCHANGED = &H1
Public Const ABN_FULLSCREENAPP = &H2
Public Const ABN_WINDOWARRANGE = &H3 '  lParam == TRUE means hide
'  flags for get state

Public Const ABS_AUTOHIDE = &H1
Public Const ABS_ALWAYSONTOP = &H2
Public Const ABE_LEFT = 0
Public Const ABE_TOP = 1
Public Const ABE_RIGHT = 2
Public Const ABE_BOTTOM = 3
Public Const EIRESID = -1
' // Shell File Operations

Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_SILENT = &H4                      '  don't create progress/report
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Public Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings

Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Public Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Public Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Public Const PO_DELETE = &H13           '  printer is being deleted
Public Const PO_RENAME = &H14           '  printer is being renamed
Public Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
                                '  if this id is set, the strings received by
                                '  the copyhook are a doubly-null terminated
                                '  list of strings.  The first is the printer
                                '  name and the second is the printer port.

Public Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.
' // End Shell File Operations
' //  Begin ShellExecuteEx and family
'  ShellExecute() and ShellExecuteEx() error codes
'  regular WinExec() codes

Public Const SE_ERR_FNF = 2                     '  file not found
Public Const SE_ERR_PNF = 3                     '  path not found
Public Const SE_ERR_ACCESSDENIED = 5            '  access denied
Public Const SE_ERR_OOM = 8                     '  out of memory
Public Const SE_ERR_DLLNOTFOUND = 32
'  Note CLASSKEY overrides CLASSNAME

Public Const SEE_MASK_CLASSNAME = &H1
Public Const SEE_MASK_CLASSKEY = &H3
'  Note INVOKEIDLIST overrides IDLIST

Public Const SEE_MASK_IDLIST = &H4
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_ICON = &H10
Public Const SEE_MASK_HOTKEY = &H20
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_CONNECTNETDRV = &H80
Public Const SEE_MASK_FLAG_DDEWAIT = &H100
Public Const SEE_MASK_DOENVSUBST = &H200
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_OPENICON = &H2                       '  get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
Public Const SHGNLI_PIDL = &H1                          '  pszLinkTo is a pidl
Public Const SHGNLI_PREFIXNAME = &H2                    '  Make name "Shortcut to xxx"
' // End SHGetFileInfo
' Copyright (C) 1993 - 1995 Microsoft Corporation
' Module Name:
'     winperf.h
' Abstract:
'     Header file for the Performance Monitor data.
'     This file contains the definitions of the data structures returned
'     by the Configuration Registry in response to a request for
'     performance data.  This file is used by both the Configuration
'     Registry and the Performance Monitor to define their interface.
'     The complete interface is described here, except for the name
'     of the node to query in the registry.  It is
'                    HKEY_PERFORMANCE_DATA.
'     By querying that node with a subkey of "Global" the caller will
'     retrieve the structures described here.
'     There is no need to RegOpenKey() the reserved handle HKEY_PERFORMANCE_DATA,
'     but the caller should RegCloseKey() the handle so that network transports
'     and drivers can be removed or installed (which cannot happen while
'     they are open for monitoring.)  Remote requests must first
'     RegConnectRegistry().
' --*/
'   Data structure definitions.
'   In order for data to be returned through the Configuration Registry
'   in a system-independent fashion, it must be self-describing.
'   In the following, all offsets are in bytes.
'
'   Data is returned through the Configuration Registry in a
'   a data block which begins with a _PERF_DATA_BLOCK structure.
'
'   The _PERF_DATA_BLOCK structure is followed by NumObjectTypes of
'   data sections, one for each type of object measured.  Each object
'   type section begins with a _PERF_OBJECT_TYPE structure.
' *****************************************************************************                                                                             *
' * winver.h -    Version management functions, types, and definitions          *
' *                                                                             *
' *               Include file for VER.DLL.  This library is                    *
' *               designed to allow version stamping of Windows executable files*
' *               and of special .VER files for DOS executable files.           *
' *                                                                             *
' *               Copyright (c) 1993, Microsoft Corp.  All rights reserved      *
' *                                                                             *
' \*****************************************************************************/
'  ----- Symbols -----

Public Const VS_VERSION_INFO = 1
Public Const VS_USER_DEFINED = 100
'  ----- VS_VERSION.dwFileFlags -----

Public Const VS_FFI_SIGNATURE = &HFEEF04BD
Public Const VS_FFI_STRUCVERSION = &H10000
Public Const VS_FFI_FILEFLAGSMASK = &H3F&
'  ----- VS_VERSION.dwFileFlags -----

Public Const VS_FF_DEBUG = &H1&
Public Const VS_FF_PRERELEASE = &H2&
Public Const VS_FF_PATCHED = &H4&
Public Const VS_FF_PRIVATEBUILD = &H8&
Public Const VS_FF_INFOINFERRED = &H10&
Public Const VS_FF_SPECIALBUILD = &H20&
'  ----- VS_VERSION.dwFileOS -----

Public Const VOS_UNKNOWN = &H0&
Public Const VOS_DOS = &H10000
Public Const VOS_OS216 = &H20000
Public Const VOS_OS232 = &H30000
Public Const VOS_NT = &H40000
Public Const VOS__BASE = &H0&
Public Const VOS__WINDOWS16 = &H1&
Public Const VOS__PM16 = &H2&
Public Const VOS__PM32 = &H3&
Public Const VOS__WINDOWS32 = &H4&
Public Const VOS_DOS_WINDOWS16 = &H10001
Public Const VOS_DOS_WINDOWS32 = &H10004
Public Const VOS_OS216_PM16 = &H20002
Public Const VOS_OS232_PM32 = &H30003
Public Const VOS_NT_WINDOWS32 = &H40004
'  ----- VS_VERSION.dwFileType -----

Public Const VFT_UNKNOWN = &H0&
Public Const VFT_APP = &H1&
Public Const VFT_DLL = &H2&
Public Const VFT_DRV = &H3&
Public Const VFT_FONT = &H4&
Public Const VFT_VXD = &H5&
Public Const VFT_STATIC_LIB = &H7&
'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----

Public Const VFT2_UNKNOWN = &H0&
Public Const VFT2_DRV_PRINTER = &H1&
Public Const VFT2_DRV_KEYBOARD = &H2&
Public Const VFT2_DRV_LANGUAGE = &H3&
Public Const VFT2_DRV_DISPLAY = &H4&
Public Const VFT2_DRV_MOUSE = &H5&
Public Const VFT2_DRV_NETWORK = &H6&
Public Const VFT2_DRV_SYSTEM = &H7&
Public Const VFT2_DRV_INSTALLABLE = &H8&
Public Const VFT2_DRV_SOUND = &H9&
Public Const VFT2_DRV_COMM = &HA&
Public Const VFT2_DRV_INPUTMETHOD = &HB&
'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT -----

Public Const VFT2_FONT_RASTER = &H1&
Public Const VFT2_FONT_VECTOR = &H2&
Public Const VFT2_FONT_TRUETYPE = &H3&
'  ----- VerFindFile() flags -----

Public Const VFFF_ISSHAREDFILE = &H1
Public Const VFF_CURNEDEST = &H1
Public Const VFF_FILEINUSE = &H2
Public Const VFF_BUFFTOOSMALL = &H4
'  ----- VerInstallFile() flags -----

Public Const VIFF_FORCEINSTALL = &H1
Public Const VIFF_DONTDELETEOLD = &H2
Public Const VIF_TEMPFILE = &H1&
Public Const VIF_MISMATCH = &H2&
Public Const VIF_SRCOLD = &H4&
Public Const VIF_DIFFLANG = &H8&
Public Const VIF_DIFFCODEPG = &H10&
Public Const VIF_DIFFTYPE = &H20&
Public Const VIF_WRITEPROT = &H40&
Public Const VIF_FILEINUSE = &H80&
Public Const VIF_OUTOFSPACE = &H100&
Public Const VIF_ACCESSVIOLATION = &H200&
Public Const VIF_SHARINGVIOLATION = &H400&
Public Const VIF_CANNOTCREATE = &H800&
Public Const VIF_CANNOTDELETE = &H1000&
Public Const VIF_CANNOTRENAME = &H2000&
Public Const VIF_CANNOTDELETECUR = &H4000&
Public Const VIF_OUTOFMEMORY = &H8000&
Public Const VIF_CANNOTREADSRC = &H10000
Public Const VIF_CANNOTREADDST = &H20000
Public Const VIF_BUFFTOOSMALL = &H40000
Public Const PROCESS_HEAP_REGION = &H1
Public Const PROCESS_HEAP_UNCOMMITTED_RANGE = &H2
Public Const PROCESS_HEAP_ENTRY_BUSY = &H4
Public Const PROCESS_HEAP_ENTRY_MOVEABLE = &H10
Public Const PROCESS_HEAP_ENTRY_DDESHARE = &H20
'  GetBinaryType return values.

Public Const SCS_32BIT_BINARY = 0
Public Const SCS_DOS_BINARY = 1
Public Const SCS_WOW_BINARY = 2
Public Const SCS_PIF_BINARY = 3
Public Const SCS_POSIX_BINARY = 4
Public Const SCS_OS216_BINARY = 5
'  Logon Support APIs

Public Const LOGON32_LOGON_INTERACTIVE = 2
Public Const LOGON32_LOGON_BATCH = 4
Public Const LOGON32_LOGON_SERVICE = 5
Public Const LOGON32_PROVIDER_DEFAULT = 0
Public Const LOGON32_PROVIDER_WINNT35 = 1
'  dwPlatformId defines:
'

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
'  Power Management APIs

Public Const AC_LINE_OFFLINE = &H0
Public Const AC_LINE_ONLINE = &H1
Public Const AC_LINE_BACKUP_POWER = &H2
Public Const AC_LINE_UNKNOWN = &HFF
Public Const BATTERY_FLAG_HIGH = &H1
Public Const BATTERY_FLAG_LOW = &H2
Public Const BATTERY_FLAG_CRITICAL = &H4
Public Const BATTERY_FLAG_CHARGING = &H8
Public Const BATTERY_FLAG_NO_BATTERY = &H80
Public Const BATTERY_FLAG_UNKNOWN = &HFF
Public Const BATTERY_PERCENTAGE_UNKNOWN = &HFF
Public Const BATTERY_LIFE_UNKNOWN = &HFFFF
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const CDM_FIRST = (WM_USER + 100)
Public Const CDM_LAST = (WM_USER + 200)
Public Const CDM_GETSPEC = (CDM_FIRST + &H0)
Public Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Public Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Public Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
Public Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Public Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Public Const CDM_SETDEFEXT = (CDM_FIRST + &H6)
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100
Public Const FR_DOWN = &H1
Public Const FR_WHOLEWORD = &H2
Public Const FR_MATCHCASE = &H4
Public Const FR_FINDNEXT = &H8
Public Const FR_REPLACE = &H10
Public Const FR_REPLACEALL = &H20
Public Const FR_DIALOGTERM = &H40
Public Const FR_SHOWHELP = &H80
Public Const FR_ENABLEHOOK = &H100
Public Const FR_ENABLETEMPLATE = &H200
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOMATCHCASE = &H800
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_ENABLETEMPLATEHANDLE = &H2000
Public Const FR_HIDEUPDOWN = &H4000
Public Const FR_HIDEMATCHCASE = &H8000
Public Const FR_HIDEWHOLEWORD = &H10000
Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000
Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400
Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"
Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2
Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const DN_DEFAULTPRN = &H1
Public Const WM_PSD_PAGESETUPDLG = (WM_USER)
Public Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Public Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Public Const WM_PSD_MARGINRECT = (WM_USER + 3)
Public Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Public Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Public Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Public Const PSD_DEFAULTMINMARGINS = &H0 '  default (printer's)
Public Const PSD_INWININIINTLMEASURE = &H0 '  1st of 4 possible
Public Const PSD_MINMARGINS = &H1 '  use caller's
Public Const PSD_MARGINS = &H2 '  use caller's
Public Const PSD_INTHOUSANDTHSOFINCHES = &H4 '  2nd of 4 possible
Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8 '  3rd of 4 possible
Public Const PSD_DISABLEMARGINS = &H10
Public Const PSD_DISABLEPRINTER = &H20
Public Const PSD_NOWARNING = &H80 '  must be same as PD_*
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_RETURNDEFAULT = &H400 '  must be same as PD_*
Public Const PSD_DISABLEPAPER = &H200
Public Const PSD_SHOWHELP = &H800 '  must be same as PD_*
Public Const PSD_ENABLEPAGESETUPHOOK = &H2000 '  must be same as PD_*
Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000 '  must be same as PD_*
Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000 '  must be same as PD_*
Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Public Const PSD_DISABLEPAGEPAINTING = &H80000
Public Const INVALID_HANDLE_VALUE = -1
'DrawEdge Constants

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10
' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.

Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.
