VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'Option Explicit
'
'Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
'Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
'Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
'Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
'Private Declare Function CryptEncryptPtr Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
'
'Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
'Private Declare Function CryptDecryptPtr Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long) As Long
'
'Private Declare Function CryptDuplicateHash Lib "advapi32.dll" (ByVal hHash As Long, pdwReserved As Long, ByVal dwFlags As Long, phHash As Long) As Long
''Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
'Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
'Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
'Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
'Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
'Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
'Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptHashDataPtr Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
'Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
'Private Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
''Private Declare Function CryptSignMessage Lib "Crypt32.dll" (pSignPara As CRYPT_SIGN_MESSAGE_PARA, ByVal fDetachedSignature As Long, ByVal cToBeSigned As Long, rgpbToBeSigned As String, rgcbToBeSigned As Long, ByVal pbSignedBlob As String, pcbSignedBlob As Long) As Long
'Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
'Private Declare Function CryptVerifySignaturePtr Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, ByVal pbSignature As Long, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As Long, ByVal dwFlags As Long) As Long
''Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32.dll" (ByRef pCert As Long, ByVal dwFlags As Long, pvReserved As Any, ByRef phCryptProv As Long, ByRef pdwKeySpec As Long, ByRef pfCallerFreeProv As Long) As Long
'Private Declare Function CryptAcquireCertificatePrivateKey Lib "Crypt32.dll" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvReserved As Long, ByVal phCryptProv As Long, ByVal pdwKeySpec As Long, ByVal pfCallerFreeProv As Long) As Long
'
''If CryptAcquireCertificatePrivateKey(ptrCertCtx, CRYPT_ACQUIRE_CACHE_FLAG, 0, VarPtr(hCryptoProv), VarPtr(dwType), VarPtr(flag)) = 0 Then
''if CryptAcquireCertificatePrivateKey(hCertCntxt, 0, IntPtr.Zero, ref hCryptProv, ref keyspec, ref freecsp)
'
'
'Private Const CAQ As String = "CryptAcquireContext"
'Private Const cch As String = "CryptCreateHash"
'Private Const CDrK As String = "CryptDeriveKey"
'Private Const CDH As String = "CryptDestroyHash"
'Private Const CDsK As String = "CryptDestroyKey"
'Private Const CDuH As String = "CryptDuplicateHash"
'Private Const CE As String = "CryptEncrypt"
'Private Const CD As String = "CryptDecrypt"
'Private Const CEK As String = "CryptExportKey"
'Private Const CIK As String = "CryptImportKey"
'Private Const CGHP As String = "CryptGetHashParam"
'Private Const CGK As String = "CryptGenKey"
'Private Const CGR As String = "CryptGenRandom"
'Private Const CGPP As String = "CryptGetProvParam"
'Private Const CGUK As String = "CryptGetUserKey"
'Private Const CHD As String = "CryptHashData"
'Private Const CRC As String = "CryptReleaseContext"
'Private Const CSHP As String = "CryptSetHashParam"
'Private Const CSKP As String = "CryptSetKeyParam"
'Private Const CSH As String = "CryptSignHash"
'Private Const CVS As String = "CryptVerifySignature"
'Private Const abClientFinished = "client finished"
'Private Const abServerFinished = "server finished"
'
''================================
''EVENTS
''================================
'Public Event Error(ByVal Number As Long, Description As String, ByVal Source As String)
'
'' exported key blob definitions
'Private Const SIMPLEBLOB As Byte = &H1
'Private Const CUR_BLOB_VERSION As Byte = &H2
'Private Const PUBLICKEYBLOB As Byte = &H6
'Private Const PRIVATEKEYBLOB As Byte = &H7
'Private Const PLAINTEXTKEYBLOB As Byte = &H8
'
'Private Type T_PUBLICKEYBLOB
'    bType    As Byte
'    bVersion As Byte
'    Reserved As Integer
'    aiKeyAlg As Long
'    magic    As Long
'    bitlen   As Long
'    pubexp   As Long
'    modulus(1 To 256) As Byte
'End Type
'Private Const T_PUBLICKEYBLOBLEN = 276
'Private Type T_EXP_PUBLICKEYBLOB
'    bPublicKey(1 To T_PUBLICKEYBLOBLEN) As Byte
'End Type
'
'Private Type T_PRIVATEKEYBLOB
'    bType    As Byte
'    bVersion As Byte
'    Reserved As Integer
'    aiKeyAlg As Long
'    magic    As Long
'    bitlen   As Long
'    pubexp   As Long
'    modulus(1 To 256)  As Byte
'    prime1(1 To 128)   As Byte
'    prime2(1 To 128)   As Byte
'    exponent1(1 To 128)       As Byte
'    exponent2(1 To 128)       As Byte
'    coefficient(1 To 128)     As Byte
'    privateExponent(1 To 256) As Byte
'End Type
'Private Const T_PRIVATEKEYBLOBLEN = 1172
'Private Type T_EXP_PRIVATEKEYBLOB
'    bPrivateKey(1 To T_PRIVATEKEYBLOBLEN) As Byte
'End Type
'
'Private Type T_SIMPLEKEYBLOB
'    bType    As Byte
'    bVersion As Byte
'    Reserved As Integer
'    aiKeyAlg As Long
'    algid    As Long
'    modulus(1 To 256) As Byte
'End Type
'Private Const T_SIMPLEKEYBLOBLEN = 268
'Private Type T_EXP_SIMPLEKEYBLOB
'    bSimpleKey(1 To T_SIMPLEKEYBLOBLEN) As Byte
'End Type
'
'Private Type SCHANNEL_ALG
'    dwUse           As Long  '4 bytes
'    algid           As Long  '4 bytes
'    cBits           As Long  '4 bytes
'    dwFlags         As Long  '4 bytes
'    dwReserved      As Long  '4 bytes
'End Type                     'Total 20 bytes
'
'Private Type CRYPT_DATA_BLOB
'    cbData As Long          'The count of bytes in the buffer pointed to by pbData.
'    pbData As Long          'A pointer to a block of data bytes.
'End Type
'
'Private Type HMAC_INFO
'    ALG_ID As Long          'HashAlgid
'    pbIStr(1 To 64) As Byte '*pbInnerString (BYTE*)0x36
'    cbIStr As Long          'len(cbInnerString)
'    pbOStr(1 To 64) As Byte '*pbOuterString (BYTE*)0x5C
'    cbOStr As Long          'len(cbOuterString)
'End Type
'
'Private Type SHAMD5_HASH
'    bHash(1 To 36) As Byte
'End Type
'
''Class Variables
'Private sContainer As New Stream
'Private sProvider As New Stream
'Private sProvType As Long
'Private PublicKey As T_EXP_PUBLICKEYBLOB
'Private NewPubKey As T_PUBLICKEYBLOB
'Private PrivateKey As T_EXP_PRIVATEKEYBLOB
'Private MasterKey As T_EXP_SIMPLEKEYBLOB
'Private NewSessionKey As T_SIMPLEKEYBLOB
'Private PriFlg As Boolean
'Private fClient As Boolean
'Private bClientRandom() As Byte
'Private bServerRandom() As Byte
'Private bPreMasterKey() As Byte
'Private Alg_Hash As Long
'
''API error function
''Private Declare Function GetLastError Lib "kernel32" () As Long
'
''API memory functions
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
'Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
''Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
''Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
''Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
''Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
''Private Declare Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
''Private Declare Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
'
''constants for API memory functions
''Private Const GMEM_MOVEABLE = &H2
''Private Const GMEM_ZEROINIT = &H40
''Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
'Private Const CRYPT_IPSEC_HMAC_KEY = &H100
''constants for Cryptography API functions
'Private Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
'Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
'Private Const MS_STRONG_PROV As String = "Microsoft Strong Cryptographic Provider"
'Private Const MS_DEF_RSA_SIG_PROV As String = "Microsoft RSA Signature Cryptographic Provider"
'Private Const MS_DEF_RSA_SCHANNEL_PROV As String = "Microsoft RSA SChannel Cryptographic Provider"
'Private Const MS_DEF_DSS_PROV As String = "Microsoft Base DSS Cryptographic Provider"
'Private Const MS_DEF_DSS_DH_PROV As String = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
'Private Const MS_ENH_DSS_DH_PROV As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
'Private Const MS_DEF_DH_SCHANNEL_PROV As String = "Microsoft DH SChannel Cryptographic Provider"
'Private Const PROV_RSA_FULL As Long = 1
'Private Const PROV_RSA_SIG As Long = 2
'Private Const PROV_DSS As Long = 3
'Private Const PROV_FORTEZZA As Long = 4
'Private Const PROV_MS_MAIL As Long = 5
'Private Const PROV_SSL As Long = 6
'Private Const PROV_STT_MER As Long = 7
'Private Const PROV_STT_ACQ As Long = 8
'Private Const PROV_STT_BRND As Long = 9
'Private Const PROV_STT_ROOT As Long = 10
'Private Const PROV_STT_ISS As Long = 11
'Private Const PROV_RSA_SCHANNEL As Long = 12
'Private Const PROV_DSS_DH As Long = 13
'Private Const PROV_EC_ECDSA_SIG As Long = 14
'Private Const PROV_EC_ECNRA_SIG As Long = 15
'Private Const PROV_EC_ECDSA_FULL As Long = 16
'Private Const PROV_EC_ECNRA_FULL As Long = 17
'Private Const PROV_DH_SCHANNEL As Long = 18
'Private Const PROV_SPYRUS_LYNKS As Long = 20
'Private Const PROV_RNG As Long = 21
'Private Const PROV_INTEL_SEC As Long = 22
'Private Const PROV_RSA_AES As Long = 24
'Private Const RSA_KEY_LEN = 256
'Private Const CRYPT_NEWKEYSET = &H8
'Private Const PP_CONTAINER = 6
'Private Const AT_KEYEXCHANGE = 1
'Private Const AT_SIGNATURE = 2
'Private Const PP_ENUMALGS As Long = 1
'
''Private Const SIMPLEBLOB = 1
'
''Algorithm Classes
'Private Const ALG_CLASS_ANY              As Long = 0
'Private Const ALG_CLASS_SIGNATURE        As Long = (1 * 2 ^ 13) '8192
'Private Const ALG_CLASS_MSG_ENCRYPT      As Long = (2 * 2 ^ 13) '16384
'Private Const ALG_CLASS_DATA_ENCRYPT     As Long = (3 * 2 ^ 13) '24576
'Private Const ALG_CLASS_HASH             As Long = (4 * 2 ^ 13) '32768
'Private Const ALG_CLASS_KEY_EXCHANGE     As Long = (5 * 2 ^ 13) '40960
'
'Private Const ALG_TYPE_ANY              As Long = 0
'Private Const ALG_TYPE_DSS              As Long = (1 * 2 ^ 9) '512
'Private Const ALG_TYPE_RSA              As Long = (2 * 2 ^ 9) '1024
'Private Const ALG_TYPE_BLOCK            As Long = (3 * 2 ^ 9) '1536
'Private Const ALG_TYPE_STREAM           As Long = (4 * 2 ^ 9) '2048
'Private Const ALG_TYPE_DH               As Long = (5 * 2 ^ 9) '2560
'Private Const ALG_TYPE_SECURECHANNEL    As Long = (6 * 2 ^ 9) '3072
'
'Private Const ALG_SID_RC2               As Long = 2
'Private Const ALG_SID_RC4               As Long = 1
'Private Const ALG_SID_MD2               As Long = 1
'Private Const ALG_SID_MD4               As Long = 2
'Private Const ALG_SID_MD5               As Long = 3
'Private Const ALG_SID_SHA1              As Long = 4
'Private Const ALG_SID_SCHANNEL_MASTER_HASH As Long = 2
'Private Const ALG_SID_SCHANNEL_MAC_KEY  As Long = 3
'Private Const ALG_SID_TLS1_MASTER       As Long = 6
'Private Const ALG_SID_SCHANNEL_ENC_KEY  As Long = 7
'Private Const ALG_SID_SSL3SHAMD5        As Long = 8
'Private Const ALG_SID_HMAC              As Long = 9
'Private Const ALG_SID_TLS1PRF           As Long = 10
'Private Const ALG_SID_RSA_ANY           As Long = 0
'
'Private Const CALG_SCHANNEL_MASTER_HASH = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MASTER_HASH) '19458
'Private Const CALG_SCHANNEL_MAC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MAC_KEY) '19459
'Private Const CALG_TLS1_MASTER = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_TLS1_MASTER) '19462
'Private Const CALG_SCHANNEL_ENC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_ENC_KEY) '19463
'Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2) '32769
'Private Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4) '32770
'Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5) '32771
'Private Const CALG_SHA1 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1) '32772
'Private Const CALG_SSL3_SHAMD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SSL3SHAMD5) '32776
'Private Const CALG_HMAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HMAC) '32777
'Private Const CALG_TLS1PRF = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_TLS1PRF) '32778
'Private Const CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2) '26114
'Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4) '26625
'Private Const CALG_RSA_KEYX = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY) '41984
'
'Private Const SCHANNEL_MAC_KEY As Long = 0
'Private Const SCHANNEL_ENC_KEY As Long = 1
'Private Const HP_HASHVAL As Long = 2 'Hash value
'Private Const HP_HASHSIZE As Long = 4
'Private Const HP_HMAC_INFO As Long = 5
'Private Const HP_TLS1PRF_LABEL As Long = 6
'Private Const HP_TLS1PRF_SEED As Long = 7
'Private Const CRYPT_SERVER = 1024 '0x00000400
'
''constants from WinErr.h
'Private Const NTE_NO_KEY As Long = &H8009000D              ' Key does not exist
'Private Const NTE_BAD_SIGNATURE As Long = &H80090006       ' Bad Signature
'
''CryptooFilterBox constants
'Private Const CFB_BUSY = 0
'Private Const CFB_READY = 1
'Private Const CFB_VALID = 2
'
'Private Const ENCRYPT_ALGORITHM = CALG_RC4
'Private Const ENCRYPT_BLOCK_SIZE = 1
'
'Private Const CRYPT_EXPORTABLE = 1
'
''KP Constants
'Private Const KP_SCHANNEL_ALG = 20
'Private Const KP_CLIENT_RANDOM = 21 'for setting the Secure Channel client random data
'Private Const KP_SERVER_RANDOM = 22 'for setting the Secure Channel server random data
'
''private property buffers
'Private lCipher As Long
'Public SInBuffer As New Stream
'
'Public SOutBuffer As New Stream
'Private sRndBuffer As String
'
'Private sPublicKey As String
'Private sPassword As String
'Private sSignature As String
'Private lStatus As Long
'Private sKeyFile As String
'Private sMsgType As String
'
''TLS Key Handles
'Private hCryptProv As Long
'Private hMasterKey As Long
'Private hReadKey As Long
'Private hWriteKey As Long
'Private hReadMAC As Long
'Private hWriteMAC As Long
'Private WRITE_MAC_KEY As String
'Private READ_MAC_KEY As String
'Private WRITE_KEY As String
'Private READ_KEY As String
'
''Counters
'Private SEND_SEQ_NUM(2) As Long
'Private RECV_SEQ_NUM(2) As Long
'
'Private flgFinal As Boolean
'
'Public Function ChooseCipher(sCiphers As String) As String
'    Dim N%
'    Alg_Hash = 0
'    N% = 1
'    Do Until Alg_Hash <> 0
'        lCipher = CLng(Asc(Mid$(sCiphers, N%, 1))) * 256 + CLng(Asc(Mid$(sCiphers, N% + 1, 1)))
'        Select Case lCipher
'            Case 4
'                Alg_Hash = CALG_MD5
'            Case 5
'                Alg_Hash = CALG_SHA1
'            Case Else
'                Alg_Hash = 0
'        End Select
'        N% = N% + 2
'    Loop
'    ChooseCipher = Mid$(sCiphers, N% - 2, 1) & Mid$(sCiphers, N% - 1, 1)
'End Function
'Public Property Get CipherToUse2() As Long
'    CipherToUse2 = lCipher
'
'End Property
'Public Property Let CipherToUse(sNewCipher As String)
'    lCipher = Asc(Left$(sNewCipher, 1)) * 256 + Asc(Right$(sNewCipher, 1))
'    Select Case lCipher
'        Case 4
'            Alg_Hash = CALG_MD5
'        Case 5
'            Alg_Hash = CALG_SHA1
'        Case Else
'            Alg_Hash = CALG_SHA1
'    End Select
'End Property
'
'
'Public Property Get ContextHandle() As Long
'    ContextHandle = hCryptProv
'End Property
'
'Public Function ExportPublicKey() As Boolean
'    Const Routine As String = "Crypto.ExportPublicKey"
'    Dim hPublicKey As Long
'    Dim hExchngKey As Long
'    Dim llen As Long
'    lStatus = CFB_BUSY
'    'Acquire Context
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Get handle to Public key
'    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get length of public key
'    'Byval 0& is equivalent to NULL in C++ even though it is declared ByRef as Any
'    If CryptExportKey(hExchngKey, hPublicKey, PUBLICKEYBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Public Key
'    If CryptExportKey(hExchngKey, hPublicKey, PUBLICKEYBLOB, 0, PublicKey, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'ReleaseHandles:
'    If hPublicKey Then CryptDestroyKey hPublicKey
'    If hExchngKey Then CryptDestroyKey hExchngKey
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Function
'
'Public Sub Hash_SHAMD5(fServer As Boolean)
'    Const Routine As String = "Crypto.Hash_SHAMD5"
'    Dim hTLS1PRF As Long
'    Dim data_blob As CRYPT_DATA_BLOB
'    Dim strHash As String
'    Dim bHandShakeHash() As Byte
'    Dim llen As Long
'    Dim bFinish() As Byte
'    Dim sBuffer As String
'    If CryptCreateHash(hCryptProv, CALG_TLS1PRF, hMasterKey, 0, hTLS1PRF) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles
'    End If
'    bHandShakeHash = StrConv(InBuffer, vbFromUnicode)
'    data_blob.pbData = VarPtr(bHandShakeHash(0))
'    data_blob.cbData = SInBuffer.Length '36
'    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_SEED, data_blob, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    If fServer Then
'        bFinish = StrConv(abServerFinished, vbFromUnicode)
'    Else
'        bFinish = StrConv(abClientFinished, vbFromUnicode)
'    End If
'    data_blob.pbData = VarPtr(bFinish(0))
'    data_blob.cbData = UBound(bFinish) + 1 '15
'    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_LABEL, data_blob, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    llen = 12
'    'Prepare buffer
'    sBuffer = String$(llen, Chr$(0))
'    If CryptGetHashParam(hTLS1PRF, HP_HASHVAL, sBuffer, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    InBuffer = Chr$(&H14) & Chr$(0) & Chr$(0) & Chr$(Len(sBuffer) Mod 256) & sBuffer
'    Call TLSEncrypt
'    flgFinal = False
'ReleaseHandles:
'    If hTLS1PRF <> 0 Then CryptDestroyHash hTLS1PRF
'End Sub
'
'Public Function ImportMasterKey() As String
'    Const Routine As String = "Crypto.ImportMasterKey"
'    Dim hPrivateKey As Long
'    Dim hExchngKey As Long
'    Dim hMasterHash As Long
'    Dim llen As Long
'    Dim TempBlob As T_SIMPLEKEYBLOB
'    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
'    Dim data As CRYPT_DATA_BLOB
'    Dim Algorithm As SCHANNEL_ALG
'    Dim sTemp As String
'    Dim bTemp() As Byte
'    lStatus = CFB_BUSY
'    'Acquire Context
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Prepare Master Key Blob
'    TempBlob.bType = SIMPLEBLOB
'    TempBlob.bVersion = CUR_BLOB_VERSION
'    TempBlob.Reserved = 0
'    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
'    TempBlob.algid = CALG_RSA_KEYX
'    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
'    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(1), RSA_KEY_LEN
'    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Import & Decrypt Pre-Master Key
'    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, 0, hMasterKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Set SCHANNEL_ENC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_ENC_KEY
'    Algorithm.algid = CALG_RC4
'    Algorithm.cBits = 128
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Set SCHANNEL_MAC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_MAC_KEY
'    Algorithm.algid = Alg_Hash
'    If Alg_Hash = CALG_MD5 Then
'        Algorithm.cBits = 128
'    Else
'        Algorithm.cBits = 160
'    End If
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add CLIENT_RANDOM
'    data.cbData = UBound(bClientRandom) + 1
'    data.pbData = VarPtr(bClientRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add SERVER_RANDOM
'    data.cbData = UBound(bServerRandom) + 1
'    data.pbData = VarPtr(bServerRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create Master Hash
'    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create the other SCHANNEL keys
'    fClient = False
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0 + CRYPT_EXPORTABLE), hReadKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    sTemp = StrConv(bPreMasterKey, vbUnicode)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    ImportMasterKey = Left$(sTemp, 48)
'    llen = 0
'    'Get WriteMAC Key Length
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get WriteMAC Key (used in HMAC routine)
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        WRITE_MAC_KEY = Left$(sTemp, 20)
'    Else
'        WRITE_MAC_KEY = Left$(sTemp, 16)
'    End If
'    llen = 0
'    'Get ReadMAC Key Length
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get ReadMAC Key (used in HMAC routine)
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        READ_MAC_KEY = Left$(sTemp, 20)
'    Else
'        READ_MAC_KEY = Left$(sTemp, 16)
'    End If
'    llen = 0
'    'Get Write Key Length
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Write Key
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    WRITE_KEY = Left$(sTemp, 16)
'    llen = 0
'    'Get Read Key Length
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Read Key
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    READ_KEY = Left$(sTemp, 16)
' '   Call DebugPrintString("Read Key:", READ_KEY)
' '   Call DebugPrintString("Write Key:", WRITE_KEY)
' '   Call DebugPrintString("Read MAC Key:", READ_MAC_KEY)
' '   Call DebugPrintString("Write MAC Key:", WRITE_MAC_KEY)
'ReleaseHandles:
'    If hPrivateKey Then CryptDestroyKey hPrivateKey
'    If hMasterHash Then CryptDestroyHash hMasterHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'    flgFinal = True
'End Function
'
'Public Function ExportMasterKey() As Boolean
'    Const Routine As String = "Crypto.ExportMasterKey"
'    Dim hPubCertKey As Long 'temporary handle for Certificate Public key
'    Dim hExchngKey As Long
'    Dim hMasterHash As Long
'    Dim llen As Long
'    Dim TempBlob As T_SIMPLEKEYBLOB
'    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
'    Dim Algorithm As SCHANNEL_ALG
'    Dim data As CRYPT_DATA_BLOB
'    Dim sTemp As String
'    Dim bTemp() As Byte
'    Dim sPreMaster As String
'    Dim tTemp As New Stream
'    lStatus = CFB_BUSY
'
'    'Acquire Context
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Get handle to Exchange key
'    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Instead of creating the encrypted Pre-Master Key with the following call,
'    'I chose to create the key, encrypt it with the exchange key, import it,
'    'and then encrypt it using the server public key. Why you ask? The encrypted
'    'Pre-Master key requires the server private key to decrypt.
'    'If CryptGenKey(hCryptProv, CALG_TLS1_MASTER, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
'    '    RaiseEvent Error(err.LastDllError, CGK, Routine)
'    '    GoTo ReleaseHandles
'    'End If
'    llen = 0
'
'
'
'    If CryptEncryptPtr(hExchngKey, 0, True, 0, tTemp.Address, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    tTemp.Clone SInBuffer
'    tTemp.Length = llen
'
'    'sTemp = InBuffer & String$(lLen - sInBuffer.Length, Chr$(0))
'    If CryptEncryptPtr(hExchngKey, 0, True, 0, tTemp.Address, SInBuffer.Length, tTemp.Length) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    bPreMasterKey = tTemp.Partial
'
'    'Prepare Master Key Blob
'    TempBlob.bType = SIMPLEBLOB
'    TempBlob.bVersion = CUR_BLOB_VERSION
'    TempBlob.Reserved = 0
'    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
'    TempBlob.algid = CALG_RSA_KEYX
'    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
'    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(1), RSA_KEY_LEN
'    'Import & Decrypt Pre-Master Key
'    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Import Certificate Public Key as blob
'    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    llen = 0
'
'    tTemp.Reset
'
'    'Get length of encryption
'    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Encrypt Pre-Master Key with Public Cert Key
'  '  tTemp.Clone SInBuffer
'    'tTemp.Length = llen
'
'    sTemp = InBuffer & String$(llen - SInBuffer.Length, Chr$(0))
' '  tTemp.Concat String$(llen - SInBuffer.Length, Chr$(0))
'
'    If CryptEncrypt(hPubCertKey, 0, True, 0, sTemp, SInBuffer.Length, Len(sTemp)) = 0 Then
'   ' If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address,  SInBuffer.Length, tTemp.Length) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    SOutBuffer.Clone tTemp
'
''    OutBuffer = sTemp 'Save Public key encrypted Pre-Master Key
'    tTemp.Reset
'
'
'
'    'Set SCHANNEL_ENC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_ENC_KEY
'    Algorithm.algid = CALG_RC4
'    Algorithm.cBits = 128
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Set SCHANNEL_MAC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_MAC_KEY
'    Algorithm.algid = Alg_Hash
'    If Alg_Hash = CALG_MD5 Then
'        Algorithm.cBits = 128
'    Else
'        Algorithm.cBits = 160
'    End If
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add CLIENT_RANDOM
'    data.cbData = UBound(bClientRandom) + 1
'    data.pbData = VarPtr(bClientRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add SERVER_RANDOM
'    data.cbData = UBound(bServerRandom) + 1
'    data.pbData = VarPtr(bServerRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create Master Hash
'    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create the other SCHANNEL keys
'    fClient = True
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    llen = 0
'    'Get WriteMAC Key Length
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get WriteMAC Key (used in HMAC routine)
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        WRITE_MAC_KEY = Left$(sTemp, 20)
'    Else
'        WRITE_MAC_KEY = Left$(sTemp, 16)
'    End If
'    llen = 0
'    'Get ReadMAC Key Length
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get ReadMAC Key (used in HMAC routine)
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        READ_MAC_KEY = Left$(sTemp, 20)
'    Else
'        READ_MAC_KEY = Left$(sTemp, 16)
'    End If
'    llen = 0
'    'Get Write Key Length
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Write Key
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    WRITE_KEY = Left$(sTemp, 16)
'    llen = 0
'    'Get Read Key Length
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Read Key
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    READ_KEY = Left$(sTemp, 16)
'  '  Call DebugPrintString("Read Key:", READ_KEY)
'  '  Call DebugPrintString("Write Key:", WRITE_KEY)
'  '  Call DebugPrintString("Read MAC Key:", READ_MAC_KEY)
'  '  Call DebugPrintString("Write MAC Key:", WRITE_MAC_KEY)
'    ExportMasterKey = True 'Success
'ReleaseHandles:
'    Set tTemp = Nothing
'
'    If hPubCertKey Then CryptDestroyKey hPubCertKey
'    If hMasterHash Then CryptDestroyHash hMasterHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    flgFinal = True
'    lStatus = CFB_READY
'End Function
'Public Sub IncSendSeqNum(Optional flgClear As Boolean)
'    If flgClear Then
'        SEND_SEQ_NUM(0) = 0
'        SEND_SEQ_NUM(1) = 0
'        SEND_SEQ_NUM(2) = 0
'        Exit Sub
'    End If
'    If SEND_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
'        SEND_SEQ_NUM(0) = SEND_SEQ_NUM(0) + 1
'    ElseIf SEND_SEQ_NUM(1) < 16777215 Then
'        SEND_SEQ_NUM(1) = SEND_SEQ_NUM(1) + 1
'        SEND_SEQ_NUM(0) = 0
'    ElseIf SEND_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
'        SEND_SEQ_NUM(2) = SEND_SEQ_NUM(2) + 1
'        SEND_SEQ_NUM(1) = 0
'    Else
'        SEND_SEQ_NUM(0) = 0
'        SEND_SEQ_NUM(1) = 0
'        SEND_SEQ_NUM(2) = 0
'    End If
'End Sub
'Public Sub IncRecvSeqNum(Optional flgClear As Boolean)
'    If flgClear Then
'        RECV_SEQ_NUM(0) = 0
'        RECV_SEQ_NUM(1) = 0
'        RECV_SEQ_NUM(2) = 0
'        Exit Sub
'    End If
'    If RECV_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
'        RECV_SEQ_NUM(0) = RECV_SEQ_NUM(0) + 1
'    ElseIf RECV_SEQ_NUM(1) < 16777215 Then
'        RECV_SEQ_NUM(1) = RECV_SEQ_NUM(1) + 1
'        RECV_SEQ_NUM(0) = 0
'    ElseIf RECV_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
'        RECV_SEQ_NUM(2) = RECV_SEQ_NUM(2) + 1
'        RECV_SEQ_NUM(1) = 0
'    Else
'        RECV_SEQ_NUM(0) = 0
'        RECV_SEQ_NUM(1) = 0
'        RECV_SEQ_NUM(2) = 0
'    End If
'End Sub
'
'Public Sub IncSendSeqSet(ByVal index0 As Byte, ByVal index1 As Byte, ByVal index2 As Byte)
'
'    SEND_SEQ_NUM(0) = index0
'    SEND_SEQ_NUM(1) = index1
'    SEND_SEQ_NUM(2) = index2
'
'End Sub
'Public Sub IncRecvSeqSet(ByVal index0 As Byte, ByVal index1 As Byte, ByVal index2 As Byte)
'
'    RECV_SEQ_NUM(0) = index0
'    RECV_SEQ_NUM(1) = index1
'    RECV_SEQ_NUM(2) = index2
'
'End Sub
'
'Public Property Get msgType() As String
'    msgType = sMsgType
'End Property
'
'Public Property Let msgType(vNewValue As String)
'    sMsgType = vNewValue
'End Property
'
'Public Function VerifySig() As Boolean
'    Const Routine As String = "VerifySig"
'    Dim hHash As Long
'    Dim hPubCertKey As Long
'    Dim sDescription As String
'    'Create a hash object.
'    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles:
'    End If
'    'Hash the data
'    If CryptHashDataPtr(hHash, ByVal SInBuffer.Address, SInBuffer.Length, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        GoTo ReleaseHandles:
'    End If
'    'Import Certificate Public Key as blob
'    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    sDescription = vbNullString
'    If CryptVerifySignature(hHash, sSignature, Len(sSignature), hPubCertKey, sDescription, 0) = 0 Then
'        VerifySig = False
'    Else
'        VerifySig = True
'    End If
'ReleaseHandles:
'    If hPubCertKey Then CryptDestroyKey hPubCertKey
'    If hHash Then CryptDestroyHash hHash
'End Function
'
'
'Public Property Get InBuffer() As String
'    InBuffer = Convert(SInBuffer.Partial)
'End Property
'Public Property Let InBuffer(vNewValue As String)
'    SInBuffer.Length = Len(vNewValue)
'    SInBuffer.Placeat Convert(vNewValue)
'End Property
'
'
'Public Property Let KeyFile(sNewFileName As String)
'    sKeyFile = sNewFileName
'End Property
'
'Public Property Get KeyFile() As String
'    KeyFile = sKeyFile
'End Property
'
'Public Property Get KeyName() As String
'    KeyName = App.CompanyName
'End Property
'Public Function MD5_AddToHash(ByVal strToHash As String, hHash) As Long
'    Const Routine As String = "Crypto.MD5_AddToHash"
'    If hHash = 0 Then
'        If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
'            RaiseEvent Error(Err.LastDllError, cch, Routine)
'            Exit Function
'        End If
'    End If
'    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        Exit Function
'    End If
'    MD5_AddToHash = hHash
'End Function
'Public Function CopyHash(hHashToCopy) As Long
'    Const Routine As String = "Crypto.CopyHash"
'    Dim phHash As Long
'    If CryptDuplicateHash(hHashToCopy, ByVal 0&, 0&, phHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDuH, Routine)
'        CopyHash = 0
'    Else
'        CopyHash = phHash
'    End If
'End Function
'
'Public Function MD5_GetHash(hHash As Long) As String
'    Const Routine As String = "Crypto.MD5_GetHash"
'    Dim strHash As String
'    Dim llen As Long
'    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    strHash = String(llen, vbNullChar)
'    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    If hHash <> 0 Then
'        CryptDestroyHash hHash
'        hHash = 0
'    End If
'    MD5_GetHash = strHash
'End Function
'Public Function MD5_Hash(ByVal TheString As String) As String
'    Const Routine As String = "Crypto.MD5_Hash"
'    Dim strHash As String
'    Dim hHash As Long
'    Dim llen As Long
'    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        Exit Function
'    End If
'    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        Exit Function
'    End If
'    'Get length of hash
'    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    strHash = String(llen, vbNull)
'    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    If hHash <> 0 Then CryptDestroyHash hHash
'    MD5_Hash = strHash
'End Function
'Public Function SHA1_AddToHash(ByVal strToHash As String, hHash) As Long
'    Const Routine As String = "Crypto.SHA1_AddToHash"
'    If hHash = 0 Then
'        If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
'            RaiseEvent Error(Err.LastDllError, cch, Routine)
'            Exit Function
'        End If
'    End If
'    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        Exit Function
'    End If
'    SHA1_AddToHash = hHash
'End Function
'Public Function SHA1_CopyHash() As Long
'
'End Function
'
'Public Function SHA1_GetHash(hHash As Long) As String
'    Const Routine As String = "Crypto.SHA1_GetHash"
'    Dim strHash As String
'    Dim llen As Long
'    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    strHash = String(llen, vbNullChar)
'    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    If hHash <> 0 Then
'        CryptDestroyHash hHash
'        hHash = 0
'    End If
'    SHA1_GetHash = strHash
'End Function
'Public Function SHA1_Hash(ByVal TheString As String) As String
'    Const Routine As String = "Crypto.SHA1_Hash"
'    Dim strHash As String
'    Dim hHash As Long
'    Dim llen As Long
'    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        Exit Function
'    End If
'    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        Exit Function
'    End If
'    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    strHash = String(llen, vbNull)
'    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        Exit Function
'    End If
'    If hHash <> 0 Then CryptDestroyHash hHash
'    SHA1_Hash = strHash
'End Function
'
'Public Property Get OutBuffer() As String
'    OutBuffer = Convert(SOutBuffer.Partial)
'End Property
'Public Property Let OutBuffer(ByRef vNewValue As String)
'    SOutBuffer.Length = Len(vNewValue)
'    SOutBuffer.Placeat Convert(vNewValue)
'End Property
'Public Property Let PubKey(NewKey As String)
'    sPublicKey = NewKey
'End Property
'
'Public Property Get PubKey() As String
'    PubKey = StrConv(PublicKey.bPublicKey(), vbUnicode)
'End Property
'
'Private Function SendSeqNum() As String
'    Dim bTmp(1 To 8) As Byte
'    CopyMemory bTmp(1), SEND_SEQ_NUM(0), 3
'    CopyMemory bTmp(4), SEND_SEQ_NUM(1), 3
'    CopyMemory bTmp(7), SEND_SEQ_NUM(2), 2
'    SendSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
'End Function
'Private Function RecvSeqNum() As String
'    Dim bTmp(1 To 8) As Byte
'    CopyMemory bTmp(1), RECV_SEQ_NUM(0), 3
'    CopyMemory bTmp(4), RECV_SEQ_NUM(1), 3
'    CopyMemory bTmp(7), RECV_SEQ_NUM(2), 2
'    RecvSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
'End Function
'
'Private Function ReadKeyFile(sFileName As String, bData() As Byte) As Boolean
'    Dim lFileHandle As Long
'    ' Reads a byte array from a file.
'    ' If Data is fixed, it must be of the expected size.
'    On Error GoTo ErrorHandler
'    lFileHandle = FreeFile
'    Open sFileName For Binary As lFileHandle
''    ReDim bData(1 To LOF(lFileHandle))
'    Get lFileHandle, , bData
'    Close lFileHandle
'    ReadKeyFile = True
'    Exit Function
'ErrorHandler:
'    ReadKeyFile = False
'End Function
'Public Property Get RndBuffer() As String
'    RndBuffer = sRndBuffer
'End Property
'Public Function GetPubKey()
'    'Check if PublicKey loaded
'    If PublicKey.bPublicKey(1) = 0 Then
'        If Not GetPubKeyBlob Then Exit Function
'    End If
'    PriFlg = False
'End Function
'Public Function EncryptSessionKey()
'    Const Routine As String = "Crypto.EncryptSessionKey"
'    Dim hPublicKey As Long
'    Dim sTemp As New Stream
'    Dim llen As Long
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    If PriFlg Then
'        'Import Private Key
'        If CryptImportKey(hCryptProv, PrivateKey, Len(PrivateKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CIK, Routine)
'            GoTo ReleaseHandles
'        End If
'    Else
'        'Import Public Key
'        If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CIK, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Get length of key
'    If CryptEncryptPtr(hPublicKey, 0, True, 0, ByVal sTemp.Address, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Prepare buffer
'
'    sTemp.Length = llen
'    sTemp.Placeat modMemory.Convert(sPassword), 0, Len(sPassword)
'   ' sTemp = sPassword & String$(llen - Len(sPassword), Chr$(0))
'    'Encrypt Session Key
'    If CryptEncryptPtr(hPublicKey, 0, True, 0, ByVal sTemp.Address, Len(sPassword), llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    SOutBuffer.Clone sTemp
'ReleaseHandles:
'    Set sTemp = Nothing
'    If hPublicKey Then CryptDestroyKey hPublicKey
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'End Function
'Public Function DecryptSessionKey()
'    Const Routine As String = "Crypto.DecryptSessionKey"
'    Dim hPublicKey As Long
'    Dim sTemp As String
'    Dim llen As Long
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    If PriFlg Then
'        'Import Private Key
'        If CryptImportKey(hCryptProv, PrivateKey, Len(PrivateKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CIK, Routine)
'            GoTo ReleaseHandles
'        End If
'    Else
'        'Import Public Key
'        If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CIK, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    sTemp = InBuffer
'    If CryptDecrypt(hPublicKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    OutBuffer = Left$(sTemp, 32)
'ReleaseHandles:
'    If hPublicKey Then CryptDestroyKey hPublicKey
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'End Function
'
'Private Function GetPubKeyBlob() As Boolean
'    If (ReadKeyFile(App.path & "\" & sKeyFile, PublicKey.bPublicKey)) Then
'        GetPubKeyBlob = True
'    Else
'        GetPubKeyBlob = False
'    End If
'End Function
'
'Public Function GetPriKey(Optional sRet As Boolean) As String
'    'Check if PrivateKey loaded
'    If PrivateKey.bPrivateKey(1) = 0 Then
'        If Not GetPriKeyBlob Then Exit Function
'    End If
'    PriFlg = True
'End Function
'Private Function GetPriKeyBlob() As Boolean
'    If (ReadKeyFile(App.path & "\" & sKeyFile, PrivateKey.bPrivateKey)) Then
'        GetPriKeyBlob = True
'    Else
'        GetPriKeyBlob = False
'    End If
'End Function
'
'
'Public Property Let ClientRandom(sClientRandom As String)
'    ReDim bClientRandom(0)
'    bClientRandom = StrConv(sClientRandom, vbFromUnicode)
'End Property
'Public Property Let ServerRandom(sServerRandom As String)
'    ReDim bServerRandom(0)
'    bServerRandom = StrConv(sServerRandom, vbFromUnicode)
'End Property
'
'Public Property Let PreMasterKey(sPreMasterKey As String)
'    ReDim bPreMasterKey(0)
'    bPreMasterKey = StrConv(sPreMasterKey, vbFromUnicode)
'End Property
'
'Public Function SetPubKey() As Boolean
'    Dim bNewKey() As Byte
'    bNewKey = StrConv(sPublicKey, vbFromUnicode)
'    NewPubKey.bType = PUBLICKEYBLOB
'    NewPubKey.bVersion = CUR_BLOB_VERSION
'    NewPubKey.Reserved = 0
'    NewPubKey.aiKeyAlg = CALG_RSA_KEYX
'    NewPubKey.magic = &H31415352 '"RSA1"
'    NewPubKey.bitlen = Len(sPublicKey) * 8
'    NewPubKey.pubexp = 65537 '0x00010001
'    CopyMemory NewPubKey.modulus(1), bNewKey(0), Len(sPublicKey)
'    CopyMemory PublicKey.bPublicKey(1), NewPubKey, Len(sPublicKey) + 20
'    SetPubKey = True
'End Function
'
'Public Property Get Signature() As String
'    Signature = sSignature
'End Property
'Public Property Let Signature(vNewValue As String)
'    sSignature = vNewValue
'End Property
'Public Sub Sign()
'    Const Routine As String = "Crypto.Sign"
'    Dim sDescription As String
'    Dim hHash As Long
'    Dim lSignatureLen As Long
'    'switch Status property
'    lStatus = CFB_BUSY
'    'init Signature property
'    sSignature = ""
'    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'        GoTo ReleaseHandles:
'    End If
'    'Create a hash object.
'    If CryptCreateHash(hCryptProv, Alg_Hash, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles:
'    End If
'    If CryptHashData(hHash, InBuffer, SInBuffer.Length, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        GoTo ReleaseHandles:
'    End If
'    'Sign hash object.
'    'Determine size of signature.
'    sDescription = vbNullChar
'    If CryptSignHash(hHash, AT_KEYEXCHANGE, sDescription, 0, vbNullString, lSignatureLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSH, Routine)
'        GoTo ReleaseHandles:
'    End If
'    sSignature = String(lSignatureLen, vbNullChar)
'    'Sign hash object (with signature key).
'    If CryptSignHash(hHash, AT_KEYEXCHANGE, sDescription, 0, sSignature, lSignatureLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSH, Routine)
'        GoTo ReleaseHandles:
'    End If
'ReleaseHandles:
'    If hHash Then CryptDestroyHash hHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Sub
'
'Public Sub TLSDecrypt()
'    Const Routine As String = "clsCrypt.TLSDecrypt"
'    Dim sBuffer As String
'    lStatus = CFB_BUSY
'    'Get handle to the provider
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'
'    sBuffer = InBuffer
'    'Decrypt data
'    If CryptDecrypt(hReadKey, 0, False, 0, sBuffer, Len(sBuffer)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Not VerifyHMAC(sBuffer) Then
'        RaiseEvent Error(&H80090002, "VerifyHMAC", Routine)
'    End If
'   ' SOutBuffer.Clone SInBuffer
'
'    OutBuffer = sBuffer
'
'
'ReleaseHandles:
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Sub
'
'Public Sub TLSEncrypt()
'    Const Routine As String = "clsCrypt.TLSEncrypt"
'    Dim sBuffer As String
'    Dim buflen As Long
'    'switch Status property
'    lStatus = CFB_BUSY
'    'Get handle to the provider
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'   ' Debug.Print "Final = " & CBool(flgFinal)
'    'Add HMAC to message
'    'Choose either HMAC routine, as both should deliver the same result
'    sBuffer = InBuffer & HMAC(Alg_Hash, WRITE_MAC_KEY, InBuffer, True)
''    sBuffer = InBuffer & HMAC2(hWriteMAC, InBuffer, True)
'    buflen = Len(sBuffer)
'    'Null string returns required buffer length
'    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, vbNullString, Len(sBuffer), buflen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    If buflen > Len(sBuffer) Then 'Adjust buffer length if necessary
'        sBuffer = sBuffer & String$(buflen - Len(sBuffer), Chr$(0))
'    End If
'    'Encrypt data
'    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, sBuffer, Len(sBuffer), buflen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    OutBuffer = sBuffer
'ReleaseHandles:
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Sub
'
'Private Function VerifyHMAC(ByRef DecryptedRecord As String) As Boolean
'    Dim N%
'    'Verify the Message Authentication Code
'    Dim AppendedHMAC As String
'    Dim CalculatedHMAC As String
'    Dim llen As Long
'    If Alg_Hash = CALG_MD5 Then
'        llen = Len(DecryptedRecord) - 16
'    Else
'        llen = Len(DecryptedRecord) - 20
'    End If
'    AppendedHMAC = Mid(DecryptedRecord, llen + 1)
'    DecryptedRecord = Mid(DecryptedRecord, 1, llen)
'    'Either HMAC routine should produce the correct result
'    CalculatedHMAC = HMAC(Alg_Hash, READ_MAC_KEY, DecryptedRecord, False)
''    CalculatedHMAC = HMAC2(hReadMAC, DecryptedRecord, False)
'    If CalculatedHMAC = AppendedHMAC Then
'        VerifyHMAC = True
'    Else
'        VerifyHMAC = False
'     '   Debug.Print "HMAC Recv = ";
'     '   For N% = 1 To Len(AppendedHMAC)
'     '       Debug.Print Right$("0" & Hex$(Asc(Mid$(AppendedHMAC, N%, 1))), 2) & " ";
'     '   Next N%
'     '   Debug.Print
'     '   Debug.Print "HMAC Calc = ";
'     '   For N% = 1 To Len(CalculatedHMAC)
'     '       Debug.Print Right$("0" & Hex$(Asc(Mid$(CalculatedHMAC, N%, 1))), 2) & " ";
'     '   Next N%
'     '   Debug.Print
'    End If
'End Function
'
'Private Function WriteFile(sFileName As String, bData() As Byte) As Boolean
'    Dim lFileHandle As Long
'    ' Writes a byte array out to a file.
'    On Error GoTo ErrorHandler
'    lFileHandle = FreeFile
'    Open sFileName For Binary As lFileHandle
'    Put lFileHandle, , bData
'    Close lFileHandle
'    WriteFile = True
'    Exit Function
'ErrorHandler:
'    WriteFile = False
'End Function
'
'Public Sub Encrypt()
'    Const Routine As String = "Crypto.Encrypt"
'    Dim hExchgKey As Long
'    Dim hHash As Long
'    Dim hKey As Long
'    Dim sCryptBuffer As String
'    Dim lCryptLength As Long
'    Dim lCryptBufLen As Long
'    'switch Status property
'    lStatus = CFB_BUSY
'    'Get handle to the provider
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Create a hash object.
'    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Hash in the password data.
'    If CryptHashData(hHash, sPassword, Len(sPassword), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Derive a session key from the hash object.
'    If CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Destroy the hash object.
'    CryptDestroyHash hHash
'    hHash = 0
'    'Prepare a string buffer for the CryptEncrypt function
'    lCryptLength = SInBuffer.Length
'    lCryptBufLen = lCryptLength * 2
'    sCryptBuffer = String(lCryptBufLen, vbNullChar)
'    LSet sCryptBuffer = InBuffer
'    'Encrypt data
'    If CryptEncrypt(hKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        'GoTo ReleaseHandles
'    End If
'    OutBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
'ReleaseHandles:
'    If hKey Then CryptDestroyKey hKey
'    If hExchgKey Then CryptDestroyKey hExchgKey
'    If hHash Then CryptDestroyHash hHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Sub
'Public Function HMAC2(h_HMAC As Long, ByVal sMsg As String, flgSend As Boolean) As String
'    'HMAC_MD5(Key, m) = MD5((Key XOR opad) ++ MD5((Key XOR ipad) ++ m)
'    '(++ means concatenate, "opad" is the bytes "5c 5c ... 5c",
'    'and "ipad" is the bytes "36 36 ... 36").
'    Const Routine As String = "Crypto.HMAC2"
'    Dim strHash As String
'    Dim hMACHash As Long
'    Dim llen As Long
'    Dim MAC_Header As String
'    Dim HMACInfo As HMAC_INFO
'    'MAC_Header = Sequence (8 bytes) + Type (1 byte) + Version (2 bytes) + msg Length (variable)
'    If flgSend Then
'        MAC_Header = SendSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
'     '   Debug.Print "MAC Header on Send:"
'    Else
'        MAC_Header = RecvSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
'       ' Debug.Print "MAC Header on Receive:"
'    End If
'    Dim N%
''    For N% = 1 To Len(MAC_Header)
''        Debug.Print Right$("0" & Hex$(Asc(Mid$(MAC_Header, N%, 1))), 2) & " ";
''    Next N%
''    Debug.Print
'    'if inner & outer strings and lengths set to 0, defaults are used
'    'They should already be zeroed, but to make sure
'    ZeroMemory HMACInfo.ALG_ID, Len(HMACInfo)
'    HMACInfo.ALG_ID = Alg_Hash
'    If CryptCreateHash(hCryptProv, CALG_HMAC, h_HMAC, 0, hMACHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptSetHashParam(hMACHash, HP_HMAC_INFO, HMACInfo, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptHashData(hMACHash, MAC_Header & sMsg, Len(MAC_Header & sMsg), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptGetHashParam(hMACHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    strHash = String$(llen, Chr$(0))
'    If CryptGetHashParam(hMACHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
'        GoTo ReleaseHandles
'    End If
'    HMAC2 = strHash
'    'Call DebugPrintString("HMAC2", HMAC2)
'ReleaseHandles:
'    If hMACHash <> 0 Then CryptDestroyHash hMACHash
'End Function
'Public Function HMAC(HashAlg As Long, sKey As String, sMsg As String, flgSend As Boolean) As String
'    'Note: if key is longer than the blocksize (64 for MD5/SHA1), the key should
'    'be hashed first to reduce it's length
'    Const Routine As String = "Crypto.HMAC"
'    Dim MAC_Header As String
'    Dim o_key_pad As String
'    Dim i_key_pad As String
'    Dim N%
'    'MAC_Header = Sequence (8 bytes) + Type (1 byte) + Version (2 bytes) + msg Length (variable)
'    If flgSend Then
'        MAC_Header = SendSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
'    '    Debug.Print "MAC Header on Send:"
'    Else
'        MAC_Header = RecvSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
'    '    Debug.Print "MAC Header on Receive:"
'    End If
' '   For N% = 1 To Len(MAC_Header)
' '       Debug.Print Right$("0" & Hex$(Asc(Mid$(MAC_Header, N%, 1))), 2) & " ";
' '   Next N%
' '   Debug.Print
'    If Len(sKey) > 64 Then
'        If HashAlg = CALG_MD5 Then
'            sKey = MD5_Hash(sKey)
'        ElseIf HashAlg = CALG_SHA1 Then
'            sKey = SHA1_Hash(sKey)
'        End If
'    End If
'    o_key_pad = String$(64, &H5C)
'    i_key_pad = String$(64, &H36)
'    For N% = 1 To Len(sKey)
'        Mid$(o_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H5C)
'        Mid$(i_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H36)
'    Next N%
'    If HashAlg = CALG_MD5 Then
'        HMAC = MD5_Hash(o_key_pad & MD5_Hash(i_key_pad & MAC_Header & sMsg))
'    ElseIf HashAlg = CALG_SHA1 Then
'        HMAC = SHA1_Hash(o_key_pad & SHA1_Hash(i_key_pad & MAC_Header & sMsg))
'    End If
'    'Call DebugPrintString("HMAC", HMAC)
'End Function
'
'Public Sub GenRandom(llen As Long)
'    Const Routine As String = "Crypto.GenRandom"
'    Dim sRandom As String
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    sRandom = String$(llen, Chr$(0))
'    If CryptGenRandom(hCryptProv, llen, sRandom) = 0 Then
'        sRndBuffer = ""
'    Else
'        sRndBuffer = sRandom
'    End If
'ReleaseHandles:
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'End Sub
'
'Private Sub DebugPrintString(sDescr As String, strToPrint As String)
'    Dim N%
'    Debug.Print sDescr & ":"
'    For N% = 1 To Len(strToPrint)
'        Debug.Print Right$("0" & Hex$(Asc(Mid$(strToPrint, N%, 1))), 2) & " ";
'        If N% Mod 16 = 0 Then Debug.Print
'    Next N%
'    Debug.Print
'End Sub
'Public Sub Decrypt()
'    Const Routine As String = "Crypto.Decrypt"
'    Dim hExchgKey As Long
'    Dim hHash As Long
'    Dim hKey As Long
'    Dim sCryptBuffer As String
'    Dim lCryptBufLen As Long
'    Dim lCryptPoint As Long
'    Dim lPasswordPoint As Long
'    Dim lPasswordCount As Long
'    'switch Status property
'    lStatus = CFB_BUSY
'    'Init OutBuffer
'    OutBuffer = ""
'    'Get handle to the default provider.
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Create a hash object.
'    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, cch, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Hash in the password data.
'    If CryptHashData(hHash, sPassword, Len(sPassword), 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CHD, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Derive a session key from the hash object.
'    If CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Destroy the hash object.
'    CryptDestroyHash hHash
'    hHash = 0
'    'Prepare sCryptBuffer for CryptDecrypt
'    lCryptBufLen = SInBuffer.Length * 2
'    sCryptBuffer = String(lCryptBufLen, vbNullChar)
'    LSet sCryptBuffer = InBuffer
'    'Decrypt data
'    If CryptDecrypt(hKey, 0, 1, 0, sCryptBuffer, lCryptBufLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    OutBuffer = Mid$(sCryptBuffer, 1, SInBuffer.Length)
'ReleaseHandles:
'    If (hKey) Then CryptDestroyKey hKey
'    If hExchgKey Then CryptDestroyKey hExchgKey
'    If hHash Then CryptDestroyHash hHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    lStatus = CFB_READY
'End Sub
'Public Sub CreateKeyPair()
'    Const Routine As String = "Crypto.CreateKeyPair"
'    Dim hPublicKey As Long
'    Dim llen As Long
'    Dim PublicKey  As T_EXP_PUBLICKEYBLOB
'    Dim PrivateKey As T_EXP_PRIVATEKEYBLOB
'    'Creates a new key pair and writes out a file for each part
'    'called "FileName.fpu" and "FileName.fpr" respectively in app.path
'    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'        'MsgBox ("Error " & Hex$(err.LastDllError) & " during CryptAcquireContext!")
'        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'        GoTo ReleaseHandles
'    End If
'    ' Create new Exchange key pair.
'    ' The use of &H80000000 guarantees a 2048 bit key.
'    If CryptGenKey(hCryptProv, AT_KEYEXCHANGE, &H8000000 Or CRYPT_EXPORTABLE, hPublicKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ' Copy the keys to variables.
'    llen = LenB(PublicKey)
'    If CryptExportKey(hPublicKey, 0, PUBLICKEYBLOB, 0, PublicKey, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    llen = LenB(PrivateKey)
'    If CryptExportKey(hPublicKey, 0, PRIVATEKEYBLOB, 0, PrivateKey, llen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ' See if either key already exists.
'    If Dir(App.path & "\" & sKeyFile & ".fpu") <> "" And _
'        Dir(App.path & "\" & sKeyFile & ".fpr") <> "" Then
'           ' If MsgBox("Key files already exist!" & vbCrLf & "Overwrite them?", _
'           '     vbYesNo) = vbNo Then Exit Sub
'    End If
'    ' Create the key files.
'    If Not (WriteFile(App.path & "\" & sKeyFile & ".fpu", PublicKey.bPublicKey)) Then
'       ' MsgBox "Write file '" & sKeyFile & ".fpu' failed!", vbExclamation
'    End If
'    If Not (WriteFile(App.path & "\" & sKeyFile & ".fpr", PrivateKey.bPrivateKey)) Then
'       ' MsgBox "Write file '" & sKeyFile & ".fpr' failed!", vbExclamation
'    End If
'ReleaseHandles:
'    ' Destroy the key.
'    CryptDestroyKey hPublicKey: hPublicKey = 0
'    ' Release Provider
'    CryptReleaseContext hCryptProv, 0: hCryptProv = 0
'End Sub
'
'Public Property Get Status() As Long
'    Status = lStatus
'End Property
'Private Function InitUser() As Long
'    Const Routine As String = "Crypto.InitUser"
'    Dim hCryptKey As Long
'    Dim avProviderData(1000) As Byte
'    Dim lProviderDataLen As Long
'    Dim lDataSize As Long
'    Dim sUserName As String
'    Dim lPoint As Long
'    Dim sBuffer As String
'    Dim strMsg As String
'    'Attempt to acquire a handle to the key container.
'    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'        'Create key container.
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, CRYPT_NEWKEYSET) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            Exit Function
'        End If
'        'Get name of key container.
'        lProviderDataLen = Len(avProviderData(0)) * (UBound(avProviderData) + 1)
'        If CryptGetProvParam(hCryptProv, PP_CONTAINER, avProviderData(0), lProviderDataLen, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CGPP, Routine)
'            avProviderData(0) = 0
'        End If
'        'Get sUserName from avProviderData()
'        lPoint = LBound(avProviderData)
'        While lPoint <= UBound(avProviderData)
'            If avProviderData(lPoint) <> 0 Then
'                sUserName = sUserName & Chr$(avProviderData(lPoint))
'            Else
'                lPoint = UBound(avProviderData)
'            End If
'            lPoint = lPoint + 1
'        Wend
'        'MsgBox ("Create key container " & sUserName)
'    End If
'    'Attempt to get handle to signature key
'    'SChannel does not support Signature Key
'    If modMemory.Convert(sProvider.Partial) <> MS_DEF_RSA_SCHANNEL_PROV Then
'        If CryptGetUserKey(hCryptProv, AT_SIGNATURE, hCryptKey) = 0 Then
'            If Err.LastDllError = NTE_NO_KEY Then
'               ' MsgBox ("System does not have Signature Key pair!" & vbCrLf & "One will be created.")
'                If CryptGenKey(hCryptProv, AT_SIGNATURE, 0, hCryptKey) = 0 Then
'                    RaiseEvent Error(Err.LastDllError, CGK, Routine)
'                    Exit Function
'                Else
'                    CryptDestroyKey hCryptProv
'                End If
'            Else
'                RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'                Exit Function
'            End If
'        End If
'    End If
'    'Attempt to get handle to exchange key
'    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hCryptKey) = 0 Then
'        If Err.LastDllError = NTE_NO_KEY Then
'           ' MsgBox ("System does not have Key Exchange pair!" & vbCrLf & "One will be created.")
'            ' The use of &H80000000 guarantees a 2048 bit key.
'            If CryptGenKey(hCryptProv, AT_KEYEXCHANGE, &H8000000, hCryptKey) = 0 Then
'                RaiseEvent Error(Err.LastDllError, CGK, Routine)
'                Exit Function
'            Else
'                CryptDestroyKey hCryptProv
'            End If
'        Else
'            RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'            Exit Function
'        End If
'    End If
'    'Enumerate the supported algorithms.
'    Dim flag As Long
'    Dim algid As Long
'    Dim dwBits  As Long
'    Dim namelen As Long
'    Dim algType As String
'    Dim lPtr As Long
'    Dim name As String
'    flag = 1
'    If CryptGetProvParam(hCryptProv, PP_ENUMALGS, ByVal vbNullString, lPtr, flag) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CGPP, Routine)
'            Exit Function
'        End If
'    Dim b() As Byte
'    Dim bName() As Byte
'    ReDim b(lPtr - 1)
'    strMsg = strMsg & vbCrLf & "ALGID" & vbTab & "dwBits" & vbTab & "algType" & String(2, vbTab) & "len" & vbTab & "Name" & vbCrLf _
'        & String(5, "-") & vbTab & String(6, "-") & vbTab & String(12, "-") & vbTab & String(3, "-") & vbTab & String(3, "-") & vbCrLf
'    While CryptGetProvParam(hCryptProv, PP_ENUMALGS, b(0), lPtr, flag) <> 0
'        flag = 2
'        CopyMemory algid, b(0), LenB(algid)
'        CopyMemory dwBits, b(4), LenB(dwBits)
'        CopyMemory namelen, b(8), LenB(namelen)
'        ReDim bName(namelen - 2)
'        CopyMemory bName(0), b(12), namelen - 1
'        name = StrConv(bName, vbUnicode)
'        Select Case (algid And &HF000)
'            Case ALG_CLASS_ANY
'                algType = "Any         "
'            Case ALG_CLASS_SIGNATURE
'                algType = "Signature   "
'            Case ALG_CLASS_MSG_ENCRYPT
'                algType = "Msg_Encrypt "
'            Case ALG_CLASS_DATA_ENCRYPT
'                algType = "Data_Encrypt"
'            Case ALG_CLASS_HASH
'                algType = "Hash        "
'            Case ALG_CLASS_KEY_EXCHANGE
'                algType = "Exchange    "
'            Case Else
'                algType = "Unknwon     "
'        End Select
'        strMsg = strMsg & algid & vbTab & Right("     " & dwBits, 6) & vbTab & algType & vbTab & namelen & vbTab & name & vbCrLf
'    Wend
'   'Debug.Print strMsg
'    InitUser = True
'End Function
'Private Sub Class_Initialize()
'    sContainer.Concat modMemory.Convert("chrome@winternet.com" & vbNullChar)
'    sProvider.Concat modMemory.Convert(MS_DEF_RSA_SCHANNEL_PROV)
'    sProvType = PROV_RSA_SCHANNEL
'
'    If Not InitUser = True Then
'      '  MsgBox ("InitUser failed")
'    End If
'End Sub
'Public Property Get Password() As String
'    Password = sPassword
'End Property
'
'Public Property Let Password(vNewValue As String)
'    sPassword = vNewValue
'End Property
'
'Private Sub Class_Terminate()
'    CryptReleaseContext hCryptProv, 0
'    Set sContainer = Nothing
'    Set sProvider = Nothing
'    Set SInBuffer = Nothing
'    Set SOutBuffer = Nothing
'End Sub
'
'Public Function ErrorCode(ByVal Code As Byte) As String
'
'    Select Case Code
'        Case 0
'
'
'            ErrorCode = "close_notify - Notifies the recipient that the sender will not send any more messages on this connection."
'
'        Case 10
'
'
'
'            ErrorCode = "unexpected_message - Received an inappropriate message This alert should never be observed in communication between proper implementations. This message is always fatal."
'
'        Case 20
'
'
'
'            ErrorCode = "bad_record_mac - Received a record with an incorrect MAC. This message is always fatal."
'
'        Case 21
'
'
'
'            ErrorCode = "decryption_failed - Decryption of a TLSCiphertext record is decrypted in an invalid way: either it was not an even multiple of the block length or its padding values, when checked, were not correct. This message is always fatal."
'
'        Case 22
'
'
'
'            ErrorCode = "record_overflow - Received a TLSCiphertext record which had a length more than 2^14+2048 bytes, or a record decrypted to a TLSCompressed record with more than 2^14+1024 bytes. This message is always fatal."
'
'        Case 30
'
'
'
'            ErrorCode = "decompression_failure - Received improper input, such as data that would expand to excessive length, from the decompression function. This message is always fatal."
'
'        Case 40
'
'
'
'            ErrorCode = "handshake_failure - Indicates that the sender was unable to negotiate an acceptable set of security parameters given the options available. This is a fatal error."
'
'        Case 42
'
'
'
'            ErrorCode = "bad_certificate - There is a problem with the certificate, for example, a certificate is corrupt, or a certificate contains signatures that cannot be verified."
'
'        Case 43
'
'
'
'            ErrorCode = "unsupported_certificate - Received an unsupported certificate type."
'
'        Case 44
'
'
'
'            ErrorCode = "certificate_revoked - Received a certificate that was revoked by its signer."
'
'        Case 45
'
'
'
'            ErrorCode = "certificate_expired - Received a certificate has expired or is not currently valid."
'
'        Case 46
'
'
'
'            ErrorCode = "certificate_unknown - An unspecified issue took place while processing the certificate that made it unacceptable."
'
'        Case 47
'
'
'
'            ErrorCode = "illegal_parameter - Violated security parameters, such as a field in the handshake was out of range or inconsistent with other fields. This is always fatal."
'
'        Case 48
'
'
'
'            ErrorCode = "unknown_ca - Received a valid certificate chain or partial chain, but the certificate was not accepted because the CA certificate could not be located or could not be matched with a known, trusted CA. This message is always fatal."
'
'        Case 49
'
'
'
'            ErrorCode = "access_denied - Received a valid certificate, but when access control was applied, the sender did not proceed with negotiation. This message is always fatal."
'
'        Case 50
'
'
'
'            ErrorCode = "decode_error - A message could not be decoded because some field was out of the specified range or the length of the message was incorrect. This message is always fatal."
'
'        Case 51
'
'
'
'            ErrorCode = "decrypt_error - Failed handshake cryptographic operation, including being unable to correctly verify a signature, decrypt a key exchange, or validate a finished message."
'
'        Case 60
'
'
'
'            ErrorCode = "export_restriction - Detected a negotiation that was not in compliance with export restrictions; for example, attempting to transfer a 1024 bit ephemeral RSA key for the RSA_EXPORT handshake method. This message is always fatal."
'
'        Case 70
'
'
'
'            ErrorCode = "protocol_version - The protocol version the client attempted to negotiate is recognized, but not supported. For example, old protocol versions might be avoided for security reasons. This message is always fatal."
'
'        Case 71
'
'
'
'            ErrorCode = "insufficient_security - Failed negotiation specifically because the server requires ciphers more secure than those supported by the client. Returned instead of handshake_failure. This message is always fatal."
'
'        Case 80
'
'
'
'            ErrorCode = "internal_error - An internal error unrelated to the peer or the correctness of the protocol makes it impossible to continue, such as a memory allocation failure. The error is not related to protocol. This message is always fatal."
'
'        Case 90
'
'
'
'            ErrorCode = "user_cancelled - Cancelled handshake for a reason that is unrelated to a protocol failure. If the user cancels an operation after the handshake is complete, just closing the connection by sending a close_notify is more appropriate. This alert should be followed by a close_notify. This message is generally a warning."
'
'        Case 100
'
'
'
'            ErrorCode = "no_renegotiation - Sent by the client in response to a hello request or sent by the server in response to a client hello after initial handshaking. Either of these would normally lead to renegotiation; when that is not appropriate, the recipient should respond with this alert; at that point, the original requester can decide whether to proceed with the connection. One case where this would be appropriate would be where a server has spawned a process to satisfy a request; the process might receive security parameters (key length, authentication, and so on) at start-up and it might be difficult to communicate changes to these parameters after that point. This message is always a warning."
'
'        Case 255
'
'
'
'            ErrorCode = "unsupported_extension"
'        Case Else
'            ErrorCode = "Unkown error code number " & Code
'
'    End Select
'
'
'End Function
''



Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptEncryptPtr Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long, ByVal dwBufLen As Long) As Long

Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptDecryptPtr Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long) As Long

Private Declare Function CryptDuplicateHash Lib "advapi32.dll" (ByVal hHash As Long, pdwReserved As Long, ByVal dwFlags As Long, phHash As Long) As Long
'Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParamPtr Lib "advapi32.dll" Alias "CryptGetHashParam" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As Long, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
Private Declare Function CryptSignHashPtr Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As Long, ByVal dwFlags As Long, ByVal pbSignature As Long, pdwSigLen As Long) As Long
'Private Declare Function CryptSignMessage Lib "Crypt32.dll" (pSignPara As CRYPT_SIGN_MESSAGE_PARA, ByVal fDetachedSignature As Long, ByVal cToBeSigned As Long, rgpbToBeSigned As String, rgcbToBeSigned As Long, ByVal pbSignedBlob As String, pcbSignedBlob As Long) As Long
Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
'Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32.dll" (ByRef pCert As Long, ByVal dwFlags As Long, pvReserved As Any, ByRef phCryptProv As Long, ByRef pdwKeySpec As Long, ByRef pfCallerFreeProv As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "Crypt32.dll" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvReserved As Long, ByVal phCryptProv As Long, ByVal pdwKeySpec As Long, ByVal pfCallerFreeProv As Long) As Long

'If CryptAcquireCertificatePrivateKey(ptrCertCtx, CRYPT_ACQUIRE_CACHE_FLAG, 0, VarPtr(hCryptoProv), VarPtr(dwType), VarPtr(flag)) = 0 Then
'if CryptAcquireCertificatePrivateKey(hCertCntxt, 0, IntPtr.Zero, ref hCryptProv, ref keyspec, ref freecsp)


Private Const CAQ As String = "CryptAcquireContext"
Private Const cch As String = "CryptCreateHash"
Private Const CDrK As String = "CryptDeriveKey"
Private Const CDH As String = "CryptDestroyHash"
Private Const CDsK As String = "CryptDestroyKey"
Private Const CDuH As String = "CryptDuplicateHash"
Private Const CE As String = "CryptEncrypt"
Private Const CD As String = "CryptDecrypt"
Private Const CEK As String = "CryptExportKey"
Private Const CIK As String = "CryptImportKey"
Private Const CGHP As String = "CryptGetHashParam"
Private Const CGK As String = "CryptGenKey"
Private Const CGR As String = "CryptGenRandom"
Private Const CGPP As String = "CryptGetProvParam"
Private Const CGUK As String = "CryptGetUserKey"
Private Const CHD As String = "CryptHashData"
Private Const CRC As String = "CryptReleaseContext"
Private Const CSHP As String = "CryptSetHashParam"
Private Const CSKP As String = "CryptSetKeyParam"
Private Const CSH As String = "CryptSignHash"
Private Const CVS As String = "CryptVerifySignature"
Private Const abClientFinished = "client finished"
Private Const abServerFinished = "server finished"

'================================
'EVENTS
'================================
Public Event Error(ByVal Number As Long, Description As String, ByVal Source As String)

' exported key blob definitions
Private Const SIMPLEBLOB As Byte = &H1
Private Const CUR_BLOB_VERSION As Byte = &H2
Private Const PUBLICKEYBLOB As Byte = &H6
Private Const PRIVATEKEYBLOB As Byte = &H7
Private Const PLAINTEXTKEYBLOB As Byte = &H8

Private Type T_PUBLICKEYBLOB
    bType    As Byte
    bVersion As Byte
    Reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 256) As Byte
End Type
Private Const T_PUBLICKEYBLOBLEN = 276
Private Type T_EXP_PUBLICKEYBLOB
    bPublicKey(1 To T_PUBLICKEYBLOBLEN) As Byte
End Type

Private Type T_PRIVATEKEYBLOB
    bType    As Byte
    bVersion As Byte
    Reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 256)  As Byte
    prime1(1 To 128)   As Byte
    prime2(1 To 128)   As Byte
    exponent1(1 To 128)       As Byte
    exponent2(1 To 128)       As Byte
    coefficient(1 To 128)     As Byte
    privateExponent(1 To 256) As Byte
End Type
Private Const T_PRIVATEKEYBLOBLEN = 1172
Private Type T_EXP_PRIVATEKEYBLOB
    bPrivateKey(1 To T_PRIVATEKEYBLOBLEN) As Byte
End Type

Private Type T_SIMPLEKEYBLOB
    bType    As Byte
    bVersion As Byte
    Reserved As Integer
    aiKeyAlg As Long
    algid    As Long
    modulus(1 To 256) As Byte
End Type
Private Const T_SIMPLEKEYBLOBLEN = 268
Private Type T_EXP_SIMPLEKEYBLOB
    bSimpleKey(1 To T_SIMPLEKEYBLOBLEN) As Byte
End Type

Private Type SCHANNEL_ALG
    dwUse           As Long  '4 bytes
    algid           As Long  '4 bytes
    cBits           As Long  '4 bytes
    dwFlags         As Long  '4 bytes
    dwReserved      As Long  '4 bytes
End Type                     'Total 20 bytes

Private Type CRYPT_DATA_BLOB
    cbData As Long          'The count of bytes in the buffer pointed to by pbData.
    pbData As Long          'A pointer to a block of data bytes.
End Type

Private Type HMAC_INFO
    ALG_ID As Long          'HashAlgid
    pbIStr(1 To 64) As Byte '*pbInnerString (BYTE*)0x36
    cbIStr As Long          'len(cbInnerString)
    pbOStr(1 To 64) As Byte '*pbOuterString (BYTE*)0x5C
    cbOStr As Long          'len(cbOuterString)
End Type

Private Type SHAMD5_HASH
    bHash(1 To 36) As Byte
End Type

'Class Variables
Private sContainer As New Stream
Private sProvider As New Stream
Private sProvType As Long
Private PublicKey As T_EXP_PUBLICKEYBLOB
Private NewPubKey As T_PUBLICKEYBLOB
Private PrivateKey As T_EXP_PRIVATEKEYBLOB
Private MasterKey As T_EXP_SIMPLEKEYBLOB
Private NewSessionKey As T_SIMPLEKEYBLOB
Private PriFlg As Boolean
Private fClient As Boolean
Private bClientRandom() As Byte
Private bServerRandom() As Byte
Private bPreMasterKey() As Byte
Private Alg_Hash As Long

'API error function
'Private Declare Function GetLastError Lib "kernel32" () As Long

'API memory functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

'constants for API memory functions
'Private Const GMEM_MOVEABLE = &H2
'Private Const GMEM_ZEROINIT = &H40
'Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'constants for Cryptography API functions
Private Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const MS_STRONG_PROV As String = "Microsoft Strong Cryptographic Provider"
Private Const MS_DEF_RSA_SIG_PROV As String = "Microsoft RSA Signature Cryptographic Provider"
Private Const MS_DEF_RSA_SCHANNEL_PROV As String = "Microsoft RSA SChannel Cryptographic Provider"
Private Const MS_DEF_DSS_PROV As String = "Microsoft Base DSS Cryptographic Provider"
Private Const MS_DEF_DSS_DH_PROV As String = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
Private Const MS_ENH_DSS_DH_PROV As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
Private Const MS_DEF_DH_SCHANNEL_PROV As String = "Microsoft DH SChannel Cryptographic Provider"
Private Const PROV_RSA_FULL As Long = 1
Private Const PROV_RSA_SIG As Long = 2
Private Const PROV_DSS As Long = 3
Private Const PROV_FORTEZZA As Long = 4
Private Const PROV_MS_MAIL As Long = 5
Private Const PROV_SSL As Long = 6
Private Const PROV_STT_MER As Long = 7
Private Const PROV_STT_ACQ As Long = 8
Private Const PROV_STT_BRND As Long = 9
Private Const PROV_STT_ROOT As Long = 10
Private Const PROV_STT_ISS As Long = 11
Private Const PROV_RSA_SCHANNEL As Long = 12
Private Const PROV_DSS_DH As Long = 13
Private Const PROV_EC_ECDSA_SIG As Long = 14
Private Const PROV_EC_ECNRA_SIG As Long = 15
Private Const PROV_EC_ECDSA_FULL As Long = 16
Private Const PROV_EC_ECNRA_FULL As Long = 17
Private Const PROV_DH_SCHANNEL As Long = 18
Private Const PROV_SPYRUS_LYNKS As Long = 20
Private Const PROV_RNG As Long = 21
Private Const PROV_INTEL_SEC As Long = 22
Private Const PROV_RSA_AES As Long = 24
Private Const RSA_KEY_LEN = 256
Private Const CRYPT_NEWKEYSET = &H8
Private Const PP_CONTAINER = 6
Private Const AT_KEYEXCHANGE = 1
Private Const AT_SIGNATURE = 2
Private Const PP_ENUMALGS As Long = 1

'Private Const SIMPLEBLOB = 1

'Algorithm Classes
Private Const ALG_CLASS_ANY              As Long = 0
Private Const ALG_CLASS_SIGNATURE        As Long = (1 * 2 ^ 13) '8192
Private Const ALG_CLASS_MSG_ENCRYPT      As Long = (2 * 2 ^ 13) '16384
Private Const ALG_CLASS_DATA_ENCRYPT     As Long = (3 * 2 ^ 13) '24576
Private Const ALG_CLASS_HASH             As Long = (4 * 2 ^ 13) '32768
Private Const ALG_CLASS_KEY_EXCHANGE     As Long = (5 * 2 ^ 13) '40960

Private Const ALG_TYPE_ANY              As Long = 0
Private Const ALG_TYPE_DSS              As Long = (1 * 2 ^ 9) '512
Private Const ALG_TYPE_RSA              As Long = (2 * 2 ^ 9) '1024
Private Const ALG_TYPE_BLOCK            As Long = (3 * 2 ^ 9) '1536
Private Const ALG_TYPE_STREAM           As Long = (4 * 2 ^ 9) '2048
Private Const ALG_TYPE_DH               As Long = (5 * 2 ^ 9) '2560
Private Const ALG_TYPE_SECURECHANNEL    As Long = (6 * 2 ^ 9) '3072

Private Const ALG_SID_RC2               As Long = 2
Private Const ALG_SID_RC4               As Long = 1
Private Const ALG_SID_MD2               As Long = 1
Private Const ALG_SID_MD4               As Long = 2
Private Const ALG_SID_MD5               As Long = 3
Private Const ALG_SID_SHA1              As Long = 4
Private Const ALG_SID_SCHANNEL_MASTER_HASH As Long = 2
Private Const ALG_SID_SCHANNEL_MAC_KEY  As Long = 3
Private Const ALG_SID_TLS1_MASTER       As Long = 6
Private Const ALG_SID_SCHANNEL_ENC_KEY  As Long = 7
Private Const ALG_SID_SSL3SHAMD5        As Long = 8
Private Const ALG_SID_HMAC              As Long = 9
Private Const ALG_SID_TLS1PRF           As Long = 10
Private Const ALG_SID_RSA_ANY           As Long = 0

Private Const CALG_SCHANNEL_MASTER_HASH = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MASTER_HASH) '19458
Private Const CALG_SCHANNEL_MAC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MAC_KEY) '19459
Private Const CALG_TLS1_MASTER = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_TLS1_MASTER) '19462
Private Const CALG_SCHANNEL_ENC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_ENC_KEY) '19463
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2) '32769
Private Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4) '32770
Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5) '32771
Private Const CALG_SHA1 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1) '32772
Private Const CALG_SSL3_SHAMD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SSL3SHAMD5) '32776
Private Const CALG_HMAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HMAC) '32777
Private Const CALG_TLS1PRF = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_TLS1PRF) '32778
Private Const CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2) '26114
Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4) '26625
Private Const CALG_RSA_KEYX = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY) '41984

Private Const SCHANNEL_MAC_KEY As Long = 0
Private Const SCHANNEL_ENC_KEY As Long = 1
Private Const HP_HASHVAL As Long = 2 'Hash value
Private Const HP_HASHSIZE As Long = 4
Private Const HP_HMAC_INFO As Long = 5
Private Const HP_TLS1PRF_LABEL As Long = 6
Private Const HP_TLS1PRF_SEED As Long = 7
Private Const CRYPT_SERVER = 1024 '0x00000400

'constants from WinErr.h
Private Const NTE_NO_KEY As Long = &H8009000D              ' Key does not exist
Private Const NTE_BAD_SIGNATURE As Long = &H80090006       ' Bad Signature

'CryptooFilterBox constants
Private Const CFB_BUSY = 0
Private Const CFB_READY = 1
Private Const CFB_VALID = 2

Private Const ENCRYPT_ALGORITHM = CALG_RC4
Private Const ENCRYPT_BLOCK_SIZE = 1

Private Const CRYPT_EXPORTABLE = 1

'KP Constants
Private Const KP_SCHANNEL_ALG = 20
Private Const KP_CLIENT_RANDOM = 21 'for setting the Secure Channel client random data
Private Const KP_SERVER_RANDOM = 22 'for setting the Secure Channel server random data

'private property buffers
Private lCipher As Long
Public SInBuffer As New Stream

Public SOutBuffer As New Stream
Private sRndBuffer As String

Private sPublicKey As String
Private sPassword As String
Private sSignature As New Stream
Private lStatus As Long
Private sKeyFile As String
Private sMsgType As String

'TLS Key Handles
Private hCryptProv As Long
Private hMasterKey As Long
Private hReadKey As Long
Private hWriteKey As Long
Private hReadMAC As Long
Private hWriteMAC As Long
Private WRITE_MAC_KEY As String
Private READ_MAC_KEY As String
Private WRITE_KEY As String
Private READ_KEY As String

'Counters
Private SEND_SEQ_NUM(2) As Long
Private RECV_SEQ_NUM(2) As Long

Private flgFinal As Boolean

Public Function ChooseCipher(sCiphers As String) As String
    Dim N%
    Alg_Hash = 0
    N% = 1
    Do Until Alg_Hash <> 0
        lCipher = CLng(Asc(Mid$(sCiphers, N%, 1))) * 256 + CLng(Asc(Mid$(sCiphers, N% + 1, 1)))
        Select Case lCipher
            Case 4
                Alg_Hash = CALG_MD5
            Case 5
                Alg_Hash = CALG_SHA1
            Case Else
                Alg_Hash = 0
        End Select
        N% = N% + 2
    Loop
    ChooseCipher = Mid$(sCiphers, N% - 2, 1) & Mid$(sCiphers, N% - 1, 1)
End Function
Public Property Get CipherToUse2() As Long
    CipherToUse2 = lCipher

End Property
Public Property Let CipherToUse(sNewCipher As String)
    lCipher = Asc(Left$(sNewCipher, 1)) * 256 + Asc(Right$(sNewCipher, 1))
    Select Case lCipher
        Case 4
            Alg_Hash = CALG_MD5
        Case 5
            Alg_Hash = CALG_SHA1
        Case Else
            Alg_Hash = CALG_SHA1
    End Select
End Property


Public Property Get ContextHandle() As Long
    ContextHandle = hCryptProv
End Property

Public Function ExportPublicKey() As Boolean
    Const Routine As String = "Crypto.ExportPublicKey"
    Dim hPublicKey As Long
    Dim hExchngKey As Long
    Dim llen As Long
    lStatus = CFB_BUSY
    'Acquire Context
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Get handle to Public key
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
        GoTo ReleaseHandles
    End If
    'Get length of public key
    'Byval 0& is equivalent to NULL in C++ even though it is declared ByRef as Any
    If CryptExportKey(hExchngKey, hPublicKey, PUBLICKEYBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Public Key
    If CryptExportKey(hExchngKey, hPublicKey, PUBLICKEYBLOB, 0, sPublicKey, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
ReleaseHandles:
    If hPublicKey Then CryptDestroyKey hPublicKey
    If hExchngKey Then CryptDestroyKey hExchngKey
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Function

Public Sub Hash_SHAMD5(fServer As Boolean)
    Const Routine As String = "Crypto.Hash_SHAMD5"
    Dim hTLS1PRF As Long
    Dim data_blob As CRYPT_DATA_BLOB
    Dim strHash As String
    Dim bHandShakeHash() As Byte
    Dim llen As Long
    Dim bFinish() As Byte
    Dim sBuffer As String
    If CryptCreateHash(hCryptProv, CALG_TLS1PRF, hMasterKey, 0, hTLS1PRF) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles
    End If
    bHandShakeHash = StrConv(InBuffer, vbFromUnicode)
    data_blob.pbData = VarPtr(bHandShakeHash(0))
    data_blob.cbData = SInBuffer.Length '36
    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_SEED, data_blob, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
        GoTo ReleaseHandles
    End If
    If fServer Then
        bFinish = StrConv(abServerFinished, vbFromUnicode)
    Else
        bFinish = StrConv(abClientFinished, vbFromUnicode)
    End If
    data_blob.pbData = VarPtr(bFinish(0))
    data_blob.cbData = UBound(bFinish) + 1 '15
    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_LABEL, data_blob, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
        GoTo ReleaseHandles
    End If
    llen = 12
    'Prepare buffer
    sBuffer = String$(llen, Chr$(0))
    If CryptGetHashParam(hTLS1PRF, HP_HASHVAL, sBuffer, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        GoTo ReleaseHandles
    End If
    InBuffer = Chr$(&H14) & Chr$(0) & Chr$(0) & Chr$(Len(sBuffer) Mod 256) & sBuffer
    Call TLSEncrypt
    flgFinal = False
ReleaseHandles:
    If hTLS1PRF <> 0 Then CryptDestroyHash hTLS1PRF
End Sub

Public Function ImportMasterKey() As String
    Const Routine As String = "Crypto.ImportMasterKey"
    Dim hPrivateKey As Long
    Dim hExchngKey As Long
    Dim hMasterHash As Long
    Dim llen As Long
    Dim TempBlob As T_SIMPLEKEYBLOB
    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
    Dim data As CRYPT_DATA_BLOB
    Dim Algorithm As SCHANNEL_ALG
    Dim sTemp As String
    Dim bTemp() As Byte
    lStatus = CFB_BUSY
    'Acquire Context
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Prepare Master Key Blob
    TempBlob.bType = SIMPLEBLOB
    TempBlob.bVersion = CUR_BLOB_VERSION
    TempBlob.Reserved = 0
    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
    TempBlob.algid = CALG_RSA_KEYX
    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(1), RSA_KEY_LEN
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
        GoTo ReleaseHandles
    End If
    'Import & Decrypt Pre-Master Key
    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, 0, hMasterKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    'Set SCHANNEL_ENC_KEY parameters
    Algorithm.dwUse = SCHANNEL_ENC_KEY
    Algorithm.algid = CALG_RC4
    Algorithm.cBits = 128
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Set SCHANNEL_MAC_KEY parameters
    Algorithm.dwUse = SCHANNEL_MAC_KEY
    Algorithm.algid = Alg_Hash
    If Alg_Hash = CALG_MD5 Then
        Algorithm.cBits = 128
    Else
        Algorithm.cBits = 160
    End If
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add CLIENT_RANDOM
    data.cbData = UBound(bClientRandom) + 1
    data.pbData = VarPtr(bClientRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, data, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add SERVER_RANDOM
    data.cbData = UBound(bServerRandom) + 1
    data.pbData = VarPtr(bServerRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, data, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Create Master Hash
    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles
    End If
    'Create the other SCHANNEL keys
    fClient = False
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0 + CRYPT_EXPORTABLE), hReadKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    sTemp = StrConv(bPreMasterKey, vbUnicode)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    ImportMasterKey = Left$(sTemp, 48)
    llen = 0
    'Get WriteMAC Key Length
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get WriteMAC Key (used in HMAC routine)
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        WRITE_MAC_KEY = Left$(sTemp, 20)
    Else
        WRITE_MAC_KEY = Left$(sTemp, 16)
    End If
    llen = 0
    'Get ReadMAC Key Length
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get ReadMAC Key (used in HMAC routine)
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        READ_MAC_KEY = Left$(sTemp, 20)
    Else
        READ_MAC_KEY = Left$(sTemp, 16)
    End If
    llen = 0
    'Get Write Key Length
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Write Key
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    WRITE_KEY = Left$(sTemp, 16)
    llen = 0
    'Get Read Key Length
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Read Key
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    READ_KEY = Left$(sTemp, 16)
 '   Call DebugPrintString("Read Key:", READ_KEY)
 '   Call DebugPrintString("Write Key:", WRITE_KEY)
 '   Call DebugPrintString("Read MAC Key:", READ_MAC_KEY)
 '   Call DebugPrintString("Write MAC Key:", WRITE_MAC_KEY)
ReleaseHandles:
    If hPrivateKey Then CryptDestroyKey hPrivateKey
    If hMasterHash Then CryptDestroyHash hMasterHash
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
    flgFinal = True
End Function

Public Function ExportMasterKey() As Boolean
    Const Routine As String = "Crypto.ExportMasterKey"
    Dim hPubCertKey As Long 'temporary handle for Certificate Public key
    Dim hExchngKey As Long
    Dim hMasterHash As Long
    Dim llen As Long
    Dim TempBlob As T_SIMPLEKEYBLOB
    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
    Dim Algorithm As SCHANNEL_ALG
    Dim data As CRYPT_DATA_BLOB
    Dim sTemp As String
    Dim bTemp() As Byte
    Dim sPreMaster As String
    Dim tTemp As New Stream
    lStatus = CFB_BUSY

    'Acquire Context
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Get handle to Exchange key
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
        GoTo ReleaseHandles
    End If
    'Instead of creating the encrypted Pre-Master Key with the following call,
    'I chose to create the key, encrypt it with the exchange key, import it,
    'and then encrypt it using the server public key. Why you ask? The encrypted
    'Pre-Master key requires the server private key to decrypt.
    'If CryptGenKey(hCryptProv, CALG_TLS1_MASTER, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
    '    RaiseEvent Error(err.LastDllError, CGK, Routine)
    '    GoTo ReleaseHandles
    'End If
    llen = 0



    If CryptEncryptPtr(hExchngKey, 0, True, 0, tTemp.Address, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    tTemp.Clone SInBuffer
    tTemp.Length = llen

    'sTemp = InBuffer & String$(lLen - sInBuffer.Length, Chr$(0))
    If CryptEncryptPtr(hExchngKey, 0, True, 0, tTemp.Address, SInBuffer.Length, tTemp.Length) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    bPreMasterKey = tTemp.Partial

    'Prepare Master Key Blob
    TempBlob.bType = SIMPLEBLOB
    TempBlob.bVersion = CUR_BLOB_VERSION
    TempBlob.Reserved = 0
    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
    TempBlob.algid = CALG_RSA_KEYX
    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(1), RSA_KEY_LEN
    'Import & Decrypt Pre-Master Key
    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    'Import Certificate Public Key as blob
    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    llen = 0

    tTemp.Reset

    'Get length of encryption
    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    'Encrypt Pre-Master Key with Public Cert Key
  '  tTemp.Clone SInBuffer
    'tTemp.Length = llen

    sTemp = InBuffer & String$(llen - SInBuffer.Length, Chr$(0))
 '  tTemp.Concat String$(llen - SInBuffer.Length, Chr$(0))

    'If CryptEncrypt(hPubCertKey, 0, True, 0, sTemp, SInBuffer.Length, Len(sTemp)) = 0 Then
    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, SInBuffer.Length, tTemp.Length) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    SOutBuffer.Clone tTemp

'    OutBuffer = sTemp 'Save Public key encrypted Pre-Master Key
    tTemp.Reset



    'Set SCHANNEL_ENC_KEY parameters
    Algorithm.dwUse = SCHANNEL_ENC_KEY
    Algorithm.algid = CALG_RC4
    Algorithm.cBits = 128
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Set SCHANNEL_MAC_KEY parameters
    Algorithm.dwUse = SCHANNEL_MAC_KEY
    Algorithm.algid = Alg_Hash
    If Alg_Hash = CALG_MD5 Then
        Algorithm.cBits = 128
    Else
        Algorithm.cBits = 160
    End If
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add CLIENT_RANDOM
    data.cbData = UBound(bClientRandom) + 1
    data.pbData = VarPtr(bClientRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, data, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add SERVER_RANDOM
    data.cbData = UBound(bServerRandom) + 1
    data.pbData = VarPtr(bServerRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, data, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Create Master Hash
    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Create the other SCHANNEL keys
    fClient = True
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    llen = 0
    'Get WriteMAC Key Length
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get WriteMAC Key (used in HMAC routine)
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        WRITE_MAC_KEY = Left$(sTemp, 20)
    Else
        WRITE_MAC_KEY = Left$(sTemp, 16)
    End If
    llen = 0
    'Get ReadMAC Key Length
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get ReadMAC Key (used in HMAC routine)
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        READ_MAC_KEY = Left$(sTemp, 20)
    Else
        READ_MAC_KEY = Left$(sTemp, 16)
    End If
    llen = 0
    'Get Write Key Length
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Write Key
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    WRITE_KEY = Left$(sTemp, 16)
    llen = 0
    'Get Read Key Length
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Read Key
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    READ_KEY = Left$(sTemp, 16)
  '  Call DebugPrintString("Read Key:", READ_KEY)
  '  Call DebugPrintString("Write Key:", WRITE_KEY)
  '  Call DebugPrintString("Read MAC Key:", READ_MAC_KEY)
  '  Call DebugPrintString("Write MAC Key:", WRITE_MAC_KEY)
    ExportMasterKey = True 'Success
ReleaseHandles:
    Set tTemp = Nothing

    If hPubCertKey Then CryptDestroyKey hPubCertKey
    If hMasterHash Then CryptDestroyHash hMasterHash
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    flgFinal = True
    lStatus = CFB_READY
End Function


'Public Function ExportMasterKey() As Boolean
'    Const Routine As String = "Crypto.ExportMasterKey"
'    Dim hPubCertKey As Long 'temporary handle for Certificate Public key
'    Dim hExchngKey As Long
'    Dim hMasterHash As Long
'    Dim lLen As Long
'    Dim TempBlob As T_SIMPLEKEYBLOB
'    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
'    Dim Algorithm As SCHANNEL_ALG
'    Dim data As CRYPT_DATA_BLOB
'    Dim sTemp As String
'    Dim bTemp() As Byte
'    Dim sPreMaster As String
'    Dim tTemp As New Stream
'    lStatus = CFB_BUSY
'
'    'Acquire Context
'    If hCryptProv = 0 Then
'        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
'            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
'            GoTo ReleaseHandles
'        End If
'    End If
'    'Get handle to Exchange key
'    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CGUK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Instead of creating the encrypted Pre-Master Key with the following call,
'    'I chose to create the key, encrypt it with the exchange key, import it,
'    'and then encrypt it using the server public key. Why you ask? The encrypted
'    'Pre-Master key requires the server private key to decrypt.
'    'If CryptGenKey(hCryptProv, CALG_TLS1_MASTER, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
'    '    RaiseEvent Error(err.LastDllError, CGK, Routine)
'    '    GoTo ReleaseHandles
'    'End If
'    lLen = 0
'
'
'
'    If CryptEncryptPtr(hExchngKey, 0, True, 0, ByVal tTemp.Address, lLen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    tTemp.Clone SInBuffer
'    tTemp.Length = lLen
'
'    'sTemp = InBuffer & String$(lLen - sInBuffer.Length, Chr$(0))
'    If CryptEncryptPtr(hExchngKey, 0, True, 0, ByVal tTemp.Address, SInBuffer.Length, tTemp.Length) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    bPreMasterKey = tTemp.Partial
'
'    'Prepare Master Key Blob
'    TempBlob.bType = SIMPLEBLOB
'    TempBlob.bVersion = CUR_BLOB_VERSION
'    TempBlob.Reserved = 0
'    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
'    TempBlob.algid = CALG_RSA_KEYX
'    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
'    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(1), RSA_KEY_LEN
'    'Import & Decrypt Pre-Master Key
'    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Import Certificate Public Key as blob
'    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CIK, Routine)
'        GoTo ReleaseHandles
'    End If
'    lLen = 0
'
'    tTemp.Reset
'
'    'Get length of encryption
'    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, lLen, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'
''    'Get length of encryption
''    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, lLen, 0) = 0 Then
''        RaiseEvent Error(Err.LastDllError, CE, Routine)
''        GoTo ReleaseHandles
''    End If
'    'Encrypt Pre-Master Key with Public Cert Key
'    tTemp.Clone SInBuffer
'   ' tTemp.Length = lLen
'
'   ' sTemp = InBuffer & String$(lLen - sInBuffer.Length, Chr$(0))
'   tTemp.Concat ToBytes(String$(lLen - SInBuffer.Length, Chr$(0)))
'
'    If CryptEncryptPtr(hPubCertKey, 0, True, 0, ByVal tTemp.Address, SInBuffer.Length, tTemp.Length) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CE, Routine)
'        GoTo ReleaseHandles
'    End If
'    SOutBuffer.Clone tTemp
'
''    OutBuffer = sTemp 'Save Public key encrypted Pre-Master Key
'    tTemp.Reset
'
'
'
'    'Set SCHANNEL_ENC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_ENC_KEY
'    Algorithm.algid = CALG_RC4
'    Algorithm.cBits = 128
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Set SCHANNEL_MAC_KEY parameters
'    Algorithm.dwUse = SCHANNEL_MAC_KEY
'    Algorithm.algid = Alg_Hash
'    If Alg_Hash = CALG_MD5 Then
'        Algorithm.cBits = 128
'    Else
'        Algorithm.cBits = 160
'    End If
'    Algorithm.dwFlags = 0
'    Algorithm.dwReserved = 0
'    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add CLIENT_RANDOM
'    data.cbData = UBound(bClientRandom) + 1
'    data.pbData = VarPtr(bClientRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Add SERVER_RANDOM
'    data.cbData = UBound(bServerRandom) + 1
'    data.pbData = VarPtr(bServerRandom(0))
'    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, data, 0) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create Master Hash
'    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CSKP, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Create the other SCHANNEL keys
'    fClient = True
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
'        GoTo ReleaseHandles
'    End If
'    lLen = 0
'    'Get WriteMAC Key Length
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get WriteMAC Key (used in HMAC routine)
'    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        WRITE_MAC_KEY = Left$(sTemp, 20)
'    Else
'        WRITE_MAC_KEY = Left$(sTemp, 16)
'    End If
'    lLen = 0
'    'Get ReadMAC Key Length
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get ReadMAC Key (used in HMAC routine)
'    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    If Alg_Hash = CALG_SHA1 Then
'        READ_MAC_KEY = Left$(sTemp, 20)
'    Else
'        READ_MAC_KEY = Left$(sTemp, 16)
'    End If
'    lLen = 0
'    'Get Write Key Length
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Write Key
'    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    WRITE_KEY = Left$(sTemp, 16)
'    lLen = 0
'    'Get Read Key Length
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    'Get Read Key
'    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CEK, Routine)
'        GoTo ReleaseHandles
'    End If
'    ReDim bTemp(Len(MasterKeyBlob) - 1)
'    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
'    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
'    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
'        RaiseEvent Error(Err.LastDllError, CD, Routine)
'        GoTo ReleaseHandles
'    End If
'    READ_KEY = Left$(sTemp, 16)
'  '  Call DebugPrintString("Read Key:", READ_KEY)
'  '  Call DebugPrintString("Write Key:", WRITE_KEY)
'  '  Call DebugPrintString("Read MAC Key:", READ_MAC_KEY)
'  '  Call DebugPrintString("Write MAC Key:", WRITE_MAC_KEY)
'    ExportMasterKey = True 'Success
'ReleaseHandles:
'    Set tTemp = Nothing
'
'    If hPubCertKey Then CryptDestroyKey hPubCertKey
'    If hMasterHash Then CryptDestroyHash hMasterHash
''    If hCryptProv Then CryptReleaseContext hCryptProv, 0
'    flgFinal = True
'    lStatus = CFB_READY
'End Function
Public Sub IncSendSeqNum(Optional flgClear As Boolean)
    If flgClear Then
        SEND_SEQ_NUM(0) = 0
        SEND_SEQ_NUM(1) = 0
        SEND_SEQ_NUM(2) = 0
        Exit Sub
    End If
    If SEND_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
        SEND_SEQ_NUM(0) = SEND_SEQ_NUM(0) + 1
    ElseIf SEND_SEQ_NUM(1) < 16777215 Then
        SEND_SEQ_NUM(1) = SEND_SEQ_NUM(1) + 1
        SEND_SEQ_NUM(0) = 0
    ElseIf SEND_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
        SEND_SEQ_NUM(2) = SEND_SEQ_NUM(2) + 1
        SEND_SEQ_NUM(1) = 0
    Else
        SEND_SEQ_NUM(0) = 0
        SEND_SEQ_NUM(1) = 0
        SEND_SEQ_NUM(2) = 0
    End If
End Sub
Public Sub IncRecvSeqNum(Optional flgClear As Boolean)
    If flgClear Then
        RECV_SEQ_NUM(0) = 0
        RECV_SEQ_NUM(1) = 0
        RECV_SEQ_NUM(2) = 0
        Exit Sub
    End If
    If RECV_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
        RECV_SEQ_NUM(0) = RECV_SEQ_NUM(0) + 1
    ElseIf RECV_SEQ_NUM(1) < 16777215 Then
        RECV_SEQ_NUM(1) = RECV_SEQ_NUM(1) + 1
        RECV_SEQ_NUM(0) = 0
    ElseIf RECV_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
        RECV_SEQ_NUM(2) = RECV_SEQ_NUM(2) + 1
        RECV_SEQ_NUM(1) = 0
    Else
        RECV_SEQ_NUM(0) = 0
        RECV_SEQ_NUM(1) = 0
        RECV_SEQ_NUM(2) = 0
    End If
End Sub

Public Sub IncSendSeqSet(ByVal index0 As Byte, ByVal index1 As Byte, ByVal index2 As Byte)

    SEND_SEQ_NUM(0) = index0
    SEND_SEQ_NUM(1) = index1
    SEND_SEQ_NUM(2) = index2

End Sub
Public Sub IncRecvSeqSet(ByVal index0 As Byte, ByVal index1 As Byte, ByVal index2 As Byte)

    RECV_SEQ_NUM(0) = index0
    RECV_SEQ_NUM(1) = index1
    RECV_SEQ_NUM(2) = index2

End Sub

Public Property Get msgType() As String
    msgType = sMsgType
End Property

Public Property Let msgType(vNewValue As String)
    sMsgType = vNewValue
End Property

Public Function VerifySig() As Boolean
    Const Routine As String = "VerifySig"
    Dim hHash As Long
    Dim hPubCertKey As Long
    Dim sDescription As String
    'Create a hash object.
    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles:
    End If
    'Hash the data
    If CryptHashData(hHash, InBuffer, SInBuffer.Length, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles:
    End If
    'Import Certificate Public Key as blob
    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    sDescription = vbNullString
    If CryptVerifySignature(hHash, Signature, Len(Signature), hPubCertKey, sDescription, 0) = 0 Then
        VerifySig = False
    Else
        VerifySig = True
    End If
ReleaseHandles:
    If hPubCertKey Then CryptDestroyKey hPubCertKey
    If hHash Then CryptDestroyHash hHash
End Function


Public Property Get InBuffer() As String
    InBuffer = Convert(SInBuffer.Partial)
End Property
Public Property Let InBuffer(vNewValue As String)
    SInBuffer.Length = Len(vNewValue)
    SInBuffer.Placeat Convert(vNewValue)
End Property


Public Property Let KeyFile(sNewFileName As String)
    sKeyFile = sNewFileName
End Property

Public Property Get KeyFile() As String
    KeyFile = sKeyFile
End Property

Public Property Get KeyName() As String
    KeyName = App.CompanyName
End Property
Public Function MD5_AddToHash(ByVal strToHash As String, hHash) As Long
    Const Routine As String = "Crypto.MD5_AddToHash"
    If hHash = 0 Then
        If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
            RaiseEvent Error(Err.LastDllError, cch, Routine)
            Exit Function
        End If
    End If
    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        Exit Function
    End If
    MD5_AddToHash = hHash
End Function
Public Function CopyHash(hHashToCopy) As Long
    Const Routine As String = "Crypto.CopyHash"
    Dim phHash As Long
    If CryptDuplicateHash(hHashToCopy, ByVal 0&, 0&, phHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDuH, Routine)
        CopyHash = 0
    Else
        CopyHash = phHash
    End If
End Function

Public Function MD5_GetHash(hHash As Long) As String
    Const Routine As String = "Crypto.MD5_GetHash"
    Dim strHash As String
    Dim llen As Long
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(llen, vbNullChar)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then
        CryptDestroyHash hHash
        hHash = 0
    End If
    MD5_GetHash = strHash
End Function
Public Function MD5_Hash(ByVal TheString As String) As String
    Const Routine As String = "Crypto.MD5_Hash"
    Dim strHash As String
    Dim hHash As Long
    Dim llen As Long
    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        Exit Function
    End If
    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    'Get length of hash
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(llen, vbNull)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then CryptDestroyHash hHash
    MD5_Hash = strHash
End Function
Public Function SHA1_AddToHash(ByVal strToHash As String, hHash) As Long
    Const Routine As String = "Crypto.SHA1_AddToHash"
    If hHash = 0 Then
        If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
            RaiseEvent Error(Err.LastDllError, cch, Routine)
            Exit Function
        End If
    End If
    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    SHA1_AddToHash = hHash
End Function
Public Function SHA1_CopyHash() As Long

End Function

Public Function SHA1_GetHash(hHash As Long) As String
    Const Routine As String = "Crypto.SHA1_GetHash"
    Dim strHash As String
    Dim llen As Long
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(llen, vbNullChar)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then
        CryptDestroyHash hHash
        hHash = 0
    End If
    SHA1_GetHash = strHash
End Function
Public Function SHA1_Hash(ByVal TheString As String) As String
    Const Routine As String = "Crypto.SHA1_Hash"
    Dim strHash As String
    Dim hHash As Long
    Dim llen As Long
    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        Exit Function
    End If
    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(llen, vbNull)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then CryptDestroyHash hHash
    SHA1_Hash = strHash
End Function

Public Property Get OutBuffer() As String
    OutBuffer = Convert(SOutBuffer.Partial)
End Property
Public Property Let OutBuffer(ByRef vNewValue As String)
    SOutBuffer.Length = Len(vNewValue)
    SOutBuffer.Placeat Convert(vNewValue)
End Property
Public Property Let PubKey(NewKey As String)
    sPublicKey = NewKey
End Property

Public Property Get PubKey() As String
    PubKey = StrConv(PublicKey.bPublicKey(), vbUnicode)
End Property

'Public Property Let PubKey(NewKey As String)
'    sPublicKey.Reset
'    sPublicKey.Concat ToBytes(NewKey)
'End Property
'Public Property Get PubKey() As String
'    PubKey = StrConv(PublicKey.bPublicKey(), vbUnicode)
'End Property


Private Function SendSeqNum() As String
    Dim bTmp(1 To 8) As Byte
    CopyMemory bTmp(1), SEND_SEQ_NUM(0), 3
    CopyMemory bTmp(4), SEND_SEQ_NUM(1), 3
    CopyMemory bTmp(7), SEND_SEQ_NUM(2), 2
    SendSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
End Function
Private Function RecvSeqNum() As String
    Dim bTmp(1 To 8) As Byte
    CopyMemory bTmp(1), RECV_SEQ_NUM(0), 3
    CopyMemory bTmp(4), RECV_SEQ_NUM(1), 3
    CopyMemory bTmp(7), RECV_SEQ_NUM(2), 2
    RecvSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
End Function

Private Function ReadKeyFile(sFileName As String, bData() As Byte) As Boolean
    Dim lFileHandle As Long
    ' Reads a byte array from a file.
    ' If Data is fixed, it must be of the expected size.
    On Error GoTo ErrorHandler
    lFileHandle = FreeFile
    Open sFileName For Binary As lFileHandle
'    ReDim bData(1 To LOF(lFileHandle))
    Get lFileHandle, , bData
    Close lFileHandle
    ReadKeyFile = True
    Exit Function
ErrorHandler:
    ReadKeyFile = False
End Function
Public Property Get RndBuffer() As String
    RndBuffer = sRndBuffer
End Property
Public Function GetPubKey()
    'Check if PublicKey loaded
    If PublicKey.bPublicKey(1) = 0 Then
        If Not GetPubKeyBlob Then Exit Function
    End If
    PriFlg = False
End Function
Public Function EncryptSessionKey()
    Const Routine As String = "Crypto.EncryptSessionKey"
    Dim hPublicKey As Long
    Dim sTemp As New Stream
    Dim llen As Long
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    If PriFlg Then
        'Import Private Key
        If CryptImportKey(hCryptProv, PrivateKey, Len(PrivateKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
            RaiseEvent Error(Err.LastDllError, CIK, Routine)
            GoTo ReleaseHandles
        End If
    Else
        'Import Public Key
        If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
            RaiseEvent Error(Err.LastDllError, CIK, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Get length of key
    If CryptEncryptPtr(hPublicKey, 0, True, 0, sTemp.Address, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    'Prepare buffer

    sTemp.Length = llen
    sTemp.Placeat modMemory.Convert(sPassword), 0, Len(sPassword)
   ' sTemp = sPassword & String$(llen - Len(sPassword), Chr$(0))
    'Encrypt Session Key
    If CryptEncryptPtr(hPublicKey, 0, True, 0, sTemp.Address, Len(sPassword), llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    SOutBuffer.Clone sTemp
ReleaseHandles:
    Set sTemp = Nothing
    If hPublicKey Then CryptDestroyKey hPublicKey
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
End Function
Public Function DecryptSessionKey()
    Const Routine As String = "Crypto.DecryptSessionKey"
    Dim hPublicKey As Long
    Dim sTemp As String
    Dim llen As Long
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    If PriFlg Then
        'Import Private Key
        If CryptImportKey(hCryptProv, PrivateKey, Len(PrivateKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
            RaiseEvent Error(Err.LastDllError, CIK, Routine)
            GoTo ReleaseHandles
        End If
    Else
        'Import Public Key
        If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPublicKey) = 0 Then
            RaiseEvent Error(Err.LastDllError, CIK, Routine)
            GoTo ReleaseHandles
        End If
    End If
    sTemp = InBuffer
    If CryptDecrypt(hPublicKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    OutBuffer = Left$(sTemp, 32)
ReleaseHandles:
    If hPublicKey Then CryptDestroyKey hPublicKey
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
End Function

Private Function GetPubKeyBlob() As Boolean
    If (ReadKeyFile(App.path & "\" & sKeyFile, PublicKey.bPublicKey)) Then
        GetPubKeyBlob = True
    Else
        GetPubKeyBlob = False
    End If
End Function

Public Function GetPriKey(Optional sRet As Boolean) As String
    'Check if PrivateKey loaded
    If PrivateKey.bPrivateKey(1) = 0 Then
        If Not GetPriKeyBlob Then Exit Function
    End If
    PriFlg = True
End Function
Private Function GetPriKeyBlob() As Boolean
    If (ReadKeyFile(App.path & "\" & sKeyFile, PrivateKey.bPrivateKey)) Then
        GetPriKeyBlob = True
    Else
        GetPriKeyBlob = False
    End If
End Function


Public Property Let ClientRandom(sClientRandom As String)
    ReDim bClientRandom(0)
    bClientRandom = StrConv(sClientRandom, vbFromUnicode)
End Property
Public Property Let ServerRandom(sServerRandom As String)
    ReDim bServerRandom(0)
    bServerRandom = StrConv(sServerRandom, vbFromUnicode)
End Property

Public Property Let PreMasterKey(sPreMasterKey As String)
    ReDim bPreMasterKey(0)
    bPreMasterKey = StrConv(sPreMasterKey, vbFromUnicode)
End Property

Public Function SetPubKey() As Boolean
    Dim bNewKey() As Byte
    bNewKey = StrConv(StrConv(PublicKey.bPublicKey(), vbUnicode), vbFromUnicode) 'StrConv(sPublicKey, vbFromUnicode)
    NewPubKey.bType = PUBLICKEYBLOB
    NewPubKey.bVersion = CUR_BLOB_VERSION
    NewPubKey.Reserved = 0
    NewPubKey.aiKeyAlg = CALG_RSA_KEYX
    NewPubKey.magic = &H31415352 '"RSA1"
    NewPubKey.bitlen = Len(sPublicKey) * 8
    NewPubKey.pubexp = 65537 '0x00010001
    CopyMemory NewPubKey.modulus(1), bNewKey(0), Len(sPublicKey)
    CopyMemory PublicKey.bPublicKey(1), NewPubKey, Len(sPublicKey) + 20
    SetPubKey = True
End Function

Public Property Get Signature() As String
    Signature = ToString(sSignature.Partial)
End Property
Public Property Let Signature(vNewValue As String)
    sSignature.Length = Len(vNewValue)
    sSignature.Placeat ToBytes(vNewValue)
End Property
Public Sub Sign()
    Const Routine As String = "Crypto.Sign"
    Dim sDescription As String
    Dim hHash As Long
    Dim lSignatureLen As Long
    'switch Status property
    lStatus = CFB_BUSY
    'init Signature property
    sSignature.Reset
    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
        GoTo ReleaseHandles:
    End If
    'Create a hash object.
    If CryptCreateHash(hCryptProv, Alg_Hash, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles:
    End If
    If CryptHashData(hHash, InBuffer, SInBuffer.Length, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles:
    End If
    'Sign hash object.
    'Determine size of signature.
   ' sDescription = vbNullChar
    If CryptSignHashPtr(hHash, AT_KEYEXCHANGE, ByVal 0&, 0, ByVal 0&, lSignatureLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSH, Routine)
        GoTo ReleaseHandles:
    End If
    sSignature.Reset
    sSignature.Concat String(lSignatureLen, vbNullChar)
    'Sign hash object (with signature key).
    If CryptSignHashPtr(hHash, AT_KEYEXCHANGE, ByVal 0&, 0, ByVal sSignature.Address, lSignatureLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSH, Routine)
        GoTo ReleaseHandles:
    End If
ReleaseHandles:
    If hHash Then CryptDestroyHash hHash
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub

Public Sub TLSDecrypt()
    Const Routine As String = "clsCrypt.TLSDecrypt"
    Dim sBuffer As String
    lStatus = CFB_BUSY
    'Get handle to the provider
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If

    sBuffer = InBuffer
    'Decrypt data
    If CryptDecrypt(hReadKey, 0, False, 0, sBuffer, Len(sBuffer)) = 0 Then
        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
        GoTo ReleaseHandles
    End If
    If Not VerifyHMAC(sBuffer) Then
        RaiseEvent Error(&H80090002, "VerifyHMAC", Routine)
    End If
   ' SOutBuffer.Clone SInBuffer

    OutBuffer = sBuffer


ReleaseHandles:
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub

Public Sub TLSEncrypt()
    Const Routine As String = "clsCrypt.TLSEncrypt"
    Dim sBuffer As String
    Dim buflen As Long
    'switch Status property
    lStatus = CFB_BUSY
    'Get handle to the provider
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
   ' Debug.Print "Final = " & CBool(flgFinal)
    'Add HMAC to message
    'Choose either HMAC routine, as both should deliver the same result
    sBuffer = InBuffer & HMAC(Alg_Hash, WRITE_MAC_KEY, InBuffer, True)
'    sBuffer = InBuffer & HMAC2(hWriteMAC, InBuffer, True)
    buflen = Len(sBuffer)
    'Null string returns required buffer length
    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, vbNullString, Len(sBuffer), buflen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    If buflen > Len(sBuffer) Then 'Adjust buffer length if necessary
        sBuffer = sBuffer & String$(buflen - Len(sBuffer), Chr$(0))
    End If
    'Encrypt data
    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, sBuffer, Len(sBuffer), buflen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    OutBuffer = sBuffer
ReleaseHandles:
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub

Private Function VerifyHMAC(ByRef DecryptedRecord As String) As Boolean
    Dim N%
    'Verify the Message Authentication Code
    Dim AppendedHMAC As String
    Dim CalculatedHMAC As String
    Dim llen As Long
    If Alg_Hash = CALG_MD5 Then
        llen = Len(DecryptedRecord) - 16
    Else
        llen = Len(DecryptedRecord) - 20
    End If
    AppendedHMAC = Mid(DecryptedRecord, llen + 1)
    DecryptedRecord = Mid(DecryptedRecord, 1, llen)
    'Either HMAC routine should produce the correct result
    CalculatedHMAC = HMAC(Alg_Hash, READ_MAC_KEY, DecryptedRecord, False)
'    CalculatedHMAC = HMAC2(hReadMAC, DecryptedRecord, False)
    If CalculatedHMAC = AppendedHMAC Then
        VerifyHMAC = True
    Else
        VerifyHMAC = False
     '   Debug.Print "HMAC Recv = ";
     '   For N% = 1 To Len(AppendedHMAC)
     '       Debug.Print Right$("0" & Hex$(Asc(Mid$(AppendedHMAC, N%, 1))), 2) & " ";
     '   Next N%
     '   Debug.Print
     '   Debug.Print "HMAC Calc = ";
     '   For N% = 1 To Len(CalculatedHMAC)
     '       Debug.Print Right$("0" & Hex$(Asc(Mid$(CalculatedHMAC, N%, 1))), 2) & " ";
     '   Next N%
     '   Debug.Print
    End If
End Function

Private Function WriteFile(sFileName As String, bData() As Byte) As Boolean
    Dim lFileHandle As Long
    ' Writes a byte array out to a file.
    On Error GoTo ErrorHandler
    lFileHandle = FreeFile
    Open sFileName For Binary As lFileHandle
    Put lFileHandle, , bData
    Close lFileHandle
    WriteFile = True
    Exit Function
ErrorHandler:
    WriteFile = False
End Function

Public Sub Encrypt()
    Const Routine As String = "Crypto.Encrypt"
    Dim hExchgKey As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim sCryptBuffer As String
    Dim lCryptLength As Long
    Dim lCryptBufLen As Long
    'switch Status property
    lStatus = CFB_BUSY
    'Get handle to the provider
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Create a hash object.
    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles
    End If
    'Hash in the password data.
    If CryptHashData(hHash, sPassword, Len(sPassword), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles
    End If
    'Derive a session key from the hash object.
    If CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    'Destroy the hash object.
    CryptDestroyHash hHash
    hHash = 0
    'Prepare a string buffer for the CryptEncrypt function
    lCryptLength = SInBuffer.Length
    lCryptBufLen = lCryptLength * 2
    sCryptBuffer = String(lCryptBufLen, vbNullChar)
    LSet sCryptBuffer = InBuffer
    'Encrypt data
    If CryptEncrypt(hKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CE, Routine)
        'GoTo ReleaseHandles
    End If
    OutBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
ReleaseHandles:
    If hKey Then CryptDestroyKey hKey
    If hExchgKey Then CryptDestroyKey hExchgKey
    If hHash Then CryptDestroyHash hHash
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub
Public Function HMAC2(h_HMAC As Long, ByVal sMsg As String, flgSend As Boolean) As String
    'HMAC_MD5(Key, m) = MD5((Key XOR opad) ++ MD5((Key XOR ipad) ++ m)
    '(++ means concatenate, "opad" is the bytes "5c 5c ... 5c",
    'and "ipad" is the bytes "36 36 ... 36").
    Const Routine As String = "Crypto.HMAC2"
    Dim strHash As String
    Dim hMACHash As Long
    Dim llen As Long
    Dim MAC_Header As String
    Dim HMACInfo As HMAC_INFO
    'MAC_Header = Sequence (8 bytes) + Type (1 byte) + Version (2 bytes) + msg Length (variable)
    If flgSend Then
        MAC_Header = SendSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
     '   Debug.Print "MAC Header on Send:"
    Else
        MAC_Header = RecvSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
       ' Debug.Print "MAC Header on Receive:"
    End If
    Dim N%
'    For N% = 1 To Len(MAC_Header)
'        Debug.Print Right$("0" & Hex$(Asc(Mid$(MAC_Header, N%, 1))), 2) & " ";
'    Next N%
'    Debug.Print
    'if inner & outer strings and lengths set to 0, defaults are used
    'They should already be zeroed, but to make sure
    ZeroMemory HMACInfo.ALG_ID, Len(HMACInfo)
    HMACInfo.ALG_ID = Alg_Hash
    If CryptCreateHash(hCryptProv, CALG_HMAC, h_HMAC, 0, hMACHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles
    End If
    If CryptSetHashParam(hMACHash, HP_HMAC_INFO, HMACInfo, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSHP, Routine)
        GoTo ReleaseHandles
    End If
    If CryptHashData(hMACHash, MAC_Header & sMsg, Len(MAC_Header & sMsg), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles
    End If
    If CryptGetHashParam(hMACHash, HP_HASHVAL, vbNullString, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        GoTo ReleaseHandles
    End If
    strHash = String$(llen, Chr$(0))
    If CryptGetHashParam(hMACHash, HP_HASHVAL, strHash, llen, 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGHP, Routine)
        GoTo ReleaseHandles
    End If
    HMAC2 = strHash
    'Call DebugPrintString("HMAC2", HMAC2)
ReleaseHandles:
    If hMACHash <> 0 Then CryptDestroyHash hMACHash
End Function
Public Function HMAC(HashAlg As Long, sKey As String, sMsg As String, flgSend As Boolean) As String
    'Note: if key is longer than the blocksize (64 for MD5/SHA1), the key should
    'be hashed first to reduce it's length
    Const Routine As String = "Crypto.HMAC"
    Dim MAC_Header As String
    Dim o_key_pad As String
    Dim i_key_pad As String
    Dim N%
    'MAC_Header = Sequence (8 bytes) + Type (1 byte) + Version (2 bytes) + msg Length (variable)
    If flgSend Then
        MAC_Header = SendSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
    '    Debug.Print "MAC Header on Send:"
    Else
        MAC_Header = RecvSeqNum & sMsgType & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
    '    Debug.Print "MAC Header on Receive:"
    End If
 '   For N% = 1 To Len(MAC_Header)
 '       Debug.Print Right$("0" & Hex$(Asc(Mid$(MAC_Header, N%, 1))), 2) & " ";
 '   Next N%
 '   Debug.Print
    If Len(sKey) > 64 Then
        If HashAlg = CALG_MD5 Then
            sKey = MD5_Hash(sKey)
        ElseIf HashAlg = CALG_SHA1 Then
            sKey = SHA1_Hash(sKey)
        End If
    End If
    o_key_pad = String$(64, &H5C)
    i_key_pad = String$(64, &H36)
    For N% = 1 To Len(sKey)
        Mid$(o_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H5C)
        Mid$(i_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H36)
    Next N%
    If HashAlg = CALG_MD5 Then
        HMAC = MD5_Hash(o_key_pad & MD5_Hash(i_key_pad & MAC_Header & sMsg))
    ElseIf HashAlg = CALG_SHA1 Then
        HMAC = SHA1_Hash(o_key_pad & SHA1_Hash(i_key_pad & MAC_Header & sMsg))
    End If
    'Call DebugPrintString("HMAC", HMAC)
End Function

Public Sub GenRandom(llen As Long)
    Const Routine As String = "Crypto.GenRandom"
    Dim sRandom As String
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    sRandom = String$(llen, Chr$(0))
    If CryptGenRandom(hCryptProv, llen, sRandom) = 0 Then
        sRndBuffer = ""
    Else
        sRndBuffer = sRandom
    End If
ReleaseHandles:
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
End Sub

Private Sub DebugPrintString(sDescr As String, strToPrint As String)
    Dim N%
    Debug.Print sDescr & ":"
    For N% = 1 To Len(strToPrint)
        Debug.Print Right$("0" & Hex$(Asc(Mid$(strToPrint, N%, 1))), 2) & " ";
        If N% Mod 16 = 0 Then Debug.Print
    Next N%
    Debug.Print
End Sub
Public Sub Decrypt()
    Const Routine As String = "Crypto.Decrypt"
    Dim hExchgKey As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim sCryptBuffer As String
    Dim lCryptBufLen As Long
    Dim lCryptPoint As Long
    Dim lPasswordPoint As Long
    Dim lPasswordCount As Long
    'switch Status property
    lStatus = CFB_BUSY
    'Init OutBuffer
    OutBuffer = ""
    'Get handle to the default provider.
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Create a hash object.
    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
        RaiseEvent Error(Err.LastDllError, cch, Routine)
        GoTo ReleaseHandles
    End If
    'Hash in the password data.
    If CryptHashData(hHash, sPassword, Len(sPassword), 0) = 0 Then
        RaiseEvent Error(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles
    End If
    'Derive a session key from the hash object.
    If CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    'Destroy the hash object.
    CryptDestroyHash hHash
    hHash = 0
    'Prepare sCryptBuffer for CryptDecrypt
    lCryptBufLen = SInBuffer.Length * 2
    sCryptBuffer = String(lCryptBufLen, vbNullChar)
    LSet sCryptBuffer = InBuffer
    'Decrypt data
    If CryptDecrypt(hKey, 0, 1, 0, sCryptBuffer, lCryptBufLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    OutBuffer = Mid$(sCryptBuffer, 1, SInBuffer.Length)
ReleaseHandles:
    If (hKey) Then CryptDestroyKey hKey
    If hExchgKey Then CryptDestroyKey hExchgKey
    If hHash Then CryptDestroyHash hHash
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub
Public Sub CreateKeyPair()
    Const Routine As String = "Crypto.CreateKeyPair"
    Dim hPublicKey As Long
    Dim llen As Long
    Dim PublicKey  As T_EXP_PUBLICKEYBLOB
    Dim PrivateKey As T_EXP_PRIVATEKEYBLOB
    'Creates a new key pair and writes out a file for each part
    'called "FileName.fpu" and "FileName.fpr" respectively in app.path
    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
        'MsgBox ("Error " & Hex$(err.LastDllError) & " during CryptAcquireContext!")
        RaiseEvent Error(Err.LastDllError, CAQ, Routine)
        GoTo ReleaseHandles
    End If
    ' Create new Exchange key pair.
    ' The use of &H80000000 guarantees a 2048 bit key.
    If CryptGenKey(hCryptProv, AT_KEYEXCHANGE, &H8000000 Or CRYPT_EXPORTABLE, hPublicKey) = 0 Then
        RaiseEvent Error(Err.LastDllError, CGK, Routine)
        GoTo ReleaseHandles
    End If
    ' Copy the keys to variables.
    llen = LenB(PublicKey)
    If CryptExportKey(hPublicKey, 0, PUBLICKEYBLOB, 0, PublicKey, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    llen = LenB(PrivateKey)
    If CryptExportKey(hPublicKey, 0, PRIVATEKEYBLOB, 0, PrivateKey, llen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ' See if either key already exists.
    If Dir(App.path & "\" & sKeyFile & ".fpu") <> "" And _
        Dir(App.path & "\" & sKeyFile & ".fpr") <> "" Then
           ' If MsgBox("Key files already exist!" & vbCrLf & "Overwrite them?", _
           '     vbYesNo) = vbNo Then Exit Sub
    End If
    ' Create the key files.
    If Not (WriteFile(App.path & "\" & sKeyFile & ".fpu", PublicKey.bPublicKey)) Then
       ' MsgBox "Write file '" & sKeyFile & ".fpu' failed!", vbExclamation
    End If
    If Not (WriteFile(App.path & "\" & sKeyFile & ".fpr", PrivateKey.bPrivateKey)) Then
       ' MsgBox "Write file '" & sKeyFile & ".fpr' failed!", vbExclamation
    End If
ReleaseHandles:
    ' Destroy the key.
    CryptDestroyKey hPublicKey: hPublicKey = 0
    ' Release Provider
    CryptReleaseContext hCryptProv, 0: hCryptProv = 0
End Sub

Public Property Get Status() As Long
    Status = lStatus
End Property
Private Function InitUser() As Long
    Const Routine As String = "Crypto.InitUser"
    Dim hCryptKey As Long
    Dim avProviderData(1000) As Byte
    Dim lProviderDataLen As Long
    Dim lDataSize As Long
    Dim sUserName As String
    Dim lPoint As Long
    Dim sBuffer As String
    Dim strMsg As String
    'Attempt to acquire a handle to the key container.
    If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, 0) = 0 Then
        'Create key container.
        If CryptAcquireContext(hCryptProv, ByVal sContainer.Address, ByVal sProvider.Address, sProvType, CRYPT_NEWKEYSET) = 0 Then
            RaiseEvent Error(Err.LastDllError, CAQ, Routine)
            Exit Function
        End If
        'Get name of key container.
        lProviderDataLen = Len(avProviderData(0)) * (UBound(avProviderData) + 1)
        If CryptGetProvParam(hCryptProv, PP_CONTAINER, avProviderData(0), lProviderDataLen, 0) = 0 Then
            RaiseEvent Error(Err.LastDllError, CGPP, Routine)
            avProviderData(0) = 0
        End If
        'Get sUserName from avProviderData()
        lPoint = LBound(avProviderData)
        While lPoint <= UBound(avProviderData)
            If avProviderData(lPoint) <> 0 Then
                sUserName = sUserName & Chr$(avProviderData(lPoint))
            Else
                lPoint = UBound(avProviderData)
            End If
            lPoint = lPoint + 1
        Wend
        'MsgBox ("Create key container " & sUserName)
    End If
    'Attempt to get handle to signature key
    'SChannel does not support Signature Key
    If modMemory.Convert(sProvider.Partial) <> MS_DEF_RSA_SCHANNEL_PROV Then
        If CryptGetUserKey(hCryptProv, AT_SIGNATURE, hCryptKey) = 0 Then
            If Err.LastDllError = NTE_NO_KEY Then
               ' MsgBox ("System does not have Signature Key pair!" & vbCrLf & "One will be created.")
                If CryptGenKey(hCryptProv, AT_SIGNATURE, 0, hCryptKey) = 0 Then
                    RaiseEvent Error(Err.LastDllError, CGK, Routine)
                    Exit Function
                Else
                    CryptDestroyKey hCryptProv
                End If
            Else
                RaiseEvent Error(Err.LastDllError, CGUK, Routine)
                Exit Function
            End If
        End If
    End If
    'Attempt to get handle to exchange key
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hCryptKey) = 0 Then
        If Err.LastDllError = NTE_NO_KEY Then
           ' MsgBox ("System does not have Key Exchange pair!" & vbCrLf & "One will be created.")
            ' The use of &H80000000 guarantees a 2048 bit key.
            If CryptGenKey(hCryptProv, AT_KEYEXCHANGE, &H8000000, hCryptKey) = 0 Then
                RaiseEvent Error(Err.LastDllError, CGK, Routine)
                Exit Function
            Else
                CryptDestroyKey hCryptProv
            End If
        Else
            RaiseEvent Error(Err.LastDllError, CGUK, Routine)
            Exit Function
        End If
    End If
    'Enumerate the supported algorithms.
    Dim flag As Long
    Dim algid As Long
    Dim dwBits  As Long
    Dim namelen As Long
    Dim algType As String
    Dim lPtr As Long
    Dim name As String
    flag = 1
    If CryptGetProvParam(hCryptProv, PP_ENUMALGS, ByVal vbNullString, lPtr, flag) = 0 Then
            RaiseEvent Error(Err.LastDllError, CGPP, Routine)
            Exit Function
        End If
    Dim B() As Byte
    Dim bName() As Byte
    ReDim B(lPtr - 1)
    strMsg = strMsg & vbCrLf & "ALGID" & vbTab & "dwBits" & vbTab & "algType" & String(2, vbTab) & "len" & vbTab & "Name" & vbCrLf _
        & String(5, "-") & vbTab & String(6, "-") & vbTab & String(12, "-") & vbTab & String(3, "-") & vbTab & String(3, "-") & vbCrLf
    While CryptGetProvParam(hCryptProv, PP_ENUMALGS, B(0), lPtr, flag) <> 0
        flag = 2
        CopyMemory algid, B(0), LenB(algid)
        CopyMemory dwBits, B(4), LenB(dwBits)
        CopyMemory namelen, B(8), LenB(namelen)
        ReDim bName(namelen - 2)
        CopyMemory bName(0), B(12), namelen - 1
        name = StrConv(bName, vbUnicode)
        Select Case (algid And &HF000)
            Case ALG_CLASS_ANY
                algType = "Any         "
            Case ALG_CLASS_SIGNATURE
                algType = "Signature   "
            Case ALG_CLASS_MSG_ENCRYPT
                algType = "Msg_Encrypt "
            Case ALG_CLASS_DATA_ENCRYPT
                algType = "Data_Encrypt"
            Case ALG_CLASS_HASH
                algType = "Hash        "
            Case ALG_CLASS_KEY_EXCHANGE
                algType = "Exchange    "
            Case Else
                algType = "Unknwon     "
        End Select
        strMsg = strMsg & algid & vbTab & Right("     " & dwBits, 6) & vbTab & algType & vbTab & namelen & vbTab & name & vbCrLf
    Wend
   'Debug.Print strMsg
    InitUser = True
End Function
Private Sub Class_Initialize()
    sContainer.Concat modMemory.Convert("chrome@winternet.com" & vbNullChar)
    sProvider.Concat modMemory.Convert(MS_DEF_RSA_SCHANNEL_PROV)
    sProvType = PROV_RSA_SCHANNEL

    If Not InitUser = True Then
      '  MsgBox ("InitUser failed")
    End If
End Sub
'Public Property Get Password() As String
'    Password = sPassword
'End Property
'
'Public Property Let Password(vNewValue As String)
'    sPassword = vNewValue
'End Property

Private Sub Class_Terminate()
    CryptReleaseContext hCryptProv, 0
    Set sContainer = Nothing
    Set sProvider = Nothing
    Set SInBuffer = Nothing
    Set SOutBuffer = Nothing
End Sub

Public Function ErrorCode(ByVal Code As Byte) As String

    Select Case Code
        Case 0


            ErrorCode = "close_notify - Notifies the recipient that the sender will not send any more messages on this connection."

        Case 10



            ErrorCode = "unexpected_message - Received an inappropriate message This alert should never be observed in communication between proper implementations. This message is always fatal."

        Case 20



            ErrorCode = "bad_record_mac - Received a record with an incorrect MAC. This message is always fatal."

        Case 21



            ErrorCode = "decryption_failed - Decryption of a TLSCiphertext record is decrypted in an invalid way: either it was not an even multiple of the block length or its padding values, when checked, were not correct. This message is always fatal."

        Case 22



            ErrorCode = "record_overflow - Received a TLSCiphertext record which had a length more than 2^14+2048 bytes, or a record decrypted to a TLSCompressed record with more than 2^14+1024 bytes. This message is always fatal."

        Case 30



            ErrorCode = "decompression_failure - Received improper input, such as data that would expand to excessive length, from the decompression function. This message is always fatal."

        Case 40



            ErrorCode = "handshake_failure - Indicates that the sender was unable to negotiate an acceptable set of security parameters given the options available. This is a fatal error."

        Case 42



            ErrorCode = "bad_certificate - There is a problem with the certificate, for example, a certificate is corrupt, or a certificate contains signatures that cannot be verified."

        Case 43



            ErrorCode = "unsupported_certificate - Received an unsupported certificate type."

        Case 44



            ErrorCode = "certificate_revoked - Received a certificate that was revoked by its signer."

        Case 45



            ErrorCode = "certificate_expired - Received a certificate has expired or is not currently valid."

        Case 46



            ErrorCode = "certificate_unknown - An unspecified issue took place while processing the certificate that made it unacceptable."

        Case 47



            ErrorCode = "illegal_parameter - Violated security parameters, such as a field in the handshake was out of range or inconsistent with other fields. This is always fatal."

        Case 48



            ErrorCode = "unknown_ca - Received a valid certificate chain or partial chain, but the certificate was not accepted because the CA certificate could not be located or could not be matched with a known, trusted CA. This message is always fatal."

        Case 49



            ErrorCode = "access_denied - Received a valid certificate, but when access control was applied, the sender did not proceed with negotiation. This message is always fatal."

        Case 50



            ErrorCode = "decode_error - A message could not be decoded because some field was out of the specified range or the length of the message was incorrect. This message is always fatal."

        Case 51



            ErrorCode = "decrypt_error - Failed handshake cryptographic operation, including being unable to correctly verify a signature, decrypt a key exchange, or validate a finished message."

        Case 60



            ErrorCode = "export_restriction - Detected a negotiation that was not in compliance with export restrictions; for example, attempting to transfer a 1024 bit ephemeral RSA key for the RSA_EXPORT handshake method. This message is always fatal."

        Case 70



            ErrorCode = "protocol_version - The protocol version the client attempted to negotiate is recognized, but not supported. For example, old protocol versions might be avoided for security reasons. This message is always fatal."

        Case 71



            ErrorCode = "insufficient_security - Failed negotiation specifically because the server requires ciphers more secure than those supported by the client. Returned instead of handshake_failure. This message is always fatal."

        Case 80



            ErrorCode = "internal_error - An internal error unrelated to the peer or the correctness of the protocol makes it impossible to continue, such as a memory allocation failure. The error is not related to protocol. This message is always fatal."

        Case 90



            ErrorCode = "user_cancelled - Cancelled handshake for a reason that is unrelated to a protocol failure. If the user cancels an operation after the handshake is complete, just closing the connection by sending a close_notify is more appropriate. This alert should be followed by a close_notify. This message is generally a warning."

        Case 100



            ErrorCode = "no_renegotiation - Sent by the client in response to a hello request or sent by the server in response to a client hello after initial handshaking. Either of these would normally lead to renegotiation; when that is not appropriate, the recipient should respond with this alert; at that point, the original requester can decide whether to proceed with the connection. One case where this would be appropriate would be where a server has spawned a process to satisfy a request; the process might receive security parameters (key length, authentication, and so on) at start-up and it might be difficult to communicate changes to these parameters after that point. This message is always a warning."

        Case 255



            ErrorCode = "unsupported_extension"
        Case Else
            ErrorCode = "Unkown error code number " & Code

    End Select


End Function
'

