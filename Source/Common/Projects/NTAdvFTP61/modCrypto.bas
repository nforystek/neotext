Attribute VB_Name = "modCrypto"
Option Explicit

'Public Enum CipherType
'    Encrypt
'    Decrypt
'End Enum
'
'Public Type PROV_ENUMALGS
'    aiAlgid As Long
'    dwBitLen As Long
'    dwNameLen As Long
'    szName As String * 20
'End Type
'
'Public Type PROV_ENUMALGS_EX
'    aiAlgid As Long
'    dwDefaultLen As Long
'    dwMinLen As Long
'    dwMaxLen As Long
'    dwProtocols As Long
'    dwNameLen As Long
'    szName As String * 20
'    dwLongNameLen As Long
'    szLongName As String * 40
'End Type
'
''  CryptoAPI Methods.
''Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptEnumProviders Lib "advapi32.dll" Alias "CryptEnumProvidersA" (ByVal dwIndex As Long, ByVal pdwReserved As Long, ByVal dwFlags As Long, pdwProvType As Long, ByVal pszProvName As String, pcbProvName As Long) As Long
''Public Declare Function CryptGetDefaultProvider Lib "advapi32.dll" (ByVal dwProvType As Long, pdwReserved As Long, ByVal dwFlags As Long, ByVal pszProvName As String, pcbProvName As Long) As Long
''
''Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
''Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
''Public Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
''Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
''Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
''Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
''Public Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
''Public Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal AlgId As Long, ByVal dwFlags As Long, phKey As Long) As Long
''Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pcbData As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptGetKeyParam Lib "advapi32.dll" (ByVal hCryptKey As Long, ByVal dwParam As Long, pbData As Any, pcbData As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
''Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
''Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
''Public Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
''Public Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
''Public Declare Function CryptSetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
''
''Public Declare Function CryptEncryptPtr Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
''Public Declare Function CryptDecryptPtr Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, pdwDataLen As Long) As Long
''Public Declare Function CryptDuplicateHash Lib "advapi32.dll" (ByVal hHash As Long, pdwReserved As Long, ByVal dwFlags As Long, phHash As Long) As Long
''Public Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
''
''Public Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, pbSignature As Any, pdwSigLen As Long) As Long
''Public Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, pbSignature As Any, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
''Public Declare Function CryptDecryptFile Lib "advapi32.dll" Alias "DecryptFileA" (ByVal lpFileName As String, ByVal dwReserved As Long) As Long
''Public Declare Function CryptEncryptFile Lib "advapi32.dll" Alias "EncryptFileA" (ByVal lpFileName As String) As Long
'
''  Constants
''  error codes.
'Public Const NTE_BAD_ALGID As Long = &H80090008
'Public Const NTE_BAD_DATA As Long = &H80090005
'Public Const NTE_BAD_FLAGS As Long = &H80090009
'Public Const NTE_BAD_HASH As Long = &H80090002
'Public Const NTE_BAD_HASH_STATE As Long = &H8009000C
'Public Const NTE_BAD_KEY As Long = &H80090003
'Public Const NTE_BAD_KEYSET As Long = &H80090016
'Public Const NTE_BAD_KEYSET_PARAM As Long = &H8009001F
'Public Const NTE_BAD_KEY_STATE As Long = &H8009000B
'Public Const NTE_BAD_LEN As Long = &H80090004
'Public Const NTE_BAD_PROVIDER As Long = &H80090013
'Public Const NTE_BAD_PROV_TYPE As Long = &H80090014
'Public Const NTE_BAD_PUBLIC_KEY As Long = &H80090015
'Public Const NTE_BAD_SIGNATURE As Long = &H80090006
'Public Const NTE_BAD_TYPE As Long = &H8009000A
'Public Const NTE_BAD_UID As Long = &H80090001
'Public Const NTE_BAD_VER As Long = &H80090007
'Public Const NTE_DOUBLE_ENCRYPT As Long = &H80090012
'Public Const NTE_EXISTS As Long = &H8009000F
'Public Const NTE_FAIL As Long = &H80090020
'Public Const NTE_KEYSET_ENTRY_BAD As Long = &H8009001A
'Public Const NTE_KEYSET_NOT_DEF As Long = &H80090019
'Public Const NTE_NO_KEY As Long = &H8009000D
'Public Const NTE_NO_MEMORY As Long = &H8009000E
'Public Const NTE_NOT_FOUND As Long = &H80090011
'Public Const NTE_PERM As Long = &H80090010
'Public Const NTE_PROVIDER_DLL_FAIL As Long = &H8009001D
'Public Const NTE_PROV_DLL_NOT_FOUND As Long = &H8009001E
'Public Const NTE_PROV_TYPE_ENTRY_BAD As Long = &H80090018
'Public Const NTE_PROV_TYPE_NO_MATCH As Long = &H8009001B
'Public Const NTE_PROV_TYPE_NOT_DEF As Long = &H80090017
'Public Const NTE_SIGNATURE_FILE_BAD As Long = &H8009001C
'Public Const NTE_SYS_ERR As Long = &H80090021
'
''  CryptoAPI Provider constants.
'Public Const PROV_DH_SCHANNEL As Long = 18
'Public Const PROV_DSS As Long = &H3
'Public Const PROV_DSS_DH As Long = &HD
'Public Const PROV_EC_ECDSA_FULL As Long = &H10
'Public Const PROV_EC_ECDSA_SIG As Long = &HE
'Public Const PROV_EC_ECNRA_FULL As Long = &H11
'Public Const PROV_EC_ECNRA_SIG As Long = &HF
'Public Const PROV_FORTEZZA As Long = &H4
'Public Const PROV_MS_EXCHANGE As Long = &H5
'Public Const PROV_RSA_FULL As Long = &H1
'Public Const PROV_RSA_SCHANNEL As Long = &HC
'Public Const PROV_RSA_SIG As Long = &H2
'Public Const PROV_SPYRUS_LYNKS As Long = &H14
'Public Const PROV_SSL As Long = &H6
'Public Const PROV_STT_ACQ As Long = &H8
'Public Const PROV_STT_BRND As Long = &H9
'Public Const PROV_STT_ISS As Long = &HB
'Public Const PROV_STT_MER As Long = &H7
'Public Const PROV_STT_ROOT As Long = &HA
'
'
''  CryptoAPI Context constants.
'Public Const CRYPT_DELETEKEYSET As Long = &H10
'Public Const CRYPT_MACHINE_KEYSET As Long = &H20
'Public Const CRYPT_NEWKEYSET As Long = &H8
'Public Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
'
''  CryptoAPI Provider Get constants.
'Public Const PP_APPLI_CERT As Long = &H12
'Public Const PP_CERTCHAIN As Long = &H9
'Public Const PP_CHANGE_PASSWORD As Long = &H7
'Public Const PP_CONTAINER As Long = &H6
'Public Const PP_ENUMALGS As Long = &H1
'Public Const PP_ENUMALGS_EX As Long = &H16
'Public Const PP_ENUMCONTAINERS As Long = &H2
'Public Const PP_IMPTYPE As Long = &H3
'Public Const PP_KEYSET_SEC_DESCR As Long = &H8
'Public Const PP_KEYSTORAGE As Long = &H11
'Public Const PP_KEY_TYPE_SUBTYPE As Long = &HA
'Public Const PP_NAME As Long = &H4
'Public Const PP_PROVTYPE As Long = &H10
'Public Const PP_SESSION_KEYSIZE As Long = &H14
'Public Const PP_SYM_KEYSIZE As Long = &H13
'Public Const PP_UI_PROMPT As Long = &H15
'Public Const PP_VERSION As Long = &H5
'Public Const PP_SIG_KEYSIZE_INC As Long = 34
'Public Const PP_KEYX_KEYSIZE_INC As Long = 35
'Public Const PP_UNIQUE_CONTAINER As Long = 36
'Public Const PP_USE_HARDWARE_RNG As Long = 38
'
''  CryptoAPI Provider Get flags constants.
'Public Const CRYPT_FIRST As Long = &H1
'Public Const CRYPT_FLAG_PCT1 As Long = &H1
'Public Const CRYPT_FLAG_SSL2 As Long = &H2
'Public Const CRYPT_FLAG_SSL3 As Long = &H4
'Public Const CRYPT_FLAG_TLS1 As Long = &H8
'Public Const CRYPT_NEXT As Long = &H2
'Public Const CRYPT_PSTORE As Long = &H2
'Public Const CRYPT_SEC_DESCR As Long = &H1
'Public Const CRYPT_UI_PROMPT As Long = &H4
'
''  CryptoAPI Provider Get flags impl. constants.
'Public Const CRYPT_IMPL_HARDWARE As Long = &H1
'Public Const CRYPT_IMPL_MIXED As Long = &H3
'Public Const CRYPT_IMPL_SOFTWARE As Long = &H2
'Public Const CRYPT_IMPL_UNKNOWN As Long = &H4
'
''  CryptoAPI Provider Set constants.
'Public Const PP_CLIENT_HWND As Long = &H1
'Public Const PP_CONTEXT_INFO As Long = &HB
'Public Const PP_DELETEKEY As Long = &H18
'Public Const PP_KEYEXCHANGE_ALG As Long = &HE
'Public Const PP_KEYEXCHANGE_KEYSIZE As Long = &HC
'Public Const PP_SIGNATURE_ALG As Long = &HF
'Public Const PP_SIGNATURE_KEYSIZE As Long = &HD
'
''  CryptoAPI Key flag constants.
'Public Const CRYPT_CREATE_SALT As Long = &H4
'Public Const CRYPT_CREATE_IV As Long = &H200
'Public Const CRYPT_DATA_KEY As Long = &H800
'Public Const CRYPT_EXPORTABLE As Long = &H1
'Public Const CRYPT_INITIATOR As Long = &H40
'Public Const CRYPT_KEK As Long = &H400
'Public Const CRYPT_NO_SALT As Long = &H10
'Public Const CRYPT_ONLINE As Long = &H80
'Public Const CRYPT_PREGEN As Long = &H40
'Public Const CRYPT_RECIPIENT As Long = &H10
'Public Const CRYPT_SERVER As Long = &H400
'Public Const CRYPT_SF As Long = &H100
'Public Const CRYPT_UPDATE_KEY As Long = &H8
'Public Const CRYPT_USER_PROTECTED As Long = &H2
'
''  CryptoAPI public/private key type constants.
'Public Const AT_KEYEXCHANGE As Long = &H1
'Public Const AT_SIGNATURE As Long = &H2
'
''  CryptoAPI algorithm classes constants.
'Public Const ALG_CLASS_ANY As Long = 0
'Public Const ALG_CLASS_SIGNATURE As Long = 8192
'Public Const ALG_CLASS_MSG_ENCRYPT As Long = 16384
'Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
'Public Const ALG_CLASS_HASH As Long = 32768
'Public Const ALG_CLASS_KEY_EXCHANGE As Long = 40960
'Public Const ALG_CLASS_ALL As Long = 57344
'
''  CryptoAPI algorithm type constants.
'Public Const ALG_TYPE_ANY As Long = 0
'Public Const ALG_TYPE_BLOCK As Long = 1536
'Public Const ALG_TYPE_RSA As Long = 1024
'Public Const ALG_TYPE_STREAM As Long = 2048
'
''  CryptoAPI algorithm SID constants.
'Public Const ALG_SID_DES As Long = 1
'Public Const ALG_SID_MD5 As Long = 3
'Public Const ALG_SID_RC2 As Long = 2
'Public Const ALG_SID_RC4 As Long = 1
'Public Const ALG_SID_RSA_ANY As Long = 0
'Public Const ALG_SID_SHA As Long = 4
'
''  CryptoAPI algorithm constants.
'Public Const CALG_3DES As Long = &H6603
'Public Const CALG_3DES_112 As Long = &H6609
'Public Const CALG_AGREEDKEY_ANY As Long = &HAA03
'Public Const CALG_CYLINK_MEK As Long = &H660C
'Public Const CALG_DES As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
'Public Const CALG_DH_EPHEM As Long = &HAA02
'Public Const CALG_DH_SF As Long = &HAA01
'Public Const CALG_DSS_SIGN As Long = &H2200
'Public Const CALG_HMAC As Long = &H8009
'Public Const CALG_HUGHES_MD5 As Long = &HA003
'Public Const CALG_KEA_KEYX As Long = &HAA04
'Public Const CALG_MAC As Long = &H8006
'Public Const CALG_MD2 As Long = &H8001
'Public Const CALG_MD4 As Long = &H8002
'Public Const CALG_MD5 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
'Public Const CALG_PCT1_MASTER As Long = &H4C04
'Public Const CALG_RC2 As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2)
'Public Const CALG_RC4 As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
'Public Const CALG_RC5 As Long = &H660D
'Public Const CALG_RSA_KEYX As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
'Public Const CALG_RSA_SIGN As Long = &H2400
'Public Const CALG_SCHANNEL_ENC_KEY As Long = &H4C07
'Public Const CALG_SCHANNEL_MAC_KEY As Long = &H4C03
'Public Const CALG_SCHANNEL_MASTER_HASH As Long = &H4C02
'Public Const CALG_SEAL As Long = &H6802
'Public Const CALG_SHA As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
'Public Const CALG_SHA1 As Long = &H8005
'Public Const CALG_SKIPJACK As Long = &H660A
'Public Const CALG_SSL2_MASTER As Long = &H4C05
'Public Const CALG_SSL3_MASTER As Long = &H4C01
'Public Const CALG_SSL3_SHAMD5 As Long = &H8008
'Public Const CALG_TEK As Long = &H660B
'Public Const CALG_TLS1_MASTER As Long = &H4C06
'
''  CryptoAPI Key parameter constants.
'Public Const KP_ALGID As Long = &H7
'Public Const KP_CERTIFICATE As Long = &H1A
'Public Const KP_CLEAR_KEY As Long = &H1B
'Public Const KP_CLIENT_RANDOM As Long = &H15
'Public Const KP_BLOCKLEN As Long = &H8
'Public Const KP_EFFECTIVE_KEYLEN As Long = &H13
'Public Const KP_G As Long = &HC
'Public Const KP_INFO As Long = &H12
'Public Const KP_IV As Long = &H1
'Public Const KP_KEYLEN As Long = &H9
'Public Const KP_MODE As Long = &H4
'Public Const KP_MODE_BITS As Long = &H5
'Public Const KP_P As Long = &HB
'Public Const KP_PADDING As Long = &H3
'Public Const KP_PERMISSIONS As Long = &H6
'Public Const KP_PRECOMP_MD5 As Long = &H18
'Public Const KP_PRECOMP_SHA As Long = &H19
'Public Const KP_PUB_EX_LEN As Long = &H1C
'Public Const KP_PUB_EX_VAL As Long = &H1D
'Public Const KP_Q As Long = &HD
'Public Const KP_RA As Long = &H10
'Public Const KP_RB As Long = &H11
'Public Const KP_RP As Long = &H17
'Public Const KP_SALT As Long = &H2
'Public Const KP_SALT_EX As Long = &HA
'Public Const KP_SCHANNEL_ALG As Long = &H14
'Public Const KP_SERVER_RANDOM As Long = &H16
'Public Const KP_X As Long = &HE
'Public Const KP_Y As Long = &HF
'
''  CryptoAPI Padding constants.
'Public Const PKCS5_PADDING As Long = &H1
'Public Const RANDOM_PADDING As Long = &H2
'Public Const ZERO_PADDING As Long = &H3
'
''  CryptoAPI mode constants.
'Public Const CRYPT_MODE_CBC As Long = &H1
'Public Const CRYPT_MODE_CFB As Long = &H4
'Public Const CRYPT_MODE_CTS As Long = &H5
'Public Const CRYPT_MODE_ECB As Long = &H2
'Public Const CRYPT_MODE_OFB As Long = &H3
'
''  CryptoAPI permission constants.
'Public Const CRYPT_DECRYPT As Long = &H2
'Public Const CRYPT_ENCRYPT As Long = &H1
'Public Const CRYPT_EXPORT As Long = &H4
'Public Const CRYPT_EXPORT_KEY As Long = &H40
'Public Const CRYPT_IMPORT_KEY As Long = &H80
'Public Const CRYPT_MAC As Long = &H20
'Public Const CRYPT_READ As Long = &H8
'Public Const CRYPT_WRITE As Long = &H10
'
''  CryptoAPI blob constants.
'Public Const OPAQUEKEYBLOB As Long = &H8
'Public Const PRIVATEKEYBLOB As Long = &H7
'Public Const PUBLICKEYBLOB As Long = &H6
'Public Const SIMPLEBLOB As Long = &H1
'Public Const SYMMETRICWRAPKEYBLOB As Long = &HB
'
''  CryptoAPI encoding constants.
'Public Const CRYPT_BASE64_ENCODING As Long = &H4
'Public Const CRYPT_HEX_ENCODING As Long = &H1
'Public Const CRYPT_NO_ENCODING As Long = &H0
'Public Const CRYPT_URL_ENCODING As Long = &H2
'Public Const CRYPT_UU_ENCODING As Long = &H3
'
''  CryptoAPI hash data constants.
'Public Const CRYPT_USERDATA As Long = &H1
'
''  CryptoAPI provider constants.
'Public Const CRYPT_MACHINE_DEFAULT As Long = &H1
'Public Const CRYPT_USER_DEFAULT As Long = &H2
'Public Const CRYPT_DELETE_DEFAULT As Long = &H4
'
''  CryptoAPI hash parameter constants.
'Public Const HP_ALGID As Long = &H1
'Public Const HP_HASHVAL As Long = &H2
'Public Const HP_HASHSIZE As Long = &H4
'Public Const HP_HMAC_INFO As Long = &H5
'
''  CryptoAPI provider-friendly names.
'Public Const MS_DEF_DH_SCHANNEL_PROV As String = "Microsoft DH SChannel Cryptographic Provider"
'Public Const MS_DEF_DSS_DH_PROV As String = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
'Public Const MS_DEF_DSS_PROV As String = "Microsoft Base DSS Cryptographic Provider"
'Public Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
'Public Const MS_DEF_RSA_SCHANNEL_PROV As String = "Microsoft RSA SChannel Cryptographic Provider"
'Public Const MS_DEF_RSA_SIG_PROV As String = "Microsoft RSA Signature Cryptographic Provider"
'Public Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
'Public Const MS_ENH_DSS_DH_PROV As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
'Public Const MS_SCARD_PROV As String = "Microsoft Base Smart Card Crypto Provider"
'Public Const MS_STRONG_PROV As String = "Microsoft Strong Cryptographic Provider"
'
''  Miscellaneous CryptoAPI constants.
'Public Const CRYPT_OAEP As Long = &H40
'Public Const MAXUIDLEN As Long = 64
'Public Const CSP_REGISTRY_KEY As String = "SOFTWARE\Microsoft\Cryptography\Defaults\Provider"
'
''  SECURITY_DESCRIPTOR constants
'Public Const OWNER_SECURITY_INFORMATION  As Long = &H1
'Public Const GROUP_SECURITY_INFORMATION As Long = &H2
'Public Const DACL_SECURITY_INFORMATION As Long = &H4
'Public Const SACL_SECURITY_INFORMATION As Long = &H8
'
'Public Const ERROR_MORE_DATA As Long = 234
'Public Const ERROR_NO_MORE_ITEMS As Long = 259

Global KeyContainer As String

Global CertObjs As NTNodes10.Collection

Public Sub ViewCertificate(ByRef cert As Certificate)

    If CertObjs Is Nothing Then
        Set CertObjs = New NTNodes10.Collection
    End If
    
    Dim CertForm As New frmCert
    CertForm.ViewCert cert
    Unload CertForm
    Set CertForm = Nothing
End Sub

Public Function CheckCertificate(ByRef cert As Certificate) As Boolean
    If CertObjs Is Nothing Then
        Set CertObjs = New NTNodes10.Collection
    End If
    Dim sn As String
    
    sn = "SN_" & Replace(cert.Terms(SerialNumber), " ", "")
    
    If CertObjs.Exists(sn) Then
        Set cert = CertObjs(sn)
        CertObjs.Remove sn
    End If
    
    If (Not cert.NoPrompt) Then 'And (LCase(AppEXE(True, True)) <> "maxservice") Then
        Dim CertForm As New frmCert
        CertForm.CheckCert cert
        Unload CertForm
        Set CertForm = Nothing
    Else 'If (LCase(AppEXE(True, True)) = "maxservice") Then
        cert.Accepted = True
    End If

    CheckCertificate = cert.Accepted
    
    If Not CertObjs Is Nothing Then
        CertObjs.Add cert, sn
    End If
End Function

Public Sub TermCerts()
    If Not CertObjs Is Nothing Then
        CertObjs.Clear
        Set CertObjs = Nothing
    End If
End Sub

