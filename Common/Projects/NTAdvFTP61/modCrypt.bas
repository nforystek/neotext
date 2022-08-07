






Attribute VB_Name = "modCryptO"
Option Explicit

Option Compare Binary

Option Private Module
'/**  modCryptoAPI
' *     this module is for constants and usage of the cryptographic providers used for encrypt decrypts
' *     treamswithasychronouskeys
' *
' *     @project     cXPLib
' *     @author      Philipp Rothmann,  mailto: dev@preneco.de
' *     @copyright   © 2000-2003 .) development
' *
' *     This program is free software; you can redistribute it and/or modify
' *     it under the terms of the GNU General Public License as published by
' *     the Free Software Foundation; either version 2 of the License, or
' *     (at your option) any later version.
' *
' *     This program is distributed in the hope that it will be useful,
' *     but WITHOUT ANY WARRANTY; without even the implied warranty of
' *     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' *     GNU General Public License for more details.
' *
' *     You should have received a copy of the GNU General Public License
' *     along with this program; if not, write to the Free Software
' *     Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
' *
' *
' *     @lastupdate  05.10.2002 00:38:34
' */

'' Functions to handle Provider

Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
                        phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
                        ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    
Public Declare Function CryptReleaseContext Lib "advapi32.dll" _
                        (ByVal hProv As Long, _
                        ByVal dwFlags As Long) As Long

'' Functions to handle and operate with Hash objects

Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
                        ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" ( _
                        ByVal hHash As Long, ByVal pbdata As String, _
                        ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptHashSessionKey Lib "advapi32.dll" ( _
                        ByVal hHash As Long, ByVal phKey As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" _
                        (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbdata As String, _
                        pdwDataLen As Long, ByVal dwFlags As Long) As Long

'' Key Functions

Public Declare Function CryptGenKey Lib "advapi32.dll" ( _
                        ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
Public Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, _
                        ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, _
                        ByVal pbdata As String, pdwDataLen As Long) As Long
Public Declare Function CryptImportKey Lib "advapi32.dll" ( _
                        ByVal hProv As Long, ByVal pbdata As String, _
                        ByVal dwDataLen As Long, ByVal hImpKey As Long, ByVal dwFlags As Long, _
                        phKey As Long) As Long
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'' Encrypting/Decrypting Functions

Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, _
                        ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbdata As Any, _
                        pdwDataLen As Long, ByVal dwBufLen As Long) As Long
''Variation for string Data blocks
Public Declare Function CryptStringEncrypt Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hKey As Long, _
                        ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbdata As String, _
                        pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, _
                        ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbdata As Any, _
                        pdwDataLen As Long) As Long
''Variation for string Data blocks
Public Declare Function CryptStringDecrypt Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hKey As Long, _
                        ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbdata As String, _
                        pdwDataLen As Long) As Long

'' Name of the provider shipped with Windows by default
'' Both support below declared hashing algorithms.
'' MS_DEF_PROV is bundled with the operating system.

Public Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL As Long = 1
Public Const CRYPT_NEWKEYSET As Long = &H8
Public Const CRYPT_EXPORTABLE As Long = &H1
Public Const AT_KEYEXCHANGE As Long = 1

''CryptGetHashParam parameter number values
Public Const HP_ALGID As Long = 1
Public Const HP_HASHVAL As Long = 2
Public Const HP_HASHSIZE As Long = 4
''
'' Exported key blob definitions
Public Const SIMPLEBLOB As Long = 1
Public Const PUBLICKEYBLOB As Long = 6
Public Const PRIVATEKEYBLOB As Long = 7
Public Const PLAINTEXTKEYBLOB As Long = 8
''Algorithm classes
Public Const ALG_CLASS_SIGNATURE As Long = 8192
Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Public Const ALG_CLASS_HASH As Long = 32768
''Algorithm types
Public Const ALG_TYPE_ANY As Long = 0
Public Const ALG_TYPE_BLOCK As Long = 1536
Public Const ALG_TYPE_STREAM As Long = 2048
''Block cipher sub ids
Public Const ALG_SID_DES As Long = 1
Public Const ALG_SID_3DES As Long = 3
Public Const ALG_SID_3DES_112 As Long = 9
Public Const ALG_SID_RC2 As Long = 2
''Stream cipher sub-ids
Public Const ALG_SID_RC4 As Long = 1
''Hash sub ids
Public Const ALG_SID_MD2 As Long = 1
Public Const ALG_SID_MD4 As Long = 2
Public Const ALG_SID_MD5 As Long = 3
Public Const ALG_SID_SHA As Long = 4
'' Algorithm identifier definitions
'' Hashing algorithms
Public Const CALG_MD2 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Public Const CALG_MD4 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Public Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Public Const CALG_SHA As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
'' Encryption/Decryption algorithms
'' Block ciphers
Public Const CALG_DES As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
Public Const CALG_3DES As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES)
Public Const CALG_3DES_112 As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES_112)
Public Const CALG_BLOCKSIZE As Byte = 16
Public Const CALG_RC2 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
'' Stream ciphers
Public Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)


Private lCryptInstances As Long, hProvider As Long

Public Function ReleaseCryptProvider() As Boolean
    
    lCryptInstances = lCryptInstances - 1
    
    If lCryptInstances = 0 Then
        CryptReleaseContext hProvider, 0
        hProvider = 0
    End If
    
End Function

'/**  InitCryptProvider
' *     initializes the cryptographic provider for each class
' *
' *     @returns     Long , the provider handle
' */
Public Function InitCryptProvider(Optional Container As String = vbNullString) As Long

  Dim hProv As Long
  Dim sProvider As String         ' Name of provider

    On Error Resume Next
    
      If hProvider = 0 Then
            sProvider = MS_ENHANCED_PROV & vbNullChar
            'Attempt to acquire a handle to the chosen key container.
            If Not CBool(CryptAcquireContext(hProv, Container, ByVal sProvider, PROV_RSA_FULL, 0)) Then
                ' Attempt to create a new key container
                If Not CBool(CryptAcquireContext(hProv, Container, ByVal sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
                    ' Attempt to get a handle to the enhanced key container
                    sProvider = MS_DEF_PROV & vbNullChar
                    If Not CBool(CryptAcquireContext(hProv, Container, ByVal sProvider, PROV_RSA_FULL, 0)) Then
                        ' Attempt to create a new key container
                        If Not CBool(CryptAcquireContext(hProv, Container, ByVal sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
                            InitCryptProvider = -1
                        End If
                    End If
                End If
            End If
            'Log.out "cCrypto.InitCryptProvider", sProvider
      hProvider = hProv
    End If
      lCryptInstances = lCryptInstances + 1
      InitCryptProvider = hProvider
    On Error GoTo 0

End Function









