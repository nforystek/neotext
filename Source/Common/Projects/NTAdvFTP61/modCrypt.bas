Attribute VB_Name = "modCrypt"
Option Explicit

' --- TLS Context Management ---
Public Declare Function TlsInit Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal serverName As String, ByRef pErr As Long) As Long

Public Declare Sub TlsClose Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long)

' --- Handshake ---
Public Declare Function TlsHandshake Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long, _
     ByVal inBuf As Long, ByVal inLen As Long, _
     ByVal outBuf As Long, ByVal outSize As Long, _
     ByRef pErr As Long) As Long

Public Declare Function TlsIsHandshakeComplete Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long) As Long

' --- Application Data Encryption / Decryption ---
Public Declare Function TlsSend Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long, _
     ByVal plain As Long, ByVal plainLen As Long, _
     ByVal outBuf As Long, ByVal outSize As Long, _
     ByRef pErr As Long) As Long

Public Declare Function TlsRecv Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long, _
     ByVal enc As Long, ByVal encLen As Long, _
     ByVal outBuf As Long, ByVal outSize As Long, _
     ByRef pErr As Long) As Long

' --- Cipher Info ---
Public Declare Function TlsGetCipherInfo Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long, _
     ByVal outBuf As String, ByVal outSize As Long) As Long

' --- Renegotiation ---
Public Declare Function TlsRenegotiate Lib "C:\Development\Neotext\Common\Binary\NTSecureTSL.dll" _
    (ByVal ctx As Long) As Long


Public Declare Function TlsGetPeerCert Lib "TlsWrapper.dll" _
    (ByVal ctx As Long, _
     ByVal outBuf As Long, ByVal outSize As Long, _
     ByRef pErr As Long) As Long

Public Declare Function TlsValidateCert Lib "TlsWrapper.dll" _
    (ByVal ctx As Long, _
     ByVal serverName As String, _
     ByRef pErr As Long) As Long

Public Declare Function TlsGetProtocolVersion Lib "TlsWrapper.dll" _
    (ByVal ctx As Long) As Long

Public Declare Function TlsGetPeerCertInfo Lib "TlsWrapper.dll" _
    (ByVal ctx As Long, _
     ByVal outBuf As String, ByVal outSize As Long, _
     ByRef pErr As Long) As Long




Public Function BytesFromString(ByVal s As String) As Byte()
    Dim b() As Byte
    If LenB(s) = 0 Then
        ReDim b(0 To -1)
    Else
        b = StrConv(s, vbFromUnicode)
    End If
    BytesFromString = b
End Function

Public Function StringFromBytes(b() As Byte) As String
    If (Not Not b) = 0 Then
        StringFromBytes = ""
    Else
        StringFromBytes = StrConv(b, vbUnicode)
    End If
End Function



