VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Internet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get LocalHost() As String
    LocalHost = modResolve.LocalHost
End Property

Public Function ResolveHost(Optional ByVal HostName As String = "") As String
    ResolveHost = modResolve.ResolveHost(IIf(HostName = "", LocalHost, HostName))
End Function

Public Function URLDecode(ByRef encodedString As String) As String
    URLDecode = modInternet.URLDecode(encodedString)
End Function

Public Function URLEncode(ByRef encodeString As String) As String
    URLEncode = modInternet.URLEncode(encodeString)
End Function

Public Function PostToWebsite(ByVal HostServerAddress As String, Optional ByVal WebFilePath As String = "/", Optional ByVal PostFormData As String = "", Optional ByVal AuthUsername As String = "", Optional ByVal AuthPassword As String = "", Optional ByVal CacheAndCookies As Boolean = False, Optional ByVal SecureSocketLayer As Boolean = False) As String
    PostToWebsite = modInternet.PostToWebsite(HostServerAddress, WebFilePath, PostFormData, AuthUsername, AuthPassword, CacheAndCookies, SecureSocketLayer)
End Function

Public Function OpenWebsite(ByVal WebLink As String) As Boolean
    OpenWebsite = modProcess.RunFile(WebLink)
End Function

