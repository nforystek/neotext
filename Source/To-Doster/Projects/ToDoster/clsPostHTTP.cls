VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPostHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text

'##########################################################################################################################
'##########################################################################################################################
'This Class Object is strictly used to post data to a web form
'##########################################################################################################################
'##########################################################################################################################



    Private Declare Function InternetOpen Lib "wininet" _
         Alias "InternetOpenA" _
            (ByVal lpszCallerName As String, _
             ByVal dwAccessType As Long, _
             ByVal lpszProxyName As String, _
             ByVal lpszProxyBypass As String, _
             ByVal dwFlags As Long) As Long

      Private Declare Function InternetConnect Lib "wininet" _
            Alias "InternetConnectA" _
            (ByVal hInternetSession As Long, _
             ByVal lpszServerName As String, _
             ByVal nProxyPort As Integer, _
             ByVal lpszUsername As String, _
             ByVal lpszPassword As String, _
             ByVal dwService As Long, _
             ByVal dwFlags As Long, _
             ByVal dwContext As Long) As Long

   Private Declare Function InternetReadFile Lib "wininet" _
            (ByVal hFile As Long, _
             ByVal sBuffer As String, _
             ByVal lNumBytesToRead As Long, _
             lNumberOfBytesRead As Long) As Integer

   Private Declare Function HttpOpenRequest Lib "wininet" _
            Alias "HttpOpenRequestA" _
            (ByVal hInternetSession As Long, _
             ByVal lpszVerb As String, _
             ByVal lpszObjectName As String, _
             ByVal lpszVersion As String, _
             ByVal lpszReferer As String, _
             ByVal lpszAcceptTypes As Long, _
             ByVal dwFlags As Long, _
             ByVal dwContext As Long) As Long

   Private Declare Function HttpSendRequest Lib "wininet" _
            Alias "HttpSendRequestA" _
            (ByVal hHttpRequest As Long, _
             ByVal sHeaders As String, _
             ByVal lHeadersLength As Long, _
             ByVal sOptional As String, _
             ByVal lOptionalLength As Long) As Boolean

   Private Declare Function InternetCloseHandle Lib "wininet" _
            (ByVal hInternetHandle As Long) As Boolean

   Private Declare Function HttpAddRequestHeaders Lib "wininet" _
             Alias "HttpAddRequestHeadersA" _
             (ByVal hHttpRequest As Long, _
             ByVal sHeaders As String, _
             ByVal lHeadersLength As Long, _
             ByVal lModifiers As Long) As Integer


Public Function PostFormData(ByVal ServerName As String, ByVal WebFile As String, ByVal postData As String) As String
            Dim hInternetOpen As Long
            Dim hInternetConnect As Long
            Dim hHttpOpenRequest As Long
            Dim bRet As Boolean

            hInternetOpen = 0
            hInternetConnect = 0
            hHttpOpenRequest = 0

            'Use registry access settings.
            Const INTERNET_OPEN_TYPE_PRECONFIG = 0
            hInternetOpen = InternetOpen("ToDoster", _
                            INTERNET_OPEN_TYPE_PRECONFIG, _
                            vbNullString, _
                            vbNullString, _
                            0)

            If hInternetOpen <> 0 Then
               'Type of service to access.
               Const INTERNET_SERVICE_HTTP = 3
               Const INTERNET_DEFAULT_HTTP_PORT = 80
               'Change the server to your server name
               hInternetConnect = InternetConnect(hInternetOpen, _
                                  ServerName, _
                                  INTERNET_DEFAULT_HTTP_PORT, _
                                  vbNullString, _
                                  "HTTP/1.0", _
                                  INTERNET_SERVICE_HTTP, _
                                  0, _
                                  0)

               If hInternetConnect <> 0 Then
                'Brings the data across the wire even if it locally cached.
                 Const INTERNET_FLAG_RELOAD = &H80000000
                 hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                                     "POST", _
                                     WebFile, _
                                     "HTTP/1.0", _
                                     vbNullString, _
                                     0, _
                                     INTERNET_FLAG_RELOAD, _
                                     0)

                  If hHttpOpenRequest <> 0 Then
                     Dim sHeader As String
                     Const HTTP_ADDREQ_FLAG_ADD = &H20000000
                     Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
                    sHeader = "Content-Type: application/x-www-form-urlencoded" _
                       & vbCrLf
                     bRet = HttpAddRequestHeaders(hHttpOpenRequest, _
                       sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE _
                       Or HTTP_ADDREQ_FLAG_ADD)

                     Dim lpszPostData As String
                     Dim lPostDataLen As Long

                     lpszPostData = postData
                     lPostDataLen = Len(lpszPostData)
                     bRet = HttpSendRequest(hHttpOpenRequest, _
                            vbNullString, _
                            0, _
                            lpszPostData, _
                            lPostDataLen)

                     Dim bDoLoop             As Boolean
                     Dim sReadBuffer         As String * 2048
                     Dim lNumberOfBytesRead  As Long
                     Dim sBuffer             As String
                     bDoLoop = True
                     While bDoLoop
                      sReadBuffer = vbNullString
                      bDoLoop = InternetReadFile(hHttpOpenRequest, _
                         sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                      sBuffer = sBuffer & _
                           Left(sReadBuffer, lNumberOfBytesRead)
                      If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                     Wend
                     
                     bRet = InternetCloseHandle(hHttpOpenRequest)
                  End If
                  bRet = InternetCloseHandle(hInternetConnect)
               End If
               bRet = InternetCloseHandle(hInternetOpen)
            End If

            
            If Left(LCase(sBuffer), 5) = "error" Then
                sBuffer = Trim(Mid(sBuffer, 6))
                Dim errNum As Long
                Dim errSource As String
                errNum = CLng(Trim(Left(sBuffer, InStr(sBuffer, " ") - 1)))
                sBuffer = Trim(Mid(sBuffer, InStr(sBuffer, " ") + 1))
                errSource = Trim(Left(sBuffer, InStr(sBuffer, " ") - 1))
                sBuffer = Trim(Mid(sBuffer, InStr(sBuffer, " ") + 1))
                Err.Raise errNum, errSource, sBuffer
            Else
                PostFormData = sBuffer
            End If


End Function


