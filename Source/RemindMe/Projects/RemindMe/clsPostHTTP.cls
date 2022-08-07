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

Option Compare Binary

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

            Const INTERNET_OPEN_TYPE_PRECONFIG = 0
            hInternetOpen = InternetOpen(App.EXEName, _
                            INTERNET_OPEN_TYPE_PRECONFIG, _
                            vbNullString, _
                            vbNullString, _
                            0)
            
            If hInternetOpen <> 0 Then

               Const INTERNET_SERVICE_HTTP = 3
               Const INTERNET_DEFAULT_HTTP_PORT = 80

               hInternetConnect = InternetConnect(hInternetOpen, _
                                  ServerName, _
                                  INTERNET_DEFAULT_HTTP_PORT, _
                                  vbNullString, _
                                  "HTTP/1.0", _
                                  INTERNET_SERVICE_HTTP, _
                                  0, _
                                  0)

               If hInternetConnect <> 0 Then

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

            PostFormData = sBuffer

End Function

Public Function URLUnEncode(encodedString As String) As String

    Dim ReturnString As String
    Dim currentChar As String

    Dim i As Long
    i = 1

    Do Until i > Len(encodedString)
        currentChar = Mid(encodedString, i, 1)

        If currentChar = "+" Then
            ReturnString = ReturnString + " "
            i = i + 1
        ElseIf currentChar = "%" Then
            currentChar = Mid(encodedString, i + 1, 2)
            ReturnString = ReturnString + Chr(val("&H" & currentChar))
            i = i + 3
        Else
            ReturnString = ReturnString + currentChar
            i = i + 1
        End If
    Loop
    
    URLUnEncode = ReturnString
    
End Function

Public Function URLEncode(encodeString As String) As String
    Dim ReturnString As String
    Dim currentChar As String

    Dim i As Long

    For i = 1 To Len(encodeString)
        currentChar = Mid(encodeString, i, 1)

        If Asc(currentChar) < 91 And Asc(currentChar) > 64 Then

            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 123 And Asc(currentChar) > 96 Then

            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 58 And Asc(currentChar) > 47 Then

            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) = 32 Then
            ReturnString = ReturnString + "+"

        Else

            If Len(Hex(Asc(currentChar))) = 1 Then
                ReturnString = ReturnString + "%0" + Hex(Asc(currentChar))
            Else
                ReturnString = ReturnString + "%" + Hex(Asc(currentChar))
            End If
        End If
    Next

    URLEncode = ReturnString
End Function


