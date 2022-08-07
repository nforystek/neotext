#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modWeb"
#Const modWeb = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32

Private Type NET_CONTROL_BLOCK
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte
   ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Private Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Public Enum eInternetFlags
    INTERNET_FLAG_RESYNCHRONIZE = &H800
    INTERNET_FLAG_NEED_FILE = &H10
    INTERNET_FLAG_RELOAD = &H80000000
    
    INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000
    INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000
    INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
    INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
    SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
    
    INTERNET_FLAG_SECURE = &H800000
    
    INTERNET_FLAG_NO_UI = &H200
    INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
    INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000
    INTERNET_FLAG_NO_COOKIES = &H80000
End Enum

Private Declare Function Netbios Lib "netapi32" (pncb As NET_CONTROL_BLOCK) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function InternetSetOption Lib "wininet" _
    Alias "InternetSetOptionA" _
    (ByVal hInternet As Long, _
    ByVal lOption As Long, _
    ByRef sBuffer As Any, _
    ByVal lBufferLength As Long) As Integer

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
        
Private Declare Function HttpEndRequest Lib "wininet" _
        Alias "HttpEndRequestA" _
        (ByVal hRequest As Long, _
        ByVal lpBuffersOut As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
             
Public Function PostToWebsite(ByVal HostServerAddress As String, Optional ByVal WebFilePath As String = "/", Optional ByVal PostFormData As String = "", Optional ByVal AuthUsername As String = "", Optional ByVal AuthPassword As String = "", Optional ByVal CacheAndCookies As Boolean = False, Optional ByVal SecureSocketLayer As Boolean = False) As String
    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim sVerb As String
    sVerb = IIf(PostFormData = "", "GET", "POST")
    Dim bRet As Boolean
    Dim lPort As Long
    Dim lTimeout As Long
    lTimeout = 10000
    Dim lFlags As eInternetFlags
    If CacheAndCookies Then
        lFlags = eInternetFlags.INTERNET_FLAG_RELOAD
    Else
        lFlags = eInternetFlags.INTERNET_FLAG_RELOAD + eInternetFlags.INTERNET_FLAG_NO_CACHE_WRITE + eInternetFlags.INTERNET_FLAG_NO_COOKIES
    End If
    
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0
    
    Const INTERNET_OPEN_TYPE_DIRECT As Long = 1
    hInternetOpen = InternetOpen(App.EXEName, _
                    INTERNET_OPEN_TYPE_DIRECT, _
                    vbNullString, _
                    vbNullString, _
                    0)
    
    If hInternetOpen <> 0 Then
        Const INTERNET_SERVICE_HTTP = 3
        Const INTERNET_DEFAULT_HTTP_PORT = 80
        Const INTERNET_DEFAULT_HTTPS_PORT = 443
        If (Left(LCase(Trim(HostServerAddress)), 5) = "https") Or SecureSocketLayer Then
            lPort = INTERNET_DEFAULT_HTTPS_PORT
            HostServerAddress = Replace(LCase(Trim(HostServerAddress)), "https://", "")
                        
            lFlags = lFlags Or eInternetFlags.INTERNET_FLAG_SECURE Or _
                 eInternetFlags.INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                 eInternetFlags.INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
                 
        ElseIf (Left(LCase(Trim(HostServerAddress)), 4) = "http") Or (Not SecureSocketLayer) Then
            lPort = INTERNET_DEFAULT_HTTP_PORT
            HostServerAddress = Replace(LCase(Trim(HostServerAddress)), "http://", "")
        Else
            lPort = INTERNET_DEFAULT_HTTP_PORT
        End If
        
        hInternetConnect = InternetConnect(hInternetOpen, _
                           HostServerAddress, _
                           lPort, _
                           AuthUsername, _
                           AuthPassword, _
                           INTERNET_SERVICE_HTTP, _
                            0, _
                           0)
    
        If hInternetConnect <> 0 Then
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                                 sVerb, _
                                 WebFilePath, _
                                 "HTTP/1.0", _
                                 vbNullString, _
                                 0, _
                                 lFlags, _
                                 0)
        
            If hHttpOpenRequest <> 0 Then
                Dim sHeader As String
                Const HTTP_ADDREQ_FLAG_ADD = &H20000000
                Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
                sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
                bRet = HttpAddRequestHeaders(hHttpOpenRequest, _
                        sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE _
                        Or HTTP_ADDREQ_FLAG_ADD)

                Const INTERNET_OPTION_CONNECT_TIMEOUT As Long = 2
                Const INTERNET_OPTION_RECEIVE_TIMEOUT As Long = 6
                Const INTERNET_OPTION_SEND_TIMEOUT As Long = 5
                Const INTERNET_OPTION_SECURITY_FLAGS As Long = 31

                Call InternetSetOption(hHttpOpenRequest, _
                                       INTERNET_OPTION_CONNECT_TIMEOUT, _
                                       lTimeout, _
                                       4)
                Call InternetSetOption(hHttpOpenRequest, _
                                       INTERNET_OPTION_RECEIVE_TIMEOUT, _
                                       lTimeout, _
                                       4)
                Call InternetSetOption(hHttpOpenRequest, _
                                       INTERNET_OPTION_SEND_TIMEOUT, _
                                       lTimeout, _
                                       4)
                                       
                If lPort = INTERNET_DEFAULT_HTTPS_PORT Then
                    Dim lSecFlag            As Long
                    lSecFlag = eInternetFlags.SECURITY_FLAG_IGNORE_UNKNOWN_CA
                    Call InternetSetOption(hHttpOpenRequest, _
                                           INTERNET_OPTION_SECURITY_FLAGS, _
                                           lSecFlag, _
                                           4)
                End If
        
                Dim lpszPostData As String
                Dim lPostDataLen As Long
        
                lpszPostData = PostFormData
                lPostDataLen = Len(lpszPostData)
                bRet = HttpSendRequest(ByVal hHttpOpenRequest, _
                        ByVal vbNullString, _
                        ByVal 0, _
                        ByVal lpszPostData, _
                        ByVal lPostDataLen)
        
                Dim bDoLoop             As Boolean
                Dim sReadBuffer         As String * 2048
                Dim lNumberOfBytesRead  As Long
                Dim sBuffer             As String
                bDoLoop = True
                Do While bDoLoop
                    sReadBuffer = vbNullString
                    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left(sReadBuffer, lNumberOfBytesRead)
                    If (Not CBool(lNumberOfBytesRead)) Or (Not bDoLoop) Then Exit Do
                    DoTasks
                Loop
            End If
        End If
    End If


    HttpEndRequest hHttpOpenRequest, 0, 0, 0
    InternetCloseHandle hHttpOpenRequest
    
    InternetCloseHandle hInternetConnect
    
    InternetCloseHandle hInternetOpen

    PostToWebsite = sBuffer

End Function

Public Function URLDecode(ByVal encodedString As String) As String

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
    
    URLDecode = ReturnString

End Function

Public Function URLEncode(ByVal encodeString As String) As String
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

Public Function GetMacAddress() As String
   
   Dim tmp As String
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT

   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)

   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT

   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   
   pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
            Or HEAP_ZERO_MEMORY, NCB.ncb_length)
            
   If Not pASTAT = 0 Then
   
        NCB.ncb_buffer = pASTAT
        Call Netbios(NCB)
        
        CopyMemory AST, NCB.ncb_buffer, Len(AST)
        
        tmp = Format$(Hex(AST.adapt.adapter_address(0)), "00") & " " & _
              Format$(Hex(AST.adapt.adapter_address(1)), "00") & " " & _
              Format$(Hex(AST.adapt.adapter_address(2)), "00") & " " & _
              Format$(Hex(AST.adapt.adapter_address(3)), "00") & " " & _
              Format$(Hex(AST.adapt.adapter_address(4)), "00") & " " & _
              Format$(Hex(AST.adapt.adapter_address(5)), "00")
        
        HeapFree GetProcessHeap(), 0, pASTAT
        
        GetMacAddress = tmp

   End If

End Function

Public Function OpenWebsite(ByVal theSite As String, Optional ByVal Silent As Boolean = False) As Integer
    
    Dim FileName As String
    Dim BrowserExec As String * 255
    Dim retVal As Integer
    Dim FileNumber As Integer

    BrowserExec = Space(255)
    FileName = AppPath & "temp.htm"
    
    FileNumber = FreeFile
    Open FileName For Output As #FileNumber
          Write #FileNumber, "<HTML> <\HTML>"
    Close #FileNumber

    retVal = FindExecutable(FileName, 0&, BrowserExec)
    BrowserExec = Trim(BrowserExec)

    If retVal <= 32 Or IsEmpty(BrowserExec) Then
        If Not Silent Then MsgBox "Could not find your default internet browser!", vbExclamation, "Open Website"
        retVal = False
    Else
        retVal = ShellExecute(0, "open", BrowserExec, """" & theSite & """", 0&, vbNormalFocus)
        If retVal <= 32 Then
            If Not Silent Then MsgBox "Unable to open site.", vbExclamation, "Open Website"
            retVal = False
        Else
            retVal = True
        End If
    End If

    Kill FileName
    OpenWebsite = retVal

End Function
