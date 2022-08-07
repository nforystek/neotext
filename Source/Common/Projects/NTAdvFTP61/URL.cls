VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "URL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary


Private Const NTP_Port = 123
Private Const HTTP_Port = 80
Private Const FTP_Port = 21

Public Enum URLTypes
    Unknown = 0
    File = 1
    Remote = 2
    ftp = 3
    HTTP = 4
    Server = 5
End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function GetSSL(ByVal theURL As String) As String
    GetSSL = (InStr(LCase(theURL), "s://") > 0)
End Function
Public Function GetShareName(ByVal theURL As String) As String

    Dim tmpInt As Integer
    tmpInt = InStr(3, theURL, "\")
    If tmpInt > 0 Then
        tmpInt = InStr(tmpInt + 1, theURL, "\")
        If tmpInt > 0 Then
            GetShareName = Left(theURL, tmpInt - 1)
        Else
            GetShareName = theURL
        End If
    Else
        GetShareName = theURL
    End If

End Function
Public Function GetUserName(ByVal theURL As String) As String

    Dim Login As String
    theURL = Trim(Replace(theURL, "s://", "://"))
    If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
        Dim testURL As String
        If Left(theURL, 7) = "http://" Then
            testURL = Mid(theURL, 8)
        Else
            testURL = Mid(theURL, 7)
        End If
        If InStr(testURL, "@") > 0 Then
            testURL = Left(testURL, InStr(testURL, "@") - 1)
            If InStr(testURL, ":") > 0 Then
                Login = Left(testURL, InStr(testURL, ":") - 1)
            Else
                Login = testURL
            End If
        Else
            Login = ""
        End If
    Else
        Login = ""
    End If
    GetUserName = Login

End Function
Public Function GetPassword(ByVal theURL As String) As String

    Dim Password As String
    theURL = Trim(Replace(theURL, "s://", "://"))
    If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
        Dim testURL As String
        If Left(theURL, 7) = "http://" Then
            testURL = Mid(theURL, 8)
        Else
            testURL = Mid(theURL, 7)
            End If
        If InStr(testURL, "@") > 0 Then
            testURL = Left(testURL, InStr(testURL, "@") - 1)
            If InStr(testURL, ":") > 0 Then
                Password = Mid(testURL, InStr(testURL, ":") + 1)
            Else
                Password = ""
            End If
        Else
            Password = ""
        End If
    Else
        Password = ""
    End If
    GetPassword = Password

End Function
Public Function GetServer(ByVal theURL As String) As String

    Dim Server As String
    theURL = Trim(Replace(theURL, "s://", "://"))
    If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
        Dim testURL As String
        If Left(theURL, 7) = "http://" Then
            testURL = Mid(theURL, 8)
        Else
            testURL = Mid(theURL, 7)
        End If
        If InStr(testURL, "@") > 0 Then
            testURL = Mid(testURL, InStr(testURL, "@") + 1)
        End If
        If InStr(testURL, ":") > 0 Then
            testURL = Left(testURL, InStr(testURL, ":") - 1)
        End If
        If InStr(testURL, "/") > 0 Then
            Server = Left(testURL, InStr(testURL, "/") - 1)
        Else
            Server = Trim(testURL)
        End If
    Else
        Server = ""
    End If
    GetServer = Server

End Function
Public Function GetFolder(ByVal theURL As String) As String

    Dim Directory As String
    theURL = Trim(Replace(theURL, "s://", "://"))
    If Left(theURL, 7) = "file://" Then
        If Mid(theURL, 8, 1) = "/" Then
            theURL = Mid(theURL, 9)
            theURL = Replace(theURL, "|", ":")
            Directory = Replace(theURL, "/", "\")
            
        Else
            theURL = Mid(theURL, 8)
            Directory = "\\" & Replace(theURL, "/", "\")
        End If
    
    ElseIf (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://" Or Left(theURL, 7) = "http://") Then
        Dim testURL As String
        If Left(theURL, 7) = "http://" Then
            testURL = Mid(theURL, 8)
        Else
            testURL = Mid(theURL, 7)
        End If
        If InStr(testURL, "/") > 0 Then
            Directory = Trim(Mid(testURL, InStr(testURL, "/")))
            If Right(Directory, 1) = "/" And Len(Directory) > 1 Then Directory = Left(Directory, Len(Directory) - 1)
        Else
            Directory = ""
        End If
    
    Else
        If Left(theURL, 2) = "\\" Then
            If Right(Directory, 1) = "\" And Len(Directory) > 3 Then Directory = Left(Directory, Len(Directory) - 1)
            Directory = theURL
        Else
            If Mid(theURL, 2, 1) = ":" Then
                Directory = Trim(theURL)
                If Len(theURL) = 2 Then
                    theURL = theURL + "\"
                Else
                    If Right(Directory, 1) = "\" And Len(Directory) > 3 Then Directory = Left(Directory, Len(Directory) - 1)
                End If
            Else
                Directory = ""
            End If
        End If
    End If
        
    GetFolder = Directory

End Function
Public Function GetPort(ByVal theURL As String) As Long
On Error GoTo catch

    Dim Port As Long
    theURL = Trim(Replace(theURL, "s://", "://"))
    If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
        Dim testURL As String
        If Left(theURL, 7) = "http://" Then
            testURL = Mid(theURL, 8)
        Else
            testURL = Mid(theURL, 7)
        End If
        If InStr(testURL, "@") > 0 Then
            testURL = Mid(testURL, InStr(testURL, "@") + 1)
        End If
        If InStr(testURL, ":") > 0 Then
            testURL = Mid(testURL, InStr(testURL, ":") + 1)
            If InStr(testURL, "/") > 0 Then
                Port = CLng(Trim(Left(testURL, InStr(testURL, "/") - 1)))
            Else
                Port = CLng(Trim(testURL))
            End If
        Else
            If Left(theURL, 6) = "ntp://" Then
                Port = NTP_Port
            Else
                If Left(theURL, 6) = "ftp://" Then
                    Port = FTP_Port
                Else
                    If Left(theURL, 7) = "http://" Then
                        Port = HTTP_Port
                    Else
                        Port = 0
                    End If
                End If
            End If
        End If
    Else
        If Left(theURL, 6) = "ntp://" Then
            Port = NTP_Port
        Else
            If Left(theURL, 6) = "ftp://" Then
                Port = FTP_Port
            Else
                If Left(theURL, 7) = "http://" Then
                    Port = HTTP_Port
                Else
                    Port = 0
                End If
            End If
        End If
    End If
    GetPort = Port

Exit Function
catch:
    Err.Clear
    GetPort = FTP_Port
End Function
Public Function GetDirChar(ByVal theURL As String) As String

If InStr(theURL, "/") > 0 Then
    GetDirChar = "/"
Else
    If InStr(theURL, "\") > 0 Then
        GetDirChar = "\"
    Else
        GetDirChar = "/"
        End If
    End If

End Function
Public Function GetParentFolder(ByVal theURL As String) As String
    
    Dim tmpStr As String
    Dim msn As String
    Dim URLType As URLTypes
    Dim dirChar As String
    theURL = Trim(Replace(theURL, "s://", "://"))
    
    dirChar = GetDirChar(theURL)
    tmpStr = Trim(theURL)

    If IsRoot(tmpStr) Then
        GetParentFolder = tmpStr
        Exit Function
    End If
    
    If Left(tmpStr, 2) = "\\" Then
        tmpStr = Mid(tmpStr, 3)
        msn = "\\" + Left(tmpStr, InStr(tmpStr, "\"))
        tmpStr = Mid(tmpStr, InStr(tmpStr, "\") + 1)
    Else
        If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(tmpStr, 7) = "http://" Then
            msn = ""
        Else
            msn = ""
        End If
    End If
    
    tmpStr = StrReverse(tmpStr)
    If Left(tmpStr, 1) = dirChar Then tmpStr = Mid(tmpStr, 2)
    If InStr(tmpStr, dirChar) > 0 Then
        tmpStr = StrReverse(Mid(tmpStr, InStr(tmpStr, dirChar) + 1))
        If Len(Trim(tmpStr)) = 0 Then tmpStr = dirChar
        If msn <> "" Then
            tmpStr = msn + tmpStr
        Else
            If (Mid(tmpStr, 2, 1) = ":") And Len(tmpStr) = 2 Then
                tmpStr = tmpStr + dirChar
            Else

            End If
        End If
        GetParentFolder = Trim(tmpStr)
    Else
        GetParentFolder = Trim(StrReverse(tmpStr))
    End If
    
End Function
Public Function GetFile(ByVal theURL As String) As String
    Dim dChar As String
    dChar = GetDirChar(theURL)
    If InStrRev(theURL, dChar) > 0 Then
        GetFile = Mid(theURL, InStrRev(theURL, dChar) + 1)
    Else
        GetFile = theURL
    End If

End Function
Public Function SetFolder(ByVal theURL As String, ByVal newDir As String) As String

    Dim tmpURL As String

    If (Left(Trim(Replace(theURL, "s://", "://")), 6) = "ftp://" Or Left(Trim(Replace(theURL, "s://", "://")), 6) = "ntp://") Or Left(Trim(Replace(theURL, "s://", "://")), 7) = "http://" Then
        If InStr(InStr(theURL, "://") + 3, theURL, "/") > 0 Then
            tmpURL = Left(theURL, InStr(InStr(theURL, "://") + 3, theURL, "/") - 1)
            tmpURL = tmpURL + newDir
        Else
            tmpURL = theURL + newDir
        End If
    Else
        tmpURL = newDir
    End If
    SetFolder = tmpURL

End Function
Public Function IsRoot(ByVal theURL As String) As Boolean
    
    Dim tmpStr As String
    Dim dirChar As String
    Dim msn As String
    dirChar = GetDirChar(theURL)
    tmpStr = Trim(theURL)
    theURL = Trim(Replace(theURL, "s://", "://"))
    
    If Left(tmpStr, 2) = "\\" Then
        tmpStr = Mid(tmpStr, 3)
        msn = "\\" + Left(tmpStr, InStr(tmpStr, "\"))
        tmpStr = Mid(tmpStr, InStr(tmpStr, "\") + 1)
    Else
        If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(tmpStr, 7) = "http://" Then
            If Left(tmpStr, 7) = "http://" Then
                tmpStr = Mid(tmpStr, 8)
            Else
                tmpStr = Mid(tmpStr, 7)
            End If
            msn = ""
        Else
            msn = ""
        End If
    End If
    
    tmpStr = StrReverse(tmpStr)
    If Left(tmpStr, 1) = dirChar Then tmpStr = Mid(tmpStr, 2)
    If InStr(tmpStr, dirChar) > 0 Then
        IsRoot = False
    Else
        IsRoot = True
    End If
        
End Function

Public Function GetType(ByVal theURL As String) As URLTypes

    Dim whatsURL As URLTypes
    theURL = Trim(Replace(theURL, "s://", "://"))
    If (LCase(Left(Trim(theURL), 6)) = "ftp://" Or LCase(Left(Trim(theURL), 6)) = "ntp://") Then
        whatsURL = URLTypes.ftp
    Else
        If (LCase(Left(Trim(theURL), 7)) = "http://") Or _
            Left(Trim(LCase(theURL)), 3) = "www" Or _
            Right(Trim(LCase(theURL)), 4) = ".txt" Or _
            Right(Trim(LCase(theURL)), 4) = ".htm" Or _
            Right(Trim(LCase(theURL)), 5) = ".html" Or _
            Right(Trim(LCase(theURL)), 4) = ".asp" Then
    
            whatsURL = URLTypes.HTTP
        Else
            If LCase(Left(Trim(theURL), 8)) = "file:///" Or LCase(Mid(Trim(theURL), 2, 1)) = ":" Then
                whatsURL = URLTypes.File
            Else
                If LCase(Left(Trim(theURL), 7)) = "file://" Or LCase(Left(Trim(theURL), 2)) = "\\" Then
                    whatsURL = URLTypes.Remote
                Else
                    whatsURL = URLTypes.File
                End If
            End If
        End If
    End If
    GetType = whatsURL

End Function

Public Function ValidateURL(ByRef theURL As String) As Boolean

    Dim URLOk As Boolean
    Dim eDir As String
    URLOk = True

    If (Left(Trim(Replace(theURL, "s://", "://")), 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "file://" Or Left(theURL, 2) = "\\" Or Mid(theURL, 2, 1) = ":" Then
        If Not (Left(Trim(Replace(theURL, "s://", "://")), 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Then
            eDir = theURL
            If LCase(Left(eDir, 7)) = "file://" Then
                If Mid(eDir, 8, 1) = "/" Then
                    eDir = Mid(eDir, 9)
                Else
                    eDir = Mid(eDir, 6)
                End If
                eDir = Replace(Replace(eDir, "|", ":"), "/", "\")
            End If
            
            If Mid(eDir, 2, 1) = ":" Or Left(eDir, 2) = "\\" Then
                If Right(eDir, 1) = "\" And Len(eDir) > 3 Then eDir = Left(eDir, Len(eDir) - 1)
                theURL = eDir
            End If
        Else
            If Right(eDir, 1) = "/" And Len(eDir) > 1 Then eDir = Left(eDir, Len(eDir) - 1)
        End If
    
        If Err > 0 Then
            Err.Clear
            URLOk = False
        End If
    Else
        URLOk = False
    End If
    
    ValidateURL = URLOk

End Function
Public Function ValidateRemoteURL(ByVal theURL As String, eServer As String, ePort As Integer, eDir As String) As String

    eDir = Mid(theURL, 7)
    If InStr(eDir, "/") > 0 Then
        eServer = Left(eDir, InStr(eDir, "/") - 1)
        eDir = Mid(eDir, InStr(eDir, "/"))
        If Right(eDir, 1) = "/" And Not (Len(eDir) = 1) Then eDir = Left(eDir, Len(eDir) - 1)
    Else
        eServer = Trim(eDir)
        eDir = "/"
    End If
    If InStr(eServer, "@") > 0 Then
        eServer = Mid(eServer, InStr(eServer, "@") + 1)
    End If
    If InStr(eServer, ":") > 0 Then
        ePort = CInt(Mid(eServer, InStr(eServer, ":") + 1))
        eServer = Left(eServer, InStr(eServer, ":") - 1)
    Else
        ePort = FTP_Port
    End If

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
            ReturnString = ReturnString + Chr(Val("&H" & currentChar))
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












