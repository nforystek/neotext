VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Private Const Base36$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Enum DetailOptions
    None = 0
    File = 1
    post = 2
End Enum

Private EventText As String
Private EventFlag As Long

Private WithEvents sck As ISocket
Attribute sck.VB_VarHelpID = -1

Private pAuthenticated As Boolean
Private pServer As String
Private pPort As Long
Private pBytes As Long

Private pAuthUser As String
Private pAuthPass As String

Private pPostAs As String
Private pGroups As String

Private pUseID As String

Private pDetails As DetailOptions
Private pFullPaths As Boolean


'Public Sub Cancel()
'    Do Until EventFlag = 10
'        If Not EventFlag = 10 Then EventFlag = -10
'        DoTasks
'    Loop
'End Sub

Public Property Get FullPaths() As Boolean
    FullPaths = pFullPaths
End Property
Public Property Let FullPaths(ByVal newVal As Boolean)
    pFullPaths = newVal
End Property

Public Property Get Details() As DetailOptions
    Details = pDetails
End Property
Public Property Let Details(ByVal newVal As DetailOptions)
    pDetails = newVal
End Property

Public Property Get Server() As String
    Server = pServer
End Property
Public Property Let Server(ByVal newVal As String)
    pServer = newVal
End Property

Public Property Get Port() As Long
    Port = pPort
End Property
Public Property Let Port(ByVal newVal As Long)
    pPort = newVal
End Property

Public Property Get MessageSize() As Long
    MessageSize = pBytes
End Property
Public Property Let MessageSize(ByVal newVal As Long)
    pBytes = newVal
End Property

Public Property Get AuthUser() As String
    AuthUser = pAuthUser
End Property
Public Property Let AuthUser(ByVal newVal As String)
    pAuthUser = newVal
End Property

Public Property Get AuthPass() As String
    AuthPass = pAuthPass
End Property
Public Property Let AuthPass(ByVal newVal As String)
    pAuthPass = newVal
End Property

Public Property Get PostAs() As String
    PostAs = pPostAs
End Property
Public Property Let PostAs(ByVal newVal As String)
    pPostAs = newVal
End Property

Public Property Get NewsGroup() As String
    NewsGroup = pGroups
End Property
Public Property Let NewsGroup(ByVal newVal As String)
    pGroups = newVal
End Property

Private Function Decim&(Sessa2$)
    Dim Posiz%, CifraT&, ValC&
    CifraT& = 0
    For Posiz% = 1 To Len(Sessa2$)
        ValC& = (InStr(Base36$, Mid$(Sessa2$, Len(Sessa2$) - Posiz% + 1, 1)) - 1) * 36 ^ (Posiz% - 1)
        CifraT& = CifraT& + ValC&
    Next Posiz%
    Decim& = CifraT&
End Function

Private Function Sessa2$(Decim&)
    Dim DeScr&, CifraT$, Cifra%
    DeScr& = Decim&
    CifraT$ = ""
    Do
        Cifra% = DeScr& Mod 36
        DeScr& = DeScr& \ 36
        CifraT$ = Mid$(Base36$, Cifra% + 1, 1) + CifraT$
    Loop Until DeScr& = 0
    Sessa2$ = CifraT$
End Function

Private Function FilePartNumber(ByVal num As Long, ByVal total As Long) As String
    Dim tmp As String
    
    If Len(Trim(CStr(total))) > 10 Then
        tmp = Trim(CStr(total))
    Else
        tmp = String(5 - Len(Trim(CStr(total))), "0") & Trim(CStr(total))
    End If
    
    If Len(Trim(CStr(num))) > 10 Then
        tmp = Trim(CStr(num)) & "/" & tmp
    Else
        tmp = String(5 - Len(Trim(CStr(num))), "0") & Trim(CStr(num)) & "/" & tmp
    End If
    FilePartNumber = tmp
End Function

Private Function MessageID() As String
    If pUseID <> "" Then
        MessageID = pUseID
    Else
        MessageID = Sessa2$(Int(1 + ((modBitValue.LongBound - 1) * Rnd))) & "." & Sessa2$(Int(1 + ((modBitValue.LongBound - 1) * Rnd))) & "@" & Sessa2$(Int(1 + ((modBitValue.LongBound - 1) * Rnd))) & Sessa2$(Int(1 + ((modBitValue.LongBound - 1) * Rnd)))
    End If
End Function

Private Function RecordDashes(ByVal mGUID As String) As String
    Dim pos As Integer
    pos = InStr(mGUID, "-")
    Do Until pos = 0
        RecordDashes = RecordDashes & pos & ";"
        pos = InStr(pos + 1, mGUID, "-")
    Loop
End Function

Private Function SetDashes(ByVal mGUID As String, ByVal mDash As String) As String
    mGUID = Replace(mGUID, "-", "")
    Dim pos As Integer
    Do Until mDash = ""
        pos = CInt(RemoveNextArg(mDash, ";"))
        mGUID = Left(mGUID, pos - 1) & "-" & Mid(mGUID, pos)
    Loop
    SetDashes = mGUID
End Function

Public Function PutBinary(ByVal FileName As String, Optional ByRef FileGUID As Variant = "") As Collection

    If pServer <> "" Then
        If pPostAs <> "" Then
            If pGroups <> "" Then

                    Dim detailheader As String
                    
                    Dim ci As Object
                    
                    Dim gid As String
                    Dim pId As String
                    
                    Set ci = CreateObject("NTCipher10.GUID")
                    gid = ci.Generate
                    
                    Dim poolDashes As String
                    Dim guidDashes As String

                    pId = ci.Generate '(FileName)

                    guidDashes = RecordDashes(gid)
                    poolDashes = RecordDashes(pId)

                    gid = SetDashes(gid, poolDashes)
                    pId = SetDashes(pId, guidDashes)
                    
                    Dim mIDS As New Collection
                    FileGUID = gid
                    
                    Dim fullSize As Long
                    Dim numParts As Long
                    Dim msgLines As Long
                    
                    Dim lineSize As Long

                    Dim padSpace As Long
                    
                    Dim firstpartheader As String
                    Dim middlepartheader As String
                    Dim finalpartheader As String

                    middlepartheader = "Path: " & sck.Host & vbLf
                    middlepartheader = middlepartheader & "From: " & PostAs & vbLf
                    middlepartheader = middlepartheader & "Newsgroups: " & NewsGroup & vbLf
                    middlepartheader = middlepartheader & "Subject: Title [" & IIf(pFullPaths, FileName, GetFileName(FileName)) & "] UU-Part []" & String("0", 11) & vbLf
                    middlepartheader = middlepartheader & "Message-ID: <" & Base36$ & ">" & vbLf
                    middlepartheader = middlepartheader & "X-Newsreader: NTAdvFTP61.Gorup" & vbLf
                    middlepartheader = middlepartheader & "X-Information: https://www.neotext.org" & vbLf
                    middlepartheader = middlepartheader & "X-GUniqueID: " & gid & vbLf
                    middlepartheader = middlepartheader & "X-PoolUID: " & pId & vbLf
                    middlepartheader = middlepartheader & vbLf
                    middlepartheader = middlepartheader & "." & vbLf
                    
                    padSpace = pBytes - Len(middlepartheader)
                    
                    If padSpace < 2 Then
                    
                        Err.Raise 8, , "Bytes too low for message sizes."
                    End If
                    
                    Dim uu As Object
                    
                    Set uu = CreateObject("NTCipher10.UUCode")
                        
                    Dim inLine As String
                    
                    Dim lCnt As Long
                    
                    If PathExists(FileName & ".uue", True) Then Kill FileName & ".uue"
                    
                    Dim fnum As Long
                    fnum = FreeFile
                    Open FileName & ".uue" For Output As #fnum
                    Close #fnum
                    
                    Open FileName & ".uue" For Binary As #fnum
                    Put #fnum, , StrConv(LoadResData(1, "PRIVATE"), vbUnicode) & vbLf & vbLf
                    Close #fnum
                    
                    uu.UUEncode FileName
                    
                    Open FileName & ".uue" For Binary As #fnum
                    Put #fnum, LOF(fnum) + 1, vbLf & StrConv(LoadResData(2, "PRIVATE"), vbUnicode) & vbLf
                    Close #fnum
                    
                    Set uu = Nothing
                    
                    fnum = FreeFile
                    Open FileName & ".uue" For Input As #fnum
                    Do
                        Line Input #fnum, inLine
                    Loop Until InStr(inLine, "begin") > 0
                    Line Input #fnum, inLine
                    Close #fnum
                    
                    lineSize = (Len(inLine) * 2) + 4
                    fullSize = ((FileLen(FileName & ".uue") - Len(StrConv(LoadResData(1, "PRIVATE"), vbUnicode) & vbLf & vbLf) - Len(vbLf & _
                        StrConv(LoadResData(2, "PRIVATE"), vbUnicode) & vbLf)) * 2) + Len(StrConv(LoadResData(1, "PRIVATE"), vbUnicode) & vbLf & vbLf)
                            
                    Open FileName & ".uue" For Input As #fnum

                    Line Input #fnum, inLine
                    Line Input #fnum, inLine
                    Line Input #fnum, inLine
                    
                    numParts = fullSize \ pBytes
                    If ((fullSize Mod pBytes) > 0) Or ((numParts = 0) And (fullSize > 0)) Then numParts = numParts + 1
                                        
                    msgLines = padSpace \ lineSize
                    If (((padSpace Mod lineSize) * numParts) \ pBytes) > 0 Then
                        numParts = numParts - (((padSpace Mod lineSize) * numParts) \ pBytes)
                    End If

                    Dim lapse As Single
                    Static dc As Boolean
                    dc = False
                    
                    If Not sck.Connected Then
                        pAuthenticated = False
                        
                        EventFlag = 0
                        If InStr(pServer, ":") > 0 Then
                            sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
                        Else
                            sck.Connect pServer, pPort
                        End If
                        pAuthenticated = False
                        lapse = Timer
                        Do Until (sck.Connected) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If Not sck.Connected Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , "Could not establish connection."
                        End If
                        
                        EventFlag = 0
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        Else
                            dc = True
                        End If
                    End If

                    If (sck.Connected) Then
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If Not pAuthenticated Then
                            If pAuthUser <> "" Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO USER " & pAuthUser & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHUSER " & pAuthUser & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            
                            If (pAuthPass <> "") Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO PASS " & pAuthPass & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHPASS " & pAuthPass & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            If (EventFlag = 3 Or EventFlag = 10) Then
                                pAuthenticated = False
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            Else
                                pAuthenticated = True
                            End If
                        End If

                        EventFlag = 0
                       ' sck.SendString "POST" & vbCrLf
                        sck.SendString "POST" & vbCrLf
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If
                        
                        detailheader = "Path: " & sck.Host & vbLf
                        detailheader = detailheader & "From: " & PostAs & vbLf
                        detailheader = detailheader & "Newsgroups: " & NewsGroup & vbLf
                        detailheader = detailheader & "Subject: Title [" & IIf(pFullPaths, FileName, GetFileName(FileName)) & "] UU-Part [" & FilePartNumber(0, numParts + 1) & "]" & vbLf
                        detailheader = detailheader & "Message-ID: <" & MessageID & ">" & vbLf
                        detailheader = detailheader & "X-Newsreader: NTAdvFTP61.Gorup" & vbLf
                        detailheader = detailheader & "X-Information: https://www.neotext.org" & vbLf
                        detailheader = detailheader & vbLf
                        detailheader = detailheader & "FileSize: " & FileLen(FileName) & vbLf
                        detailheader = detailheader & "FileDateTime: " & FileDateTime(FileName) & vbLf
                        detailheader = detailheader & "FileGUID: " & gid & vbLf
                        detailheader = detailheader & vbLf
                        detailheader = detailheader & "MessageID(s):" & vbLf

                        firstpartheader = "Path: " & sck.Host & vbLf
                        firstpartheader = firstpartheader & "From: " & PostAs & vbLf
                        firstpartheader = firstpartheader & "Newsgroups: " & NewsGroup & vbLf
                        firstpartheader = firstpartheader & "Subject: Title [" & IIf(pFullPaths, FileName, GetFileName(FileName)) & "] UU-Part [" & FilePartNumber(1, numParts + 1) & "]" & vbLf
                        mIDS.Add "<" & MessageID & ">"
                        detailheader = detailheader & mIDS(mIDS.count) & vbLf
                        firstpartheader = firstpartheader & "Message-ID: " & mIDS(mIDS.count) & vbLf
                        firstpartheader = firstpartheader & "X-Newsreader: NTAdvFTP61.Gorup" & vbLf
                        firstpartheader = firstpartheader & "X-Information: https://www.neotext.org" & vbLf
                        firstpartheader = firstpartheader & "X-GUniqueID: " & gid & vbLf
                        firstpartheader = firstpartheader & "X-PoolUID: " & pId & vbLf
                        firstpartheader = firstpartheader & vbLf
                        firstpartheader = firstpartheader & StrConv(LoadResData(1, "PUBLIC"), vbUnicode) & vbLf
                        firstpartheader = firstpartheader & vbLf
                        lCnt = 0
                        Do While (lCnt < msgLines) And (Left(inLine, 3) <> "end")
                            Line Input #fnum, inLine
                            firstpartheader = firstpartheader & StrConv(inLine, vbUnicode) & vbLf
                            lCnt = lCnt + 1
                        Loop
                        firstpartheader = firstpartheader & "." & vbLf
        
                        EventFlag = 0
                        sck.SendString firstpartheader
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If
                        
                        If fullSize > 0 Then
                        
                            Dim cnt As Long
                            For cnt = 1 To numParts - 1
                            
                                EventFlag = 0
                                sck.SendString "POST" & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    Close #fnum
                                    sck.Disconnect
                                    Err.Raise 8, , EventText
                                End If

                                pId = ci.Generate '(FileName)
                                poolDashes = RecordDashes(pId)
                                pId = SetDashes(pId, guidDashes)
                                gid = SetDashes(gid, poolDashes)
                                
                                middlepartheader = "Path: " & sck.Host & vbLf
                                middlepartheader = middlepartheader & "From: " & PostAs & vbLf
                                middlepartheader = middlepartheader & "Newsgroups: " & NewsGroup & vbLf
                                middlepartheader = middlepartheader & "Subject: Title [" & IIf(pFullPaths, FileName, GetFileName(FileName)) & "] UU-Part [" & FilePartNumber(cnt + 1, numParts + 1) & "]" & vbLf
                                mIDS.Add "<" & MessageID & ">"
                                detailheader = detailheader & mIDS(mIDS.count) & vbLf
                                middlepartheader = middlepartheader & "Message-ID: " & mIDS(mIDS.count) & vbLf
                                middlepartheader = middlepartheader & "X-Newsreader: NTAdvFTP61.Gorup" & vbLf
                                middlepartheader = middlepartheader & "X-Information: https://www.neotext.org" & vbLf
                                middlepartheader = middlepartheader & "X-GUniqueID: " & gid & vbLf
                                middlepartheader = middlepartheader & "X-PoolUID: " & pId & vbLf
                                middlepartheader = middlepartheader & vbLf
                                lCnt = 0
                                Do While (lCnt < msgLines) And (Left(inLine, 3) <> "end")
                                    Line Input #fnum, inLine
                                    middlepartheader = middlepartheader & StrConv(inLine, vbUnicode) & vbLf
                                    lCnt = lCnt + 1
                                Loop
                                middlepartheader = middlepartheader & "." & vbLf

                                EventFlag = 0
                                sck.SendString middlepartheader
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    Close #fnum
                                    sck.Disconnect
                                    Err.Raise 8, , EventText
                                End If

                            Next

                        End If

                        EventFlag = 0
                        sck.SendString "POST" & vbLf
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If

                        pId = ci.Generate '(FileName)
                        poolDashes = RecordDashes(pId)
                        pId = SetDashes(pId, guidDashes)
                        gid = SetDashes(gid, poolDashes)
                                
                        finalpartheader = "Path: " & sck.Host & vbLf
                        finalpartheader = finalpartheader & "From: " & PostAs & vbLf
                        finalpartheader = finalpartheader & "Newsgroups: " & NewsGroup & vbLf
                        finalpartheader = finalpartheader & "Subject: Title [" & IIf(pFullPaths, FileName, GetFileName(FileName)) & "] UU-Part [" & FilePartNumber(numParts + 1, numParts + 1) & "]" & vbLf
                        mIDS.Add "<" & MessageID & ">"
                        detailheader = detailheader & mIDS(mIDS.count) & vbLf
                        finalpartheader = finalpartheader & "Message-ID: " & mIDS(mIDS.count) & vbLf
                        finalpartheader = finalpartheader & "X-Newsreader: NTAdvFTP61.Gorup" & vbLf
                        finalpartheader = finalpartheader & "X-Information: https://www.neotext.org" & vbLf
                        finalpartheader = finalpartheader & "X-GUniqueID: " & gid & vbLf
                        finalpartheader = finalpartheader & "X-PoolUID: " & pId & vbLf
                        finalpartheader = finalpartheader & vbLf
                        finalpartheader = finalpartheader & StrConv(LoadResData(2, "PUBLIC"), vbUnicode) & vbLf
                        finalpartheader = finalpartheader & vbLf
                        finalpartheader = finalpartheader & "." & vbLf
    
                        detailheader = detailheader & vbLf & "." & vbLf
     
                        EventFlag = 0
                        sck.SendString finalpartheader
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            Close #fnum
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If
                        
                        If ((pDetails And DetailOptions.post) = DetailOptions.post) Then
                            detailheader = NextArg(detailheader, "Message-ID:") & "Message-ID: <" & MessageID & ">" & vbLf & _
                                    RemoveArg(RemoveArg(detailheader, "Message-ID:"), vbLf)
                        End If

                        If ((pDetails And DetailOptions.File) = DetailOptions.File) Then
                            WriteFile FileName & ".nws", detailheader
                        End If
                        
                        If ((pDetails And DetailOptions.post) = DetailOptions.post) Then
                                   
                            EventFlag = 0
                            sck.SendString "POST" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                Close #fnum
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            End If
                        
                            EventFlag = 0
                            sck.SendString detailheader
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                Close #fnum
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            End If
                        End If
                        
                        Set PutBinary = mIDS
                        
                        Set mIDS = Nothing
                        Set ci = Nothing
                    Else
                        
                        sck.Disconnect
                        Err.Raise 8, , EventText
                    End If
                        
                    If dc Then
                        EventFlag = 0
                        sck.SendString "QUIT" & vbLf
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        Else
                            sck.Disconnect
                        End If
                    End If
                        
                    Close #fnum
                    
            Else
                Err.Raise 8, "NewsGroup unspecified."
            End If
        Else
            Err.Raise 8, "Post As unspecified."
        End If
    Else
        Err.Raise 8, "Server unspecified."
    End If

    EventFlag = 10
End Function

Public Sub GetBinary(ByVal FileName As String, Optional ByRef FileGUID As Variant = "", Optional ByVal MessageIDs As Object = Nothing)
    If pServer <> "" Then
        If pPostAs <> "" Then
            If pGroups <> "" Then

                    If (GetFileExt(FileName, True, True) = "nws") And PathExists(FileName, True) Then
                        Dim txt As String
                        Dim ret As String
                        
                        txt = ReadFile(FileName)
                        If (InStr(txt, "MessageID(s):") > 0) Then
                            If (MessageIDs Is Nothing) Then
                                Set MessageIDs = New Collection
                            ElseIf MessageIDs.count > 0 Then
                                Do Until MessageIDs.count = 0
                                    MessageIDs.Remove 1
                                Loop
                            End If
                            RemoveNextArg txt, "Subject: Title ["
                            FileName = GetFilePath(FileName) & "\" & GetFileName(RemoveNextArg(txt, "] UU-Part [00000/"))
                            RemoveNextArg txt, "MessageID(s):" & vbLf
                            Do Until txt = ""
                                ret = RemoveNextArg(txt, vbLf)
                                If (Not (ret = "")) And (Not (ret = ".")) Then MessageIDs.Add ret
                            Loop
                        End If
                    End If
                
                    If (Not (MessageIDs Is Nothing)) Or IsMissing(MessageIDs) Then
                        If MessageIDs.count = 0 Then
                            Set MessageIDs = Nothing
                        End If
                    End If

                    Dim ci As Object
                    Dim gid As String
                    
                    Set ci = CreateObject("NTCipher10.GUID")
                    gid = RecordDashes(ci.Generate)
                    Set ci = Nothing
                    
                    Dim lapse As Single
                    Static dc As Boolean
                    dc = False
                    
                    If Not sck.Connected Then
                        pAuthenticated = False
                        
                        EventFlag = 0
                        If InStr(pServer, ":") > 0 Then
                            sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
                        Else
                            sck.Connect pServer, pPort
                        End If
                        pAuthenticated = False
                        lapse = Timer
                        Do Until (sck.Connected) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If Not sck.Connected Then
                            sck.Disconnect
                            Err.Raise 8, , "Could not establish connection."
                        End If
                        
                        EventFlag = 0
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        Else
                            dc = True
                        End If
                    End If

                    If (sck.Connected) Then
                    
                        If Not pAuthenticated Then
                            If pAuthUser <> "" Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO USER " & pAuthUser & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHUSER " & pAuthUser & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            
                            If (pAuthPass <> "") Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO PASS " & pAuthPass & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHPASS " & pAuthPass & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            If (EventFlag = 3 Or EventFlag = 10) Then
                                pAuthenticated = False
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            Else
                                pAuthenticated = True
                            End If
                        End If
                        
                        EventFlag = 0
                        sck.SendString "GROUP " & pGroups & vbLf
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            'sck.Disconnect
                            'Err.Raise 8, , EventText
                        End If
                        
                        If RemoveNextArg(EventText, " ") = "211" Then
                            Dim aCnt As Long
                            Dim aStart As Long
                            Dim aEnd As Long
                            Dim aDone As Boolean
                            
                            aCnt = RemoveNextArg(EventText, " ")
                            aStart = RemoveNextArg(EventText, " ")
                            aEnd = RemoveNextArg(EventText, " ")

                            If (MessageIDs Is Nothing) Then
                                EventFlag = 0
                                sck.SendString "STAT " & aStart & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    'sck.Disconnect
                                    'Err.Raise 8, , EventText
                                End If
                            End If


                            Dim Articles As New Collection
                            Dim msgID As String
                            
                            Do
                                
                                If Not (MessageIDs Is Nothing) Then
                                    If MessageIDs.count > 0 Then
                                        EventFlag = 0
                                        sck.SendString "HEAD " & MessageIDs.Item(1) & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag > 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If
                                        
                                        If (InStr(LCase(EventText), "subject: title [" & LCase(IIf(pFullPaths, FileName, GetFileName(FileName))) & "] uu-part [") > 0) Then
                                            If (NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf) <> "") Or _
                                                (NextArg(RemoveArg(EventText, "X-PoolUID: "), vbLf) <> "") Then
                                                If FileGUID = "" Then
                                                    FileGUID = SetDashes(Replace(NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf), "-", ""), gid)
                                                    msgID = MessageIDs(1)
                                                    Articles.Add EventText, msgID
                                                    If (InStr(LCase(EventText), "x-newsreader: ntadvftp61.gorup") = 0) Then
                                                        sck.Disconnect
                                                        Err.Raise 8, , "Invalid X-Newsreader in postings header."
                                                    End If
                                                ElseIf (Replace(FileGUID, "-", "") = Replace(NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf), "-", "")) Then
                                                    msgID = MessageIDs(1)
                                                    Articles.Add EventText, msgID
                                                    If (InStr(LCase(EventText), "x-newsreader: ntadvftp61.gorup") = 0) Then
                                                        sck.Disconnect
                                                        Err.Raise 8, , "Invalid X-Newsreader in postings header."
                                                    End If
                                                End If
                                            End If
                                        End If
                                        MessageIDs.Remove 1
                                        aDone = (MessageIDs.count = 0)

                                    End If
                                Else
                                    EventFlag = 0
                                    sck.SendString "HEAD" & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag > 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                    If EventFlag = 3 Or EventFlag = 10 Then
                                        sck.Disconnect
                                        Err.Raise 8, , EventText
                                    End If
                                    Debug.Print EventText
                                    If (InStr(LCase(EventText), "subject: title [" & LCase(IIf(pFullPaths, FileName, GetFileName(FileName))) & "] uu-part [") > 0) Then

                                        If (NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf) <> "") And _
                                            (NextArg(RemoveArg(EventText, "X-PoolUID: "), vbLf) <> "") Then
                                            If FileGUID = "" Then
                                                FileGUID = SetDashes(Replace(NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf), "-", ""), gid)
                                                msgID = "A" & NextArg(Replace(NextArg(EventText, vbLf), "221 ", ""), " ")
                                                Articles.Add EventText, msgID
                                                aDone = CLng(LTrimStrip(RemoveQuotedArg(EventText, "/", "]"), "0")) = Articles.count
                                                If (InStr(LCase(EventText), "x-newsreader: ntadvftp61.gorup") = 0) Then
                                                    sck.Disconnect
                                                    Err.Raise 8, , "Invalid X-Newsreader in postings header."
                                                End If
                                            ElseIf Replace(FileGUID, "-", "") = Replace(NextArg(RemoveArg(EventText, "X-GUniqueID: "), vbLf), "-", "") Then
                                                msgID = "A" & NextArg(Replace(NextArg(EventText, vbLf), "221 ", ""), " ")
                                                Articles.Add EventText, msgID
                                                aDone = CLng(LTrimStrip(RemoveQuotedArg(EventText, "/", "]"), "0")) = Articles.count
                                                If (InStr(LCase(EventText), "x-newsreader: ntadvftp61.gorup") = 0) Then
                                                    sck.Disconnect
                                                    Err.Raise 8, , "Invalid X-Newsreader in postings header."
                                                End If
                                            End If
                                            
                                        End If
                                    End If


                                    If (Not aDone) Then
                                    
                                        EventFlag = 0
                                        sck.SendString "NEXT" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            aDone = True
                                            'sck.Disconnect
                                            'Err.Raise 8, , EventText
                                        End If
                                        
                                    End If
                            
                                End If

                            Loop Until aDone
                            
                            Dim newids As Collection
                            Set newids = New Collection
                            
                            If Articles.count > 0 Then
                                
                                Dim fnum As Long
                                
                                fnum = FreeFile
                                
                                Open FileName & ".uud" For Output As #fnum
                                
                                Print #fnum, StrConv(LoadResData(1, "PRIVATE"), vbUnicode) & vbLf & vbLf
                                
                                aStart = 1: aCnt = 1
                                aEnd = Articles.count + 1
                                
                                Dim uud As String
                                Dim tmp As String
                                Dim lin As String
                                Dim aheader As String
                                Dim asubject As String
                                Dim prior As String
                                Dim atp As Long
                                Dim pos As Long

                                Do Until aCnt > aEnd
                                                                           
                                    For aStart = 1 To Articles.count

                                        aheader = Articles(aStart)
                                        
                                        asubject = "title [" & LCase(IIf(pFullPaths, FileName, GetFileName(FileName))) & "] uu-part [" & FilePartNumber(aCnt, aEnd) & "]" & vbLf
                                        If LCase(Left(LCase(RemoveArg(aheader, "Subject: ")), Len(asubject))) = asubject Then
                                            If Not MessageIDs Is Nothing Then
                                                msgID = RemoveArg(Replace(NextArg(aheader, vbLf), "221 ", ""), " ")
                                                newids.Add msgID
                                            Else
                                                msgID = NextArg(Replace(NextArg(aheader, vbLf), "221 ", ""), " ")
                                                newids.Add RemoveArg(Replace(NextArg(aheader, vbLf), "221 ", ""), " ")
                                            End If
                                            
                                            aheader = ""
                                        
                                            EventFlag = 0
                                            sck.SendString "BODY " & msgID & vbLf
                                            lapse = Timer
                                            Do Until (EventFlag > 0) Or (Timer - lapse > 30)
                                                DoTasks
                                            Loop
                                            If EventFlag = 3 Or EventFlag = 10 Then
                                                Close #fnum
                                                sck.Disconnect
                                                Err.Raise 8, , EventText
                                            End If

                                            EventText = Replace(EventText, StrConv(LoadResData(1, "PUBLIC"), vbUnicode) & vbLf & vbLf, "")
                                            EventText = Replace(EventText, StrConv(LoadResData(2, "PUBLIC"), vbUnicode) & vbLf & vbLf, "")
                                            
                                            EventText = RTrimStrip(Replace(EventText, Chr(0), ""), vbLf)
                                            
                                            aheader = RemoveArg(EventText, vbLf)

                                            If Left(aheader, 5) = "begin" Then
                                                pos = Len(NextArg(RemoveArg(aheader, vbLf), vbLf))
                                            End If
                                            
                                            atp = 1
                                            Do
                                                atp = InStr(atp, aheader, vbLf)
                                                If atp > 0 Then
                                                    If (Mid(aheader, atp, 3) <> vbLf & "M") Then
                                                        If (Len(Mid(aheader, atp + 2)) > pos) Then
                                                            aheader = Left(aheader, atp - 1) & Mid(aheader, atp + 2)
                                                        Else
                                                            atp = 0
                                                        End If
                                                    ElseIf ((atp > (pos * 2) + 4)) Then
                                                        If (Mid(aheader, atp - pos - 2, 3) <> vbLf & "M") Then
                                                            If (Len(Mid(aheader, atp + 2)) > pos) Then
                                                                aheader = Left(aheader, atp - 1) & Mid(aheader, atp + 2)
                                                            Else
                                                                atp = 0
                                                            End If
                                                        Else
                                                            If (atp + 2 < Len(aheader)) Then
                                                                atp = atp + 2
                                                            Else
                                                                atp = 0
                                                            End If
                                                        End If
                                                    Else
                                                        If (atp + 2 < Len(aheader)) Then
                                                            atp = atp + 2
                                                        Else
                                                            atp = 0
                                                        End If
                                                    End If
                                                End If
                                            Loop Until atp = 0

                                            Print #fnum, aheader

                                        Else
                                            msgID = ""
                                        End If
                                    
                                    Next
                                
                                    aCnt = aCnt + 1
                                Loop
                                
                                Print #fnum, vbLf & StrConv(LoadResData(2, "PRIVATE"), vbUnicode) & vbLf
                                
                                Close #fnum
                                sck.Disconnect
                                
                                If PathExists(FileName, True) Then Kill FileName
                                
                                Dim uu As Object
                                
                                Set uu = CreateObject("NTCipher10.UUCode")
                                If PathExists(FileName & ".uud", True) Then Kill FileName & ".uud"
                                uu.UUDecode FileName & ".uud"
                                Set uu = Nothing
                                
                                If MessageIDs Is Nothing Then
                                    Set MessageIDs = newids
                                    Set newids = Nothing
                                End If
                                
                            End If
                        
                        Else
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If
                
                        If dc Then
                            EventFlag = 0
                            sck.SendString "QUIT" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            Else
                                sck.Disconnect
                            End If
                        End If
                        
                    Else
                        sck.Disconnect
                        Err.Raise 8, , EventText
                    End If
            Else
                Err.Raise 8, "NewsGroup unspecified."
            End If
        Else
            Err.Raise 8, "Post As unspecified."
        End If
    Else
        Err.Raise 8, "Server unspecified."
    End If
    
    EventFlag = 10
End Sub

'Public Sub VerticalMov()
'    If pServer <> "" Then
'        If pPostAs <> "" Then
'
'            Dim gid As String
'            Dim lapse As Single
'            Static dc As Boolean
'            dc = False
'
'            If Not sck.Connected Then
'
'                pAuthenticated = False
'                EventFlag = 0
'                If InStr(pServer, ":") > 0 Then
'                    sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
'                Else
'                    sck.Connect pServer, pPort
'                End If
'                lapse = Timer
'                Do Until (sck.Connected) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If Not sck.Connected Then
'                    sck.Disconnect
'                    Err.Raise 8, , "Could not establish connection."
'                End If
'
'                EventFlag = 0
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    sck.Disconnect
'                    Err.Raise 8, , EventText
'                Else
'                    dc = True
'                End If
'
'            End If
'
'            If (sck.Connected) Then
'
'                        If Not pAuthenticated Then
'                            If pAuthUser <> "" Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO USER " & pAuthUser & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHUSER " & pAuthUser & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'
'                            If (pAuthPass <> "") Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO PASS " & pAuthPass & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHPASS " & pAuthPass & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'                            If (EventFlag = 3 Or EventFlag = 10) Then
'                                pAuthenticated = False
'                                sck.Disconnect
'                                Err.Raise 8, , EventText
'                            Else
'                                pAuthenticated = True
'                            End If
'                        End If
'
'                EventFlag = 0
'                sck.sendstring "GROUP " & Replace(NewsGroup, ",", "") & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                Dim msgpst As String
'
'                EventFlag = 0
'                sck.sendstring "POST" & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                msgpst = "Path: " & sck.host & vblf
'                msgpst = msgpst & "From: " & PostAs & vblf
'                msgpst = msgpst & "Newsgroups: " & NewsGroup & vblf
'                msgpst = msgpst & "Subject: ." & vblf
'                msgpst = msgpst & "Message-ID: <" & MessageID & ">" & vblf
'                msgpst = msgpst & "X-Newsreader: NTAdvFTP61.Gorup" & vblf
'                msgpst = msgpst & "X-Information: https://www.neotext.org" & vblf
'                msgpst = msgpst & vblf
'                msgpst = msgpst & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & ".." & vblf & _
'                                "." & vblf
'
'                EventFlag = 0
'                sck.sendstring msgpst
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                If dc Then
'                    EventFlag = 0
'                    sck.sendstring "QUIT" & vblf
'                    lapse = Timer
'                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                        DoTasks
'                    Loop
'                    If EventFlag = 3 Or EventFlag = 10 Then
'                        sck.Disconnect
'                        Err.Raise 8, , EventText
'                    Else
'                        sck.Disconnect
'                    End If
'                End If
'
'            Else
'                sck.Disconnect
'                Err.Raise 8, , EventText
'            End If
'
'        Else
'            Err.Raise 8, "Post As unspecified."
'        End If
'    Else
'        Err.Raise 8, "Server unspecified."
'    End If
'
'    EventFlag = 10
'End Sub
'
'
'Public Sub HorizontalMov()
'    If pServer <> "" Then
'        If pPostAs <> "" Then
'
'            Dim gid As String
'            Dim lapse As Single
'            Static dc As Boolean
'            dc = False
'
'            If Not sck.Connected Then
'
'                pAuthenticated = False
'                EventFlag = 0
'                If InStr(pServer, ":") > 0 Then
'                    sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
'                Else
'                    sck.Connect pServer, pPort
'                End If
'                lapse = Timer
'                Do Until (sck.Connected) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If Not sck.Connected Then
'                    sck.Disconnect
'                    Err.Raise 8, , "Could not establish connection."
'                End If
'
'                EventFlag = 0
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    sck.Disconnect
'                    Err.Raise 8, , EventText
'                Else
'                    dc = True
'                End If
'
'            End If
'
'            If (sck.Connected) Then
'
'                        If Not pAuthenticated Then
'                            If pAuthUser <> "" Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO USER " & pAuthUser & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHUSER " & pAuthUser & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'
'                            If (pAuthPass <> "") Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO PASS " & pAuthPass & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHPASS " & pAuthPass & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'                            If (EventFlag = 3 Or EventFlag = 10) Then
'                                pAuthenticated = False
'                                sck.Disconnect
'                                Err.Raise 8, , EventText
'                            Else
'                                pAuthenticated = True
'                            End If
'                        End If
'
'                EventFlag = 0
'                sck.sendstring "GROUP " & Replace(NewsGroup, ",", "") & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
''                EventFlag = 0
''                sck.sendstring "GROUP control.motion" & vblf
''                lapse = Timer
''                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
''                    DoTasks
''                Loop
''                If EventFlag = 3 Or EventFlag = 10 Then
''                    'sck.Disconnect
''                    'Err.Raise 8, , EventText
''                End If
'
'                Dim msgpst As String
'
'                EventFlag = 0
'                sck.sendstring "POST" & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                msgpst = "Path: " & sck.host & vblf
'                msgpst = msgpst & "From: " & PostAs & vblf
'                msgpst = msgpst & "Newsgroups: " & NewsGroup & vblf
'                msgpst = msgpst & "Subject: ............................................................" & vblf
'                msgpst = msgpst & "Message-ID: <" & MessageID & ">" & vblf
'                msgpst = msgpst & "X-Newsreader: NTAdvFTP61.Gorup" & vblf
'                msgpst = msgpst & "X-Information: https://www.neotext.org" & vblf
'                msgpst = msgpst & vblf
'                msgpst = msgpst & "............................................................" & vblf
'                msgpst = msgpst & "." & vblf
'
'                EventFlag = 0
'                sck.sendstring msgpst
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                If dc Then
'                    EventFlag = 0
'                    sck.sendstring "QUIT" & vblf
'                    lapse = Timer
'                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                        DoTasks
'                    Loop
'                    If EventFlag = 3 Or EventFlag = 10 Then
'                        sck.Disconnect
'                        Err.Raise 8, , EventText
'                    Else
'                        sck.Disconnect
'                    End If
'                End If
'
'            Else
'                sck.Disconnect
'                Err.Raise 8, , EventText
'            End If
'
'        Else
'            Err.Raise 8, "Post As unspecified."
'        End If
'    Else
'        Err.Raise 8, "Server unspecified."
'    End If
'
'    EventFlag = 10
'End Sub

Public Function PostCancel(ByRef Groups As Collection, ByVal WildCards As Collection) As Boolean
    If pServer <> "" Then
        If pPostAs <> "" Then

            Dim gid As String
            Dim lapse As Single
            Static dc As Boolean
            dc = False

            If Not sck.Connected Then

                pAuthenticated = False
                EventFlag = 0
                If InStr(pServer, ":") > 0 Then
                    sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
                Else
                    sck.Connect pServer, pPort
                End If
                lapse = Timer
                Do Until (sck.Connected) Or (Timer - lapse > 30)
                    DoTasks
                Loop
                If Not sck.Connected Then
                    sck.Disconnect
                    Err.Raise 8, , "Could not establish connection."
                End If

                EventFlag = 0
                lapse = Timer
                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                    DoTasks
                Loop
                If EventFlag = 3 Or EventFlag = 10 Then
                    sck.Disconnect
                    Err.Raise 8, , EventText
                Else
                    dc = True
                End If

            End If

            If (sck.Connected) Then

                        If Not pAuthenticated Then
                            If pAuthUser <> "" Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO USER " & pAuthUser & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHUSER " & pAuthUser & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If

                            If (pAuthPass <> "") Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO PASS " & pAuthPass & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHPASS " & pAuthPass & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            If (EventFlag = 3 Or EventFlag = 10) Then
                                pAuthenticated = False
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            Else
                                pAuthenticated = True
                            End If
                        End If

                Dim mpath As Boolean

                Do Until Groups.count = 0

                    EventFlag = 0
                    sck.SendString "GROUP " & Groups(1) & vbLf
                    lapse = Timer
                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                        DoTasks
                    Loop
'                    If EventFlag = 3 Or EventFlag = 10 Then
'                        sck.Disconnect
'                        Err.Raise 8, , EventText
'                    End If

                    If RemoveNextArg(EventText, " ") = "211" Then
                        Dim aCnt As Long
                        Dim aStart As Long
                        Dim aEnd As Long
                        Dim aDone As Boolean

                        aCnt = RemoveNextArg(EventText, " ")
                        aStart = RemoveNextArg(EventText, " ")
                        aEnd = RemoveNextArg(EventText, " ")

                        EventFlag = 0
                        sck.SendString "STAT" & vbLf
                        lapse = Timer
                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                            DoTasks
                        Loop
                        If EventFlag = 3 Or EventFlag = 10 Then
                            sck.Disconnect
                            Err.Raise 8, , EventText
                        End If

                        If RemoveNextArg(EventText, " ") = "223" Then

                            Dim Articles As New Collection
                            Dim msgID As String
                            Dim wcard As Variant
                            Dim msgpst As String
                            Dim temp As String

                            Do

                                EventFlag = 0
                                sck.SendString "HEAD" & vbLf
                                lapse = Timer
                                Do Until (EventFlag > 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop

                                For Each wcard In WildCards

                                    If TestWildCard(EventText, "*" & wcard & "*") Then

                                        temp = EventText
                                        If InStr(temp, "From: ") > 0 Then
                                            RemoveNextArg temp, "From:"
                                            msgID = "From: " & NextArg(temp, vbLf)
                                        ElseIf InStr(temp, "Sender: ") > 0 Then
                                            RemoveNextArg temp, "Sender:"
                                            msgID = "Sender: " & NextArg(temp, vbLf)
                                        End If
                                        temp = EventText
                                        RemoveNextArg temp, "Message-ID:"
                                        msgID = msgID & "|" & NextArg(temp, vbLf)
                                        Articles.Add msgID

                                        EventFlag = 0
                                        sck.SendString "GROUP control.cancel" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If


                                        PostCancel = True
                                        mpath = Not mpath
                                        EventFlag = 0
                                        sck.SendString "POST" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If
                                        mpath = Not mpath

                                        msgpst = ""
                                        'msgpst = "Path: " & IIf(mpath, "0", "zzz.3") & vblf
                                        msgpst = msgpst & NextArg(msgID, "|") & vbLf
                                        msgpst = msgpst & "Control: cancel " & RemoveArg(msgID, "|") & vbLf
                                        msgpst = msgpst & "Newsgroups: control.cancel, " & Groups(1) & vbLf
                                        msgpst = msgpst & "Article-ID: <" & MessageID & ">" & vbLf
                                        msgpst = msgpst & "Subject: " & RemoveArg(msgID, "|") & vbLf
                                        msgpst = msgpst & "" & RemoveArg(msgID, "|") & vbLf
                                        msgpst = msgpst & vbLf & "." & vbLf
                                        EventFlag = 0
                                        Debug.Print msgpst

                                        sck.SendString msgpst
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If
                                        
                                        EventFlag = 0
                                        sck.SendString "NEXT" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If

                                        EventFlag = 0
                                        sck.SendString "GROUP " & Groups(1) & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If

                                        EventFlag = 0
                                        sck.SendString "STAT" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If


                                        Exit For
                                    End If
                                Next

                                EventFlag = 0
                                sck.SendString "NEXT" & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop

                            Loop Until Left(EventText, 3) = "430" Or Left(EventText, 3) = "421"



                            EventFlag = 0
                            sck.SendString "GROUP control.cancel" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            End If


                            If Articles.count > 0 Then

                                For Each wcard In Articles

                                    Do

                                        PostCancel = True
                                        mpath = Not mpath
                                        EventFlag = 0
                                        sck.SendString "POST" & vbLf
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If

                                        msgpst = ""
                                        'msgpst = "Path: " & IIf(mpath, "0", "zzz.3") & vbLf
                                        msgpst = msgpst & "Control: cancel " & RemoveArg(wcard, "|") & vbLf
                                        msgpst = msgpst & "Newsgroups: control.cancel" & vbLf
                                        msgpst = msgpst & NextArg(wcard, "|") & vbLf
                                        msgpst = msgpst & "Article-ID: <" & MessageID & ">" & vbLf
                                        msgpst = msgpst & "Subject: " & RemoveArg(wcard, "|") & vbLf
                                        msgpst = msgpst & "" & RemoveArg(wcard, "|") & vbLf
                                        msgpst = msgpst & vbLf & "." & vbLf
                                        EventFlag = 0
                                        Debug.Print msgpst

                                        sck.SendString msgpst
                                        lapse = Timer
                                        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                            DoTasks
                                        Loop
                                        If EventFlag = 3 Or EventFlag = 10 Then
                                            sck.Disconnect
                                            Err.Raise 8, , EventText
                                        End If


                                    Loop While (EventFlag = 3 Or EventFlag = 10 Or Left(EventText, 3) = "421")


                                Next

                            End If

                            EventFlag = 0
                            sck.SendString "STAT" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            End If

                            EventFlag = 0
                            sck.SendString "GROUP " & Groups(1) & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            End If

                            EventFlag = 0
                            sck.SendString "LAST" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop

                            EventFlag = 0
                            sck.SendString "STAT" & vbLf
                            lapse = Timer
                            Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                DoTasks
                            Loop
                            If EventFlag = 3 Or EventFlag = 10 Then
                                'sck.Disconnect
                                'Err.Raise 8, , EventText
                            End If

                        End If


                    Else
'                        sck.Disconnect
'                        Err.Raise 8, , EventText
                    End If

                    Groups.Remove 1
                Loop

                If dc Then
                    EventFlag = 0
                    sck.SendString "QUIT" & vbLf
                    lapse = Timer
                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                        DoTasks
                    Loop
                    If EventFlag = 3 Or EventFlag = 10 Then
                        sck.Disconnect
                        Err.Raise 8, , EventText
                    Else
                        sck.Disconnect
                    End If
                End If

            Else
                sck.Disconnect
                Err.Raise 8, , EventText
            End If

        Else
            Err.Raise 8, "Post As unspecified."
        End If
    Else
        Err.Raise 8, "Server unspecified."
    End If

    EventFlag = 10
End Function

Public Sub Connect()
    If pServer <> "" Then

        If Not sck.Connected Then

            Dim gid As String
            Dim lapse As Single

            EventFlag = 0
            If InStr(pServer, ":") > 0 Then
                sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
            Else
                sck.Connect pServer, pPort
            End If
            lapse = Timer
            Do Until (sck.Connected) Or (Timer - lapse > 30)
                DoTasks
            Loop
            If Not sck.Connected Then
                sck.Disconnect
                Err.Raise 8, , "Could not establish connection."
            End If

            EventFlag = 0
            lapse = Timer
            Do Until (EventFlag <> 0) Or (Timer - lapse > 30) Or sck.Connected
                DoTasks
            Loop
            If EventFlag = 3 Or EventFlag = 10 Then
                sck.Disconnect
                Err.Raise 8, , EventText
            End If
            
            
        'Else
            'Err.Raise 8, "Connection in use."
        End If
        
        If sck.Connected Then

                        If Not pAuthenticated Then
                            If pAuthUser <> "" Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO USER " & pAuthUser & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHUSER " & pAuthUser & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            
                            If (pAuthPass <> "") Then
                                EventFlag = 0
                                sck.SendString "AUTHINFO PASS " & pAuthPass & vbLf
                                lapse = Timer
                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                    DoTasks
                                Loop
                                If EventFlag = 3 Or EventFlag = 10 Then
                                    EventFlag = 0
                                    sck.SendString "AUTHPASS " & pAuthPass & vbLf
                                    lapse = Timer
                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
                                        DoTasks
                                    Loop
                                End If
                            End If
                            If (EventFlag = 3 Or EventFlag = 10) Then
                                pAuthenticated = False
                                sck.Disconnect
                                Err.Raise 8, , EventText
                            Else
                                pAuthenticated = True
                            End If
                        End If

        End If

    Else
        Err.Raise 8, "Server unspecified."
    End If

    EventFlag = 10
End Sub

Public Sub Disconnect()

    If sck.Connected Then

        Dim lapse As Single

        EventFlag = 0
        sck.SendString "QUIT" & vbLf
        lapse = Timer
        Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
            DoTasks
        Loop
        If EventFlag = 3 Or EventFlag = 10 Then
            'sck.Disconnect
            'Err.Raise 8, , EventText
        End If
        sck.Disconnect

        pAuthenticated = False
        
    End If

    EventFlag = 10
End Sub

Private Sub Class_Initialize()
    Randomize
    Set sck = New ISocket
    EventFlag = 0
    pPort = 119
    pBytes = 300000
    pDetails = DetailOptions.None
End Sub

Private Sub Class_Terminate()
    Set sck = Nothing
End Sub

'Public Sub PostMessage(ByVal Subject As String, ByVal Body As String)
'    If pServer <> "" Then
'        If pPostAs <> "" Then
'
'    If Subject <> "" Then
'        If Body <> "" Then
'
'            Dim gid As String
'            Dim lapse As Single
'            Static dc As Boolean
'            dc = False
'
'            If Not sck.Connected Then
'
'                pAuthenticated = False
'                EventFlag = 0
'                If InStr(pServer, ":") > 0 Then
'                    sck.Connect NextArg(pServer, ":"), RemoveArg(pServer, ":")
'                Else
'                    sck.Connect pServer, pPort
'                End If
'                lapse = Timer
'                Do Until (sck.Connected) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If Not sck.Connected Then
'                    sck.Disconnect
'                    Err.Raise 8, , "Could not establish connection."
'                End If
'
'                EventFlag = 0
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    sck.Disconnect
'                    Err.Raise 8, , EventText
'                Else
'                    dc = True
'                End If
'
'            End If
'
'            If (sck.Connected) Then
'
'                        If Not pAuthenticated Then
'                            If pAuthUser <> "" Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO USER " & pAuthUser & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHUSER " & pAuthUser & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'
'                            If (pAuthPass <> "") Then
'                                EventFlag = 0
'                                sck.sendstring "AUTHINFO PASS " & pAuthPass & vblf
'                                lapse = Timer
'                                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                    DoTasks
'                                Loop
'                                If EventFlag = 3 Or EventFlag = 10 Then
'                                    EventFlag = 0
'                                    sck.sendstring "AUTHPASS " & pAuthPass & vblf
'                                    lapse = Timer
'                                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                                        DoTasks
'                                    Loop
'                                End If
'                            End If
'                            If (EventFlag = 3 Or EventFlag = 10) Then
'                                pAuthenticated = False
'                                sck.Disconnect
'                                Err.Raise 8, , EventText
'                            Else
'                                pAuthenticated = True
'                            End If
'                        End If
'
'                EventFlag = 0
'                sck.sendstring "GROUP " & Replace(NewsGroup, ",", "") & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                Dim msgpst As String
'
'                EventFlag = 0
'                sck.sendstring "POST" & vblf
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                msgpst = "Path: " & sck.host & vblf
'                msgpst = msgpst & "From: " & PostAs & vblf
'                msgpst = msgpst & "Newsgroups: " & NewsGroup & vblf
'                msgpst = msgpst & "Subject: " & Subject & vblf
'                msgpst = msgpst & "Message-ID: <" & MessageID & ">" & vblf
'                msgpst = msgpst & "X-Newsreader: NTAdvFTP61.Gorup" & vblf
'                msgpst = msgpst & "X-Advertised: https://www.neotext.org" & vblf
'                msgpst = msgpst & vblf & Body & vblf
'                msgpst = msgpst & "." & vblf
'
'                EventFlag = 0
'                sck.sendstring msgpst
'                lapse = Timer
'                Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                    DoTasks
'                Loop
'                If EventFlag = 3 Or EventFlag = 10 Then
'                    'sck.Disconnect
'                    'Err.Raise 8, , EventText
'                End If
'
'                If dc Then
'                    EventFlag = 0
'                    sck.sendstring "QUIT" & vblf
'                    lapse = Timer
'                    Do Until (EventFlag <> 0) Or (Timer - lapse > 30)
'                        DoTasks
'                    Loop
'                    If EventFlag = 3 Or EventFlag = 10 Then
'                        sck.Disconnect
'                        Err.Raise 8, , EventText
'                    Else
'                        sck.Disconnect
'                    End If
'                End If
'
'            Else
'                sck.Disconnect
'                Err.Raise 8, , EventText
'            End If
'
'        Else
'            Err.Raise 8, "Message body not supplied."
'        End If
'    Else
'        Err.Raise 8, "Message subject not supplied."
'    End If
'
'        Else
'            Err.Raise 8, "Post As unspecified."
'        End If
'    Else
'        Err.Raise 8, "Server unspecified."
'    End If
'
'    EventFlag = 10
'End Sub
Private Sub sck_Connected()
    EventFlag = 1
    pUseID = ""
End Sub

Private Sub sck_DataArriving()
    Dim inLine As String
    Dim inData As String
    If sck.ReadString(inData) Then
    
        Do Until inData = ""
            inLine = Replace(RemoveNextArg(inData, vbLf), vbCr, "")
            Debug.Print "NNTP->: " & inLine
            
            Select Case Left(inLine, 3)
                Case "."
                    If EventFlag = -1 Then
                        EventFlag = 2
                    End If
                Case "221", "222", "501", "502", "480"
                    EventText = inLine & vbLf
                    EventFlag = -1

                Case "200", "340", "240", "381", "281", "205", "211", "223", "441", "335", "411"
                    EventText = inLine
                    If (InStr(inLine, "message-ID") > 0) And (InStr(inLine, "<") > 0) And (InStr(inLine, ">") > 0) Then
                        pUseID = RemoveQuotedArg(inLine, "<", ">")
                    End If
                    EventFlag = 2
                Case Else
                    If EventFlag = -1 Then
                        EventText = EventText & inLine & vbLf
                    Else
                        EventText = inLine
                        EventFlag = 3
                    End If
            End Select
        Loop
    End If
End Sub

Private Sub sck_Disconnected()
    sck.Disconnect
    If EventFlag = 1 Then EventFlag = 4
    pAuthenticated = False
End Sub







Private Sub sck_SendComplete()
                        EventFlag = 5
End Sub
