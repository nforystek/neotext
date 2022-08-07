#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modHosts"
#Const modHosts = -1
Option Explicit
'TOP DOWN

Option Private Module

Public myTimer As String
Public myFetch As String
Public myWanIP As String
Public myNoWeb As Boolean
Public noSave1 As Boolean
Public noSave2 As Boolean
Public myHosts As String
Public myLocal As String
Public myDates As String
Public myWebIP As String

Public LastRun As String
Public LastSet As String
Public LastWan As String

Public Function PreformWanIP()
    If myNoWeb Then
        If Not LastRun = "" Then LastRun = ""
    Else
        If (LastRun = "") Then LastRun = DateAdd("h", -25, Now)
        If (DateDiff("h", LastRun, Now) > 24) Or ((LastWan <> myWanIP) And DateDiff("h", LastRun, Now) > 1) Then
            LastRun = Now
            myWanIP = ServiceTest
            Debug.Print myWanIP
            
            If SubmitWanIP(myWanIP) Then
                LastWan = myWanIP
            Else
                ModifyHostFiles
            End If
        Else
            ModifyHostFiles
        End If
    End If
    
End Function

Public Function ServiceTest() As String
    Dim newip As String
    Dim html As String
    Dim pos As Long
    
    Static setas As String
    
    If ((Trim(myFetch) <> "") And ((LastWan <> setas) Or (setas = ""))) Then

        If PathExists(AppPath & "Ident.sub", True) Then
            newip = NextArg(ReadFile(AppPath & "Ident.sub"), ".")
        Else
            newip = GetMachineName & GetUserLoginName
        End If

        html = modInternet.PostToWebsite(myFetch, "/", "action=wanip&subhost=" & newip)
    
        If IsIpAddress(NextArg(html, vbCrLf)) Then
            newip = NextArg(html, vbCrLf)
            If Not PathExists(AppPath & "Ident.sub", True) Then
                WriteFile AppPath & "Ident.sub", RemoveArg(html, vbCrLf)
            End If
        Else
            pos = InStr(html, ".")
            Do Until (pos = 0) Or IsIpAddress(newip)
                newip = Mid(html, pos - 3, 15)
                Do While Not IsNumeric(Right(newip, 1)) And (newip <> "")
                    newip = Left(newip, Len(newip) - 1)
                Loop
                pos = InStr(pos + 1, html, ".")
            Loop
    
            If Not IsIpAddress(newip) Then
    
                Dim temp As String
                temp = html
                temp = Mid(temp, InStr(temp, "<span id=""ipa"">") + 4)
    
                newip = " "
                Do Until temp = "" Or IsIpAddress(Left(newip, Len(newip) - 1))
                    RemoveNextArg temp, "<span>"
                    newip = newip & RemoveNextArg(temp, "</span>") & "."
                Loop
                If newip <> "" Then newip = Trim(Left(newip, Len(newip) - 1))
            End If
        
            If Not IsIpAddress(newip) And Not newip = "" Then
                newip = myWanIP
            End If

        End If
    End If

    If Not setas = newip And IsIpAddress(newip) Then
        setas = newip
        ServiceTest = newip
    Else
        ServiceTest = setas
    End If
End Function

Public Function FailSend(Optional ByVal testBack As String = "") As Boolean
    If (testBack <> "") Then
        If Not IsGuid(testBack) Then
            myTimer = Now
            FailSend = True
        Else
            FailSend = False
        End If
    ElseIf myTimer = "" Then
        FailSend = False
    Else
        If DateDiff("s", Now, myTimer) > 1441 Then
            FailSend = False
            myTimer = ""
        Else
            FailSend = True
        End If
    End If
End Function

Public Function SubmitWanIP(ByVal WanIP As String) As Boolean
    If Not FailSend() And (WanIP <> "") Then
            
        Dim subhost As String
        If PathExists(AppPath & "Ident.sub", True) Then
            subhost = NextArg(ReadFile(AppPath & "Ident.sub"), ".")
        Else
            subhost = GetMachineName & GetUserLoginName
        End If
        
        Dim cnt As Long
        Dim max As Long
        Dim newhost As String
        max = Len(subhost)
        If (max > 13) Then max = 13
        For cnt = 1 To max
            If InStr("abcdefghijklmnopqrstuvwxyz", Mid(subhost, cnt, 1)) > 0 Or InStr(UCase("abcdefghijklmnopqrstuvwxyz"), Mid(subhost, cnt, 1)) > 0 Then
                newhost = newhost & Mid(subhost, cnt, 1)
            End If
        Next
        subhost = newhost
            
        
        If (subhost <> "") Then

            Dim guid1st As String
            Dim guidOut As String
            guid1st = GUID()

            guidOut = modInternet.PostToWebsite(myWebIP, "/", "action=start&subhost=" & subhost & "&hash=" & modInternet.URLEncode(guid1st), , , True)
            If Not FailSend(guidOut) Then
                guidOut = modInternet.PostToWebsite(myWebIP, "/", "action=" & modInternet.URLEncode(WanIP) & "&subhost=" & subhost & "&hash=" & modInternet.URLEncode(guidOut), , , True)
                If Not FailSend(guidOut) Then
                    guidOut = modInternet.PostToWebsite(myWebIP, "/", "action=final&subhost=" & subhost & "&hash=" & modInternet.URLEncode(guidOut), , , True)
                    If guidOut = guid1st Then
                        SubmitWanIP = True
                    Else
                        SubmitWanIP = False
                    End If
                Else
                    SubmitWanIP = False
                End If
            Else
                SubmitWanIP = False
            End If
        Else
            SubmitWanIP = False
        End If
        
    Else
        SubmitWanIP = False
    End If
    myTimer = Now

End Function

Private Function CheckDate()
    If PathExists(App.Path & "\WanIP.txt", True) Then
        CheckDate = CheckDate & GetFileDate(App.Path & "\WanIP.txt")
    Else
        CheckDate = "#INVALID#"
    End If
    If PathExists(App.Path & "\Local.txt", True) Then
        CheckDate = CheckDate & GetFileDate(App.Path & "\Local.txt")
    Else
        CheckDate = "#INVALID#"
    End If
    If PathExists(App.Path & "\Hosts.txt", True) Then
        CheckDate = CheckDate & GetFileDate(App.Path & "\Hosts.txt")
    Else
        CheckDate = "#INVALID#"
    End If
End Function

Public Function ModifyHostFiles() As Boolean
    If myDates <> CheckDate() Then
         myDates = ""
         Dim preText As String
         If PathExists(App.Path & "\WanIP.txt", True) Then
             preText = Replace(Replace(Replace(ReadFile(App.Path & "\WanIP.txt"), "|", vbCrLf), vbTab & vbTab, vbTab), vbTab, "    ") & vbCrLf & vbCrLf
         End If
         If PathExists(App.Path & "\Hosts.txt", True) Then
             myHosts = Replace(Replace(Replace(ReadFile(App.Path & "\Hosts.txt"), "|", vbCrLf), vbTab & vbTab, vbTab), vbTab, "   ") & vbCrLf & vbCrLf
        End If
         If PathExists(App.Path & "\Local.txt", True) Then
             myLocal = Replace(Replace(Replace(ReadFile(App.Path & "\Local.txt"), "|", vbCrLf), vbTab & vbTab, vbTab), vbTab, "   ") & vbCrLf & vbCrLf
        End If
        
        Dim outtext As String
        Dim count As Long
        Dim ip As Variant
        Dim c As Collection
        Set c = GetPortIP
        
        For Each ip In c
            count = count + 1
            If InStr(preText, "%localip" & Trim(count) & "%") > 0 Then
                preText = Replace(preText, "%localip" & Trim(count) & "%", CStr(ip))
                myLocal = Replace(myLocal, "%localip" & Trim(count) & "%", CStr(ip))
                myHosts = Replace(myHosts, "%localip" & Trim(count) & "%", CStr(ip))
            ElseIf InStr(preText, "%localip%") > 0 Then
                preText = Replace(preText, "%localip%", CStr(ip))
                myLocal = Replace(myLocal, "%localip%", CStr(ip))
                myHosts = Replace(myHosts, "%localip%", CStr(ip))
            End If
            If myWanIP = ip Then
                myNoWeb = True
            End If
        Next

        If Not myWanIP = "" Then
           If InStr(preText, "%localip0%") > 0 Then
               preText = Replace(preText, "%localip0%", CStr(myWanIP))
               myLocal = Replace(myLocal, "%localip0%", CStr(myWanIP))
               myHosts = Replace(myHosts, "%localip0%", CStr(myWanIP))
           ElseIf InStr(preText, "%localip%") > 0 Then
               preText = Replace(preText, "%localip%", CStr(myWanIP))
               myLocal = Replace(myLocal, "%localip%", CStr(myWanIP))
               myHosts = Replace(myHosts, "%localip%", CStr(myWanIP))
           End If
        End If

        myLocal = Replace(myLocal & vbCrLf & preText & vbCrLf, vbCrLf & vbCrLf, vbCrLf)
        myHosts = Replace(myHosts & vbCrLf & preText & vbCrLf, vbCrLf & vbCrLf, vbCrLf)
        
        If Replace(ReadFile(SysPath & "DRIVERS\ETC\HOSTS"), vbCrLf, "") <> Replace(myHosts, vbCrLf, "") Then
            myHosts = Replace(myHosts, vbCrLf & vbCrLf, vbCrLf)
            If PathExists(App.Path & "\Hosts.txt", True) Then
                WriteFile SysPath & "DRIVERS\ETC\HOSTS", myHosts
            End If
        End If
                
        If Replace(ReadFile(SysPath & "DRIVERS\ETC\LMHOSTS.SAM"), vbCrLf, "") <> Replace(myLocal, vbCrLf, "") Then
            myLocal = Replace(myLocal, vbCrLf & vbCrLf, vbCrLf)
            If PathExists(App.Path & "\Local.txt", True) Then
                WriteFile SysPath & "DRIVERS\ETC\LMHOSTS.SAM", myLocal
            End If
        End If
         
         myDates = CheckDate
     End If
End Function

