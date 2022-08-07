Attribute VB_Name = "modMain"
#Const [True] = -1
#Const [False] = 0
#Const modMain = -1
Option Explicit
'TOP DOWN

Option Compare Binary

#If Not DLL = -1 Then
    Public wService As clsIdent
#End If

Public WinsockControl As Boolean

Public Sub Log(ByVal Text As String)

    Dim FileNum As Integer
    FileNum = FreeFile
    Open AppPath & "Ident.log" For Append As #FileNum
        Print #FileNum, Text
    Close #FileNum

End Sub

Public Sub Main()
    
    On Local Error GoTo 0
    #If Not DLL = -1 Then
        Set wService = New clsIdent
        wService.Controller.ServiceName = "Identd"
        wService.Controller.Account = ".\LocalSystem"
        wService.Controller.Password = "*"
        wService.Controller.DisplayName = "Ident Service"
        wService.Controller.Description = "UDP self aware service listing on TCP port 113 waiting to respond to inbound identity requests. (RFC 1413 - Identification Protocol)"
        wService.Controller.AutoStart = True
        wService.Controller.Interactive = False
    
        Select Case Command
            Case "install"
                wService.Controller.Install
                Set wService = Nothing
                End
            Case "uninstall"
                wService.Controller.Uninstall
                Set wService = Nothing
                End
            Case Else
                'If IsCompiled Then
                    wService.Controller.StartService
                'End If
    #End If
                Load frmIdent
                
                
    #If Not DLL = -1 Then
        End Select
    #End If
End Sub

Public Sub StopService()
    'frmIdent.Shutdown
   ' Do Until Forms.count = 1
        frmIdent.Shutdown
        Unload frmIdent
    'Loop
    #If Not DLL = -1 Then
        Set wService = Nothing
    #End If
End Sub

Public Sub LoadINI()
    Dim inData As String
    Dim inLine As String
    #If DLL = -1 Then
        inData = "[Settings]" & vbCrLf & _
                "IdleTimeout = 180" & vbCrLf & _
                "WinsOnlySys = false" & vbCrLf & _
                "IncludeComp = false" & vbCrLf & _
                "AdapterAddr = every" & vbCrLf & _
                "StandsAlone = true" & vbCrLf & _
                "StandOnNeed = false" & vbCrLf & _
                "ProcessNeed = ident" & vbCrLf & _
                "UserNameSID = false" & vbclrf & _
                "ServicePort =" & vbCrLf
    #Else
        If Not PathExists(AppPath & "Ident.ini", True) And PathExists(AppPath & "Ident.new", True) Then
            Name AppPath & "Ident.new" As AppPath & "Ident.ini"
        End If
        inData = ReadFile(AppPath & "Ident.new") & vbCrLf & ReadFile(AppPath & "Ident.ini")
    #End If
    Do Until inData = ""
        inLine = RemoveNextArg(inData, vbCrLf)
        inLine = RemoveNextArg(inLine, "//")
        inLine = Replace(inLine, vbTab, "    ")
        If (Not (inLine = "")) And (InStr(inLine, "=") > 0) Then
            Select Case LCase(NextArg(inLine, "="))
                Case "idletimeout"
                    If IsNumeric(RemoveArg(inLine, "=")) Then
                        frmIdent.IdleTimeout = CLng(CDbl(RemoveArg(inLine, "=")))
                    End If
                Case "winsonlysys"
                    Select Case LCase(RemoveArg(inLine, "="))
                        Case "true", 1, "1", "-1", "yes", "on"
                            frmIdent.WinsOnlySys = True
                        Case "false", 0, "0", "no", "off", ""
                            frmIdent.WinsOnlySys = False
                    End Select
                Case "includecomp"
                    Select Case LCase(RemoveArg(inLine, "="))
                        Case "true", 1, "1", "-1", "yes", "on"
                            frmIdent.IncludeComp = True
                        Case "false", 0, "0", "no", "off", ""
                            frmIdent.IncludeComp = False
                    End Select
                Case "standsalone"
                    Select Case LCase(RemoveArg(inLine, "="))
                        Case "true", 1, "1", "-1", "yes", "on"
                            frmIdent.StandsAlone = True
                        Case "false", 0, "0", "no", "off", ""
                            frmIdent.StandsAlone = False
                    End Select
                Case "standonneed"
                    Select Case LCase(RemoveArg(inLine, "="))
                        Case "true", 1, "1", "-1", "yes", "on"
                            frmIdent.StandOnNeed = True
                        Case "false", 0, "0", "no", "off", ""
                            frmIdent.StandOnNeed = False
                    End Select
                Case "processneed"
                    If Not frmIdent.ExistsInSetting(frmIdent.ProcessNeed, RemoveArg(inLine, "=")) Then
                        frmIdent.ProcessNeed.Add RemoveArg(inLine, "="), RemoveArg(inLine, "=")
                    End If
                Case "adapteraddr"
                    If Not frmIdent.ExistsInSetting(frmIdent.AdapterAddr, RemoveArg(inLine, "=")) Then
                        frmIdent.AdapterAddr.Add LCase(RemoveArg(inLine, "=")), LCase(RemoveArg(inLine, "="))
                    End If
                Case "usernamesid"
                    Select Case LCase(RemoveArg(inLine, "="))
                        Case "true", 1, "1", "-1", "yes", "on"
                            frmIdent.UserNameSID = True
                        Case "false", 0, "0", "no", "off", ""
                            frmIdent.UserNameSID = False
                    End Select
                Case "serviceport"
                    inLine = RemoveArg(inLine, "=")
                    Do Until inLine = ""
                        frmIdent.ServicePort.Add RemoveNextArg(inLine, ",")
                    Loop

            End Select
        End If
    Loop
End Sub

