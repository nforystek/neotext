#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMain"



#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Public Const st_Starting = 0
Public Const st_Transfering = 1
Public Const st_QueuedLocally = 2
Public Const st_QueuedRemotely = 3
Public Const st_Stopped = 4
Public Const st_Finished = 5

Public Const AppName = "Max-FTP"

Public dbSettings As clsSettings
Public MaxEvents As clsEventLog
Public ThreadManager As clsThreadManager

Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)


Public Sub MaxDBError(ByVal Number As Long, ByVal Description As String, ByRef Retry As Boolean)
    'Handles any and all database errors that technically should not occur and
    'immediately halts progress of the application until an ultimatim is met;
    'either it continues to retry by user request and time outs, or shutdown.
    If (Number = -2147467259) Or (Number = 3709) Then
        
        ShutDownMaxFTP True
        
        MsgBox "User permissions insufficient to run Max-FTP, your user account must be part of a group with" & vbCrLf & _
                "permissions to access and modify the Max-FTP database located under the installation directory." & vbCrLf & _
                "Contact your administrator to set up proper user group privlidges for Max-FTP or your account." & vbCrLf & vbCrLf & _
                "(Unable to write to database or folders at: " & AppPath & ")", vbInformation + vbOKOnly, AppName
    
    Else
    
        frmDBError.ShowError Number & " " & Description
                
        Do Until frmDBError.Visible = False
            modCommon.DoTasks
        Loop
        Select Case frmDBError.IsOk
            Case 0
                ShutDownMaxFTP True
            Case 2
                Retry = True
        End Select
    End If
End Sub


Public Sub Main()
    EnableMachinePrivileges
    
'%LICENSE%
    
           
tryit: On Local Error GoTo catch
On Error GoTo catch
    
    'the first sub executed of the app (before all forms is the app model)
    'handles the multiple occurances of itself also command line forwards
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, MaxMainFormCaption)
    If (Not hwnd = 0) Or App.PrevInstance Then
        If Not (Command = "") Then
            MessageQueueAdd MaxFileName, Command
        End If
    Else
        Load frmMain
        
        Set dbSettings = New clsSettings
        Set MaxEvents = New clsEventLog
        Set ThreadManager = New clsThreadManager

        If SetupUser(dbSettings.GetUserLoginName) Then
            InitSystemTray
            LoadFTPClientGraphics
            LoadScheduleGraphics
            ResetScheduleStatus

            If Not ExecuteFunction(Command) Then
                Dim newClient3 As New frmFTPClientGUI
                newClient3.LoadClient
                newClient3.ShowClient
            End If
            
            If dbSettings.GetProfileSetting("TipOfDay") Then
                frmTipOfDay.Show
            End If

            frmMain.timGlobal.enabled = True
            
            Dim dbNet As New clsNetwork
            dbNet.OpenSessionDrives
            Set dbNet = Nothing

        Else
            MsgBox "Error creating or loading User Profile for '" + dbSettings.GetUserLoginName + "'.  You may need to reinstall " & AppName & ".", vbCritical, AppName
        End If

    End If

GoTo final
 On Error GoTo 0
catch:

    If Err Then
        Debug.Print "Error: " & Err.Description
        'MsgBox Err.Description, vbExclamation, App.EXEName
    End If
    
final: On Error Resume Next
    
On Error GoTo -1
End Sub

Private Sub ResetScheduleStatus()
    'this resets the status of all schedule operations that are
    'impropperly left hanging in case the application or service
    'was not actively running as there is also a message status
    'per occurance as it happens rather then an individual state
    If (ProcessRunning(ServiceFileName) = 0) Then
        Dim dbConn As New clsDBConnection
        dbConn.dbQuery "UPDATE Operations SET Status = 'stopped' WHERE Status='running';"
        Set dbConn = Nothing
    End If
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String)
    'this is the main command interpeter that handles all commands
    'applicitable which can come from command line or message queue
    Dim hasCmd As Boolean
    If Trim(CommandLine) <> "" Then
        Dim InParams As String
        Dim InCommand As String
        CommandLine = Replace(Replace(Replace(CommandLine, "+", " "), "%20", " "), "%", " ")
        
        Do Until CommandLine = ""
        
            If InStr(CommandLine, vbCrLf) > 0 Then
                InParams = RemoveNextArg(CommandLine, vbCrLf)
            ElseIf Left(CommandLine, 1) = "/" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "/")
            ElseIf Left(CommandLine, 1) = "-" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "-")
            Else
                InParams = LCase(CommandLine)
                CommandLine = ""
            End If
            
            InCommand = LCase(RemoveNextArg(InParams, " "))
            
            hasCmd = True
            Select Case Replace(InCommand, " ", "")
                Case "file", "open"
                    OpenWebsite InParams, False
                Case "client"
                    If InParams <> "" Then
                        Dim ftpSite1 As New frmFavoriteSite
                        ftpSite1.LoadSite InParams
                    
                        Dim newClient1 As New frmFTPClientGUI
                        newClient1.LoadClient
                        newClient1.ShowClient
                        newClient1.FTPOpenSite ftpSite1
                        
                        Unload ftpSite1
                    Else
                        Dim newClient2 As New frmFTPClientGUI
                        newClient2.LoadClient
                        newClient2.ShowClient
                    End If
                Case "schedule"
                    frmSchManager.Show
                Case "favorites"
                    frmFavorites.Show
                Case "options", "setup"
                    frmSetup.Show
                Case "activeapp", "activecache", "activeappcache"
                    frmActiveCache.ShowForm
                Case "fileassoc"
                    frmFileAssoc.Show
                Case "netdrives", "network"
                    frmNetDrives.Show
                Case "about"
                    frmAbout.Show
                Case "status"
                    frmMain.ProcessScheduleStatus InParams
                Case Else

                    hasCmd = False
                    
            End Select

        Loop
    End If
    
    ExecuteFunction = hasCmd
    
End Function

Sub ShutDownActiveX(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long)
    
    KillTimer hwnd, nIDEvent
    WindowTerminate hwnd
    ShutDownMaxFTP
    
End Sub

Public Sub ShutDownMaxFTP(Optional ByVal DBError As Boolean = False)
    'this is the main opposite of the sub main and handles
    'everything needed to occur when the application closes
    frmMain.timGlobal.enabled = False
    
    TrayIcon False
    
    If (Not DBError) And (Not (dbSettings Is Nothing)) Then CleanUser
    
    Dim cnt As Integer
    Dim frmcnt As Long
    
    cnt = 0
    Do Until cnt > Forms.Count - 1
        If frmMain.timGlobal.enabled Then Exit Sub
        If IsUnloadForm(Forms(cnt)) Then
            Unload Forms(cnt)
        Else
            cnt = cnt + 1
        End If
        
    Loop
    
    cnt = 0
    Do Until cnt > Forms.Count - 1
        If frmMain.timGlobal.enabled Then Exit Sub
        If TypeName(Forms(cnt)) = "frmMain" Then
            cnt = cnt + 1
        Else
            frmcnt = Forms.Count
            Unload Forms(cnt)
            If frmcnt = Forms.Count Then cnt = cnt + 1
        End If
    Loop

    If (Not DBError) Then
        If Not dbSettings Is Nothing Then
            If (Not dbSettings.GetPublicSetting("ServiceSession")) Then
                Dim dbNet As New clsNetwork
                dbNet.CloseSessionDrives
                Set dbNet = Nothing
            End If
        End If
    End If
    
    Set dbSettings = Nothing
    Set MaxEvents = Nothing
    Set ThreadManager = Nothing

    Unload frmMain

End Sub

Public Function IsHiddenForm(ByRef frm As Form) As Boolean
    'this fuction returns whether the form is to be hidden
    'minimized when the application is in the system tray
    Dim frmTypeName As String
    frmTypeName = TypeName(frm)
    If frmTypeName = "frmFTPClientGUI" Then
        IsHiddenForm = (Left(frm.MyDescription, 6) = "Client")
    Else
        IsHiddenForm = ((frmTypeName = "frmSchOperations") Or (frmTypeName = "frmActiveCache"))
    End If

End Function

Public Function IsVisibleClient(ByRef frm As Form) As Boolean
    'this function returns whether the form is a ftp client
    If TypeName(frm) = "frmFTPClientGUI" Then
        IsVisibleClient = (Left(frm.MyDescription, 6) = "Client")
    Else
        IsVisibleClient = False
    End If
End Function

Public Function IsActiveForm(ByRef frm As Form) As Boolean
    'this function returns whether the form is a ftp client
    If TypeName(frm) = "frmFTPClientGUI" Then
        IsActiveForm = (Left(frm.MyDescription, 6) = "Active")
    Else
        IsActiveForm = False
    End If
End Function

Public Function IsUnloadForm(ByRef frm As Form) As Boolean
    'this function returns whether the form is to impose
    'upon the status of keeping the application running
    'if it exists vs forms that are allowed disposable
    Dim frmTypeName As String
    frmTypeName = TypeName(frm)
    
    If frmTypeName = "frmFTPClientGUI" Then
        IsUnloadForm = (Left(frm.MyDescription, 6) = "Client")
    Else
        IsUnloadForm = (frmTypeName = "frmSchOpProperties" Or frmTypeName = "frmSchManager" Or frmTypeName = "frmSchProperties" Or frmTypeName = "frmSchOperations" Or frmTypeName = "frmFavorites" Or frmTypeName = "frmActiveCache" Or frmTypeName = "frmFavoriteSite" Or frmTypeName = "frmFileAssoc" Or frmTypeName = "frmNetConnection" Or frmTypeName = "frmNetDrives" Or frmTypeName = "frmSetup" Or frmTypeName = "frmTransfer")
    End If
    
End Function

Public Function PromptAbortClose(ByVal MsgText As String) As Boolean
    'this is a message hook that is used to catch and determine
    'cancellation of unloading forms based on the users settings
    If dbSettings.GetProfileSetting("PromptAbortClose") Then
        Dim cnt As Integer
        Dim frm As frmFTPClientGUI
        cnt = 0
        Do Until cnt > Forms.Count - 1
            If TypeName(Forms(cnt)) = "frmFTPClientGUI" Then
                Set frm = Forms(cnt)
                If frm.myClient0.ConnectedState Or frm.myClient1.ConnectedState Then
                    PromptAbortClose = (MsgBox(MsgText, vbQuestion + vbYesNo, AppName) = vbYes)
                    Exit Function
                End If
            End If
            cnt = cnt + 1
        Loop
    End If
    PromptAbortClose = True

End Function

Public Function GetTopMostClientGUI() As Form
    'this itterates through all forms and
    'returns the top most open ftp client
    Dim frms As Form
    Dim found As Boolean
    found = False
    For Each frms In Forms
        If TypeName(frms) = "frmFTPClientGUI" Then
            If frms.HasFocus Then
                found = True
                Exit For
            End If
        End If
    Next
    
    If found Then
        Set GetTopMostClientGUI = frms
    Else
        Set GetTopMostClientGUI = Nothing
    End If

End Function

Public Function GetFormByID(ByVal FormID As String) As Form
    'this itterates through all forms and
    'returns the form matching the formID
    'which is a unique client installment

    Dim test As String
    Dim cnt As Integer
    For cnt = 0 To Forms.Count - 1
        If TypeName(Forms(cnt)) = "frmFTPClientGUI" Then
            test = Forms(cnt).MyDescription
            If LCase(test) = LCase(FormID) Then
                Set GetFormByID = Forms(cnt)
                Exit Function
            End If
        End If
    Next
    
    Set GetFormByID = Nothing

End Function

Public Sub GotoHelp()
    'called by the gui when the help is requested
    'and determines the proper location by option
    'of the install and ensures a help is reached
    If IsDocumentationInstalled Then
        OpenWebsite AppPath & "help\index.htm", False
    Else
        OpenWebsite NeoTextWebSite & "/ipub/help/max-ftp"
    End If
End Sub

Public Sub LoadFTPClientGraphics()
    'initialized the imagelists of all
    'button icons for ftp client forms
    frmMain.imgClient(0).ListImages.Clear
    frmMain.imgClient(1).ListImages.Clear
    frmMain.imgClient16(0).ListImages.Clear
    frmMain.imgClient16(1).ListImages.Clear
    
    frmMain.imgClient(0).ImageHeight = GetSkinDimension("toolbarbutton_height")
    frmMain.imgClient(0).ImageWidth = GetSkinDimension("toolbarbutton_width")
    frmMain.imgClient(1).ImageHeight = GetSkinDimension("toolbarbutton_height")
    frmMain.imgClient(1).ImageWidth = GetSkinDimension("toolbarbutton_width")

    frmMain.imgClient16(0).ImageHeight = 16
    frmMain.imgClient16(0).ImageWidth = 16
    frmMain.imgClient16(1).ImageHeight = 16
    frmMain.imgClient16(1).ImageWidth = 16

    LoadButton frmMain.imgClient(1), "uplevel", 1
    LoadButton frmMain.imgClient(1), "back", 2
    LoadButton frmMain.imgClient(1), "forward", 3
    LoadButton frmMain.imgClient(1), "stop", 4
    LoadButton frmMain.imgClient(1), "refresh", 5
    LoadButton frmMain.imgClient(1), "newfolder", 6
    LoadButton frmMain.imgClient(1), "delete", 7
    LoadButton frmMain.imgClient(1), "cut", 8
    LoadButton frmMain.imgClient(1), "copy", 9
    LoadButton frmMain.imgClient(1), "paste", 10
    
    LoadButton frmMain.imgClient16(1), "go", 1
    LoadButton frmMain.imgClient16(1), "close", 2
    LoadButton frmMain.imgClient16(1), "browse", 3
    
    LoadButton frmMain.imgClient(0), "uplevelout", 1
    LoadButton frmMain.imgClient(0), "backout", 2
    LoadButton frmMain.imgClient(0), "forwardout", 3
    LoadButton frmMain.imgClient(0), "stopout", 4
    LoadButton frmMain.imgClient(0), "refreshout", 5
    LoadButton frmMain.imgClient(0), "newfolderout", 6
    LoadButton frmMain.imgClient(0), "deleteout", 7
    LoadButton frmMain.imgClient(0), "cutout", 8
    LoadButton frmMain.imgClient(0), "copyout", 9
    LoadButton frmMain.imgClient(0), "pasteout", 10
    
    LoadButton frmMain.imgClient16(0), "goout", 1
    LoadButton frmMain.imgClient16(0), "closeout", 2
    LoadButton frmMain.imgClient16(0), "browseout", 3

    frmMain.imgFiles.MaskColor = GetSkinColor("list_transparentcolor")
    
    frmMain.imgClient(0).MaskColor = GetSkinColor("toolout_transparentcolor")
    frmMain.imgClient(0).UseMaskColor = True
    frmMain.imgClient(1).MaskColor = GetSkinColor("toolover_transparentcolor")
    frmMain.imgClient(1).UseMaskColor = True
    frmMain.imgClient16(0).MaskColor = GetSkinColor("toolout_transparentcolor")
    frmMain.imgClient16(0).UseMaskColor = True
    frmMain.imgClient16(1).MaskColor = GetSkinColor("toolover_transparentcolor")
    frmMain.imgClient16(1).UseMaskColor = True
            
End Sub

Public Sub LoadScheduleGraphics()
    'initialized the imagelists of all
    'button icons for scheduler forms
    frmMain.imgSchedule(0).ListImages.Clear
    frmMain.imgSchedule(1).ListImages.Clear
    
    frmMain.imgSchedule(0).ImageHeight = GetSkinDimension("toolbarbutton_height")
    frmMain.imgSchedule(0).ImageWidth = GetSkinDimension("toolbarbutton_width")
    frmMain.imgSchedule(1).ImageHeight = GetSkinDimension("toolbarbutton_height")
    frmMain.imgSchedule(1).ImageWidth = GetSkinDimension("toolbarbutton_width")
    
    LoadButton frmMain.imgSchedule(1), "schedule_add", 1
    LoadButton frmMain.imgSchedule(1), "schedule_edit", 2
    LoadButton frmMain.imgSchedule(1), "schedule_delete", 3
    LoadButton frmMain.imgSchedule(1), "schedule_up", 4
    LoadButton frmMain.imgSchedule(1), "schedule_down", 5
    LoadButton frmMain.imgSchedule(1), "schedule_run", 6
    LoadButton frmMain.imgSchedule(1), "schedule_stop", 7
    LoadButton frmMain.imgSchedule(1), "schedule_events", 8
    LoadButton frmMain.imgSchedule(1), "schedule_servicestop", 9
    LoadButton frmMain.imgSchedule(1), "schedule_servicestart", 10

    LoadButton frmMain.imgSchedule(0), "schedule_addout", 1
    LoadButton frmMain.imgSchedule(0), "schedule_editout", 2
    LoadButton frmMain.imgSchedule(0), "schedule_deleteout", 3
    LoadButton frmMain.imgSchedule(0), "schedule_upout", 4
    LoadButton frmMain.imgSchedule(0), "schedule_downout", 5
    LoadButton frmMain.imgSchedule(0), "schedule_runout", 6
    LoadButton frmMain.imgSchedule(0), "schedule_stopout", 7
    LoadButton frmMain.imgSchedule(0), "schedule_eventsout", 8
    LoadButton frmMain.imgSchedule(0), "schedule_servicestopout", 9
    LoadButton frmMain.imgSchedule(0), "schedule_servicestartout", 10

    frmMain.imgSchedule(0).MaskColor = GetSkinColor("schedule_toolout_transparentcolor")
    frmMain.imgSchedule(0).UseMaskColor = True
    frmMain.imgSchedule(1).MaskColor = GetSkinColor("schedule_toolover_transparentcolor")
    frmMain.imgSchedule(1).UseMaskColor = True
        
End Sub

Public Function MapFolder(ByVal RootURL As String, ByVal vURL As String) As String
    'concatenates vURL to the RootURL properly by blind path specifications
    Dim checkURL As New NTAdvFTP61.URL
    If checkURL.GetType(RootURL) = URLTypes.ftp Or checkURL.GetType(RootURL) = URLTypes.HTTP Then
        RootURL = checkURL.GetFolder(RootURL)
        vURL = Replace(vURL, "\", "/")
        
        If Left(vURL, 1) = "/" And Right(RootURL, 1) = "/" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "/" And Right(RootURL, 1) <> "/" Then
            vURL = RootURL & "/" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "/" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "/"
        
    Else
        vURL = Replace(vURL, "/", "\")
        If Left(vURL, 1) = "\" And Right(RootURL, 1) = "\" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "\" And Right(RootURL, 1) <> "\" Then
            vURL = RootURL & "\" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "\" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "\"
    
    End If
    MapFolder = vURL
    Set checkURL = Nothing

End Function







