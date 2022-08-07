VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Operation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private oID As Long
Private oUserID As Long
Private oScheduleID As Long

Private oAction As String
Private oOverwrite As Boolean
Private oOnlyNewer As Boolean
Private oSubFolders As Boolean

Private oWildCard As String
Private oRenameNew As String

Private oSURL As String
Private oSLogin As String
Private oSPass As String
Private oSPort As Long
Private oSPasv As Boolean
Private oSData As String
Private oSAdap As Long
Private oSSSL As Long

Private oDURL As String
Private oDLogin As String
Private oDPass As String
Private oDPort As Long
Private oDPasv As Boolean
Private oDData As String
Private oDAdap As Long
Private oDSSL As Long

Private ClientIndex As Long

Private oPause As Boolean
Private oFlag As Boolean
Private oElapse As String

Private WithEvents oSClient As NTAdvFTP61.Client
Attribute oSClient.VB_VarHelpID = -1
Private WithEvents oDClient As NTAdvFTP61.Client
Attribute oDClient.VB_VarHelpID = -1

Private oCancelOperation As Boolean

Private Sub WaitForAction(ByRef dbSettings As clsSettings, ByVal MyID As String)
    oElapse = Now
    Do Until (oFlag Or DateDiff("s", oElapse, Now) > dbSettings.GetProfileSetting("TimeOut"))
        If oPause Then oElapse = Now
        DoEvents
        Sleep 1
    Loop
    If Not oFlag Then
        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Error: The servers response has timed out."
    Else
        oFlag = False
    End If
End Sub

Public Property Get PauseTransfers() As Boolean
    PauseTransfers = oPause
End Property
Public Property Let PauseTransfers(ByVal RHS As Boolean)
    oPause = RHS
    If Not oSClient Is Nothing Then
        oSClient.PauseTransfers = RHS
    End If
    If Not oDClient Is Nothing Then
        oDClient.PauseTransfers = RHS
    End If
End Property
Public Property Get CancelOperation() As Boolean
    CancelOperation = oCancelOperation
End Property
Public Property Let CancelOperation(ByVal newValue As Boolean)
    oCancelOperation = newValue
End Property

Public Property Get UserID() As Long
    UserID = oUserID
End Property
Public Property Let UserID(ByVal newValue As Long)
    oUserID = newValue
End Property

Public Property Get ScheduleID() As Long
    ScheduleID = oScheduleID
End Property
Public Property Let ScheduleID(ByVal newValue As Long)
    oScheduleID = newValue
End Property

Public Property Get ID() As Long
    ID = oID
End Property
Public Property Let ID(ByVal newValue As Long)
    oID = newValue
End Property

Public Property Get Action() As String
    Action = oAction
End Property
Public Property Let Action(ByVal newValue As String)
    oAction = newValue
End Property

Public Property Get WaitForScript() As Boolean
    WaitForScript = oOverwrite
End Property

Public Property Let WaitForScript(ByVal newval As Boolean)
    oOverwrite = newval
End Property

Public Property Get Overwrite() As Boolean
    Overwrite = oOverwrite
End Property
Public Property Let Overwrite(ByVal newValue As Boolean)
    oOverwrite = newValue
End Property

Public Property Get ForceTerminate() As Boolean
    ForceTerminate = oOnlyNewer
End Property
Public Property Let ForceTerminate(ByVal newval As Boolean)
    oOnlyNewer = newval
End Property

Public Property Get OnlyNewer() As Boolean
    OnlyNewer = oOnlyNewer
End Property
Public Property Let OnlyNewer(ByVal newValue As Boolean)
    oOnlyNewer = newValue
End Property

Public Property Get NoInterface() As Boolean
    NoInterface = (Not clsService.Controller.Interactive) Or (oSubFolders And clsService.Controller.Interactive)
End Property

Public Property Let NoInterface(ByVal newval As Boolean)
    oSubFolders = Not newval
End Property

Public Property Get SubFolders() As Boolean
    SubFolders = oSubFolders
End Property
Public Property Let SubFolders(ByVal newValue As Boolean)
    oSubFolders = newValue
End Property

Public Property Get mprjScript() As String
    mprjScript = oWildCard
End Property

Public Property Get WildCard() As String
    WildCard = oWildCard
End Property
Public Property Let WildCard(ByVal newValue As String)
    oWildCard = newValue
End Property

Public Property Get Seconds() As Long
    If IsNumeric(oRenameNew) Then
        Seconds = CLng(oRenameNew)
    End If
End Property
Public Property Let Seconds(ByVal newval As Long)
    oRenameNew = CStr(newval)
End Property

Public Property Get RenameNew() As String
    RenameNew = oRenameNew
End Property
Public Property Let RenameNew(ByVal newValue As String)
    oRenameNew = newValue
End Property

Public Property Get SURL() As String
    SURL = oSURL
End Property
Public Property Let SURL(ByVal newValue As String)
    oSURL = newValue
End Property
Public Property Get SLogin() As String
    SLogin = oSLogin
End Property
Public Property Let SLogin(ByVal newValue As String)
    oSLogin = newValue
End Property
Public Property Get SPass() As String
    SPass = oSPass
End Property
Public Property Let SPass(ByVal newValue As String)
    oSPass = newValue
End Property
Public Property Get sport() As Long
    sport = oSPort
End Property
Public Property Let sport(ByVal newValue As Long)
    oSPort = newValue
End Property
Public Property Get SPasv() As Boolean
    SPasv = oSPasv
End Property
Public Property Let SPasv(ByVal newValue As Boolean)
    oSPasv = newValue
End Property
Public Property Get SData() As String
    SData = oSData
End Property
Public Property Let SData(ByVal newValue As String)
    oSData = newValue
End Property
Public Property Get SAdap() As Long
    SAdap = oSAdap
End Property
Public Property Let SAdap(ByVal newValue As Long)
    oSAdap = newValue
End Property
Public Property Get SSSL() As Long
    SSSL = oSSSL
End Property
Public Property Let SSSL(ByVal newValue As Long)
    oSSSL = newValue
End Property

Public Property Get DURL() As String
    DURL = oDURL
End Property
Public Property Let DURL(ByVal newValue As String)
    oDURL = newValue
End Property
Public Property Get DLogin() As String
    DLogin = oDLogin
End Property
Public Property Let DLogin(ByVal newValue As String)
    oDLogin = newValue
End Property
Public Property Get DPass() As String
    DPass = oDPass
End Property
Public Property Let DPass(ByVal newValue As String)
    oDPass = newValue
End Property
Public Property Get DPort() As Long
    DPort = oDPort
End Property
Public Property Let DPort(ByVal newValue As Long)
    oDPort = newValue
End Property
Public Property Get DPasv() As Boolean
    DPasv = oDPasv
End Property
Public Property Let DPasv(ByVal newValue As Boolean)
    oDPasv = newValue
End Property
Public Property Get DData() As String
    DData = oDData
End Property
Public Property Let DData(ByVal newValue As String)
    oDData = newValue
End Property
Public Property Get DAdap() As Long
    DAdap = oDAdap
End Property
Public Property Let DAdap(ByVal newValue As Long)
    oDAdap = newValue
End Property

Public Property Get DSSL() As Long
    DSSL = oDSSL
End Property
Public Property Let DSSL(ByVal newValue As Long)
    oDSSL = newValue
End Property

Public Sub RunOperation()
    On Error Resume Next
    
    oCancelOperation = False
    Dim MyID As String
    Dim tmpName As String
    Dim enc As New NTCipher10.ncode
    Dim dbSettings As New clsSettings
    If dbSettings.LoadUser(oUserID) Or (ScheduleID = -1) Then
        If Not (ScheduleID = -1) Then
            tmpName = dbSettings.CryptKey("", oUserID)
        
            MyID = "Schedule ID: " & oScheduleID & ", Operation ID: " & oID
    
            SetOperationStatus "running"
        Else
            MyID = "Command Line"
        End If
        
        If LCase(Trim(oAction)) = "script" Then
    
            MaxEvents.AddEvent dbSettings, MyID, AppPath & MaxIDEFileName, "Begin " & UCase(Left(oAction, 1)) & LCase(Mid(oAction, 2))
            Dim fPID As Long
            
            If WaitForScript Then
                
                fPID = IIf(NoInterface, vbHide, vbNormalFocus)
    
                If Seconds > 0 Then
                    MaxEvents.AddEvent dbSettings, MyID, AppPath & MaxIDEFileName, "Running " & mprjScript & " Waiting " & Seconds & " Seconds"
                    fPID = MODPROCESS.RunProcess(AppPath & MaxIDEFileName, IIf(NoInterface, "/noui /exec  ", "/exec ") & mprjScript, fPID, Seconds)
                Else
                    MaxEvents.AddEvent dbSettings, MyID, AppPath & MaxIDEFileName, "Running " & mprjScript & " Waiting Indefinite"
                    fPID = MODPROCESS.RunProcess(AppPath & MaxIDEFileName, IIf(NoInterface, "/noui /exec ", "/exec ") & mprjScript, fPID, True)
                    
                End If
                If ForceTerminate Then
                    If ProcessRunning(fPID) > 0 Then KillApp MaxIDEFileName
                End If
            Else
                fPID = vbHide
                MaxEvents.AddEvent dbSettings, MyID, AppPath & MaxIDEFileName, "Running " & mprjScript & " Without Waiting"
                MODPROCESS.RunProcess AppPath & MaxIDEFileName, IIf(NoInterface, "/noui /exec  ", "/exec ") & mprjScript, fPID, False
            End If
    
            MaxEvents.AddEvent dbSettings, MyID, AppPath & MaxIDEFileName, "Finish " & UCase(Left(oAction, 1)) & LCase(Mid(oAction, 2))
        
            If Not (ScheduleID = -1) Then
                
                If Err.Description <> "" Then
                    SetOperationStatus "error"
                Else
                    SetOperationStatus "stopped"
                End If
                
            End If
               
        Else
    
            ClientIndex = ThreadManager.AddClients()

            Dim tmpValue As String
            
            Dim dcID As String
    
            Set oSClient = ThreadManager.GetClients(ClientIndex).FTPClient1
    
            If Not (ScheduleID = -1) Then
                oSClient.timeout = dbSettings.GetProfileSetting("TimeOut")
                oSClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                oSClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                oSClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                oSClient.LargeFileMode = dbSettings.GetProfileSetting("LargeFileMode")
            End If
            oSClient.LogBytes = 0
                                
            If ScheduleID = -1 Then
                oSClient.url = oSURL
                oSClient.Username = oSLogin
                oSClient.Password = oSPass
            Else
                tmpValue = oSURL
                If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                oSClient.url = tmpValue
        
                tmpValue = oSLogin
                If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                oSClient.Username = tmpValue
        
                tmpValue = oSPass
                If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                oSClient.Password = tmpValue
            End If
    
            oSClient.Port = oSPort
            oSClient.ConnectionMode = IIf(oSPasv, "PASV", "PORT")
            If Not (oSData = "") Then oSClient.DataPortRange = oSData
            If Not (oSAdap = 0) Then oSClient.NetAdapter = oSAdap
            oSClient.ImplicitSSL = oSSSL
            
            If Not (ScheduleID = -1) Then
                Select Case LCase(Trim(oAction))
                    Case "copy", "move", "delete"
                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Begin " & UCase(Left(oAction, 1)) & LCase(Mid(oAction, 2)), oSClient, oDClient
                End Select
            End If
            
            oSClient.Connect
    
            If IsCopyOrMove() Then
    
                Set oDClient = ThreadManager.GetClients(ClientIndex).FTPClient2
    
                If Not (ScheduleID = -1) Then

                    oDClient.timeout = dbSettings.GetProfileSetting("TimeOut")
                    oDClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                    oDClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                    oDClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                    oDClient.LargeFileMode = dbSettings.GetProfileSetting("LargeFileMode")
                End If
                oDClient.LogBytes = 0

                If ScheduleID = -1 Then
                    oDClient.url = oDURL
                    oDClient.Username = oDLogin
                    oDClient.Password = oDPass
                Else
                    tmpValue = oDURL
                    If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                    oDClient.url = tmpValue
        
                    tmpValue = oDLogin
                    If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                    oDClient.Username = tmpValue
        
                    tmpValue = oDPass
                    If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
                    oDClient.Password = tmpValue
                End If
                
                oDClient.Port = oDPort
                oDClient.ConnectionMode = IIf(oDPasv, "PASV", "PORT")
                If Not (oDData = "") Then oDClient.DataPortRange = oDData
                If Not (oSAdap = 0) Then oDClient.NetAdapter = oSAdap
                oDClient.ImplicitSSL = oDSSL
                
                oDClient.Connect
            End If
    
            Select Case LCase(Trim(oAction))
                Case "copy", "move", "delete"
                    If (Not oCancelOperation) Then
                        RunOperationRecursive dbSettings, MyID
                    End If
    
                Case "folder"
                    oSClient.MakeFolder oWildCard
                Case "rename"
                    oSClient.Rename oWildCard, oRenameNew
            End Select
    
'            If IsCopyOrMove() Then
'                Do While oSClient.Transfering Or oDClient.Transfering
'                    DoEvents
'                Loop
'            End If
            oSClient.Disconnect
            If IsCopyOrMove() Then
                oDClient.Disconnect
            End If

            If Not (ScheduleID = -1) Then
                Select Case LCase(Trim(oAction))
                    Case "copy", "move", "delete"
                        If (Not oCancelOperation) Then
                            MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Finish " & UCase(Left(oAction, 1)) & LCase(Mid(oAction, 2)), oSClient, oDClient
                        Else
                            MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Canceled " & UCase(Left(oAction, 1)) & LCase(Mid(oAction, 2)), oSClient, oDClient
                        End If
                    Case "folder"
                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "New Folder " & oWildCard, oSClient, oDClient
                    Case "rename"
                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Rename " & oWildCard & " to " & oRenameNew, oSClient, oDClient
                End Select
            End If

    
            If Not (ScheduleID = -1) Then
                
                If Trim(oSClient.GetLastError) <> "" Then
                    SetOperationStatus "error"
                ElseIf Not oDClient Is Nothing Then
                    If Trim(oDClient.GetLastError) <> "" Then
                        SetOperationStatus "error"
                    ElseIf Err.Description <> "" Then
                        SetOperationStatus "error"
                    Else
                        SetOperationStatus "stopped"
                    End If
                ElseIf Err.Description <> "" Then
                    SetOperationStatus "error"
                Else
                    SetOperationStatus "stopped"
                End If
                
            End If
            
            ThreadManager.RemoveClients ClientIndex
        End If
        Set oSClient = Nothing
        Set oDClient = Nothing
    

        Set dbSettings = Nothing
        Set enc = Nothing
    
    End If
    
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Function RunOperationRecursive(ByRef dbSettings As clsSettings, ByVal MyID As String, Optional ByVal oFolder As String = "", Optional ByVal Matched As Boolean = False) As Boolean
    
    Dim sArray() As String
    Dim dArray() As String

    Dim cnt As Long
    Dim skipCnt As Long
    Dim oFileName As String
    Dim oFileSize As String
    Dim oFileDate As String
    Dim dFileSize As String
    Dim dFIleDate As String
    
    Dim attempt As Integer
    Dim fa As clsFileAssoc

    If (Not oFolder = "") And (Not oCancelOperation) Then
        oSClient.ChangeFolderRelative oFolder
        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Change Folder " & oFolder, oSClient, oDClient
        If IsCopyOrMove() Then
            oDClient.ChangeFolderRelative oFolder
            MaxEvents.AddEvent dbSettings, MyID, oDClient.url, "Change Folder " & oFolder, oSClient, oDClient
        End If
    End If

    If (Not oCancelOperation) Then
        Dim ofile As String
        ofile = oSClient.ListContents()
        WaitForAction dbSettings, MyID
        If PathExists(ofile, True) Then
            oSClient.ParseListing ReadFile(ofile), sArray
            Kill ofile
        End If
        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Listing " & oWildCard, oSClient, oDClient
        If IsCopyOrMove() Then
            ofile = oDClient.ListContents()
            WaitForAction dbSettings, MyID
            If PathExists(ofile, True) Then
                oDClient.ParseListing ReadFile(ofile), dArray
                Kill ofile
            End If
            MaxEvents.AddEvent dbSettings, MyID, oDClient.url, "Listing " & oWildCard, oSClient, oDClient
        End If
    End If

    If (Not IsEmptyArray(sArray)) And (Not oCancelOperation) Then
        cnt = 0
        Do Until (cnt > UBound(sArray)) Or oCancelOperation

            oFileName = RemoveNextArg(sArray(cnt), "|")
            oFileSize = RemoveNextArg(sArray(cnt), "|")
            oFileDate = RemoveNextArg(sArray(cnt), "|")
            If (Left(oFileName, 1) = "/") And (oFileSize = "<DIR>" Or oFileSize = "" Or oFileSize = "0") Then
                If oSubFolders Or WildCardMatch(oFileName, oWildCard) Then

                    If IsCopyOrMove() And (Not oCancelOperation) Then
                        If (Not IsFileOnArray(dArray, oFileName)) Then
                            oDClient.MakeFolder Mid(oFileName, 2)
                            MaxEvents.AddEvent dbSettings, MyID, oDClient.url, "Create Folder " & Mid(oFileName, 2), oSClient, oDClient
                        End If
                    End If

                    If (WildCardMatch(oFileName, oWildCard) Or RunOperationRecursive(dbSettings, MyID, Mid(oFileName, 2), WildCardMatch(oFileName, oWildCard) Or Matched)) And (Not oCancelOperation) Then

                        If IsDeleteOrMove() And (Not oCancelOperation) Then
                            oSClient.RemoveFolder Mid(oFileName, 2)
                            MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Remove Folder " & Mid(oFileName, 2), oSClient, oDClient
                        End If
                    Else
                        skipCnt = skipCnt + 1
                    End If
                Else
                    skipCnt = skipCnt + 1
                End If
            Else

                If (WildCardMatch(oFileName, oWildCard) Or Matched) Then
                    If IsCopyOrMove() And (Not oCancelOperation) Then

                        If CDec(CDbl(oFileSize)) > CDec(CDbl(HighBound())) Then
                            Err.Raise 8, App.EXEName, "Transfering for file sizes over " & CDec(CDbl(HighBound)) & " bytes not implementated."
                        End If
                        
                        Set fa = New clsFileAssoc
                        oSClient.TransferType = fa.GetTransferType(GetFileExt(oFileName))
                        oDClient.TransferType = fa.GetTransferType(GetFileExt(oFileName))
                        oSClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(oFileName))
                        oDClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(oFileName))
                        Set fa = Nothing
                        
                        oSClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                        oSClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                        oSClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                        oDClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                        oDClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                        oDClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                        oSClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                                                IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)
                        oDClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                                                IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)
                        
                        If IsFileOnArray(dArray, oFileName, dFileSize, dFIleDate) Then
                            If (CDbl(oFileSize) > LongBound) Or (CDbl(dFileSize) > LongBound) Or dbSettings.GetProfileSetting("LargeFileMode") Then
                                oSClient.LargeFileMode = True
                                oDClient.LargeFileMode = True
                            Else
                                oSClient.LargeFileMode = False
                                oDClient.LargeFileMode = False
                            End If
                        
                            If (oOverwrite And (Not oOnlyNewer)) Then
                                MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfering " & oFileName & " to " & oDClient.Folder, oSClient, oDClient
                                oSClient.TransferFile oFileName, oDClient, , CDbl(oFileSize)
                                WaitForAction dbSettings, MyID
                                MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfered " & oFileName & " to " & oDClient.Folder, oSClient, oDClient

                            ElseIf (oOverwrite And oOnlyNewer) Then
                                If IsDate(oFileDate) And IsDate(dFIleDate) Then
                                    If DateDiff("s", oFileDate, dFIleDate) < 0 Then
                                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfering " & oFileName & " to " & oDClient.Folder, oSClient, oDClient
                                        oSClient.TransferFile oFileName, oDClient, , CDbl(oFileSize)
                                        WaitForAction dbSettings, MyID
                                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfered " & oFileName & " to " & oDClient.Folder, oSClient, oDClient
                                    End If
                                End If
                            End If
                        Else
                            MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfering " & oFileName & " to " & oDClient.Folder, oSClient, oDClient
                            oSClient.TransferFile oFileName, oDClient, , CDbl(oFileSize)
                            WaitForAction dbSettings, MyID
                            MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Transfered " & oFileName & " to " & oDClient.Folder, oSClient, oDClient
                        End If

                    End If

                    If IsDeleteOrMove() And (Not oCancelOperation) Then
                       
                        oSClient.RemoveFile oFileName
                        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Remove File " & oFileName, oSClient, oDClient
                    End If
                Else
                    skipCnt = skipCnt + 1
                End If
            End If

            cnt = cnt + 1
        Loop

    End If


    If (Not oFolder = "") And (Not oCancelOperation) Then
        oSClient.ChangeFolderRelative ".."
        MaxEvents.AddEvent dbSettings, MyID, oSClient.url, "Change Folder .. (To Parent)", oSClient, oDClient
        If IsCopyOrMove() Then
            oDClient.ChangeFolderRelative ".."
            MaxEvents.AddEvent dbSettings, MyID, oDClient.url, "Change Folder .. (To Parent)", oSClient, oDClient
        End If
    End If

    ReDim sArray(0) As String
    ReDim dArray(0) As String

    RunOperationRecursive = (skipCnt = 0)
    
End Function

Private Sub SetOperationStatus(ByVal Status As String)
    
    Dim dbSchedule As New clsDBSchedule

    dbSchedule.SetOperationValue oID, "Status", Status
    If oCancelOperation Then
        dbSchedule.SetOperationValue oID, "LastRun", "Canceled: " & CStr(Now)
    Else
        Select Case Status
            Case "stopped"
                dbSchedule.SetOperationValue oID, "LastRun", "Finished: " & CStr(Now)
            Case "error"
                dbSchedule.SetOperationValue oID, "LastRun", "Error: " & CStr(Now)
        End Select
    End If
    Set dbSchedule = Nothing

    If (ProcessRunning(MaxFileName) <> 0) Or IsDebugger("MaxService") Then
        MessageQueueAdd MaxFileName, "/status " & Trim(oScheduleID) & " " & oID
    End If

End Sub

Private Function IsCopyOrMove() As Boolean
    IsCopyOrMove = ((LCase(Trim(oAction)) = "copy") Or (LCase(Trim(oAction)) = "move"))
End Function
Private Function IsDeleteOrMove() As Boolean
    IsDeleteOrMove = ((LCase(Trim(oAction)) = "delete") Or (LCase(Trim(oAction)) = "move"))
End Function

Public Sub StopOperation()
    oCancelOperation = True
    If Not oDClient Is Nothing Then
        If oDClient.Transfering Then oDClient.CancelTransfer
    End If
    If Not oSClient Is Nothing Then
        If oSClient.Transfering Then oSClient.CancelTransfer
    End If
End Sub

Private Sub Class_Terminate()
    StopOperation
End Sub

Private Sub oDClient_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
    oElapse = Now
    oFlag = True
End Sub

Private Sub oDClient_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    oElapse = Now
End Sub

Private Sub oDClient_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    oElapse = Now
    oFlag = True
End Sub

Private Sub oSClient_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
    oElapse = Now
    oFlag = True
End Sub

Private Sub oSClient_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    oElapse = Now
End Sub

Private Sub oSClient_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    oElapse = Now
    oFlag = True
End Sub
