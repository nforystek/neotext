VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public MyParent As Object
Public MyListItem As ListItem

Public MyDescription As String
Public ClientIndex As Long

Public ThreadState As Integer

Public sOwnerID As String

Public sAction As String

Public SourceURL As String
Public sFolder As String

Public DestURL As String
Public dFolder As String

Public sFileName As String
Public sFileSize As Double
Public sResumeByte As Double
Public dFileSize As Double

Public PerSecond As String
Public TransferTime As Long
Public TransferRate As Double
Public LastReceivedBytes As Double
Private ftpRates(2) As Double

Public WithEvents ftpSClient As NTAdvFTP61.Client
Attribute ftpSClient.VB_VarHelpID = -1
Public WithEvents ftpDClient As NTAdvFTP61.Client
Attribute ftpDClient.VB_VarHelpID = -1

Public Function ThreadOpen()

    If MyListItem Is Nothing Then
        Set MyListItem = MyParent.ListView1.ListItems.Add(, MyDescription, "Starting")
    End If
    MyListItem.SubItems(1) = "0%"
    MyListItem.SubItems(2) = sFileSize
    MyListItem.SubItems(3) = "Unknown"
    MyListItem.SubItems(4) = sFileName
    MyListItem.SubItems(5) = sFolder
    MyListItem.SubItems(6) = dFolder
End Function
Public Function ThreadStatus(ByVal pProgress As String, ByVal pFileBytes As String, ByVal pRate As String)
    On Error Resume Next
    MyListItem.SubItems(1) = pProgress
    MyListItem.SubItems(2) = pFileBytes
    MyListItem.SubItems(3) = pRate
End Function

Public Sub SetSource(ByVal HostURL As String, ByVal Port As Long, ByVal Username As String, ByVal Password As String, ByVal SSL As Boolean)
    ftpSClient.URL = HostURL
    ftpSClient.Port = Port
    ftpSClient.Username = Username
    ftpSClient.Password = Password
    ftpSClient.ImplicitSSL = SSL
    
    SourceURL = HostURL

End Sub

Public Sub SetDestination(ByVal HostURL As String, ByVal Port As Long, ByVal Username As String, ByVal Password As String, ByVal SSL As Boolean)
    ftpDClient.URL = HostURL
    ftpDClient.Port = Port
    ftpDClient.Username = Username
    ftpDClient.Password = Password
    ftpDClient.ImplicitSSL = SSL
    DestURL = HostURL

End Sub

Public Sub SetTransfer(ByVal Action As String, ByVal SourceFolder As String, ByVal DestFolder As String, ByVal Filename As String, ByVal FileSize As Double, ByVal ResumeByte As Double, ByVal DestFileSize As Double)

    sAction = Action
    sFileName = Filename
    sFileSize = FileSize
    dFileSize = DestFileSize
    sResumeByte = ResumeByte
    sFolder = SourceFolder
    dFolder = DestFolder
    
End Sub

Public Sub StartTransfer()
        
    ftpSClient.Connect
    ftpDClient.Connect
    
    ftpSClient.ChangeFolderAbsolute sFolder
    ftpDClient.ChangeFolderAbsolute dFolder

    If (CDbl(sFileSize) > modBitValue.LongBound) Or (CDbl(dFileSize) > modBitValue.LongBound) Or dbSettings.GetProfileSetting("LargeFileMode") Then
        ftpSClient.LargeFileMode = True
        ftpDClient.LargeFileMode = True
    Else
        ftpSClient.LargeFileMode = False
        ftpDClient.LargeFileMode = False
    End If

    Dim fa As clsFileAssoc
    Set fa = New clsFileAssoc
        
    ftpSClient.TransferType = fa.GetTransferType(GetFileExt(sFileName))
    ftpDClient.TransferType = fa.GetTransferType(GetFileExt(sFileName))
   ' ftpSClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(sFileName))
   ' ftpDClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(sFileName))
    Set fa = Nothing
    
    ftpSClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
    ftpSClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
    ftpSClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
    ftpDClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
    ftpDClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
    ftpDClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
    ftpSClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                            IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)
    ftpDClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                            IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)
                                            
    ftpSClient.TransferFile sFileName, ftpDClient, sResumeByte, sFileSize
        
    SetState st_Transfering

End Sub

Public Sub QueueLocally()

    SetState st_QueuedLocally

End Sub

Public Sub LoadTransfer(ByVal ClientDescription As String)
On Error GoTo errcatch
    
    ClientIndex = ThreadManager.AddClients()
    MyDescription = ClientDescription & ClientIndex
    
    Set ftpSClient = ThreadManager.GetClients(ClientIndex).FTPClient1
    Set ftpDClient = ThreadManager.GetClients(ClientIndex).FTPClient2
    
    ftpSClient.timeout = dbSettings.GetProfileSetting("TimeOut")
    ftpSClient.PauseOnStandBy = dbSettings.GetPublicSetting("ServiceStandBy")
    ftpSClient.LogBytes = dbSettings.GetClientSetting("LogFileSize")
    ftpSClient.ImplicitSSL = dbSettings.GetClientSetting("SSL")
    
    ftpDClient.timeout = dbSettings.GetProfileSetting("TimeOut")
    ftpDClient.PauseOnStandBy = dbSettings.GetPublicSetting("ServiceStandBy")
    ftpDClient.LogBytes = dbSettings.GetClientSetting("LogFileSize")
    ftpDClient.ImplicitSSL = dbSettings.GetClientSetting("SSL")
    
    If dbSettings.GetProfileSetting("ConnectionMode") = 1 Then
        ftpSClient.ConnectionMode = "PORT"
        ftpDClient.ConnectionMode = "PORT"
    Else
        ftpSClient.ConnectionMode = "PASV"
        ftpDClient.ConnectionMode = "PASV"
    End If

    modCommon.DoTasks

Exit Sub
errcatch:
    MsgBox Err.Description
    ftpSClient.CancelTransfer
    ftpDClient.CancelTransfer
End Sub

Private Sub Class_Terminate()
    If Not (ThreadManager Is Nothing) Then
        ThreadManager.RemoveClients ClientIndex
    End If
        
    Set MyParent = Nothing
    Set ftpSClient = Nothing
    Set ftpDClient = Nothing

End Sub

Public Sub ftpSClient_FileComplete()

    If LCase(sAction) = "move" Then
    
        ftpSClient.RemoveFile sFileName
    
    End If
    
    ftpSClient.Disconnect
    ftpDClient.Disconnect
    
    SetState st_Finished

End Sub



Public Function GUICancel() As Boolean
    Dim boolCancel As Boolean
    If ThreadState = 2 Or (Not dbSettings.GetProfileSetting("PromptAbortClose")) Then
        boolCancel = True
    Else
        If MsgBox("Do you want to cancel transfering this file - [" & sFileName & "]?", vbQuestion + vbYesNo, "Max-FTP") = vbYes Then
            boolCancel = True
        Else
            boolCancel = False
        End If
    End If

    If boolCancel = True Then
        CancelFileTransfer
    End If

    GUICancel = boolCancel
End Function

Public Sub CancelFileTransfer()

    If Not ftpSClient.ConnectedState() = False Then ftpSClient.Disconnect
    If Not ftpDClient.ConnectedState() = False Then ftpDClient.Disconnect
    SetState st_Stopped
    
    TransferRate = 0
    LastReceivedBytes = 0
End Sub

Private Function SetState(ByVal NewState As Integer)
    ThreadState = NewState
        
    Select Case NewState
        Case st_Starting
            MyListItem.Text = "Starting"
            MyListItem.SmallIcon = "running"
        Case st_Transfering
            MyListItem.Text = "Transfering"
            MyListItem.SmallIcon = "running"
        Case st_QueuedLocally
            MyListItem.Text = "Queued Locally"
            MyListItem.SmallIcon = "paused"
            MyListItem.SubItems(3) = ""
        Case st_QueuedRemotely
            MyListItem.Text = "Queued Remotely"
            MyListItem.SmallIcon = "paused"
            MyListItem.SubItems(3) = ""
        Case st_Stopped
            MyListItem.Text = "Stopped"
            MyListItem.SmallIcon = "stopped"
            MyListItem.SubItems(3) = ""
        Case st_Finished
            MyListItem.Text = "Finished"
            MyListItem.SmallIcon = "stopped"
            MyListItem.SubItems(3) = ""
    End Select

End Function


Private Sub ftpSClient_FileProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    
    Dim percentDone As Integer
    percentDone = CInt((ReceivedBytes / sFileSize) * 100)
        
    If (Timer - TransferTime) >= 1 Then
        TransferTime = Timer
        
        Dim newRate As Double
        newRate = CLng(CDbl(ReceivedBytes - LastReceivedBytes) / 1000)
        
        If ftpRates(0) = 0 And ftpRates(1) = 0 And ftpRates(2) = 0 Then
            ftpRates(0) = newRate
        ElseIf ftpRates(1) = 0 And ftpRates(2) = 0 Then
            ftpRates(1) = newRate
        ElseIf ftpRates(2) = 0 Then
            ftpRates(2) = newRate
            newRate = CLng(CDbl(ftpRates(0) + ftpRates(1) + ftpRates(2)) / 3)
            ftpRates(0) = 0
        ElseIf ftpRates(1) = 0 Then
            ftpRates(1) = newRate
            newRate = CLng(CDbl(ftpRates(0) + ftpRates(1) + ftpRates(2)) / 3)
            ftpRates(2) = 0
        ElseIf ftpRates(0) = 0 Then
            ftpRates(0) = newRate
            newRate = CLng(CDbl(ftpRates(0) + ftpRates(1) + ftpRates(2)) / 3)
            ftpRates(1) = 0
        End If
            
        TransferRate = newRate
        LastReceivedBytes = ReceivedBytes
    End If
        
    If sFileSize > 0 Then

        If TransferRate > 0 Then
            ThreadStatus percentDone & "%", Trim(ReceivedBytes) & " of " & Trim(sFileSize), TransferRate & " KB/Sec"
        Else
            ThreadStatus percentDone & "%", Trim(ReceivedBytes) & " of " & Trim(sFileSize), "Unknown"
        End If
        
    Else
        ThreadStatus "Unknown", Trim(ReceivedBytes) & " of Unknown Size", "Unknown"
    End If
    
End Sub
