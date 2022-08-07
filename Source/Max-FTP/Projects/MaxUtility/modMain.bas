Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public dbSettings As New clsSettings
Public InstallBackupFile  As String
Public Const AppName = "Max-FTP Database Utility"

Public Sub MaxDBError(ByVal Number As Long, ByVal Description As String, ByRef Retry As Boolean)
    If dbSettings.GetUserLoginName = "" Then
        Retry = True
    Else
        If (Number = -2147467259) Or (Number = 3709) Then
            
                MsgBox "User permissions insufficient to run Max-FTP, your user account must be part of a group with" & vbCrLf & _
                    "permissions to access and modify the Max-FTP database located under the installation directory." & vbCrLf & _
                    "Contact your administrator to set up proper user group privlidges for Max-FTP or your account." & vbCrLf & vbCrLf & _
                    "(Unable to write to sub folders or database: " & DatabaseFilePath & ")", vbInformation + vbOKOnly, AppName
            
            End
        Else
        
            frmDBError.ShowError Description
                    
            Do Until frmDBError.Visible = False
                DoTasks
            Loop
            Select Case frmDBError.IsOk
                Case 0
                    End
                Case 2
                    Retry = True
            End Select
        End If
    End If
End Sub

Public Sub Main()
    '%LICENSE%
    
    
tryit: On Error GoTo catch
    EnableMachinePrivileges
    
    InstallBackupFile = Replace(AppPath & "installer" & MaxDBBackupExt, GetFilePath(Left(AppPath, Len(AppPath) - 1)), GetAllUsersAppDataFolder)
    If Not PathExists(GetFilePath(InstallBackupFile), False) Then MkDir GetFilePath(InstallBackupFile)
   
    ExecuteFunction IIf(Command = "", "/utility", Command)
    
GoTo final
catch: On Error GoTo 0

    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next

On Error GoTo -1
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String)
    Dim hasCmd As Boolean
    Dim tmp As String
        
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
                Case "utility"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Then
                        Set dbSettings = Nothing
                        
                        Load frmUtility
                        frmUtility.ShowForm
                        hasCmd = False
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "open"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Then
                        Set dbSettings = Nothing
                        If PathExists(InParams, True) Then
                            Load frmUtility
                            frmUtility.Text1.Text = InParams
                            frmUtility.Text2.Text = InParams
                            hasCmd = False
                        End If
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "reset", "sreset"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityReset bo_AllOptions, (InCommand = "sreset")
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "backup", "sbackup"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityBackup bo_AllOptions, InParams, (InCommand = "sbackup")
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "restore", "srestore"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        If PathExists(InParams, True) Then
                            UtilityRestore bo_None, InParams, (InCommand = "srestore")
                        End If
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "compact", "scompact"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityCompact (InCommand = "scompact")
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "setupreset"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityReset bo_AllOptions, True
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "setupbackup"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityBackup bo_AllOptions, InstallBackupFile, True
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "setuprestore"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        Set dbSettings = Nothing
                        UtilityRestore bo_AllOptions, InstallBackupFile, True
                        Kill InstallBackupFile
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "setupinitial"
                    If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or dbSettings.NoUsers Or (dbSettings.GetUserLoginName = "") Then
                        tmp = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Max-FTP", "InstallUser", "")
                        If Not (tmp = "") Then
                            SetupUser tmp
                            If dbSettings.LoadUser(tmp) Then
                                If (dbSettings.CurrentUserAccessRights = AccessRights.ar_Administrator) Then
                                    dbSettings.SetPublicSetting "ServiceNetwork", CBool(InParams)
                                End If
                            End If
                        End If
                    Else
                        MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                    End If
                Case "stop"
                    If dbSettings.GetPublicSetting("ServiceAllowAny") Or dbSettings.NoUsers Then
                        Set dbSettings = Nothing
                        NetStop IIf((InParams = ""), MaxServiceName, InParams), IIf((InParams = ""), ServiceFileName, "")
                    Else
                        If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                        If (dbSettings.CurrentUserAccessRights = ar_Administrator) Then
                            Set dbSettings = Nothing
                            NetStop IIf((InParams = ""), MaxServiceName, InParams), IIf((InParams = ""), ServiceFileName, "")
                        Else
                            MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                        End If
                    End If
                Case "start"
                    If dbSettings.GetPublicSetting("ServiceAllowAny") Or dbSettings.NoUsers Then
                        Set dbSettings = Nothing
                        NetStart IIf((InParams = ""), MaxServiceName, InParams), IIf((InParams = ""), ServiceFileName, "")
                    Else
                        If (dbSettings.GetUserLoginName <> "") Then dbSettings.LoadUser dbSettings.GetUserLoginName
                        If (dbSettings.CurrentUserAccessRights = ar_Administrator) Then
                            Set dbSettings = Nothing
                            NetStart IIf((InParams = ""), MaxServiceName, InParams), IIf((InParams = ""), ServiceFileName, "")
                        Else
                            MsgBox "Your Max-FTP access rights are not sufficient to perform this action.", vbInformation + vbOKOnly, AppName
                        End If
                    End If
                Case "closeall"
                    Load frmWarning
                    If Not (frmWarning.Label2.Caption = "") Then
                        frmWarning.Show
                        Do While frmWarning.Visible
                            frmWarning.TestApps
                            DoTasks
                        Loop
                    End If
                    Unload frmWarning
                    
                Case Else
                    hasCmd = False
                    
            End Select

        Loop
    End If

    Set dbSettings = Nothing
    
    ExecuteFunction = hasCmd

End Function

Public Sub UtilityCompact(Optional ByVal Silent As Boolean = False)
    Dim dbArchive As New clsArchive
    dbArchive.CompactDatabase Silent
    Set dbArchive = Nothing
End Sub

Public Sub UtilityReset(ByVal Options As Long, Optional ByVal Silent As Boolean = False, Optional ByVal Compact As Boolean = True)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not Silent Then
        go = (MsgBox("Are you sure you want to reset the Max-FTP database?", vbQuestion + vbYesNo, AppName) = vbYes)
    End If
    
    If go Then
    
        Dim dbArchive As New clsArchive
        dbArchive.ResetDatabase Options, Silent, True
        If Compact Then dbArchive.CompactDatabase
        
        If Not Silent Then
            If errRet = "" Then
                MsgBox "Database reset complete.", vbInformation, AppName
            Else
                MsgBox "Error: " & errRet, vbInformation, AppName
            End If
        End If
        Set dbArchive = Nothing
    End If
End Sub

Public Sub UtilityBackup(ByVal Options As Long, ByVal FileName As String, Optional ByVal Silent As Boolean = False)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not PathExists(GetFilePath(FileName), False) Then
        If Not Silent Then MsgBox "The specified path does not exist.", vbExclamation, AppName
    Else
        If Not Silent Then
            go = (MsgBox("Are you sure you want to backup the Max-FTP database to" & vbCrLf & FileName & "?", vbQuestion + vbYesNo, AppName) = vbYes)
        End If
        
        If go Then
            Dim dbArchive As New clsArchive
            dbArchive.ExportToBackup Options, FileName, Silent
            
            If Not Silent Then
                If errRet = "" Then
                    MsgBox "Database backup complete.", vbInformation, AppName
                Else
                    MsgBox "Error: " & errRet, vbInformation, AppName
                End If
            End If
            Set dbArchive = Nothing
        End If
    End If
End Sub

Public Sub UtilityRestore(ByVal Options As Long, ByVal FileName As String, Optional ByVal Silent As Boolean = False, Optional ByVal Compact As Boolean = True)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not PathExists(FileName, True) Then
        If Not Silent Then MsgBox "The specified file does not exist.", vbExclamation, AppName
    Else
        If Not Silent Then
            go = (MsgBox("Are you sure you want to restore the Max-FTP database from" & vbCrLf & FileName & "?", vbQuestion + vbYesNo, AppName) = vbYes)
        End If
        
        If go Then
            Dim dbArchive As New clsArchive
            If Compact Then dbArchive.CompactDatabase
            dbArchive.ImportFromBackup Options, FileName, Silent
            If Compact Then dbArchive.CompactDatabase
            
            If Not Silent Then
                If errRet = "" Then
                    MsgBox "Database restore complete.", vbInformation, AppName
                Else
                    MsgBox "Error: " & errRet, vbInformation, AppName
                End If
            End If
            Set dbArchive = Nothing
        End If
    End If
End Sub

