#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public InstallBackupFile  As String

Public Sub Main()
'%LICENSE%
    
tryit: On Error GoTo catch

    InstallBackupFile = Replace(AppPath & "installer" & DBBackupExt, GetFilePath(Left(AppPath, Len(AppPath) - 1)), GetCurrentAppDataFolder)
    If Not PathExists(GetFilePath(InstallBackupFile), False) Then MkDir GetFilePath(InstallBackupFile)

    
    ExecuteFunction Command


GoTo final
catch: On Error GoTo 0

    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next


On Error GoTo -1
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String) As Boolean
    Dim HasCmd As Boolean
    
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
            
            HasCmd = True
            Select Case Replace(InCommand, " ", "")
                Case "setupreset"
                    UtilityReset bo_AllOptions, True
                Case "setupbackup"
                    UtilityBackup bo_AllOptions, InstallBackupFile, True
                Case "setuprestore"
                    UtilityRestore bo_AllOptions, InstallBackupFile, True
                    Kill InstallBackupFile
                Case "stop"
                    NetStop ServiceName, ServiceFileName
                Case "start"
                    NetStart ServiceName, ServiceFileName
                    
                Case "reset"
                    UtilityReset bo_AllOptions, IIf((InParams = ""), False, CBool(InParams))
                Case "compact"
                    UtilityCompact IIf((InParams = ""), False, CBool(InParams))
                Case "backup"
                    UtilityBackup bo_AllOptions, InParams, IIf((InParams = ""), False, CBool(InParams))
                Case "restore"
                    If PathExists(InParams, True) Then
                        UtilityRestore bo_None, InParams, IIf((InParams = ""), False, CBool(InParams))
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
                    
                    HasCmd = False
            End Select
        Loop
    End If
    
    ExecuteFunction = HasCmd
End Function

Public Sub UtilityCompact(Optional ByVal Silent As Boolean = False)
    Dim dbArchive As New clsArchive
    dbArchive.CompactDatabase Silent
    Set dbArchive = Nothing
End Sub

Public Sub UtilityReset(ByVal Options As Integer, Optional ByVal Silent As Boolean = False, Optional ByVal Compact As Boolean = True)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not Silent Then
        go = (MsgBox("Are you sure you want to reset the RemindMe database?", vbQuestion + vbYesNo, AppName) = vbYes)
    End If
    
    If go Then
        Dim dbArchive As New clsArchive
        
        If Compact Then dbArchive.CompactDatabase
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

Public Sub UtilityBackup(ByVal Options As Integer, ByVal FileName As String, Optional ByVal Silent As Boolean = False)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not PathExists(GetFilePath(FileName), False) Then
        If Not Silent Then MsgBox "The specified path does not exist.", vbExclamation, AppName
    Else
        If Not Silent Then
            go = (MsgBox("Are you sure you want to backup the RemindMe database to" & vbCrLf & FileName & "?", vbQuestion + vbYesNo, AppName) = vbYes)
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

Public Sub UtilityRestore(ByVal Options As Integer, ByVal FileName As String, Optional ByVal Silent As Boolean = False, Optional ByVal Compact As Boolean = True)
    Dim errRet As String
    Dim go As Boolean
    go = True
    
    If Not PathExists(FileName, True) Then
        If Not Silent Then MsgBox "The specified file does not exist.", vbExclamation, AppName
    Else
        If Not Silent Then
            go = (MsgBox("Are you sure you want to restore the RemindMe database from" & vbCrLf & FileName & "?", vbQuestion + vbYesNo, AppName) = vbYes)
        End If
        
        If go Then
            Dim dbArchive As New clsArchive
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


