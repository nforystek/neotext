VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Executes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

'following are in order
Public backup As String 'restore is rollback
Public Remove As String 'apply is rollback
'midway
Public Restore As String 'backup is rollback
Public Initial As String 'remove is rollback
'this is not as ordered
Public Service As String 'identity if service

Private Stage As Integer

Public Function NextExecute(ByRef InfoText As String, Optional ByVal Rollback As Boolean = False) As Boolean
    Dim cmd As String
    Dim arg As String
    If Program.Installed Or Rollback Then
        If Stage > 8 Or Stage = 0 Then Stage = 8
        Select Case Stage
            Case 4
                NextExecute = False
            Case 5
                cmd = NextArg(Service, "|")
                arg = "uninstall"
                NextExecute = True
            Case 6
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Service, "|"), True) Then
                    InfoText = "Removing service: " & RemoveArg(Service, "|")
                End If
                NextExecute = True
            Case 7
                cmd = NextArg(Remove, "|")
                arg = RemoveArg(Remove, "|")
                NextExecute = True
            Case 8
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Remove, "|"), True) Then
                    InfoText = "Removal specifics: " & NextArg(Remove, "|")
                End If
                NextExecute = True
        End Select
        Stage = Stage - 1
    Else
        Select Case Stage
            Case 0
                If Program.Restore Then
                    If SimSilence = ZeroUI Then
                        InfoText = "Silent install, skipping restore point."
                    Else
                        If PathExists(GetSystem32Folder & "wbem\wmic.exe", True) Then
                            InfoText = "Creating restore point: Before " & IIf(Program.Installed, "uninstallation", "installation") & " of " & Program.Display
                        Else
                            InfoText = "Unable to create restore point: wmic.exe not found"
                        End If
                    End If
                End If
                NextExecute = True
            Case 1
                If Not SimSilence = ZeroUI Then
                    If Program.Restore Then
                        cmd = SearchPath("powershell.*", -1, GetSystem32Folder & "WindowsPowershell", FirstOnly)
                        If cmd <> "" Then
                        
                            RunProcess cmd, "-NoProfile -ExecutionPolicy Bypass -Command ""& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -NoExit -Command ""Checkpoint-Computer -Description ""Before " & IIf(Program.Installed, "uninstallation", "installation") & " of " & Program.Display & """ -RestorePointType ""MODIFY_SETTINGS""; "" ' "" -Verb RunAs}"""
                        
                        ElseIf PathExists(GetSystem32Folder & "wbem\wmic.exe", True) And Not IsDebugger Then
                            RunProcess GetSystem32Folder & "wbem\Wmic.exe", "/Namespace:\\root\default Path SystemRestore Call CreateRestorePoint ""Before " & IIf(Program.Installed, "uninstallation", "installation") & " of " & Program.Display & """, 100, 12", vbNormal, True
                        End If
                        cmd = ""
                        
                    End If
                End If
                NextExecute = True
            Case 2
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(backup, "|"), True) Then
                    InfoText = "Preforming backup: " & NextArg(backup, "|")
                Else
                    Restore = ""
                End If
                NextExecute = True
            Case 3
    
                cmd = NextArg(backup, "|")
                arg = RemoveArg(backup, "|")
                NextExecute = True
            Case 4
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Service, "|"), True) Then
                    InfoText = "Removing service: " & RemoveArg(Service, "|")
                End If
                NextExecute = True
            Case 5
                cmd = NextArg(Service, "|")
                arg = "uninstall"
                NextExecute = True
            Case 6
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Remove, "|"), True) Then
                    InfoText = "Removal specifics: " & NextArg(Remove, "|")
                End If
                NextExecute = True
            Case 7
                cmd = NextArg(Remove, "|")
                arg = RemoveArg(Remove, "|")
                NextExecute = True
            Case 8
                NextExecute = False
            Case 9
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Service, "|"), True) Then
                    InfoText = "Installing service: " & NextArg(Service, "|")
                End If
                NextExecute = True
            Case 10
                cmd = NextArg(Service, "|")
                arg = "install"
                NextExecute = True
            Case 11
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Restore, "|"), True) Then
                    InfoText = "Preforming restore: " & NextArg(Restore, "|")
                End If
                NextExecute = True
            Case 12
                cmd = NextArg(Restore, "|")
                arg = RemoveArg(Restore, "|")
                NextExecute = True
            Case 13
                If PathExists(InstallLocation & "\" & Program.AppValue & "\" & NextArg(Initial, "|"), True) Then
                    InfoText = "Installing specifics: " & NextArg(Initial, "|")
                End If
                NextExecute = True
            Case 14
                cmd = NextArg(Initial, "|")
                arg = RemoveArg(Initial, "|")
        End Select
        Stage = Stage + 1
    End If
    If cmd <> "" Then
        
        If PathExists(InstallLocation & "\" & Program.AppValue & "\" & cmd, True) Then
            RunProcess InstallLocation & "\" & Program.AppValue & "\" & cmd, Trim(arg), vbHide, True
        Else
            InfoText = ""
        End If
    End If
End Function
