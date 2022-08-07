Attribute VB_Name = "modUser"

#Const modUser = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Private Const REG_FORCE_RESTORE As Long = 8&
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"

Private Const REG_OPTION_VOLATILE = 1

Private Type LUID
  lowpart As Long
  highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type

Private Declare Function AdjustTokenPrivileges Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "ADVAPI32.DLL" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "ADVAPI32.DLL" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


Public Function EnableMachinePrivileges() As Boolean
  On Error Resume Next
  Dim seName As String
  seName = SE_MACHINE_ACCOUNT_NAME
  Dim p_lngRtn As Long
  Dim p_lngToken As Long
  Dim p_lngBufferLen As Long
  Dim p_typLUID As LUID
  Dim p_typTokenPriv As TOKEN_PRIVILEGES
  Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

  p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
  If p_lngRtn = 0 Then
    Exit Function
  End If
  If Err.LastDllError <> 0 Then
    Exit Function
  End If
  p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
  If p_lngRtn = 0 Then
    Exit Function
  End If
  p_typTokenPriv.PrivilegeCount = 1
  p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
  p_typTokenPriv.Privileges.pLuid = p_typLUID
  EnableMachinePrivileges = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

Public Function SetupUser(ByVal Username As String) As Boolean
    
        
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    dbConn.rsQuery rs, "SELECT * FROM Users;"
    If rsEnd(rs) Then
        dbConn.rsQuery rs, "DELETE * FROM PublicSettings;"
        
        dbConn.rsQuery rs, "INSERT INTO PublicSettings (ServiceInstallGUID) VALUES ('" & Replace(modGuid.GUID, "-", "") & "');"
        dbSettings.DefaultAccess = ar_Administrator
    Else

        dbConn.rsQuery rs, "SELECT * FROM PublicSettings;"
        If rsEnd(rs) Then
            dbConn.rsQuery rs, "INSERT INTO PublicSettings (ServiceInstallGUID) VALUES ('" & Replace(modGuid.GUID, "-", "") & "');"
            dbSettings.DefaultAccess = ar_Administrator
        Else
            dbSettings.DefaultAccess = ar_CommonUser
        End If
    
   End If
    
    dbConn.rsQuery rs, "SELECT * FROM Users WHERE UserName='" & Replace(Username, "'", "''") & "';"
    If rsEnd(rs) Then
            
        dbConn.rsQuery rs, "INSERT INTO Users (UserName, AccessRights) VALUES ('" & Replace(Username, "'", "''") & "'," & dbSettings.DefaultAccess & ");"
        
        If dbConn.rsQuery(rs, "SELECT * FROM Users WHERE UserName = '" & Replace(Username, "'", "''") & "';") Then
            dbSettings.LoadUser rs("ID")
            
            dbConn.rsQuery rs, "INSERT INTO ProfileSettings (ParentID) VALUES (" & dbSettings.CurrentUserID & ");"
            dbConn.rsQuery rs, "INSERT INTO ClientSettings (ParentID) VALUES (" & dbSettings.CurrentUserID & ");"
            dbConn.rsQuery rs, "INSERT INTO ScheduleSettings (ParentID) VALUES (" & dbSettings.CurrentUserID & ");"
            dbConn.rsQuery rs, "INSERT INTO ScriptingSettings (ParentID) VALUES (" & dbSettings.CurrentUserID & ");"
            
            SetupUser = True
        Else
            SetupUser = False
        End If

    Else
        dbSettings.LoadUser rs("ID")
        SetupUser = True
    End If
    
    rsClose rs
End Function

Private Sub CleanTemp()
    On Error Resume Next
    
    If PathExists(GetTemporaryFolder, False) Then

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim af As Object
        Set af = fso.GetFolder(GetTemporaryFolder)
        
        Dim fn As Object
        On Error Resume Next
        
        For Each fn In af.Files
            If fn.name Like "maxFTP*.lst" Then
            
                Kill fn.path
                If Not (Err.Number = 0) Then Err.Clear
            End If
        Next
    
        Set fso = Nothing
    
    End If
    On Error GoTo 0
End Sub
Private Sub CleanFolder2(ByVal FirstLevel As Boolean, ByRef af As Object)

    Dim sf As Object
    Dim fn As Object
    
    On Error Resume Next
    
    For Each sf In af.SubFolders
    
        SetAttr sf.path, VbFileAttribute.vbNormal
        If Not (Err.Number = 0) Then Err.Clear
        
        CleanFolder2 False, sf
    
    Next
    
    For Each fn In af.Files
    
        SetAttr fn.path, VbFileAttribute.vbNormal
        If Not (Err.Number = 0) Then Err.Clear
        
        Kill fn.path
        If Not (Err.Number = 0) Then Err.Clear
    
    Next
    
    If Not FirstLevel Then RmDir af.path
    If Not (Err.Number = 0) Then Err.Clear
    
    On Error GoTo 0
    
End Sub
Private Sub CleanFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    If PathExists(path) Then
    
        SetAttr path, VbFileAttribute.vbNormal
        If Not (Err.Number = 0) Then Err.Clear
        
        CleanFolder2 True, fso.GetFolder(path)

    End If
    
    Set fso = Nothing
    
    On Error GoTo 0
End Sub

Private Sub CleanActiveApp(ByRef dbConn As clsDBConnection, ByRef rs As ADODB.Recordset)

    Do
        On Error Resume Next
        SetAttr Replace(GetTemporaryFolder & "\" & ActiveAppFolder & rs("SubFolder") & "\" & rs("FileName"), "\\", "\"), VbFileAttribute.vbNormal
        Kill Replace(GetTemporaryFolder & "\" & ActiveAppFolder & rs("SubFolder") & "\" & rs("FileName"), "\\", "\")
        If InStr(GetTemporaryFolder & "\" & ActiveAppFolder & rs("SubFolder") & "\", "\\") = 0 Then RemovePath GetTemporaryFolder & "\" & ActiveAppFolder & rs("SubFolder")
        If Not (Err.Number = 0) Then Err.Clear
        On Error GoTo 0
        
        dbConn.dbQuery "DELETE * FROM ActiveApp WHERE ID=" & rs("ID") & ";"
        
        rs.MoveNext
    Loop Until rsEnd(rs)
    
End Sub
Public Sub CleanUser(Optional ByVal IsIDE As Boolean = False)
    
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT * FROM Users WHERE UserName='" & Replace(dbSettings.GetUserLoginName, "'", "''") & "';"
        
    If Not rsEnd(rs) Then
        Dim uID As Long
        uID = rs("ID")

        dbConn.rsQuery rs, "SELECT * FROM ActiveApp WHERE ParentID=" & uID & ";"
        If Not rsEnd(rs) Then
        
            If ((dbSettings.GetClientSetting("ActiveAppAsk") And (Not dbSettings.RemoveProfile))) And (Not IsIDE) Then
            
                Select Case MsgBox("Do you want to remove your Active App Cache files?" & vbCrLf & vbCrLf & "(Note: These files may be visible by other users)", vbQuestion + vbYesNo, AppName)
                    Case vbYes
                    
                        CleanActiveApp dbConn, rs
    
                End Select
                
            ElseIf dbSettings.RemoveProfile And (Not IsIDE) Then
                
                CleanActiveApp dbConn, rs
                
            End If
            
        End If
    
        If dbSettings.RemoveProfile Then
            
            dbConn.rsQuery rs, "DELETE * FROM History WHERE ParentID = " & uID & ";"
            dbConn.rsQuery rs, "DELETE * FROM SiteCache WHERE ParentID = " & uID & ";"
            dbConn.rsQuery rs, "DELETE * FROM ClientSettings WHERE ParentID = " & uID & ";"
            dbConn.rsQuery rs, "DELETE * FROM ProfileSettings WHERE ParentID = " & uID & ";"
            dbConn.rsQuery rs, "DELETE * FROM ScheduleSettings WHERE ParentID = " & uID & ";"

            dbConn.rsQuery rs, "DELETE * FROM ScriptingSettings WHERE ParentID = " & uID & ";"
            dbConn.rsQuery rs, "DELETE * FROM SessionDrives WHERE ParentID = " & uID & ";"

            dbConn.rsQuery rs, "SELECT * FROM Schedules WHERE ParentID = " & uID & ";"
            If Not rsEnd(rs) Then
                Dim col As New Collection
                
                Do
                
                    dbConn.dbQuery "DELETE * FROM Operations WHERE ParentID=" & rs("ID") & ";"
                    col.Add CLng(rs("ID"))
                    
                    rs.MoveNext
                Loop Until rsEnd(rs)
                
                dbConn.dbQuery "DELETE * FROM Schedules WHERE ParentID=" & uID & ";"
                
                If Not (ProcessRunning(ServiceFileName) = 0) Then
                    Dim Sid As Variant
                    For Each Sid In col
                        MessageQueueAdd ServiceFileName, "/loadschedule " & Sid
                    Next
                End If
                
                ClearCollection col
            End If
            
            dbConn.dbQuery "DELETE * FROM Users WHERE ID=" & uID & ";"
        
        End If

    End If
    
    rsClose rs
    Set dbConn = Nothing

    CleanTemp
    CleanFolder AppPath & TempFolder

End Sub






