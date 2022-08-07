VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dbConn As clsDBConnection

Private Const pDefaultUser = "WhoYouAre"
Private Const pDefaultComp = "WhereYouBe"

Public DefaultAccess As String

Public Enum AccessRights
    ar_Administrator = 3
    ar_CommonUser = 1
    ar_Undetermined = 0
End Enum

Public CurrentUserID As Long
Public CurrentUserProfile As String
Public CurrentUserAccessRights As Integer

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long

Public Property Get DefaultUser() As String
    DefaultUser = pDefaultUser
End Property
Public Property Get DefaultComp() As String
    DefaultComp = pDefaultComp
End Property

Public Function NoUsers() As Boolean
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    dbConn.rsQuery rs, "SELECT * FROM Users;"
    
    NoUsers = rsEnd(rs)
    
    rsClose rs
    
    Set dbConn = Nothing
End Function

Public Property Get RemoveProfile() As Boolean
    RemoveProfile = False
    
    If (CurrentUserAccessRights = ar_Administrator) Then
        RemoveProfile = False
    ElseIf GetPublicSetting("ServiceReadOnly") Then
        RemoveProfile = True
    Else
    
        Dim hToken As Long
        Dim sAllUser As String
        sAllUser = String(255, Chr(0))
        
        GetAllUsersProfileDirectory sAllUser, 255
        sAllUser = Replace(sAllUser, Chr(0), "")
        
        Dim sDefault As String
        GetDefaultUserProfileDirectory sDefault, 255
        sDefault = Replace(sDefault, Chr(0), "")
        
        Dim sLibrary As String
        sLibrary = String(255, Chr(0))
    
        OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
        GetUserProfileDirectory hToken, sLibrary, 255
        sLibrary = Replace(sLibrary, Chr(0), "")
        
        If ((GetAttr(sLibrary) And vbReadOnly) = vbReadOnly) Or _
            (LCase(Trim(sDefault)) = LCase(Trim(sLibrary))) Or _
            (LCase(Trim(sAllUser)) = LCase(Trim(sLibrary))) Then
            
            RemoveProfile = True
        End If
        
    End If
End Property

Public Function GetUserLoginName(Optional ByVal ExtInfo As Boolean = False) As String

    If ExtInfo Then
        Dim hToken As Long
        
        Dim sAllUser As String
        sAllUser = String(255, Chr(0))
        
        GetAllUsersProfileDirectory sAllUser, 255
        sAllUser = Replace(sAllUser, Chr(0), "")
        
        Dim sDefault As String
        GetDefaultUserProfileDirectory sDefault, 255
        sDefault = Replace(sDefault, Chr(0), "")
        
        Dim sLibrary As String
        sLibrary = String(255, Chr(0))
    
        OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
        GetUserProfileDirectory hToken, sLibrary, 255
        sLibrary = Replace(sLibrary, Chr(0), "")
    End If
    Dim sBuffer As String
    
    sBuffer = modUserInfo.GetUserLoginName
    
    If ExtInfo Then
        If (LCase(Trim(sAllUser)) = LCase(Trim(sLibrary))) Then
            GetUserLoginName = sBuffer & " (Shared Profile)"
        ElseIf (LCase(Trim(sDefault)) = LCase(Trim(sLibrary))) Or ((GetAttr(sLibrary) And vbReadOnly) = vbReadOnly) Then
            GetUserLoginName = sBuffer & " (Readonly Profile)"
        Else
            If ((Not (LCase(Trim(GetFileName(sLibrary))) = LCase(Trim(sBuffer)))) _
            And (Not (LCase(Trim(GetFileName(sAllUser))) = LCase(Trim(sBuffer))))) Then
    
                GetUserLoginName = sBuffer & " (" & GetFileName(sLibrary) & ")"
            Else
                GetUserLoginName = sBuffer & " (Windows Profile)"
            End If
        End If
    Else
        GetUserLoginName = sBuffer
    End If
    
End Function

'Public Function GetMachineName() As String
'
'    GetMachineName = DefaultComp
'
'    If CBool(GetPublicSetting("ServiceNetwork")) = True Then
'        Static strdomain As String
'        If strdomain = "" Then strdomain = GetDomain
'
'
'        If strdomain = "" Then
'            Static showmsgbox As Boolean
'            If Not showmsgbox Then
'
'                MsgBox "WARNING!!!!  Unable to resolve the domain or workgroup, it is recomended" & vbCrLf & _
'                        "that you disable the Network roaming allowance of users data, the domain" & vbCrLf & _
'                        "must be aquired to properly share encrypt/decrypt data of Max-FTP users." & vbCrLf _
'                        , vbCritical + vbOKOnly, "Warning"
'                showmsgbox = True
'            End If
'
'        End If
'        GetMachineName = strdomain
'
'    Else
'
'        Dim sBuffer As String
'        Dim lSize As Long
'        sBuffer = Space$(255)
'        lSize = Len(sBuffer)
'        Call GetCompName(sBuffer, lSize)
'        If lSize > 0 Then
'            GetMachineName = Replace(Left$(sBuffer, lSize), Chr(0), "")
'        End If
'
'    End If
'
'End Function

Public Function GetMachineName() As String
    Static CacheName As String
    If CacheName = "" Then
        Dim sBuffer As String
        
        sBuffer = modUserInfo.GetMachineName
        If CBool(GetPublicSetting("ServiceNetwork")) = True Then

            CacheName = GetDomainByMachine(sBuffer)

        Else
            CacheName = sBuffer

        End If

        If CacheName = "" Then CacheName = DefaultComp

    End If

    GetMachineName = CacheName

End Function

Public Function CryptKey(Optional Buck As String = "", Optional ParentID As Long = -1) As String
        
    If Not (ParentID = -1) Then
    
        Dim rs As New ADODB.Recordset
        dbConn.rsQuery rs, "SELECT * FROM Users WHERE ID=" & ParentID & ";"
        
        CryptKey = GetPublicSetting("ServiceInstallGUID") & GetPublicSetting("ServiceInstallGUID") & SecurityID & Me.GetMachineName & rs("UserName") & Buck

        If rs.State <> 0 Then rs.Close
        Set rs = Nothing

    Else
        CryptKey = GetPublicSetting("ServiceInstallGUID") & SecurityID & Me.GetMachineName & Buck
    End If

End Function

Public Function LoadUser(ByVal UserID As String)
    Dim rs As New ADODB.Recordset
    CurrentUserAccessRights = ar_Undetermined
    CurrentUserID = 0
    CurrentUserProfile = ""
    
    If dbConn.rsQuery(rs, "SELECT * FROM Users WHERE " & IIf(IsNumeric(UserID), "ID=" & UserID & ";", "UserName='" & Replace(UserID, "'", "''") & "';")) Then
        If Not rsEnd(rs) Then
            CurrentUserProfile = rs("UserName")
            CurrentUserID = rs("ID")
            CurrentUserAccessRights = rs("AccessRights")
            LoadUser = True
        Else
            LoadUser = False
        End If
    Else
        LoadUser = False
    End If

    rsClose rs
End Function

Public Function GetPublicSetting(ByVal Column As String) As Variant
    Dim sValue As Variant
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM PublicSettings;") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If
    
    rsClose rs
        
    GetPublicSetting = sValue

End Function
Public Function GetProfileSetting(ByVal Column As String) As Variant
    Dim sValue As Variant
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM ProfileSettings WHERE ParentID = " & CurrentUserID & ";") Then
        sValue = rs.Fields(Column).Value
    Else
        sValue = Null
    End If
    
    rsClose rs
        
    GetProfileSetting = sValue

End Function
Public Function GetClientSetting(ByVal Column As String) As Variant
    Dim sValue As Variant
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM ClientSettings WHERE ParentID = " & CurrentUserID & ";") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If

    rsClose rs
    
    GetClientSetting = sValue

End Function
Public Function GetScheduleSetting(ByVal Column As String) As Variant
    Dim sValue As Variant
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM ScheduleSettings WHERE ParentID = " & CurrentUserID & ";") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If

    rsClose rs
    
    GetScheduleSetting = sValue

End Function
Public Function GetScriptingSetting(ByVal Column As String) As Variant
    Dim sValue As Variant
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM ScriptingSettings WHERE ParentID = " & CurrentUserID & ";") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If

    rsClose rs
    
    GetScriptingSetting = sValue

End Function

Public Sub SetPublicSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE PublicSettings SET " & Column & "='" & Replace(Value, "'", "''") & "';"
    Else
        dbConn.dbQuery "UPDATE PublicSettings SET " & Column & "=" & Value & ";"
    End If
    
End Sub
Public Sub SetProfileSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE ProfileSettings SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ParentID=" & CurrentUserID & ";"
    Else
        dbConn.dbQuery "UPDATE ProfileSettings SET " & Column & "=" & Value & " WHERE ParentID=" & CurrentUserID & ";"
    End If

End Sub
Public Sub SetClientSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE ClientSettings SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ParentID=" & CurrentUserID & ";"
    Else
        dbConn.dbQuery "UPDATE ClientSettings SET " & Column & "=" & Value & " WHERE ParentID=" & CurrentUserID & ";"
    End If

End Sub
Public Sub SetScheduleSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE ScheduleSettings SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ParentID=" & CurrentUserID & ";"
    Else
        dbConn.dbQuery "UPDATE ScheduleSettings SET " & Column & "=" & Value & " WHERE ParentID=" & CurrentUserID & ";"
    End If

End Sub
Public Sub SetScriptingSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE ScriptingSettings SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ParentID=" & CurrentUserID & ";"
    Else
        dbConn.dbQuery "UPDATE ScriptingSettings SET " & Column & "=" & Value & " WHERE ParentID=" & CurrentUserID & ";"
    End If

End Sub

Private Sub Class_Initialize()
    Set dbConn = New clsDBConnection
End Sub

Private Sub Class_Terminate()
    Set dbConn = Nothing
End Sub
