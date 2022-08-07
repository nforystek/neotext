


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function WNetAddConnection2 Lib "mpr" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Const NO_ERROR = 0

Private Const CONNECT_UPDATE_PROFILE = &H1

Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_REMEMBERED = &H3
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2

Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_ALREADY_ASSIGNED = 85&
Private Const ERROR_BAD_DEV_TYPE = 66&
Private Const ERROR_BAD_DEVICE = 1200&
Private Const ERROR_BAD_NET_NAME = 67&
Private Const ERROR_BAD_PROFILE = 1206&
Private Const ERROR_BAD_PROVIDER = 1204&
Private Const ERROR_BUSY = 170&
Private Const ERROR_CANCELLED = 1223&
Private Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Private Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Private Const ERROR_EXTENDED_ERROR = 1208&
Private Const ERROR_INVALID_PASSWORD = 86&
Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Public Function ConnectNetworkDrive(ByVal DriveLetter As String, ByVal ShareName As String, Optional ByVal Username As String = "", Optional ByVal Password As String = "", Optional ByVal ShowPassBox As Boolean = False) As Boolean
On Error GoTo catch
    Dim connectionOK As String
    Dim netOk As Long

    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset

    If ShowPassBox Then
        netOk = ERROR_INVALID_PASSWORD
    Else
        netOk = -1
    End If
        
    DriveLetter = Left(DriveLetter, 1)
    
    Do
        If (netOk = ERROR_INVALID_PASSWORD Or netOk = ERROR_ACCESS_DENIED) Then
            If IsAppMaxFTP Then
                frmEquipment.ShowShareBox "Password required for " & ShareName, Username, Password
                Do Until frmEquipment.Visible = False
                    DoEvents
                    Sleep 1
                Loop
            Else
                netOk = ERROR_CANCELLED
            End If
        Else
            frmEquipment.IsOk = True
        End If
        
        If frmEquipment.IsOk And netOk <> ERROR_CANCELLED Then
            netOk = AddConnection(ShareName, frmEquipment.Username, frmEquipment.Password, DriveLetter + ":")
            Select Case netOk
                Case NO_ERROR
                    ShareName = DriveLetter + ":\" + ShareName
                Case ERROR_INVALID_PASSWORD
                    If IsNetDrivesVisible Then
                        MsgBox "Net Drive: INVALID_PASSWORD", vbOKOnly + vbExclamation, AppName
                    Else
                        MaxEvents.AddEvent dbSettings, "Mounting", "Drive letter " + DriveLetter + ":", "Error: INVALID_PASSWORD"
                    End If
                Case ERROR_ACCESS_DENIED
                    If IsNetDrivesVisible Then
                        MsgBox "Net Drive: ACCESS_DENIED", vbOKOnly + vbExclamation, AppName
                    End If
                Case Else
                    If PathExists(DriveLetter + ":") Then
                        netOk = NO_ERROR
                    Else
                        netOk = ERROR_CANCELLED
                        If IsNetDrivesVisible Then
                            MsgBox "Network connection error, can not mount drive " + DriveLetter + ":", vbOKOnly + vbExclamation, AppName
                        End If
                    End If
            End Select
        Else
            netOk = ERROR_CANCELLED
        End If
    
    Loop Until netOk = NO_ERROR Or netOk = ERROR_CANCELLED

    If (netOk = NO_ERROR) And (Username = "" And Password = "") Then
    
        Dim enc As New NTCipher10.ncode
        
        dbConn.rsQuery rs, "UPDATE SessionDrives SET Username='" & enc.EncryptString(Replace(frmEquipment.Username, "'", "''"), dbSettings.CryptKey) & "', Password='" & enc.EncryptString(Replace(frmEquipment.Password, "'", "''"), dbSettings.CryptKey(Replace(frmEquipment.Username, "'", "''"))) & "' WHERE DriveLetter='" & Replace(DriveLetter, "'", "''") & "' AND ParentID=" & dbSettings.CurrentUserID & ";"
        
        Set enc = Nothing
            
    End If
    
    Unload frmEquipment
        
    rsClose rs
    Set dbConn = Nothing
    
    RefreshDriveLists
    
    ConnectNetworkDrive = (netOk = NO_ERROR)

Exit Function
catch:
    If IsNetDrivesVisible Then
        MsgBox "Error: " + Err.Source + "  " + str(Err) + "  " + (Error), vbOKOnly + vbExclamation, AppName
    End If
    Err.Clear

End Function

Public Function IsNetDrivesVisible() As Boolean

    Dim cnt As Integer
    Dim vis As Boolean
    vis = False
    For cnt = 0 To Forms.Count - 1
        If TypeName(Forms(cnt)) = "frmNetDrives" Then
            If Forms(cnt).Visible Then
                vis = True
                Exit For
            End If
        End If
    Next
    
    IsNetDrivesVisible = vis

End Function

Public Function AddConnection(ByVal MyShareName As String, ByVal MyLogin As String, ByVal MyPWD As String, ByVal UseLetter As String) As Integer

    Dim NetR As NETRESOURCE
    Dim ErrInfo As Long
    NetR.dwScope = RESOURCE_GLOBALNET
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = UseLetter
    NetR.lpRemoteName = MyShareName

    ErrInfo = WNetAddConnection2(NetR, MyPWD, MyLogin, CONNECT_UPDATE_PROFILE)
    
    AddConnection = ErrInfo

End Function

Public Function CancelConnection(ByVal UseLetter As String) As Integer

    Dim ErrInfo As Long
    Dim strLocalName As String

    If Len(UseLetter) = 1 Then
        strLocalName = UseLetter & ":"
    Else
        strLocalName = UseLetter
    End If
    ErrInfo = WNetCancelConnection2(strLocalName, CONNECT_UPDATE_PROFILE, False)
        
    CancelConnection = ErrInfo

End Function

Public Function OpenSessionDrives()
    Dim enc As New NTCipher10.ncode
    Dim usr As String
    Dim pas As String
    
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT * FROM SessionDrives WHERE ParentID=" & dbSettings.CurrentUserID & ";"
    Do Until rsEnd(rs)
                        
        If Not ((rs("Username") & "") = "") Then
            usr = enc.DecryptString(rs("Username") & "", dbSettings.CryptKey)
            pas = enc.DecryptString(rs("Password") & "", dbSettings.CryptKey(usr))
        End If
        
        ConnectNetworkDrive rs("DriveLetter"), rs("ShareName"), usr, pas
            
        rs.MoveNext
    Loop

    rsClose rs
    Set dbConn = Nothing
    
    Set enc = Nothing
    
    RefreshDriveLists
End Function

Public Function CloseSessionDrives()

    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT * FROM SessionDrives WHERE ParentID=" & dbSettings.CurrentUserID & ";"
    Do Until rsEnd(rs)
                        
        CancelConnection rs("DriveLetter")
        
        rs.MoveNext
    Loop

    rsClose rs
    Set dbConn = Nothing

    RefreshDriveLists
End Function

Public Sub RefreshDriveLists()
    Dim frm As Object
    For Each frm In Forms
        If IsVisibleClient(frm) Then
            If frm.pViewDrives(0).Visible Then frm.pViewDrives(0).Refresh
            If frm.pViewDrives(1).Visible Then frm.pViewDrives(1).Refresh
        End If
    Next
End Sub
