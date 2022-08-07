#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modShared"
#Const modShared = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module

Public Const WebSite = "http://www.neotext.org"

Public Const AppName = "RemindMe"

Public Const SecurityID = "nt1783405269"

Public Const DBFileName = "remindme.mdb"

Public Const DBBackupExt = ".madb"

Public Const BackupFolder = "Backup\"

Public Const RemindMeFileName = "RemindMe.exe"

Public Const ServiceFileName = "RmdMeSrv.exe"

Public Const UtilityFileName = "Utility.exe"

Public Const ServiceName = "RemindMeService"

Public Function rsEnd(ByRef rs As ADODB.Recordset) As Boolean
    rsEnd = (rs.EOF Or rs.BOF)
End Function
Public Sub rsClose(ByRef rs As ADODB.Recordset, Optional ByVal SetNothing As Boolean = True)
    If Not rs.State = 0 Then rs.Close
    If SetNothing Then Set rs = Nothing
End Sub

Public Function DatabaseFilePath() As String
    Static retVal As String
    If retVal = "" Then
        retVal = AppPath & DBFileName
        If PathExists(Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare), True) And _
            (Not PathExists(retVal, True)) Then
            retVal = Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare)
        End If
    End If
    DatabaseFilePath = retVal
End Function
